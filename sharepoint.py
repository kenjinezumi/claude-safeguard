"""Scoped read-only SharePoint imports via Graph; OAuth stays on this Windows PC.
No model-callable configuration/login, remote writes, arbitrary HTTP or shell input.
"""
from pathlib import Path
from urllib.parse import urlsplit, unquote, parse_qs, quote
import hashlib
import json
import os
import re
import subprocess
import tempfile
import time
import uuid

BASE = Path(__file__).resolve().parent
GRAPH = 'https://graph.microsoft.com/v1.0'
SCOPES = ['https://graph.microsoft.com/Sites.Selected']
LIMIT = 25 * 1024 * 1024
TEXT_TYPES = {'.pdf', '.docx', '.pptx', '.odt', '.ods', '.odp', '.txt', '.csv',
              '.tsv', '.json', '.xml', '.html', '.htm', '.md', '.log'}
FIELDS = 'id,name,size,file,folder,package,remoteItem,parentReference,webUrl,eTag,deleted'
MESSAGES = {
    'SP-SETUP': 'Run Connect-SharePoint.cmd locally with the settings file supplied by IT.',
    'SP-LOGIN': 'Run Connect-SharePoint.cmd locally to sign in with your approved Microsoft work account.',
    'SP-ACCESS': 'Microsoft denied access. IT must check site grants, user access and Conditional Access policy.',
    'SP-NOT-FOUND': 'The SharePoint item was not found. Browse the approved folder to get its exact path.',
    'SP-SCOPE': 'The item or link is outside the configured SharePoint folder, or is a linked/unsupported item.',
    'SP-LINK': 'Use a direct document URL or browse the approved SharePoint folder. Opaque sharing links are unsupported.',
    'SP-LIMIT': 'A download, listing or converted document exceeded the supported limits.',
    'SP-TYPE': 'This file format is unsupported. XLSX, supported Office documents, text PDFs and UTF-8 text are accepted.',
    'SP-CHANGED': 'The source changed during import. Retry after editing has finished.',
    'SP-REVIEW': 'Local checks refused the content or conversion. Review the source locally; do not paste blocked data.',
    'SP-NETWORK': 'Microsoft could not be reached or throttled the request. Retry later or ask IT to check connectivity.',
    'SP-CHECK': 'The SharePoint operation stopped. No source content is included in this diagnostic.',
}
class SharePointError(Exception):
    def __init__(self, code):
        self.code = code if code in MESSAGES else 'SP-CHECK'
        super().__init__(self.code)

def fail(code):
    raise SharePointError(code)

def private_dir():
    if os.name != 'nt' or not os.environ.get('LOCALAPPDATA'):
        fail('SP-SETUP')
    path = Path(os.environ['LOCALAPPDATA']) / 'claude-safeguard-auth'
    path.mkdir(exist_ok=True)
    from folder_access import no_links
    no_links(path)
    return path

def checked_config(raw):
    if not isinstance(raw, dict): fail('SP-SETUP')
    keys = {'tenant_id', 'client_id', 'hostname', 'drive_id', 'root_item_id'}
    if set(raw) != keys: fail('SP-SETUP')
    c = dict(raw)
    for k in ('tenant_id', 'client_id'):
        try:
            c[k] = str(uuid.UUID(c[k]))
            if int(uuid.UUID(c[k])) == 0: fail('SP-SETUP')
        except (ValueError, TypeError, AttributeError): fail('SP-SETUP')
    if not isinstance(c['hostname'], str) or not re.fullmatch(r'[a-z0-9][a-z0-9-]*\.sharepoint\.com', c['hostname']):
        fail('SP-SETUP')
    for k in ('drive_id', 'root_item_id'):
        if not isinstance(c[k], str) or not re.fullmatch(r'[A-Za-z0-9!_.-]{1,200}', c[k]): fail('SP-SETUP')
    return c

def read_config():
    try:
        p = private_dir() / 'sharepoint-settings.json'
        from folder_access import no_links
        no_links(p)
        if p.stat().st_size > 16384: fail('SP-SETUP')
        return checked_config(json.loads(p.read_text(encoding='utf-8-sig')))
    except SharePointError: raise
    except Exception: fail('SP-SETUP')

def identity(c):
    return hashlib.sha256(json.dumps(c, sort_keys=True).encode()).hexdigest()

def auth_app(c):
    try:
        import msal
        from msal_extensions import build_encrypted_persistence, PersistedTokenCache
        cache_path = private_dir() / (identity(c) + '.bin')
        if cache_path.exists():
            from folder_access import no_links
            no_links(cache_path)
        persistence = build_encrypted_persistence(str(cache_path))
        if not persistence.is_encrypted: fail('SP-SETUP')
        return msal.PublicClientApplication(c['client_id'],
            authority='https://login.microsoftonline.com/' + c['tenant_id'],
            token_cache=PersistedTokenCache(persistence), enable_pii_log=False)
    except SharePointError: raise
    except Exception: fail('SP-SETUP')

def token(c):
    app = auth_app(c)
    accounts = app.get_accounts()
    if len(accounts) != 1: fail('SP-LOGIN')
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result or 'access_token' not in result: fail('SP-LOGIN')
    return result['access_token']

def relative_path(value):
    if not isinstance(value, str) or len(value) > 1800: fail('SP-SCOPE')
    if value == '': return ''
    if any(x in value for x in ('\\', ':', '%', '?', '#')) or any(ord(x) < 32 for x in value): fail('SP-SCOPE')
    if any(p in ('', '.', '..') for p in value.split('/')): fail('SP-SCOPE')
    return value

def url_path(value, hostname):
    if not isinstance(value, str) or len(value) > 4096: fail('SP-LINK')
    u = urlsplit(value)
    if u.scheme != 'https' or u.hostname != hostname or u.port not in (None, 443) or u.username or u.password: fail('SP-SCOPE')
    path = unquote(u.path)
    # Normal document URL, the /:w:/r/... browser form, or library view id=.
    path = re.sub(r'^/:[a-z]:/r/', '/', path)
    if path.lower().endswith('/forms/allitems.aspx'):
        vals = parse_qs(u.query).get('id', [])
        if len(vals) != 1: fail('SP-LINK')
        path = vals[0]
    elif path.startswith('/:') or path.lower().endswith('.aspx'):
        fail('SP-LINK')
    return '/' + relative_path(path.strip('/'))

class Client:
    def __init__(self, config, bearer, session=None):
        self.config = checked_config(config)
        self.bearer = bearer
        if session is None:
            import requests
            session = requests.Session()
        self.session = session
        self.drive = '/drives/' + quote(config['drive_id'], safe='')
        self.root_id = config['root_item_id']
        self._root = None

    def item_path(self, item_id):
        if not isinstance(item_id, str) or not re.fullmatch(r'[A-Za-z0-9!_.-]{1,200}', item_id): fail('SP-SCOPE')
        return self.drive + '/items/' + quote(item_id, safe='')

    def get(self, url, authenticated):
        u = urlsplit(url)
        expected_host = 'graph.microsoft.com' if authenticated else self.config['hostname']
        if u.scheme != 'https' or u.hostname != expected_host or u.port not in (None, 443) or u.username or u.password: fail('SP-SCOPE')
        if authenticated and not u.path.startswith('/v1.0' + self.drive + '/'): fail('SP-SCOPE')
        headers = {'Authorization': 'Bearer ' + self.bearer} if authenticated else {}
        try:
            response = self.session.get(url, headers=headers, allow_redirects=False, stream=True, timeout=(10, 45), auth=lambda prepared: prepared)
        except Exception: fail('SP-NETWORK')
        if response.status_code in (401, 403, 404, 429) or response.status_code >= 500:
            code = {401:'SP-LOGIN',403:'SP-ACCESS',404:'SP-NOT-FOUND'}.get(response.status_code,'SP-NETWORK')
            response.close(); fail(code)
        return response

    @staticmethod
    def bounded(response, limit):
        data = bytearray(); start = time.monotonic()
        declared = response.headers.get('Content-Length')
        if declared and int(declared) > limit: fail('SP-LIMIT')
        for chunk in response.iter_content(chunk_size=65536):
            data.extend(chunk)
            if len(data) > limit or time.monotonic()-start > 60: fail('SP-LIMIT')
        return bytes(data)

    def json(self, path):
        response = self.get(GRAPH + path, True)
        try:
            if response.status_code != 200: fail('SP-NETWORK')
            value = json.loads(self.bounded(response, 1024*1024))
            if not isinstance(value, dict): fail('SP-CHECK')
            return value
        finally: response.close()

    def metadata(self, item_id):
        return self.json(self.item_path(item_id) + '?$select=' + FIELDS)

    def root(self):
        if self._root is None:
            item = self.metadata(self.root_id)
            if item.get('id') != self.root_id or 'folder' not in item or any(k in item for k in ('remoteItem','package','deleted')): fail('SP-SCOPE')
            url_path(item.get('webUrl'), self.config['hostname'])
            self._root = item
        return self._root

    def path_from_url(self, url):
        base = url_path(self.root()['webUrl'], self.config['hostname']).rstrip('/')
        candidate = url_path(url, self.config['hostname'])
        if candidate.casefold() == base.casefold(): return ''
        if not candidate.casefold().startswith(base.casefold()+'/'): fail('SP-SCOPE')
        return relative_path(candidate[len(base)+1:])

    def within(self, item):
        current = item; seen = set()
        for _ in range(40):
            item_id = current.get('id')
            if not item_id or item_id in seen or any(k in current for k in ('remoteItem','package','deleted')): fail('SP-SCOPE')
            seen.add(item_id)
            if item_id == self.root_id: return item
            parent = current.get('parentReference', {})
            if parent.get('driveId') != self.config['drive_id'] or not parent.get('id'): fail('SP-SCOPE')
            current = self.metadata(parent['id'])
        fail('SP-SCOPE')

    def resolve(self, path):
        path = relative_path(path)
        if not path: return self.root()
        self.root()
        item = self.json(self.item_path(self.root_id) + ':/' + quote(path, safe='/') + '?$select=' + FIELDS)
        self.within(item)
        if self.path_from_url(item.get('webUrl')).casefold() != path.casefold(): fail('SP-SCOPE')
        return item

    def download(self, item, destination):
        if 'file' not in item or item.get('size', LIMIT+1) > LIMIT: fail('SP-LIMIT')
        response = self.get(GRAPH + self.item_path(item['id']) + '/content', True)
        try:
            # Graph returns a preauthenticated download URL. Never forward OAuth.
            if response.status_code == 200:
                data = self.bounded(response, LIMIT)
            else:
                if response.status_code != 302: fail('SP-NETWORK')
                location = response.headers.get('Location', '')
                response.close()
                for _ in range(3):
                    response = self.get(location, False)
                    if response.status_code == 200: break
                    if response.status_code not in (301,302,303,307,308): fail('SP-NETWORK')
                    location = response.headers.get('Location','')
                    response.close()
                if response.status_code != 200: fail('SP-NETWORK')
                data = self.bounded(response, LIMIT)
            # Detect source edits during download; retain no stale import.
            after = self.within(self.metadata(item['id']))
            if not item.get('eTag') or after.get('eTag') != item['eTag']: fail('SP-CHANGED')
            destination.write_bytes(data)
        finally: response.close()

PAGES = {}
def browse(client, path, scan, page=None):
    item = client.resolve(path)
    if 'folder' not in item: fail('SP-TYPE')
    children = client.item_path(item['id']) + '/children'
    next_url = None
    if page:
        saved = PAGES.pop(page, None)
        if not saved or saved[:2] != (identity(client.config), path) or saved[2] < time.monotonic(): fail('SP-SCOPE')
        next_url = saved[3]
    path_query = children + '?$top=200&$select=' + FIELDS
    if next_url:
        if not next_url.startswith(GRAPH + children + '?'): fail('SP-SCOPE')
        path_query = next_url[len(GRAPH):]
    data = client.json(path_query)
    values = data.get('value')
    if not isinstance(values, list) or len(values)>200: fail('SP-LIMIT')
    entries=[]
    for child in values:
        try:
            name = relative_path(child['name'])
            if '/' in name or not name or any(k in child for k in ('remoteItem','package','deleted')): raise ValueError()
            parent=child.get('parentReference',{})
            if parent.get('driveId')!=client.config['drive_id'] or parent.get('id')!=item['id']: raise ValueError()
            rel = (path+'/' if path else '')+name
            scan(rel)
            entries.append({'path':rel,'type':'folder' if 'folder' in child else 'file',
                            'can_import':Path(name).suffix.lower() in TEXT_TYPES|{'.xlsx'}})
        except Exception:
            entries.append({'type':'blocked','reason':'Unsupported or potentially sensitive name withheld.'})
    page_token=None
    if data.get('@odata.nextLink'):
        next_url=data['@odata.nextLink']
        if not isinstance(next_url,str) or not next_url.startswith(GRAPH+children+'?'): fail('SP-SCOPE')
        if len(PAGES)>=100: PAGES.clear()
        page_token=uuid.uuid4().hex
        PAGES[page_token]=(identity(client.config),path,time.monotonic()+900,next_url)
    return {'entries':entries,'next_page':page_token,'source':'Approved SharePoint folder'}

def guard_file(path):
    powershell = Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe'
    try:
        result = subprocess.run([str(powershell),'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass',
            '-File',str(BASE/'guard.ps1'),'-CheckFile',str(path)], capture_output=True, encoding='utf-8', timeout=20)
        check = json.loads(result.stdout)
        if result.returncode != 0 or check.get('status') != 'passed': fail('SP-REVIEW')
    except SharePointError: raise
    except Exception: fail('SP-CHECK')

def import_document(client, path, workbooks, load, scan):
    item = client.resolve(path)
    scan(path)
    ext=Path(item.get('name','')).suffix.lower()
    if ext not in TEXT_TYPES|{'.xlsx'}: fail('SP-TYPE')
    if 'file' not in item or any(k in item for k in ('remoteItem','package','deleted')): fail('SP-SCOPE')
    destination=None
    try:
        # Raw downloads never enter the Claude project. Cleanup is normal deletion.
        with tempfile.TemporaryDirectory(prefix='download-',dir=private_dir()) as folder:
            raw=Path(folder)/('source'+ext)
            client.download(item, raw)
            if ext=='.xlsx':
                workbooks.mkdir(exist_ok=True)
                name=uuid.uuid4().hex+'.xlsx';destination=workbooks/name
                destination.write_bytes(raw.read_bytes())
                wb=load(name)
                result={'workbook':name,'sheets':wb.sheetnames,'source':'SharePoint','source_unchanged':True}
            else:
                from convert import extract
                text=extract(raw)
                if not text.strip() or len(text.encode('utf-8'))>1024*1024: fail('SP-LIMIT')
                scan(text)
                approved=BASE/'Approved-Data';approved.mkdir(exist_ok=True)
                name='prepared-'+uuid.uuid4().hex+'.txt';destination=approved/name
                destination.write_text(text,encoding='utf-8')
                guard_file(destination)
                result={'prepared_file':name,'folder':'Approved-Data','source':'SharePoint',
                    'note':'Read this exact prepared file using scanned Read. Text extraction may omit images/layout; review completeness.'}
            scan(json.dumps(result))
            return result
    except Exception:
        if destination is not None: destination.unlink(missing_ok=True)
        raise

def execute(request, workbooks, load, scan):
    # Fixed messages only; Graph bodies, token cache and preauthenticated URLs
    # must never become tool output, including on network/auth failures.
    try:
        c=read_config()
        operation=request.get('operation')
        if operation not in ('sharepoint_status','sharepoint_browse','sharepoint_import'): fail('SP-SCOPE')
        client=Client(c,token(c))
        if operation=='sharepoint_status':
            client.root()
            return {'status':'connected','access':'read-only approved SharePoint folder','credentials':'kept locally'}
        if 'url' in request and 'path' in request: fail('SP-SCOPE')
        path=client.path_from_url(request['url']) if 'url' in request else relative_path(request.get('path',''))
        scan(path)
        if operation=='sharepoint_browse': return browse(client,path,scan,request.get('page'))
        return import_document(client,path,workbooks,load,scan)
    except SharePointError as error:
        return {'error':error.code,'message':MESSAGES[error.code]}
    except Exception:
        return {'error':'SP-REVIEW','message':MESSAGES['SP-REVIEW']}

def main():
    import sys
    try:
        action=sys.argv[1]
        if action=='--configure':
            source=Path(sys.argv[2])
            if source.stat().st_size>16384: fail('SP-SETUP')
            c=checked_config(json.loads(source.read_text(encoding='utf-8-sig')))
            target=private_dir()/'sharepoint-settings.json'
            temp=target.with_suffix('.tmp')
            temp.write_text(json.dumps(c),encoding='utf-8');temp.replace(target)
            print('SharePoint settings saved locally. Next: Microsoft sign-in.')
        elif action=='--sign-in':
            c=read_config();app=auth_app(c)
            # Dedicated cache for this configuration: prevent ambiguous accounts.
            for account in app.get_accounts(): app.remove_account(account)
            result=app.acquire_token_interactive(scopes=SCOPES,prompt='select_account',timeout=180)
            if 'access_token' not in result: fail('SP-LOGIN')
            Client(c,result['access_token']).root()
            print('Connected to the approved SharePoint folder. Start a fresh safeguard session.')
        elif action=='--sign-out':
            c=read_config();app=auth_app(c)
            for account in app.get_accounts(): app.remove_account(account)
            print('Local SharePoint account cache cleared. Imported files remain. Close Claude to end any in-flight request.')
        else: fail('SP-SETUP')
        return 0
    except SharePointError as error:
        print(error.code+': '+MESSAGES[error.code]);return 1
    except Exception:
        print('SP-CHECK: Setup stopped. No credentials or source values included. Contact IT.');return 1

if __name__=='__main__':
    raise SystemExit(main())
