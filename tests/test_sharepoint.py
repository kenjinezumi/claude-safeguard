"""Offline Graph transport tests. No Microsoft account/network/real user files.
Includes actual document conversion; text imports run the actual PowerShell guard
when SAFEGUARD_TEST_PWSH names a PowerShell binary (Linux portable runtime is OK).
"""
from pathlib import Path
from urllib.parse import unquote, urlsplit
import copy
import io
import json
import os
import shutil
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import sharepoint as sp
import spreadsheet as sheets

CONFIG={'tenant_id':'11111111-1111-1111-1111-111111111111',
        'client_id':'22222222-2222-2222-2222-222222222222',
        'hostname':'example.sharepoint.com','drive_id':'drive1','root_item_id':'root1'}
URL='https://example.sharepoint.com/sites/Test/Documents/Approved'
PWSH=os.environ.get('SAFEGUARD_TEST_PWSH') or shutil.which('pwsh')

class Response:
    def __init__(self,status=200,payload=None,data=b'',headers=None):
        self.status_code=status
        self.data=json.dumps(payload).encode() if payload is not None else data
        self.headers=headers or {}
        self.closed=False
    def iter_content(self,chunk_size):
        for i in range(0,len(self.data),chunk_size): yield self.data[i:i+chunk_size]
    def close(self): self.closed=True

class Session:
    def __init__(self):
        self.calls=[];self.redirect='https://example.sharepoint.com/download?secret=DO_NOT_LEAK'
        self.source=b'Meal allowance GBP 20.'
        self.items={
            'root1':{'id':'root1','name':'Approved','folder':{},'webUrl':URL,'eTag':'r1'},
            'sub1':{'id':'sub1','name':'Policies','folder':{},'webUrl':URL+'/Policies','eTag':'s1',
                    'parentReference':{'driveId':'drive1','id':'root1'}},
            'file1':{'id':'file1','name':'Policy.txt','size':100,'file':{},'eTag':'f1',
                     'webUrl':URL+'/Policies/Policy.txt','parentReference':{'driveId':'drive1','id':'sub1'}}}
        self.next=None;self.changed=False;self.status=None
    def get(self,url,**kwargs):
        self.calls.append((url,kwargs))
        if self.status: return Response(self.status,payload={'private_error':'DO_NOT_LEAK'})
        if url.startswith('https://example.sharepoint.com/download'):
            if self.changed: self.items['file1']['eTag']='f2'
            return Response(data=self.source)
        if not url.startswith(sp.GRAPH+'/drives/drive1/items/'):
            raise AssertionError('Unexpected network destination')
        path=unquote(urlsplit(url).path.split('/items/',1)[1])
        if path.endswith('/content'): return Response(302,headers={'Location':self.redirect})
        if path.endswith('/children'):
            item_id=path.split('/')[0]
            data={'value':[copy.deepcopy(v) for v in self.items.values() if v.get('parentReference',{}).get('id')==item_id]}
            if self.next: data['@odata.nextLink']=self.next
            return Response(payload=data)
        if ':/' in path:
            _,rel=path.split(':/',1)
            for item in self.items.values():
                if item['webUrl']==URL+'/'+rel: return Response(payload=copy.deepcopy(item))
            return Response(404)
        return Response(payload=copy.deepcopy(self.items[path]))

class SharePointTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(prefix='sharepoint-test-')
        self.root=Path(self.tmp.name);self.work=self.root/'work';self.work.mkdir()
        self.private=self.root/'auth';self.private.mkdir()
        self.books=self.work/'Workbooks';self.books.mkdir()
        self.session=Session();self.client=sp.Client(CONFIG,'TOKEN_DO_NOT_LEAK',self.session)
        self.patches=[patch.object(sp,'BASE',self.work),patch.object(sp,'private_dir',return_value=self.private),patch.object(sheets,'ROOT',self.books)]
        for p in self.patches:p.start()
        sp.PAGES.clear()
    def tearDown(self):
        for p in reversed(self.patches):p.stop()
        self.tmp.cleanup()
    def error(self,code,callback):
        with self.assertRaises(sp.SharePointError) as cm:callback()
        self.assertEqual(cm.exception.code,code)
    def test_config_rejects_missing_or_extra_credentials(self):
        self.error('SP-SETUP',lambda:sp.checked_config({**CONFIG,'client_secret':'private'}))
        self.error('SP-SETUP',lambda:sp.checked_config({**CONFIG,'hostname':'example.sharepoint.com.evil.test'}))
    def test_traversal_and_absolute_paths_rejected(self):
        for path in ['../a','Policies/../a','/root','a//b','a\\b','https://example.com','%2e%2e/a','a\nb']:
            with self.subTest(path=path):self.error('SP-SCOPE',lambda:sp.relative_path(path))
    def test_direct_url_and_browser_url(self):
        self.assertEqual(self.client.path_from_url(URL+'/Policies/Policy.txt?web=1'),'Policies/Policy.txt')
        self.assertEqual(self.client.path_from_url('https://example.sharepoint.com/:w:/r/sites/Test/Documents/Approved/Policies/Policy.txt?x=1'),'Policies/Policy.txt')
    def test_url_wrong_host_scheme_or_root_rejected(self):
        for url in ['http://example.sharepoint.com/sites/Test/Documents/Approved/a.txt',URL.replace('example.sharepoint.com','evil.test')+'/a',URL+'-other/a',URL+'/../a',URL+'/%252e%252e/a']:
            with self.subTest(url=url):self.error('SP-SCOPE',lambda:self.client.path_from_url(url))
    def test_opaque_link_rejected(self):
        self.error('SP-LINK',lambda:self.client.path_from_url('https://example.sharepoint.com/:w:/s/Test/OPAQUE'))
    def test_resolve_proves_ancestry(self):
        self.assertEqual(self.client.resolve('Policies/Policy.txt')['id'],'file1')
    def test_cross_drive_and_remote_item_refused(self):
        self.session.items['file1']['parentReference']['driveId']='other'
        self.error('SP-SCOPE',lambda:self.client.resolve('Policies/Policy.txt'))
        self.session.items['file1']['parentReference']['driveId']='drive1'
        self.session.items['file1']['remoteItem']={'id':'elsewhere'}
        self.error('SP-SCOPE',lambda:self.client.resolve('Policies/Policy.txt'))
    def test_ancestor_loop_refused(self):
        self.session.items['sub1']['parentReference']['id']='sub1'
        self.error('SP-SCOPE',lambda:self.client.resolve('Policies/Policy.txt'))
    def test_listing_scans_names(self):
        self.session.items['file1']['name']='person@example.org.txt'
        result=sp.browse(self.client,'Policies',sheets.scan)
        self.assertEqual(result['entries'][0]['type'],'blocked')
        self.assertNotIn('person@example.org',json.dumps(result))
    def test_pagination_opaque_and_bound_to_path(self):
        self.session.next=sp.GRAPH+'/drives/drive1/items/root1/children?$skiptoken=DO_NOT_LEAK'
        result=sp.browse(self.client,'',sheets.scan)
        self.assertTrue(result['next_page']);self.assertNotIn('DO_NOT_LEAK',json.dumps(result))
        self.error('SP-SCOPE',lambda:sp.browse(self.client,'Policies',sheets.scan,result['next_page']))
    def test_external_nextlink_rejected(self):
        self.session.next='https://evil.test/next'
        self.error('SP-SCOPE',lambda:sp.browse(self.client,'',sheets.scan))
    def test_download_never_forwards_token(self):
        item=self.client.resolve('Policies/Policy.txt');out=self.root/'download'
        self.client.download(item,out)
        self.assertEqual(out.read_bytes(),self.session.source)
        for url,options in self.session.calls:
            self.assertFalse(options['allow_redirects'])
            self.assertTrue(callable(options['auth']))
            if url.startswith(sp.GRAPH):self.assertEqual(options['headers']['Authorization'],'Bearer TOKEN_DO_NOT_LEAK')
            else:self.assertNotIn('Authorization',options['headers'])
    def test_redirect_to_other_host_blocked_before_request(self):
        self.session.redirect='https://evil.test/steal'
        self.error('SP-SCOPE',lambda:self.client.download(self.client.resolve('Policies/Policy.txt'),self.root/'x'))
        self.assertFalse(any('evil.test' in url for url,_ in self.session.calls))
    def test_changed_source_not_saved(self):
        self.session.changed=True;out=self.root/'x'
        self.error('SP-CHANGED',lambda:self.client.download(self.client.resolve('Policies/Policy.txt'),out))
        self.assertFalse(out.exists())
    def test_size_limit_before_download(self):
        self.session.items['file1']['size']=sp.LIMIT+1
        self.error('SP-LIMIT',lambda:self.client.download(self.client.resolve('Policies/Policy.txt'),self.root/'x'))
        self.assertFalse(any(url.endswith('/content') for url,_ in self.session.calls))
    def test_stream_limit_without_content_length(self):
        old=sp.LIMIT
        with patch.object(sp,'LIMIT',8):
            item=self.client.resolve('Policies/Policy.txt');item['size']=1
            self.error('SP-LIMIT',lambda:self.client.download(item,self.root/'x'))
    def test_download_error_cleans_raw_temp(self):
        self.session.redirect='https://evil.test/a'
        self.error('SP-SCOPE',lambda:sp.import_document(self.client,'Policies/Policy.txt',self.books,sheets.load,sheets.scan))
        self.assertEqual(list(self.private.iterdir()),[])
        self.assertEqual(list(self.books.iterdir()),[])
    def test_unsupported_extension_before_download(self):
        self.session.items['file1']['name']='Policy.exe'
        self.error('SP-TYPE',lambda:sp.import_document(self.client,'Policies/Policy.txt',self.books,sheets.load,sheets.scan))
        self.assertFalse(any(url.endswith('/content') for url,_ in self.session.calls))
    def test_sensitive_text_not_imported(self):
        self.session.source=b'SECRET_SENTINEL person@example.org'
        with self.assertRaises(ValueError):sp.import_document(self.client,'Policies/Policy.txt',self.books,sheets.load,sheets.scan)
        self.assertEqual(list(self.private.iterdir()),[])
        self.assertFalse((self.work/'Approved-Data').exists())
    def test_workbook_import_real_scanner(self):
        import openpyxl
        wb=openpyxl.Workbook();wb.active.append(['Department','Actual']);wb.active.append(['Newsroom',200])
        data=io.BytesIO();wb.save(data);self.session.source=data.getvalue()
        self.session.items['file1']['name']='Policy.xlsx'
        result=sp.import_document(self.client,'Policies/Policy.txt',self.books,sheets.load,sheets.scan)
        self.assertTrue((self.books/result['workbook']).exists())
        self.assertEqual(sheets.load(result['workbook']).active['B2'].value,200)
        self.assertEqual(list(self.private.iterdir()),[])
    def test_sensitive_workbook_removed(self):
        import openpyxl
        wb=openpyxl.Workbook();wb.active['A1']='person@example.org'
        data=io.BytesIO();wb.save(data);self.session.source=data.getvalue()
        self.session.items['file1']['name']='Policy.xlsx'
        with self.assertRaises(ValueError):sp.import_document(self.client,'Policies/Policy.txt',self.books,sheets.load,sheets.scan)
        self.assertEqual(list(self.books.iterdir()),[])
        self.assertEqual(list(self.private.iterdir()),[])
    @unittest.skipUnless(PWSH,'Requires PowerShell')
    def test_docx_conversion_and_real_guard(self):
        actual_source=Path(__file__).resolve().parents[1]
        shutil.copy2(actual_source/'guard.ps1',self.work/'guard.ps1')
        (self.work/'CLAUDE.md').write_text('Synthetic policy')
        windows=Path(os.environ.get('SystemRoot', str(self.root/'windows')))
        if os.name=='posix':
            windows=self.root/'windows';exe=windows/'System32/WindowsPowerShell/v1.0/powershell.exe'
            exe.parent.mkdir(parents=True);exe.symlink_to(PWSH)
        data=io.BytesIO()
        with zipfile.ZipFile(data,'w') as z:z.writestr('word/document.xml','<document><p><t>Meal allowance GBP 20.</t></p></document>')
        self.session.source=data.getvalue();self.session.items['file1']['name']='Policy.docx'
        with patch.dict(os.environ,{'SystemRoot':str(windows)}):
            result=sp.import_document(self.client,'Policies/Policy.txt',self.books,sheets.load,sheets.scan)
        self.assertIn('Meal allowance',(self.work/'Approved-Data'/result['prepared_file']).read_text())
        self.assertEqual(list(self.private.iterdir()),[])
    def test_guard_failure_removes_prepared_copy(self):
        with patch.object(sp,'guard_file',side_effect=sp.SharePointError('SP-REVIEW')):
            self.error('SP-REVIEW',lambda:sp.import_document(self.client,'Policies/Policy.txt',self.books,sheets.load,sheets.scan))
        self.assertEqual(list((self.work/'Approved-Data').iterdir()),[])
    def test_error_never_contains_graph_body_or_token(self):
        self.session.status=403
        with patch.object(sp,'read_config',return_value=CONFIG),patch.object(sp,'token',return_value='SECRET_TOKEN'),patch.object(sp,'Client',return_value=self.client):
            result=sp.execute({'operation':'sharepoint_status'},self.books,sheets.load,sheets.scan)
        self.assertEqual(result['error'],'SP-ACCESS')
        self.assertNotIn('DO_NOT_LEAK',json.dumps(result));self.assertNotIn('SECRET_TOKEN',json.dumps(result))
    def test_unconfigured_returns_actionable_error(self):
        result=sp.execute({'operation':'sharepoint_status'},self.books,sheets.load,sheets.scan)
        self.assertEqual(result['error'],'SP-SETUP')

if __name__=='__main__':unittest.main(verbosity=2)
