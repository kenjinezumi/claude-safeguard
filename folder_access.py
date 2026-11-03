"""Read-only source folder access. The user selects the root outside Claude."""
from pathlib import Path
import json,os,stat,uuid,itertools
CONFIG=Path(__file__).resolve().parent/'approved-folder.json'

def no_links(p):
    for node in [p,*p.parents]:
        s=node.lstat()
        if stat.S_ISLNK(s.st_mode) or getattr(s,'st_file_attributes',0)&1024:
            raise ValueError('Linked paths refused')

def scoped(relative=''):
    raw=json.loads(CONFIG.read_text(encoding='utf-8-sig'))['root']
    if raw.startswith(('\\\\','//')):raise ValueError('Network roots not supported')
    root=Path(raw)
    if not root.is_absolute() or not root.is_dir():raise ValueError('Select a valid local folder')
    no_links(root);root=root.resolve(strict=True)
    if not isinstance(relative,str) or '\\' in relative or ':' in relative or relative.startswith('/') or '..' in Path(relative).parts:
        raise ValueError('Use relative paths within the approved folder')
    requested=root/relative
    no_links(requested)
    resolved=requested.resolve(strict=True)
    if not resolved.is_relative_to(root):raise ValueError('Outside approved folder')
    return root,resolved

def execute(r,workbooks,load,scan):
    root,p=scoped(r.get('path',''))
    if r['operation']=='browse':
        if not p.is_dir():raise ValueError('Not a directory')
        items=[]
        for child in itertools.islice(p.iterdir(),501):
            if len(items)>=500:return {'entries':items,'truncated':True,'note':'Open a narrower subfolder to see additional entries.'}
            try:
                no_links(child)
                relative=child.relative_to(root).as_posix();scan(relative)
                is_dir=child.is_dir()
                items.append({'path':relative,'type':'folder' if is_dir else 'file','bytes':None if is_dir else child.stat().st_size,'can_import':not is_dir and child.suffix.lower()=='.xlsx'})
            except Exception:
                items.append({'type':'blocked','reason':'Linked, inaccessible or potentially identifying name withheld'})
        return {'entries':items,'truncated':False,'root':'User-selected folder (absolute path withheld)'}
    if r['operation']=='import_file':
        if not p.is_file() or p.suffix.lower()!='.xlsx':raise ValueError('Only XLSX source import supported')
        if p.stat().st_size>25*1024*1024:raise ValueError('Too large')
        workbooks.mkdir(exist_ok=True);name=uuid.uuid4().hex+'.xlsx';dest=workbooks/name
        try:
            # Bounded read; source is never written.
            with p.open('rb') as f:data=f.read(25*1024*1024+1)
            if len(data)>25*1024*1024:raise ValueError('Too large')
            dest.write_bytes(data);wb=load(name)
            return {'workbook':name,'sheets':wb.sheetnames,'source_unchanged':True}
        except Exception:
            dest.unlink(missing_ok=True);raise
    raise ValueError('Unknown folder operation')
