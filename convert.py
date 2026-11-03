"""Local text extraction only. No network, macros, formulas, or external links executed."""
from pathlib import Path
import sys, zipfile, xml.etree.ElementTree as ET, re, csv, io
LIMIT=1024*1024

def xml(data):
    if b'<!DOCTYPE' in data.upper() or b'<!ENTITY' in data.upper():
        raise ValueError('DTD forbidden')
    return ET.fromstring(data)

def extract(source):
    if source.stat().st_size > 25*LIMIT: raise ValueError('Too large')
    ext=source.suffix.lower()
    if ext in {'.txt','.csv','.tsv','.json','.xml','.html','.htm','.md','.log'}:
        text=source.read_text(encoding='utf-8-sig')
        if '\x00' in text: raise ValueError('Binary')
        return text
    if ext=='.pdf':
        from pypdf import PdfReader
        reader=PdfReader(source)
        if reader.is_encrypted: raise ValueError('Encrypted')
        if len(reader.pages)>200: raise ValueError('Too many pages')
        return '\n'.join(f'PAGE {i+1}\n{page.extract_text() or "[No extractable text: image content omitted]"}' for i,page in enumerate(reader.pages))
    if ext not in {'.docx','.xlsx','.pptx','.odt','.ods','.odp'}: raise ValueError('Unsupported')
    with zipfile.ZipFile(source) as z:
        if len(z.infolist())>10000 or sum(i.file_size for i in z.infolist())>50*LIMIT: raise ValueError('Archive too large')
        if ext=='.xlsx':
            ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            strings=[]
            if 'xl/sharedStrings.xml' in z.namelist():
                strings=[''.join(t.text or '' for t in n.findall('.//s:t',ns)) for n in xml(z.read('xl/sharedStrings.xml')).findall('s:si',ns)]
            out=io.StringIO(); w=csv.writer(out)
            for name in sorted(n for n in z.namelist() if re.fullmatch(r'xl/worksheets/sheet[0-9]+.xml',n)):
                w.writerow(['SHEET PART',name,'Cell coordinates retained; values may be cached/stale'])
                for cell in xml(z.read(name)).findall('.//s:c',ns):
                    value=cell.find('s:v',ns); value=value.text if value is not None else ''
                    if cell.get('t')=='s': value=strings[int(value)]
                    elif cell.get('t')=='inlineStr': value=''.join(t.text or '' for t in cell.findall('.//s:t',ns))
                    formula=cell.find('s:f',ns)
                    w.writerow([cell.get('r'),value,('Formula: '+(formula.text or '')) if formula is not None else ''])
            return out.getvalue()
        names=sorted(n for n in z.namelist() if
            (ext=='.docx' and n.startswith('word/') and n.endswith('.xml')) or
            (ext=='.pptx' and (n.startswith('ppt/slides/slide') or n.startswith('ppt/notesSlides/notesSlide')) and n.endswith('.xml')) or
            (ext in {'.odt','.ods','.odp'} and n=='content.xml'))
        parts=[]
        for name in names:
            root=xml(z.read(name))
            if ext in {'.odt','.ods','.odp'}:
                paragraphs=[''.join(n.itertext()) for n in root.iter() if n.tag.rsplit('}',1)[-1] in {'p','h'}]
            else:
                paragraphs=[''.join(t.text or '' for t in n.iter() if t.tag.rsplit('}',1)[-1]=='t') for n in root.iter() if n.tag.rsplit('}',1)[-1]=='p']
            if paragraphs: parts.append('PART '+name+'\n'+'\n'.join(paragraphs))
        return '\n\n'.join(parts)

if __name__=='__main__':
    try:
        text=extract(Path(sys.argv[1]))
        if not text.strip() or len(text.encode('utf-8'))>LIMIT: raise ValueError('Empty/too large')
        Path(sys.argv[2]).write_text(text,encoding='utf-8')
    except Exception:
        print('Local extraction failed; no content uploaded.',file=sys.stderr)
        sys.exit(1)
