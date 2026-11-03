"""Local, bounded spreadsheet operations; never evaluates Python or shell input."""
from pathlib import Path
import json,re,uuid,zipfile,io,os
import openpyxl
import pandas as pd
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.chart import BarChart,LineChart,PieChart,Reference
from openpyxl.utils.cell import range_boundaries
ROOT=Path(__file__).resolve().parent/'Workbooks'
PATTERNS=[r'(?i)[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}',r'(?<!\w)\+[1-9][0-9 ()-]{9,}[0-9](?!\w)',r'(?i)\b(password|api[_ -]?key|secret|access[_ -]?token)\s*[:=]',r'(?i)\b(passport|national insurance|date of birth|full name|email address|phone number|home address|source identity|safeguarding|confidential|restricted)\b']
def scan(text):
    if any(re.search(p,text) for p in PATTERNS): raise ValueError('Possible sensitive content; review locally.')
def check_formula(value):
    if isinstance(value,str) and value.startswith('=') and re.search(r'(?i)(https?:|file:|\\\\|\[|\]|\||WEBSERVICE|HYPERLINK|RTD|DDE|IMAGE|STOCKHISTORY)',value):
        raise ValueError('External formula refused')
def path(name):
    if not re.fullmatch(r'[a-f0-9]{32}\.xlsx',name): raise ValueError('Use a workbook ID returned by list/create.')
    p=ROOT/name
    from folder_access import no_links
    no_links(p.parent)
    if p.exists() or p.is_symlink(): no_links(p)
    return p

def load(name):
    p=path(name)
    with p.open('rb') as stream:
        data=stream.read(25*1024*1024+1)
    if len(data)>25*1024*1024: raise ValueError('Workbook exceeds 25 MiB')
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if sum(i.file_size for i in z.infolist())>50*1024*1024: raise ValueError('Expanded workbook too large')
        for n in z.namelist():
            if any(x in n.lower() for x in ['vbaproject','externallinks/','embeddings/']): raise ValueError('Macros, external links or embedded objects unsupported')
            if n.endswith(('.xml','.rels')):
                t=z.read(n).decode('utf-8',errors='strict'); scan(t)
                if '<!DOCTYPE' in t or '<!ENTITY' in t: raise ValueError('DTD refused')
    wb=openpyxl.load_workbook(io.BytesIO(data),keep_links=False)
    if sum(w.max_row*w.max_column for w in wb)>250000: raise ValueError('Workbook exceeds 250,000 cells')
    for ws in wb:
        for row in ws:
            for c in row:
                if isinstance(c.value,str): scan(c.value)
                if c.data_type=='f': check_formula(c.value)
    return wb

def frame(wb,sheet):
    rows=list(wb[sheet].values)
    if not rows: raise ValueError('No table')
    if any(x is None for x in rows[0]) or len(set(rows[0]))!=len(rows[0]): raise ValueError('Unique nonempty headers required in row 1')
    if any(isinstance(x,str) and x.startswith('=') for row in rows for x in row): raise ValueError('Table operations require values, not formulas; use locally recalculated values copy')
    return pd.DataFrame(rows[1:],columns=rows[0])

def putframe(wb,df,title):
    if title in wb.sheetnames: raise ValueError('Choose a new result sheet name')
    if df.size>250000: raise ValueError('Result too large')
    ws=wb.create_sheet(title); ws.append([str(x) for x in df.columns])
    for row in df.itertuples(index=False,name=None):
        values=[None if pd.isna(v) else v for v in row]; ws.append(values)
        for c in ws[ws.max_row]:
            if isinstance(c.value,str) and c.value.startswith('='): c.data_type='s'
    ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
    for c in ws[1]:c.font=Font(bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='194B6B')

def save(wb):
    ROOT.mkdir(exist_ok=True); name=uuid.uuid4().hex+'.xlsx'
    wb.calculation=openpyxl.workbook.properties.CalcProperties(fullCalcOnLoad=True)
    p=path(name);wb.save(p)
    try: load(name)
    except Exception: p.unlink();raise
    return {'workbook':name,'saved_locally':True,'formula_note':'Excel must recalculate formulas on opening. Original preserved; each edit creates a new version.'}

def execute(r):
    scan(json.dumps(r,default=str))
    op=r['operation']
    if op in ['sharepoint_status','sharepoint_browse','sharepoint_import']:
        from sharepoint import execute as sharepoint_execute
        return sharepoint_execute(r,ROOT,load,scan)
    if op in ['browse','import_file']:
        from folder_access import execute as folder_execute
        return folder_execute(r,ROOT,load,scan)
    if op=='list':
        ROOT.mkdir(exist_ok=True); result=[]
        for p in ROOT.glob('*.xlsx'):
            try:w=load(p.name);result.append({'workbook':p.name,'sheets':w.sheetnames})
            except Exception:result.append({'workbook':p.name,'blocked':True})
        return result
    if op=='create':
        w=openpyxl.Workbook();w.active.title=r.get('sheet','Data')
        for row in r.get('rows',[]): w.active.append(row)
        return save(w)
    w=load(r['workbook']); sheet=r.get('sheet',w.sheetnames[0]);ws=w[sheet]
    if op=='inspect':return {'sheets':[{'name':s.title,'rows':s.max_row,'columns':s.max_column,'state':s.sheet_state} for s in w]}
    if op=='read':
        a,b,c,d=range_boundaries(r.get('range','A1:J20'))
        if not all([a,b,c,d]) or (c-a+1)*(d-b+1)>5000:raise ValueError('Read at most 5000 cells per call')
        return [[x for x in row] for row in ws.iter_rows(min_row=b,max_row=d,min_col=a,max_col=c,values_only=True)]
    if op in ['dashboard','fill_template']:
        from finance import apply
        report=apply(w,r,load,frame)
        result=save(w);result['report']=report;return result
    if op=='write':
        for coordinate,value in r['cells'].items():
            if not re.fullmatch('[A-Z]{1,3}[1-9][0-9]{0,5}',coordinate):raise ValueError('Invalid cell')
            check_formula(value)
            target=ws[coordinate]
            if target.data_type=='f' and not r.get('replace_formulas',False):raise ValueError('Formula target requires explicit replace_formulas approval')
            if ws.protection.sheet and target.protection.locked:raise ValueError('Protected target')
            target.value=value
    elif op=='add_sheet':w.create_sheet(r['name'])
    elif op=='format':
        a,b,c,d=range_boundaries(r['range'])
        if not all([a,b,c,d]) or (c-a+1)*(d-b+1)>250000:raise ValueError('Range too large')
        for row in ws.iter_rows(min_row=b,max_row=d,min_col=a,max_col=c):
            for cell in row:
                if 'number_format' in r:cell.number_format=r['number_format']
                if 'bold' in r:cell.font=Font(bold=bool(r['bold']))
                if 'fill' in r:cell.fill=PatternFill('solid',fgColor=r['fill'])
        if 'freeze' in r:ws.freeze_panes=r['freeze']
        if r.get('autofilter'):ws.auto_filter.ref=r['range']
    elif op=='chart':
        chart={'bar':BarChart,'line':LineChart,'pie':PieChart}[r.get('kind','bar')]()
        chart.title=r.get('title','Chart')
        chart.add_data(Reference(ws,range_string=f"'{sheet}'!{r['data_range']}"),titles_from_data=True)
        chart.set_categories(Reference(ws,range_string=f"'{sheet}'!{r['category_range']}"));ws.add_chart(chart,r.get('anchor','H2'))
    elif op in ['clean','sort','filter','aggregate','join','append']:
        df=frame(w,sheet)
        if op=='clean':
            if r.get('trim',True):df=df.map(lambda x:x.strip() if isinstance(x,str) else x)
            if r.get('drop_empty',True):df=df.replace('',None).dropna(how='all')
            if r.get('deduplicate',True):df=df.drop_duplicates(subset=r.get('keys'))
            if r.get('rename'):df=df.rename(columns=r['rename'])
        elif op=='sort':df=df.sort_values(r['columns'],ascending=r.get('ascending',True),kind='stable')
        elif op=='filter':
            col=df[r['column']];v=r.get('value');operator=r.get('comparison','eq')
            mask={'eq':lambda:col.eq(v),'ne':lambda:col.ne(v),'gt':lambda:col.gt(v),'lt':lambda:col.lt(v),'contains':lambda:col.astype(str).str.contains(str(v),regex=False,na=False),'notnull':lambda:col.notna()}[operator]()
            df=df[mask]
        elif op=='aggregate':
            func=r.get('function','sum')
            if func not in ['sum','mean','count','min','max','median']:raise ValueError('Unsupported aggregate')
            df=df.groupby(r['group_by'],dropna=False)[r['values']].agg(func).reset_index()
        else:
            other=frame(load(r.get('other_workbook',r['workbook'])),r['other_sheet'])
            if op=='append':df=pd.concat([df,other],ignore_index=True)
            else:
                how=r.get('how','left')
                if how not in ['left','right','inner','outer']:raise ValueError('Unsupported join')
                validate=r.get('validate','many_to_one')
                if validate not in ['one_to_one','one_to_many','many_to_one']:raise ValueError('Many-to-many financial joins are unsupported')
                df=df.merge(other,on=r['keys'],how=how,validate=validate)
        putframe(w,df,r.get('result_sheet','Result'))
    else:raise ValueError('Unknown operation')
    return save(w)

DESCRIPTION='''SharePoint operations (requires local Connect-SharePoint.cmd setup and Microsoft sign-in): sharepoint_status() checks connection without revealing account details. sharepoint_browse(path?:relative folder, page?:opaque next_page token) lists approved SharePoint entries; default path empty string. sharepoint_import(path:exact relative file OR url:direct SharePoint file URL) downloads and scans locally, returning a workbook ID for XLSX or prepared_file for a supported document. Read prepared_file using scanned Read in Approved-Data. Browse instead of guessing names. Short opaque sharing links are unsupported. Imports preserve remote originals; there are no upload, edit-remote, permission or sharing operations. Account tokens, settings, raw download URLs and unscanned content must never be requested. SP-SETUP/SP-LOGIN: ask user to run Connect-SharePoint.cmd locally, never try another tool. SP-ACCESS: ask IT to verify the configured site grant and user access. Folder operations: browse(path?:relative folder, empty string for root) lists up to 500 direct entries; browse returned folder paths to explore subfolders. import_file(path:relative XLSX file) copies and scans a workbook, returning its workbook ID. Source root is chosen by user in Select working folder shortcut, never set by the model. Only approved source folders, no links or path escapes. Imported copies can then be edited and joined. Finance operations: dashboard(workbook,sheet,group_by:[columns],budget_column,actual_column,currency:ISO code,currency_column?,result_sheet?). Amounts must be numeric with at most two decimals and one declared currency. Produces calculated KPI totals, reconciliation, grouped variances, formatted table and chart. fill_template(workbook:template ID,source_workbook?,mappings:[{target_sheet,target_cell,source_workbook?,source_sheet,source_column,where:{column:value},aggregation:single/sum/count,replace_existing?:false}]). Missing or ambiguous matches, protected targets and formula overwrite STOP the whole operation. All transformations preserve originals. Local Excel workflow on approved XLSX copies. Request.operation: list; create(sheet,rows); inspect(workbook); read(workbook,sheet,range); write(workbook,sheet,cells:{A1:value}); add_sheet(workbook,name); format(workbook,sheet,range,number_format?,bold?,fill?,freeze?,autofilter?); chart(workbook,sheet,kind:bar/line/pie,data_range,category_range,anchor?,title?); clean(workbook,sheet,trim?,drop_empty?,deduplicate?,keys?,rename?,result_sheet); sort(workbook,sheet,columns,ascending?,result_sheet); filter(workbook,sheet,column,comparison:eq/ne/gt/lt/contains/notnull,value?,result_sheet); aggregate(workbook,sheet,group_by:[columns],values:[columns],function:sum/mean/count/min/max/median,result_sheet); join(workbook,sheet,other_workbook?,other_sheet,keys,how?,validate?,result_sheet); append(workbook,sheet,other_workbook?,other_sheet,result_sheet). All mutations return a NEW workbook ID: use it in the next step. Files saved in local Workbooks folder. Tables require unique first-row headers and literal values. Formulas are written, not calculated here. Existing advanced Excel features may not be preserved. Never claim a passing privacy scan makes data anonymous.'''
if __name__=='__main__':
    import sys
    if len(sys.argv)==3 and sys.argv[1]=='--check':
        try: load(sys.argv[2]); sys.exit(0)
        except Exception: sys.exit(1)
    from mcp.server.fastmcp import FastMCP
    mcp=FastMCP('claude_safeguard')
    @mcp.tool(description=DESCRIPTION)
    def spreadsheet(request:dict)->str:
        try:
            result=execute(request);text=json.dumps(result,default=str,ensure_ascii=False);scan(text)
            if len(text)>100000:raise ValueError('Response too large')
            return text
        except Exception:
            return json.dumps({'error':'Operation stopped: possible sensitive content, unsupported feature, invalid request or limits exceeded. Review locally; no source values included in this error.'})
    mcp.run(transport='stdio')
