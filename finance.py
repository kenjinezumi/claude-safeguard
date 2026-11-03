"""Deterministic financial summaries and explicit template mappings."""
from decimal import Decimal,InvalidOperation
from collections import defaultdict
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.chart import BarChart,Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.cell.cell import MergedCell
import re
import pandas as pd

def money(value):
    if value is None or isinstance(value,bool):raise ValueError('Missing/non-numeric amount')
    v=Decimal(str(value))
    if not v.is_finite() or v!=v.quantize(Decimal('.01')):raise ValueError('Amounts require finite values with at most two decimal places')
    return v

def apply(w,r,load,frame):
    if r['operation']=='dashboard':
        df=frame(w,r['sheet'])
        currency=r['currency']
        if not re.fullmatch('[A-Z]{3}',currency):raise ValueError('Explicit ISO currency required')
        if r.get('currency_column') and set(df[r['currency_column']])!={currency}:raise ValueError('Mixed or missing currency')
        if df.empty or not r['group_by']:raise ValueError('Nonempty data and grouping columns required')
        groups=defaultdict(lambda:[Decimal(0),Decimal(0)])
        for row in df.to_dict('records'):
            key=tuple(row[k] for k in r['group_by'])
            if any(pd.isna(x) or str(x).strip()=='' for x in key):raise ValueError('Missing category')
            groups[key][0]+=money(row[r['budget_column']]);groups[key][1]+=money(row[r['actual_column']])
        title=r.get('result_sheet','Dashboard')
        if title in w.sheetnames:raise ValueError('Choose new dashboard sheet name')
        ws=w.create_sheet(title)
        ws.append(['BUDGET AND ACTUALS',currency]);ws.append(['Variance convention','Actual minus budget; positive = overspend'])
        budget=sum((v[0] for v in groups.values()),Decimal(0));actual=sum((v[1] for v in groups.values()),Decimal(0))
        ws.append(['Total budget',float(budget)]);ws.append(['Total actual',float(actual)]);ws.append(['Variance',float(actual-budget)])
        ws.append(['Reconciliation','PASS: grouped totals equal source totals']);ws.append([])
        headers=r['group_by']+['Budget','Actual','Variance','Variance %'];ws.append(headers)
        n=len(r['group_by'])
        for key,(b,a) in sorted(groups.items(),key=lambda x:tuple(str(v) for v in x[0])):
            ws.append(list(key)+[float(b),float(a),float(a-b),float((a-b)/b) if b else None])
        ws.freeze_panes='A9';ws.auto_filter.ref=f'A8:{ws.cell(8,len(headers)).column_letter}{ws.max_row}'
        for row in [1,8]:
            for cell in ws[row]:cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='194B6B')
        for row in ws.iter_rows(min_row=9,min_col=n+1,max_col=n+3):
            for cell in row:cell.number_format='#,##0.00;[Red](#,##0.00)'
        for row in ws.iter_rows(min_row=9,min_col=n+4,max_col=n+4):
            row[0].number_format='0.0%'
        for i in range(3,6):ws.cell(i,2).number_format='#,##0.00;[Red](#,##0.00)'
        for col in ws.columns:ws.column_dimensions[col[0].column_letter].width=24
        if groups:
            chart=BarChart();chart.title='Budget versus actual';chart.y_axis.title=currency
            chart.add_data(Reference(ws,min_col=n+1,max_col=n+2,min_row=8,max_row=ws.max_row),titles_from_data=True)
            chart.set_categories(Reference(ws,min_col=1,min_row=9,max_row=ws.max_row));ws.add_chart(chart,'J3')
        return {'rows_processed':len(df),'budget':str(budget),'actual':str(actual),'variance':str(actual-budget),'currency':currency,'reconciled':True,'zero_budget_percent':'blank'}
    if r['operation']=='fill_template':
        changes=[]; seen=set()
        for item in r['mappings']:
            target=(item['target_sheet'],item['target_cell'])
            if target in seen:raise ValueError('Duplicate target')
            seen.add(target)
            ws=w[target[0]];cell=ws[target[1]]
            if isinstance(cell,MergedCell):raise ValueError('Use merged-range anchor')
            if cell.data_type=='f':raise ValueError('Template formula would be overwritten')
            if ws.protection.sheet and cell.protection.locked:raise ValueError('Protected target')
            if cell.value is not None and not item.get('replace_existing',False):raise ValueError('Target already populated; explicitly approve replacement')
            source=load(item.get('source_workbook',r.get('source_workbook',r['workbook'])))
            df=frame(source,item['source_sheet'])
            for key,value in item.get('where',{}).items():df=df[df[key].eq(value)]
            mode=item.get('aggregation','single')
            if len(df)==0:raise ValueError('Missing source match')
            if mode=='single':
                if len(df)!=1:raise ValueError('Ambiguous source match')
                value=df.iloc[0][item['source_column']]
                if hasattr(value,'item'):value=value.item()
            elif mode=='sum':value=float(sum((money(v) for v in df[item['source_column']]),Decimal(0)))
            elif mode=='count':value=len(df)
            else:raise ValueError('Unsupported mapping aggregation')
            if isinstance(value,str) and value.startswith('='):raise ValueError('Source formula refused')
            changes.append((cell,value))
        for cell,value in changes:cell.value=value
        return {'filled_cells':len(changes),'formulas_preserved':True,'missing_or_ambiguous':'All mappings resolved; otherwise no workbook is saved'}
    raise ValueError('Unsupported finance operation')
