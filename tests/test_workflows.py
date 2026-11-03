"""Synthetic workbook workflows, including the actual MCP stdio protocol."""
from pathlib import Path
import asyncio
import sys
import tempfile
import unittest
from unittest.mock import patch
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import spreadsheet as sheets
import runtime_check

class ProtocolTests(unittest.TestCase):
    def test_mcp_initialize_list_create_dashboard_and_refusal(self):
        runtime_check.imports()
        asyncio.run(runtime_check.smoke())

class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='safeguard finance ')
        self.root = Path(self.temp.name)
        self.patch = patch.object(sheets, 'ROOT', self.root)
        self.patch.start()
    def tearDown(self):
        self.patch.stop()
        self.temp.cleanup()
    def run_op(self, operation, **kwargs):
        return sheets.execute(dict(operation=operation, **kwargs))
    def source(self):
        return self.run_op('create', sheet='Actuals', rows=[['Department','Budget','Actual','Currency'],
            ['Newsroom',1000,1100,'GBP'],['Operations',2000,1800,'GBP'],['Newsroom',500,400,'GBP']])['workbook']
    def test_dashboard_template_across_files_and_original_preservation(self):
        source = self.source()
        original = sheets.path(source).read_bytes()
        result = self.run_op('dashboard', workbook=source, sheet='Actuals', group_by=['Department'],
                             budget_column='Budget', actual_column='Actual', currency='GBP', currency_column='Currency')
        self.assertEqual({k: result['report'][k] for k in ['budget','actual','variance']},
                         {'budget':'3500','actual':'3300','variance':'-200'})
        self.assertEqual(len(sheets.load(result['workbook'])['Dashboard']._charts), 1)
        template = self.run_op('create', sheet='Report', rows=[['Department','Actual','Check'],
                 ['Newsroom',None,'=B2'],['Operations',None,None]])['workbook']
        book = sheets.load(template)
        book['Report']['B2'].number_format = '#,##0.00'
        styled = sheets.save(book)['workbook']
        out = self.run_op('fill_template', workbook=styled, source_workbook=source, mappings=[
            {'target_sheet':'Report','target_cell':'B2','source_sheet':'Actuals','source_column':'Actual',
             'where':{'Department':'Newsroom'},'aggregation':'sum'},
            {'target_sheet':'Report','target_cell':'B3','source_sheet':'Actuals','source_column':'Actual',
             'where':{'Department':'Operations'},'aggregation':'single'}])
        report = sheets.load(out['workbook'])['Report']
        self.assertEqual((report['B2'].value, report['B3'].value), (1500,1800))
        self.assertEqual(report['C2'].value, '=B2')
        self.assertEqual(report['B2'].number_format, '#,##0.00')
        self.assertEqual(sheets.path(source).read_bytes(), original)
        self.assertIsNone(sheets.load(styled)['Report']['B2'].value)
    def test_fill_stops_atomically_on_second_invalid_mapping(self):
        source = self.source()
        template = self.run_op('create', sheet='Report', rows=[['Report'],[None,None,'=1+1']])['workbook']
        first = dict(target_sheet='Report',target_cell='B2',source_sheet='Actuals',source_column='Actual',
                     where={'Department':'Operations'},aggregation='single')
        for change in [dict(where={'Department':'Newsroom'}), dict(target_cell='C2'),dict(where={'Department':'Absent'})]:
            second = dict(first, target_cell='B3')
            second.update(change)
            before = set(self.root.iterdir())
            with self.assertRaises(ValueError):
                self.run_op('fill_template',workbook=template,source_workbook=source,mappings=[first,second])
            self.assertEqual(before,set(self.root.iterdir()))
            self.assertIsNone(sheets.load(template)['Report']['B2'].value)
    def test_missing_category_amount_mixed_currency_and_precision_refused(self):
        for rows in [
            [['A',1,2,'GBP'],[None,1,2,'GBP']],
            [['A',1,None,'GBP']],
            [['A',1,2,'GBP'],['B',1,2,'USD']],
            [['A',1,2.001,'GBP']], []]:
            source=self.run_op('create',sheet='Data',rows=[['Department','Budget','Actual','Currency']]+rows)['workbook']
            before=set(self.root.iterdir())
            with self.assertRaises(ValueError):
                self.run_op('dashboard',workbook=source,sheet='Data',group_by=['Department'],budget_column='Budget',
                            actual_column='Actual',currency='GBP',currency_column='Currency')
            self.assertEqual(before,set(self.root.iterdir()))
    def test_clean_join_aggregate_chart_filter_sort_append(self):
        source=self.run_op('create',sheet='Data',rows=[['Month','Views'],['Jan',10],['Feb',20],['Jan',10]])['workbook']
        clean=self.run_op('clean',workbook=source,sheet='Data',result_sheet='Clean')['workbook']
        self.assertEqual(sheets.load(clean)['Clean'].max_row,3)
        other=self.run_op('create',sheet='Lookup',rows=[['Month','Group'],['Jan','A'],['Feb','B']])['workbook']
        joined=self.run_op('join',workbook=source,sheet='Data',other_workbook=other,other_sheet='Lookup',keys=['Month'],result_sheet='Joined')['workbook']
        self.assertEqual(sheets.load(joined)['Joined']['C2'].value,'A')
        aggregate=self.run_op('aggregate',workbook=joined,sheet='Joined',group_by=['Group'],values=['Views'],result_sheet='Summary')['workbook']
        self.assertEqual(sheets.load(aggregate)['Summary']['B2'].value,20)
        chart=self.run_op('chart',workbook=aggregate,sheet='Summary',data_range='B1:B3',category_range='A2:A3')['workbook']
        self.assertEqual(len(sheets.load(chart)['Summary']._charts),1)
        filtered=self.run_op('filter',workbook=source,sheet='Data',column='Views',comparison='gt',value=15,result_sheet='Filtered')['workbook']
        self.assertEqual(sheets.load(filtered)['Filtered']['A2'].value,'Feb')
        sorted_book=self.run_op('sort',workbook=source,sheet='Data',columns=['Views'],ascending=False,result_sheet='Sorted')['workbook']
        self.assertEqual(sheets.load(sorted_book)['Sorted']['B2'].value,20)
        appended=self.run_op('append',workbook=source,sheet='Data',other_sheet='Data',result_sheet='Appended')['workbook']
        self.assertEqual(sheets.load(appended)['Appended'].max_row,7)
        self.assertEqual(sheets.load(source)['Data'].max_row,4)
    def test_external_formula_refused_on_create_and_import(self):
        import openpyxl
        import uuid
        before=set(self.root.iterdir())
        with self.assertRaises(ValueError):
            self.run_op('create',rows=[['=WEBSERVICE("https://example.invalid")']])
        self.assertEqual(before,set(self.root.iterdir()))
        book=openpyxl.Workbook()
        book.active['A1']='=HYPERLINK("https://example.invalid","Open")'
        name=uuid.uuid4().hex+'.xlsx';book.save(self.root/name)
        with self.assertRaises(ValueError): sheets.load(name)
    def test_many_to_many_join_refused(self):
        source=self.source()
        with self.assertRaises(ValueError):
            self.run_op('join',workbook=source,sheet='Actuals',other_sheet='Actuals',
                        keys=['Department'],validate='many_to_many',result_sheet='Join')

if __name__ == '__main__': unittest.main(verbosity=2)
