from pathlib import Path
import json
import os
import sys
import tempfile
import unittest
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import folder_access as folders
import spreadsheet as sheets

class FolderTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(prefix='safeguard folder ')
        self.root=Path(self.temp.name)
        self.source=self.root/'source';self.source.mkdir()
        self.books=self.root/'Workbooks';self.books.mkdir()
        config=self.root/'approved-folder.json'
        config.write_text(json.dumps({'root':str(self.source)}),encoding='utf-8-sig')
        self.patches=[patch.object(folders,'CONFIG',config),patch.object(sheets,'ROOT',self.books)]
        for p in self.patches:p.start()
    def tearDown(self):
        for p in reversed(self.patches):p.stop()
        self.temp.cleanup()
    def test_browse_nested_import_leaves_source_unchanged(self):
        nested=self.source/'month';nested.mkdir()
        book=sheets.execute({'operation':'create','rows':[['Amount'],[50]]})['workbook']
        content=(self.books/book).read_bytes()
        (nested/'source.xlsx').write_bytes(content)
        listed=sheets.execute({'operation':'browse','path':'month'})
        self.assertEqual(listed['entries'][0]['path'],'month/source.xlsx')
        imported=sheets.execute({'operation':'import_file','path':'month/source.xlsx'})
        self.assertEqual(sheets.load(imported['workbook']).active['A2'].value,50)
        self.assertEqual((nested/'source.xlsx').read_bytes(),content)
    def test_escape_absolute_and_network_paths_refused(self):
        for value in ['../other','/other','C:\\other','\\\\host\\share','nested/../../other']:
            with self.assertRaises((ValueError,FileNotFoundError)):folders.scoped(value)
    def test_sensitive_name_withheld(self):
        (self.source/'person@example.org.xlsx').write_bytes(b'')
        result=sheets.execute({'operation':'browse'})
        self.assertNotIn('person@example.org',json.dumps(result))
        self.assertEqual(result['entries'][0]['type'],'blocked')
    def test_sensitive_workbook_not_retained(self):
        import openpyxl
        book=openpyxl.Workbook();book.active['A1']='person@example.org'
        book.save(self.source/'source.xlsx')
        with self.assertRaises(ValueError):sheets.execute({'operation':'import_file','path':'source.xlsx'})
        self.assertEqual(list(self.books.iterdir()),[])
    @unittest.skipUnless(os.name=='posix','Symlink fixture needs Unix')
    def test_linked_source_refused(self):
        (self.source/'link').symlink_to(self.root,target_is_directory=True)
        with self.assertRaises(ValueError):folders.scoped('link')

if __name__=='__main__':unittest.main(verbosity=2)
