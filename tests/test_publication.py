import importlib.util
import json
from pathlib import Path
import re
import shutil
import tempfile
import unittest
import openpyxl

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('publisher', ROOT/'scripts/build_site.py')
publisher = importlib.util.module_from_spec(spec)
spec.loader.exec_module(publisher)


class PublicationTests(unittest.TestCase):
    def prepare(self, root):
        for name in ('HK_Mortgage_Dashboard.html', 'hk_mortgage_data.xlsx'):
            shutil.copy2(ROOT/name, root/name)

    def test_published_excel_and_html_match_exactly(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            self.prepare(root)
            # A secret placed beside the workbook must never enter the website.
            (root/'.dashboard-sync.json').write_text('private configuration')
            publisher.build_site(root)
            site=root/'site'
            self.assertEqual({p.name for p in site.iterdir()}, {'index.html','hk_mortgage_data.xlsx','.nojekyll'})
            self.assertEqual((site/'hk_mortgage_data.xlsx').read_bytes(), (root/'hk_mortgage_data.xlsx').read_bytes())
            self.assertEqual((site/'index.html').read_bytes(), (root/'HK_Mortgage_Dashboard.html').read_bytes())

    def test_mismatch_cannot_overwrite_existing_site(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            self.prepare(root)
            site=root/'site'
            site.mkdir()
            (site/'index.html').write_text('previous verified site')
            book=openpyxl.load_workbook(root/'hk_mortgage_data.xlsx')
            book.worksheets[0].cell(3,2).value=123456
            book.save(root/'hk_mortgage_data.xlsx')
            book.close()
            with self.assertRaises(ValueError): publisher.build_site(root)
            self.assertEqual((site/'index.html').read_text(),'previous verified site')

    def test_changed_values_are_rejected_even_with_copied_revision(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            self.prepare(root)
            html=root/'HK_Mortgage_Dashboard.html'
            text=html.read_text(encoding='utf-8')
            pattern=r'(<script id="workbook-data" type="application/json">)(.*?)(</script>)'
            match=re.search(pattern,text,re.S)
            data=json.loads(match[2])
            data['series']['registrations'][0][1]=123456
            html.write_text(text[:match.start(2)]+json.dumps(data)+text[match.end(2):],encoding='utf-8')
            with self.assertRaises(ValueError): publisher.build_site(root)
            self.assertFalse((root/'site').exists())


if __name__=='__main__': unittest.main()
