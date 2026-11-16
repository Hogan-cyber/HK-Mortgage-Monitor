import importlib.util
import io
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('offline_updater', ROOT / 'update_mortgage_data.py')
updater = importlib.util.module_from_spec(spec)
spec.loader.exec_module(updater)
PATTERN = r'<script id="workbook-data" type="application/json">(.*?)</script>'


class OfflineDashboardTests(unittest.TestCase):
    def test_imported_writer_handles_a_western_console(self):
        with tempfile.TemporaryDirectory() as folder:
            base=Path(folder)
            self.copy_bundle(base)
            output=io.BytesIO()
            stream=io.TextIOWrapper(output,encoding='cp1252',errors='strict')
            with patch.object(updater,'BASE',base), patch.object(updater.sys,'stdout',stream):
                updater.update_dashboard(base/'hk_mortgage_data.xlsx')
                stream.flush()
                self.assertIn('已更新',output.getvalue().decode('utf-8'))

    def test_old_but_importable_excel_library_is_rejected(self):
        actual_version = updater.importlib.metadata.version
        def versions(name):
            return '3.0.10' if name == 'openpyxl' else actual_version(name)
        with patch.object(updater.importlib.metadata, 'version', side_effect=versions):
            with self.assertRaises(ImportError):
                updater.check_dependencies()

    def copy_bundle(self, base):
        for name in ['update_mortgage_data.py','hk_mortgage_data.xlsx','HK_Mortgage_Dashboard.html']:
            shutil.copy2(ROOT / name, base / name)

    def test_render_preserves_excel_and_embeds_every_reviewed_value(self):
        self.assertTrue(hasattr(updater, 'update_dashboard'), 'Standalone dashboard writer is missing')
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            self.copy_bundle(base)
            book = base / 'hk_mortgage_data.xlsx'
            original = book.read_bytes()
            expected = updater.dashboard_snapshot(book)
            with patch.object(updater, 'BASE', base), patch.object(updater.urllib.request, 'build_opener', side_effect=AssertionError('No network allowed')):
                updater.update_dashboard(book)
            html = (base / 'HK_Mortgage_Dashboard.html').read_text(encoding="utf-8")
            payload = json.loads(re.search(PATTERN, html, re.S).group(1))
            self.assertEqual(payload['series'], expected['series'])
            self.assertEqual(payload['revision'], expected['revision'])
            self.assertEqual(book.read_bytes(), original)
            self.assertIn("connect-src 'none'", html)
            self.assertNotRegex(html, r'<(?:script|link)[^>]+(?:src|href)=["\'](?:https?:|/)')
            self.assertNotIn('chatgpt.site', html)
            self.assertNotIn('netlify.app', html)

    def test_copied_folder_works_from_another_working_directory_without_network(self):
        self.assertTrue(hasattr(updater, 'update_dashboard'), 'Standalone dashboard writer is missing')
        with tempfile.TemporaryDirectory(prefix='Mortgage 同事 folder ') as folder:
            base = Path(folder)
            self.copy_bundle(base)
            guard = "import socket,runpy,sys\ndef deny(*a,**kw): raise RuntimeError('NETWORK BLOCKED')\nsocket.socket.connect=deny\nsys.argv=[sys.argv[1],'--dashboard-only']\nrunpy.run_path(sys.argv[0],run_name='__main__')"
            result = subprocess.run([sys.executable,'-c',guard,str(base/'update_mortgage_data.py')],
                                    cwd=base.parent,capture_output=True,text=True,encoding="utf-8")
            self.assertEqual(result.returncode,0,result.stderr)
            self.assertTrue((base/'dashboard_status.json').is_file())

    def test_invalid_workbook_cannot_replace_the_last_good_dashboard(self):
        self.assertTrue(hasattr(updater, 'update_dashboard'), 'Standalone dashboard writer is missing')
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            self.copy_bundle(base)
            previous = (base/'HK_Mortgage_Dashboard.html').read_bytes()
            book = base/'hk_mortgage_data.xlsx'
            book.write_bytes(b'not a workbook')
            with patch.object(updater,'BASE',base), self.assertRaises(Exception):
                updater.update_dashboard(book)
            self.assertEqual((base/'HK_Mortgage_Dashboard.html').read_bytes(), previous)

    def test_failed_file_replace_keeps_last_good_dashboard(self):
        self.assertTrue(hasattr(updater, 'update_dashboard'), 'Standalone dashboard writer is missing')
        with tempfile.TemporaryDirectory() as folder:
            base = Path(folder)
            self.copy_bundle(base)
            previous = (base/'HK_Mortgage_Dashboard.html').read_bytes()
            with patch.object(updater,'BASE',base), patch.object(updater.os,'replace',side_effect=PermissionError('File locked')), self.assertRaises(Exception):
                updater.update_dashboard(base/'hk_mortgage_data.xlsx')
            self.assertEqual((base/'HK_Mortgage_Dashboard.html').read_bytes(), previous)

    def test_offline_recalculation_updates_both_workbook_and_dashboard(self):
        self.assertTrue(hasattr(updater, 'update_dashboard'), 'Standalone dashboard writer is missing')
        with tempfile.TemporaryDirectory(prefix='Mortgage 移动 folder ') as folder:
            base=Path(folder)
            self.copy_bundle(base)
            result=subprocess.run([sys.executable,str(base/'update_mortgage_data.py'),'--recalculate-only'],cwd=base.parent,capture_output=True,text=True,encoding="utf-8")
            self.assertEqual(result.returncode,0,result.stderr)
            html=(base/'HK_Mortgage_Dashboard.html').read_text(encoding="utf-8")
            rows=json.loads(re.search(PATTERN,html,re.S).group(1))['series']['nonbank']
            october=next(row for row in rows if row[0]=='2026-10')
            source_rows = updater.dashboard_snapshot(ROOT/'hk_mortgage_data.xlsx')['series']['nonbank']
            expected = next(row for row in source_rows if row[0]=='2026-10')
            self.assertEqual(october, expected)
            self.assertTrue((base/'hk_mortgage_data.previous.xlsx').is_file())


if __name__ == '__main__':
    unittest.main()
