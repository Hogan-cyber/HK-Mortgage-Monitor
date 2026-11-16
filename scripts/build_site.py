"""Publish only the reviewed dashboard and its exact saved workbook."""
import hashlib
import json
from pathlib import Path
import re
import shutil
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from update_mortgage_data import dashboard_snapshot


def build_site(root=ROOT):
    root=Path(root)
    workbook=root/'hk_mortgage_data.xlsx'
    dashboard=root/'HK_Mortgage_Dashboard.html'
    html=dashboard.read_text(encoding='utf-8')
    matches=re.findall(r'<script id="workbook-data" type="application/json">(.*?)</script>',html,re.S)
    if len(matches)!=1:
        raise ValueError('The dashboard must contain exactly one workbook snapshot.')
    embedded=json.loads(matches[0])
    expected=dashboard_snapshot(workbook)
    if embedded.get('revision')!=hashlib.sha256(workbook.read_bytes()).hexdigest():
        raise ValueError('Dashboard is not generated from the current saved Excel workbook.')
    if embedded.get('schemaVersion')!=expected['schemaVersion'] or embedded.get('series')!=expected['series']:
        raise ValueError('Dashboard values differ from the saved Excel workbook.')
    if re.search(r'<(?:script|link)[^>]+(?:src|href)=["\'](?:https?:|/)',html,re.I):
        raise ValueError('The dashboard must be self-contained.')
    site=root/'site'
    site.mkdir(exist_ok=True)
    # The site is build output. Never publish the repository root or local caches.
    for entry in site.iterdir():
        if entry.name not in {'index.html','hk_mortgage_data.xlsx','.nojekyll'}:
            raise ValueError(f'Unexpected file in publication output: {entry.name}')
    shutil.copyfile(workbook,site/'hk_mortgage_data.xlsx')
    shutil.copyfile(dashboard,site/'index.html')
    (site/'.nojekyll').write_text('',encoding='utf-8')
    print('Verified website:',site)
    print('Workbook revision:',embedded['revision'])
    return site


if __name__=='__main__':
    build_site()
