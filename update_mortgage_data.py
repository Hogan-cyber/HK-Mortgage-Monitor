# -*- coding: utf-8 -*-
"""
香港住宅按揭数据自动采集与整合脚本
===================================
可每天运行，自动更新按揭、货币供应及利率，整合进同一个 Excel 文件：

  [1] 中原按揭 (Centaline Mortgage) — 现楼/楼花按揭登记宗数 (可追溯至 2012 年)
      来源: https://api.centamortgage.com/api/Infomation/GetStatistics
      (数据口径: 基于土地注册处登记记录, 由中原按揭研究部拆分出「现楼按揭」与「楼花按揭」)

  [2] HKMA 住宅按揭统计调查 — 新申请/新批出(一手/二手/转按)/新取用 宗数 + 金额
      来源: HKMA Monthly Statistical Bulletin 表 T0307 (官方 Excel)
      URL : https://www.hkma.gov.hk/media/eng/doc/market-data-and-statistics/monthly-statistical-bulletin/T0307.xlsx
      (只覆盖认可机构=银行, 财务公司/放债人不在统计范围)

  [3] 土地注册处 (Land Registry) — 楼宇按揭/其他按揭(法定押记)/撤销按揭 登记宗数
      来源: https://www.landreg.gov.hk/json/monthly_stat/stat{MM}{YY}/t1.json
      (官方原始口径, 不区分住宅/非住宅, 且不公布按揭金额)

  [4] M3 — HKMA T020201.xls，港元及所有货币，月度，HK$ million
  [5] HSBC Prime — HKMA T060401.xls 历史及 HSBC 官网最新生效利率
  [6] HIBOR — HKMA T060303.xls 历史及 HKAB 每日 11:15 定盘，跳过休市日

  [7] 非银行 — 延续 Excel 原有公式，按登记月与上月银行买卖放款估算配对。

输出: hk_mortgage_data.xlsx (多 sheet)

运行: python update_mortgage_data.py
可选参数: python update_mortgage_data.py --landreg-months 36   # 土地注册处回溯月数
更新数据和本地 Dashboard: python update_mortgage_data.py
仅用现有数据重算公式和 Dashboard: python update_mortgage_data.py --recalculate-only
仅从已保存 Excel 更新 Dashboard: python update_mortgage_data.py --dashboard-only
只更新 Excel: python update_mortgage_data.py --no-dashboard
Dashboard 为独立 HTML，不需要账号、服务器或联网。抓取新数据需要联网。
"""

import json
import os
import re
import sys
import argparse
import urllib.request
import urllib.parse
import subprocess
import shutil
import time
import hashlib
import gzip
import math
import importlib
import importlib.metadata
import platform
from pathlib import Path
from html import unescape
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from copy import copy
from datetime import timedelta, timezone, date
import io
import calendar
import tempfile
from zipfile import ZipFile
from xml.etree import ElementTree


def configure_console():
    """Keep Chinese output usable in Windows consoles and redirected task logs."""
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            try:
                stream.reconfigure(encoding='utf-8', errors='backslashreplace')
            except (OSError, ValueError):
                pass  # Some IDE consoles manage their own encoding.


if __name__ == '__main__':
    configure_console()

HK_TZ = timezone(timedelta(hours=8))
def hk_now():
    return datetime.now(HK_TZ)

DEPENDENCIES = ('openpyxl>=3.1,<4', 'xlrd>=2.0.1,<3',
                'beautifulsoup4>=4.12,<5', 'pypdf>=5,<7')
DEPENDENCY_VERSIONS = (
    ('openpyxl', 'openpyxl', (3, 1, 0), (4, 0, 0)),
    ('xlrd', 'xlrd', (2, 0, 1), (3, 0, 0)),
    ('bs4', 'beautifulsoup4', (4, 12, 0), (5, 0, 0)),
    ('pypdf', 'pypdf', (5, 0, 0), (7, 0, 0)),
)


def check_dependencies():
    """Reject old or unsupported installed packages before touching the workbook."""
    for module, distribution, minimum, maximum in DEPENDENCY_VERSIONS:
        importlib.import_module(module)
        raw_version = importlib.metadata.version(distribution)
        match = re.fullmatch(r'(\d+)\.(\d+)(?:\.(\d+))?(?:\.post\d+)?', raw_version)
        version = tuple(int(part or 0) for part in match.groups()) if match else ()
        if not minimum <= version < maximum:
            raise ImportError(f'{distribution} {raw_version} 不在支持的版本范围内。')


def ensure_dependencies():
    """Prepare a machine-specific local Python environment only when needed."""
    if sys.version_info < (3, 10):
        raise RuntimeError('更新数据需要 Python 3.10 或更高版本；查看 HTML 不需要 Python。')
    try:
        check_dependencies()
        return
    except ImportError:
        if __name__ != '__main__':
            raise
    if os.environ.get('HK_MORTGAGE_MANAGED_RUNTIME') == '1':
        raise RuntimeError('本地依赖环境未准备成功。请检查网络后重新运行。')
    import venv
    identity = '|'.join((sys.executable, sys.version, sys.platform, platform.machine()))
    tag = hashlib.sha256(identity.encode('utf-8')).hexdigest()[:12]
    runtime = Path(__file__).resolve().parent / '.mortgage-runtime' / (sys.platform + '-' + tag)
    executable = runtime / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')
    print('正在为这台电脑准备数据更新所需的 Python 库，请保持联网。', flush=True)
    if not executable.is_file():
        venv.EnvBuilder(with_pip=True).create(runtime)
    # Run the same check in the managed environment; do not trust a copied cache.
    probe_environment = dict(os.environ, HK_MORTGAGE_DEPENDENCY_PROBE='1')
    probe = subprocess.run([str(executable), str(Path(__file__).resolve())],
                           env=probe_environment, stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL)
    if probe.returncode:
        subprocess.run([str(executable), '-m', 'pip', 'install', '--upgrade', '--disable-pip-version-check',
                        '--retries', '2', '--timeout', '30', *DEPENDENCIES], check=True)
    environment = dict(os.environ, HK_MORTGAGE_MANAGED_RUNTIME='1')
    result = subprocess.run([str(executable), str(Path(__file__).resolve()), *sys.argv[1:]],
                            env=environment)
    raise SystemExit(result.returncode)


if __name__ == '__main__' and os.environ.get('HK_MORTGAGE_DEPENDENCY_PROBE') == '1':
    try:
        check_dependencies()
    except ImportError:
        raise SystemExit(1)
    raise SystemExit(0)

try:
    ensure_dependencies()
except Exception as exc:
    print(f'无法准备更新环境：{exc}\n已有 Excel 和 Dashboard 不会被删除。', file=sys.stderr)
    raise SystemExit(1)
import openpyxl
import xlrd
import bs4
import pypdf

# ----------------------------------------------------------------------------
# 常量
# ----------------------------------------------------------------------------
HKMA_XLSX_URL = ("https://www.hkma.gov.hk/media/eng/doc/market-data-and-statistics/"
                 "monthly-statistical-bulletin/T0307.xlsx")
CENTALINE_STAT_URL = ("https://api.centamortgage.com/api/Infomation/"
                      "GetStatistics?languageType=sc&year={year}")
LANDREG_JSON_URL = "https://www.landreg.gov.hk/json/monthly_stat/stat{mm:02d}{yy:02d}/t1.json"

MONTH_EN = ["January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"]

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

OUT_XLSX = "hk_mortgage_data.xlsx"
BASE = Path(__file__).resolve().parent
CACHE = BASE / '.mortgage_cache'
ISSUES = []
EVIDENCE = []
START_YEAR = 2012
REFRESH_HISTORY = False
PREVIOUS = {}
TABLE_RESULTS = []
FIELDS = ['现楼按揭宗数','楼花按揭宗数','私人住宅现楼宗数','资助房屋宗数']


# ----------------------------------------------------------------------------
# 工具函数
# ----------------------------------------------------------------------------
def http_get(url, binary=False, payload=None, cache=False, extra_headers=None):
    """Use normal HTTPS verification; curl also supports system proxy settings on macOS."""
    url = urllib.parse.quote(url, safe=':/?=&%+#')
    headers = dict(HEADERS, **(extra_headers or {}))
    body = json.dumps(payload, ensure_ascii=False).encode() if payload is not None else None
    key = hashlib.sha256(url.encode() + (body or b'')).hexdigest()
    cached = CACHE / (key + '.gz')
    if cache and cached.exists() and time.time() - cached.stat().st_mtime < 86400:
        data = gzip.decompress(cached.read_bytes())
        return data if binary else data.decode('utf-8', errors='replace')
    error = None
    for attempt in range(2):
        try:
            if shutil.which('curl'):
                command = ['curl', '-sSL', '--fail', '--connect-timeout', '10', '--max-time', '35', url]
                for name, value in headers.items():
                    command += ['-H', name + ': ' + value]
                if body is not None:
                    command += ['-H', 'Content-Type: application/json', '--data-binary', '@-']
                result = subprocess.run(command, input=body, capture_output=True, timeout=40)
                if result.returncode:
                    raise OSError(result.stderr.decode(errors='replace').strip())
                data = result.stdout
            else:
                if body is not None:
                    headers['Content-Type'] = 'application/json'
                req = urllib.request.Request(url, data=body, headers=headers)
                with urllib.request.urlopen(req, timeout=35) as resp:
                    data = resp.read()
            if not data:
                raise ValueError('Empty source response')
            if cache:
                CACHE.mkdir(exist_ok=True)
                cached.write_bytes(gzip.compress(data))
            break
        except Exception as exc:
            error = exc
            if attempt == 1 or '429' in str(exc):
                raise OSError(f'{url}: {error}') from exc
            time.sleep(0.5)
    return data if binary else data.decode("utf-8", errors="ignore")


def http_get_json(url):
    return json.loads(http_get(url))


def month_key(y, m):
    date(int(y),int(m),1)  # Reject impossible source months before accepting evidence.
    return f"{int(y):04d}-{int(m):02d}"


def parse_int(s):
    """解析可能带千分位逗号的字符串为 int; 无法解析返回 None"""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return int(s) if math.isfinite(s) and float(s).is_integer() else None
    t = str(s).strip().replace(",", "").replace(" ", "")
    if t in ("", "-", "—", "N/A", "n/a", "None"):
        return None
    try:
        value=float(t)
        return int(value) if math.isfinite(value) and value.is_integer() else None
    except ValueError:
        return None


def parse_float(s):
    """解析为 float, 保留小数; 无法解析返回 None"""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    t = str(s).strip().replace(",", "").replace(" ", "")
    if t in ("", "-", "—", "N/A", "n/a", "None"):
        return None
    try:
        value=float(t)
        return value if math.isfinite(value) else None
    except ValueError:
        return None


# ----------------------------------------------------------------------------
# [1] 中原按揭 — 现楼 / 楼花 (2012 至今)
# ----------------------------------------------------------------------------
def fetch_centaline():
    """返回 {month_key: {'现楼': int, '楼花': int}}"""
    out = {}
    years = range(max(2012, START_YEAR), hk_now().year + 1)
    for y in years:
        try:
            d = http_get_json(CENTALINE_STAT_URL.format(year=y))
            if not isinstance(d,dict) or not isinstance(d.get('StatisticsList'),list):
                raise ValueError('Centaline statistics response format changed')
        except Exception as e:
            ISSUES.append({'source': 'Centaline API', 'period': str(y), 'error': str(e)})
            print(f"  [中原] 拉取 {y} 年失败: {e}")
            continue
        for r in d.get("StatisticsList", []):
            now_floor = parse_int(r.get("NowFloorCount"))
            fut_floor = parse_int(r.get("FutureFloorCount"))
            # 现楼/楼花均无值的月份跳过
            if now_floor is None and fut_floor is None:
                continue
            key = month_key(r["Year"], r["Month"])
            out[key] = {"现楼按揭宗数": now_floor, "楼花按揭宗数": fut_floor}
    return out


# ----------------------------------------------------------------------------
# [1b] 中原按揭 detail 页兜底 — 补最新月 + 提取「私人住宅现楼/资助房屋」
# ----------------------------------------------------------------------------
def article_text(raw_html):
    # Read only the article, never related-news cards or the navigation.
    articles = re.findall(r'<article\b[^>]*>(.*?)</article>', raw_html, re.S | re.I)
    if articles:
        body = '\n'.join(articles)
    else:
        nuxt = re.search(r'window\.__NUXT__=(.*?)</script>', raw_html, re.S)
        if nuxt:
            vals = re.findall(r'value:"((?:[^"\\]|\\.)*)"', nuxt.group(1))
            body = '\n'.join(json.loads('"' + v + '"') for v in vals)
        else:
            body = re.sub(r'<(script|style)\b.*?</\1>', '', raw_html, flags=re.S | re.I)
    body = re.sub(r'</(?:p|h[1-6]|tr|div)>', '。', body, flags=re.I)
    text = unescape(re.sub(r'<[^>]+>', '', body))
    text = re.sub(r'\s+', '', text)
    return text.translate(str.maketrans('现楼约登记录资助计划夹转贷数', '現樓約登記錄資助計劃夾轉貸數'))


def extract_report(raw_html, month=None):
    """Return observed monthly totals and supporting phrases; never estimate missing data."""
    text = article_text(raw_html)
    number = r'([\d,，]+)宗'
    private = r'(?<![一二手])私人住宅現樓(?:按揭|樓按)(?:合約)?(?:登記量|登記宗數|登記)?'
    levels = {
        '私人住宅现楼宗数': [private + r'(?:錄得|錄|有|共|為|達)?[（(]?' + number,
            r'(?<![一二手])私人住宅按揭(?:合約)?(?:有|共)?'+number+
            r'[。；]{1,3}扣除[^。]{0,100}一手現樓[^。]{0,100}。[估計\s]*二手私人住宅現樓按揭'],
        '现楼按揭宗数': [r'樓宇按揭(?:合約)?[（(]不包括樓花[）)](?:有|共|錄得)?' + number,
                        r'(?<!住宅)現樓按揭(?:登記量|登記)?(?:錄得|錄|有|共|為|達)' + number],
        '楼花按揭宗数': [r'(?:私人住宅)?樓花按揭(?:合約)?(?:登記量|登記宗數|登記)?(?:錄得|錄|有|共|為|達)?' + number],
        '资助房屋宗数': [r'居屋[、,，]夾屋(?:[、,，]|及)+(?:及)?租置計劃(?:的)?按揭(?:合約)?(?:有|共|錄得)?' + number,
                       r'資助(?:性)?房屋按揭登記[^。；]{0,45}?至' + number],
    }
    out = {}; evidence = {}
    for field, patterns in levels.items():
        candidates = []
        for pattern in patterns:
            for match in re.finditer(pattern, text):
                start = max(text.rfind(mark, 0, match.start()) for mark in ['。','；','，','：']) + 1
                prefix = text[start:match.start()]
                if re.search(r'全年|個月|首[\d一二三四五六七八九十]+月|季度|季[度末]?|上半年|下半年|預計|估計|預期|料|展望|有望', prefix):
                    continue
                for number_month, chinese in enumerate(['一','二','三','四','五','六','七','八','九','十','十一','十二'],1):
                    prefix=re.sub(r'(?<![一二三四五六七八九十])'+chinese+'月',str(number_month)+'月',prefix)
                dates = re.findall(r'(?:(20\d{2}|19\d{2})年)?(\d{1,2})月份?', prefix)
                if not dates and re.search(r'(?:19|20)\d{2}年?',prefix):
                    continue  # A year-only observation is not a monthly value.
                if month and dates:
                    year, mon = dates[-1]
                    if int(mon) != int(month[5:]) or (year and int(year) != int(month[:4])):
                        continue
                value = int(match.group(1).replace(',', '').replace('，', ''))
                candidates.append((value, text[start:match.end()],bool(dates)))
        # An explicitly dated monthly observation outranks an undated clipped excerpt.
        if any(c[2] for c in candidates):
            candidates=[c for c in candidates if c[2]]
        values = {c[0] for c in candidates}
        if len(values) == 1:
            out[field] = candidates[0][0]
            evidence[field] = candidates[0][1]
        else:
            out[field] = None
            if len(values) > 1:
                evidence[field] = 'CONFLICT: ' + ' | '.join(x[1] for x in candidates)
    return out, evidence


def parse_centaline_detail(raw_html, month=None):
    return extract_report(raw_html, month)[0]


def search_excerpt_html(text):
    # Search separates non-contiguous excerpts with spaces. Never join them into
    # one invented sentence (e.g. a trailing '2017' + a new excerpt '3月').
    text = re.sub(r'</?em>', '', text)
    # A search highlight can also split a phrase at its internal HTML boundary.
    text = re.sub(r'(?<=[\u4e00-\u9fff]) +(?=(?:現樓|按揭|樓按|合約|登記|[（(]?\d[\d,]*宗))', '', text)
    return '<article>' + ''.join('<p>' + part + '</p>' for part in re.split(r'\s+', text) if part) + '</article>'


def extract_search_report(text, month):
    values,phrases=extract_report(search_excerpt_html(text),month)
    # An excerpt can start halfway through '二手私人住宅...'. Require a
    # monthly prefix for private totals in snippets; otherwise fetch full text.
    phrase=phrases.get('私人住宅现楼宗数','')
    prefix=phrase.split('私人住宅')[0]
    # The named total followed immediately by subtracting new completed units
    # to arrive at a secondary subtotal also establishes the metric's scope.
    context=article_text(search_excerpt_html(text))
    has_breakdown=bool(phrase and re.search(re.escape(phrase)+r'[。；]{1,3}扣除[^。]{0,100}一手現樓[^。]{0,100}。估計二手私人住宅現樓按揭',context))
    if not re.search(r'[\d一二三四五六七八九十]月份?',prefix) and not has_breakdown:
        values['私人住宅现楼宗数']=None
    elif has_breakdown:
        start=context.find(phrase)
        phrases['私人住宅现楼宗数']=context[start:context.find('二手私人住宅現樓按揭',start)+len('二手私人住宅現樓按揭')]
    return values,phrases


def expand_clipped_report(article):
    """Use numbers in a clipped excerpt only as search terms, never as data.

    Accept a value only after the same official article returns its full named
    metric, month and number. This needs no manually maintained URL/month list.
    """
    text=article_text(search_excerpt_html(article.get('sourceText','')))
    candidates=re.findall(r'私人住宅現樓(?:按揭|樓按)(?:合約)?(?:有)?([\d,]+)',text)
    candidates+=re.findall(r'(?:^|。)([\d,]+)宗，按月',text)
    for number in list(dict.fromkeys(candidates))[:3]:
        data=json.loads(http_get('https://hk.centanet.com/info/api/Article/Search',
            payload={'SeriesIds':[281],'Keyword':'現樓按揭'+number,'Offset':0,'Size':20},cache=True))
        for item in data.get('data',[]):
            if item.get('webUrl')!=article['webUrl']:
                continue
            snippet=item.get('highlight',{}).get('plainBody','')
            values,phrases=extract_search_report(snippet,article['month'])
            if values.get('私人住宅现楼宗数') is not None:
                return values,phrases
    return {},{}


def mortgage_article_html(url,recent=False):
    """The Mortgage site sometimes returns an empty JS shell; read its public article API."""
    match=re.search(r'_(\d+)$',url)
    if not match:
        return ''
    data=json.loads(http_get('https://www.centamortgage.com/centanetApi/Article/GetDetail?articleId='+match[1],
                            extra_headers={'AppName':'Mortgage'},cache=not REFRESH_HISTORY and not recent))
    parts=[]
    def visit(node):
        if isinstance(node,dict):
            if node.get('alias')=='rte' and isinstance(node.get('value'),str):
                parts.append(node['value'])
            for value in node.values():
                if isinstance(value,(dict,list)): visit(value)
        elif isinstance(node,list):
            for value in node:visit(value)
    visit(data.get('grid',{}))
    if isinstance(data.get('body'),str):parts.append(data['body'])
    return '<article>'+''.join(parts)+'</article>'


def research_article_html(url,recent=False):
    """Read the same public article when its presentation page is unavailable."""
    match=re.search(r'/(\d+)(?:\?.*)?$',url)
    if not match:raise ValueError('Research article ID missing')
    data=json.loads(http_get('https://hk.centanet.com/info/api/Article/GetDetail?articleId='+match[1],
                            cache=not REFRESH_HISTORY and not recent))
    if not isinstance(data,dict) or not isinstance(data.get('body'),str):
        raise ValueError('Research article API returned no body')
    attachments=data.get('attachments') or []
    links=''.join('<a href="'+item['fileUrl']+'"></a>' for item in attachments
                  if isinstance(item,dict) and isinstance(item.get('fileUrl'),str))
    return '<article>'+data['body']+'</article>'+links


def attachment_urls(raw):
    raw=unescape(raw.replace('\\u002F','/').replace('\\u0026','&'))
    return list(dict.fromkeys(url for url in re.findall(r'https://hkcms\.centanet\.com/[^"<>\s\\]+',raw)
               if re.search(r'\.(?:xls|xlsx|pdf)(?:$|&)',url,re.I)))


def parse_attachment_rows(rows, url):
    """Read named metrics under explicit month headers, including the previous month."""
    aliases={'私人住宅現樓':'私人住宅现楼宗数','居屋、夾屋、租置計劃':'资助房屋宗数',
             '樓宇按揭合約總數':'现楼按揭宗数','私人住宅樓花':'楼花按揭宗数'}
    columns={}; result=[]
    for row in rows:
        # New period headings start a new table; annual/cumulative columns must
        # never inherit the preceding monthly header.
        if any(re.search(r'20\d{2}年|累計|累计|季度|全年',str(cell)) for cell in row):
            columns={}
        for col,cell in enumerate(row):
            match=re.fullmatch(r'\s*(20\d{2})年(\d{1,2})月(?:宗數)?\s*',str(cell))
            if match: columns[col]=month_key(match[1],match[2])
        if not row:continue
        label=re.sub(r'\s+','',str(row[0])).replace('，','、').replace('及','、')
        field=aliases.get(label)
        if not field:continue
        for col,month in columns.items():
            if col>=len(row) or month<f'{START_YEAR}-01' or month>hk_now().strftime('%Y-%m'):continue
            value=parse_int(row[col])
            if value is None or value<0:continue
            result.append({'month':month,'field':field,'value':value,'url':url,
                           'text':f'{month} | {row[0]} | {row[col]}','retrieval':'official_attachment'})
    return result


def read_attachment(url, recent=False):
    raw=http_get(url,binary=True,cache=not recent and not REFRESH_HISTORY)
    if '.pdf' in url.lower():
        from pypdf import PdfReader
        rows=[]
        for page in PdfReader(io.BytesIO(raw)).pages:
            for line in page.extract_text(extraction_mode='layout').splitlines():
                rows.append(re.split(r'\s{2,}',line.strip()))
        return parse_attachment_rows(rows,url)
    import xlrd
    if '.xlsx' in url.lower():
        book=openpyxl.load_workbook(io.BytesIO(raw),read_only=True,data_only=True)
        result=[r for sheet in book for r in parse_attachment_rows(list(sheet.values),url)]
        book.close();return result
    book=xlrd.open_workbook(file_contents=raw)
    return [r for sheet in book.sheets() for r in parse_attachment_rows([sheet.row_values(i) for i in range(sheet.nrows)],url)]


def cached_article(url):
    key=hashlib.sha256(urllib.parse.quote(url,safe=':/?=&%+#').encode()).hexdigest()
    path=CACHE/(key+'.gz')
    if path.exists() and time.time()-path.stat().st_mtime<86400:
        return gzip.decompress(path.read_bytes()).decode('utf-8',errors='replace')
    return None


def fetch_centaline_detail_fallback():
    """
    自动查找历史月报并读取官网原文，返回 {month_key: dict}。
    可选的 centaline_detail_urls.json 支持补充已知文章地址，不是运行前提。
    补全 GetStatistics 未提供的私人住宅现楼/资助房屋及最新月份。
    """
    articles = discover_reports()
    cfg_path = BASE / 'centaline_detail_urls.json'
    if cfg_path.exists():
        for key, url in json.loads(cfg_path.read_text(encoding='utf-8')).items():
            if re.fullmatch(r'\d{4}-\d{2}', key):
                articles.insert(0, {'month': key, 'webUrl': url})
    threshold = hk_now().year * 12 + hk_now().month - 4
    table_lookup_months=set()
    for month,row in PREVIOUS.items():
        if any(row.get(k) is None for k in FIELDS):
            n=int(month[:4])*12+int(month[5:])
            table_lookup_months.update([month,month_key(n//12,n%12+1)])
    needed = [
        a for a in articles if REFRESH_HISTORY or a['month'] not in PREVIOUS
        or a['month'] in table_lookup_months
        or int(a['month'][:4]) * 12 + int(a['month'][5:]) >= threshold
    ]
    print(f'  [中原月报] 找到 {len(articles)} 篇；本次读取 {len(needed)} 篇', flush=True)
    results = []
    failures = {}
    def read(a):
        # Official search exposes source text as well as its canonical article URL.
        # Prefer a complete dated observation there, avoiding unnecessary article requests.
        snippet = a.get('sourceText', '')
        extracted, phrases = extract_search_report(snippet, a['month'])
        recent=int(a['month'][:4])*12+int(a['month'][5:])>=threshold
        # Every output column matters. A private total alone does not finish a month.
        complete=all(extracted.get(field) is not None for field in FIELDS)
        if complete and a['month'] not in table_lookup_months and not recent and not REFRESH_HISTORY:
            a = dict(a, retrieval='official_search')
            return a, extracted, phrases
        host=urllib.parse.urlparse(a['webUrl']).netloc
        try:
            cached=cached_article(a['webUrl']) if not recent and not REFRESH_HISTORY else None
            try:
                if failures.get(host,0)>=3 and cached is None:
                    raise OSError('Article host temporarily unavailable; trying public article API')
                raw = cached or http_get(a['webUrl'], cache=not REFRESH_HISTORY and not recent)
            except Exception:
                failures[host]=failures.get(host,0)+1
                article_id=a['webUrl'].rstrip('/').split('/')[-1]
                api_cache=cached_article('https://hk.centanet.com/info/api/Article/GetDetail?articleId='+article_id) if not recent and not REFRESH_HISTORY else None
                if host!='hk.centanet.com' or (failures.get('research_api',0)>=3 and api_cache is None):raise
                try:
                    raw=research_article_html(a['webUrl'],recent)
                    a=dict(a,retrieval='official_article_api')
                except Exception:
                    failures['research_api']=failures.get('research_api',0)+1
                    raise
            if not article_text(raw) and host=='www.centamortgage.com':
                raw=mortgage_article_html(a['webUrl'],recent)
                a=dict(a,retrieval='official_article_api')
            if not article_text(raw):
                raise ValueError('Article returned no body text; observation remains unavailable')
            values, evidence = extract_report(raw, a['month'])
            for url in attachment_urls(raw):
                try:
                    TABLE_RESULTS.extend(read_attachment(url,recent))
                except Exception as exc:
                    ISSUES.append({'source':'Centaline attachment','period':a['month'],'url':url,'error':str(exc)})
            for field in FIELDS:
                if values.get(field) is None and extracted.get(field) is not None and not evidence.get(field,'').startswith('CONFLICT:'):
                    values[field]=extracted[field];evidence[field]=phrases[field]
            return a, values, evidence
        except Exception as exc:
            failures[host]=failures.get(host,0)+1
            if '429' in str(exc):
                failures[host]=3
            ISSUES.append({'source':'Centaline article','period':a['month'],'url':a['webUrl'],'error':str(exc)})
            if extracted.get('私人住宅现楼宗数') is None and host=='hk.centanet.com' and '429' not in str(exc):
                try:
                    expanded,support=expand_clipped_report(a)
                    for field,value in expanded.items():
                        if value is not None:extracted[field]=value;phrases[field]=support[field]
                except Exception:pass
            return dict(a,retrieval='official_search'),extracted,phrases
    with ThreadPoolExecutor(max_workers=2) as pool:
        pending = {pool.submit(read, a): a for a in needed}
        for n, future in enumerate(as_completed(pending), 1):
            a = pending[future]
            try:
                results.append(future.result())
            except Exception as exc:
                ISSUES.append({'source':'Centaline report','period':a['month'],'url':a['webUrl'],'error':str(exc)})
            if n % 25 == 0 or n == len(needed):
                print(f'  [中原月报] {n}/{len(needed)}', flush=True)
    out = {}
    # Research reports preferred for historical classifications; newest Mortgage release fills gaps.
    for a, values, evidence in sorted(results, key=lambda r: ('hk.centanet.com/info/' not in r[0]['webUrl'], r[0]['webUrl'])):
        row = out.setdefault(a['month'], {})
        for field, value in values.items():
            if value is None:
                if evidence.get(field, '').startswith('CONFLICT:'):
                    ISSUES.append({'period':a['month'],'field':field,'url':a['webUrl'],'error':evidence[field]})
                continue
            EVIDENCE.append({'month':a['month'],'field':field,'value':value,'url':a['webUrl'],'text':evidence[field], 'retrieval':a.get('retrieval','article_html')})
            if field in row and row[field] != value:
                ISSUES.append({'period':a['month'],'field':field,'url':a['webUrl'],'error':f'Report conflict: {row[field]} vs {value}'})
            elif field not in row:
                row[field] = value
                if field == '私人住宅现楼宗数':
                    row['私人住宅来源'] = a['webUrl']
    grouped={}
    for item in TABLE_RESULTS:
        grouped.setdefault((item['month'],item['field']),[]).append(item)
    for (month,field),items in grouped.items():
        distinct={x['value'] for x in items}
        if len(distinct)!=1:
            ISSUES.append({'source':'Centaline attachments','period':month,'field':field,'error':f'Conflicting table values: {sorted(distinct)}'})
            continue
        # Explicit table headers resolve prose typos; never recalculate the number.
        out.setdefault(month,{})[field]=items[0]['value']
        EVIDENCE.extend(items)
    return out


def discover_reports():
    """Page through the public search used by Centaline's two websites; no search engine or AI."""
    sources = [
        ('https://hk.centanet.com/info/api/Article/Search', {'SeriesIds':[281], 'Keyword':'居屋、夾屋、租置計劃'}, {}),
        ('https://hk.centanet.com/info/api/Article/Search', {'SeriesIds':[281], 'Keyword':'私人住宅現樓樓按'}, {}),
        ('https://hk.centanet.com/info/api/Article/Search', {'SeriesIds':[281], 'Keyword':'其中一手市場'}, {}),
        ('https://hk.centanet.com/info/api/Article/Search', {'SeriesIds':[281], 'Keyword':'私人住宅現樓按揭'}, {}),
        ('https://hk.centanet.com/info/api/Article/Search', {'SeriesIds':[281], 'Keyword':'登記私人住宅現樓'}, {}),
        ('https://hk.centanet.com/info/api/Article/Search', {'SeriesIds':[281], 'Keyword':'按揭合約'}, {}),
        ('https://hk.centanet.com/info/api/Article/Search', {'SeriesIds':[281], 'TagCodes':['166']}, {}),
        ('https://www.centamortgage.com/centanetApi/Article/Search', {'SeriesIds':[805], 'Lang':'hk', 'CalledByCentanet':False}, {'AppName':'Mortgage'}),
    ]
    found = {}; seen = set()
    for url, query, headers in sources:
        offset = 0
        try:
            while True:
                data = json.loads(http_get(url, payload=dict(query, Offset=offset, Size=100), extra_headers=headers, cache=False))
                items = data.get('data')
                if not isinstance(items, list) or not isinstance(data.get('count'), int):
                    raise ValueError('Article search schema changed')
                if not items:
                    if offset < data['count']:
                        raise ValueError('Article search stopped before the reported total')
                    break
                for item in items:
                    title = item.get('subject', '')
                    if not re.search(r'按揭合[約约]|按揭登[記记].*統計|樓按.*[佔占]|[转轉]按登[記记]|私人住宅現樓|居屋.*夾屋', title):
                        continue
                    snippet = '\n'.join([item.get('plainBody',''), item.get('highlight',{}).get('plainBody','')])
                    match = re.search(r'(20\d{2}|19\d{2})年(\d{1,2})月份?', title)
                    if not match:
                        # Annual/forecast headlines can contain genuine monthly source tables.
                        # The candidate date only routes extraction; every accepted value
                        # must still pass the original-text/explicit-table month checks.
                        match=re.search(r'(20\d{2}|19\d{2})年(\d{1,2})月份?',re.sub(r'<[^>]+>','',snippet))
                    if not match or not START_YEAR <= int(match[1]) <= hk_now().year:
                        continue
                    month = month_key(int(match[1]), int(match[2]))
                    link = item.get('webUrl')
                    if not link or not link.startswith(('https://hk.centanet.com/', 'https://www.centamortgage.com/')):
                        continue
                    if (month, link) not in seen:
                        found[link] = {'month':month, 'webUrl':link, 'subject':title, 'sourceText':snippet}
                        seen.add((month, link))
                    else:
                        found[link]['sourceText'] += '\n' + snippet
                offset += len(items)
                if offset >= data['count']:
                    break
                if offset > 10000:
                    raise ValueError('Unexpected article search size')
        except Exception as exc:
            ISSUES.append({'source':'Centaline discovery','url':url,'error':str(exc)})
    (BASE / 'discovered_reports.json').write_text(json.dumps(list(found.values()), ensure_ascii=False, indent=2), encoding='utf-8')
    return list(found.values())


# ----------------------------------------------------------------------------
# [2] HKMA — 住宅按揭统计调查 (表 T0307)
# ----------------------------------------------------------------------------
def fetch_hkma():
    """
    下载 T0307.xlsx 并解析。
    返回 {month_key: {新申请宗数, 新批出宗数, 新批出金额_HK$Mn, 一手宗数, 二手宗数, 转按宗数,
                    一手金额, 二手金额, 转按金额, 新取用宗数, 新取用金额, 平均贷款额, 未偿还总额}}
    """
    global HKMA_REMOTE_OK
    HKMA_REMOTE_OK=False
    xlsx_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "T0307.xlsx")
    try:
        data = http_get(HKMA_XLSX_URL, binary=True)
        with open(xlsx_path, "wb") as f:
            f.write(data)
        print(f"  [HKMA] 已下载 T0307.xlsx ({len(data)} bytes)")
        HKMA_REMOTE_OK=True
    except Exception as e:
        ISSUES.append({'source':'HKMA', 'error':str(e), 'action':'Use last available local source if present'})
        print(f"  [HKMA] 下载失败: {e}, 尝试使用本地缓存 {xlsx_path}")

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["T3.7"]
    rows = list(ws.iter_rows(values_only=True))
    validate_hkma_layout(rows)

    # 列映射 (0-indexed), 来自 T3.7 表头
    C = {
        "outstanding": 3,      # 未偿还按揭总额 (HK$ Mn)
        "applications": 19,    # 新申请宗数
        "drawn_val": 22,       # 新取用金额 (HK$ Mn)
        "drawn_no": 25,        # 新取用宗数
        "approved_val": 28,    # 新批出金额 (HK$ Mn)
        "approved_no": 31,     # 新批出宗数
        "avg_size": 33,        # 平均贷款额 (HK$ Mn)
        "primary_val": 39,     # 一手金额 (HK$ Mn)
        "secondary_val": 40,   # 二手金额 (HK$ Mn)
        "refi_val": 41,        # 转按金额 (HK$ Mn)
        "primary_no": 43,      # 一手宗数
        "secondary_no": 44,    # 二手宗数
        "refi_no": 45,         # 转按宗数
    }

    def cell(r, c):
        return r[c] if c < len(r) else None

    out = {}
    year = None
    for r in rows:
        # row[0] = 年份(仅新年份行有值, 合并单元格), row[1] = 英文月份名
        if r[0] is not None and isinstance(r[0], (int, float)):
            year = int(r[0])
        month_name = r[1] if len(r) > 1 else None
        if month_name not in MONTH_EN:
            continue
        m = MONTH_EN.index(month_name) + 1
        key = month_key(year, m)
        out[key] = {
            "新申请宗数": parse_int(cell(r, C["applications"])),
            "新批出宗数": parse_int(cell(r, C["approved_no"])),
            "新批出金额_HK$Mn": parse_int(cell(r, C["approved_val"])),
            "一手宗数": parse_int(cell(r, C["primary_no"])),
            "二手宗数": parse_int(cell(r, C["secondary_no"])),
            "转按宗数": parse_int(cell(r, C["refi_no"])),
            "一手金额_HK$Mn": parse_int(cell(r, C["primary_val"])),
            "二手金额_HK$Mn": parse_int(cell(r, C["secondary_val"])),
            "转按金额_HK$Mn": parse_int(cell(r, C["refi_val"])),
            "新取用宗数": parse_int(cell(r, C["drawn_no"])),
            "新取用金额_HK$Mn": parse_int(cell(r, C["drawn_val"])),
            "平均贷款额_HK$Mn": parse_float(cell(r, C["avg_size"])),
            "未偿还总额_HK$Mn": parse_int(cell(r, C["outstanding"])),
        }
    wb.close()
    validate_hkma(out)
    return out


def validate_hkma_layout(rows):
    expected={3:['Outstanding loans','Value'],19:['New applications received','Number'],
              22:['New loans drawn down','Value'],25:['Number'],28:['New loans approved','Value'],
              31:['Number'],33:['Average size'],39:['Value','Primary market'],
              40:['Secondary market'],41:['Refinancing'],43:['Number','Primary market'],
              44:['Secondary market'],45:['Refinancing']}
    for col,labels in expected.items():
        cells=[re.sub(r'\s+',' ',str(r[col])).strip() for r in rows[:20] if col<len(r) and r[col] is not None]
        for label in labels:
            if not any(text.startswith(label) for text in cells):
                raise ValueError(f'HKMA column {col+1}: expected heading {label}; source layout changed')


def validate_hkma(rows):
    required=['新申请宗数','新批出宗数','一手宗数','二手宗数','转按宗数',
              '新批出金额_HK$Mn','一手金额_HK$Mn','二手金额_HK$Mn','转按金额_HK$Mn',
              '新取用宗数','新取用金额_HK$Mn','平均贷款额_HK$Mn','未偿还总额_HK$Mn']
    if not rows:
        raise ValueError('HKMA returned no monthly records')
    for month,row in rows.items():
        for field in required:
            value=row.get(field)
            if not isinstance(value,(int,float)) or not math.isfinite(value) or value<0:
                raise ValueError(f'HKMA {month}: missing/invalid {field}; check source layout')
        if row['新批出宗数']!=sum(row[k] for k in ['一手宗数','二手宗数','转按宗数']):
            raise ValueError(f'HKMA {month}: approved count does not equal category counts')
        if abs(row['新批出金额_HK$Mn']-sum(row[k] for k in ['一手金额_HK$Mn','二手金额_HK$Mn','转按金额_HK$Mn']))>2:
            raise ValueError(f'HKMA {month}: approved amount does not reconcile within rounding')


# ----------------------------------------------------------------------------
# [3] 土地注册处 — 按揭登记宗数 (官方原始口径)
# ----------------------------------------------------------------------------
def fetch_landreg(months_back=24):
    """
    拉取最近 N 个月的 t1.json, 提取三类按揭登记宗数。
    返回 {month_key: {'楼宇按揭宗数', '其他按揭法定押记宗数', '撤销按揭宗数'}}
    """
    out = {}
    now = hk_now()
    for i in range(months_back):
        # 从当前月往前推
        total = now.year * 12 + (now.month - 1) - i - 1
        y, m = total // 12, total % 12 + 1
        fname = f"stat{m:02d}{y % 100:02d}"
        url = LANDREG_JSON_URL.format(mm=m, yy=y % 100)
        try:
            d = http_get_json(url)
            if not isinstance(d,list) or not all(isinstance(r,dict) for r in d):
                raise ValueError('Land Registry response format changed')
        except Exception as e:
            ISSUES.append({'source':'Land Registry','period':month_key(y,m),'error':str(e)})
            print(f"  [土地注册处] {y}-{m:02d} 拉取失败: {e}")
            continue
        rec = {"楼宇按揭宗数": None, "其他按揭法定押记宗数": None, "撤销按揭宗数": None}
        for row in d:
            desc = row.get("Description", "")
            if "Total Number of Urban" not in desc:
                continue
            units = parse_int(row.get("Units"))
            if "Building Mortgages/Building Legal Charges" in desc:
                rec["楼宇按揭宗数"] = units
            elif "Other Mortgages/Legal Charges" in desc:
                rec["其他按揭法定押记宗数"] = units
            elif "Receipts/Discharges/Releases" in desc:
                rec["撤销按揭宗数"] = units
        if any(v is None for v in rec.values()):
            ISSUES.append({'source':'Land Registry','period':month_key(y,m),'error':'Expected mortgage categories missing; old values retained'})
        if any(v is not None for v in rec.values()):
            out[month_key(y, m)] = rec
    return out


# ----------------------------------------------------------------------------
# 写 Excel (多 sheet)
# ----------------------------------------------------------------------------
MSB_ROOT='https://www.hkma.gov.hk/media/eng/doc/market-data-and-statistics/monthly-statistical-bulletin/'
HSBC_URL='https://www.hsbc.com.hk/investments/market-information/hk/lending-rate/'
MACRO_LAYOUT={
    'M3':['月份','港元M3(HK$Mn)','所有货币M3(HK$Mn)'],
    'HSBC Prime':['生效日期','HSBC Prime Rate(%)'],
    'HIBOR Daily':['日期','Overnight(%)','1 Week(%)','1 Month(%)','3 Months(%)','6 Months(%)','12 Months(%)']}
MACRO_ROWS={}; MACRO_FRESH={}; MACRO_CHECKS={}


def xls_book(url):
    import xlrd
    return xlrd.open_workbook(file_contents=http_get(url,binary=True))


def rate_cell(cell):
    import xlrd
    if cell.ctype in (xlrd.XL_CELL_EMPTY,xlrd.XL_CELL_BLANK,xlrd.XL_CELL_ERROR):return None
    value=parse_float(cell.value)
    if value is None or not 0<=value<=100:raise ValueError('Invalid interest rate')
    return value


def fetch_m3():
    # Same HKMA source and column checks as the existing repository collector.
    book=xls_book(MSB_ROOT+'T020201.xls'); sheet=book.sheet_by_name('T2.2.1 (new series)')
    if str(sheet.cell_value(7,11)).strip()!='M3' or sheet.cell_value(9,11)!='HK$2' or sheet.cell_value(9,13)!='Total':
        raise ValueError('HKMA M3 workbook columns changed')
    months={name:i for i,name in enumerate(calendar.month_abbr) if i}; year=None; out={}
    for i in range(sheet.nrows):
        row=sheet.row_values(i);mon=months.get(str(row[1]).strip())
        if not mon:continue
        if row[0] not in ('',None):year=parse_int(row[0])
        if not year:raise ValueError('M3 month has no valid year')
        period=month_key(year,mon)
        if period<f'{START_YEAR}-01' or period>=hk_now().strftime('%Y-%m'):continue
        values=[parse_float(row[c]) for c in (11,13)]
        if any(v is None or v<0 for v in values):raise ValueError('Invalid M3 value')
        out[period]=dict(zip(MACRO_LAYOUT['M3'][1:],values))
    if not out:raise ValueError('No monthly M3 observations')
    MACRO_CHECKS['M3']={'checked_at':hk_now().isoformat(),'source':MSB_ROOT+'T020201.xls'}
    return out


def fetch_prime(previous=None):
    import xlrd
    from bs4 import BeautifulSoup
    rates={}; history_ok=True
    try:
        book=xls_book(MSB_ROOT+'T060401.xls'); sheet=book.sheet_by_name('T6.4.1')
        header=' '.join(str(sheet.cell_value(i,10)) for i in (7,8,9))
        if 'Best lending rate' not in header:raise ValueError('Prime workbook column changed')
        notes=' '.join(str(c) for i in range(max(0,sheet.nrows-15),sheet.nrows) for c in sheet.row_values(i))
        if 'Shanghai Banking Corporation Limited' not in notes:raise ValueError('Prime is not attributed to HSBC')
        for i in range(sheet.nrows):
            if sheet.cell(i,0).ctype!=xlrd.XL_CELL_DATE:continue
            day=xlrd.xldate_as_datetime(sheet.cell_value(i,0),book.datemode).date().isoformat()
            value=rate_cell(sheet.cell(i,10))
            if value is not None and day<=hk_now().date().isoformat():rates[day]=value
    except Exception as exc:
        if not previous:raise
        history_ok=False
        rates={d:r[MACRO_LAYOUT['HSBC Prime'][1]] for d,r in previous.items()}
        ISSUES.append({'source':'Prime HKMA history','error':str(exc),'action':'Preserve history; check HSBC current independently'})
    # HSBC's current page can publish a change ahead of the monthly HKMA workbook.
    soup=BeautifulSoup(http_get(HSBC_URL),'html.parser'); text=soup.get_text(' ',strip=True)
    current=re.search(r'Current Hong Kong Dollar Best Lending Rate:\s*(\d+(?:\.\d+)?)%',text)
    if not current:raise ValueError('HSBC current rate heading missing')
    for tr in soup.select('table tr'):
        cells=[c.get_text(' ',strip=True) for c in tr.select('td')]
        if len(cells)!=2 or not re.fullmatch(r'\d{1,2} [A-Za-z]{3} \d{4}',cells[0]):continue
        day=datetime.strptime(cells[0],'%d %b %Y').date().isoformat(); value=float(cells[1].rstrip('%'))
        if day>hk_now().date().isoformat() or not 0<=value<=100:raise ValueError('Invalid HSBC effective rate')
        if day in rates and rates[day]!=value:raise ValueError('HSBC / HKMA effective-rate conflict')
        rates[day]=value
    if not rates or rates[max(rates)]!=float(current[1]):raise ValueError('Prime history differs from current HSBC rate')
    # Include the last actual effective date before the requested start as its baseline.
    prior=[d for d in rates if d<f'{START_YEAR}-01-01']; baseline=max(prior) if prior else None
    out={d:{MACRO_LAYOUT['HSBC Prime'][1]:v} for d,v in sorted(rates.items()) if d>=f'{START_YEAR}-01-01' or d==baseline}
    MACRO_CHECKS['HSBC Prime']={'checked_at':hk_now().isoformat(),'source':HSBC_URL,'current_rate':float(current[1]),'history_fresh':history_ok}
    return out


def parse_hkab(data, day):
    labels=['Overnight','1 Week','1 Month','3 Months','6 Months','12 Months']
    if not isinstance(data,dict):raise ValueError('Invalid HKAB response')
    match=re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})',str(data.get('date','')))
    if data.get('date') is not None and (not match or date(*map(int,match.groups()))!=day):
        raise ValueError('HKAB response date differs from requested date')
    for key,value in [('year',day.year),('month',day.month),('day',day.day)]:
        if data.get(key) is not None and str(data[key])!=str(value):raise ValueError('HKAB date fields conflict')
    if data.get('isHoliday') is True:
        if any(data.get(k) is not None for k in labels):raise ValueError('Holiday contains fixing values')
        return None
    if not match:raise ValueError('HKAB fixing has no observation date')
    if day.weekday()>=5:raise ValueError('Unexpected weekend HIBOR fixing')
    if str(data.get('time','')).replace(' ','').lower()!='11:15am':raise ValueError('Wrong HKAB fixing time')
    result={}
    for label in labels:
        value=data.get(label)
        if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or not 0<=value<=100:
            raise ValueError(f'HKAB {label}: invalid fixing')
        result[label+'(%)']=float(value)
    return result


def fetch_hibor(previous=None):
    import xlrd
    out={}; now=hk_now(); today=now.date()
    history_ok=True
    try:
        book=xls_book(MSB_ROOT+'T060303.xls'); sheet=book.sheet_by_name('T6.3.3 (HIBOR Fixing)')
        if sheet.cell_value(7,5)!='1-month' or '11:15' not in str(sheet.cell_value(7,0)):
            raise ValueError('HIBOR fixing workbook header changed')
        for i in range(sheet.nrows):
            if sheet.cell(i,0).ctype!=xlrd.XL_CELL_DATE:continue
            day=xlrd.xldate_as_datetime(sheet.cell_value(i,0),book.datemode).date()
            if day.year<START_YEAR or day>today or day.weekday()>=5:continue
            values=[rate_cell(sheet.cell(i,col)) for col in [3,4,5,6,7,9]]
            if any(v is not None for v in values):out[day.isoformat()]=dict(zip(MACRO_LAYOUT['HIBOR Daily'][1:],values))
    except Exception as exc:
        if not previous:raise
        history_ok=False
        out={}
        ISSUES.append({'source':'HIBOR HKMA history','error':str(exc),'action':'Preserve history; check HKAB daily independently'})
    history=out or previous or {}
    if not history:raise ValueError('No HIBOR fixing history')
    last=today if (now.hour,now.minute)>=(11,30) else today-timedelta(days=1)
    # Recheck recent fixings every run, also repairing isolated holes in that window.
    start=min(date.fromisoformat(max(history))+timedelta(days=1),last-timedelta(days=7))
    if (last-start).days>62:raise ValueError('HIBOR source more than 62 days behind; manual source review required')
    holidays=[]; failures=[]
    for n in range(max(0,(last-start).days+1)):
        day=start+timedelta(days=n)
        if day.weekday()>=5:continue
        url=f'https://www.hkab.org.hk/api/hibor?year={day.year}&month={day.month}&day={day.day}'
        try:
            row=parse_hkab(http_get_json(url),day)
            if row is None:holidays.append(day.isoformat())
            else:out[day.isoformat()]=row
        except Exception as exc:
            failures.append(day.isoformat());ISSUES.append({'source':'HKAB daily HIBOR','period':day.isoformat(),'url':url,'error':str(exc)})
    expected=last
    while expected.weekday()>=5 or expected.isoformat() in holidays:expected-=timedelta(days=1)
    MACRO_CHECKS['HIBOR Daily']={'checked_at':now.isoformat(),'source':'https://www.hkab.org.hk/','expected_date':expected.isoformat(),'holidays_skipped':holidays,'failed_dates':failures,'history_fresh':history_ok}
    return out


def collect_macro(previous):
    for name,fetcher in [('M3',fetch_m3),('HSBC Prime',fetch_prime),('HIBOR Daily',fetch_hibor)]:
        try:
            fresh=fetcher(previous.get(name,{})) if name in ('HSBC Prime','HIBOR Daily') else fetcher()
            MACRO_FRESH[name]=fresh
        except Exception as exc:
            fresh={};MACRO_FRESH[name]={};ISSUES.append({'source':name,'error':str(exc)})
            print(f'提醒：{name} 未取得本次数据，保留已有值。{exc}',flush=True)
        MACRO_ROWS[name]=merge_rows(previous.get(name,{}),fresh)
        if not MACRO_ROWS[name]:raise ValueError(f'{name}: no verified history available; workbook not saved')
        print(f'  {name}: {len(MACRO_ROWS[name])} 行；最新 {max(MACRO_ROWS[name])}',flush=True)


SOURCE_HEADERS = {
    "中原按揭": ("月份", *FIELDS),
    "HKMA": (
        "月份", "新申请宗数", "新批出宗数", "新批出金额(HK$Mn)",
        "一手宗数", "二手宗数", "转按宗数", "一手金额(HK$Mn)",
        "二手金额(HK$Mn)", "转按金额(HK$Mn)", "新取用宗数",
        "新取用金额(HK$Mn)", "平均贷款额(HK$Mn)", "未偿还总额(HK$Mn)",
    ),
    "土地注册处": ("月份", "楼宇按揭宗数", "其他按揭/法定押记宗数", "撤销按揭宗数"),
}
NONBANK_SHEET = "非银行"
NONBANK_HEADERS = ("Date", "非银行贷款宗数", "真实买卖总数")


def period_key(value, daily=False):
    """Accept the workbook's ISO text dates and native Excel dates."""
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d" if daily else "%Y-%m")
    if not isinstance(value, str):
        return None
    pattern = r"\d{4}-\d{2}-\d{2}" if daily else r"\d{4}-\d{2}"
    if not re.fullmatch(pattern, value):
        return None
    try:
        date.fromisoformat(value if daily else value + "-01")
    except ValueError:
        return None
    return value


def source_field(header):
    return header.replace("(HK$Mn)", "_HK$Mn").replace(
        "其他按揭/法定押记宗数", "其他按揭法定押记宗数"
    )


def check_headers(sheet, headers):
    actual = tuple(sheet.cell(1, column).value for column in range(1, len(headers) + 1))
    if actual != tuple(headers):
        raise ValueError(f"{sheet.title}：表头已改变，停止写入。预期 {tuple(headers)}")


def period_rows(sheet, daily=False):
    rows = {}
    for row_number, (value,) in enumerate(
        sheet.iter_rows(min_row=2, max_col=1, values_only=True), start=2
    ):
        key = period_key(value, daily)
        if key is None:
            continue
        if key in rows:
            raise ValueError(f"{sheet.title}：日期 {key} 重复，停止写入。")
        rows[key] = row_number
    return rows


def append_period(sheet, key, template_row, width):
    row = sheet.max_row + 1
    sheet.cell(row, 1, key)
    if template_row is not None:
        for column in range(1, width + 1):
            sheet.cell(row, column)._style = copy(sheet.cell(template_row, column)._style)
    return row


def update_source_sheet(sheet, records, headers, *, macro=False):
    """Update matched dates and leave other cells and the order of existing rows intact."""
    check_headers(sheet, headers)
    daily = sheet.title in ("HSBC Prime", "HIBOR Daily")
    rows = period_rows(sheet, daily)
    template = max(rows.values(), default=None)
    for key, record in sorted(records.items()):
        if period_key(key, daily) is None:
            raise ValueError(f"{sheet.title}：无法识别日期 {key}。")
        row = rows.get(key)
        if row is None:
            if not any(record.get(h if macro else source_field(h)) is not None for h in headers[1:]):
                continue
            row = append_period(sheet, key, template, len(headers))
        for column, header in enumerate(headers[1:], 2):
            value = record.get(header if macro else source_field(header))
            cell = sheet.cell(row, column)
            if value is not None and cell.value != value:
                cell.value = value
            if macro and template is None:
                cell.number_format = {
                    "M3": "#,##0.000",
                    "HSBC Prime": "0.000",
                    "HIBOR Daily": "0.00000",
                }[sheet.title]


def previous_month(key):
    return (date.fromisoformat(key + "-01") - timedelta(days=1)).strftime("%Y-%m")


def count_value(value):
    return type(value) in (int, float) and math.isfinite(value) and value >= 0


def update_nonbank_sheet(workbook):
    """Extend the attached workbook's formulas and calculate their saved values."""
    centaline = workbook["中原按揭"]
    bank = workbook["HKMA"]
    check_headers(centaline, SOURCE_HEADERS["中原按揭"])
    check_headers(bank, SOURCE_HEADERS["HKMA"])
    centaline_rows = period_rows(centaline)
    bank_rows = period_rows(bank)
    if not bank_rows:
        raise ValueError("HKMA：没有可计算的月份。")

    if NONBANK_SHEET in workbook:
        output = workbook[NONBANK_SHEET]
    else:
        output = workbook.create_sheet(NONBANK_SHEET, 1)
        output.append(NONBANK_HEADERS)
        for column in range(1, 4):
            output.cell(1, column)._style = copy(centaline.cell(1, column)._style)
        output.column_dimensions["A"].width = 14
        output.column_dimensions["B"].width = 22
        output.column_dimensions["C"].width = 22
        output.freeze_panes = "A2"
    check_headers(output, NONBANK_HEADERS)
    output_rows = period_rows(output)
    template = max(output_rows.values(), default=None)
    first_bank_month = min(bank_rows)
    eligible_months = set(bank_rows) | {key for key in centaline_rows if key >= first_bank_month}
    for key in sorted(eligible_months - output_rows.keys()):
        output_rows[key] = append_period(output, key, template, 3)

    purchase_drawdowns = {}
    cached_values = {}
    for key, row in output_rows.items():
        bank_row = bank_rows.get(key)
        formula = None
        if bank_row is not None:
            approved, primary, secondary, drawn = (
                bank.cell(bank_row, column).value for column in (3, 5, 6, 11)
            )
            if all(count_value(value) for value in (approved, primary, secondary, drawn)):
                if approved > 0 and primary + secondary <= approved:
                    purchase_drawdowns[key] = (primary + secondary) / approved * drawn
                    formula = f"=(HKMA!E{bank_row}+HKMA!F{bank_row})/HKMA!C{bank_row}*HKMA!K{bank_row}"
                    cached_values[f"C{row}"] = purchase_drawdowns[key]
        output.cell(row, 3).value = formula

    for key, row in output_rows.items():
        prior = previous_month(key)
        registration_row = centaline_rows.get(key)
        formula = None
        if registration_row is not None and prior in purchase_drawdowns:
            completed = centaline.cell(registration_row, 2).value
            off_plan = centaline.cell(registration_row, 3).value
            if count_value(completed) and count_value(off_plan):
                formula = (
                    f"=SUM(中原按揭!B{registration_row}:C{registration_row})-C{output_rows[prior]}"
                )
                cached_values[f"B{row}"] = completed + off_plan - purchase_drawdowns[prior]
        output.cell(row, 2).value = formula
    return cached_values


def cache_formula_results(source, destination, sheet_number, values):
    """Save numeric formula results alongside the Excel formulas for local readers."""
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    part = f"xl/worksheets/sheet{sheet_number}.xml"
    with ZipFile(source) as original, ZipFile(destination, "w") as result:
        for info in original.infolist():
            data = original.read(info)
            if info.filename == part:
                xml = ElementTree.fromstring(data)
                found = set()
                for cell in xml.iter(f"{{{namespace}}}c"):
                    address = cell.get("r")
                    if address not in values:
                        continue
                    if cell.find(f"{{{namespace}}}f") is None:
                        raise ValueError(f"{NONBANK_SHEET}!{address}：公式缺失。")
                    value_node = cell.find(f"{{{namespace}}}v")
                    if value_node is None:
                        value_node = ElementTree.SubElement(cell, f"{{{namespace}}}v")
                    value_node.text = repr(values[address])
                    found.add(address)
                if found != values.keys():
                    raise ValueError("非银行：无法保存全部计算结果。")
                data = ElementTree.tostring(xml, encoding="utf-8")
            result.writestr(info, data)


def save_existing_workbook(workbook, out_path, cached_values, original_bytes):
    """Validate the full saved file before replacing it at the same path."""
    out_path = Path(out_path)
    with tempfile.TemporaryDirectory(prefix=".mortgage-save-", dir=out_path.parent) as folder:
        raw_path = Path(folder) / "workbook.xlsx"
        final_path = Path(folder) / "checked.xlsx"
        workbook.save(raw_path)
        cache_formula_results(
            raw_path, final_path, workbook.index(workbook[NONBANK_SHEET]) + 1, cached_values
        )
        check = openpyxl.load_workbook(final_path, read_only=True, data_only=True)
        try:
            if check.sheetnames != workbook.sheetnames:
                raise ValueError("保存验证失败：工作表清单有变动。")
            actual = {
                cell.coordinate: cell.value
                for row in check[NONBANK_SHEET].iter_rows(min_col=2, max_col=3)
                for cell in row if cell.value is not None
            }
            for address, value in cached_values.items():
                if address not in actual or not math.isclose(actual[address], value, abs_tol=1e-9):
                    raise ValueError(f"保存验证失败：{NONBANK_SHEET}!{address} 计算结果不一致。")
        finally:
            check.close()
        if out_path.read_bytes() != original_bytes:
            raise RuntimeError("Excel 在写入期间发生变动。请保存并关闭 Excel 后重新运行。")
        backup = Path(folder) / "backup.xlsx"
        backup.write_bytes(original_bytes)
        backup.replace(out_path.with_name(out_path.stem + ".previous.xlsx"))
        final_path.replace(out_path)
    print(f"已更新: {out_path}")


def open_existing_workbook(out_path):
    out_path = Path(out_path)
    if not out_path.is_file():
        raise FileNotFoundError(f"请先放入要更新的工作簿：{out_path}")
    lock_file = out_path.with_name("~$" + out_path.name)
    if lock_file.exists():
        raise RuntimeError("请保存并关闭 Excel 工作簿，再运行更新。")
    content = out_path.read_bytes()
    return openpyxl.load_workbook(io.BytesIO(content), rich_text=True), content


def write_xlsx(centaline, hkma, landreg, out_path, *, expected_original=None):
    workbook, original_bytes = open_existing_workbook(out_path)
    try:
        if expected_original is not None and original_bytes != expected_original:
            raise RuntimeError("Excel 在抓取期间发生变动。请保存并关闭 Excel 后重新运行。")
        for name, data in [("中原按揭", centaline), ("HKMA", hkma), ("土地注册处", landreg)]:
            update_source_sheet(workbook[name], data, SOURCE_HEADERS[name])
        for name, headers in MACRO_LAYOUT.items():
            if name not in MACRO_ROWS:
                continue
            if name not in workbook:
                sheet = workbook.create_sheet(name)
                sheet.append(headers)
                sheet.freeze_panes = "A2"
                for column, label in enumerate(headers, 1):
                    sheet.cell(1, column)._style = copy(workbook["中原按揭"]["A1"]._style)
                    sheet.column_dimensions[openpyxl.utils.get_column_letter(column)].width = max(16, len(label) + 7)
            update_source_sheet(workbook[name], MACRO_ROWS[name], headers, macro=True)
        if "说明" in workbook and workbook["说明"]["A3"].value == "生成时间":
            workbook["说明"]["B3"] = hk_now().strftime("%Y-%m-%d %H:%M:%S")
        results = update_nonbank_sheet(workbook)
        save_existing_workbook(workbook, out_path, results, original_bytes)
    finally:
        workbook.close()


def recalculate_existing(out_path):
    workbook, original_bytes = open_existing_workbook(out_path)
    try:
        results = update_nonbank_sheet(workbook)
        save_existing_workbook(workbook, out_path, results, original_bytes)
    finally:
        workbook.close()



# -----------------------------------------------------------------------------
# 主流程
# -----------------------------------------------------------------------------

def load_existing(path):
    result = {name:{} for name in ['中原按揭','HKMA','土地注册处',*MACRO_LAYOUT]}
    if not Path(path).exists():
        return result
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for name in result:
        if name not in wb.sheetnames:
            continue
        sheet = wb[name]
        headers = MACRO_LAYOUT[name] if name in MACRO_LAYOUT else SOURCE_HEADERS[name]
        check_headers(sheet, headers)
        rows = sheet.iter_rows(min_row=2, max_col=len(headers), values_only=True)
        for row in rows:
            key = period_key(row[0], name in ("HSBC Prime", "HIBOR Daily"))
            if key is None:
                continue
            if key in result[name]:
                raise ValueError(f"{name}：日期 {key} 重复，停止写入。")
            result[name][key] = {
                (h if name in MACRO_LAYOUT else source_field(h)): value
                for h, value in zip(headers[1:], row[1:])
            }
    wb.close()
    return result


def merge_rows(previous, fresh):
    out = {k:dict(v) for k,v in previous.items()}
    for month, values in fresh.items():
        out.setdefault(month, {}).update({k:v for k,v in values.items() if v is not None})
    return out


def verified_table_values(previous,evidence):
    return {(r['month'],r['field']):r['value'] for r in evidence
            if r.get('retrieval')=='official_attachment' and previous.get(r['month'],{}).get(r['field'])==r['value']}


def dated_source_values(evidence):
    return {(r['month'],r['field'],r['value']) for r in evidence if r.get('retrieval')=='official_attachment'
            or re.search(r'[\d一二三四五六七八九十]月份?',r.get('text','').split('按揭')[0])}


def merge_centaline_api(previous, fresh, evidence):
    out=merge_rows(previous,fresh)
    verified={(month,field):value for month,field,value in dated_source_values(evidence) if previous.get(month,{}).get(field)==value}
    for (month,field),value in verified.items():
        if fresh.get(month,{}).get(field) is not None and fresh[month][field]!=value:
            out[month][field]=value
    return out


def align_months(rows, start, end):
    first=int(start[:4])*12+int(start[5:])-1
    last=int(end[:4])*12+int(end[5:])-1
    return {month_key(n//12,n%12+1):dict(rows.get(month_key(n//12,n%12+1),{})) for n in range(first,last+1)}


def dashboard_snapshot(path):
    """Export only the saved workbook's reviewed data, including formula results."""
    path = Path(path)
    content = path.read_bytes()
    layouts = {
        'registrations': ('中原按揭', SOURCE_HEADERS['中原按揭']),
        'bank': ('HKMA', SOURCE_HEADERS['HKMA']),
        'registry': ('土地注册处', SOURCE_HEADERS['土地注册处']),
        'm3': ('M3', MACRO_LAYOUT['M3']),
        'prime': ('HSBC Prime', MACRO_LAYOUT['HSBC Prime']),
        'hibor': ('HIBOR Daily', MACRO_LAYOUT['HIBOR Daily']),
        'nonbank': (NONBANK_SHEET, NONBANK_HEADERS),
    }
    book = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    series = {}
    try:
        for key, (name, headers) in layouts.items():
            sheet = book[name]
            check_headers(sheet, headers)
            records = {}
            for row in sheet.iter_rows(min_row=2, max_col=len(headers), values_only=True):
                period = period_key(row[0], key in ('prime', 'hibor'))
                if period is None:
                    continue
                if period in records:
                    raise ValueError(f'{name}: duplicate observation {period}')
                values = [None if value in (None, '') else value for value in row[1:]]
                if any(value is not None and (type(value) not in (int, float) or not math.isfinite(value)) for value in values):
                    raise ValueError(f'{name} {period}: invalid saved value')
                records[period] = [period, *values]
            series[key] = [records[period] for period in sorted(records)]
        if not any(row[1] is not None for row in series['nonbank']):
            raise ValueError('非银行公式没有已保存的结果；请先运行更新或在 Excel 中计算后保存。')
    finally:
        book.close()
    if path.read_bytes() != content:
        raise RuntimeError('Excel 在同步期间发生变动，请保存后重试。')
    return {
        'schemaVersion': 1, 'revision': hashlib.sha256(content).hexdigest(),
        'workbook': path.name,
        'savedAt': datetime.fromtimestamp(path.stat().st_mtime, timezone.utc).isoformat(),
        'series': series, 'stale': False,
    }


def update_dashboard(path):
    """Atomically replace the embedded data in the portable dashboard file."""
    destination = BASE / 'HK_Mortgage_Dashboard.html'
    if not destination.is_file():
        raise FileNotFoundError('请把 HK_Mortgage_Dashboard.html 与 Python 文件放在同一文件夹。')
    payload = dashboard_snapshot(path)
    payload['generatedAt'] = hk_now().isoformat()
    pattern = r'(<script id="workbook-data" type="application/json">)(.*?)(</script>)'
    template = destination.read_text(encoding='utf-8')
    if len(re.findall(pattern, template, re.DOTALL)) != 1:
        raise ValueError('Dashboard 文件的数据区损坏，已保留原文件。')
    encoded = json.dumps(payload, ensure_ascii=False, allow_nan=False, separators=(',', ':'))
    for char, escaped in (('<', '\\u003c'), ('>', '\\u003e'), ('&', '\\u0026'),
                          ('\u2028', '\\u2028'), ('\u2029', '\\u2029')):
        encoded = encoded.replace(char, escaped)
    html = re.sub(pattern, lambda match: match[1] + encoded + match[3], template, flags=re.DOTALL)
    verified = json.loads(re.search(pattern, html, re.DOTALL)[2])
    if verified != payload:
        raise ValueError('Dashboard 数据核对失败，已保留原文件。')
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', newline='\n',
                                         dir=BASE, prefix='.dashboard-', suffix='.tmp', delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(html)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, destination)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
    receipt = {'status':'updated', 'mode':'local', 'file':destination.name,
               'revision':payload['revision'], 'checked_at':payload['generatedAt']}
    (BASE / 'dashboard_status.json').write_text(json.dumps(receipt, ensure_ascii=False, indent=2), encoding='utf-8')
    message = f'本地 Dashboard 已更新并逐项核对：{destination.name}'
    with (BASE / 'update_status.txt').open('a', encoding='utf-8') as feedback:
        feedback.write('\n' + message + '\n')
    print(message, flush=True)
    return True


def main():
    global START_YEAR, REFRESH_HISTORY, PREVIOUS
    started=time.monotonic()
    ISSUES.clear();EVIDENCE.clear();TABLE_RESULTS.clear();MACRO_ROWS.clear();MACRO_FRESH.clear();MACRO_CHECKS.clear()
    parser = argparse.ArgumentParser()
    parser.add_argument("--landreg-months", type=int, default=24,
                        help="土地注册处回溯月数 (默认 24)")
    parser.add_argument('--start-year', type=int, default=2012, help='中原历史起始年，默认2012，与原表现楼/楼花一致')
    parser.add_argument('--refresh-history', action='store_true', help='重新检查所有历史月报；默认复用已提取历史，更新近期及缺值')
    parser.add_argument('--output', default=str(BASE / OUT_XLSX), help='输出Excel路径')
    parser.add_argument('--recalculate-only', action='store_true', help='只重算现有 Excel 和本地 Dashboard，不抓取新数据')
    parser.add_argument('--dashboard-only', action='store_true', help='只从已保存 Excel 更新本地 Dashboard，不联网')
    parser.add_argument('--no-dashboard', action='store_true', help='只更新 Excel')
    parser.add_argument('--sync-only', dest='dashboard_only', action='store_true', help=argparse.SUPPRESS)
    parser.add_argument('--no-sync', dest='no_dashboard', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.dashboard_only:
        if args.no_dashboard or args.recalculate_only:
            parser.error('--dashboard-only 不能与其他运行模式同时使用')
        update_dashboard(args.output)
        return
    if args.recalculate_only:
        recalculate_existing(args.output)
        if not args.no_dashboard:
            update_dashboard(args.output)
        return
    if not Path(args.output).is_file():
        parser.error(f'请先放入要更新的工作簿：{args.output}')
    if not 1900 <= args.start_year <= hk_now().year or not 1 <= args.landreg_months <= 600:
        parser.error('Invalid start year or month range')
    START_YEAR = args.start_year
    REFRESH_HISTORY = args.refresh_history
    initial_workbook = Path(args.output).read_bytes()
    previous = load_existing(args.output)
    PREVIOUS = previous['中原按揭']
    evidence_path=BASE/'source_evidence.json'
    existing_evidence=json.loads(evidence_path.read_text(encoding='utf-8')) if evidence_path.exists() else []
    protected=verified_table_values(PREVIOUS,existing_evidence)
    protected_dated={(month,field):value for month,field,value in dated_source_values(existing_evidence)
                     if PREVIOUS.get(month,{}).get(field)==value}

    print("=" * 70)
    print("香港住宅按揭数据自动采集与整合")
    print("=" * 70)

    print("\n[1/4] 拉取中原按揭 (现楼/楼花, 2012至今)...")
    api = fetch_centaline()
    centaline = merge_centaline_api(PREVIOUS, api, existing_evidence)
    print(f"  中原按揭(GetStatistics): {len(centaline)} 个月")

    # detail 页兜底: 补最新月 + 私人住宅现楼/资助房屋
    fallback = fetch_centaline_detail_fallback()
    table_values={(r['month'],r['field'],r['value']) for r in EVIDENCE if r.get('retrieval')=='official_attachment'}
    dated_values=dated_source_values(EVIDENCE)
    for k, values in fallback.items():
        row = centaline.setdefault(k, {})
        for field, value in values.items():
            if value is None:
                continue
            if (k,field) in protected and (k,field,value) not in table_values:
                continue  # A failed attachment fetch must not undo a verified historical correction.
            if (k,field) in protected_dated and (k,field,value) not in dated_values:
                continue
            if api.get(k, {}).get(field) is not None and api[k][field] != value:
                if (k,field,value) in dated_values:
                    row[field]=value
                    origin='official attachment table' if (k,field,value) in table_values else 'explicit dated Centaline report'
                    ISSUES.append({'source':'API vs report','period':k,'field':field,'error':f'{api[k][field]} vs {value}; used {origin}','resolved':True})
                else:
                    ISSUES.append({'source':'API vs report','period':k,'field':field,'error':f'{api[k][field]} vs {value}; no corroborating table, kept API'})
            else:
                row[field] = value
    centaline = {k:v for k,v in centaline.items() if int(k[:4]) >= START_YEAR and any(isinstance(x,(int,float)) for x in v.values())}
    now=hk_now()
    last_completed=month_key((now.year*12+now.month-2)//12,(now.year*12+now.month-2)%12+1)
    centaline=align_months(centaline, f'{START_YEAR}-01', min(max(centaline),last_completed))
    print(f"  中原按揭(含 detail 兜底): {len(centaline)} 个月")

    print("\n[2/4] 拉取 HKMA 住宅按揭统计 (表 T0307)...")
    hkma_fresh = fetch_hkma()
    hkma = merge_rows(previous['HKMA'], hkma_fresh)
    print(f"  HKMA: {len(hkma)} 个月")

    print(f"\n[3/4] 拉取土地注册处按揭登记 (最近 {args.landreg_months} 个月)...")
    landreg_fresh = fetch_landreg(args.landreg_months)
    landreg = merge_rows(previous['土地注册处'], landreg_fresh)
    print(f"  土地注册处: {len(landreg)} 个月")

    print('\n[4/4] 更新 M3、HSBC Prime 和每日 HIBOR...',flush=True)
    collect_macro(previous)

    print("\n写入 Excel...")
    if not centaline or not hkma_fresh or not landreg_fresh:
        raise RuntimeError('A required source returned no data; existing workbook has not been overwritten')
    for month, row in centaline.items():
        total, private, subsidised = [row.get(k) for k in ['现楼按揭宗数','私人住宅现楼宗数','资助房屋宗数']]
        if total is not None and private is not None and private > total:
            ISSUES.append({'source':'Cross-source check','period':month,'error':f'Private residential {private} exceeds API completed total {total}; original source values retained, not suitable for automatic residual calculation'})
        if all(x is not None for x in [total,private,subsidised]) and private+subsidised > total:
            ISSUES.append({'period':month,'error':'Private + subsidised exceeds completed total; review sources'})
    out_path = args.output
    write_xlsx(centaline, hkma, landreg, out_path, expected_original=initial_workbook)
    evidence_path = BASE / 'source_evidence.json'
    existing_evidence = json.loads(evidence_path.read_text(encoding='utf-8')) if evidence_path.exists() else []
    unique = {(r['month'],r['field'],r['url']):r for r in existing_evidence + EVIDENCE
              if centaline.get(r['month'],{}).get(r['field'])==r['value']}
    evidence_path.write_text(json.dumps(list(unique.values()),ensure_ascii=False,indent=2),encoding='utf-8')
    report = {'checked_at':hk_now().isoformat(timespec='seconds'), 'issues':ISSUES,
              'centaline_months':len(centaline),
              'private_months':sum(r.get('私人住宅现楼宗数') is not None for r in centaline.values()),
              'private_missing':[k for k,v in sorted(centaline.items()) if v.get('私人住宅现楼宗数') is None],
              'new_private_months':[k for k,v in sorted(centaline.items()) if v.get('私人住宅现楼宗数') is not None and PREVIOUS.get(k,{}).get('私人住宅现楼宗数') is None]}
    report['status']='partial' if ISSUES or report['private_missing'] else 'complete'
    report['duration_seconds']=round(time.monotonic()-started,1)
    report['coverage']={name:{'start':min(rows),'end':max(rows),'months':len(rows)} for name,rows in [('Centaline',centaline),('HKMA',hkma),('Land Registry',landreg)]}
    report['missing_by_field']={field:[k for k,v in centaline.items() if v.get(field) is None] for field in FIELDS}
    report['macro_checks']=MACRO_CHECKS
    fresh_centaline=merge_rows(api,fallback)
    current={'中原按揭':centaline,'HKMA':hkma,'土地注册处':landreg,**MACRO_ROWS}
    fresh={'中原按揭':fresh_centaline,'HKMA':hkma_fresh if HKMA_REMOTE_OK else {},'土地注册处':landreg_fresh,**MACRO_FRESH}
    report['update_feedback']=update_feedback(previous,current,fresh)
    report['status']='partial' if ISSUES or any(report['missing_by_field'].values()) or any(v['target_missing'] for v in report['update_feedback'].values()) else 'complete'
    (BASE / 'collection_report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    notices=[f'检查时间（香港）：{hk_now():%Y-%m-%d %H:%M:%S}']
    for name,state in report['update_feedback'].items():
        line=f"{name}：已有数据至 {state['latest']}；本次新增 {state['new_periods']} 个日期/月；补齐 {state['filled_values']} 个空值；修订 {state['revised_values']} 个数值。"
        if name=='HSBC Prime':
            line=line.replace('已有数据至','最新实际生效日')
            if 'current_rate' in MACRO_CHECKS.get(name,{}):line+=f" 官网本次核对的当前利率：{MACRO_CHECKS[name]['current_rate']}%。"
        if state['target_missing']:line+=f" 提醒：目标 {state['target']} 的数据尚未取得（可能尚未发布或抓取失败），请检查日志。"
        elif state['fresh_target_missing']:line+=' 提醒：目标日期的源数据本次未核实，现有历史值已保留。'
        elif not state['new_periods'] and not state['revised_values'] and not state['filled_values']:line+=' 本次未发现新数据，已重新检查来源。'
        if not state['fresh_records']:line+=' 提醒：本次没有取到可用源数据，当前显示的是保留的历史。'
        notices.append(line)
    notices.append('缺值/冲突详情见 collection_report.json；本文件只作反馈，不写入 Excel。')
    (BASE/'update_status.txt').write_text('\n'.join(notices),encoding='utf-8')
    print('\n'+'\n'.join(notices),flush=True)
    print(f"私人住宅现楼: {report['private_months']}/{len(centaline)}个月；本次补入{len(report['new_private_months'])}个月；异常/冲突{len(ISSUES)}项", flush=True)

    # 简要预览
    print("\n" + "=" * 70)
    print("最近 6 个月预览 (中原现楼/楼花 vs HKMA新取用 vs 土地注册处其他按揭)")
    print("=" * 70)
    keys = sorted(set(centaline) | set(hkma) | set(landreg))[-6:]
    print(f"{'月份':<9}{'现楼':>7}{'楼花':>7}{'HKMA取用':>10}{'其他按揭':>9}")
    for k in keys:
        c = centaline.get(k, {})
        h = hkma.get(k, {})
        l = landreg.get(k, {})
        def f(x):
            return f"{x:,}" if x is not None else "-"
        print(f"{k:<9}{f(c.get('现楼按揭宗数')):>7}{f(c.get('楼花按揭宗数')):>7}"
              f"{f(h.get('新取用宗数')):>10}{f(l.get('其他按揭法定押记宗数')):>9}")

    if not args.no_dashboard:
        update_dashboard(args.output)


def update_feedback(previous,current,fresh,now=None):
    now=now or hk_now(); month_end=now.date().replace(day=1)-timedelta(days=1)
    feedback={}
    for name,rows in current.items():
        prior=previous.get(name,{}); latest=max(rows) if rows else None
        target=month_end.strftime('%Y-%m') if name in ['中原按揭','HKMA','土地注册处','M3'] else None
        if name=='HIBOR Daily':target=MACRO_CHECKS.get(name,{}).get('expected_date')
        required=FIELDS if name=='中原按揭' else ['新取用宗数','新取用金额_HK$Mn'] if name=='HKMA' else ['楼宇按揭宗数','其他按揭法定押记宗数','撤销按揭宗数'] if name=='土地注册处' else MACRO_LAYOUT.get(name,[])[1:]
        missing=[field for field in required if target and rows.get(target,{}).get(field) is None]
        fresh_missing=bool(target and any(fresh.get(name,{}).get(target,{}).get(field) is None for field in required))
        def changed(a,b):
            if isinstance(a,(int,float)) and isinstance(b,(int,float)):
                return not math.isclose(a,b,rel_tol=1e-12,abs_tol=1e-9)
            return a!=b
        revised=sum(1 for k,v in rows.items() for field,value in v.items() if value is not None and k in prior and prior[k].get(field) is not None and changed(prior[k][field],value))
        filled=sum(1 for k,v in rows.items() for field,value in v.items() if isinstance(value,(int,float)) and k in prior and prior[k].get(field) is None)
        feedback[name]={'latest':latest,'target':target,'target_missing':bool(missing or target and (latest is None or latest<target)),
                        'target_missing_fields':missing,'fresh_target_missing':fresh_missing,
                        'new_periods':len(set(rows)-set(prior)),'filled_values':filled,'revised_values':revised,'fresh_records':len(fresh.get(name,{}))}
    return feedback


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        failure={'status':'failed','checked_at':hk_now().isoformat(timespec='seconds'),
                 'error':str(exc),'issues':ISSUES}
        (BASE/'last_failure.json').write_text(json.dumps(failure,ensure_ascii=False,indent=2),encoding='utf-8')
        (BASE/'update_status.txt').write_text(f'更新失败：{hk_now().isoformat()}\n{exc}\n请查看 last_failure.json。',encoding='utf-8')
        print(f'更新未完成：{exc}\n详情见 last_failure.json',file=sys.stderr)
        sys.exit(1)
