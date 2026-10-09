#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
农夫开票明细 · 网站自动抓取工具
==============================
功能：
  登录供应链系统（默认 http://gyl.zyred.com，可在 config.json 的 base_url 字段覆盖），
  自动抓取「购进查询」「返厂查询」的商品明细数据，
  输出与人工复制格式一致的 xlsx（正单 / 返单 两个 sheet）。

用法：
  python scrape.py --config config.json
  python scrape.py --user 你的登录账号 --password 你的密码 \
                   --gysbm 你的供应商编码 --bmbm 你的商户编码 \
                   --start-date 20250901 --end-date 20261009 \
                   --output 开票明细_网站抓取.xlsx

依赖：requests、openpyxl、beautifulsoup4（首次运行缺失会自动尝试安装）

注意：账号密码为敏感信息，请通过 config.json 或在命令行临时传入，
      切勿将包含真实密码的 config.json 提交到 git 或转发给他人。
"""

import sys
import os
import json
import re
import argparse
import subprocess
from calendar import monthrange

# ---------- 依赖自动安装 ----------
_missing = []
for _pkg in ['requests', 'openpyxl', 'bs4']:
    try:
        __import__(_pkg)
    except ImportError:
        _missing.append(_pkg)
if _missing:
    print("未检测到依赖 {}，正在尝试自动安装……".format(", ".join(_missing)))
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + _missing)
    except Exception as e:
        print("自动安装失败，请手动执行：pip install {}".format(" ".join(_missing)))
        print("错误：", e)
        sys.exit(1)

import requests
from bs4 import BeautifulSoup
import openpyxl


# 目标系统地址；如系统地址变更，可在此修改，或放入 config.json 的 base_url 字段。
BASE_URL = "http://gyl.zyred.com"
LOGIN_URL = BASE_URL + "/bymagq/login.do"
GOUJIN_QUERY_URL = BASE_URL + "/query/queryGoujinAction.do"
FANCHANG_QUERY_URL = BASE_URL + "/query/queryFanchangAction.do"


def clean_text(v):
    if v is None:
        return ""
    return str(v).replace('\xa0', '').replace('\u200b', '').strip()


def to_number(v):
    """把字符串转成 int/float；失败返回原值。"""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return v
    s = str(v).replace(',', '').replace('######', '').strip()
    if s == '':
        return None
    try:
        if '.' in s:
            return float(s)
        return int(s)
    except ValueError:
        return v


def parse_table(html_text, title_keywords):
    """在 HTML 中查找标题包含指定关键字的表格，返回 (header, data_rows)。
    如果找到的是标题表（只有 1 行 1 格），则取下一个同层 table 作为数据表。
    自动跳过合计行。"""
    soup = BeautifulSoup(html_text, 'html.parser')
    tables = soup.find_all('table')
    for idx, tbl in enumerate(tables):
        rows = tbl.find_all('tr')
        if not rows:
            continue
        first_text = clean_text(rows[0].get_text())
        if any(kw in first_text for kw in title_keywords):
            # 若该表是标题表（第一行只有 1 格且为标题文字），取紧随其后的数据表
            if len(rows[0].find_all(['th', 'td'])) == 1 and idx + 1 < len(tables):
                tbl = tables[idx + 1]
                rows = tbl.find_all('tr')
            header = [clean_text(c.get_text()) for c in rows[0].find_all(['th', 'td'])]
            data = []
            for r in rows[1:]:
                cells = [clean_text(c.get_text()) for c in r.find_all(['td', 'th'])]
                if not cells or all(c == '' for c in cells):
                    continue
                # 跳过合计/总计行：第一格包含“合计”或“总计”
                if cells and ('合计' in cells[0] or '总计' in cells[0]):
                    continue
                data.append(cells)
            return header, data
    return [], []


def login(session, username, password):
    """登录系统并返回会话。"""
    session.get(LOGIN_URL, timeout=20)
    r = session.post(LOGIN_URL, data={
        'person_name': username,
        'password': password
    }, timeout=20, allow_redirects=True)
    if '供应链' not in r.text and 'frameset' not in r.text:
        raise RuntimeError("登录失败：返回页面异常，请检查账号密码。")
    return session


def fetch_order_summary(session, url, params, title_keywords):
    """抓取订单/返厂汇总页，返回 (header, rows, html_text)。"""
    r = session.post(url, data=params, timeout=60, allow_redirects=True)
    if r.status_code != 200:
        raise RuntimeError("查询失败，HTTP {}".format(r.status_code))
    header, rows = parse_table(r.text, title_keywords)
    return header, rows, r.text


def extract_detail_links(html_text):
    """从汇总页 HTML 中提取所有 detail.jsp 的相对/绝对链接。
    返回绝对 URL 列表（基于当前 BASE_URL）。"""
    base_context = BASE_URL + "/query/"
    links = []
    for m in re.finditer(r'href\s*=\s*"([^"]+\.jsp\?[^"]*)"', html_text, re.IGNORECASE):
        href = m.group(1)
        if not href.startswith('http'):
            # 相对路径；汇总页本身在 /query/queryXxxAction.do，因此 detail.jsp 在同一目录
            href = base_context + href
        links.append(href)
    return links


def fetch_detail_rows(session, detail_url, title_keywords):
    """抓取一个明细页，返回该页的商品明细行（已经是清洗后的字符串列表）。"""
    r = session.get(detail_url, timeout=20, allow_redirects=True)
    header, rows = parse_table(r.text, title_keywords)
    return header, rows


def parse_detail_row_to_sheet(row, expected_header):
    """根据明细页表头映射到目标 sheet 的 8 列。"""
    # 购进目标列：商品编码, 商品名称, 规格型号, 计量单位, 入库数量, 不含税价, 含税价, 入库金额含税
    # 返厂目标列：商品编码, 商品名称, 规格型号, 计量单位, 返厂数量, 不含税价, 含税价, 返厂金额含税
    mapping = {}
    for i, h in enumerate(expected_header):
        h = h.replace(' ', '')
        if '商品编码' in h:
            mapping['code'] = i
        elif '商品名称' in h:
            mapping['name'] = i
        elif '规格型号' in h:
            mapping['spec'] = i
        elif '计量单位' in h:
            mapping['unit'] = i
        elif h in ('入库数量', '返厂数量'):
            mapping['qty'] = i
        elif '不含税价' in h:
            mapping['price_no_tax'] = i
        elif '含税价' in h and '不含税' not in h:
            mapping['price_tax'] = i
        elif '入库金额含税' in h or '返厂金额含税' in h:
            mapping['amount'] = i

    def get(key):
        idx = mapping.get(key)
        return row[idx] if idx is not None and idx < len(row) else None

    return [
        get('code'), get('name'), get('spec'), get('unit'),
        to_number(get('qty')),
        to_number(get('price_no_tax')),
        to_number(get('price_tax')),
        to_number(get('amount')),
    ]


def scrape_goujin(session, config):
    """抓取购进（正单）明细，返回 [(row), ...] 和表头。"""
    params = {
        't_start': config['start_date'],
        't_end': config['end_date'],
        'bmbm': config['bmbm'],
        'gysbm': config['gysbm'],
        'cxms': '1',
        'button1': '查询',
        'spbm': '', 'rkdh': '', 'ISBN': '', 'BookName': '',
        'mdbm': '%', 'lcbm': '%', 'dlbm': '', 'd_wljk': '0'
    }
    header, rows, html = fetch_order_summary(session, GOUJIN_QUERY_URL, params,
                                              ['商品入库汇总', '按单号'])
    print("购进汇总：{} 个订单".format(len(rows)))

    detail_links = extract_detail_links(html)
    print("购进明细页链接数：{}".format(len(detail_links)))

    sheet_header = ['商品编码', '商品名称', '规格型号', '计量单位', '入库数量', '不含税价', '含税价', '入库金额含税']
    all_rows = []
    for idx, url in enumerate(detail_links, 1):
        try:
            _, drows = fetch_detail_rows(session, url, ['商品入库明细', '按单号'])
            for r in drows:
                all_rows.append(parse_detail_row_to_sheet(r, ['商品编码', '商品名称', '规格型号', '计量单位', '入库数量', '不含税价', '含税价', '入库金额含税', '入库金额不含税', '入库时间', '属性', '门店编码', '订货单号']))
            print("  [{}] {} -> {} 行".format(idx, url.split('?')[1].split('&')[0], len(drows)))
        except Exception as e:
            print("  [{}] 抓取失败 {}: {}".format(idx, url, e))
    return sheet_header, all_rows


def scrape_fanchang(session, config):
    """抓取返厂（返单）明细，返回 [(row), ...] 和表头。"""
    params = {
        't_start': config['start_date'],
        't_end': config['end_date'],
        'bmbm': config['bmbm'],
        'gysbm': config['gysbm'],
        'cxms': '1',
        'button1': '查询',
        'spbm': '', 'fcdh': '', 'ISBN': '', 'BookName': '',
        'mdbm': '%', 'lcbm': '%', 'dlbm': '', 'd_wljk': '0'
    }
    header, rows, html = fetch_order_summary(session, FANCHANG_QUERY_URL, params,
                                              ['商品返厂汇总', '按单号'])
    print("返厂汇总：{} 个订单".format(len(rows)))

    detail_links = extract_detail_links(html)
    print("返厂明细页链接数：{}".format(len(detail_links)))

    sheet_header = ['商品编码', '商品名称', '规格型号', '计量单位', '返厂数量', '不含税价', '含税价', '返厂金额含税']
    all_rows = []
    for idx, url in enumerate(detail_links, 1):
        try:
            _, drows = fetch_detail_rows(session, url, ['商品返厂明细', '按单号'])
            for r in drows:
                all_rows.append(parse_detail_row_to_sheet(r, ['商品编码', '商品名称', '规格型号', '计量单位', '返厂数量', '不含税价', '含税价', '返厂金额含税', '返厂金额不含税', '返厂时间', '属性', '门店编码']))
            print("  [{}] {} -> {} 行".format(idx, url.split('?')[1].split('&')[0], len(drows)))
        except Exception as e:
            print("  [{}] 抓取失败 {}: {}".format(idx, url, e))
    return sheet_header, all_rows


def write_xlsx(output_path, zheng_header, zheng_rows, fan_header, fan_rows):
    """写入 xlsx，sheet 名为「正单」「返单」。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '正单'
    ws.append(zheng_header)
    for r in zheng_rows:
        ws.append(r)

    ws2 = wb.create_sheet('返单')
    ws2.append(fan_header)
    for r in fan_rows:
        ws2.append(r)

    wb.save(output_path)
    print("已保存原始抓取文件：{}".format(output_path))
    print("  正单行数：{}，返单行数：{}".format(len(zheng_rows), len(fan_rows)))


def load_config():
    parser = argparse.ArgumentParser(description='自动抓取供应链系统正单/返单数据')
    parser.add_argument('--config', '-c', help='JSON 配置文件路径')
    parser.add_argument('--user', '-u', help='登录用户名')
    parser.add_argument('--password', '-p', help='登录密码')
    parser.add_argument('--gysbm', help='供应商编码（由系统提供）')
    parser.add_argument('--bmbm', help='商户编码（必填，由系统提供）')
    parser.add_argument('--base-url', help='系统地址，默认 http://gyl.zyred.com')
    parser.add_argument('--start-date', help='起始日期 YYYYMMDD')
    parser.add_argument('--end-date', help='结束日期 YYYYMMDD')
    parser.add_argument('--month', '-m', help='便捷指定月份，格式 YYYYMM（自动取该月第一天到最后一天）')
    parser.add_argument('--output', '-o', help='输出 xlsx 路径')
    parser.add_argument('--no-consolidate', action='store_true', help='不执行按品类合计')
    args = parser.parse_args()

    cfg = {}
    if args.config:
        with open(args.config, 'r', encoding='utf-8') as f:
            cfg = json.load(f)

    # CLI overrides config
    for key in ['user', 'password', 'gysbm', 'bmbm', 'base_url',
                'start_date', 'end_date', 'output']:
        val = getattr(args, key)
        if val:
            cfg[key] = val

    # --month 便捷指定月份（CLI 或 config 均可）
    month_val = args.month or cfg.get('month')
    if month_val:
        y, m = int(month_val[:4]), int(month_val[4:])
        cfg['start_date'] = month_val + '01'
        cfg['end_date'] = month_val + str(monthrange(y, m)[1])

    # defaults
    cfg.setdefault('base_url', BASE_URL)

    required = ['user', 'password', 'gysbm', 'bmbm', 'start_date', 'end_date', 'output']
    missing = [k for k in required if not cfg.get(k)]
    if missing:
        parser.error("缺少必要参数：{}".format(", ".join(missing)))

    return cfg, args.no_consolidate


def main():
    cfg, no_consolidate = load_config()

    # 应用可配置的系统地址
    global BASE_URL, LOGIN_URL, GOUJIN_QUERY_URL, FANCHANG_QUERY_URL
    BASE_URL = cfg.get('base_url', BASE_URL)
    LOGIN_URL = BASE_URL + "/bymagq/login.do"
    GOUJIN_QUERY_URL = BASE_URL + "/query/queryGoujinAction.do"
    FANCHANG_QUERY_URL = BASE_URL + "/query/queryFanchangAction.do"

    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    })

    print("登录中...")
    login(session, cfg['user'], cfg['password'])
    print("登录成功。")

    print("\n【1/2】抓取购进（正单）明细...")
    zheng_header, zheng_rows = scrape_goujin(session, cfg)

    print("\n【2/2】抓取返厂（返单）明细...")
    fan_header, fan_rows = scrape_fanchang(session, cfg)

    write_xlsx(cfg['output'], zheng_header, zheng_rows, fan_header, fan_rows)

    if not no_consolidate:
        try:
            # 自动调用 consolidate 生成合并表
            from consolidate import consolidate
            base, ext = os.path.splitext(cfg['output'])
            merged_path = base + '_按品类合计' + (ext or '.xlsx')
            consolidate(cfg['output'], merged_path)
        except Exception as e:
            print("按品类合计失败，需手动运行 consolidate.py：{}".format(e))


if __name__ == '__main__':
    main()
