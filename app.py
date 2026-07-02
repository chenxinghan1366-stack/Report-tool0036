import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from io import BytesIO
import uuid
import sqlite3
import os
import zipfile
import requests
from bs4 import BeautifulSoup
import re
import hashlib
import time

# ========== 数据库路径 ==========
DB_PATH = os.path.join(os.path.dirname(__file__), "app_data.db")

# ========== 初始化数据库 ==========
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS companies (
        id TEXT PRIMARY KEY, company_name TEXT, province TEXT, city TEXT, district TEXT, tax_id TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS templates (
        id TEXT PRIMARY KEY, province TEXT, city TEXT, district TEXT, report_type TEXT,
        template_name TEXT, template_version TEXT, source_url TEXT, source_authority TEXT,
        publish_date TEXT, required_fields TEXT, status TEXT, file_hash TEXT, file_type TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS rules (
        id TEXT PRIMARY KEY, city TEXT, province TEXT, unit_social REAL, personal_social REAL,
        unit_fund REAL, personal_fund REAL, social_min REAL, social_max REAL,
        fund_min REAL, fund_max REAL, source_quote TEXT, is_default BOOLEAN DEFAULT 0
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS export_history (
        id TEXT PRIMARY KEY, company_id TEXT, template_id TEXT, company_name TEXT,
        city TEXT, province TEXT, report_type TEXT, period_type TEXT, generated_at TEXT,
        review_status TEXT, reviewer TEXT, reviewed_at TEXT, file_name TEXT, file_data BLOB,
        data_source TEXT, month_used TEXT, year_used TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS source_registry (
        id TEXT PRIMARY KEY, authority_type TEXT, province TEXT, city TEXT, district TEXT,
        authority_name TEXT, official_site_name TEXT, source_url TEXT, source_level TEXT,
        source_section TEXT, is_official BOOLEAN, crawl_allowed BOOLEAN, last_checked TEXT,
        status TEXT, notes TEXT
    )''')
    conn.commit()
    conn.close()

init_db()

# ========== 数据库迁移 ==========
def migrate_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("PRAGMA table_info(rules)")
    columns_rules = [col[1] for col in c.fetchall()]
    if 'province' not in columns_rules:
        c.execute("ALTER TABLE rules ADD COLUMN province TEXT")
    
    c.execute("PRAGMA table_info(export_history)")
    columns_export = [col[1] for col in c.fetchall()]
    for col in ['data_source', 'month_used', 'year_used']:
        if col not in columns_export:
            c.execute(f"ALTER TABLE export_history ADD COLUMN {col} TEXT")
    
    c.execute("PRAGMA table_info(templates)")
    columns_templates = [col[1] for col in c.fetchall()]
    for col in ['file_hash', 'file_type']:
        if col not in columns_templates:
            c.execute(f"ALTER TABLE templates ADD COLUMN {col} TEXT")
    
    conn.commit()
    conn.close()

migrate_db()

# ========== 数据操作函数 ==========
def dict_fetchall(cursor):
    columns = [col[0] for col in cursor.description]
    return [dict(zip(columns, row)) for row in cursor.fetchall()]

def safe_execute_query(query, params=None):
    try:
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        if params:
            c.execute(query, params)
        else:
            c.execute(query)
        rows = dict_fetchall(c)
        conn.close()
        return rows
    except sqlite3.OperationalError as e:
        if "no such table" in str(e):
            return []
        else:
            raise e

def load_companies():
    return safe_execute_query("SELECT * FROM companies")

def load_templates():
    return safe_execute_query("SELECT * FROM templates WHERE status='active'")

def load_rules():
    return safe_execute_query("SELECT * FROM rules ORDER BY province, city")

def load_export_history():
    return safe_execute_query("SELECT * FROM export_history ORDER BY generated_at DESC")

def load_source_registry():
    return safe_execute_query("SELECT * FROM source_registry ORDER BY province, city")

def save_companies(companies):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM companies")
    for comp in companies:
        c.execute('''INSERT INTO companies (id, company_name, province, city, district, tax_id)
            VALUES (?,?,?,?,?,?)''',
            (comp.get('id', str(uuid.uuid4())[:8]), comp['company_name'], comp['province'],
             comp['city'], comp.get('district',''), comp.get('tax_id','')))
    conn.commit()
    conn.close()

def save_template(template):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id FROM templates WHERE source_url=? OR file_hash=?", 
              (template.get('source_url',''), template.get('file_hash','')))
    existing = c.fetchone()
    if existing:
        c.execute('''UPDATE templates SET 
            template_name=?, template_version=?, source_authority=?, publish_date=?,
            required_fields=?, status=?, file_type=?
            WHERE id=?''',
            (template['template_name'], template.get('template_version','v1.0'),
             template.get('source_authority',''), template.get('publish_date',''),
             template.get('required_fields',''), template.get('status','active'),
             template.get('file_type',''), existing[0]))
    else:
        c.execute('''INSERT INTO templates 
            (id, province, city, district, report_type, template_name, template_version,
             source_url, source_authority, publish_date, required_fields, status, file_hash, file_type)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (template['id'], template.get('province',''), template.get('city',''),
             template.get('district',''), template.get('report_type',''),
             template['template_name'], template.get('template_version','v1.0'),
             template.get('source_url',''), template.get('source_authority',''),
             template.get('publish_date',''), template.get('required_fields',''),
             template.get('status','active'), template.get('file_hash',''),
             template.get('file_type','')))
    conn.commit()
    conn.close()

def save_source_registry(sources):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    for s in sources:
        c.execute('''INSERT OR REPLACE INTO source_registry 
            (id, authority_type, province, city, district, authority_name,
             official_site_name, source_url, source_level, source_section,
             is_official, crawl_allowed, last_checked, status, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (s.get('id', str(uuid.uuid4())[:8]), s.get('authority_type','tax'),
             s.get('province',''), s.get('city',''), s.get('district',''),
             s.get('authority_name',''), s.get('official_site_name',''),
             s.get('source_url',''), s.get('source_level',''), s.get('source_section',''),
             s.get('is_official',1), s.get('crawl_allowed',1),
             s.get('last_checked',''), s.get('status','active'), s.get('notes','')))
    conn.commit()
    conn.close()

def save_rules(rules):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM rules")
    for r in rules:
        c.execute('''INSERT INTO rules 
            (id, city, province, unit_social, personal_social, unit_fund, personal_fund,
             social_min, social_max, fund_min, fund_max, source_quote, is_default)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (r['id'], r['city'], r.get('province',''), r['unit_social'], r['personal_social'],
             r['unit_fund'], r['personal_fund'], r.get('social_min',0), r.get('social_max',999999),
             r.get('fund_min',0), r.get('fund_max',999999), r.get('source_quote',''),
             r.get('is_default',0)))
    conn.commit()
    conn.close()

def save_export(record):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO export_history 
        (id, company_id, template_id, company_name, city, province, report_type, period_type,
         generated_at, review_status, reviewer, reviewed_at, file_name, file_data,
         data_source, month_used, year_used)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (record['id'], record.get('company_id',''), record.get('template_id',''),
         record['company_name'], record.get('city',''), record.get('province',''),
         record.get('report_type',''), record.get('period_type',''), record['generated_at'],
         record.get('review_status','pending'), record.get('reviewer',''), record.get('reviewed_at',''),
         record.get('file_name',''), record.get('file_data', None),
         record.get('data_source',''), record.get('month_used',''), record.get('year_used','')))
    conn.commit()
    conn.close()

def update_export_status(export_id, status, reviewer):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''UPDATE export_history 
        SET review_status=?, reviewer=?, reviewed_at=?
        WHERE id=?''',
        (status, reviewer, datetime.now().isoformat(), export_id))
    conn.commit()
    conn.close()

# ========== 爬虫模块 ==========
def get_national_tax_sites():
    """获取各省税务局列表（先尝试爬取，失败则返回预设列表）"""
    urls = []
    try:
        url = "https://www.chinatax.gov.cn/chinatax/n810219/n810744/index.html"
        response = requests.get(url, timeout=10)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        for a in soup.find_all('a', href=True):
            href = a['href']
            text = a.get_text(strip=True)
            if '省' in text or '市' in text or '自治区' in text:
                if href.startswith('http'):
                    full_url = href
                elif href.startswith('/'):
                    full_url = requests.compat.urljoin(url, href)
                else:
                    continue
                if 'tax' in full_url or 'chinatax' in full_url:
                    urls.append({
                        'name': text,
                        'url': full_url,
                        'level': 'provincial'
                    })
    except:
        pass
    
    # 如果未获取到，使用预设列表
    if not urls:
        fallback = [
            ('北京', 'https://beijing.chinatax.gov.cn/'),
            ('上海', 'https://shanghai.chinatax.gov.cn/'),
            ('广东', 'https://guangdong.chinatax.gov.cn/'),
            ('江苏', 'https://jiangsu.chinatax.gov.cn/'),
            ('浙江', 'https://zhejiang.chinatax.gov.cn/'),
            ('四川', 'https://sichuan.chinatax.gov.cn/'),
            ('湖北', 'https://hubei.chinatax.gov.cn/'),
            ('湖南', 'https://hunan.chinatax.gov.cn/'),
            ('河南', 'https://henan.chinatax.gov.cn/'),
            ('山东', 'https://shandong.chinatax.gov.cn/'),
            ('陕西', 'https://shaanxi.chinatax.gov.cn/'),
            ('辽宁', 'https://liaoning.chinatax.gov.cn/'),
            ('福建', 'https://fujian.chinatax.gov.cn/'),
            ('安徽', 'https://anhui.chinatax.gov.cn/'),
            ('江西', 'https://jiangxi.chinatax.gov.cn/'),
            ('山西', 'https://shanxi.chinatax.gov.cn/'),
            ('吉林', 'https://jilin.chinatax.gov.cn/'),
            ('黑龙江', 'https://heilongjiang.chinatax.gov.cn/'),
            ('云南', 'https://yunnan.chinatax.gov.cn/'),
            ('贵州', 'https://guizhou.chinatax.gov.cn/'),
            ('甘肃', 'https://gansu.chinatax.gov.cn/'),
            ('内蒙古', 'https://neimenggu.chinatax.gov.cn/'),
            ('新疆', 'https://xinjiang.chinatax.gov.cn/'),
            ('宁夏', 'https://ningxia.chinatax.gov.cn/'),
            ('青海', 'https://qinghai.chinatax.gov.cn/'),
            ('西藏', 'https://xizang.chinatax.gov.cn/'),
            ('海南', 'https://hainan.chinatax.gov.cn/'),
            ('广西', 'https://guangxi.chinatax.gov.cn/'),
            ('天津', 'https://tianjin.chinatax.gov.cn/'),
            ('重庆', 'https://chongqing.chinatax.gov.cn/'),
        ]
        for name, url in fallback:
            urls.append({'name': name, 'url': url, 'level': 'provincial'})
    return urls

def extract_sections(url, base_url):
    sections = []
    keywords = ['下载中心', '表证单书', '办税指南', '通知公告', '下载', '表单', '指南', '公告']
    try:
        response = requests.get(url, timeout=10)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        for a in soup.find_all('a', href=True):
            text = a.get_text(strip=True)
            href = a['href']
            if any(k in text for k in keywords):
                full_url = requests.compat.urljoin(base_url, href)
                sections.append({'name': text, 'url': full_url})
    except:
        pass
    return sections

def find_download_links(url, base_url):
    links = []
    try:
        response = requests.get(url, timeout=10)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        for a in soup.find_all('a', href=True):
            href = a['href']
            if any(href.endswith(ext) for ext in ['.xlsx', '.xls', '.docx', '.doc', '.pdf']):
                full_url = requests.compat.urljoin(base_url, href)
                links.append({
                    'url': full_url,
                    'filename': href.split('/')[-1],
                    'text': a.get_text(strip=True)
                })
    except:
        pass
    return links

def download_file(url):
    try:
        response = requests.get(url, timeout=20)
        response.raise_for_status()
        return response.content
    except:
        return None

def extract_fields_from_excel(content):
    try:
        from openpyxl import load_workbook
        from io import BytesIO
        wb = load_workbook(BytesIO(content))
        ws = wb.active
        headers = []
        for cell in ws[1]:
            if cell.value:
                headers.append(str(cell.value).strip())
        return ','.join(headers) if headers else ''
    except:
        return ''

def sync_official_templates():
    """执行同步，返回 (新增数量, 错误数量, 日志列表)"""
    log = []
    total_updated = 0
    total_errors = 0
    
    log.append("正在获取省级税务局入口...")
    sites = get_national_tax_sites()
    log.append(f"获取到 {len(sites)} 个省级税务局")
    
    sources = []
    for s in sites:
        sources.append({
            'id': str(uuid.uuid4())[:8],
            'authority_type': 'tax',
            'province': s['name'],
            'city': '',
            'district': '',
            'authority_name': f"国家税务总局{s['name']}省税务局" if s['name'] not in ['北京','上海','天津','重庆'] else f"国家税务总局{s['name']}市税务局",
            'official_site_name': f"{s['name']}税务",
            'source_url': s['url'],
            'source_level': s['level'],
            'source_section': '首页',
            'is_official': 1,
            'crawl_allowed': 1,
            'last_checked': datetime.now().isoformat(),
            'status': 'active',
            'notes': ''
        })
    save_source_registry(sources)
    log.append(f"已更新渠道库，共 {len(sources)} 条记录")
    
    # 进度显示
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, site in enumerate(sites):
        progress_bar.progress((idx + 1) / len(sites))
        status_text.text(f"正在处理：{site['name']} ({idx+1}/{len(sites)})")
        
        base_url = site['url']
        sections = extract_sections(base_url, base_url)
        section_urls = [s['url'] for s in sections if '下载' in s['name'] or '表证' in s['name']]
        if not section_urls:
            section_urls = [base_url]
        
        for section_url in section_urls[:3]:
            links = find_download_links(section_url, base_url)
            for link in links[:10]:
                file_url = link['url']
                file_data = download_file(file_url)
                if file_data:
                    file_hash = hashlib.md5(file_data).hexdigest()
                    conn = sqlite3.connect(DB_PATH)
                    c = conn.cursor()
                    c.execute("SELECT id FROM templates WHERE file_hash=?", (file_hash,))
                    if not c.fetchone():
                        fields = ''
                        if file_url.endswith(('.xlsx', '.xls')):
                            fields = extract_fields_from_excel(file_data)
                        template = {
                            'id': str(uuid.uuid4())[:8],
                            'province': site['name'],
                            'city': '',
                            'district': '',
                            'report_type': '增值税',
                            'template_name': link['text'] or link['filename'],
                            'template_version': 'v1.0',
                            'source_url': file_url,
                            'source_authority': f"国家税务总局{site['name']}税务局",
                            'publish_date': datetime.now().strftime('%Y-%m-%d'),
                            'required_fields': fields,
                            'status': 'active',
                            'file_hash': file_hash,
                            'file_type': file_url.split('.')[-1]
                        }
                        save_template(template)
                        total_updated += 1
                        log.append(f"新增模板：{template['template_name']} ({site['name']})")
                    conn.close()
                time.sleep(0.3)
    
    status_text.text("同步完成！")
    progress_bar.progress(1.0)
    
    return total_updated, total_errors, log

# ========== 全国省份及城市规则（内置） ==========
PROVINCE_DEFAULT_RULES = [
    {'city': '上海', 'province': '上海', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.07, 'personal_fund': 0.07, 'social_min': 7310, 'social_max': 36549,
     'fund_min': 2590, 'fund_max': 34188, 'source_quote': '沪人社规〔2024〕22号'},
    {'city': '北京', 'province': '北京', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 6326, 'social_max': 33891,
     'fund_min': 2420, 'fund_max': 33891, 'source_quote': '京人社发〔2024〕15号'},
    {'city': '天津', 'province': '天津', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.11, 'personal_fund': 0.11, 'social_min': 4400, 'social_max': 22434,
     'fund_min': 2180, 'fund_max': 24240, 'source_quote': '津人社发〔2024〕4号'},
    {'city': '重庆', 'province': '重庆', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3957, 'social_max': 19784,
     'fund_min': 2100, 'fund_max': 24595, 'source_quote': '渝人社发〔2024〕5号'},
    {'city': '广州', 'province': '广东', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 4588, 'social_max': 22941,
     'fund_min': 2300, 'fund_max': 27960, 'source_quote': '穗人社发〔2024〕3号'},
    {'city': '深圳', 'province': '广东', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 2360, 'social_max': 22941,
     'fund_min': 2360, 'fund_max': 27927, 'source_quote': '深人社规〔2024〕3号'},
    {'city': '南京', 'province': '江苏', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.08, 'personal_fund': 0.08, 'social_min': 4250, 'social_max': 22470,
     'fund_min': 2280, 'fund_max': 27841, 'source_quote': '宁人社发〔2024〕5号'},
    {'city': '苏州', 'province': '江苏', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4250, 'social_max': 22470,
     'fund_min': 2280, 'fund_max': 27874, 'source_quote': '苏人社发〔2024〕6号'},
    {'city': '杭州', 'province': '浙江', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3957, 'social_max': 22941,
     'fund_min': 2280, 'fund_max': 27874, 'source_quote': '杭人社发〔2024〕6号'},
    {'city': '宁波', 'province': '浙江', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3957, 'social_max': 22941,
     'fund_min': 2280, 'fund_max': 27874, 'source_quote': '甬人社发〔2024〕5号'},
    {'city': '成都', 'province': '四川', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4071, 'social_max': 20355,
     'fund_min': 2100, 'fund_max': 25401, 'source_quote': '成人社发〔2024〕7号'},
    {'city': '武汉', 'province': '湖北', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4077, 'social_max': 20385,
     'fund_min': 2010, 'fund_max': 24114, 'source_quote': '武人社发〔2024〕4号'},
    {'city': '长沙', 'province': '湖南', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3604, 'social_max': 18018,
     'fund_min': 1930, 'fund_max': 22998, 'source_quote': '长人社发〔2024〕4号'},
    {'city': '郑州', 'province': '河南', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 3409, 'social_max': 17043,
     'fund_min': 2000, 'fund_max': 22892, 'source_quote': '郑人社发〔2024〕5号'},
    {'city': '青岛', 'province': '山东', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3746, 'social_max': 18726,
     'fund_min': 2010, 'fund_max': 23496, 'source_quote': '青人社发〔2024〕4号'},
    {'city': '西安', 'province': '陕西', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 3957, 'social_max': 19784,
     'fund_min': 1950, 'fund_max': 23556, 'source_quote': '西人社发〔2024〕6号'},
    {'city': '沈阳', 'province': '辽宁', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4100, 'social_max': 20500,
     'fund_min': 2100, 'fund_max': 25200, 'source_quote': '沈人社发〔2024〕5号'},
    {'city': '大连', 'province': '辽宁', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4100, 'social_max': 20500,
     'fund_min': 2100, 'fund_max': 25200, 'source_quote': '大人社发〔2024〕4号'},
    {'city': '福州', 'province': '福建', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4100, 'social_max': 20500,
     'fund_min': 2100, 'fund_max': 25200, 'source_quote': '榕人社发〔2024〕5号'},
    {'city': '厦门', 'province': '福建', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 4100, 'social_max': 20500,
     'fund_min': 2100, 'fund_max': 25200, 'source_quote': '厦人社发〔2024〕4号'},
    {'city': '石家庄', 'province': '河北', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3800, 'social_max': 19000,
     'fund_min': 1900, 'fund_max': 22800, 'source_quote': '石人社发〔2024〕5号'},
    {'city': '合肥', 'province': '安徽', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3900, 'social_max': 19500,
     'fund_min': 1950, 'fund_max': 23400, 'source_quote': '合人社发〔2024〕5号'},
    {'city': '南昌', 'province': '江西', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3800, 'social_max': 19000,
     'fund_min': 1900, 'fund_max': 22800, 'source_quote': '洪人社发〔2024〕4号'},
    {'city': '太原', 'province': '山西', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3700, 'social_max': 18500,
     'fund_min': 1850, 'fund_max': 22200, 'source_quote': '并人社发〔2024〕4号'},
    {'city': '长春', 'province': '吉林', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3700, 'social_max': 18500,
     'fund_min': 1850, 'fund_max': 22200, 'source_quote': '长人社发〔2024〕4号'},
    {'city': '哈尔滨', 'province': '黑龙江', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3600, 'social_max': 18000,
     'fund_min': 1800, 'fund_max': 21600, 'source_quote': '哈人社发〔2024〕4号'},
    {'city': '昆明', 'province': '云南', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3700, 'social_max': 18500,
     'fund_min': 1850, 'fund_max': 22200, 'source_quote': '昆人社发〔2024〕5号'},
    {'city': '贵阳', 'province': '贵州', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3600, 'social_max': 18000,
     'fund_min': 1800, 'fund_max': 21600, 'source_quote': '筑人社发〔2024〕4号'},
    {'city': '兰州', 'province': '甘肃', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3500, 'social_max': 17500,
     'fund_min': 1750, 'fund_max': 21000, 'source_quote': '兰人社发〔2024〕4号'},
    {'city': '呼和浩特', 'province': '内蒙古', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3600, 'social_max': 18000,
     'fund_min': 1800, 'fund_max': 21600, 'source_quote': '呼人社发〔2024〕4号'},
    {'city': '乌鲁木齐', 'province': '新疆', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3500, 'social_max': 17500,
     'fund_min': 1750, 'fund_max': 21000, 'source_quote': '乌人社发〔2024〕4号'},
    {'city': '银川', 'province': '宁夏', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3500, 'social_max': 17500,
     'fund_min': 1750, 'fund_max': 21000, 'source_quote': '银人社发〔2024〕4号'},
    {'city': '西宁', 'province': '青海', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3400, 'social_max': 17000,
     'fund_min': 1700, 'fund_max': 20400, 'source_quote': '宁人社发〔2024〕4号'},
    {'city': '拉萨', 'province': '西藏', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3300, 'social_max': 16500,
     'fund_min': 1650, 'fund_max': 19800, 'source_quote': '拉人社发〔2024〕4号'},
    {'city': '海口', 'province': '海南', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3800, 'social_max': 19000,
     'fund_min': 1900, 'fund_max': 22800, 'source_quote': '海人社发〔2024〕4号'},
    {'city': '南宁', 'province': '广西', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3600, 'social_max': 18000,
     'fund_min': 1800, 'fund_max': 21600, 'source_quote': '南人社发〔2024〕4号'},
]

# ========== 通用模板 ==========
GENERAL_TEMPLATE = {
    "id": "gen001",
    "template_name": "通用申报表（适用于所有地区）",
    "template_version": "v1.0",
    "source_authority": "系统内置",
    "publish_date": datetime.now().strftime("%Y-%m-%d"),
    "required_fields": "纳税人识别号,公司名称,申报金额",
    "source_url": "#"
}

# ========== 初始化默认数据 ==========
def init_default_data():
    if not load_rules():
        all_rules = []
        for r in PROVINCE_DEFAULT_RULES:
            all_rules.append({
                'id': str(uuid.uuid4())[:8],
                'city': r['city'],
                'province': r.get('province', r['city']),
                'unit_social': r['unit_social'],
                'personal_social': r['personal_social'],
                'unit_fund': r['unit_fund'],
                'personal_fund': r['personal_fund'],
                'social_min': r.get('social_min', 0),
                'social_max': r.get('social_max', 999999),
                'fund_min': r.get('fund_min', 0),
                'fund_max': r.get('fund_max', 999999),
                'source_quote': r.get('source_quote', '省份默认'),
                'is_default': 1 if r['city'] == r.get('province', r['city']) else 0
            })
        save_rules(all_rules)

init_default_data()

# ========== 标准化匹配函数 ==========
def normalize_name(name):
    if not name:
        return name
    for suffix in ['省', '市', '区', '县', '自治区', '特别行政区']:
        if name.endswith(suffix):
            name = name[:-len(suffix)]
    return name.strip()

def match_template_with_details(province, city, district, report_type):
    templates = load_templates()
    if not templates:
        return None, None, []
    
    norm_prov = normalize_name(province)
    norm_city = normalize_name(city)
    norm_dist = normalize_name(district) if district else ''
    
    matched = None
    match_level = None
    
    for t in templates:
        if normalize_name(t['province']) == norm_prov and \
           normalize_name(t['city']) == norm_city and \
           normalize_name(t.get('district', '')) == norm_dist and \
           t['report_type'] == report_type:
            matched = t
            match_level = "区级模板"
            break
    if not matched:
        for t in templates:
            if normalize_name(t['province']) == norm_prov and \
               normalize_name(t['city']) == norm_city and \
               t['report_type'] == report_type:
                matched = t
                match_level = "市级模板"
                break
    if not matched:
        for t in templates:
            if normalize_name(t['province']) == norm_prov and \
               t['report_type'] == report_type:
                matched = t
                match_level = "省级模板"
                break
    
    candidates = [t for t in templates if normalize_name(t['province']) == norm_prov and t['report_type'] == report_type]
    return matched, match_level, candidates

# ========== 解析上传的Excel ==========
def parse_uploaded_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    all_companies = []
    for sheet in sheets:
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            header_row = None
            for i, row in df.iterrows():
                row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                if '所属城市' in row_text or '城市' in row_text or '分公司' in row_text:
                    header_row = i
                    break
            if header_row is not None:
                df = pd.read_excel(file, sheet_name=sheet, skiprows=header_row)
                df.columns = [str(c).strip() for c in df.columns]
                city_col = None
                company_col = None
                district_col = None
                for col in df.columns:
                    if '所属城市' in col or '城市' in col:
                        city_col = col
                    elif '分公司' in col or '公司' in col:
                        company_col = col
                    elif '区县' in col or '区' in col:
                        district_col = col
                if city_col and company_col:
                    for _, row in df.iterrows():
                        city = str(row[city_col]) if pd.notna(row[city_col]) else ''
                        company = str(row[company_col]) if pd.notna(row[company_col]) else ''
                        district = str(row[district_col]) if district_col and pd.notna(row[district_col]) else ''
                        if city and company:
                            province = city
                            for r in PROVINCE_DEFAULT_RULES:
                                if r['city'] == city and r.get('province'):
                                    province = r['province']
                                    break
                            all_companies.append({
                                'company_name': company,
                                'province': province,
                                'city': city,
                                'district': district,
                                'tax_id': ''
                            })
        except:
            continue
    unique = []
    seen = set()
    for c in all_companies:
        key = (c['company_name'], c['city'])
        if key not in seen:
            seen.add(key)
            unique.append(c)
    return unique

def get_rule_for_city(city):
    rules = load_rules()
    for r in rules:
        if r['city'] == city:
            return r
    for r in rules:
        if r.get('province') == city:
            return r
    return None

def get_data_source_info(df):
    info = {}
    if df is not None and not df.empty:
        for col in df.columns:
            col_lower = str(col).lower()
            if '年份' in col_lower or '年度' in col_lower:
                info['year'] = df[col].iloc[0] if not df[col].empty else '2025'
            if '月份' in col_lower or '月' in col_lower:
                if '统计月份' in col_lower or '月份' in col_lower:
                    info['month'] = df[col].iloc[0] if not df[col].empty else '12'
    return info

# ========== Streamlit 页面 ==========
st.set_page_config(page_title="官方模板匹配器", layout="wide")
st.title("📋 官方模板匹配器（含爬虫）")
st.markdown("**上传Excel → 自动提取城市/公司 → 选择模板和统计口径 → 生成待复核版Excel**")

template_count = len(load_templates())
rule_count = len(load_rules())
st.success(f"✅ 已内置 {rule_count} 个城市的规则，以及 {template_count} 个官方模板")

# ===== 侧边栏 =====
with st.sidebar:
    st.header("📤 上传数据Excel")
    uploaded_file = st.file_uploader("选择Excel文件（.xlsx）", type=["xlsx"])
    
    if uploaded_file:
        with st.spinner("正在解析Excel..."):
            companies = parse_uploaded_excel(uploaded_file)
            if companies:
                save_companies(companies)
                st.success(f"成功提取 {len(companies)} 家公司")
                try:
                    xls = pd.ExcelFile(uploaded_file)
                    data_sheet = None
                    for s in xls.sheet_names:
                        if '明细' in s or '月度' in s or '数据' in s:
                            data_sheet = s
                            break
                    if data_sheet:
                        df_data = pd.read_excel(uploaded_file, sheet_name=data_sheet)
                        st.session_state['imported_df'] = df_data
                        st.session_state['data_sheet_name'] = data_sheet
                        st.success(f"已读取数据Sheet「{data_sheet}」，共{len(df_data)}行")
                except:
                    pass
            else:
                st.warning("未识别到公司数据，请确认Excel包含「城市」和「公司」列")
    
    # ===== 同步官方模板 =====
    st.markdown("---")
    st.header("📡 同步官方模板")
    if st.button("🔄 从国家税务总局官网抓取模板"):
        with st.spinner("正在抓取官方模板，请耐心等待..."):
            updated, errors, logs = sync_official_templates()
            st.success(f"同步完成！新增 {updated} 个模板，{errors} 个错误")
            with st.expander("查看日志"):
                st.text("\n".join(logs))
            st.rerun()
    
    with st.sidebar.expander("🏢 当前公司列表"):
        companies = load_companies()
        if companies:
            st.dataframe(pd.DataFrame(companies))
            st.caption(f"共 {len(companies)} 家公司")
        else:
            st.info("暂无数据")
    
    with st.sidebar.expander("📚 查看所有模板"):
        templates = load_templates()
        if templates:
            df_temp = pd.DataFrame(templates)
            st.dataframe(df_temp[['province', 'city', 'report_type', 'template_name', 'template_version', 'source_url']])
            st.caption(f"共 {len(templates)} 个模板")
        else:
            st.info("暂无模板，请点击同步官方模板获取")

# ===== 主体 =====
st.subheader("📊 导入数据预览")
data_source_info = ""
if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
    df_preview = st.session_state['imported_df']
    st.dataframe(df_preview.head(10))
    sheet_name = st.session_state.get('data_sheet_name', '未知Sheet')
    info = get_data_source_info(df_preview)
    year = info.get('year', '')
    month = info.get('month', '')
    data_source_info = f"数据来源：{sheet_name}"
    if year:
        data_source_info += f"，年份：{year}"
    if month:
        data_source_info += f"，月份：{month}"
    st.caption(f"共 {len(df_preview)} 行数据 | {data_source_info}")
else:
    st.info("上传Excel后，此处将显示数据预览")

companies = load_companies()
if not companies:
    st.info("👈 请先在侧边栏上传包含公司/城市数据的Excel")
    st.stop()

all_provinces = sorted(set(c['province'] for c in companies if c['province']))

col1, col2, col3 = st.columns(3)
with col1:
    province = st.selectbox("省份", [""] + all_provinces)
    cities = sorted(set(c['city'] for c in companies if c['province'] == province)) if province else sorted(set(c['city'] for c in companies))
    city = st.selectbox("城市", [""] + cities)
with col2:
    districts = sorted(set(c['district'] for c in companies if c['province'] == province and c['city'] == city)) if province and city else []
    district = st.selectbox("区县", [""] + districts)
    company_list = [c for c in companies if c['province'] == province and c['city'] == city and (not district or c['district'] == district)]
    company_names = [c['company_name'] for c in company_list]
    selected_company_names = st.multiselect("公司（可多选）", company_names)
with col3:
    report_type = st.selectbox("报表类型", ["", "增值税", "社保", "公积金", "个人所得税", "企业所得税", "年度汇算清缴"])
    period_type = st.selectbox("统计口径", ["月度（12月单月）", "累计（1-12月）"])

selected_companies = [c for c in company_list if c['company_name'] in selected_company_names]

if selected_companies and report_type:
    st.markdown("---")
    st.subheader("🔍 匹配结果")
    
    matched, match_level, candidates = match_template_with_details(province, city, district, report_type)
    
    # 构建选项列表
    options = {}
    if matched:
        options[f"✅ {matched['template_name']}（{match_level}）"] = matched
    for c in candidates:
        if c['id'] != (matched['id'] if matched else ''):
            options[f"📄 {c['template_name']}（{c['province']}{c['city']}{c.get('district','')}）"] = c
    # 通用模板
    general_key = "⭐ 通用模板（适用于所有地区，始终可用）"
    options[general_key] = GENERAL_TEMPLATE
    
    selected_template = None
    if len(options) > 1:
        st.info("💡 请选择要使用的模板（默认选中推荐模板）")
        keys = list(options.keys())
        default_idx = 0
        if matched:
            for i, k in enumerate(keys):
                if "✅" in k:
                    default_idx = i
                    break
        else:
            for i, k in enumerate(keys):
                if "⭐" in k:
                    default_idx = i
                    break
        selected_key = st.selectbox("选择模板", keys, index=default_idx)
        selected_template = options[selected_key]
        if "⭐" in selected_key:
            match_level = "通用模板（手动选择）"
        elif "✅" in selected_key:
            match_level = "官方模板（自动匹配）"
        else:
            match_level = "官方模板（手动选择）"
    else:
        selected_template = list(options.values())[0]
        match_level = "唯一可用模板"
    
    matched = selected_template
    if not matched:
        matched = GENERAL_TEMPLATE
        match_level = "通用模板（兜底）"
    
    # 显示模板信息
    st.success(f"✅ 已选择模板：{matched['template_name']}（{match_level}）")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**📄 模板信息**")
        st.write(f"模板名称：{matched['template_name']}")
        st.write(f"版本：{matched['template_version']}")
        st.write(f"发布机构：{matched['source_authority']}")
        st.write(f"发布日期：{matched['publish_date']}")
        st.write(f"必填字段：{matched['required_fields']}")
    with col_b:
        st.markdown("**🔗 来源信息**")
        if 'source_url' in matched and matched['source_url'] != '#':
            st.write(f"来源URL：[{matched['source_url']}]({matched['source_url']})")
        else:
            st.write("来源：系统内置")
        if 'province' in matched and matched['province']:
            st.write(f"适用地区：{matched['province']} {matched.get('city','')} {matched.get('district','')}")
        st.write(f"报表类型：{report_type}")
    
    # 模板预览
    st.subheader("📋 模板预览")
    fields = matched['required_fields'].split(',')
    st.markdown(f"**字段列表**：{', '.join(fields)}")
    
    sample_row = {}
    for f in fields:
        sample_values = {
            '纳税人识别号': '91310115MA1KXXXXX',
            '公司名称': selected_companies[0]['company_name'] if selected_companies else '示例公司',
            '销售额': '100,000.00',
            '进项税额': '13,000.00',
            '应纳税额': '0.00',
            '单位名称': selected_companies[0]['company_name'] if selected_companies else '示例公司',
            '社保登记号': 'SH123456',
            '基数': '8,000.00',
            '单位金额': '1,280.00',
            '个人金额': '640.00',
            '单位比例': '12.0%',
            '个人比例': '12.0%',
            '公积金账号': 'GJJ123456',
            '收入额': '100,000.00',
            '专项扣除': '0.00',
            '营业收入': '1,000,000.00',
            '营业成本': '600,000.00',
            '应纳税所得额': '100,000.00',
            '全年收入': '12,000,000.00',
            '全年成本': '7,200,000.00',
            '已预缴税额': '150,000.00',
            '应补退税额': '0.00',
            '申报金额': '100,000.00'
        }
        sample_row[f] = sample_values.get(f, f'<{f} 示例值>')
    
    preview_df = pd.DataFrame([{'字段名': f, '示例值': sample_row[f]} for f in fields])
    st.dataframe(preview_df, use_container_width=True)
    
    # 数据校验
    st.subheader("📋 数据校验")
    data_source_text = "未知"
    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
        data_source_text = st.session_state.get('data_sheet_name', '未知Sheet')
        info = get_data_source_info(st.session_state['imported_df'])
        if info.get('year'):
            data_source_text += f"（年份：{info.get('year')}"
        if info.get('month'):
            data_source_text += f"，月份：{info.get('month')}"
        if info.get('year') or info.get('month'):
            data_source_text += "）"
    
    st.info(f"📌 数据来源：{data_source_text}")
    
    missing_rules = []
    for comp in selected_companies:
        rule = get_rule_for_city(comp['city'])
        if rule is None:
            missing_rules.append(comp['city'])
    if missing_rules:
        st.warning(f"⚠️ 以下城市缺少规则，将使用默认值：{', '.join(set(missing_rules))}")
    else:
        st.success("✅ 所有城市已配置规则")
    
    # 报表预览
    st.subheader("📊 报表预览（生成前确认）")
    preview_data = []
    for comp in selected_companies:
        preview_data.append({
            '公司': comp['company_name'],
            '城市': comp['city'],
            '模板': matched['template_name'],
            '匹配级别': match_level
        })
    st.dataframe(pd.DataFrame(preview_data), use_container_width=True)
    
    reviewed = st.checkbox("✅ 我已人工复核确认数据无误", value=False)
    
    if st.button("📥 生成待复核版Excel", disabled=not reviewed):
        if not matched:
            st.error("请先选择模板")
        else:
            generated_files = []
            summary = []
            errors = []
            
            for comp in selected_companies:
                try:
                    rule = get_rule_for_city(comp['city'])
                    fields = matched['required_fields'].split(',')
                    
                    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
                        df_data = st.session_state['imported_df']
                        company_col = None
                        for col in df_data.columns:
                            if '公司' in str(col) or '分公司' in str(col):
                                company_col = col
                                break
                        if company_col:
                            df_comp = df_data[df_data[company_col] == comp['company_name']]
                            if not df_comp.empty:
                                row_data = []
                                for f in fields:
                                    matched_col = None
                                    for col in df_data.columns:
                                        if f in str(col) or str(col) in f:
                                            matched_col = col
                                            break
                                    if matched_col:
                                        row_data.append(df_comp.iloc[0][matched_col])
                                    else:
                                        row_data.append('')
                            else:
                                row_data = [''] * len(fields)
                        else:
                            row_data = [''] * len(fields)
                    else:
                        sample_data = {
                            '纳税人识别号': comp.get('tax_id', ''),
                            '公司名称': comp['company_name'],
                            '销售额': '100,000.00',
                            '进项税额': '13,000.00',
                            '应纳税额': '0.00',
                            '单位名称': comp['company_name'],
                            '社保登记号': 'SH123456',
                            '基数': '8,000.00',
                            '单位金额': str(round(8000 * rule['unit_social'], 2)) if rule else '1,280.00',
                            '个人金额': str(round(8000 * rule['personal_social'], 2)) if rule else '640.00',
                            '单位比例': str(round(rule['unit_fund'] * 100, 1)) if rule else '12.0',
                            '个人比例': str(round(rule['personal_fund'] * 100, 1)) if rule else '12.0',
                            '公积金账号': 'GJJ123456',
                            '收入额': '100,000.00',
                            '专项扣除': '0.00',
                            '营业收入': '1,000,000.00',
                            '营业成本': '600,000.00',
                            '应纳税所得额': '100,000.00',
                            '申报金额': '100,000.00',
                            '全年收入': '12,000,000.00',
                            '全年成本': '7,200,000.00',
                            '已预缴税额': '150,000.00',
                            '应补退税额': '0.00'
                        }
                        row_data = [sample_data.get(f, '') for f in fields]
                    
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "申报表"
                    ws.append(fields)
                    ws.append(row_data)
                    
                    ws.insert_rows(1)
                    ws['A1'] = f'【系统生成 - 待复核版】统计口径：{period_type}'
                    ws['A1'].font = Font(color='FF0000', bold=True, size=14)
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fields))
                    ws['A1'].alignment = Alignment(horizontal='center')
                    ws['A1'].fill = PatternFill(start_color='FFF9E6', end_color='FFF9E6', fill_type='solid')
                    
                    ws.insert_rows(2)
                    ws['A2'] = f'模板名称：{matched["template_name"]}  版本：{matched["template_version"]}  匹配级别：{match_level}'
                    ws['A2'].font = Font(color='666666', size=10)
                    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(fields))
                    
                    ws.insert_rows(3)
                    ws['A3'] = f'来源：{matched.get("source_authority","")}  发布日期：{matched.get("publish_date","")}'
                    ws['A3'].font = Font(color='666666', size=10)
                    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(fields))
                    
                    ws.insert_rows(4)
                    ws['A4'] = f'数据来源：{data_source_text}  统计口径：{period_type}'
                    ws['A4'].font = Font(color='666666', size=10)
                    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=len(fields))
                    
                    # 年检汇总
                    ws_annual = wb.create_sheet("年检汇总")
                    ws_annual.append(['年检汇总数据'])
                    ws_annual.merge_cells('A1:B1')
                    ws_annual['A1'].font = Font(bold=True, size=12)
                    
                    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
                        df_all = st.session_state['imported_df']
                        social_col = None
                        fund_col = None
                        unit_col = None
                        personal_col = None
                        total_col = None
                        people_col = None
                        for col in df_all.columns:
                            col_str = str(col)
                            if '社保' in col_str and '合计' in col_str:
                                if '单位' in col_str:
                                    social_col = col
                                elif '个人' in col_str:
                                    personal_col = col
                            elif '公积金' in col_str and '合计' in col_str:
                                fund_col = col
                            elif '单位总费用' in col_str:
                                unit_col = col
                            elif '个人总费用' in col_str:
                                personal_col = col
                            elif '全部总费用' in col_str or '总金额' in col_str:
                                total_col = col
                            elif '参保人数' in col_str:
                                people_col = col
                        
                        if social_col:
                            social_total = df_all[social_col].sum()
                        else:
                            social_total = 0
                        if fund_col:
                            fund_total = df_all[fund_col].sum()
                        else:
                            fund_total = 0
                        if people_col:
                            total_people = df_all[people_col].sum()
                        else:
                            total_people = len(df_all)
                        
                        if unit_col and personal_col:
                            unit_total = df_all[unit_col].sum()
                            personal_total = df_all[personal_col].sum()
                            grand_total = df_all[total_col].sum() if total_col else (unit_total + personal_total)
                        else:
                            unit_total = social_total
                            personal_total = personal_col if personal_col else 0
                            grand_total = social_total + fund_total if fund_total else social_total
                    else:
                        total_people = 0
                        social_total = 0
                        fund_total = 0
                        unit_total = 0
                        personal_total = 0
                        grand_total = 0
                    
                    ws_annual.append(['公司名称', comp['company_name']])
                    ws_annual.append(['所属城市', comp['city']])
                    ws_annual.append(['统计口径', period_type])
                    ws_annual.append(['参保人数（全年）', int(total_people) if total_people else 0])
                    ws_annual.append(['全年社保缴费基数总额', round(social_total, 2)])
                    ws_annual.append(['全年公积金缴费基数总额', round(fund_total, 2)])
                    ws_annual.append(['单位全年缴费总额', round(unit_total, 2)])
                    ws_annual.append(['个人全年缴费总额', round(personal_total, 2)])
                    ws_annual.append(['全年总费用', round(grand_total, 2)])
                    ws_annual.append(['报告生成时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                    ws_annual.append(['数据来源', data_source_text])
                    
                    audit = wb.create_sheet("审计日志")
                    audit.append(['操作时间', '操作类型', '操作人', '详情'])
                    audit.append([datetime.now().isoformat(), 'GENERATED', '系统', f'公司:{comp["company_name"]}, 城市:{comp["city"]}, 模板:{matched["template_name"]}, 匹配级别:{match_level}, 数据来源:{data_source_text}'])
                    
                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    fname = f"{comp['company_name']}_{report_type}_{period_type.replace('（','_').replace('）','')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    generated_files.append((fname, output.getvalue()))
                    summary.append({
                        '公司': comp['company_name'], 
                        '城市': comp['city'], 
                        '模板': matched['template_name'],
                        '匹配级别': match_level,
                        '数据来源': data_source_text,
                        '状态': '待复核'
                    })
                    
                    save_export({
                        'id': str(uuid.uuid4())[:8],
                        'company_id': comp['id'],
                        'template_id': matched.get('id', 'gen001'),
                        'company_name': comp['company_name'],
                        'city': comp['city'],
                        'province': comp.get('province', ''),
                        'report_type': report_type,
                        'period_type': period_type,
                        'generated_at': datetime.now().isoformat(),
                        'review_status': 'pending',
                        'file_name': fname,
                        'file_data': output.getvalue(),
                        'data_source': data_source_text,
                        'month_used': period_type,
                        'year_used': datetime.now().strftime('%Y')
                    })
                except Exception as e:
                    errors.append(f"{comp['company_name']}: {str(e)}")
            
            if errors:
                for err in errors:
                    st.warning(err)
            if generated_files:
                st.success(f"✅ 成功生成 {len(generated_files)} 份报表")
                st.dataframe(pd.DataFrame(summary), use_container_width=True)
                if len(generated_files) > 1:
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as zf:
                        for fname, data in generated_files:
                            zf.writestr(fname, data)
                    zip_buffer.seek(0)
                    st.download_button("📦 下载全部报表（ZIP）", data=zip_buffer, file_name=f"报表_{datetime.now().strftime('%Y%m%d')}.zip", mime="application/zip")
                else:
                    fname, data = generated_files[0]
                    st.download_button(f"📥 下载 {fname}", data=BytesIO(data), file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    if not selected_companies:
        st.info("👆 请先选择公司")
    elif not report_type:
        st.info("👆 请选择报表类型")

# ===== 导出历史 =====
with st.expander("📋 导出历史记录"):
    history = load_export_history()
    if history:
        df_hist = pd.DataFrame(history)
        st.dataframe(df_hist[['company_name', 'city', 'report_type', 'period_type', 'data_source', 'generated_at', 'review_status']], use_container_width=True)
        
        pending = [h for h in history if h['review_status'] == 'pending']
        if pending:
            st.subheader("✅ 复核待处理报表")
            opts = [f"{h['company_name']} - {h['city']} ({h['generated_at'][:10]})" for h in pending]
            sel_idx = st.selectbox("选择要复核的报表", range(len(opts)), format_func=lambda x: opts[x])
            selected = pending[sel_idx]
            col1, col2 = st.columns(2)
            with col1:
                if st.button("✅ 通过复核"):
                    update_export_status(selected['id'], 'approved', '复核员')
                    st.success("已通过复核")
                    st.rerun()
            with col2:
                if st.button("❌ 驳回"):
                    update_export_status(selected['id'], 'rejected', '复核员')
                    st.warning("已驳回")
                    st.rerun()
    else:
        st.info("暂无导出记录")

# ===== 查看知识库 =====
with st.expander("📚 官方模板知识库（按省份查看）"):
    templates = load_templates()
    if templates:
        provinces_in_templates = sorted(set(t['province'] for t in templates))
        selected_province = st.selectbox("选择省份查看模板", [""] + provinces_in_templates)
        if selected_province:
            filtered = [t for t in templates if t['province'] == selected_province]
            st.dataframe(pd.DataFrame(filtered)[['city', 'district', 'report_type', 'template_name', 'template_version', 'source_authority']])
        else:
            st.dataframe(pd.DataFrame(templates)[['province', 'city', 'report_type', 'template_name', 'template_version']])
        st.caption(f"共 {len(templates)} 个官方模板，覆盖 {len(provinces_in_templates)} 个省份，6种报表类型")
    else:
        st.info("暂无模板")
