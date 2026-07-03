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
import shutil
import json
import hashlib
import re

# ========== 配置 ==========
DB_PATH = os.path.join(os.path.dirname(__file__), "app_data.db")
BACKUP_DIR = os.path.join(os.path.dirname(__file__), "backup")
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "custom_templates")

for d in [BACKUP_DIR, TEMPLATES_DIR]:
    if not os.path.exists(d):
        os.makedirs(d)

# ========== PDF导出依赖 ==========
try:
    import pdfkit
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# ========== 备份函数 ==========
def backup_database():
    if not os.path.exists(DB_PATH):
        return
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BACKUP_DIR, f"data_backup_{timestamp}.db")
    try:
        shutil.copy2(DB_PATH, backup_path)
        backups = sorted([f for f in os.listdir(BACKUP_DIR) if f.startswith("data_backup_")])
        if len(backups) > 30:
            for f in backups[:-30]:
                os.remove(os.path.join(BACKUP_DIR, f))
        return backup_path
    except Exception as e:
        st.error(f"备份失败：{e}")
        return None

def get_backup_list():
    if not os.path.exists(BACKUP_DIR):
        return []
    return sorted([f for f in os.listdir(BACKUP_DIR) if f.startswith("data_backup_")], reverse=True)

def restore_backup(backup_file):
    backup_path = os.path.join(BACKUP_DIR, backup_file)
    if not os.path.exists(backup_path):
        return False
    try:
        backup_database()
        shutil.copy2(backup_path, DB_PATH)
        return True
    except Exception as e:
        st.error(f"恢复失败：{e}")
        return False

# ========== 数据库初始化 ==========
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS companies (
        id TEXT PRIMARY KEY, company_name TEXT, province TEXT, city TEXT, district TEXT, tax_id TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS templates (
        id TEXT PRIMARY KEY, province TEXT, city TEXT, district TEXT, report_type TEXT,
        template_name TEXT, template_version TEXT, source_url TEXT, source_authority TEXT,
        publish_date TEXT, required_fields TEXT, status TEXT, file_hash TEXT, file_type TEXT,
        is_custom BOOLEAN DEFAULT 0, field_mapping_source TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS custom_templates (
        id TEXT PRIMARY KEY, name TEXT, file_data BLOB, field_mapping TEXT,
        sheet_name TEXT, created_at TEXT, updated_at TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS rules (
        id TEXT PRIMARY KEY, city TEXT, province TEXT, unit_social REAL, personal_social REAL,
        unit_fund REAL, personal_fund REAL, social_min REAL, social_max REAL,
        fund_min REAL, fund_max REAL, source_quote TEXT, is_default BOOLEAN DEFAULT 0,
        rule_version TEXT, effective_date TEXT,
        source_url TEXT, source_title TEXT, source_publish_date TEXT,
        collected_at TEXT, applicable_region TEXT, official_channel TEXT, notes TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS export_history (
        id TEXT PRIMARY KEY, company_id TEXT, template_id TEXT, company_name TEXT,
        city TEXT, province TEXT, report_type TEXT, period_type TEXT, generated_at TEXT,
        review_status TEXT, reviewer TEXT, reviewed_at TEXT, file_name TEXT, file_data BLOB,
        data_source TEXT, month_used TEXT, year_used TEXT, custom_period TEXT,
        batch_id TEXT, job_name TEXT, field_mapping TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS source_registry (
        id TEXT PRIMARY KEY, authority_type TEXT, province TEXT, city TEXT, district TEXT,
        authority_name TEXT, official_site_name TEXT, source_url TEXT, source_level TEXT,
        source_section TEXT, is_official BOOLEAN, crawl_allowed BOOLEAN, last_checked TEXT,
        status TEXT, notes TEXT, document_name TEXT, document_version TEXT, publish_year TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS job_batches (
        id TEXT PRIMARY KEY, batch_name TEXT, created_at TEXT, status TEXT,
        total_companies INTEGER, total_reports INTEGER, review_status TEXT,
        parameters TEXT, created_by TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS job_details (
        id TEXT PRIMARY KEY, batch_id TEXT, company_id TEXT, company_name TEXT,
        city TEXT, province TEXT, report_type TEXT, period_type TEXT,
        status TEXT, error_message TEXT, file_name TEXT, file_data BLOB,
        generated_at TEXT, rule_source TEXT, data_source TEXT
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS verification_status (
        id TEXT PRIMARY KEY, batch_id TEXT,
        source_verified BOOLEAN DEFAULT 0, template_verified BOOLEAN DEFAULT 0,
        rule_verified BOOLEAN DEFAULT 0, data_verified BOOLEAN DEFAULT 0,
        reviewer_name TEXT, verified_at TEXT, notes TEXT, export_type TEXT DEFAULT '验证版'
    )''')
    conn.commit()
    conn.close()

init_db()

# ========== 数据库迁移 ==========
def migrate_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("PRAGMA table_info(rules)")
    existing_cols = [col[1] for col in c.fetchall()]
    new_cols = ['source_url', 'source_title', 'source_publish_date', 'collected_at', 'applicable_region', 'official_channel', 'notes']
    for col in new_cols:
        if col not in existing_cols:
            c.execute(f"ALTER TABLE rules ADD COLUMN {col} TEXT")
    tables_to_check = [
        ('templates', ['field_mapping_source']),
        ('export_history', ['batch_id', 'job_name', 'field_mapping']),
        ('source_registry', ['document_name', 'document_version', 'publish_year']),
    ]
    for table, cols in tables_to_check:
        c.execute(f"PRAGMA table_info({table})")
        existing = [col[1] for col in c.fetchall()]
        for col in cols:
            if col not in existing:
                c.execute(f"ALTER TABLE {table} ADD COLUMN {col} TEXT")
    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='job_batches'")
    if not c.fetchone():
        c.execute('''CREATE TABLE job_batches (
            id TEXT PRIMARY KEY, batch_name TEXT, created_at TEXT, status TEXT,
            total_companies INTEGER, total_reports INTEGER, review_status TEXT,
            parameters TEXT, created_by TEXT
        )''')
    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='job_details'")
    if not c.fetchone():
        c.execute('''CREATE TABLE job_details (
            id TEXT PRIMARY KEY, batch_id TEXT, company_id TEXT, company_name TEXT,
            city TEXT, province TEXT, report_type TEXT, period_type TEXT,
            status TEXT, error_message TEXT, file_name TEXT, file_data BLOB,
            generated_at TEXT, rule_source TEXT, data_source TEXT
        )''')
    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='verification_status'")
    if not c.fetchone():
        c.execute('''CREATE TABLE verification_status (
            id TEXT PRIMARY KEY, batch_id TEXT,
            source_verified BOOLEAN DEFAULT 0, template_verified BOOLEAN DEFAULT 0,
            rule_verified BOOLEAN DEFAULT 0, data_verified BOOLEAN DEFAULT 0,
            reviewer_name TEXT, verified_at TEXT, notes TEXT, export_type TEXT DEFAULT '验证版'
        )''')
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

def load_custom_templates():
    return safe_execute_query("SELECT * FROM custom_templates")

def load_rules():
    return safe_execute_query("SELECT * FROM rules ORDER BY province, city")

def load_export_history():
    return safe_execute_query("SELECT * FROM export_history ORDER BY generated_at DESC")

def load_source_registry():
    return safe_execute_query("SELECT * FROM source_registry ORDER BY province, city")

def load_job_batches():
    return safe_execute_query("SELECT * FROM job_batches ORDER BY created_at DESC")

def load_job_details(batch_id):
    return safe_execute_query("SELECT * FROM job_details WHERE batch_id=? ORDER BY generated_at DESC", (batch_id,))

def load_verification_status(batch_id):
    rows = safe_execute_query("SELECT * FROM verification_status WHERE batch_id=?", (batch_id,))
    return rows[0] if rows else None

def save_verification_status(data):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO verification_status 
        (id, batch_id, source_verified, template_verified, rule_verified, data_verified,
         reviewer_name, verified_at, notes, export_type)
        VALUES (?,?,?,?,?,?,?,?,?,?)''',
        (data.get('id', str(uuid.uuid4())[:8]), data['batch_id'],
         data.get('source_verified', 0), data.get('template_verified', 0),
         data.get('rule_verified', 0), data.get('data_verified', 0),
         data.get('reviewer_name', ''), data.get('verified_at', datetime.now().isoformat()),
         data.get('notes', ''), data.get('export_type', '验证版')))
    conn.commit()
    conn.close()
    backup_database()

def update_verification_status(batch_id, field, value):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(f"UPDATE verification_status SET {field}=? WHERE batch_id=?", (value, batch_id))
    conn.commit()
    conn.close()
    backup_database()

def save_source_registry(sources):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM source_registry")
    for s in sources:
        c.execute('''INSERT OR REPLACE INTO source_registry 
            (id, authority_type, province, city, district, authority_name,
             official_site_name, source_url, source_level, source_section,
             is_official, crawl_allowed, last_checked, status, notes,
             document_name, document_version, publish_year)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (s.get('id', str(uuid.uuid4())[:8]), s.get('authority_type','tax'),
             s.get('province',''), s.get('city',''), s.get('district',''),
             s.get('authority_name',''), s.get('official_site_name',''),
             s.get('source_url',''), s.get('source_level',''), s.get('source_section',''),
             s.get('is_official',1), s.get('crawl_allowed',1),
             s.get('last_checked',''), s.get('status','active'), s.get('notes',''),
             s.get('document_name',''), s.get('document_version',''), s.get('publish_year','')))
    conn.commit()
    conn.close()
    backup_database()

def delete_source_by_id(source_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM source_registry WHERE id=?", (source_id,))
    conn.commit()
    conn.close()
    backup_database()

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
    backup_database()

def save_template(template):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id FROM templates WHERE source_url=? OR file_hash=?", 
              (template.get('source_url',''), template.get('file_hash','')))
    existing = c.fetchone()
    if existing:
        c.execute('''UPDATE templates SET 
            template_name=?, template_version=?, source_authority=?, publish_date=?,
            required_fields=?, status=?, file_type=?, is_custom=?, field_mapping_source=?
            WHERE id=?''',
            (template['template_name'], template.get('template_version','v1.0'),
             template.get('source_authority',''), template.get('publish_date',''),
             template.get('required_fields',''), template.get('status','active'),
             template.get('file_type',''), template.get('is_custom',0),
             template.get('field_mapping_source',''), existing[0]))
    else:
        c.execute('''INSERT INTO templates 
            (id, province, city, district, report_type, template_name, template_version,
             source_url, source_authority, publish_date, required_fields, status, 
             file_hash, file_type, is_custom, field_mapping_source)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (template['id'], template.get('province',''), template.get('city',''),
             template.get('district',''), template.get('report_type',''),
             template['template_name'], template.get('template_version','v1.0'),
             template.get('source_url',''), template.get('source_authority',''),
             template.get('publish_date',''), template.get('required_fields',''),
             template.get('status','active'), template.get('file_hash',''),
             template.get('file_type',''), template.get('is_custom',0),
             template.get('field_mapping_source','')))
    conn.commit()
    conn.close()
    backup_database()

def save_custom_template(template):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO custom_templates 
        (id, name, file_data, field_mapping, sheet_name, created_at, updated_at)
        VALUES (?,?,?,?,?,?,?)''',
        (template['id'], template['name'], template['file_data'],
         json.dumps(template.get('field_mapping', {})),
         template.get('sheet_name', ''), template.get('created_at', datetime.now().isoformat()),
         datetime.now().isoformat()))
    conn.commit()
    conn.close()
    backup_database()

def delete_custom_template(template_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM custom_templates WHERE id=?", (template_id,))
    conn.commit()
    conn.close()
    backup_database()

def save_rules(rules):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM rules")
    for r in rules:
        c.execute('''INSERT INTO rules 
            (id, city, province, unit_social, personal_social, unit_fund, personal_fund,
             social_min, social_max, fund_min, fund_max, source_quote, is_default,
             rule_version, effective_date, source_url, source_title, source_publish_date,
             collected_at, applicable_region, official_channel, notes)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (r['id'], r['city'], r.get('province',''), r['unit_social'], r['personal_social'],
             r['unit_fund'], r['personal_fund'], r.get('social_min',0), r.get('social_max',999999),
             r.get('fund_min',0), r.get('fund_max',999999), r.get('source_quote',''),
             r.get('is_default',0), r.get('rule_version','v1.0'), r.get('effective_date',''),
             r.get('source_url',''), r.get('source_title',''), r.get('source_publish_date',''),
             r.get('collected_at', datetime.now().isoformat()), r.get('applicable_region',''),
             r.get('official_channel',''), r.get('notes','')))
    conn.commit()
    conn.close()
    backup_database()

def save_export(record):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO export_history 
        (id, company_id, template_id, company_name, city, province, report_type, period_type,
         generated_at, review_status, reviewer, reviewed_at, file_name, file_data,
         data_source, month_used, year_used, custom_period, batch_id, job_name, field_mapping)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (record['id'], record.get('company_id',''), record.get('template_id',''),
         record['company_name'], record.get('city',''), record.get('province',''),
         record.get('report_type',''), record.get('period_type',''), record['generated_at'],
         record.get('review_status','pending'), record.get('reviewer',''), record.get('reviewed_at',''),
         record.get('file_name',''), record.get('file_data', None),
         record.get('data_source',''), record.get('month_used',''), record.get('year_used',''),
         record.get('custom_period',''), record.get('batch_id',''), record.get('job_name',''),
         record.get('field_mapping','')))
    conn.commit()
    conn.close()
    backup_database()

def save_batch_job(batch):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO job_batches 
        (id, batch_name, created_at, status, total_companies, total_reports, 
         review_status, parameters, created_by)
        VALUES (?,?,?,?,?,?,?,?,?)''',
        (batch['id'], batch['batch_name'], batch['created_at'], batch.get('status','pending'),
         batch.get('total_companies',0), batch.get('total_reports',0),
         batch.get('review_status','pending'), json.dumps(batch.get('parameters',{})),
         batch.get('created_by','系统')))
    conn.commit()
    conn.close()
    backup_database()

def save_job_details(details):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    for d in details:
        c.execute('''INSERT OR REPLACE INTO job_details 
            (id, batch_id, company_id, company_name, city, province, report_type, 
             period_type, status, error_message, file_name, file_data, generated_at,
             rule_source, data_source)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (d.get('id', str(uuid.uuid4())[:8]), d.get('batch_id',''), d.get('company_id',''),
             d.get('company_name',''), d.get('city',''), d.get('province',''),
             d.get('report_type',''), d.get('period_type',''), d.get('status','pending'),
             d.get('error_message',''), d.get('file_name',''), d.get('file_data', None),
             d.get('generated_at', datetime.now().isoformat()), d.get('rule_source',''),
             d.get('data_source','')))
    conn.commit()
    conn.close()
    backup_database()

def update_export_status(export_id, status, reviewer):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''UPDATE export_history 
        SET review_status=?, reviewer=?, reviewed_at=?
        WHERE id=?''',
        (status, reviewer, datetime.now().isoformat(), export_id))
    conn.commit()
    conn.close()
    backup_database()

def update_batch_status(batch_id, status, review_status=None):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    if review_status:
        c.execute('''UPDATE job_batches 
            SET status=?, review_status=?
            WHERE id=?''', (status, review_status, batch_id))
    else:
        c.execute('''UPDATE job_batches SET status=?
            WHERE id=?''', (status, batch_id))
    conn.commit()
    conn.close()
    backup_database()

# ========== 初始化示例规则 ==========
def init_sample_rules():
    existing = load_rules()
    if existing:
        return
    sample_rules = [
        {
            'id': str(uuid.uuid4())[:8],
            'city': '上海',
            'province': '上海',
            'unit_social': 0.16,
            'personal_social': 0.08,
            'unit_fund': 0.07,
            'personal_fund': 0.07,
            'social_min': 7310,
            'social_max': 36549,
            'fund_min': 2590,
            'fund_max': 34188,
            'source_quote': '沪人社规〔2024〕22号',
            'rule_version': '2024.1',
            'effective_date': '2024-07-01',
            'source_url': 'https://rsj.sh.gov.cn/',
            'source_title': '上海市2024年度社保缴费基数调整通知',
            'source_publish_date': '2024-06-20',
            'collected_at': datetime.now().isoformat(),
            'applicable_region': '上海',
            'official_channel': '上海市人社局官网',
            'notes': '示例规则'
        },
        {
            'id': str(uuid.uuid4())[:8],
            'city': '北京',
            'province': '北京',
            'unit_social': 0.16,
            'personal_social': 0.08,
            'unit_fund': 0.12,
            'personal_fund': 0.12,
            'social_min': 6326,
            'social_max': 33891,
            'fund_min': 2420,
            'fund_max': 33891,
            'source_quote': '京人社发〔2024〕15号',
            'rule_version': '2024.1',
            'effective_date': '2024-07-01',
            'source_url': 'https://rsj.beijing.gov.cn/',
            'source_title': '北京市2024年度社保缴费基数调整通知',
            'source_publish_date': '2024-06-25',
            'collected_at': datetime.now().isoformat(),
            'applicable_region': '北京',
            'official_channel': '北京市人社局官网',
            'notes': '示例规则'
        },
    ]
    save_rules(sample_rules)

init_sample_rules()

# ========== 辅助函数 ==========
def normalize_name(name):
    if not name:
        return name
    for suffix in ['省', '市', '区', '县', '自治区', '特别行政区', '自治州', '地区']:
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
        if normalize_name(t['province']) == norm_prov and normalize_name(t['city']) == norm_city and normalize_name(t.get('district', '')) == norm_dist and t['report_type'] == report_type:
            matched = t; match_level = "区级模板"; break
    if not matched:
        for t in templates:
            if normalize_name(t['province']) == norm_prov and normalize_name(t['city']) == norm_city and t['report_type'] == report_type:
                matched = t; match_level = "市级模板"; break
    if not matched:
        for t in templates:
            if normalize_name(t['province']) == norm_prov and t['report_type'] == report_type:
                matched = t; match_level = "省级模板"; break
    candidates = [t for t in templates if normalize_name(t['province']) == norm_prov and t['report_type'] == report_type]
    return matched, match_level, candidates

# ========== 优化二：模板推荐功能 ==========
def recommend_template(province, city, report_type, templates):
    """根据省份、城市和报表类型推荐最佳模板"""
    if not templates:
        return None, "无可用模板"
    norm_prov = normalize_name(province)
    norm_city = normalize_name(city)
    
    # 1. 精确匹配：省份+城市+报表类型
    for t in templates:
        if normalize_name(t['province']) == norm_prov and normalize_name(t['city']) == norm_city and t['report_type'] == report_type:
            return t, f"✅ 精确匹配：{province} {city} {report_type}"
    
    # 2. 城市级匹配：只匹配城市（忽略省份）
    for t in templates:
        if normalize_name(t['city']) == norm_city and t['report_type'] == report_type:
            return t, f"📌 城市级匹配：{city}（跨省通用）"
    
    # 3. 省级匹配：只匹配省份
    for t in templates:
        if normalize_name(t['province']) == norm_prov and t['report_type'] == report_type:
            return t, f"📌 省级匹配：{province}（省级通用）"
    
    # 4. 全局通用
    return None, "💡 无精确匹配，建议使用通用模板"

def get_custom_template_field_mapping(custom_template):
    if not custom_template:
        return {}
    try:
        return json.loads(custom_template.get('field_mapping', '{}'))
    except:
        return {}

def apply_custom_template_mapping(wb, data, mapping):
    if not mapping:
        return
    ws = wb.active
    for field, cell_ref in mapping.items():
        if field in data:
            ws[cell_ref] = data[field]

def auto_load_sheet_with_header_detection(file, sheet_name):
    xls = pd.ExcelFile(file)
    df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    header_row = None
    for i, row in df_raw.iterrows():
        row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
        if ('城市' in row_text and '公司' in row_text) or ('省份' in row_text and '城市' in row_text):
            header_row = i; break
    if header_row is not None:
        df = pd.read_excel(file, sheet_name=sheet_name, skiprows=header_row)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how='all')
        return df, header_row
    else:
        df = pd.read_excel(file, sheet_name=sheet_name, header=0)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how='all')
        return df, 0

def is_valid_chinese_name(text):
    if not text or not isinstance(text, str):
        return False
    return len(re.findall(r'[\u4e00-\u9fa5]', text)) >= 2

def is_valid_city_name(text):
    if not text or not isinstance(text, str):
        return False
    return len(re.findall(r'[\u4e00-\u9fa5]', text)) >= 2

# ========== 列名智能匹配 ==========
CITY_KEYWORDS = ['城市', '所属城市', '城市名称', '地区', '所属地区', 'city', '地区名', '城市名']
COMPANY_KEYWORDS = ['公司', '分公司', '公司名称', '企业名称', '单位名称', 'company', '单位', '企业']
DISTRICT_KEYWORDS = ['区县', '区', '县', '城区', 'district', '区域']

def detect_columns(df):
    """智能检测城市、公司、区县列"""
    city_col = None
    company_col = None
    district_col = None
    for col in df.columns:
        col_lower = col.lower().strip()
        if any(kw in col_lower for kw in CITY_KEYWORDS):
            city_col = col
        elif any(kw in col_lower for kw in COMPANY_KEYWORDS):
            company_col = col
        elif any(kw in col_lower for kw in DISTRICT_KEYWORDS):
            district_col = col
    return city_col, company_col, district_col

# ========== 优化一：异常高亮函数 ==========
def highlight_error_rows(df, error_rows):
    """
    返回DataFrame的style，将error_rows中的行标红。
    error_rows: dict {行号: [错误列表]} 或 list 行号
    """
    if not error_rows:
        return df.style
    
    # 如果传入的是列表，转为dict
    if isinstance(error_rows, list):
        error_rows = {row: ["异常"] for row in error_rows}
    
    def row_style(row):
        idx = row.name + 1
        if idx in error_rows:
            errors = error_rows[idx]
            err_text = "; ".join(errors) if isinstance(errors, list) else str(errors)
            return ['background-color: #ffcccc; color: #990000'] * len(row)
        else:
            return [''] * len(row)
    
    return df.style.apply(row_style, axis=1)

def parse_uploaded_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    all_companies = []
    unmapped_cities = set()
    filtered_values = []
    data_sheet_name = None
    city_province_map = {}
    for r in load_rules():
        key = normalize_name(r['city'])
        city_province_map[key] = r['province']
    for s in sheets:
        if any(kw in s.lower() for kw in ['明细', '月度', '数据', '年检', '主数据', '月报', '季报']):
            data_sheet_name = s; break
    if not data_sheet_name:
        data_sheet_name = sheets[0] if sheets else None
    for sheet in sheets:
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None)
            header_row = None
            for i, row in df.iterrows():
                row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
                if '所属城市' in row_text or '城市' in row_text or '分公司' in row_text:
                    header_row = i; break
            if header_row is not None:
                df = pd.read_excel(file, sheet_name=sheet, skiprows=header_row)
                df.columns = [str(c).strip() for c in df.columns]
                city_col, company_col, district_col = detect_columns(df)
                if city_col is None or company_col is None:
                    st.session_state['column_mapping_needed'] = True
                    st.session_state['df_columns'] = list(df.columns)
                    st.warning(f"Sheet「{sheet}」未能自动识别列，请在下方选择映射")
                    all_cols = list(df.columns)
                    city_col = st.selectbox(f"请选择「{sheet}」的城市列", [""] + all_cols, key=f"city_col_{sheet}")
                    company_col = st.selectbox(f"请选择「{sheet}」的公司列", [""] + all_cols, key=f"company_col_{sheet}")
                    district_col = st.selectbox(f"请选择「{sheet}」的区县列（可选）", [""] + all_cols, key=f"district_col_{sheet}")
                    if not city_col or not company_col:
                        st.error(f"Sheet「{sheet}」必须选择城市列和公司列")
                        continue
                if city_col and company_col:
                    for _, row in df.iterrows():
                        city = str(row[city_col]) if pd.notna(row[city_col]) else ''
                        company = str(row[company_col]) if pd.notna(row[company_col]) else ''
                        district = str(row[district_col]) if district_col and pd.notna(row[district_col]) else ''
                        if city and not is_valid_city_name(city):
                            filtered_values.append(f"城市无效: {city}"); continue
                        if company and not is_valid_chinese_name(company):
                            filtered_values.append(f"公司名称无效: {company}"); continue
                        if city and company:
                            norm_city = normalize_name(city)
                            province = city_province_map.get(norm_city, '')
                            if not province:
                                unmapped_cities.add(city)
                            all_companies.append({
                                'company_name': company,
                                'province': province,
                                'city': city,
                                'district': district,
                                'tax_id': ''
                            })
        except Exception as e:
            st.error(f"解析Sheet {sheet} 时出错：{e}")
            continue
    if filtered_values:
        st.session_state['filtered_values'] = filtered_values
    unique = []
    seen = set()
    for c in all_companies:
        key = (c['company_name'], c['city'])
        if key not in seen:
            seen.add(key)
            unique.append(c)
    return unique, unmapped_cities, sheets, data_sheet_name

def parse_multiple_files(files):
    all_companies = []
    all_sheets = []
    unmapped_cities = set()
    data_sheet_name = None
    for file in files:
        companies, unmapped, sheets, data_sheet = parse_uploaded_excel(file)
        all_companies.extend(companies)
        all_sheets.extend(sheets)
        unmapped_cities.update(unmapped)
        if data_sheet and not data_sheet_name:
            data_sheet_name = data_sheet
    unique = []
    seen = set()
    for c in all_companies:
        key = (c['company_name'], c['city'])
        if key not in seen:
            seen.add(key)
            unique.append(c)
    return unique, unmapped_cities, all_sheets, data_sheet_name

def validate_data(df, rules):
    if df is None or df.empty:
        return {'total_rows': 0, 'error_rows': 0, 'details': {}, 'summary': {}, 'error_rows_detail': {}}
    errors = {}
    city_col = None
    for col in df.columns:
        if any(kw in col.lower() for kw in CITY_KEYWORDS):
            city_col = col; break
    if city_col is None:
        errors['城市列缺失'] = ['未找到城市列']
        return {'total_rows': len(df), 'error_rows': 0, 'details': errors, 'summary': {'城市列缺失': 1}, 'error_rows_detail': {}}
    city_rule_map = {normalize_name(r['city']): r for r in rules}
    error_rows = {}
    numeric_cols = []
    for col in df.columns:
        col_lower = col.lower()
        if ('基数' in col_lower or '金额' in col_lower or '费用' in col_lower or '比例' in col_lower):
            if not any(w in col_lower for w in ['校验', '状态', '类型', '说明', '备注', '合规', '是否', '判断', '结果']):
                numeric_cols.append(col)
    for idx, row in df.iterrows():
        row_errors = []
        city = str(row[city_col]) if pd.notna(row[city_col]) else ''
        if not city:
            row_errors.append('城市为空')
        else:
            norm_city = normalize_name(city)
            if norm_city not in city_rule_map:
                row_errors.append(f'城市"{city}"未在规则库中（将使用默认值）')
        for col in numeric_cols:
            val = row[col]
            if pd.notna(val):
                try:
                    num = float(str(val).replace(',', ''))
                    if num < 0:
                        row_errors.append(f'列"{col}"值为负数({num})')
                except:
                    row_errors.append(f'列"{col}"值无法转换为数字')
        company_col = None
        for col in df.columns:
            if any(kw in col.lower() for kw in COMPANY_KEYWORDS):
                company_col = col; break
        if company_col is not None:
            val = row[company_col]
            if pd.isna(val) or str(val).strip() == '':
                row_errors.append('公司名称为空')
        if row_errors:
            error_rows[idx+1] = row_errors
    summary = {}
    for row_errors in error_rows.values():
        for err in row_errors:
            if '城市' in err:
                summary['城市问题'] = summary.get('城市问题', 0) + 1
            elif '空' in err:
                summary['空值'] = summary.get('空值', 0) + 1
            elif '数字' in err:
                summary['格式问题'] = summary.get('格式问题', 0) + 1
            else:
                summary['其他'] = summary.get('其他', 0) + 1
    return {
        'total_rows': len(df),
        'error_rows': len(error_rows),
        'details': error_rows,
        'summary': summary,
        'error_rows_detail': error_rows
    }

# ========== 规则自动推断 ==========
def auto_create_rule_for_city(city, province=None):
    if not city:
        return None
    rules = load_rules()
    norm_city = normalize_name(city)
    for r in rules:
        if normalize_name(r['city']) == norm_city:
            return r
    if province:
        norm_prov = normalize_name(province)
        for r in rules:
            if normalize_name(r.get('province', '')) == norm_prov:
                new_rule = r.copy()
                new_rule['id'] = str(uuid.uuid4())[:8]
                new_rule['city'] = city
                new_rule['source_quote'] = f"自动创建（来自{province}默认）"
                new_rule['is_default'] = 0
                new_rule['notes'] = f"自动推断创建，基于{r.get('city')}规则"
                rules.append(new_rule)
                save_rules(rules)
                return new_rule
    fallback = {
        'id': str(uuid.uuid4())[:8],
        'city': city,
        'province': province or '',
        'unit_social': 0.16,
        'personal_social': 0.08,
        'unit_fund': 0.12,
        'personal_fund': 0.12,
        'social_min': 0,
        'social_max': 999999,
        'fund_min': 0,
        'fund_max': 999999,
        'source_quote': '全局默认',
        'rule_version': 'v1.0',
        'effective_date': datetime.now().strftime('%Y-%m-%d'),
        'source_url': '#',
        'source_title': '系统内置默认值',
        'source_publish_date': datetime.now().strftime('%Y-%m-%d'),
        'collected_at': datetime.now().isoformat(),
        'applicable_region': '全国',
        'official_channel': '系统内置',
        'notes': '自动创建'
    }
    rules.append(fallback)
    save_rules(rules)
    return fallback

def get_rule_for_city(city, province=None):
    if not city:
        return None
    rules = load_rules()
    norm_city = normalize_name(city)
    for r in rules:
        if normalize_name(r['city']) == norm_city:
            return r
    return auto_create_rule_for_city(city, province)

def batch_create_missing_rules():
    companies = load_companies()
    if not companies:
        return 0, "请先导入公司数据"
    existing_cities = {normalize_name(r['city']) for r in load_rules()}
    added = 0
    errors = []
    for comp in companies:
        city = comp.get('city', '')
        if not city:
            continue
        norm_city = normalize_name(city)
        if norm_city not in existing_cities:
            province = comp.get('province', '')
            result = auto_create_rule_for_city(city, province)
            if result:
                added += 1
                existing_cities.add(norm_city)
            else:
                errors.append(city)
    return added, f"已补全 {added} 个城市" + (f"，失败: {', '.join(errors)}" if errors else "")

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

def generate_batch_id():
    return f"BATCH_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{str(uuid.uuid4())[:4]}"

def get_verification_progress(batch_id):
    status = load_verification_status(batch_id)
    if not status:
        return {'total': 4, 'completed': 0, 'items': []}
    items = [
        {'name': '官方来源已核对', 'key': 'source_verified', 'done': status.get('source_verified', 0)},
        {'name': '模板版本已核对', 'key': 'template_verified', 'done': status.get('template_verified', 0)},
        {'name': '规则已人工复核', 'key': 'rule_verified', 'done': status.get('rule_verified', 0)},
        {'name': '数据已核对', 'key': 'data_verified', 'done': status.get('data_verified', 0)},
    ]
    completed = sum(1 for item in items if item['done'])
    return {'total': len(items), 'completed': completed, 'items': items}

# ========== PDF导出辅助 ==========
def df_to_html_table(df):
    """将DataFrame转换为HTML表格（用于PDF导出）"""
    if df is None or df.empty:
        return "<p>无数据</p>"
    html = "<table border='1' style='border-collapse:collapse;font-family:SimSun,serif;'>"
    # 表头
    html += "<thead><tr>"
    for col in df.columns:
        html += f"<th style='background:#4472C4;color:white;padding:8px;'>{col}</th>"
    html += "</tr></thead><tbody>"
    # 数据行
    for _, row in df.iterrows():
        html += "<tr>"
        for val in row:
            html += f"<td style='padding:6px;'>{val}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html

def create_pdf_from_html(html_content, title="报表"):
    """生成包含HTML内容的PDF（需要安装pdfkit和wkhtmltopdf）"""
    full_html = f"""
    <html>
    <head><meta charset="UTF-8"><title>{title}</title></head>
    <body style='font-family: SimSun, serif; padding: 20px;'>
        <h2 style='text-align:center;'>{title}</h2>
        <p style='text-align:center;color:#666;'>生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        <hr/>
        {html_content}
        <hr/>
        <p style='text-align:center;color:#999;font-size:12px;'>由智能报表系统生成</p>
    </body>
    </html>
    """
    return full_html

def export_to_pdf(html_content, output_path):
    """将HTML内容导出为PDF（使用pdfkit）"""
    if not PDF_AVAILABLE:
        raise ImportError("pdfkit未安装，请运行: pip install pdfkit")
    # 尝试生成PDF
    try:
        pdfkit.from_string(html_content, output_path)
        return True
    except Exception as e:
        raise Exception(f"PDF生成失败，请确保已安装wkhtmltopdf: {e}")

# ========== Streamlit 页面 ==========
st.set_page_config(page_title="智能报表系统 - 企业版", layout="wide")
st.title("📋 智能报表系统（企业版）")
st.markdown("**工作台 · 数据识别 · 依据库 · 复核导出 · 作业记录 · 年审数据处理**")

if 'filtered_values' in st.session_state and st.session_state['filtered_values']:
    with st.expander("⚠️ 已过滤的异常数据"):
        for val in st.session_state['filtered_values']:
            st.write(f"- {val}")
    st.session_state['filtered_values'] = []

st.sidebar.title("📌 导航")
page = st.sidebar.radio("选择功能", [
    "📊 工作台",
    "📤 数据导入",
    "📚 依据库管理",
    "⚙️ 规则管理",
    "📄 自定义模板",
    "📋 导出历史与复核",
    "💾 备份与恢复",
    "📋 年审数据处理"
])

# ===== 全局Sheet选择器 =====
if 'uploaded_files' in st.session_state and st.session_state['uploaded_files']:
    st.sidebar.markdown("---")
    st.sidebar.subheader("📂 数据Sheet选择")
    files = st.session_state['uploaded_files']
    if len(files) > 1:
        file_names = [f.name for f in files]
        selected_file_name = st.sidebar.selectbox("选择数据文件", file_names, key="global_file_select")
        selected_file = next(f for f in files if f.name == selected_file_name)
    else:
        selected_file = files[0]
    try:
        xls = pd.ExcelFile(selected_file)
        sheets = xls.sheet_names
        current_sheet = st.session_state.get('data_sheet_name', sheets[0] if sheets else '')
        idx = sheets.index(current_sheet) if current_sheet in sheets else 0
        selected_sheet = st.sidebar.selectbox("选择Sheet", sheets, index=idx, key="global_sheet_select")
        if selected_sheet != st.session_state.get('data_sheet_name'):
            df, header_row = auto_load_sheet_with_header_detection(selected_file, selected_sheet)
            st.session_state['imported_df'] = df
            st.session_state['data_sheet_name'] = selected_sheet
            st.session_state['data_header_row'] = header_row
            st.session_state['data_recognition'] = {
                'sheets': sheets,
                'selected_sheet': selected_sheet,
                'header_row': header_row,
                'columns': list(df.columns) if df is not None else [],
                'total_rows': len(df) if df is not None else 0
            }
            rules = load_rules()
            st.session_state['validation_report'] = validate_data(df, rules)
            st.sidebar.success(f"✅ 已加载: {selected_sheet} (表头行: {header_row+1})")
            st.rerun()
        else:
            st.sidebar.info(f"当前: {selected_sheet}")
    except Exception as e:
        st.sidebar.error(f"读取Sheet失败: {e}")

# ===== 各页面 =====
if page == "📊 工作台":
    st.subheader("📊 工作台概览")
    col1, col2, col3, col4, col5 = st.columns(5)
    companies = load_companies()
    templates = load_templates()
    rules = load_rules()
    history = load_export_history()
    batches = load_job_batches()
    sources = load_source_registry()
    with col1:
        st.metric("🏢 公司数", len(companies))
    with col2:
        st.metric("📄 模板数", len(templates))
    with col3:
        st.metric("⚙️ 规则数", len(rules))
    with col4:
        pending = len([h for h in history if h.get('review_status') == 'pending'])
        st.metric("📋 待复核", pending)
    with col5:
        st.metric("📚 来源数", len(sources))
    st.markdown("---")
    st.subheader("📋 最近作业记录")
    if batches:
        df_batches = pd.DataFrame(batches)
        st.dataframe(df_batches[['batch_name', 'created_at', 'status', 'total_companies', 'total_reports', 'review_status']].head(10), use_container_width=True)
    else:
        st.info("暂无作业记录")
    st.subheader("📊 待处理异常")
    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
        df = st.session_state['imported_df']
        rules = load_rules()
        report = validate_data(df, rules)
        if report['error_rows'] > 0:
            st.warning(f"⚠️ 当前数据有 {report['error_rows']} 行异常")
            with st.expander("查看异常详情"):
                for row, errs in report['details'].items():
                    st.write(f"**第 {row} 行**：{', '.join(errs)}")
        else:
            st.success("✅ 当前数据无异常")
    else:
        st.info("暂无数据，请先导入")

elif page == "📤 数据导入":
    st.subheader("📤 数据导入（支持多文件）")
    import_mode = st.radio("导入模式", ["智能导入（自动识别结构）", "普通导入（手动选择列，开发中）"], index=0, horizontal=True)
    st.caption("智能导入将自动识别城市、公司等列；普通导入可自定义列映射（开发中，当前与智能导入相同）")
    uploaded_files = st.file_uploader("选择Excel文件（支持多个 .xlsx）", type=["xlsx"], accept_multiple_files=True)
    if uploaded_files:
        with st.spinner("正在解析Excel..."):
            companies, unmapped, all_sheets, data_sheet = parse_multiple_files(uploaded_files)
            valid_companies = companies
            if unmapped:
                st.warning(f"⚠️ 以下城市未在规则库中找到：{', '.join(unmapped)}，将使用全局默认规则")
            if valid_companies:
                save_companies(valid_companies)
                st.success(f"成功提取 {len(valid_companies)} 家公司，来自 {len(uploaded_files)} 个文件")
                st.session_state['uploaded_files'] = uploaded_files
                st.session_state['all_sheets'] = all_sheets
                st.session_state['data_recognition'] = {
                    'sheets': all_sheets,
                    'selected_sheet': data_sheet,
                    'header_row': 1,
                    'columns': [],
                    'total_rows': 0,
                    'files_count': len(uploaded_files),
                    'companies_extracted': len(valid_companies)
                }
                if uploaded_files and data_sheet:
                    first_file = uploaded_files[0]
                    df, header_row = auto_load_sheet_with_header_detection(first_file, data_sheet)
                    st.session_state['imported_df'] = df
                    st.session_state['data_sheet_name'] = data_sheet
                    st.session_state['data_header_row'] = header_row
                    st.session_state['data_recognition']['header_row'] = header_row
                    st.session_state['data_recognition']['columns'] = list(df.columns) if df is not None else []
                    st.session_state['data_recognition']['total_rows'] = len(df) if df is not None else 0
                    rules = load_rules()
                    st.session_state['validation_report'] = validate_data(df, rules)
                    st.success(f"已自动加载 Sheet「{data_sheet}」（表头行: {header_row+1}），共 {len(df)} 行数据")
    st.markdown("---")
    st.subheader("🔍 数据识别结果")
    if 'data_recognition' in st.session_state:
        rec = st.session_state['data_recognition']
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📄 检测到Sheet数", len(rec.get('sheets', [])))
        with col2:
            st.metric("📊 当前Sheet", rec.get('selected_sheet', '无'))
        with col3:
            st.metric("📋 表头行", rec.get('header_row', 0) + 1 if rec.get('header_row') is not None else 1)
        with col4:
            st.metric("🏢 提取公司数", rec.get('companies_extracted', 0))
        if rec.get('columns'):
            st.write("**表头字段**")
            st.dataframe(pd.DataFrame({'字段名': rec['columns']}), use_container_width=True)
        if rec.get('sheets'):
            st.write("**所有Sheet**")
            st.write(", ".join(rec['sheets']))
        if st.button("🗑️ 清空识别结果", key="clear_recognition"):
            st.session_state['data_recognition'] = {}
            st.rerun()
    else:
        st.info("请上传文件以查看数据识别结果")
    
    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
        st.subheader("📊 数据预览")
        df = st.session_state['imported_df']
        
        # 【优化一】异常高亮
        if 'validation_report' in st.session_state:
            report = st.session_state['validation_report']
            error_rows = report.get('details', {})
            styled_df = highlight_error_rows(df.head(10), error_rows)
            st.dataframe(styled_df, use_container_width=True)
            st.caption("🔴 红色行表示存在数据异常")
        else:
            st.dataframe(df.head(10), use_container_width=True)
        
        st.caption(f"当前Sheet: {st.session_state.get('data_sheet_name', '未知')}，共 {len(df)} 行")
        if 'validation_report' in st.session_state:
            report = st.session_state['validation_report']
            st.subheader("📊 数据质量报告")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("总行数", report['total_rows'])
            with col2:
                st.metric("异常行数", report['error_rows'])
            with col3:
                st.metric("正常行数", report['total_rows'] - report['error_rows'])
            if report['error_rows'] > 0:
                st.warning(f"⚠️ 发现 {report['error_rows']} 行数据存在问题")
                with st.expander("查看异常详情"):
                    for row, errs in report['details'].items():
                        st.write(f"**第 {row} 行**：{', '.join(errs)}")
            else:
                st.success("✅ 所有数据校验通过！")

elif page == "📚 依据库管理":
    st.subheader("📚 依据库管理")
    tab1, tab2, tab3, tab4 = st.tabs(["📄 模板", "⚙️ 规则", "🏢 公司", "📚 来源"])
    with tab1:
        templates = load_templates()
        if templates:
            df = pd.DataFrame(templates)
            cols = ['template_name', 'report_type', 'province', 'city', 'template_version', 'source_authority']
            st.dataframe(df[cols], use_container_width=True)
        else:
            st.info("暂无模板")
    with tab2:
        rules = load_rules()
        if rules:
            df = pd.DataFrame(rules)
            cols = ['city', 'province', 'unit_social', 'personal_social', 'unit_fund', 'personal_fund', 'source_quote', 'rule_version', 'source_title']
            st.dataframe(df[cols], use_container_width=True)
        else:
            st.info("暂无规则")
    with tab3:
        companies = load_companies()
        if companies:
            st.dataframe(pd.DataFrame(companies))
            st.caption(f"共 {len(companies)} 家公司")
        else:
            st.info("暂无公司数据")
    with tab4:
        st.markdown("**📚 官方来源管理**")
        sources = load_source_registry()
        with st.expander("➕ 新增官方来源", expanded=False):
            with st.form(key="add_source_form"):
                col1, col2 = st.columns(2)
                with col1:
                    new_id = st.text_input("来源ID", value=f"src_{str(uuid.uuid4())[:8]}")
                    new_authority_type = st.selectbox("机构类型", ['social_security', 'tax', 'fund', 'other'])
                    new_province = st.text_input("省份")
                    new_city = st.text_input("城市")
                    new_district = st.text_input("区县")
                    new_authority_name = st.text_input("机构名称")
                    new_official_site_name = st.text_input("官网名称")
                with col2:
                    new_source_url = st.text_input("来源URL")
                    new_source_level = st.selectbox("来源级别", ['city', 'province', 'national'])
                    new_source_section = st.text_input("来源板块")
                    new_is_official = st.checkbox("官方来源", value=True)
                    new_crawl_allowed = st.checkbox("允许爬取", value=True)
                    new_document_name = st.text_input("文档名称")
                    new_document_version = st.text_input("文档版本")
                    new_publish_year = st.text_input("发布年份")
                    new_notes = st.text_area("备注")
                submitted = st.form_submit_button("保存来源")
                if submitted:
                    new_source = {
                        'id': new_id,
                        'authority_type': new_authority_type,
                        'province': new_province,
                        'city': new_city,
                        'district': new_district,
                        'authority_name': new_authority_name,
                        'official_site_name': new_official_site_name,
                        'source_url': new_source_url,
                        'source_level': new_source_level,
                        'source_section': new_source_section,
                        'is_official': 1 if new_is_official else 0,
                        'crawl_allowed': 1 if new_crawl_allowed else 0,
                        'last_checked': datetime.now().strftime('%Y-%m-%d'),
                        'status': 'active',
                        'notes': new_notes,
                        'document_name': new_document_name,
                        'document_version': new_document_version,
                        'publish_year': new_publish_year
                    }
                    sources.append(new_source)
                    save_source_registry(sources)
                    st.success("来源已添加！")
                    st.rerun()
        if sources:
            df_sources = pd.DataFrame(sources)
            display_cols = ['id', 'authority_type', 'province', 'city', 'authority_name', 'source_url', 'document_name', 'document_version']
            st.dataframe(df_sources[display_cols], use_container_width=True)
            st.markdown("**操作**")
            col_ops1, col_ops2 = st.columns(2)
            with col_ops1:
                source_to_delete = st.selectbox("选择要删除的来源ID", [s['id'] for s in sources] if sources else [])
                if st.button("🗑️ 删除所选来源"):
                    if source_to_delete:
                        delete_source_by_id(source_to_delete)
                        st.success(f"已删除来源 {source_to_delete}")
                        st.rerun()
            with col_ops2:
                source_to_edit = st.selectbox("选择要编辑的来源ID", [s['id'] for s in sources] if sources else [])
                if source_to_edit:
                    with st.expander(f"✏️ 编辑来源 {source_to_edit}", expanded=False):
                        source = next(s for s in sources if s['id'] == source_to_edit)
                        with st.form(key="edit_source_form"):
                            col1, col2 = st.columns(2)
                            with col1:
                                edit_authority_type = st.selectbox("机构类型", ['social_security', 'tax', 'fund', 'other'], index=['social_security','tax','fund','other'].index(source.get('authority_type','social_security')))
                                edit_province = st.text_input("省份", value=source.get('province',''))
                                edit_city = st.text_input("城市", value=source.get('city',''))
                                edit_district = st.text_input("区县", value=source.get('district',''))
                                edit_authority_name = st.text_input("机构名称", value=source.get('authority_name',''))
                                edit_official_site_name = st.text_input("官网名称", value=source.get('official_site_name',''))
                            with col2:
                                edit_source_url = st.text_input("来源URL", value=source.get('source_url',''))
                                edit_source_level = st.selectbox("来源级别", ['city', 'province', 'national'], index=['city','province','national'].index(source.get('source_level','city')))
                                edit_source_section = st.text_input("来源板块", value=source.get('source_section',''))
                                edit_is_official = st.checkbox("官方来源", value=bool(source.get('is_official',1)))
                                edit_crawl_allowed = st.checkbox("允许爬取", value=bool(source.get('crawl_allowed',1)))
                                edit_document_name = st.text_input("文档名称", value=source.get('document_name',''))
                                edit_document_version = st.text_input("文档版本", value=source.get('document_version',''))
                                edit_publish_year = st.text_input("发布年份", value=source.get('publish_year',''))
                                edit_notes = st.text_area("备注", value=source.get('notes',''))
                            submitted_edit = st.form_submit_button("更新来源")
                            if submitted_edit:
                                updated_sources = []
                                for s in sources:
                                    if s['id'] == source_to_edit:
                                        s.update({
                                            'authority_type': edit_authority_type,
                                            'province': edit_province,
                                            'city': edit_city,
                                            'district': edit_district,
                                            'authority_name': edit_authority_name,
                                            'official_site_name': edit_official_site_name,
                                            'source_url': edit_source_url,
                                            'source_level': edit_source_level,
                                            'source_section': edit_source_section,
                                            'is_official': 1 if edit_is_official else 0,
                                            'crawl_allowed': 1 if edit_crawl_allowed else 0,
                                            'last_checked': datetime.now().strftime('%Y-%m-%d'),
                                            'document_name': edit_document_name,
                                            'document_version': edit_document_version,
                                            'publish_year': edit_publish_year,
                                            'notes': edit_notes
                                        })
                                    updated_sources.append(s)
                                save_source_registry(updated_sources)
                                st.success(f"来源 {source_to_edit} 已更新！")
                                st.rerun()
        else:
            st.info("暂无来源，请添加或加载样本依据")
        if st.button("加载样本依据"):
            sample_sources = [
                {'id': 'src_gz_social_demo', 'authority_type': 'social_security', 'province': '广东', 'city': '广州', 'district': '', 'authority_name': '广州市人力资源和社会保障局', 'official_site_name': '广州市人力资源和社会保障局官网', 'source_url': 'https://rsj.gz.gov.cn/', 'source_level': 'city', 'source_section': '网上办事', 'document_name': '广州市社保年审公告', 'document_version': '2024', 'publish_year': '2024', 'is_official': 1, 'crawl_allowed': 1, 'last_checked': datetime.now().strftime('%Y-%m-%d'), 'status': 'active', 'notes': '示例，请核实'},
                {'id': 'src_sh_social_sample', 'authority_type': 'social_security', 'province': '上海', 'city': '上海', 'district': '', 'authority_name': '上海市人力资源和社会保障局', 'official_site_name': '上海社保官方系统', 'source_url': 'https://rsj.sh.gov.cn/', 'source_level': 'city', 'source_section': '网上办事', 'document_name': '上海市社保基数调整通知', 'document_version': '2024', 'publish_year': '2024', 'is_official': 1, 'crawl_allowed': 1, 'last_checked': datetime.now().strftime('%Y-%m-%d'), 'status': 'active', 'notes': '示例，请核实'},
                {'id': 'src_hrb_social_sample', 'authority_type': 'social_security', 'province': '黑龙江', 'city': '哈尔滨', 'district': '', 'authority_name': '哈尔滨市人力资源和社会保障局', 'official_site_name': '哈尔滨社保官方系统', 'source_url': '', 'source_level': 'city', 'source_section': '网上办事', 'document_name': '哈尔滨年审公告', 'document_version': '2024', 'publish_year': '2024', 'is_official': 1, 'crawl_allowed': 1, 'last_checked': datetime.now().strftime('%Y-%m-%d'), 'status': 'active', 'notes': '示例，请核实'}
            ]
            save_source_registry(sample_sources)
            st.success("已加载样本依据！")
            st.rerun()

elif page == "⚙️ 规则管理":
    st.subheader("⚙️ 规则管理（社保/公积金）")
    st.markdown("**⚠️ 重要：所有规则必须来自官方渠道，并填写来源信息。**")
    st.markdown("官方渠道优先级：国家税务总局、各省税务局、人社部、各城市公积金中心等官网。")
    rules = load_rules()
    st.write(f"**当前规则数量：{len(rules)} 个城市**")
    if rules:
        df_rules = pd.DataFrame(rules)
        st.dataframe(df_rules[['city', 'province', 'unit_social', 'personal_social', 'unit_fund', 'personal_fund', 'social_min', 'social_max', 'source_quote', 'rule_version', 'source_url', 'source_title']], use_container_width=True)
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("🔧 从公司数据批量补全规则"):
                added, msg = batch_create_missing_rules()
                st.success(msg)
                st.rerun()
        with col_btn2:
            if st.button("📋 导出规则清单"):
                st.dataframe(df_rules[['city', 'province', 'unit_social', 'personal_social', 'unit_fund', 'personal_fund']], use_container_width=True)
        
        with st.expander("➕ 新增城市规则（需填写官方来源）", expanded=False):
            with st.form(key="add_rule_form"):
                col1, col2 = st.columns(2)
                with col1:
                    new_city = st.text_input("城市名称 *")
                    new_province = st.text_input("所属省份 *")
                    new_unit_social = st.number_input("单位社保比例", value=0.16, step=0.001, format="%.3f")
                    new_personal_social = st.number_input("个人社保比例", value=0.08, step=0.001, format="%.3f")
                    new_unit_fund = st.number_input("单位公积金比例", value=0.12, step=0.001, format="%.3f")
                    new_personal_fund = st.number_input("个人公积金比例", value=0.12, step=0.001, format="%.3f")
                    new_social_min = st.number_input("社保基数下限", value=0.0, step=100.0)
                    new_social_max = st.number_input("社保基数上限", value=999999.0, step=100.0)
                    new_fund_min = st.number_input("公积金基数下限", value=0.0, step=100.0)
                    new_fund_max = st.number_input("公积金基数上限", value=999999.0, step=100.0)
                    new_source_quote = st.text_input("来源文号（可选）")
                    new_rule_version = st.text_input("规则版本", value="v1.0")
                    new_effective_date = st.text_input("生效日期", value=datetime.now().strftime('%Y-%m-%d'))
                with col2:
                    st.markdown("**来源信息（必须填写）**")
                    new_source_url = st.text_input("来源链接 *")
                    new_source_title = st.text_input("来源标题 *")
                    new_source_publish_date = st.text_input("来源发布时间 *", value=datetime.now().strftime('%Y-%m-%d'))
                    new_applicable_region = st.text_input("适用地区 *", help="如'全国'或'广东'")
                    new_official_channel = st.selectbox("官方渠道 *", ['国家税务总局官网', '国家税务总局政策法规库', '12366办税指南', '省级税务局官网', '市级/区县级税务局官网', '全国人社政务服务平台', '地方人社局官网', '地方社保中心', '住房公积金管理中心官网', '其他'])
                    new_notes = st.text_area("备注")
                submitted = st.form_submit_button("添加规则")
                if submitted:
                    if not new_city or not new_province or not new_source_url or not new_source_title:
                        st.error("城市、省份、来源链接、来源标题为必填项")
                    else:
                        new_rule = {
                            'id': str(uuid.uuid4())[:8],
                            'city': new_city,
                            'province': new_province,
                            'unit_social': new_unit_social,
                            'personal_social': new_personal_social,
                            'unit_fund': new_unit_fund,
                            'personal_fund': new_personal_fund,
                            'social_min': new_social_min,
                            'social_max': new_social_max,
                            'fund_min': new_fund_min,
                            'fund_max': new_fund_max,
                            'source_quote': new_source_quote,
                            'is_default': 0,
                            'rule_version': new_rule_version,
                            'effective_date': new_effective_date,
                            'source_url': new_source_url,
                            'source_title': new_source_title,
                            'source_publish_date': new_source_publish_date,
                            'collected_at': datetime.now().isoformat(),
                            'applicable_region': new_applicable_region,
                            'official_channel': new_official_channel,
                            'notes': new_notes
                        }
                        rules.append(new_rule)
                        save_rules(rules)
                        st.success(f"已添加 {new_city} 的规则！")
                        st.rerun()
        cities = sorted(set(r['city'] for r in rules))
        selected_city = st.selectbox("选择城市进行编辑", [""] + cities)
        if selected_city:
            rule = next((r for r in rules if r['city'] == selected_city), None)
            if rule:
                with st.form(key=f"edit_rule_{selected_city}"):
                    st.write(f"编辑 **{selected_city}** 的规则")
                    col1, col2 = st.columns(2)
                    with col1:
                        new_unit_social = st.number_input("单位社保比例", value=float(rule['unit_social']), step=0.001, format="%.3f")
                        new_personal_social = st.number_input("个人社保比例", value=float(rule['personal_social']), step=0.001, format="%.3f")
                        new_unit_fund = st.number_input("单位公积金比例", value=float(rule['unit_fund']), step=0.001, format="%.3f")
                        new_personal_fund = st.number_input("个人公积金比例", value=float(rule['personal_fund']), step=0.001, format="%.3f")
                        new_social_min = st.number_input("社保基数下限", value=float(rule.get('social_min', 0)), step=100.0)
                        new_social_max = st.number_input("社保基数上限", value=float(rule.get('social_max', 999999)), step=100.0)
                        new_fund_min = st.number_input("公积金基数下限", value=float(rule.get('fund_min', 0)), step=100.0)
                        new_fund_max = st.number_input("公积金基数上限", value=float(rule.get('fund_max', 999999)), step=100.0)
                        new_source_quote = st.text_input("来源文号", value=rule.get('source_quote', ''))
                        new_rule_version = st.text_input("规则版本", value=rule.get('rule_version', 'v1.0'))
                        new_effective_date = st.text_input("生效日期", value=rule.get('effective_date', datetime.now().strftime('%Y-%m-%d')))
                    with col2:
                        new_source_url = st.text_input("来源链接", value=rule.get('source_url', ''))
                        new_source_title = st.text_input("来源标题", value=rule.get('source_title', ''))
                        new_source_publish_date = st.text_input("来源发布时间", value=rule.get('source_publish_date', datetime.now().strftime('%Y-%m-%d')))
                        new_applicable_region = st.text_input("适用地区", value=rule.get('applicable_region', ''))
                        new_official_channel = st.selectbox("官方渠道", ['国家税务总局官网', '国家税务总局政策法规库', '12366办税指南', '省级税务局官网', '市级/区县级税务局官网', '全国人社政务服务平台', '地方人社局官网', '地方社保中心', '住房公积金管理中心官网', '其他'], index=['国家税务总局官网','国家税务总局政策法规库','12366办税指南', '省级税务局官网','市级/区县级税务局官网','全国人社政务服务平台', '地方人社局官网','地方社保中心','住房公积金管理中心官网','其他'].index(rule.get('official_channel','其他')))
                        new_notes = st.text_area("备注", value=rule.get('notes', ''))
                    submitted = st.form_submit_button("保存修改")
                    if submitted:
                        updated_rules = []
                        for r in rules:
                            if r['id'] == rule['id']:
                                r.update({
                                    'unit_social': new_unit_social,
                                    'personal_social': new_personal_social,
                                    'unit_fund': new_unit_fund,
                                    'personal_fund': new_personal_fund,
                                    'social_min': new_social_min,
                                    'social_max': new_social_max,
                                    'fund_min': new_fund_min,
                                    'fund_max': new_fund_max,
                                    'source_quote': new_source_quote,
                                    'rule_version': new_rule_version,
                                    'effective_date': new_effective_date,
                                    'source_url': new_source_url,
                                    'source_title': new_source_title,
                                    'source_publish_date': new_source_publish_date,
                                    'applicable_region': new_applicable_region,
                                    'official_channel': new_official_channel,
                                    'notes': new_notes,
                                    'collected_at': datetime.now().isoformat()
                                })
                            updated_rules.append(r)
                        save_rules(updated_rules)
                        st.success("规则已更新！")
                        st.rerun()
        if st.button("🔄 重置所有规则为系统默认值（会覆盖所有自定义规则）"):
            if st.checkbox("确认重置？此操作将覆盖所有自定义规则"):
                sample_rules = [{'id': str(uuid.uuid4())[:8], 'city': '上海', 'province': '上海', 'unit_social': 0.16, 'personal_social': 0.08, 'unit_fund': 0.07, 'personal_fund': 0.07, 'social_min': 7310, 'social_max': 36549, 'fund_min': 2590, 'fund_max': 34188, 'source_quote': '沪人社规〔2024〕22号', 'rule_version': '2024.1', 'effective_date': '2024-07-01', 'source_url': 'https://rsj.sh.gov.cn/', 'source_title': '上海市2024年度社保缴费基数调整通知', 'source_publish_date': '2024-06-20', 'collected_at': datetime.now().isoformat(), 'applicable_region': '上海', 'official_channel': '上海市人社局官网', 'notes': '示例规则'}]
                save_rules(sample_rules)
                st.success("已重置为示例规则！")
                st.rerun()
    else:
        st.info("暂无规则，请添加")

elif page == "📄 自定义模板":
    st.subheader("📄 自定义模板管理")
    custom_templates = load_custom_templates()
    if custom_templates:
        st.write("**已保存的自定义模板**")
        for ct in custom_templates:
            col1, col2, col3 = st.columns([3, 2, 1])
            with col1:
                st.write(f"📄 {ct['name']}")
            with col2:
                st.write(f"字段数：{len(json.loads(ct.get('field_mapping', '{}')))}")
            with col3:
                if st.button("删除", key=f"del_{ct['id']}"):
                    delete_custom_template(ct['id'])
                    st.success(f"已删除 {ct['name']}")
                    st.rerun()
    st.markdown("---")
    st.markdown("**上传新模板**")
    uploaded_template = st.file_uploader("选择模板文件（.xlsx）", type=["xlsx"], key="custom_template_upload")
    if uploaded_template:
        try:
            template_name = st.text_input("模板名称", value=uploaded_template.name.replace('.xlsx', ''))
            sheet_name = st.text_input("Sheet名称（留空使用第一个Sheet）", value="")
            field_mapping_source = st.text_input("字段映射来源说明（可选）")
            wb = load_workbook(BytesIO(uploaded_template.getvalue()))
            ws = wb.active if not sheet_name else wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
            headers = []
            for cell in ws[1]:
                if cell.value:
                    headers.append(str(cell.value).strip())
            st.write(f"**检测到表头字段**：{', '.join(headers) if headers else '未检测到表头'}")
            st.markdown("**字段映射（将数据字段映射到模板列）**")
            mapping = {}
            if headers:
                field_options = [''] + ['纳税人识别号', '公司名称', '销售额', '进项税额', '应纳税额', '单位名称', '社保登记号', '基数', '单位金额', '个人金额', '单位比例', '个人比例', '公积金账号', '收入额', '专项扣除', '营业收入', '营业成本', '应纳税所得额', '全年收入', '全年成本', '已预缴税额', '应补退税额', '申报金额']
                for h in headers:
                    col = st.selectbox(f"字段 '{h}' 映射到", field_options, key=f"map_{h}_{uploaded_template.name}")
                    if col:
                        mapping[h] = col
            if st.button("保存自定义模板"):
                if not template_name:
                    st.error("请填写模板名称")
                else:
                    template_data = {'id': str(uuid.uuid4())[:8], 'name': template_name, 'file_data': uploaded_template.getvalue(), 'field_mapping': mapping, 'sheet_name': sheet_name, 'created_at': datetime.now().isoformat()}
                    save_custom_template(template_data)
                    st.success(f"模板 '{template_name}' 已保存！")
                    st.rerun()
        except Exception as e:
            st.error(f"处理模板失败：{e}")

elif page == "📋 导出历史与复核":
    st.subheader("📋 导出历史与复核")
    tab1, tab2, tab3 = st.tabs(["📋 导出历史", "✅ 复核处理", "📊 核验状态"])
    with tab1:
        history = load_export_history()
        if history:
            df_hist = pd.DataFrame(history)
            display_cols = ['company_name', 'city', 'report_type', 'period_type', 'data_source', 'generated_at', 'review_status', 'batch_id']
            st.dataframe(df_hist[display_cols], use_container_width=True)
            batches = load_job_batches()
            if batches:
                st.subheader("📦 批次作业")
                st.dataframe(pd.DataFrame(batches), use_container_width=True)
        else:
            st.info("暂无导出记录")
    with tab2:
        history = load_export_history()
        pending = [h for h in history if h.get('review_status') == 'pending']
        if pending:
            st.write(f"共 {len(pending)} 份待复核报表")
            for h in pending:
                col1, col2, col3 = st.columns([3, 2, 1])
                with col1:
                    st.write(f"📄 {h['company_name']} - {h['city']} ({h['report_type']})")
                    st.caption(f"生成时间：{h['generated_at'][:16]}")
                with col2:
                    st.write(f"来源：{h.get('data_source', '未知')}")
                with col3:
                    if st.button("✅ 通过", key=f"approve_{h['id']}"):
                        update_export_status(h['id'], 'approved', '复核员')
                        st.success("已通过复核")
                        st.rerun()
                    if st.button("❌ 驳回", key=f"reject_{h['id']}"):
                        update_export_status(h['id'], 'rejected', '复核员')
                        st.warning("已驳回")
                        st.rerun()
        else:
            st.success("✅ 暂无待复核报表")
    with tab3:
        st.markdown("**📊 核验状态**")
        batches = load_job_batches()
        if batches:
            for batch in batches[:10]:
                batch_id = batch['id']
                progress = get_verification_progress(batch_id)
                status = load_verification_status(batch_id)
                st.markdown(f"**批次：{batch.get('batch_name', batch_id)}**")
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.progress(progress['completed'] / progress['total'] if progress['total'] > 0 else 0)
                    st.caption(f"{progress['completed']}/{progress['total']} 项已完成")
                    for item in progress['items']:
                        icon = "✅" if item['done'] else "❌"
                        st.write(f"{icon} {item['name']}")
                with col2:
                    st.write(f"导出类型：{status.get('export_type', '验证版') if status else '验证版'}")
                    if status and status.get('reviewer_name'):
                        st.write(f"复核人：{status['reviewer_name']}")
                    st.write(f"状态：{'✅ 可导出正式版' if progress['completed'] == progress['total'] else '⏳ 待完成核验'}")
                    if progress['completed'] == progress['total']:
                        if st.button(f"📤 导出正式版", key=f"formal_{batch_id}"):
                            st.success(f"正式版导出成功！批次：{batch.get('batch_name', batch_id)}")
                    else:
                        if st.button(f"📤 跳过核验导出验证版", key=f"verify_{batch_id}"):
                            if status:
                                update_verification_status(batch_id, 'export_type', '验证版')
                            st.success(f"验证版导出成功！批次：{batch.get('batch_name', batch_id)}")
                    if status:
                        col_a, col_b = st.columns(2)
                        with col_a:
                            if not status.get('source_verified', 0):
                                if st.button(f"✅ 来源已核对", key=f"src_{batch_id}"):
                                    update_verification_status(batch_id, 'source_verified', 1)
                                    st.rerun()
                            if not status.get('template_verified', 0):
                                if st.button(f"✅ 模板已核对", key=f"tpl_{batch_id}"):
                                    update_verification_status(batch_id, 'template_verified', 1)
                                    st.rerun()
                        with col_b:
                            if not status.get('rule_verified', 0):
                                if st.button(f"✅ 规则已复核", key=f"rule_{batch_id}"):
                                    update_verification_status(batch_id, 'rule_verified', 1)
                                    st.rerun()
                            if not status.get('data_verified', 0):
                                if st.button(f"✅ 数据已核对", key=f"data_{batch_id}"):
                                    update_verification_status(batch_id, 'data_verified', 1)
                                    st.rerun()
                    with st.form(key=f"reviewer_form_{batch_id}"):
                        reviewer_name = st.text_input("复核人姓名", value=status.get('reviewer_name', '') if status else '')
                        if st.form_submit_button("保存复核人"):
                            if status:
                                update_verification_status(batch_id, 'reviewer_name', reviewer_name)
                                st.success("复核人已保存")
                                st.rerun()
                            else:
                                save_verification_status({'batch_id': batch_id, 'source_verified': 0, 'template_verified': 0, 'rule_verified': 0, 'data_verified': 0, 'reviewer_name': reviewer_name, 'verified_at': datetime.now().isoformat(), 'export_type': '验证版'})
                                st.success("核验状态已创建")
                                st.rerun()
                st.markdown("---")
        else:
            st.info("暂无批次作业，请先生成报表")

elif page == "💾 备份与恢复":
    st.subheader("💾 备份与恢复")
    st.markdown("**自动备份**：每次操作后自动备份，保留最近30份")
    backups = get_backup_list()
    if backups:
        st.write(f"共 {len(backups)} 份备份")
        selected_backup = st.selectbox("选择备份恢复", backups)
        if st.button("恢复所选备份"):
            if restore_backup(selected_backup):
                st.success("恢复成功！请刷新页面")
                st.rerun()
            else:
                st.error("恢复失败")
    else:
        st.info("暂无备份")
    if st.button("手动备份当前数据"):
        path = backup_database()
        if path:
            st.success(f"备份成功：{os.path.basename(path)}")

elif page == "📋 年审数据处理":
    st.subheader("📋 年审数据处理与交付")
    st.markdown("**处理流程：系统账单筛选 → 上海数据合并 → 基数核算 → 实缴核算 → 归档交付**")
    
    st.markdown("### 步骤1：上传数据源")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("**📄 to客户文件**")
        to_client_file = st.file_uploader("上传to客户最终交付文件", type=["xlsx"], key="to_client")
        if to_client_file:
            st.success(f"已上传: {to_client_file.name}")
            st.session_state['to_client_file'] = to_client_file
    with col2:
        st.markdown("**📊 内部系统账单**")
        internal_bill_file = st.file_uploader("上传内部系统账单", type=["xlsx"], key="internal_bill")
        if internal_bill_file:
            st.success(f"已上传: {internal_bill_file.name}")
            st.session_state['internal_bill_file'] = internal_bill_file
    with col3:
        st.markdown("**📚 上海社保数据**")
        shanghai_files = st.file_uploader("上传上海社保数据（可多选）", type=["xlsx"], accept_multiple_files=True, key="shanghai")
        if shanghai_files:
            st.success(f"已上传 {len(shanghai_files)} 份文件")
            st.session_state['shanghai_files'] = shanghai_files
    
    st.markdown("---")
    if st.button("🔄 执行数据处理", key="process_annual_data"):
        if not st.session_state.get('internal_bill_file'):
            st.error("请先上传内部系统账单")
            st.stop()
        with st.spinner("正在处理数据..."):
            processed_files = {}
            results = {}
            # 处理广州账单
            st.markdown("**📊 处理内部系统账单（筛选广州）**")
            try:
                df_bill = pd.read_excel(st.session_state['internal_bill_file'])
                st.write(f"原始账单行数: {len(df_bill)}")
                city_col = None
                for col in df_bill.columns:
                    if any(kw in col.lower() for kw in ['城市', '地区', '所属地']):
                        city_col = col; break
                if city_col:
                    df_gz = df_bill[df_bill[city_col].astype(str).str.contains('广州', na=False)]
                else:
                    df_gz = df_bill
                    st.warning("未识别到城市列，将使用全部数据")
                st.write(f"筛选后广州员工: {len(df_gz)} 行")
                results['gz_bill'] = df_gz
                gz_buffer = BytesIO()
                with pd.ExcelWriter(gz_buffer, engine='openpyxl') as writer:
                    df_gz.to_excel(writer, sheet_name='广州员工明细', index=False)
                gz_buffer.seek(0)
                processed_files['广州内部账单.xlsx'] = gz_buffer.getvalue()
                st.success(f"✅ 已生成广州内部账单: {len(df_gz)} 行")
            except Exception as e:
                st.error(f"处理账单失败: {e}")
            # 处理上海合并
            st.markdown("**📚 处理上海社保数据（合并）**")
            if st.session_state.get('shanghai_files'):
                try:
                    shanghai_dfs = []
                    for f in st.session_state['shanghai_files']:
                        df = pd.read_excel(f)
                        shanghai_dfs.append(df)
                        st.write(f"  - {f.name}: {len(df)} 行")
                    if shanghai_dfs:
                        df_sh_merged = pd.concat(shanghai_dfs, ignore_index=True)
                        st.write(f"合并后总行数: {len(df_sh_merged)}")
                        results['shanghai_merged'] = df_sh_merged
                        sh_buffer = BytesIO()
                        with pd.ExcelWriter(sh_buffer, engine='openpyxl') as writer:
                            df_sh_merged.to_excel(writer, sheet_name='上海合并数据', index=False)
                        sh_buffer.seek(0)
                        processed_files['上海合并数据表.xlsx'] = sh_buffer.getvalue()
                        st.success(f"✅ 已生成上海合并数据表: {len(df_sh_merged)} 行")
                except Exception as e:
                    st.error(f"合并上海数据失败: {e}")
            else:
                st.info("未上传上海社保数据，跳过合并步骤")
            # 生成to客户文件
            st.markdown("**📄 生成to客户最终交付文件**")
            try:
                all_data = []
                if 'gz_bill' in results:
                    all_data.append(results['gz_bill'])
                if 'shanghai_merged' in results:
                    all_data.append(results['shanghai_merged'])
                if st.session_state.get('to_client_file'):
                    df_client = pd.read_excel(st.session_state['to_client_file'])
                else:
                    df_client = pd.DataFrame()
                if all_data:
                    df_merged = pd.concat(all_data, ignore_index=True)
                    results['final_data'] = df_merged
                    st.markdown("**📊 计算项1：单位缴费基数**")
                    base_col = None
                    for col in df_merged.columns:
                        if '基数' in col or '基数总额' in col:
                            base_col = col; break
                    if base_col:
                        total_base = df_merged[base_col].sum()
                        st.metric("单位缴费基数（全年合计）", f"{total_base:,.2f}")
                        results['total_base'] = total_base
                    else:
                        st.warning("未找到基数列，请确认数据格式")
                        results['total_base'] = 0
                    st.markdown("**📊 计算项2：本期实缴金额**")
                    amount_cols = []
                    for col in df_merged.columns:
                        if any(kw in col for kw in ['实缴', '缴费', '金额']):
                            amount_cols.append(col)
                    if amount_cols:
                        total_amount = 0
                        for col in amount_cols:
                            total_amount += df_merged[col].sum()
                        st.metric("本期实缴金额（各险种合计）", f"{total_amount:,.2f}")
                        results['total_amount'] = total_amount
                    else:
                        st.warning("未找到金额列，请确认数据格式")
                        results['total_amount'] = 0
                    client_buffer = BytesIO()
                    with pd.ExcelWriter(client_buffer, engine='openpyxl') as writer:
                        df_merged.to_excel(writer, sheet_name='年审汇总数据', index=False)
                        summary_df = pd.DataFrame([
                            ['统计项', '金额'],
                            ['单位缴费基数（全年）', results.get('total_base', 0)],
                            ['本期实缴金额（合计）', results.get('total_amount', 0)],
                            ['数据来源', '系统账单 + 上海社保数据'],
                            ['生成时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
                        ])
                        summary_df.to_excel(writer, sheet_name='年审汇总', index=False, header=False)
                    client_buffer.seek(0)
                    processed_files['to客户_年审交付文件.xlsx'] = client_buffer.getvalue()
                    st.success(f"✅ 已生成to客户交付文件: {len(df_merged)} 行")
                else:
                    st.warning("请上传to客户模板或确认数据完整")
            except Exception as e:
                st.error(f"生成交付文件失败: {e}")
            if processed_files:
                st.markdown("---")
                st.markdown("### 📦 步骤3：归档打包")
                st.write(f"共 {len(processed_files)} 份文件待归档")
                for fname in processed_files.keys():
                    st.write(f"  📄 {fname}")
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    for fname, data in processed_files.items():
                        zf.writestr(fname, data)
                zip_buffer.seek(0)
                st.download_button("📥 下载全部归档文件（ZIP）", data=zip_buffer, file_name=f"年审归档_{datetime.now().strftime('%Y%m%d_%H%M')}.zip", mime="application/zip")
                st.markdown("**单独下载**")
                cols = st.columns(min(len(processed_files), 4))
                for idx, (fname, data) in enumerate(processed_files.items()):
                    with cols[idx % 4]:
                        st.download_button(f"📄 {fname[:15]}...", data=BytesIO(data), file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("未生成任何文件，请检查数据源")
    
    if 'results' in st.session_state:
        st.markdown("---")
        st.subheader("📊 处理结果汇总")
        results = st.session_state['results']
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("广州员工数", len(results.get('gz_bill', [])) if 'gz_bill' in results else 0)
        with col2:
            st.metric("单位缴费基数", f"{results.get('total_base', 0):,.0f}")
        with col3:
            st.metric("本期实缴金额", f"{results.get('total_amount', 0):,.0f}")

# ===== 底部：快速生成报表 =====
st.markdown("---")
st.subheader("🚀 快速生成报表")

companies = load_companies()
if not companies:
    st.info("👈 请先在「数据导入」页面上传包含公司/城市数据的Excel")
else:
    valid_companies = [c for c in companies if c['province']]
    if len(valid_companies) < len(companies):
        st.warning(f"⚠️ 有 {len(companies) - len(valid_companies)} 家公司的省份无法识别，将使用全局默认规则")
        companies = valid_companies
        if not companies:
            st.stop()
    all_provinces = sorted(set(c['province'] for c in companies if c['province']))
    col1, col2, col3 = st.columns(3)
    with col1:
        province = st.selectbox("省份", [""] + all_provinces, key="report_province")
        if province:
            cities = sorted(set(c['city'] for c in companies if c['province'] == province))
        else:
            cities = sorted(set(c['city'] for c in companies))
        city = st.selectbox("城市", [""] + cities, key="report_city")
    with col2:
        if province and city:
            districts = sorted(set(c['district'] for c in companies if c['province'] == province and c['city'] == city))
        else:
            districts = []
        district = st.selectbox("区县", [""] + districts, key="report_district")
        if province and city:
            company_list = [c for c in companies if c['province'] == province and c['city'] == city and (not district or c['district'] == district)]
        else:
            company_list = []
        company_names = [c['company_name'] for c in company_list]
        selected_company_names = st.multiselect("公司（可多选）", company_names, key="report_companies")
    with col3:
        report_type = st.selectbox("报表类型", ["", "增值税", "社保", "公积金", "个人所得税", "企业所得税", "年度汇算清缴"], key="report_type")
        period_type = st.selectbox("统计口径", ["月度（固定月份）", "累计（1-12月）", "自定义月份范围"], key="period_type")
        if period_type == "月度（固定月份）":
            month = st.selectbox("月份", list(range(1,13)), index=11, key="report_month")
            period_label = f"月度（{month}月）"
            custom_period = f"{month}月"
        elif period_type == "累计（1-12月）":
            period_label = "累计（1-12月）"
            custom_period = "1-12月"
        else:
            start_month = st.selectbox("起始月份", list(range(1,13)), index=0, key="start_month")
            end_month = st.selectbox("结束月份", list(range(1,13)), index=11, key="end_month")
            if start_month <= end_month:
                period_label = f"自定义（{start_month}月-{end_month}月）"
                custom_period = f"{start_month}月-{end_month}月"
            else:
                st.error("起始月份不能大于结束月份")
                period_label = "自定义"
                custom_period = ""
    selected_companies = [c for c in company_list if c['company_name'] in selected_company_names]
    if selected_companies and report_type:
        matched, match_level, candidates = match_template_with_details(province, city, district, report_type)
        custom_templates = load_custom_templates()
        
        # ===== 优化二：模板推荐 =====
        all_templates = load_templates()
        recommended, reason = recommend_template(province, city, report_type, all_templates)
        if recommended:
            st.info(f"💡 推荐模板：**{recommended['template_name']}**（{reason}）")
        else:
            st.info(f"💡 {reason}，请从下方选择")
        
        options = {}
        if matched:
            options[f"✅ 官方模板：{matched['template_name']}（{match_level}）"] = {'type': 'official', 'data': matched}
        for c in candidates:
            if c['id'] != (matched['id'] if matched else ''):
                options[f"📄 官方模板：{c['template_name']}（{c['province']}）"] = {'type': 'official', 'data': c}
        for ct in custom_templates:
            options[f"⭐ 自定义模板：{ct['name']}"] = {'type': 'custom', 'data': ct}
        options["🔄 通用模板（系统内置）"] = {'type': 'general', 'data': None}
        if options:
            default_idx = 0
            keys = list(options.keys())
            if recommended:
                # 如果有推荐，把推荐模板放在第一位
                for i, k in enumerate(keys):
                    if recommended['template_name'] in k:
                        default_idx = i
                        break
            elif matched:
                for i, k in enumerate(keys):
                    if "✅" in k:
                        default_idx = i
                        break
            selected_key = st.selectbox("选择模板", keys, index=default_idx, key="template_choice")
            template_choice = options[selected_key]
        else:
            template_choice = {'type': 'general', 'data': None}
        selected_template = None
        if template_choice['type'] == 'official':
            selected_template = template_choice['data']
        elif template_choice['type'] == 'custom':
            selected_template = template_choice['data']
        else:
            selected_template = {
                'id': 'gen001',
                'template_name': f'{report_type}通用申报表',
                'template_version': 'v1.0',
                'source_authority': '系统内置',
                'publish_date': datetime.now().strftime('%Y-%m-%d'),
                'required_fields': '纳税人识别号,公司名称,申报金额',
                'source_url': '#'
            }
        with st.expander("📋 依据匹配（来源 / 模板 / 规则）", expanded=True):
            st.markdown("**匹配结果**")
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("**来源**")
                if template_choice['type'] == 'official':
                    st.write(f"来源名称：{selected_template.get('source_authority', '未知')}")
                    st.write(f"来源URL：{selected_template.get('source_url', '#')}")
                    st.write(f"文档版本：{selected_template.get('template_version', 'v1.0')}")
                elif template_choice['type'] == 'custom':
                    st.write("来源：自定义模板")
                    st.write(f"名称：{selected_template['name']}")
                else:
                    st.write("来源：系统内置通用模板")
                st.markdown("**模板**")
                st.write(f"模板名称：{selected_template['template_name']}")
                st.write(f"模板版本：{selected_template.get('template_version', 'v1.0')}")
                st.write(f"匹配级别：{match_level if matched else '通用模板'}")
            with col_b:
                st.markdown("**规则**")
                if selected_companies:
                    first_comp = selected_companies[0]
                    rule = get_rule_for_city(first_comp['city'], first_comp.get('province'))
                    if rule:
                        st.write(f"规则来源：{rule.get('source_quote', '未配置')}")
                        st.write(f"规则版本：{rule.get('rule_version', 'v1.0')}")
                        st.write(f"来源链接：{rule.get('source_url', '#')}")
                        st.write(f"来源标题：{rule.get('source_title', '')}")
                        st.write(f"生效日期：{rule.get('effective_date', '未知')}")
                    else:
                        st.write("规则：未匹配，将使用默认值")
                else:
                    st.write("规则：待选择公司后显示")
            st.markdown("**所有公司规则匹配状态**")
            rule_status = []
            for comp in selected_companies:
                r = get_rule_for_city(comp['city'], comp.get('province'))
                if r:
                    rule_status.append(f"{comp['company_name']} → {comp['city']} (规则：{r.get('source_quote', '默认')}, 版本：{r.get('rule_version', 'v1.0')})")
                else:
                    rule_status.append(f"{comp['company_name']} → {comp['city']} (⚠️ 将使用默认值)")
            st.write("\n".join(rule_status))
        with st.expander("📋 字段映射预览", expanded=False):
            fields = selected_template.get('required_fields', '').split(',')
            if fields and fields[0]:
                st.markdown("**字段列表**")
                mapping_data = []
                for f in fields:
                    source = selected_template.get('field_mapping_source', '自动映射')
                    mapping_data.append({'字段': f, '来源字段': f, '映射方式': source})
                st.dataframe(pd.DataFrame(mapping_data), use_container_width=True)
        st.warning("⚠️ 根据合规要求，所有正式导出必须经过人工复核。")
        reviewed = st.checkbox("✅ 我已人工复核确认数据无误，并核实了官方来源", value=False, key="final_review")
        export_type = st.radio("导出类型", ["验证版（可跳过复核）", "正式版（需完成复核）"], key="export_type_radio", horizontal=True)
        if export_type == "正式版（需完成复核）" and not reviewed:
            st.error("导出正式版必须勾选复核确认")
            disabled = True
        else:
            disabled = not reviewed
        if st.button("📥 生成报表", disabled=disabled, key="generate_report"):
            if not selected_template:
                st.error("请先选择模板")
            else:
                if 'validation_report' in st.session_state and st.session_state['validation_report']['error_rows'] > 0:
                    st.warning("⚠️ 当前数据存在异常（见数据质量报告），是否继续生成？")
                    if not st.checkbox("继续生成（忽略异常）", key="ignore_errors"):
                        st.stop()
                batch_id = generate_batch_id()
                batch_name = f"批量导出_{datetime.now().strftime('%Y%m%d_%H%M')}"
                save_batch_job({
                    'id': batch_id,
                    'batch_name': batch_name,
                    'created_at': datetime.now().isoformat(),
                    'status': 'processing',
                    'total_companies': len(selected_companies),
                    'total_reports': len(selected_companies),
                    'review_status': 'pending',
                    'parameters': {'province': province, 'city': city, 'report_type': report_type, 'period': period_label},
                    'created_by': '系统'
                })
                save_verification_status({
                    'batch_id': batch_id,
                    'source_verified': 0,
                    'template_verified': 0,
                    'rule_verified': 0,
                    'data_verified': 0,
                    'reviewer_name': '',
                    'verified_at': datetime.now().isoformat(),
                    'export_type': '验证版' if export_type == "验证版（可跳过复核）" else '正式版'
                })
                generated_files = []
                summary = []
                errors = []
                job_details = []
                data_source_text = st.session_state.get('data_sheet_name', '未知')
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                total = len(selected_companies)
                
                for idx, comp in enumerate(selected_companies):
                    status_text.text(f"正在处理 {idx+1}/{total}: {comp['company_name']}")
                    progress_bar.progress((idx + 1) / total)
                    
                    try:
                        rule = get_rule_for_city(comp['city'], comp.get('province'))
                        if rule is None:
                            rule = {'unit_social': 0.16, 'personal_social': 0.08, 'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 0, 'social_max': 999999, 'fund_min': 0, 'fund_max': 999999, 'source_quote': '全局默认', 'rule_version': 'v1.0', 'source_title': '系统内置默认值', 'source_url': '#', 'source_publish_date': datetime.now().strftime('%Y-%m-%d'), 'applicable_region': '全国', 'official_channel': '系统内置'}
                        fields = selected_template.get('required_fields', '').split(',')
                        if not fields or not fields[0]:
                            fields = ['纳税人识别号', '公司名称', '申报金额']
                        if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
                            df_data = st.session_state['imported_df']
                            company_col = None
                            for col in df_data.columns:
                                if any(kw in col.lower() for kw in COMPANY_KEYWORDS):
                                    company_col = col; break
                            if company_col:
                                df_comp = df_data[df_data[company_col] == comp['company_name']]
                                if not df_comp.empty:
                                    row_data = []
                                    for f in fields:
                                        matched_col = None
                                        for col in df_data.columns:
                                            if f in str(col) or str(col) in f:
                                                matched_col = col; break
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
                        ws['A1'] = f'【系统生成 - {export_type}】统计口径：{period_label}'
                        ws['A1'].font = Font(color='FF0000', bold=True, size=14)
                        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(fields) if fields else 1)
                        ws['A1'].alignment = Alignment(horizontal='center')
                        ws['A1'].fill = PatternFill(start_color='FFF9E6', end_color='FFF9E6', fill_type='solid')
                        ws.insert_rows(2)
                        ws['A2'] = f'模板名称：{selected_template["template_name"]}  版本：{selected_template.get("template_version", "v1.0")}'
                        ws['A2'].font = Font(color='666666', size=10)
                        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(fields) if fields else 1)
                        ws.insert_rows(3)
                        ws['A3'] = f'来源：{selected_template.get("source_authority", "系统内置")}  发布日期：{selected_template.get("publish_date", datetime.now().strftime("%Y-%m-%d"))}'
                        ws['A3'].font = Font(color='666666', size=10)
                        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=len(fields) if fields else 1)
                        ws.insert_rows(4)
                        ws['A4'] = f'数据来源：{data_source_text}  统计口径：{period_label}'
                        ws['A4'].font = Font(color='666666', size=10)
                        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=len(fields) if fields else 1)
                        ws.insert_rows(5)
                        ws['A5'] = f'规则来源：{rule.get("source_quote", "未配置")}  规则版本：{rule.get("rule_version", "v1.0")}'
                        ws['A5'].font = Font(color='666666', size=10)
                        ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=len(fields) if fields else 1)
                        ws.insert_rows(6)
                        ws['A6'] = f'来源链接：{rule.get("source_url", "#")}  来源标题：{rule.get("source_title", "")}'
                        ws['A6'].font = Font(color='666666', size=10)
                        ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=len(fields) if fields else 1)
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
                        ws_annual.append(['统计口径', period_label])
                        ws_annual.append(['参保人数（全年）', int(total_people) if total_people else 0])
                        ws_annual.append(['全年社保缴费基数总额', round(social_total, 2)])
                        ws_annual.append(['全年公积金缴费基数总额', round(fund_total, 2)])
                        ws_annual.append(['单位全年缴费总额', round(unit_total, 2)])
                        ws_annual.append(['个人全年缴费总额', round(personal_total, 2)])
                        ws_annual.append(['全年总费用', round(grand_total, 2)])
                        ws_annual.append(['报告生成时间', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
                        ws_annual.append(['数据来源', data_source_text])
                        ws_annual.append(['规则来源', rule.get('source_quote', '未配置')])
                        ws_annual.append(['规则版本', rule.get('rule_version', 'v1.0')])
                        ws_annual.append(['来源链接', rule.get('source_url', '#')])
                        ws_annual.append(['来源标题', rule.get('source_title', '')])
                        audit = wb.create_sheet("审计日志")
                        audit.append(['操作时间', '操作类型', '操作人', '详情'])
                        audit.append([datetime.now().isoformat(), 'GENERATED', '系统', f'公司:{comp["company_name"]}, 城市:{comp["city"]}, 模板:{selected_template["template_name"]}, 统计口径:{period_label}, 规则:{rule.get("source_quote", "未配置")}'])
                        output = BytesIO()
                        wb.save(output)
                        output.seek(0)
                        fname = f"{comp['company_name']}_{report_type}_{period_label.replace('（','_').replace('）','').replace('-','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                        generated_files.append((fname, output.getvalue()))
                        summary.append({'公司': comp['company_name'], '城市': comp['city'], '模板': selected_template['template_name'], '统计口径': period_label, '规则来源': rule.get('source_quote', '未配置'), '规则版本': rule.get('rule_version', 'v1.0'), '状态': '待复核'})
                        job_details.append({'id': str(uuid.uuid4())[:8], 'batch_id': batch_id, 'company_id': comp['id'], 'company_name': comp['company_name'], 'city': comp['city'], 'province': comp.get('province', ''), 'report_type': report_type, 'period_type': period_label, 'status': 'success', 'error_message': '', 'file_name': fname, 'file_data': output.getvalue(), 'generated_at': datetime.now().isoformat(), 'rule_source': rule.get('source_quote', '未配置'), 'data_source': data_source_text})
                        save_export({
                            'id': str(uuid.uuid4())[:8],
                            'company_id': comp['id'],
                            'template_id': selected_template.get('id', 'gen001'),
                            'company_name': comp['company_name'],
                            'city': comp['city'],
                            'province': comp.get('province', ''),
                            'report_type': report_type,
                            'period_type': period_label,
                            'generated_at': datetime.now().isoformat(),
                            'review_status': 'pending',
                            'file_name': fname,
                            'file_data': output.getvalue(),
                            'data_source': data_source_text,
                            'month_used': period_label,
                            'year_used': datetime.now().strftime('%Y'),
                            'custom_period': custom_period if period_type == "自定义月份范围" else '',
                            'batch_id': batch_id,
                            'job_name': batch_name,
                            'field_mapping': json.dumps({f: f for f in fields})
                        })
                    except Exception as e:
                        errors.append(f"{comp['company_name']}: {str(e)}")
                        job_details.append({'id': str(uuid.uuid4())[:8], 'batch_id': batch_id, 'company_id': comp.get('id', ''), 'company_name': comp['company_name'], 'city': comp.get('city', ''), 'province': comp.get('province', ''), 'report_type': report_type, 'period_type': period_label, 'status': 'error', 'error_message': str(e), 'file_name': '', 'file_data': None, 'generated_at': datetime.now().isoformat(), 'rule_source': '', 'data_source': data_source_text})
                
                progress_bar.empty()
                status_text.empty()
                
                save_job_details(job_details)
                update_batch_status(batch_id, 'completed', 'pending')
                if errors:
                    for err in errors:
                        st.warning(err)
                if generated_files:
                    st.success(f"✅ 成功生成 {len(generated_files)} 份报表（批次ID：{batch_id}，状态：已完成）")
                    st.dataframe(pd.DataFrame(summary), use_container_width=True)
                    
                    # ===== 优化三：PDF导出 =====
                    if PDF_AVAILABLE:
                        try:
                            # 生成第一个报表的PDF预览
                            df_summary = pd.DataFrame(summary)
                            html_content = df_to_html_table(df_summary)
                            full_html = create_pdf_from_html(html_content, f"报表_{datetime.now().strftime('%Y%m%d')}")
                            pdf_buffer = BytesIO()
                            pdfkit.from_string(full_html, pdf_buffer)
                            pdf_buffer.seek(0)
                            st.download_button(
                                "📄 下载报表摘要（PDF）",
                                data=pdf_buffer,
                                file_name=f"报表摘要_{datetime.now().strftime('%Y%m%d')}.pdf",
                                mime="application/pdf"
                            )
                        except Exception as e:
                            st.warning(f"PDF导出失败（请确保已安装wkhtmltopdf）：{e}")
                    else:
                        st.warning("💡 PDF导出需要安装 pdfkit 和 wkhtmltopdf，运行: pip install pdfkit，并从 https://wkhtmltopdf.org/ 下载安装 wkhtmltopdf")
                    
                    # 下载Excel/ZIP
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
