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
        rule_version TEXT, effective_date TEXT
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
    conn.commit()
    conn.close()

init_db()

# ========== 数据库迁移 ==========
def migrate_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    tables_to_check = [
        ('templates', ['field_mapping_source']),
        ('rules', ['rule_version', 'effective_date']),
        ('export_history', ['batch_id', 'job_name', 'field_mapping']),
        ('source_registry', ['document_name', 'document_version', 'publish_year']),
    ]
    for table, cols in tables_to_check:
        c.execute(f"PRAGMA table_info({table})")
        existing_cols = [col[1] for col in c.fetchall()]
        for col in cols:
            if col not in existing_cols:
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
    conn.commit()
    conn.close()

migrate_db()

# ========== 数据操作 ==========
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
             rule_version, effective_date)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (r['id'], r['city'], r.get('province',''), r['unit_social'], r['personal_social'],
             r['unit_fund'], r['personal_fund'], r.get('social_min',0), r.get('social_max',999999),
             r.get('fund_min',0), r.get('fund_max',999999), r.get('source_quote',''),
             r.get('is_default',0), r.get('rule_version','v1.0'), r.get('effective_date','')))
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

# ========== 全国默认规则 ==========
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
    {'city': '东莞', 'province': '广东', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 4588, 'social_max': 22941,
     'fund_min': 1900, 'fund_max': 25431, 'source_quote': '东人社发〔2024〕6号'},
    {'city': '佛山', 'province': '广东', 'unit_social': 0.15, 'personal_social': 0.08,
     'unit_fund': 0.10, 'personal_fund': 0.10, 'social_min': 4588, 'social_max': 22941,
     'fund_min': 1900, 'fund_max': 25431, 'source_quote': '佛人社发〔2024〕5号'},
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
    {'city': '济南', 'province': '山东', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 3746, 'social_max': 18726,
     'fund_min': 2010, 'fund_max': 23496, 'source_quote': '济人社发〔2024〕5号'},
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
    {'city': '襄阳', 'province': '湖北', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 0, 'social_max': 999999,
     'fund_min': 0, 'fund_max': 999999, 'source_quote': '系统默认（建议核实）'},
    {'city': '绵阳', 'province': '四川', 'unit_social': 0.16, 'personal_social': 0.08,
     'unit_fund': 0.12, 'personal_fund': 0.12, 'social_min': 0, 'social_max': 999999,
     'fund_min': 0, 'fund_max': 999999, 'source_quote': '系统默认（建议核实）'},
]

# ========== 确保默认规则 ==========
def ensure_default_rules():
    existing_rules = load_rules()
    existing_cities = {normalize_name(r['city']) for r in existing_rules}
    added = 0
    for dr in PROVINCE_DEFAULT_RULES:
        norm_city = normalize_name(dr['city'])
        if norm_city not in existing_cities:
            new_rule = {
                'id': str(uuid.uuid4())[:8],
                'city': dr['city'],
                'province': dr.get('province', dr['city']),
                'unit_social': dr['unit_social'],
                'personal_social': dr['personal_social'],
                'unit_fund': dr['unit_fund'],
                'personal_fund': dr['personal_fund'],
                'social_min': dr.get('social_min', 0),
                'social_max': dr.get('social_max', 999999),
                'fund_min': dr.get('fund_min', 0),
                'fund_max': dr.get('fund_max', 999999),
                'source_quote': dr.get('source_quote', '省份默认'),
                'is_default': 1,
                'rule_version': 'v1.0',
                'effective_date': datetime.now().strftime('%Y-%m-%d')
            }
            existing_rules.append(new_rule)
            added += 1
    if added > 0:
        save_rules(existing_rules)
    return added

def fix_companies_province():
    companies = load_companies()
    if not companies:
        return 0, 0
    city_province = {}
    for dr in PROVINCE_DEFAULT_RULES:
        city_province[normalize_name(dr['city'])] = dr['province']
    for r in load_rules():
        city_province[normalize_name(r['city'])] = r['province']
    
    fixed = 0
    updated_companies = []
    for comp in companies:
        original_province = comp['province']
        city = comp['city']
        norm_city = normalize_name(city)
        correct_province = city_province.get(norm_city)
        if correct_province and correct_province != original_province:
            comp['province'] = correct_province
            fixed += 1
        elif not correct_province and original_province != '':
            comp['province'] = ''
            fixed += 1
        updated_companies.append(comp)
    if fixed > 0:
        save_companies(updated_companies)
    return fixed, len(companies)

# ========== 标准化函数 ==========
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

# ========== 自动检测表头并读取Sheet（修复版） ==========
def auto_load_sheet_with_header_detection(file, sheet_name):
    xls = pd.ExcelFile(file)
    df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None)
    
    header_row = None
    for i, row in df_raw.iterrows():
        row_text = ' '.join([str(v) for v in row.values if pd.notna(v)])
        # 更严格的检测：必须同时包含"城市"和"公司"或"省份"
        if ('城市' in row_text and '公司' in row_text) or ('省份' in row_text and '城市' in row_text):
            header_row = i
            break
    
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

# ========== 判断是否为有效中文名称 ==========
def is_valid_chinese_name(text):
    """判断文本是否为有效的中文名称（至少包含2个中文字符）"""
    if not text or not isinstance(text, str):
        return False
    # 至少包含2个中文字符
    chinese_chars = re.findall(r'[\u4e00-\u9fa5]', text)
    return len(chinese_chars) >= 2

def is_valid_city_name(text):
    """判断是否为有效的城市名（至少包含2个中文字符）"""
    if not text or not isinstance(text, str):
        return False
    chinese_chars = re.findall(r'[\u4e00-\u9fa5]', text)
    return len(chinese_chars) >= 2

# ========== 解析Excel（过滤异常数据） ==========
def parse_uploaded_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    all_companies = []
    unmapped_cities = set()
    filtered_values = []  # 记录被过滤的异常值
    
    city_province_map = {}
    for dr in PROVINCE_DEFAULT_RULES:
        key = normalize_name(dr['city'])
        city_province_map[key] = dr['province']
    for r in load_rules():
        key = normalize_name(r['city'])
        city_province_map[key] = r['province']
    
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
                        
                        # ---- 过滤异常数据 ----
                        # 1. 城市必须是有效中文名称
                        if city and not is_valid_city_name(city):
                            filtered_values.append(f"城市无效: {city}")
                            continue
                        
                        # 2. 公司名称必须是有效中文名称（至少包含2个中文字符）
                        if company and not is_valid_chinese_name(company):
                            filtered_values.append(f"公司名称无效: {company}")
                            continue
                        
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
    
    # 如果有过滤掉的异常值，在session中记录以便显示
    if filtered_values:
        st.session_state['filtered_values'] = filtered_values
    
    unique = []
    seen = set()
    for c in all_companies:
        key = (c['company_name'], c['city'])
        if key not in seen:
            seen.add(key)
            unique.append(c)
    return unique, unmapped_cities, sheets

def parse_multiple_files(files):
    all_companies = []
    all_sheets = []
    unmapped_cities = set()
    for file in files:
        companies, unmapped, sheets = parse_uploaded_excel(file)
        all_companies.extend(companies)
        all_sheets.extend(sheets)
        unmapped_cities.update(unmapped)
    unique = []
    seen = set()
    for c in all_companies:
        key = (c['company_name'], c['city'])
        if key not in seen:
            seen.add(key)
            unique.append(c)
    return unique, unmapped_cities, all_sheets

# ========== 数据校验函数 ==========
def validate_data(df, rules):
    if df is None or df.empty:
        return {'total_rows': 0, 'error_rows': 0, 'details': {}, 'summary': {}, 'error_rows_detail': {}}
    
    errors = {}
    city_col = None
    for col in df.columns:
        if '城市' in col:
            city_col = col
            break
    if city_col is None:
        errors['城市列缺失'] = ['未找到城市列']
        return {'total_rows': len(df), 'error_rows': 0, 'details': errors, 'summary': {'城市列缺失': 1}, 'error_rows_detail': {}}
    
    city_rule_map = {normalize_name(r['city']): r for r in rules}
    error_rows = {}
    
    numeric_cols = []
    for col in df.columns:
        col_lower = col.lower()
        if ('基数' in col_lower or '金额' in col_lower or '费用' in col_lower or '比例' in col_lower):
            exclude_words = ['校验', '状态', '类型', '说明', '备注', '合规', '是否', '判断', '结果']
            if not any(w in col_lower for w in exclude_words):
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
            if '公司' in col or '分公司' in col:
                company_col = col
                break
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

# ========== 核心：获取规则（含省份兜底和全局默认） ==========
def get_rule_for_city(city, province=None):
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
                fallback = r.copy()
                fallback['source_quote'] = f"省份默认（{province}）"
                return fallback
    return {
        'unit_social': 0.16,
        'personal_social': 0.08,
        'unit_fund': 0.12,
        'personal_fund': 0.12,
        'social_min': 0,
        'social_max': 999999,
        'fund_min': 0,
        'fund_max': 999999,
        'source_quote': '全局默认'
    }

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

# ========== Streamlit 页面 ==========
st.set_page_config(page_title="智能报表系统 - 企业版", layout="wide")
st.title("📋 智能报表系统（企业版）")
st.markdown("**工作台 · 数据识别 · 依据库 · 复核导出 · 作业记录**")

ensure_default_rules()
fixed, total = fix_companies_province()
if fixed > 0:
    st.success(f"✅ 自动修复了 {fixed}/{total} 家公司的省份数据")

# 显示被过滤的异常值提示
if 'filtered_values' in st.session_state and st.session_state['filtered_values']:
    with st.expander("⚠️ 已过滤的异常数据"):
        for val in st.session_state['filtered_values']:
            st.write(f"- {val}")
    st.session_state['filtered_values'] = []

# ===== 侧边栏导航 =====
st.sidebar.title("📌 导航")
page = st.sidebar.radio("选择功能", [
    "📊 工作台",
    "📤 数据导入",
    "📚 依据库管理",
    "⚙️ 规则管理",
    "📄 自定义模板",
    "📋 导出历史与复核",
    "💾 备份与恢复"
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
            rules = load_rules()
            st.session_state['validation_report'] = validate_data(df, rules)
            st.sidebar.success(f"✅ 已加载: {selected_sheet} (表头行: {header_row+1})")
            st.rerun()
        else:
            st.sidebar.info(f"当前: {selected_sheet}")
    except Exception as e:
        st.sidebar.error(f"读取Sheet失败: {e}")

# ===== 工作台 =====
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

# ===== 数据导入 =====
elif page == "📤 数据导入":
    st.subheader("📤 数据导入（支持多文件）")
    
    import_mode = st.radio("导入模式", ["智能导入（自动识别结构）", "普通导入（手动选择列，开发中）"], index=0, horizontal=True)
    st.caption("智能导入将自动识别城市、公司等列；普通导入可自定义列映射（开发中，当前与智能导入相同）")
    
    uploaded_files = st.file_uploader(
        "选择Excel文件（支持多个 .xlsx）", 
        type=["xlsx"], 
        accept_multiple_files=True
    )
    
    if uploaded_files:
        with st.spinner("正在解析Excel..."):
            companies, unmapped, all_sheets = parse_multiple_files(uploaded_files)
            # 过滤掉异常公司（名称无效的已在解析时过滤）
            valid_companies = companies
            if unmapped:
                st.warning(f"⚠️ 以下城市未在规则库中找到：{', '.join(unmapped)}，将使用全局默认规则，请在规则管理中补充以获得更准确的数据。")
            if valid_companies:
                save_companies(valid_companies)
                st.success(f"成功提取 {len(valid_companies)} 家公司，来自 {len(uploaded_files)} 个文件")
                st.session_state['uploaded_files'] = uploaded_files
                st.session_state['all_sheets'] = all_sheets
                if uploaded_files:
                    first_file = uploaded_files[0]
                    xls = pd.ExcelFile(first_file)
                    sheets = xls.sheet_names
                    default_sheet = None
                    for kw in ['明细', '月度', '数据', '工资', '社保', '员工', '月报']:
                        for s in sheets:
                            if kw in s:
                                default_sheet = s
                                break
                        if default_sheet:
                            break
                    if not default_sheet and sheets:
                        default_sheet = sheets[0]
                    if default_sheet:
                        df, header_row = auto_load_sheet_with_header_detection(first_file, default_sheet)
                        st.session_state['imported_df'] = df
                        st.session_state['data_sheet_name'] = default_sheet
                        st.session_state['data_header_row'] = header_row
                        rules = load_rules()
                        st.session_state['validation_report'] = validate_data(df, rules)
                        st.success(f"已自动加载 Sheet「{default_sheet}」（表头行: {header_row+1}），共 {len(df)} 行数据")
            else:
                st.warning("未识别到有效公司数据，请检查数据格式")
    
    if 'imported_df' in st.session_state and st.session_state['imported_df'] is not None:
        st.subheader("📊 数据预览")
        df = st.session_state['imported_df']
        st.dataframe(df.head(10), use_container_width=True)
        st.caption(f"当前Sheet: {st.session_state.get('data_sheet_name', '未知')}，共 {len(df)} 行，表头行: {st.session_state.get('data_header_row', 0)+1}")
        
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

# ===== 依据库管理 =====
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
            cols = ['city', 'province', 'unit_social', 'personal_social', 'unit_fund', 'personal_fund', 'source_quote']
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
        sources = load_source_registry()
        if sources:
            df = pd.DataFrame(sources)
            cols = ['authority_name', 'province', 'city', 'source_url', 'document_name', 'document_version']
            st.dataframe(df[cols], use_container_width=True)
        else:
            st.info("暂无来源，可点击下方初始化样本依据")
        
        if st.button("加载样本依据"):
            sample_sources = [
                {'id': 'src_gz_social_demo', 'authority_type': 'social_security', 
                 'province': '广东', 'city': '广州', 'district': '', 
                 'authority_name': '广州市人力资源和社会保障局',
                 'official_site_name': '广州市人力资源和社会保障局官网',
                 'source_url': 'https://rsj.gz.gov.cn/',
                 'document_name': '广州市社保年审公告', 'document_version': '2024',
                 'publish_year': '2024'},
                {'id': 'src_sh_social_sample', 'authority_type': 'social_security',
                 'province': '上海', 'city': '上海', 'district': '',
                 'authority_name': '上海市人力资源和社会保障局',
                 'official_site_name': '上海社保官方系统',
                 'source_url': 'https://rsj.sh.gov.cn/',
                 'document_name': '上海市社保基数调整通知', 'document_version': '2024',
                 'publish_year': '2024'},
                {'id': 'src_km_social_sample', 'authority_type': 'social_security',
                 'province': '云南', 'city': '昆明', 'district': '',
                 'authority_name': '昆明市人力资源和社会保障局',
                 'official_site_name': '昆明社保官方系统',
                 'source_url': '', 'document_name': '昆明市社保年检通知',
                 'document_version': '2024', 'publish_year': '2024'}
            ]
            save_source_registry(sample_sources)
            st.success("已加载样本依据！")
            st.rerun()

# ===== 规则管理 =====
elif page == "⚙️ 规则管理":
    st.subheader("⚙️ 规则管理（社保/公积金）")
    rules = load_rules()
    st.write(f"**当前规则数量：{len(rules)} 个城市**")
    
    if rules:
        df_rules = pd.DataFrame(rules)
        st.dataframe(df_rules[['city', 'province', 'unit_social', 'personal_social', 'unit_fund', 'personal_fund', 'social_min', 'social_max', 'source_quote']], use_container_width=True)
        
        cities = sorted(set(r['city'] for r in rules))
        selected_city = st.selectbox("选择城市进行编辑", [""] + cities)
        if selected_city:
            rule = next((r for r in rules if r['city'] == selected_city), None)
            if rule:
                with st.form(key=f"edit_rule_{selected_city}"):
                    st.write(f"编辑 **{selected_city}** 的规则")
                    new_unit_social = st.number_input("单位社保比例", value=float(rule['unit_social']), step=0.001, format="%.3f")
                    new_personal_social = st.number_input("个人社保比例", value=float(rule['personal_social']), step=0.001, format="%.3f")
                    new_unit_fund = st.number_input("单位公积金比例", value=float(rule['unit_fund']), step=0.001, format="%.3f")
                    new_personal_fund = st.number_input("个人公积金比例", value=float(rule['personal_fund']), step=0.001, format="%.3f")
                    new_social_min = st.number_input("社保基数下限", value=float(rule.get('social_min', 0)), step=100.0)
                    new_social_max = st.number_input("社保基数上限", value=float(rule.get('social_max', 999999)), step=100.0)
                    new_fund_min = st.number_input("公积金基数下限", value=float(rule.get('fund_min', 0)), step=100.0)
                    new_fund_max = st.number_input("公积金基数上限", value=float(rule.get('fund_max', 999999)), step=100.0)
                    new_source = st.text_input("来源文号", value=rule.get('source_quote', ''))
                    new_rule_version = st.text_input("规则版本", value=rule.get('rule_version', 'v1.0'))
                    new_effective_date = st.text_input("生效日期", value=rule.get('effective_date', datetime.now().strftime('%Y-%m-%d')))
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
                                    'source_quote': new_source,
                                    'rule_version': new_rule_version,
                                    'effective_date': new_effective_date
                                })
                            updated_rules.append(r)
                        save_rules(updated_rules)
                        st.success("规则已更新！")
                        st.rerun()
        
        with st.expander("➕ 新增城市规则"):
            with st.form(key="add_rule"):
                new_city = st.text_input("城市名称")
                new_province = st.text_input("所属省份")
                new_unit_social = st.number_input("单位社保比例", value=0.16, step=0.001, format="%.3f")
                new_personal_social = st.number_input("个人社保比例", value=0.08, step=0.001, format="%.3f")
                new_unit_fund = st.number_input("单位公积金比例", value=0.12, step=0.001, format="%.3f")
                new_personal_fund = st.number_input("个人公积金比例", value=0.12, step=0.001, format="%.3f")
                new_social_min = st.number_input("社保基数下限", value=0.0, step=100.0)
                new_social_max = st.number_input("社保基数上限", value=999999.0, step=100.0)
                new_fund_min = st.number_input("公积金基数下限", value=0.0, step=100.0)
                new_fund_max = st.number_input("公积金基数上限", value=999999.0, step=100.0)
                new_source = st.text_input("来源文号")
                new_rule_version = st.text_input("规则版本", value="v1.0")
                new_effective_date = st.text_input("生效日期", value=datetime.now().strftime('%Y-%m-%d'))
                submitted = st.form_submit_button("添加")
                if submitted and new_city:
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
                        'source_quote': new_source,
                        'is_default': 0,
                        'rule_version': new_rule_version,
                        'effective_date': new_effective_date
                    }
                    rules.append(new_rule)
                    save_rules(rules)
                    st.success(f"已添加 {new_city} 的规则！")
                    st.rerun()
        
        if st.button("🔄 重置所有规则为系统默认值"):
            if st.checkbox("确认重置？此操作将覆盖所有自定义规则"):
                default_rules = []
                for r in PROVINCE_DEFAULT_RULES:
                    default_rules.append({
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
                        'is_default': 1,
                        'rule_version': 'v1.0',
                        'effective_date': datetime.now().strftime('%Y-%m-%d')
                    })
                save_rules(default_rules)
                st.success("已重置为系统默认规则！")
                st.rerun()
    else:
        st.info("暂无规则，请先重置或添加")

# ===== 自定义模板 =====
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
            field_mapping_source = st.text_input("字段映射来源说明（可选）", placeholder="如：根据XX文件字段对应表")
            
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
                field_options = [''] + ['纳税人识别号', '公司名称', '销售额', '进项税额', '应纳税额', '单位名称', 
                                        '社保登记号', '基数', '单位金额', '个人金额', '单位比例', '个人比例', 
                                        '公积金账号', '收入额', '专项扣除', '营业收入', '营业成本', '应纳税所得额', 
                                        '全年收入', '全年成本', '已预缴税额', '应补退税额', '申报金额']
                for h in headers:
                    col = st.selectbox(
                        f"字段 '{h}' 映射到",
                        field_options,
                        key=f"map_{h}_{uploaded_template.name}"
                    )
                    if col:
                        mapping[h] = col
            
            if st.button("保存自定义模板"):
                if not template_name:
                    st.error("请填写模板名称")
                else:
                    template_data = {
                        'id': str(uuid.uuid4())[:8],
                        'name': template_name,
                        'file_data': uploaded_template.getvalue(),
                        'field_mapping': mapping,
                        'sheet_name': sheet_name,
                        'created_at': datetime.now().isoformat()
                    }
                    save_custom_template(template_data)
                    st.success(f"模板 '{template_name}' 已保存！")
                    st.rerun()
        except Exception as e:
            st.error(f"处理模板失败：{e}")

# ===== 导出历史与复核 =====
elif page == "📋 导出历史与复核":
    st.subheader("📋 导出历史与复核")
    
    tab1, tab2 = st.tabs(["📋 导出历史", "✅ 复核处理"])
    
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

# ===== 备份与恢复 =====
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
            if matched:
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
        
        # ---- 依据匹配详情 ----
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
                    rule_status.append(f"{comp['company_name']} → {comp['city']} (规则：{r.get('source_quote', '默认')})")
                else:
                    rule_status.append(f"{comp['company_name']} → {comp['city']} (⚠️ 将使用默认值)")
            st.write("\n".join(rule_status))
        
        # ---- 字段映射预览 ----
        with st.expander("📋 字段映射预览", expanded=False):
            fields = selected_template.get('required_fields', '').split(',')
            if fields and fields[0]:
                st.markdown("**字段列表**")
                mapping_data = []
                for f in fields:
                    source = selected_template.get('field_mapping_source', '自动映射')
                    mapping_data.append({'字段': f, '来源字段': f, '映射方式': source})
                st.dataframe(pd.DataFrame(mapping_data), use_container_width=True)
        
        reviewed = st.checkbox("✅ 我已人工复核确认数据无误", value=False, key="final_review")
        
        if st.button("📥 生成待复核版Excel", disabled=not reviewed, key="generate_report"):
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
                
                generated_files = []
                summary = []
                errors = []
                job_details = []
                data_source_text = st.session_state.get('data_sheet_name', '未知')
                
                for comp in selected_companies:
                    try:
                        rule = get_rule_for_city(comp['city'], comp.get('province'))
                        if rule is None:
                            rule = {
                                'unit_social': 0.16,
                                'personal_social': 0.08,
                                'unit_fund': 0.12,
                                'personal_fund': 0.12,
                                'social_min': 0,
                                'social_max': 999999,
                                'fund_min': 0,
                                'fund_max': 999999,
                                'source_quote': '全局默认'
                            }
                        
                        fields = selected_template.get('required_fields', '').split(',')
                        if not fields or not fields[0]:
                            fields = ['纳税人识别号', '公司名称', '申报金额']
                        
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
                        ws['A1'] = f'【系统生成 - 待复核版】统计口径：{period_label}'
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
                        ws['A4'] = f'数据来源：{data_source_text}  统计口径：{period_label}  规则来源：{rule.get("source_quote", "未配置")}'
                        ws['A4'].font = Font(color='666666', size=10)
                        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=len(fields) if fields else 1)
                        
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
                        
                        audit = wb.create_sheet("审计日志")
                        audit.append(['操作时间', '操作类型', '操作人', '详情'])
                        audit.append([datetime.now().isoformat(), 'GENERATED', '系统', f'公司:{comp["company_name"]}, 城市:{comp["city"]}, 模板:{selected_template["template_name"]}, 统计口径:{period_label}, 规则:{rule.get("source_quote", "未配置")}'])
                        
                        output = BytesIO()
                        wb.save(output)
                        output.seek(0)
                        
                        fname = f"{comp['company_name']}_{report_type}_{period_label.replace('（','_').replace('）','').replace('-','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                        generated_files.append((fname, output.getvalue()))
                        summary.append({
                            '公司': comp['company_name'],
                            '城市': comp['city'],
                            '模板': selected_template['template_name'],
                            '统计口径': period_label,
                            '规则来源': rule.get('source_quote', '未配置'),
                            '状态': '待复核'
                        })
                        
                        job_details.append({
                            'id': str(uuid.uuid4())[:8],
                            'batch_id': batch_id,
                            'company_id': comp['id'],
                            'company_name': comp['company_name'],
                            'city': comp['city'],
                            'province': comp.get('province', ''),
                            'report_type': report_type,
                            'period_type': period_label,
                            'status': 'success',
                            'error_message': '',
                            'file_name': fname,
                            'file_data': output.getvalue(),
                            'generated_at': datetime.now().isoformat(),
                            'rule_source': rule.get('source_quote', '未配置'),
                            'data_source': data_source_text
                        })
                        
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
                        job_details.append({
                            'id': str(uuid.uuid4())[:8],
                            'batch_id': batch_id,
                            'company_id': comp.get('id', ''),
                            'company_name': comp['company_name'],
                            'city': comp.get('city', ''),
                            'province': comp.get('province', ''),
                            'report_type': report_type,
                            'period_type': period_label,
                            'status': 'error',
                            'error_message': str(e),
                            'file_name': '',
                            'file_data': None,
                            'generated_at': datetime.now().isoformat(),
                            'rule_source': '',
                            'data_source': data_source_text
                        })
                
                save_job_details(job_details)
                update_batch_status(batch_id, 'completed', 'pending')
                
                if errors:
                    for err in errors:
                        st.warning(err)
                if generated_files:
                    st.success(f"✅ 成功生成 {len(generated_files)} 份报表（批次ID：{batch_id}，状态：已完成）")
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
