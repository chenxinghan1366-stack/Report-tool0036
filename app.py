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
        is_custom BOOLEAN DEFAULT 0
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS custom_templates (
        id TEXT PRIMARY KEY, name TEXT, file_data BLOB, field_mapping TEXT,
        sheet_name TEXT, created_at TEXT, updated_at TEXT
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
        data_source TEXT, month_used TEXT, year_used TEXT, custom_period TEXT
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
    for col in ['data_source', 'month_used', 'year_used', 'custom_period']:
        if col not in columns_export:
            c.execute(f"ALTER TABLE export_history ADD COLUMN {col} TEXT")
    
    c.execute("PRAGMA table_info(templates)")
    columns_templates = [col[1] for col in c.fetchall()]
    for col in ['file_hash', 'file_type', 'is_custom']:
        if col not in columns_templates:
            c.execute(f"ALTER TABLE templates ADD COLUMN {col} { 'BOOLEAN DEFAULT 0' if col == 'is_custom' else 'TEXT' }")
    
    c.execute("PRAGMA table_info(custom_templates)")
    columns_custom = [col[1] for col in c.fetchall()]
    if not columns_custom:
        c.execute('''CREATE TABLE IF NOT EXISTS custom_templates (
            id TEXT PRIMARY KEY, name TEXT, file_data BLOB, field_mapping TEXT,
            sheet_name TEXT, created_at TEXT, updated_at TEXT
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
            required_fields=?, status=?, file_type=?, is_custom=?
            WHERE id=?''',
            (template['template_name'], template.get('template_version','v1.0'),
             template.get('source_authority',''), template.get('publish_date',''),
             template.get('required_fields',''), template.get('status','active'),
             template.get('file_type',''), template.get('is_custom',0), existing[0]))
    else:
        c.execute('''INSERT INTO templates 
            (id, province, city, district, report_type, template_name, template_version,
             source_url, source_authority, publish_date, required_fields, status, file_hash, file_type, is_custom)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (template['id'], template.get('province',''), template.get('city',''),
             template.get('district',''), template.get('report_type',''),
             template['template_name'], template.get('template_version','v1.0'),
             template.get('source_url',''), template.get('source_authority',''),
             template.get('publish_date',''), template.get('required_fields',''),
             template.get('status','active'), template.get('file_hash',''),
             template.get('file_type',''), template.get('is_custom',0)))
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
             social_min, social_max, fund_min, fund_max, source_quote, is_default)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (r['id'], r['city'], r.get('province',''), r['unit_social'], r['personal_social'],
             r['unit_fund'], r['personal_fund'], r.get('social_min',0), r.get('social_max',999999),
             r.get('fund_min',0), r.get('fund_max',999999), r.get('source_quote',''),
             r.get('is_default',0)))
    conn.commit()
    conn.close()
    backup_database()

def save_export(record):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT OR REPLACE INTO export_history 
        (id, company_id, template_id, company_name, city, province, report_type, period_type,
         generated_at, review_status, reviewer, reviewed_at, file_name, file_data,
         data_source, month_used, year_used, custom_period)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (record['id'], record.get('company_id',''), record.get('template_id',''),
         record['company_name'], record.get('city',''), record.get('province',''),
         record.get('report_type',''), record.get('period_type',''), record['generated_at'],
         record.get('review_status','pending'), record.get('reviewer',''), record.get('reviewed_at',''),
         record.get('file_name',''), record.get('file_data', None),
         record.get('data_source',''), record.get('month_used',''), record.get('year_used',''),
         record.get('custom_period','')))
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

# ========== 初始化/补全规则 ==========
def ensure_default_rules():
    """确保所有默认城市规则都存在，缺失则插入，已存在则保留原有值（不覆盖）"""
    existing_rules = load_rules()
    existing_cities = {normalize_name(r['city']) for r in existing_rules}
    
    added = 0
    for dr in PROVINCE_DEFAULT_RULES:
        norm_city = normalize_name(dr['city'])
        if norm_city not in existing_cities:
            new_rule = {
                'id': str(uuid.uuid4())[:8],
                'city': dr['city'],
                'province': dr['province'],
                'unit_social': dr['unit_social'],
                'personal_social': dr['personal_social'],
                'unit_fund': dr['unit_fund'],
                'personal_fund': dr['personal_fund'],
                'social_min': dr.get('social_min', 0),
                'social_max': dr.get('social_max', 999999),
                'fund_min': dr.get('fund_min', 0),
                'fund_max': dr.get('fund_max', 999999),
                'source_quote': dr.get('source_quote', '省份默认'),
                'is_default': 1
            }
            existing_rules.append(new_rule)
            added += 1
    
    if added > 0:
        save_rules(existing_rules)
        print(f"[系统] 已补全 {added} 个缺失的城市规则")

# 执行补全
ensure_default_rules()

# ========== 标准化匹配 ==========
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

# ========== 自定义模板 ==========
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

# ========== 【核心修复】解析Excel：使用默认规则映射省份 ==========
def parse_uploaded_excel(file):
    xls = pd.ExcelFile(file)
    sheets = xls.sheet_names
    all_companies = []
    
    # 构建城市→省份映射（直接使用默认规则，不依赖数据库）
    city_province_map = {}
    for dr in PROVINCE_DEFAULT_RULES:
        city_province_map[normalize_name(dr['city'])] = dr['province']
    
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
                            norm_city = normalize_name(city)
                            province = city_province_map.get(norm_city, city)  # 找不到就用城市名本身（直辖市也如此）
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

# ========== 获取规则（去后缀匹配） ==========
def get_rule_for_city(city):
    if not city:
        return None
    rules = load_rules()
    norm_city = normalize_name(city)
    for r in rules:
        if normalize_name(r['city']) == norm_city:
            return r
    # 尝试用省份匹配（直辖市）
    for r in rules:
        if normalize_name(r.get('province', '')) == norm_city:
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
st.set_page_config(page_title="智能报表系统", layout="wide")
st.title("📋 智能报表系统（含自定义模板与备份）")
st.markdown("**上传Excel → 选择模板 → 选择统计口径 → 生成待复核版Excel**")

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
    
    # ===== 备份管理 =====
    with st.sidebar.expander("💾 备份与恢复"):
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
    
    # ===== 公司列表（增加清空按钮） =====
    with st.sidebar.expander("🏢 当前公司列表"):
        companies = load_companies()
        if companies:
            st.dataframe(pd.DataFrame(companies))
            st.caption(f"共 {len(companies)} 家公司")
            if st.button("🗑️ 清空所有公司数据", key="clear_companies"):
                if st.checkbox("确认清空？此操作不可恢复", key="confirm_clear"):
                    save_companies([])
                    st.success("已清空所有公司数据")
                    st.rerun()
        else:
            st.info("暂无数据")
    
    with st.sidebar.expander("📚 查看所有模板"):
        templates = load_templates()
        if templates:
            df_temp = pd.DataFrame(templates)
            st.dataframe(df_temp[['province', 'city', 'report_type', 'template_name', 'template_version']])
            st.caption(f"共 {len(templates)} 个模板")
        else:
            st.info("暂无模板")

    # ===== 规则管理面板 =====
    with st.sidebar.expander("⚙️ 规则管理（社保/公积金）"):
        st.markdown("**当前所有城市规则**")
        rules = load_rules()
        if rules:
            df_rules = pd.DataFrame(rules)
            st.dataframe(df_rules[['city', 'province', 'unit_social', 'personal_social', 'unit_fund', 'personal_fund', 'source_quote']], use_container_width=True)
            
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
                        new_social_min = st.number_input("社保基数下限", value=float(rule.get('social_min', 0)), step=100)
                        new_social_max = st.number_input("社保基数上限", value=float(rule.get('social_max', 999999)), step=100)
                        new_fund_min = st.number_input("公积金基数下限", value=float(rule.get('fund_min', 0)), step=100)
                        new_fund_max = st.number_input("公积金基数上限", value=float(rule.get('fund_max', 999999)), step=100)
                        new_source = st.text_input("来源文号", value=rule.get('source_quote', ''))
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
                                        'source_quote': new_source
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
                    new_social_min = st.number_input("社保基数下限", value=0, step=100)
                    new_social_max = st.number_input("社保基数上限", value=999999, step=100)
                    new_fund_min = st.number_input("公积金基数下限", value=0, step=100)
                    new_fund_max = st.number_input("公积金基数上限", value=999999, step=100)
                    new_source = st.text_input("来源文号")
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
                            'is_default': 0
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
                            'is_default': 1
                        })
                    save_rules(default_rules)
                    st.success("已重置为系统默认规则！")
                    st.rerun()
        else:
            st.info("暂无规则，请初始化或添加")

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
    if province:
        cities = sorted(set(c['city'] for c in companies if c['province'] == province))
    else:
        cities = sorted(set(c['city'] for c in companies))
    city = st.selectbox("城市", [""] + cities)
with col2:
    if province and city:
        districts = sorted(set(c['district'] for c in companies if c['province'] == province and c['city'] == city))
    else:
        districts = []
    district = st.selectbox("区县", [""] + districts)
    if province and city:
        company_list = [c for c in companies if c['province'] == province and c['city'] == city and (not district or c['district'] == district)]
    else:
        company_list = []
    company_names = [c['company_name'] for c in company_list]
    selected_company_names = st.multiselect("公司（可多选）", company_names)
with col3:
    report_type = st.selectbox("报表类型", ["", "增值税", "社保", "公积金", "个人所得税", "企业所得税", "年度汇算清缴"])
    
    period_type = st.selectbox("统计口径", ["月度（固定月份）", "累计（1-12月）", "自定义月份范围"])
    if period_type == "月度（固定月份）":
        month = st.selectbox("月份", list(range(1,13)), index=11)
        period_label = f"月度（{month}月）"
        custom_period = f"{month}月"
    elif period_type == "累计（1-12月）":
        month = None
        period_label = "累计（1-12月）"
        custom_period = "1-12月"
    else:
        st.markdown("**选择月份范围**")
        start_month = st.selectbox("起始月份", list(range(1,13)), index=0)
        end_month = st.selectbox("结束月份", list(range(1,13)), index=11)
        if start_month <= end_month:
            month = None
            period_label = f"自定义（{start_month}月-{end_month}月）"
            custom_period = f"{start_month}月-{end_month}月"
        else:
            st.error("起始月份不能大于结束月份")
            period_label = "自定义"
            custom_period = ""

selected_companies = [c for c in company_list if c['company_name'] in selected_company_names]

if selected_companies and report_type:
    st.markdown("---")
    st.subheader("🔍 匹配结果")
    
    matched, match_level, candidates = match_template_with_details(province, city, district, report_type)
    
    custom_templates = load_custom_templates()
    template_choice = None
    
    options = {}
    if matched:
        options[f"✅ 官方模板：{matched['template_name']}（{match_level}）"] = {'type': 'official', 'data': matched}
    for c in candidates:
        if c['id'] != (matched['id'] if matched else ''):
            options[f"📄 官方模板：{c['template_name']}（{c['province']}）"] = {'type': 'official', 'data': c}
    for ct in custom_templates:
        options[f"⭐ 自定义模板：{ct['name']}"] = {'type': 'custom', 'data': ct}
    options["🔄 通用模板（系统内置）"] = {'type': 'general', 'data': None}
    
    if len(options) > 1:
        st.info("💡 请选择要使用的模板")
        default_idx = 0
        keys = list(options.keys())
        if matched:
            for i, k in enumerate(keys):
                if "✅" in k:
                    default_idx = i
                    break
        selected_key = st.selectbox("选择模板", keys, index=default_idx)
        template_choice = options[selected_key]
    else:
        template_choice = list(options.values())[0]
    
    selected_template = None
    template_type = "通用"
    if template_choice['type'] == 'official':
        selected_template = template_choice['data']
        template_type = "官方模板"
        match_level = "官方模板"
    elif template_choice['type'] == 'custom':
        selected_template = template_choice['data']
        template_type = "自定义模板"
        match_level = "自定义模板"
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
        match_level = "通用模板"
    
    st.success(f"✅ 已选择模板：{selected_template['template_name']}（{match_level}）")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**📄 模板信息**")
        st.write(f"模板名称：{selected_template['template_name']}")
        st.write(f"版本：{selected_template.get('template_version', 'v1.0')}")
        st.write(f"发布机构：{selected_template.get('source_authority', '系统内置')}")
        st.write(f"发布日期：{selected_template.get('publish_date', datetime.now().strftime('%Y-%m-%d'))}")
        st.write(f"必填字段：{selected_template.get('required_fields', '')}")
    with col_b:
        st.markdown("**🔗 来源信息**")
        st.write(f"来源URL：[{selected_template.get('source_url', '#')}]({selected_template.get('source_url', '#')})")
        if 'province' in selected_template and selected_template['province']:
            st.write(f"适用地区：{selected_template['province']} {selected_template.get('city','')} {selected_template.get('district','')}")
        st.write(f"报表类型：{report_type}")
        st.write(f"统计口径：{period_label}")
    
    if template_type == "自定义模板" and selected_template:
        mapping = get_custom_template_field_mapping(selected_template)
        if mapping:
            st.info(f"📌 字段映射：{', '.join([f'{k}→{v}' for k, v in mapping.items()])}")
    
    st.subheader("📋 模板预览")
    fields = selected_template.get('required_fields', '').split(',')
    if fields and fields[0]:
        st.markdown(f"**字段列表**：{', '.join(fields)}")
        sample_row = {}
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
        for f in fields:
            sample_row[f] = sample_values.get(f, f'<{f} 示例值>')
        preview_df = pd.DataFrame([{'字段名': f, '示例值': sample_row[f]} for f in fields if f])
        if not preview_df.empty:
            st.dataframe(preview_df, use_container_width=True)
    
    # ===== 数据校验 =====
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
    st.info(f"📌 统计口径：{period_label}")
    
    # ---- 规则匹配显示 ----
    rule_status = []
    for comp in selected_companies:
        rule = get_rule_for_city(comp['city'])
        if rule is None:
            # 从默认规则中查找（兜底）
            default_rule = next((dr for dr in PROVINCE_DEFAULT_RULES if normalize_name(dr['city']) == normalize_name(comp['city'])), None)
            if default_rule:
                rule_status.append(f"{comp['company_name']} → {comp['city']} (将使用默认规则：{default_rule['source_quote']})")
            else:
                rule_status.append(f"{comp['company_name']} → {comp['city']} (⚠️ 无任何规则，使用通用默认值)")
        else:
            rule_status.append(f"{comp['company_name']} → {comp['city']} (规则: {rule.get('source_quote', '系统默认')})")
    st.info("📌 规则匹配情况：\n" + "\n".join(rule_status))
    
    missing = [comp['city'] for comp in selected_companies if get_rule_for_city(comp['city']) is None]
    if missing:
        st.info(f"ℹ️ 部分城市({', '.join(set(missing))})未在规则库中，将使用系统默认或通用值。")
    else:
        st.success("✅ 所有城市已匹配到规则")
    
    # ===== 报表预览 =====
    st.subheader("📊 报表预览（生成前确认）")
    preview_data = []
    for comp in selected_companies:
        rule = get_rule_for_city(comp['city'])
        if rule is None:
            default_rule = next((dr for dr in PROVINCE_DEFAULT_RULES if normalize_name(dr['city']) == normalize_name(comp['city'])), None)
            rule_source = default_rule['source_quote'] if default_rule else '通用默认'
        else:
            rule_source = rule.get('source_quote', '未配置')
        preview_data.append({
            '公司': comp['company_name'],
            '城市': comp['city'],
            '模板': selected_template['template_name'],
            '匹配级别': match_level,
            '统计口径': period_label,
            '规则来源': rule_source
        })
    st.dataframe(pd.DataFrame(preview_data), use_container_width=True)
    
    reviewed = st.checkbox("✅ 我已人工复核确认数据无误", value=False)
    
    if st.button("📥 生成待复核版Excel", disabled=not reviewed):
        if not selected_template:
            st.error("请先选择模板")
        else:
            generated_files = []
            summary = []
            errors = []
            
            for comp in selected_companies:
                try:
                    # 获取规则（含兜底）
                    rule = get_rule_for_city(comp['city'])
                    if rule is None:
                        default_rule = next((dr for dr in PROVINCE_DEFAULT_RULES if normalize_name(dr['city']) == normalize_name(comp['city'])), None)
                        if default_rule:
                            st.warning(f"⚠️ 城市 {comp['city']} 未在规则库中找到，将使用系统默认规则（{default_rule['source_quote']}）")
                            rule = {
                                'unit_social': default_rule['unit_social'],
                                'personal_social': default_rule['personal_social'],
                                'unit_fund': default_rule['unit_fund'],
                                'personal_fund': default_rule['personal_fund'],
                                'social_min': default_rule.get('social_min', 0),
                                'social_max': default_rule.get('social_max', 999999),
                                'fund_min': default_rule.get('fund_min', 0),
                                'fund_max': default_rule.get('fund_max', 999999),
                                'source_quote': default_rule.get('source_quote', '系统默认')
                            }
                        else:
                            st.warning(f"⚠️ 城市 {comp['city']} 没有任何规则，将使用通用默认值（16%/8%）")
                            rule = {
                                'unit_social': 0.16,
                                'personal_social': 0.08,
                                'unit_fund': 0.12,
                                'personal_fund': 0.12,
                                'social_min': 0,
                                'social_max': 999999,
                                'fund_min': 0,
                                'fund_max': 999999,
                                'source_quote': '系统默认'
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
                    ws['A2'] = f'模板名称：{selected_template["template_name"]}  版本：{selected_template.get("template_version", "v1.0")}  匹配级别：{match_level}'
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
                    audit.append([datetime.now().isoformat(), 'GENERATED', '系统', f'公司:{comp["company_name"]}, 城市:{comp["city"]}, 模板:{selected_template["template_name"]}, 匹配级别:{match_level}, 数据来源:{data_source_text}, 统计口径:{period_label}, 规则:{rule.get("source_quote", "未配置")}'])
                    
                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    fname = f"{comp['company_name']}_{report_type}_{period_label.replace('（','_').replace('）','').replace('-','_')}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    generated_files.append((fname, output.getvalue()))
                    summary.append({
                        '公司': comp['company_name'], 
                        '城市': comp['city'], 
                        '模板': selected_template['template_name'],
                        '匹配级别': match_level,
                        '统计口径': period_label,
                        '规则来源': rule.get('source_quote', '未配置'),
                        '状态': '待复核'
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
                        'custom_period': custom_period if period_type == "自定义月份范围" else ''
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

# ===== 自定义模板管理 =====
with st.expander("📄 自定义模板管理（上传自己的模板）"):
    st.markdown("上传您自己的Excel模板，并映射字段到系统数据")
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
                for h in headers:
                    col = st.selectbox(
                        f"字段 '{h}' 映射到",
                        [''] + ['纳税人识别号', '公司名称', '销售额', '进项税额', '应纳税额', '单位名称', '社保登记号', '基数', '单位金额', '个人金额', '单位比例', '个人比例', '公积金账号', '收入额', '专项扣除', '营业收入', '营业成本', '应纳税所得额', '全年收入', '全年成本', '已预缴税额', '应补退税额', '申报金额'],
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

# ===== 导出历史 =====
with st.expander("📋 导出历史记录"):
    history = load_export_history()
    if history:
        df_hist = pd.DataFrame(history)
        st.dataframe(df_hist[['company_name', 'city', 'report_type', 'period_type', 'data_source', 'generated_at', 'review_status', 'custom_period']], use_container_width=True)
        
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
        provinces_in_templates = sorted(set(t['province'] for t in templates if t['province']))
        selected_province = st.selectbox("选择省份查看模板", [""] + provinces_in_templates)
        if selected_province:
            filtered = [t for t in templates if t['province'] == selected_province]
            st.dataframe(pd.DataFrame(filtered)[['city', 'district', 'report_type', 'template_name', 'template_version', 'source_authority']])
        else:
            st.dataframe(pd.DataFrame(templates)[['province', 'city', 'report_type', 'template_name', 'template_version']])
        st.caption(f"共 {len(templates)} 个官方模板")
    else:
        st.info("暂无模板")
