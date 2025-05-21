# metadata_db.py
# 用于管理文件元数据的数据库操作
import sqlite3
import json
import os

DB_PATH = "file_metadata.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS file_metadata (
            file_path TEXT PRIMARY KEY,
            modified_time TEXT,
            metadata_json TEXT
        )
    ''')
    conn.commit()
    conn.close()

def upsert_metadata(file_path, modified_time, metadata_dict):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    metadata_json = json.dumps(metadata_dict, ensure_ascii=False)
    c.execute('''
        INSERT INTO file_metadata (file_path, modified_time, metadata_json)
        VALUES (?, ?, ?)
        ON CONFLICT(file_path) DO UPDATE SET
            modified_time=excluded.modified_time,
            metadata_json=excluded.metadata_json
    ''', (file_path, modified_time, metadata_json))
    conn.commit()
    conn.close()

def get_cached_metadata(file_path, current_modified_time):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT modified_time, metadata_json FROM file_metadata WHERE file_path = ?', (file_path,))
    row = c.fetchone()
    conn.close()

    if row and row[0] == current_modified_time:
        return json.loads(row[1])
    return None

def load_all_metadata():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT metadata_json FROM file_metadata')
    rows = c.fetchall()
    conn.close()
    return [json.loads(row[0]) for row in rows]
