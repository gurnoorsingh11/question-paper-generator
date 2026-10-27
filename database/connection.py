import sqlite3
from pathlib import Path
from config import DATABASE_PATH

def get_connection():
    Path(DATABASE_PATH).parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DATABASE_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

def init_db():
    with get_connection() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS question_banks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            source_filename TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bank_id INTEGER NOT NULL REFERENCES question_banks(id) ON DELETE CASCADE,
            parent_id INTEGER REFERENCES questions(id) ON DELETE CASCADE,
            question_text TEXT NOT NULL,
            answer_text TEXT NOT NULL DEFAULT '',
            marks INTEGER,
            eligible_sections TEXT NOT NULL DEFAULT 'A,B,C',
            question_type TEXT NOT NULL DEFAULT 'standalone',
            approved INTEGER NOT NULL DEFAULT 0,
            position INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            department TEXT NOT NULL DEFAULT '',
            subject_name TEXT NOT NULL DEFAULT '',
            subject_code TEXT NOT NULL DEFAULT '',
            semester TEXT NOT NULL DEFAULT '',
            m_code TEXT NOT NULL DEFAULT '',
            exam_date TEXT NOT NULL DEFAULT '',
            duration_minutes INTEGER NOT NULL DEFAULT 180,
            maximum_marks INTEGER NOT NULL DEFAULT 60,
            instructions TEXT NOT NULL DEFAULT '',
            sections_json TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS generated_papers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            template_id INTEGER,
            title TEXT NOT NULL,
            paper_json TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        """)
