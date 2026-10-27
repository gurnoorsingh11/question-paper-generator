import json
from database.connection import get_connection

def create_bank(name, filename, questions):
    with get_connection() as conn:
        cur = conn.execute(
            "INSERT INTO question_banks(name, source_filename) VALUES (?, ?)",
            (name, filename),
        )
        bank_id = cur.lastrowid
        for i, q in enumerate(questions):
            conn.execute("""
                INSERT INTO questions
                (bank_id, question_text, answer_text, marks, eligible_sections, question_type, approved, position)
                VALUES (?, ?, ?, ?, ?, 'standalone', 0, ?)
            """, (bank_id, q["question_text"], q.get("answer_text", ""),
                  q.get("marks"), "A,B,C", i))
        return bank_id

def list_banks():
    with get_connection() as conn:
        return [dict(r) for r in conn.execute(
            "SELECT b.*, COUNT(q.id) AS question_count FROM question_banks b "
            "LEFT JOIN questions q ON q.bank_id=b.id GROUP BY b.id ORDER BY b.id DESC"
        ).fetchall()]

def get_questions(bank_id):
    with get_connection() as conn:
        return [dict(r) for r in conn.execute(
            "SELECT * FROM questions WHERE bank_id=? ORDER BY position, id", (bank_id,)
        ).fetchall()]

def update_question(qid, text, answer, marks, sections, approved):
    with get_connection() as conn:
        conn.execute("""
            UPDATE questions SET question_text=?, answer_text=?, marks=?,
            eligible_sections=?, approved=? WHERE id=?
        """, (text, answer, marks, ",".join(sections), int(approved), qid))
