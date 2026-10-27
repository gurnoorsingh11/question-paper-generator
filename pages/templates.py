import json
import streamlit as st
from database.connection import get_connection

DEFAULT_SECTIONS = [
    {"name": "A", "count": 10, "marks_each": 2, "attempt": 10, "minimum_attempt": 10},
    {"name": "B", "count": 4, "marks_each": 8, "attempt": 3, "minimum_attempt": 2},
    {"name": "C", "count": 4, "marks_each": 8, "attempt": 2, "minimum_attempt": 2},
]

def render():
    st.header("Exam Templates")
    st.caption("Create reusable patterns. Attempt rules determine the maximum marks obtainable.")
    with st.form("new_template"):
        name = st.text_input("Template name", "SVIET Standard Theory Paper")
        department = st.text_input("Department", "Computer Science & Engineering")
        subject = st.text_input("Subject name")
        code = st.text_input("Subject code")
        semester = st.text_input("Semester", "1")
        m_code = st.text_input("M. Code")
        exam_date = st.text_input("Examination date")
        duration = st.number_input("Duration (minutes)", 30, 600, 180, 15)
        maximum = st.number_input("Declared maximum marks", 1, 1000, 60)
        st.markdown("#### Section A")
        a_count = st.number_input("A: questions offered", 1, 100, 10)
        a_marks = st.number_input("A: marks each", 1, 100, 2)
        a_attempt = st.number_input("A: questions to attempt", 1, 100, 10)
        st.markdown("#### Section B")
        b_count = st.number_input("B: questions offered", 1, 100, 4)
        b_marks = st.number_input("B: marks each", 1, 100, 8)
        b_attempt = st.number_input("B: questions to attempt", 0, 100, 3)
        b_min = st.number_input("B: minimum attempts", 0, 100, 2)
        st.markdown("#### Section C")
        c_count = st.number_input("C: questions offered", 1, 100, 4)
        c_marks = st.number_input("C: marks each", 1, 100, 8)
        c_attempt = st.number_input("C: questions to attempt", 0, 100, 2)
        c_min = st.number_input("C: minimum attempts", 0, 100, 2)
        instructions = st.text_area("Instructions to candidates", "1. SECTION-A is COMPULSORY.\n2. Attempt the prescribed number of questions from Sections B and C.\n3. Follow the minimum selection requirement for each section.")
        submitted = st.form_submit_button("Save template")
    if submitted:
        sections = [
            {"name": "A", "count": int(a_count), "marks_each": int(a_marks), "attempt": int(a_attempt), "minimum_attempt": int(a_attempt)},
            {"name": "B", "count": int(b_count), "marks_each": int(b_marks), "attempt": int(b_attempt), "minimum_attempt": int(b_min)},
            {"name": "C", "count": int(c_count), "marks_each": int(c_marks), "attempt": int(c_attempt), "minimum_attempt": int(c_min)},
        ]
        calc = sum(s["attempt"] * s["marks_each"] for s in sections)
        if any(s["attempt"] > s["count"] or s["minimum_attempt"] > s["attempt"] for s in sections):
            st.error("Invalid section attempt rules. Attempts cannot exceed questions offered; minimum attempts cannot exceed attempts.")
        elif calc != int(maximum):
            st.error(f"Attempt rules calculate {calc} marks, but declared maximum is {maximum}. Correct one of these values before saving.")
        else:
            with get_connection() as conn:
                conn.execute("""
                    INSERT INTO templates(name, department, subject_name, subject_code, semester,
                    m_code, exam_date, duration_minutes, maximum_marks, instructions, sections_json)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (name, department, subject, code, semester, m_code, exam_date, int(duration),
                      int(maximum), instructions, json.dumps(sections)))
            st.success("Template saved.")
            st.rerun()
    with get_connection() as conn:
        rows = conn.execute("SELECT * FROM templates ORDER BY id DESC").fetchall()
    st.subheader("Saved templates")
    if not rows:
        st.info("No templates saved yet.")
    for row in rows:
        with st.expander(f"#{row['id']} · {row['name']} — {row['maximum_marks']} marks"):
            st.write(f"Department: {row['department']} · Subject: {row['subject_name']} · Semester: {row['semester']}")
            st.write(f"Duration: {row['duration_minutes']} minutes · Subject code: {row['subject_code']} · M. Code: {row['m_code']}")
            st.json(json.loads(row["sections_json"]))
            st.write(row["instructions"])
