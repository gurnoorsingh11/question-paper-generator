import io
import random
import re
from datetime import date

import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="SVIET Question Paper Generator", page_icon="📄", layout="centered")

MARK_OPTIONS = [2, 5, 10]


def clean_question(text):
    text = str(text).strip()
    text = re.sub(r"^\s*(?:Q(?:uestion)?\s*)?\d+\s*[.)、:-]\s*", "", text, flags=re.I)
    return re.sub(r"\s+", " ", text).strip()


def parse_word(file):
    doc = Document(file)
    rows = []
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        if re.match(r"^(unit\s+[ivx\d]+|course\s*(?:name|code)?\s*:|subject\s*:|department\s*:|semester\s*:)", text, re.I):
            continue
        text = clean_question(text)
        if len(text) >= 8:
            rows.append({"Question": text, "Marks": None})
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if not any(cells):
                continue
            question = next((c for c in cells if len(c) >= 8), "")
            if question and question.lower() not in {"question", "questions", "question text"}:
                rows.append({"Question": clean_question(question), "Marks": None})
    return rows


def parse_excel(file):
    workbook = pd.ExcelFile(file)
    rows = []
    for sheet in workbook.sheet_names:
        df = pd.read_excel(workbook, sheet_name=sheet, header=None).fillna("")
        for _, row in df.iterrows():
            cells = [str(value).strip() for value in row.tolist() if str(value).strip()]
            if not cells:
                continue
            joined = " ".join(cells)
            if re.match(r"^(unit\s+[ivx\d]+|question\s*(?:text)?|questions?)\s*$", joined, re.I):
                continue
            # Prefer a cell that looks like a full question; otherwise combine the row.
            question = max(cells, key=len)
            if len(question) >= 8:
                rows.append({"Question": clean_question(question), "Marks": None})
    return rows


def make_docx(meta, paper):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = section.bottom_margin = __import__("docx").shared.Inches(0.7)
    p = doc.add_paragraph()
    p.alignment = 1
    run = p.add_run(meta["college"])
    run.bold = True
    run.font.size = __import__("docx").shared.Pt(15)
    p = doc.add_paragraph()
    p.alignment = 1
    run = p.add_run("QUESTION PAPER")
    run.bold = True
    run.font.size = __import__("docx").shared.Pt(13)
    doc.add_paragraph(f"Department: {meta['department']}\tSemester: {meta['semester']}")
    doc.add_paragraph(f"Subject: {meta['subject']} ({meta['subject_code']})")
    doc.add_paragraph(f"M. Code: {meta['m_code']}\tDuration: {meta['duration']}\tMaximum Marks: {meta['total_marks']}")
    doc.add_paragraph("Instructions: Attempt the questions as directed by the examiner.")
    current_marks = None
    number = 0
    for item in paper:
        if item["Marks"] != current_marks:
            current_marks = item["Marks"]
            doc.add_paragraph(f"Section — {current_marks} Marks", style="Heading 2")
        number += 1
        doc.add_paragraph(f"{number}. {item['Question']} ({item['Marks']} marks)")
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


st.title("SVIET Question Paper Generator")
st.caption("Upload questions → assign marks → generate → preview → download Word")

with st.expander("1. Paper details", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        college = st.text_input("College name", "Swami Vivekanand Institute of Engineering and Technology")
        department = st.text_input("Department", "Computer Science & Engineering")
        subject = st.text_input("Subject name")
        subject_code = st.text_input("Subject code")
    with c2:
        semester = st.text_input("Semester")
        m_code = st.text_input("M. Code")
        duration = st.text_input("Exam duration", "3 Hours")
        total_marks = st.number_input("Maximum marks", min_value=1, max_value=300, value=70, step=1)

st.subheader("2. Upload question bank")
upload = st.file_uploader("Choose a Word (.docx) or Excel (.xlsx/.xls) file", type=["docx", "xlsx", "xls"])
if upload:
    try:
        if upload.name.lower().endswith(".docx"):
            imported = parse_word(upload)
        else:
            imported = parse_excel(upload)
        # Remove exact duplicates while preserving order.
        seen, unique = set(), []
        for row in imported:
            key = row["Question"].casefold()
            if key and key not in seen:
                seen.add(key)
                unique.append(row)
        st.session_state["questions"] = unique
        st.session_state["source_name"] = upload.name
        st.success(f"Imported {len(unique)} possible questions from {upload.name}. Please review them below.")
    except Exception as exc:
        st.error(f"Couldn't read this file: {exc}")

if "questions" in st.session_state and st.session_state["questions"]:
    st.subheader("3. Review questions and assign marks")
    st.caption("Choose 2, 5, or 10 marks for every question you want the app to use. Leave unused questions unassigned.")
    df = pd.DataFrame(st.session_state["questions"])
    if "Marks" not in df.columns:
        df["Marks"] = None
    edited = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "Question": st.column_config.TextColumn("Question", required=True, width="large"),
            "Marks": st.column_config.SelectboxColumn("Marks", options=MARK_OPTIONS, required=False),
        },
        key="question_editor",
    )
    st.session_state["edited_questions"] = edited
    ready = edited.dropna(subset=["Marks"]).copy()
    ready["Marks"] = ready["Marks"].astype(int)
    st.write(f"Classified questions ready to use: **{len(ready)}**")
    counts = {mark: int((ready["Marks"] == mark).sum()) for mark in MARK_OPTIONS}

    st.subheader("4. Choose how many questions to include")
    a, b, c = st.columns(3)
    with a:
        n2 = st.number_input(f"2-mark questions (available: {counts[2]})", min_value=0, max_value=counts[2], value=min(5, counts[2]), step=1)
    with b:
        n5 = st.number_input(f"5-mark questions (available: {counts[5]})", min_value=0, max_value=counts[5], value=min(4, counts[5]), step=1)
    with c:
        n10 = st.number_input(f"10-mark questions (available: {counts[10]})", min_value=0, max_value=counts[10], value=min(3, counts[10]), step=1)
    calculated_total = int(n2) * 2 + int(n5) * 5 + int(n10) * 10
    st.write(f"Marks from selected questions: **{calculated_total}**")
    if calculated_total != int(total_marks):
        st.warning(f"The selected questions add up to {calculated_total} marks, but Maximum marks is {int(total_marks)}. Adjust either the question counts or Maximum marks.")

    if st.button("Generate / Shuffle Paper", type="primary", use_container_width=True):
        if not subject.strip():
            st.error("Enter the subject name first.")
        elif calculated_total != int(total_marks):
            st.error("The question marks must equal Maximum marks before generating the paper.")
        else:
            paper = []
            for mark, count in [(2, int(n2)), (5, int(n5)), (10, int(n10))]:
                pool = ready[ready["Marks"] == mark]["Question"].astype(str).tolist()
                paper.extend({"Question": q, "Marks": mark} for q in random.sample(pool, count))
            random.shuffle(paper)
            st.session_state["generated_paper"] = paper
            st.session_state["generated_meta"] = {
                "college": college, "department": department, "subject": subject,
                "subject_code": subject_code, "semester": semester, "m_code": m_code,
                "duration": duration, "total_marks": int(total_marks),
            }

if "generated_paper" in st.session_state:
    st.subheader("5. Preview")
    meta = st.session_state["generated_meta"]
    st.markdown(f"**{meta['subject']} ({meta['subject_code']})**  \nDepartment: {meta['department']} | Semester: {meta['semester']}  \nDuration: {meta['duration']} | Maximum marks: {meta['total_marks']}")
    for i, question in enumerate(st.session_state["generated_paper"], start=1):
        st.write(f"**{i}.** {question['Question']}  *({question['Marks']} marks)*")
    docx_bytes = make_docx(meta, st.session_state["generated_paper"])
    st.download_button(
        "Download Word Paper (.docx)", data=docx_bytes,
        file_name=f"{re.sub(r'[^A-Za-z0-9_-]+', '_', meta['subject']).strip('_') or 'Question_Paper'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )
else:
    st.info("Your generated paper will appear here after you classify questions and click Generate / Shuffle Paper.")

st.divider()
st.caption("Simple version: no account, database, templates, or automatic marks classification. Uploaded questions are reviewed manually before use.")
