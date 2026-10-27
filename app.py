from __future__ import annotations

import io
import random
import re
from datetime import datetime
from typing import Any

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt

st.set_page_config(
    page_title="SVIET Question Paper Generator",
    page_icon="📄",
    layout="wide",
)

SUPPORTED_MARKS = [2, 5, 10]
EMPTY_BANK = pd.DataFrame({"Question": pd.Series(dtype="str")})


def clean_text(value: Any) -> str:
    """Normalize whitespace without removing meaningful question wording."""
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def strip_question_number(text: str) -> str:
    # Only remove a leading question number; leave subparts such as (a) intact.
    return re.sub(
        r"^\s*(?:(?:Q(?:uestion)?\s*)?\d{1,3}\s*[.)\-:]\s*)",
        "",
        text,
        flags=re.IGNORECASE,
    ).strip()


def looks_like_heading(text: str) -> bool:
    if not text:
        return True
    patterns = [
        r"^(?:question\s*)?bank(?:\s+of\s+questions)?$",
        r"^questions?$",
        r"^question\s+text$",
        r"^s\.?\s*no\.?$",
        r"^serial\s+no\.?$",
        r"^unit\s*(?:[ivx]+|\d+)(?:\s*[:\-].*)?$",
        r"^(?:course|subject|department|semester|college|marks?)\s*(?:name|code)?\s*:?$",
    ]
    return any(re.match(pattern, text, flags=re.IGNORECASE) for pattern in patterns)


def deduplicate_questions(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        question = clean_text(value)
        question = strip_question_number(question)
        if len(question) < 3 or looks_like_heading(question):
            continue
        key = re.sub(r"\W+", " ", question.casefold()).strip()
        if key and key not in seen:
            result.append(question)
            seen.add(key)
    return result


def parse_docx(uploaded_file) -> list[str]:
    doc = Document(io.BytesIO(uploaded_file.getvalue()))
    values: list[str] = []
    # Read paragraphs and table cells. Deduplication below handles overlap.
    for paragraph in doc.paragraphs:
        text = clean_text(paragraph.text)
        if text and not looks_like_heading(text):
            values.append(text)
    for table in doc.tables:
        for row in table.rows:
            cells = [clean_text(cell.text) for cell in row.cells]
            cells = [cell for cell in cells if cell]
            if not cells:
                continue
            # Most banks put the question in one cell, sometimes with S. No. in another.
            candidate = max(cells, key=len)
            if not looks_like_heading(candidate):
                values.append(candidate)
    return deduplicate_questions(values)


def parse_excel(uploaded_file) -> list[str]:
    book = pd.ExcelFile(io.BytesIO(uploaded_file.getvalue()))
    values: list[str] = []
    for sheet_name in book.sheet_names:
        frame = pd.read_excel(book, sheet_name=sheet_name, header=None, dtype=object).fillna("")
        for _, row in frame.iterrows():
            cells = [clean_text(value) for value in row.tolist()]
            cells = [cell for cell in cells if cell]
            if not cells:
                continue
            # Ignore obvious column labels; otherwise take the most question-like cell.
            if len(cells) <= 3 and all(
                re.fullmatch(r"(?:s\.?\s*no\.?|sr\.?\s*no\.?|question|questions|question text|marks?)", c, re.I)
                for c in cells
            ):
                continue
            candidate = max(cells, key=len)
            if not looks_like_heading(candidate):
                values.append(candidate)
    return deduplicate_questions(values)


def read_question_file(uploaded_file) -> list[str]:
    if uploaded_file is None:
        return []
    filename = uploaded_file.name.lower()
    if filename.endswith(".docx"):
        return parse_docx(uploaded_file)
    if filename.endswith((".xlsx", ".xls")):
        return parse_excel(uploaded_file)
    raise ValueError("Use a Word .docx or Excel .xlsx/.xls file.")


def init_state() -> None:
    defaults = {
        "short_bank": [],
        "long_bank": [],
        "generated_paper": None,
        "last_short_import": None,
        "last_long_import": None,
        "last_template_name": None,
        "short_editor_data": None,
        "long_editor_data": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def add_questions(existing: list[str], incoming: list[str]) -> tuple[list[str], int]:
    combined = list(existing)
    seen = {re.sub(r"\W+", " ", q.casefold()).strip() for q in combined}
    added = 0
    for question in incoming:
        key = re.sub(r"\W+", " ", clean_text(question).casefold()).strip()
        if key and key not in seen:
            combined.append(clean_text(question))
            seen.add(key)
            added += 1
    return combined, added


def load_upload(uploaded_file, bank_key: str, import_key: str, editor_key: str) -> None:
    if uploaded_file is None:
        return
    fingerprint = f"{uploaded_file.name}:{hash(uploaded_file.getvalue())}"
    if st.session_state.get(import_key) == fingerprint:
        return
    try:
        parsed = read_question_file(uploaded_file)
        combined, added = add_questions(st.session_state[bank_key], parsed)
        st.session_state[bank_key] = combined
        st.session_state[import_key] = fingerprint
        st.session_state.pop(editor_key, None)
        st.session_state.generated_paper = None
        st.success(f"Imported {added} new question(s). {len(combined)} available in this bank.")
    except Exception as exc:
        st.error(f"Could not read **{uploaded_file.name}**: {exc}")


def build_question_document(meta: dict[str, Any], sections: dict[str, list[str]], template_bytes: bytes | None) -> bytes:
    """Build a printable DOCX. If supplied, keep the template and append the paper after a page break."""
    if template_bytes:
        try:
            doc = Document(io.BytesIO(template_bytes))
        except Exception as exc:
            raise ValueError(f"The template could not be opened as a Word document: {exc}") from exc
        doc.add_page_break()
    else:
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    def add_centered(text: str, size: int, bold: bool = False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(text)
        run.bold = bold
        run.font.size = Pt(size)
        return p

    add_centered(meta["college"], 14, True)
    add_centered("QUESTION PAPER", 13, True)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.add_run(f"Department: {meta['department']}\t\tSemester: {meta['semester']}")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.add_run(f"Subject: {meta['subject']}\t\tSubject Code: {meta['subject_code']}")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.add_run(
        f"M. Code: {meta['m_code']}\t\tDuration: {meta['duration']}\t\tMaximum Marks: {meta['total_marks']}"
    )
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("Instructions: ")
    r.bold = True
    p.add_run(meta["instructions"])

    def add_section_heading(title: str):
        # Direct formatting works even when an uploaded template has no Heading styles.
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(3)
        r = p.add_run(title)
        r.bold = True
        r.font.size = Pt(12)

    def add_questions(title: str, questions: list[str], marks_label: str | None = None, numbering: bool = True):
        if not questions:
            return
        add_section_heading(title)
        for index, question in enumerate(questions, start=1):
            label = f"{index}. " if numbering else ""
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_together = True
            p.add_run(label + question)
            if marks_label:
                r = p.add_run(f" ({marks_label} marks)")
                r.italic = True

    add_questions("SECTION A — Short Answer Questions", sections.get("A", []), "2")
    add_questions("SECTION B — Long Answer Questions", sections.get("B", []), meta["b_marks"])
    add_questions("SECTION C — Long Answer Questions", sections.get("C", []), meta["c_marks"])

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def make_selection(bank: list[str], count: int, excluded: set[str] | None = None) -> list[str]:
    excluded = excluded or set()
    pool = [q for q in bank if q not in excluded]
    return random.sample(pool, min(max(0, int(count)), len(pool)))


def safe_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return value.strip("._-") or "Question_Paper"


def add_manual_question(bank_key: str, text_key: str, editor_key: str, bank_label: str) -> None:
    """Callback runs before widgets render, so resetting the editor cannot conflict with Streamlit."""
    question = clean_text(st.session_state.get(text_key, ""))
    if len(question) < 3:
        st.session_state["manual_feedback"] = "Enter the question text first."
        st.session_state["manual_feedback_kind"] = "warning"
        return
    updated, added = add_questions(st.session_state.get(bank_key, []), [question])
    st.session_state[bank_key] = updated
    st.session_state.generated_paper = None
    st.session_state.pop(editor_key, None)
    if added:
        st.session_state["manual_feedback"] = f"Question added to {bank_label}."
        st.session_state["manual_feedback_kind"] = "success"
    else:
        st.session_state["manual_feedback"] = "That question is already in the bank."
        st.session_state["manual_feedback_kind"] = "warning"


init_state()

st.title("SVIET Question Paper Generator")
st.caption("A simple workflow: import two question banks → set counts → shuffle → preview → download Word paper.")

with st.expander("1. Exam details", expanded=True):
    left, right = st.columns(2)
    with left:
        college = st.text_input("College name", "Swami Vivekanand Institute of Engineering and Technology")
        department = st.text_input("Department", "Computer Science & Engineering")
        subject = st.text_input("Subject name", placeholder="e.g. Database Management Systems")
        subject_code = st.text_input("Subject code", placeholder="e.g. BCSE-301")
        semester = st.text_input("Semester", placeholder="e.g. 5th")
    with right:
        m_code = st.text_input("M. Code", placeholder="Optional")
        duration = st.text_input("Exam duration", "3 Hours")
        total_marks = st.number_input("Maximum marks", min_value=1, max_value=300, value=70, step=1)
        b_marks = st.selectbox("Marks per Section B question", [5, 10], index=1)
        c_marks = st.selectbox("Marks per Section C question", [5, 10], index=1)
    instructions = st.text_area(
        "Instructions printed on paper",
        "Attempt the questions as directed. All questions carry marks as indicated.",
        height=68,
    )

st.subheader("2. Import question banks")
short_col, long_col = st.columns(2)
with short_col:
    st.markdown("**Section A — 2-mark questions**")
    short_upload = st.file_uploader(
        "Upload short-question bank (Word or Excel)",
        type=["docx", "xlsx", "xls"],
        key="short_upload",
        help="Each paragraph, table row, or spreadsheet row should contain one question.",
    )
    load_upload(short_upload, "short_bank", "last_short_import", "short_bank_editor")
    st.metric("Short questions available", len(st.session_state.short_bank))
with long_col:
    st.markdown("**Sections B and C — long questions**")
    long_upload = st.file_uploader(
        "Upload long-question bank (Word or Excel)",
        type=["docx", "xlsx", "xls"],
        key="long_upload",
        help="Sections B and C draw from this same bank and do not repeat a question in one paper.",
    )
    load_upload(long_upload, "long_bank", "last_long_import", "long_bank_editor")
    st.metric("Long questions available", len(st.session_state.long_bank))

with st.expander("Optional: use a Word exam template", expanded=False):
    template_upload = st.file_uploader(
        "Upload .docx template (optional)",
        type=["docx"],
        key="template_upload",
        help="The template's existing pages are preserved. The generated question paper starts after a page break.",
    )
    if template_upload:
        st.caption(f"Template selected: {template_upload.name}. Existing template content will be kept.")

st.subheader("3. Review or add questions")
review_short, review_long = st.tabs(["Section A bank", "Sections B/C bank"])
with review_short:
    st.caption("Edit imported questions directly, or add one manually below.")
    short_df = pd.DataFrame({"Question": st.session_state.short_bank})
    if short_df.empty:
        short_df = EMPTY_BANK.copy()
    edited_short = st.data_editor(
        short_df,
        key="short_bank_editor",
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={"Question": st.column_config.TextColumn("Question text", required=True, width="large")},
    )
    st.session_state.short_bank = deduplicate_questions(edited_short["Question"].fillna("").astype(str).tolist()) if "Question" in edited_short else []
    with st.form("add_short_form", clear_on_submit=True):
        st.text_area("Add a 2-mark question", placeholder="Type one complete question…", key="new_short")
        st.form_submit_button(
            "Add to Section A",
            on_click=add_manual_question,
            args=("short_bank", "new_short", "short_bank_editor", "Section A"),
        )
    if st.session_state.get("manual_feedback"):
        if st.session_state.get("manual_feedback_kind") == "success":
            st.success(st.session_state["manual_feedback"])
        else:
            st.warning(st.session_state["manual_feedback"])
        st.session_state.pop("manual_feedback", None)
        st.session_state.pop("manual_feedback_kind", None)
with review_long:
    st.caption("The same long-question bank feeds Sections B and C. A question will not appear in both sections in one paper.")
    long_df = pd.DataFrame({"Question": st.session_state.long_bank})
    if long_df.empty:
        long_df = EMPTY_BANK.copy()
    edited_long = st.data_editor(
        long_df,
        key="long_bank_editor",
        num_rows="dynamic",
        hide_index=True,
        use_container_width=True,
        column_config={"Question": st.column_config.TextColumn("Question text", required=True, width="large")},
    )
    st.session_state.long_bank = deduplicate_questions(edited_long["Question"].fillna("").astype(str).tolist()) if "Question" in edited_long else []
    with st.form("add_long_form", clear_on_submit=True):
        st.text_area("Add a long-answer question", placeholder="Type one complete question…", key="new_long")
        st.form_submit_button(
            "Add to Sections B/C",
            on_click=add_manual_question,
            args=("long_bank", "new_long", "long_bank_editor", "Sections B/C"),
        )
    if st.session_state.get("manual_feedback"):
        if st.session_state.get("manual_feedback_kind") == "success":
            st.success(st.session_state["manual_feedback"])
        else:
            st.warning(st.session_state["manual_feedback"])
        st.session_state.pop("manual_feedback", None)
        st.session_state.pop("manual_feedback_kind", None)

st.subheader("4. Set question counts")
count_left, count_mid, count_right = st.columns(3)
with count_left:
    count_a = st.number_input("Section A: 2-mark questions", min_value=0, max_value=200, value=10, step=1)
with count_mid:
    count_b = st.number_input("Section B: long questions", min_value=0, max_value=100, value=5, step=1)
with count_right:
    count_c = st.number_input("Section C: long questions", min_value=0, max_value=100, value=3, step=1)

short_available = len(st.session_state.short_bank)
long_available = len(st.session_state.long_bank)
if count_a > short_available:
    st.info(f"Section A requests {count_a}, but only {short_available} are available. The app will use all available short questions.")
if count_b + count_c > long_available:
    st.info(f"Sections B and C request {count_b + count_c} distinct long questions, but only {long_available} are available. The app will use all available long questions, without duplicates.")

button_cols = st.columns(4)
with button_cols[0]:
    generate_clicked = st.button("🎲 Shuffle All", type="primary", use_container_width=True)
with button_cols[1]:
    shuffle_a_clicked = st.button("Shuffle A", use_container_width=True)
with button_cols[2]:
    shuffle_b_clicked = st.button("Shuffle B", use_container_width=True)
with button_cols[3]:
    shuffle_c_clicked = st.button("Shuffle C", use_container_width=True)

if st.button("Clear generated paper"):
    st.session_state.generated_paper = None
    st.rerun()

meta = {
    "college": college.strip() or "Swami Vivekanand Institute of Engineering and Technology",
    "department": department.strip() or "Department",
    "subject": subject.strip() or "Subject",
    "subject_code": subject_code.strip() or "—",
    "semester": semester.strip() or "—",
    "m_code": m_code.strip() or "—",
    "duration": duration.strip() or "—",
    "total_marks": int(total_marks),
    "b_marks": int(b_marks),
    "c_marks": int(c_marks),
    "instructions": instructions.strip() or "Attempt the questions as directed.",
}

if generate_clicked or shuffle_a_clicked or shuffle_b_clicked or shuffle_c_clicked:
    previous = st.session_state.generated_paper
    if previous:
        sections = {key: list(previous["sections"].get(key, [])) for key in ("A", "B", "C")}
    else:
        sections = {"A": [], "B": [], "C": []}

    if generate_clicked or not previous:
        sections["A"] = make_selection(st.session_state.short_bank, int(count_a))
        selected_long = make_selection(st.session_state.long_bank, int(count_b) + int(count_c))
        sections["B"] = selected_long[: min(int(count_b), len(selected_long))]
        sections["C"] = selected_long[len(sections["B"]):len(sections["B"]) + int(count_c)]
    elif shuffle_a_clicked:
        sections["A"] = make_selection(st.session_state.short_bank, int(count_a))
    elif shuffle_b_clicked:
        excluded = set(sections["C"])
        sections["B"] = make_selection(st.session_state.long_bank, int(count_b), excluded)
    elif shuffle_c_clicked:
        excluded = set(sections["B"])
        sections["C"] = make_selection(st.session_state.long_bank, int(count_c), excluded)

    template_bytes = template_upload.getvalue() if template_upload else None
    try:
        output_bytes = build_question_document(meta, sections, template_bytes)
        st.session_state.generated_paper = {
            "meta": meta,
            "sections": sections,
            "bytes": output_bytes,
            "filename": safe_filename(f"{meta['subject']}_Question_Paper") + ".docx",
        }
    except Exception as exc:
        st.error(f"Could not create the Word paper: {exc}")

paper = st.session_state.generated_paper
if paper:
    st.subheader("5. Preview paper")
    meta = paper["meta"]
    st.markdown(f"**{meta['college']}**  \n**QUESTION PAPER**  \n{meta['subject']} ({meta['subject_code']}) · Semester {meta['semester']} · {meta['duration']} · Maximum marks: {meta['total_marks']}")
    section_labels = [("A", "Section A — 2 marks each"), ("B", f"Section B — {meta['b_marks']} marks each"), ("C", f"Section C — {meta['c_marks']} marks each")]
    for section_key, label in section_labels:
        questions = paper["sections"][section_key]
        if questions:
            st.markdown(f"**{label}** · {len(questions)} question(s)")
            for i, question in enumerate(questions, 1):
                st.write(f"{i}. {question}")
        else:
            st.caption(f"{label}: no questions selected")
    st.download_button(
        "⬇️ Download Word paper (.docx)",
        data=paper["bytes"],
        file_name=paper["filename"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )
    st.caption("The downloaded file is editable in Microsoft Word and can be printed. Use Shuffle A, Shuffle B, or Shuffle C to change only that section.")
else:
    st.info("Import or add questions, set the counts, then click **Shuffle All** to preview the paper. Individual section shuffle buttons preserve the other sections.")

with st.expander("How to prepare your question bank", expanded=False):
    st.markdown(
        "- **Word:** put one question per paragraph or one question per table row.\n"
        "- **Excel:** put one question per row; the longest non-empty cell in each row is treated as the question.\n"
        "- You can edit imported questions in the tables or add questions manually.\n"
        "- Section A uses only the short-question bank. Sections B and C share the long-question bank.\n"
        "- If you request more questions than are available, the app uses all available questions instead of failing.\n"
        "- A supplied Word template is kept as-is; the generated paper is appended after a page break."
    )
