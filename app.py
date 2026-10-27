import io
import random
import re

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt

st.set_page_config(
    page_title="SVIET Question Paper Generator",
    page_icon="📄",
    layout="centered",
)

st.title("SVIET Question Paper Generator")
st.write(
    "Upload a separate 2-mark question bank and a long-question bank "
    "for Sections B/C. Select random questions, shuffle sections, "
    "and download a printable Word paper."
)


# ---------- QUESTION BANK READERS ----------

def normalize(text):
    return re.sub(r"\s+", " ", str(text)).strip()


def looks_like_heading(text):
    value = normalize(text).strip(" :-")
    return bool(
        not value
        or re.match(
            r"^(unit\s+[ivx\d]+|module\s+\d+|chapter\s+\d+|"
            r"question\s*(bank|text)?|questions?|s\.?\s*no\.?|serial\s*no\.?)$",
            value,
            flags=re.IGNORECASE,
        )
    )


def unique(items):
    result = []
    seen = set()

    for item in items:
        text = normalize(item)

        if len(text) < 3 or looks_like_heading(text):
            continue

        if text.lower() in {
            "question", "question text", "s.no", "s.no.", "serial no"
        }:
            continue

        key = text.casefold()

        if key not in seen:
            seen.add(key)
            result.append(text)

    return result


def read_docx(uploaded_file):
    uploaded_file.seek(0)
    doc = Document(uploaded_file)
    items = []

    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            items.append(paragraph.text.strip())

    for table in doc.tables:
        for row in table.rows:
            cells = [
                cell.text.strip()
                for cell in row.cells
                if cell.text.strip()
            ]

            if cells:
                items.append(max(cells, key=len))

    return unique(items)


def read_excel(uploaded_file):
    uploaded_file.seek(0)
    workbook = pd.ExcelFile(uploaded_file)
    items = []

    for sheet_name in workbook.sheet_names:
        frame = pd.read_excel(
            workbook,
            sheet_name=sheet_name,
            header=None,
        ).fillna("")

        for _, row in frame.iterrows():
            cells = [
                normalize(value)
                for value in row.tolist()
                if normalize(value)
            ]

            if cells:
                items.append(max(cells, key=len))

    return unique(items)


def read_bank(uploaded_file):
    name = uploaded_file.name.lower()

    if name.endswith(".docx"):
        return read_docx(uploaded_file)

    if name.endswith((".xlsx", ".xls")):
        return read_excel(uploaded_file)

    raise ValueError(
        "Upload a Word (.docx) or Excel (.xlsx/.xls) file."
    )


# ---------- WORD PAPER CREATION ----------

def create_default_template(meta):
    doc = Document()
    section = doc.sections[0]

    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    p = doc.add_paragraph()
    p.alignment = 1
    run = p.add_run(meta["college"])
    run.bold = True
    run.font.size = Pt(15)

    p = doc.add_paragraph()
    p.alignment = 1
    run = p.add_run("QUESTION PAPER")
    run.bold = True
    run.font.size = Pt(13)

    doc.add_paragraph(
        f"Department: {meta['department']}     "
        f"Semester: {meta['semester']}"
    )

    doc.add_paragraph(
        f"Subject: {meta['subject']}     "
        f"Subject Code: {meta['subject_code']}"
    )

    doc.add_paragraph(
        f"Exam Duration: {meta['duration']}     "
        f"Maximum Marks: {meta['maximum_marks']}"
    )

    if meta["m_code"]:
        doc.add_paragraph(f"M. Code: {meta['m_code']}")

    doc.add_paragraph(
        "Instructions: Attempt questions as directed."
    )
    doc.add_paragraph("")

    return doc


def append_section(doc, title, questions, start_number, marks_label=None):
    if not questions:
        return start_number

    heading = doc.add_paragraph()
    heading.style = "Heading 2"
    heading.add_run(title)

    number = start_number

    for question in questions:
        label = f"{number}. {question}"

        if marks_label:
            label += f" ({marks_label})"

        doc.add_paragraph(label)
        number += 1

    return number


def build_docx(
    meta,
    section_a,
    section_b,
    section_c,
    template_bytes=None,
):
    if template_bytes:
        doc = Document(io.BytesIO(template_bytes))
        doc.add_page_break()
    else:
        doc = create_default_template(meta)

    next_number = 1

    next_number = append_section(
        doc,
        "SECTION A — SHORT QUESTIONS",
        section_a,
        next_number,
        "2 marks each",
    )

    next_number = append_section(
        doc,
        "SECTION B — LONG QUESTIONS",
        section_b,
        next_number,
    )

    append_section(
        doc,
        "SECTION C — LONG QUESTIONS",
        section_c,
        next_number,
    )

    output = io.BytesIO()
    doc.save(output)

    return output.getvalue()


# ---------- EXAM DETAILS ----------

with st.expander("1. Exam details", expanded=True):
    col1, col2 = st.columns(2)

    with col1:
        college = st.text_input(
            "College / institute",
            "Swami Vivekanand Institute of Engineering and Technology",
        )

        department = st.text_input(
            "Department",
            "Computer Science & Engineering",
        )

        subject = st.text_input("Subject name")
        subject_code = st.text_input("Subject code")

    with col2:
        semester = st.text_input("Semester")
        m_code = st.text_input("M. Code (optional)")
        duration = st.text_input("Exam duration", "3 Hours")

        maximum_marks = st.number_input(
            "Maximum marks",
            min_value=1,
            max_value=300,
            value=70,
            step=1,
        )


# ---------- TEMPLATE UPLOAD ----------

st.subheader("2. Upload exam layout template (optional)")

template_upload = st.file_uploader(
    "Word template (.docx)",
    type=["docx"],
    key="layout_template",
    help=(
        "Upload your desired exam layout. Existing content is preserved; "
        "generated questions are added after a page break."
    ),
)

if template_upload is not None:
    st.success(f"Template selected: {template_upload.name}")
    template_bytes = template_upload.getvalue()
else:
    template_bytes = None


# ---------- TWO QUESTION BANKS ----------

st.subheader("3. Upload question banks")

col_a, col_b = st.columns(2)

with col_a:
    short_upload = st.file_uploader(
        "Section A — 2-mark questions",
        type=["docx", "xlsx", "xls"],
        key="short_bank",
        help="Upload a separate file containing short-answer questions.",
    )

with col_b:
    long_upload = st.file_uploader(
        "Sections B/C — long questions",
        type=["docx", "xlsx", "xls"],
        key="long_bank",
        help="Upload a separate file containing long-answer questions.",
    )


def load_if_changed(upload, state_key, display_name):
    if upload is None:
        st.session_state.pop(state_key, None)
        st.session_state.pop(state_key + "_fingerprint", None)
        return []

    fingerprint = f"{upload.name}:{upload.size}"

    if st.session_state.get(state_key + "_fingerprint") != fingerprint:
        try:
            st.session_state[state_key] = read_bank(upload)
            st.session_state[state_key + "_fingerprint"] = fingerprint
            st.session_state.pop("selected_paper", None)

        except Exception as exc:
            st.error(f"Couldn't read {display_name}: {exc}")
            return []

    return st.session_state.get(state_key, [])


short_questions = load_if_changed(
    short_upload,
    "short_questions",
    "the 2-mark bank",
)

long_questions = load_if_changed(
    long_upload,
    "long_questions",
    "the long-question bank",
)

if short_upload is not None:
    st.caption(
        f"Section A bank: {len(short_questions)} questions loaded."
    )

if long_upload is not None:
    st.caption(
        f"Long-question bank: {len(long_questions)} questions loaded."
    )


# ---------- QUESTION COUNTS ----------

st.subheader("4. Choose how many questions to print")

c1, c2, c3 = st.columns(3)

with c1:
    count_a = st.number_input(
        "Section A (2 marks each)",
        min_value=0,
        max_value=200,
        value=10,
        step=1,
    )

with c2:
    count_b = st.number_input(
        "Section B long questions",
        min_value=0,
        max_value=200,
        value=5,
        step=1,
    )

with c3:
    count_c = st.number_input(
        "Section C long questions",
        min_value=0,
        max_value=200,
        value=3,
        step=1,
    )

st.caption(
    "No 2-, 5-, or 10-mark availability checks are performed. "
    "Questions are selected from the two uploaded banks. If you request "
    "more questions than a bank contains, all available questions from "
    "that bank will be used."
)

if int(count_a) > len(short_questions):
    st.info(
        f"Section A requests {int(count_a)} questions, but the bank "
        f"contains {len(short_questions)}. All available short questions "
        "will be used."
    )

if int(count_b) + int(count_c) > len(long_questions):
    st.info(
        f"Sections B and C request {int(count_b) + int(count_c)} long "
        f"questions, but the bank contains {len(long_questions)}. "
        "Available questions will be used without duplicates between B and C."
    )


# ---------- RANDOM SELECTION ----------

def choose_paper():
    selected_a = random.sample(
        short_questions,
        min(int(count_a), len(short_questions)),
    )

    available_long = long_questions.copy()
    random.shuffle(available_long)

    b_count = min(int(count_b), len(available_long))
    selected_b = available_long[:b_count]

    remaining_long = available_long[b_count:]

    c_count = min(int(count_c), len(remaining_long))
    selected_c = remaining_long[:c_count]

    st.session_state["selected_paper"] = {
        "A": selected_a,
        "B": selected_b,
        "C": selected_c,
    }


action1, action2 = st.columns(2)

with action1:
    if st.button(
        "🎲 Generate Random Paper",
        type="primary",
        use_container_width=True,
    ):
        if not short_questions and not long_questions:
            st.error("Upload at least one question bank first.")
        else:
            choose_paper()
            st.rerun()

with action2:
    if st.button(
        "🔀 Shuffle All Sections",
        use_container_width=True,
    ):
        if not short_questions and not long_questions:
            st.error("Upload at least one question bank first.")
        else:
            choose_paper()
            st.rerun()


# ---------- PREVIEW AND SECTION SHUFFLE ----------

selected = st.session_state.get("selected_paper")

if selected is not None:
    st.subheader("5. Preview exam paper")

    st.markdown(
        f"**{subject or 'Subject'}**  \n"
        f"Code: {subject_code or '—'} | "
        f"Semester: {semester or '—'}  \n"
        f"Duration: {duration} | "
        f"Maximum marks: {maximum_marks}"
    )

    if m_code.strip():
        st.caption(f"M. Code: {m_code}")

    section_specs = [
        (
            "A",
            "SECTION A — SHORT QUESTIONS (2 MARKS EACH)",
            int(count_a),
        ),
        (
            "B",
            "SECTION B — LONG QUESTIONS",
            int(count_b),
        ),
        (
            "C",
            "SECTION C — LONG QUESTIONS",
            int(count_c),
        ),
    ]

    for section_key, heading, requested in section_specs:
        questions = selected.get(section_key, [])

        heading_col, shuffle_col = st.columns([4, 1])

        with heading_col:
            st.markdown(f"### {heading}")

        with shuffle_col:
            if st.button(
                "Shuffle",
                key=f"shuffle_{section_key}",
                use_container_width=True,
            ):
                if section_key == "A":
                    selected["A"] = random.sample(
                        short_questions,
                        min(int(count_a), len(short_questions)),
                    )

                else:
                    other_key = "C" if section_key == "B" else "B"
                    reserved = set(selected.get(other_key, []))

                    pool = [
                        question
                        for question in long_questions
                        if question not in reserved
                    ]

                    selected[section_key] = random.sample(
                        pool,
                        min(requested, len(pool)),
                    )

                st.session_state["selected_paper"] = selected
                st.rerun()

        if not questions:
            st.caption("No questions selected for this section.")

        else:
            for number, question in enumerate(questions, start=1):
                st.write(f"**{number}.** {question}")

    meta = {
        "college": college,
        "department": department,
        "subject": subject,
        "subject_code": subject_code,
        "semester": semester,
        "duration": duration,
        "maximum_marks": int(maximum_marks),
        "m_code": m_code,
    }

    try:
        docx_data = build_docx(
            meta,
            selected.get("A", []),
            selected.get("B", []),
            selected.get("C", []),
            template_bytes=template_bytes,
        )

        filename_subject = re.sub(
            r"[^A-Za-z0-9_-]+",
            "_",
            subject,
        ).strip("_")

        if not filename_subject:
            filename_subject = "SVIET_Exam"

        st.download_button(
            "⬇️ Download Word Exam Paper",
            data=docx_data,
            file_name=f"{filename_subject}_Question_Paper.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

        st.caption(
            "Open the downloaded Word document and choose Print."
        )

    except Exception as exc:
        st.error(f"Could not create the Word paper: {exc}")


st.divider()

st.caption(
    "Simple workflow: one 2-mark bank for Section A, one long-question "
    "bank for Sections B/C, optional Word layout template, random selection, "
    "shuffle, preview, and Word download."
)