import streamlit as st
import pandas as pd
from parsers.excel_parser import parse_excel
from parsers.word_parser import parse_word
from services.question_bank import create_bank, list_banks, get_questions, update_question

def render():
    st.header("Question Banks")
    with st.expander("Import a new question bank", expanded=True):
        name = st.text_input("Question bank name", placeholder="e.g. Engineering Physics Unit 1")
        uploaded = st.file_uploader("Upload Excel or Word file", type=["xlsx", "xls", "docx"], key="bank_upload")
        if st.button("Import questions", type="primary"):
            if not name.strip() or not uploaded:
                st.error("Enter a bank name and choose a file.")
            else:
                try:
                    parsed = parse_excel(uploaded) if uploaded.name.lower().endswith((".xlsx", ".xls")) else parse_word(uploaded)
                    if not parsed:
                        st.error("No questions detected. Use a Question column in Excel or numbered paragraphs / question tables in Word.")
                    else:
                        bank_id = create_bank(name.strip(), uploaded.name, parsed)
                        st.success(f"Imported {len(parsed)} questions into bank #{bank_id}. All questions start unapproved.")
                        st.rerun()
                except Exception as e:
                    st.error(f"Import failed: {e}")
    banks = list_banks()
    if not banks:
        st.info("No saved question banks yet.")
        return
    labels = {f"#{b['id']} · {b['name']} ({b['question_count']} questions)": b["id"] for b in banks}
    selected_label = st.selectbox("Choose a question bank", list(labels))
    bank_id = labels[selected_label]
    questions = get_questions(bank_id)
    st.caption("Review each imported question. Assign marks, select eligible sections, and approve it before generation.")
    for q in questions:
        with st.expander(f"Question #{q['id']} · {q['question_text'][:100]}"):
            with st.form(f"edit_{q['id']}"):
                qtext = st.text_area("Question text", q["question_text"])
                answer = st.text_area("Answer / solution (optional)", q["answer_text"])
                marks_options = [0, 1, 2, 3, 4, 5, 6, 8, 10, 12, 15, 20]
                current_marks = q["marks"] if q["marks"] in marks_options else 0
                marks = st.selectbox("Marks (0 = unclassified)", marks_options, index=marks_options.index(current_marks))
                current_sections = [x for x in q["eligible_sections"].split(",") if x in ["A", "B", "C"]]
                sections = st.multiselect("Eligible sections", ["A", "B", "C"], default=current_sections)
                approved = st.checkbox("Approved for paper generation", value=bool(q["approved"]))
                if st.form_submit_button("Save question"):
                    if not qtext.strip():
                        st.error("Question text is required.")
                    elif marks == 0 and approved:
                        st.error("Assign marks before approving.")
                    elif not sections and approved:
                        st.error("Select at least one eligible section.")
                    else:
                        update_question(q["id"], qtext.strip(), answer, None if marks == 0 else marks, sections, approved)
                        st.success("Saved.")
                        st.rerun()
