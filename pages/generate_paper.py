import json
import streamlit as st
from database.connection import get_connection
from services.question_bank import list_banks, get_questions
from services.marks_validator import validate_template
from services.paper_generator import generate_paper
from services.document_exporter import export_paper

def render():
    st.header("Generate Paper")
    with get_connection() as conn:
        templates = [dict(r) for r in conn.execute("SELECT * FROM templates ORDER BY id DESC").fetchall()]
    banks = list_banks()
    if not templates:
        st.warning("Create and save an exam template first.")
        return
    if not banks:
        st.warning("Import and classify a question bank first.")
        return
    template_labels = {f"#{t['id']} · {t['name']} ({t['maximum_marks']} marks)": t for t in templates}
    bank_labels = {f"#{b['id']} · {b['name']} ({b['question_count']} questions)": b["id"] for b in banks}
    chosen_template_label = st.selectbox("Exam template", list(template_labels))
    template = template_labels[chosen_template_label]
    chosen_bank_label = st.selectbox("Question bank", list(bank_labels))
    bank_id = bank_labels[chosen_bank_label]
    sections = json.loads(template["sections_json"])
    validation = validate_template(sections, template["maximum_marks"])
    if not validation["valid"]:
        for error in validation["errors"]:
            st.error(error)
        return
    if st.button("Generate / Shuffle Paper", type="primary"):
        questions = get_questions(bank_id)
        try:
            paper = generate_paper(questions, sections)
            st.session_state["paper_preview"] = {
                "template": template, "sections": sections,
                "paper": paper, "bank_id": bank_id
            }
        except ValueError as e:
            st.error(str(e))
    preview = st.session_state.get("paper_preview")
    if not preview or preview["template"]["id"] != template["id"] or preview["bank_id"] != bank_id:
        return
    st.divider()
    st.subheader("Preview")
    st.caption("Review this paper before downloading. Shuffle generates a new selection.")
    t = preview["template"]
    st.markdown(f"### {t['department']}")
    st.markdown(f"**{t['subject_name']}** · Subject code: {t['subject_code']} · Semester: {t['semester']}")
    st.write(f"M. Code: {t['m_code']} · Duration: {t['duration_minutes']} minutes · Maximum marks: {t['maximum_marks']}")
    st.markdown("**Instructions to candidates**")
    st.text(t["instructions"])
    qnum = 1
    for section in preview["paper"]:
        st.markdown(f"#### SECTION-{section['section']}")
        for q in section["questions"]:
            st.markdown(f"**{qnum}.** {q['question_text']} **[{section['marks_each']} marks]**")
            if q.get("answer_text"):
                with st.expander(f"Reveal answer for Q{qnum}"):
                    st.write(q["answer_text"])
            qnum += 1
    metadata = {
        "department": t["department"], "subject_name": t["subject_name"],
        "subject_code": t["subject_code"], "semester": t["semester"],
        "programme_semester": f"B. Tech. · {t['department']} · Semester {t['semester']}",
        "m_code": t["m_code"], "exam_date": t["exam_date"],
        "duration": t["duration_minutes"], "maximum_marks": t["maximum_marks"],
    }
    docx = export_paper(metadata, preview["sections"], preview["paper"], t["instructions"])
    st.download_button("Download Word paper", docx, file_name="sviet_question_paper.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
