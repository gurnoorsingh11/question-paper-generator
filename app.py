import streamlit as st
from database.connection import init_db
from pages import question_banks, templates, generate_paper

st.set_page_config(page_title="SVIET Question Paper Generator", page_icon="📄", layout="wide")
init_db()

st.title("SVIET Question Paper Generation System")
st.caption("Manage question banks, configure examination templates, and generate reviewable Word papers.")

page = st.sidebar.radio(
    "Navigation",
    ["Dashboard", "Question Banks", "Templates", "Generate Paper"],
)

if page == "Dashboard":
    st.subheader("Dashboard")
    st.write(
        "This starter application separates saved question banks, reusable exam templates, "
        "and generated papers. Begin by importing a question bank, then classify and approve "
        "questions before generating a paper."
    )
    st.info("Recommended workflow: Question Banks → Templates → Generate Paper → Preview → Download Word.")
    st.markdown("""
    **Important safeguards**
    - Questions must be reviewed and assigned marks before use.
    - Questions with subparts stay together as a single question group.
    - Attempt rules and maximum marks are validated before export.
    - This starter stores data in a local SQLite database (`data/app.db`).
    """)
elif page == "Question Banks":
    question_banks.render()
elif page == "Templates":
    templates.render()
else:
    generate_paper.render()
