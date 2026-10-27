import pandas as pd

QUESTION_HEADERS = {"question", "questions", "questiontext", "questionstatement", "problem"}
ANSWER_HEADERS = {"answer", "answers", "solution", "solutions", "answertext"}
MARK_HEADERS = {"marks", "mark", "points"}

def parse_excel(file):
    """Parse Excel sheets with a Question column; optional Answer and Marks columns."""
    result = []
    sheets = pd.read_excel(file, sheet_name=None)
    for sheet_name, df in sheets.items():
        if df.empty:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        normalized = {str(c).strip().lower().replace(" ", ""): c for c in df.columns}
        q_col = next((normalized[k] for k in normalized if k in QUESTION_HEADERS), None)
        a_col = next((normalized[k] for k in normalized if k in ANSWER_HEADERS), None)
        m_col = next((normalized[k] for k in normalized if k in MARK_HEADERS), None)
        if q_col is None:
            # Accept a headerless sheet as column A = question, B = answer, C = marks.
            df = pd.read_excel(file, sheet_name=sheet_name, header=None)
            q_col = df.columns[0]
            a_col = df.columns[1] if len(df.columns) > 1 else None
            m_col = df.columns[2] if len(df.columns) > 2 else None
        for _, row in df.iterrows():
            q = row.get(q_col, "")
            if pd.isna(q) or not str(q).strip():
                continue
            ans = row.get(a_col, "") if a_col is not None else ""
            marks = row.get(m_col, None) if m_col is not None else None
            try:
                marks = int(marks) if pd.notna(marks) and str(marks).strip() else None
            except (TypeError, ValueError):
                marks = None
            result.append({
                "question_text": str(q).strip(),
                "answer_text": "" if pd.isna(ans) else str(ans).strip(),
                "marks": marks,
                "source": sheet_name,
            })
    return result
