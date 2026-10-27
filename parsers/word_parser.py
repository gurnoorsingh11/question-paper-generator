import re
from docx import Document

NUMBERED = re.compile(r"^\s*(?:(?:question|q)\s*)?\d+\s*[\).:\-]\s*(.+)$", re.I)
ANSWER = re.compile(r"^\s*(?:answer|ans|solution)\s*[:\-]\s*(.*)$", re.I)
MARKS = re.compile(r"\[(\d+)\s*marks?\]", re.I)

def parse_word(file):
    """Best-effort parser for numbered .docx paragraphs and Question/Answer tables.
    Complex formatting should be checked in the review screen before approval.
    """
    doc = Document(file)
    items = []
    for table in doc.tables:
        rows = [[c.text.strip() for c in r.cells] for r in table.rows]
        if not rows:
            continue
        headers = [x.lower().replace(" ", "") for x in rows[0]]
        q_idx = next((i for i, x in enumerate(headers) if x in {"question", "questions", "questiontext"}), None)
        a_idx = next((i for i, x in enumerate(headers) if x in {"answer", "solution", "answers"}), None)
        m_idx = next((i for i, x in enumerate(headers) if x in {"marks", "mark"}), None)
        data = rows[1:] if q_idx is not None else rows
        q_idx = q_idx if q_idx is not None else 0
        for row in data:
            if q_idx >= len(row) or not row[q_idx]:
                continue
            answer = row[a_idx] if a_idx is not None and a_idx < len(row) else ""
            raw_marks = row[m_idx] if m_idx is not None and m_idx < len(row) else ""
            match = re.search(r"\d+", raw_marks or "")
            items.append({"question_text": row[q_idx], "answer_text": answer,
                          "marks": int(match.group()) if match else None, "source": "Word table"})
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    current_q, current_a = None, []
    numbered_format = any(NUMBERED.match(line) for line in lines)
    def save():
        if current_q:
            text = current_q.strip()
            m = MARKS.search(text)
            marks = int(m.group(1)) if m else None
            text = MARKS.sub("", text).strip()
            items.append({"question_text": text, "answer_text": "\n".join(current_a), "marks": marks, "source": "Word paragraph"})
    for line in lines:
        qm, am = NUMBERED.match(line), ANSWER.match(line)
        if qm:
            save()
            current_q, current_a = qm.group(1).strip(), []
        elif am and current_q:
            if am.group(1).strip():
                current_a.append(am.group(1).strip())
        elif current_q and current_a:
            current_a.append(line)
        elif current_q and numbered_format:
            current_q += "\n" + line
        else:
            save()
            current_q, current_a = line, []
    save()
    # Exact duplicate text is collapsed; review imported items before approval.
    seen, unique = set(), []
    for item in items:
        key = item["question_text"].strip().casefold()
        if key not in seen:
            seen.add(key)
            unique.append(item)
    return unique
