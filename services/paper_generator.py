import random
from collections import defaultdict

def generate_paper(questions, sections, seed=None):
    """Select approved question records without duplicates across sections.
    This first version treats each record as a whole question group.
    """
    rng = random.Random(seed)
    remaining = list(questions)
    result = []
    for section in sections:
        name = section["name"]
        count = int(section["count"])
        marks = int(section["marks_each"])
        eligible = [
            q for q in remaining
            if q.get("approved")
            and q.get("marks") == marks
            and name[-1].upper() in (q.get("eligible_sections") or "A,B,C").split(",")
        ]
        if len(eligible) < count:
            raise ValueError(
                f"{name} needs {count} approved {marks}-mark questions, "
                f"but only {len(eligible)} are eligible."
            )
        selected = rng.sample(eligible, count)
        for q in selected:
            remaining.remove(q)
        result.append({"section": name, "marks_each": marks, "questions": selected})
    return result
