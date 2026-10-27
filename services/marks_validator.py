def validate_template(sections, declared_maximum):
    errors = []
    offered = 0
    calculated_maximum = 0
    for section in sections:
        name = section.get("name", "Section")
        count = int(section.get("count", 0))
        marks = int(section.get("marks_each", 0))
        attempt = int(section.get("attempt", count))
        minimum = int(section.get("minimum_attempt", 0))
        if count < 0 or marks < 0 or attempt < 0 or minimum < 0:
            errors.append(f"{name}: values cannot be negative.")
            continue
        if attempt > count:
            errors.append(f"{name}: attempts cannot exceed questions offered.")
        if minimum > attempt or minimum > count:
            errors.append(f"{name}: minimum attempts are inconsistent.")
        offered += count
        calculated_maximum += attempt * marks
    if calculated_maximum != int(declared_maximum):
        errors.append(
            f"Declared maximum is {declared_maximum}, but attempt rules calculate "
            f"{calculated_maximum} marks."
        )
    return {"valid": not errors, "errors": errors, "questions_offered": offered,
            "calculated_maximum": calculated_maximum}

def validate_subpart_marks(parent_marks, subpart_marks):
    return sum(int(x) for x in subpart_marks) == int(parent_marks)
