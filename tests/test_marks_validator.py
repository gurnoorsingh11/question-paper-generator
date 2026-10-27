from services.marks_validator import validate_template, validate_subpart_marks

def test_valid_choice_pattern():
    sections = [
        {"name":"A", "count":10, "marks_each":2, "attempt":10, "minimum_attempt":10},
        {"name":"B", "count":4, "marks_each":8, "attempt":3, "minimum_attempt":2},
        {"name":"C", "count":4, "marks_each":8, "attempt":2, "minimum_attempt":2},
    ]
    result = validate_template(sections, 60)
    assert result["valid"]

def test_wrong_maximum_rejected():
    sections = [{"name":"A", "count":5, "marks_each":2, "attempt":5, "minimum_attempt":5}]
    assert not validate_template(sections, 20)["valid"]

def test_subpart_sum():
    assert validate_subpart_marks(8, [3, 3, 2])
    assert not validate_subpart_marks(8, [3, 3, 3])
