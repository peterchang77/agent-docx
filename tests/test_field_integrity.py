"""Tests for the field-character integrity validator (P2).

Covers ``validate_field_chars`` and its wiring into ``DOCXSchemaValidator`` so a
truncated/unbalanced field FAILS LOUDLY instead of packing silently.
"""

import lxml.etree

from agent_docx.validators.field_integrity import (
    validate_field_chars,
    validate_field_counts_against_original,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
BEGIN = f'<w:fldChar w:fldCharType="begin"/>'
SEP = f'<w:fldChar w:fldCharType="separate"/>'
END = f'<w:fldChar w:fldCharType="end"/>'


def _parse(body: str):
    xml = (
        f'<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>'
    )
    return lxml.etree.fromstring(xml)


def test_balanced_field_passes():
    root = _parse(f"<w:p><w:r>{BEGIN}<w:instrText> X </w:instrText>{SEP}</w:r>"
                  f"<w:r><w:t>[1]</w:t></w:r><w:r>{END}</w:r></w:p>")
    assert validate_field_chars(root) == []


def test_validator_flags_orphan_field_end():
    """P2 test 4: an unbalanced (orphaned end) fldChar is reported with a line."""
    # An "end" with no preceding "begin".
    root = _parse(f"<w:p><w:r>{END}</w:r></w:p>")
    errors = validate_field_chars(root)
    assert errors, "expected an orphan-end error"
    assert any("end" in e.lower() and "begin" in e.lower() for e in errors)
    # sourceline is included for an actionable message.
    assert any("Line" in e for e in errors)


def test_validator_flags_unclosed_begin():
    """A begin with no end is reported as a truncated field."""
    root = _parse(f"<w:p><w:r>{BEGIN}<w:instrText> X </w:instrText>{SEP}</w:r></w:p>")
    errors = validate_field_chars(root)
    assert errors
    assert any("never closed" in e for e in errors)


def test_nested_fields_balance_correctly():
    """Nested fields (TOC inside a citation, etc.) must balance via the stack."""
    root = _parse(
        f"<w:p><w:r>{BEGIN}</w:r>"
        f"<w:r>{BEGIN}<w:instrText> inner </w:instrText>{END}</w:r>"
        f"<w:r>{END}</w:r></w:p>"
    )
    assert validate_field_chars(root) == []


def test_instr_count_check_flags_dropped_citation():
    """P2 helper: dropping a citation changes the instrText multiset."""
    original = _parse(
        f"<w:p><w:r>{BEGIN}<w:instrText> ADDIN EN.CITE </w:instrText>{SEP}</w:r>"
        f"<w:r><w:t>[1]</w:t></w:r><w:r>{END}</w:r></w:p>"
    )
    # Edited doc: the whole field is gone (balanced, but instrText lost).
    edited = _parse("<w:p><w:r><w:t>[1]</w:t></w:r></w:p>")
    assert validate_field_counts_against_original(edited, original)


def test_wired_into_docx_schema_validator(tmp_path):
    """The P2 check is wired into pack --check / validate (fails loudly)."""
    from field_fixtures import make_unpacked

    # Unbalanced: begin + separate but no end (the corruption signature).
    body = (
        f'<w:document xmlns:w="{W}"><w:body><w:p>'
        f'<w:r>{BEGIN}<w:instrText> ADDIN X </w:instrText>{SEP}</w:r>'
        f'<w:r><w:t>[1]</w:t></w:r>'
        f"</w:p></w:body></w:document>"
    )
    make_unpacked(tmp_path, body)

    from agent_docx.validators import DOCXSchemaValidator

    validator = DOCXSchemaValidator(tmp_path, verbose=True)
    assert validator.validate_field_integrity() is False
