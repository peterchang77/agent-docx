"""Tests for the field-aware --diff / redlining fingerprint (P4).

The text-only diff is blind to field structure: dropping a citation's field
wrapper while keeping its display glyph (e.g. "[1]") yields IDENTICAL extracted
text. These tests verify the field fingerprint comparison catches that loss.
"""

from pathlib import Path

from agent_docx.validators import RedliningValidator

from field_fixtures import citation_in_paragraph, make_docx, make_unpacked


def _original_with_field(tmp_path: Path) -> Path:
    return make_docx(tmp_path / "original.docx", citation_in_paragraph("[1]"))


def _edited_field_dropped(tmp_path: Path) -> Path:
    """Same display text as the original, but the field wrapper is gone."""
    from field_fixtures import _wrap_document

    body = _wrap_document(
        '    <w:p>\n'
        '      <w:r><w:t xml:space="preserve">see </w:t></w:r>'
        '<w:r><w:t>[1]</w:t></w:r>'
        '<w:r><w:t xml:space="preserve"> for details.</w:t></w:r>\n'
        "    </w:p>"
    )
    return make_unpacked(tmp_path / "edited", body)


def test_diff_flags_dropped_citation(tmp_path):
    """P4 test 5: dropping a field wrapper but keeping display text is reported."""
    original = _original_with_field(tmp_path)
    edited = _edited_field_dropped(tmp_path)

    validator = RedliningValidator(edited, original, author="Peter")
    ok = validator.validate()

    # The display text is identical, so the OLD validator would have passed.
    # The field-aware check must now FAIL.
    assert ok is False

    # And specifically the fingerprint changed.
    import xml.etree.ElementTree as ET

    mod_root = ET.parse(edited / "word" / "document.xml").getroot()
    import tempfile, zipfile

    with tempfile.TemporaryDirectory() as td:
        with zipfile.ZipFile(original) as z:
            z.extractall(td)
        orig_root = ET.parse(Path(td) / "word" / "document.xml").getroot()
    mod_fp = validator._extract_field_fingerprint(mod_root)
    orig_fp = validator._extract_field_fingerprint(orig_root)
    assert mod_fp != orig_fp
    # Original had one citation (begin/separate/end + one instrText).
    assert orig_fp[0] == 1 and orig_fp[2] == 1 and len(orig_fp[3]) == 1
    # Edited lost it entirely.
    assert mod_fp[0] == 0 and len(mod_fp[3]) == 0


def test_pack_diff_flags_dropped_citation(tmp_path):
    """The pack --diff path also reports the field-count change."""
    import subprocess
    import sys

    original = _original_with_field(tmp_path)
    edited = _edited_field_dropped(tmp_path)

    result = subprocess.run(
        [sys.executable, "-m", "agent_docx.cli",
         "pack", str(edited), "--diff", "--original", str(original)],
        capture_output=True, text=True,
    )
    assert "No untracked text differences" in result.stdout, result.stdout
    assert "Field/citation fingerprint changed" in result.stdout, result.stdout
    assert result.returncode != 0


def test_matching_fields_pass(tmp_path):
    """When fields are intact, the field-aware check passes."""
    original = _original_with_field(tmp_path)
    # Edited == exact copy of the original (no changes at all).
    edited = make_unpacked(tmp_path / "edited", citation_in_paragraph("[1]"))

    validator = RedliningValidator(edited, original, author="Peter")
    assert validator.validate() is True
