"""Tests for field-aware run splitting and the field-spanning preflight guard.

Covers P1 (field/complex-content-aware ``split_run_at`` / ``_get_text``) and P3
(preflight + post-edit self-check) from scratch/docx-tool-improvements.md.
"""

from pathlib import Path

import defusedxml.minidom

from agent_docx.core.tracked_changes import (
    _field_fingerprint,
    _get_text,
    split_run_at,
    tracked_edit,
)

from field_fixtures import (
    break_in_run_paragraph,
    count_fldchars,
    field_begin_after_anchor_paragraph,
    field_end_after_text_paragraph,
)


def _write(tmp_path: Path, document_xml: str) -> Path:
    doc = tmp_path / "word" / "document.xml"
    doc.parent.mkdir(parents=True, exist_ok=True)
    doc.write_text(document_xml, encoding="utf-8")
    return doc


# ---------------------------------------------------------------------------
# P1: split_run_at / _get_text preserve non-text children
# ---------------------------------------------------------------------------


def test_get_text_concatenates_all_text_children():
    dom = defusedxml.minidom.parseString(break_in_run_paragraph())
    run = dom.getElementsByTagName("w:r")[0]
    # Run is <w:t>foo</w:t><w:br/><w:t>bar</w:t> -> concatenated text "foobar".
    assert _get_text(run) == "foobar"


def test_split_preserves_break_and_drawing(tmp_path):
    """Splitting a run that mixes <w:t> + <w:br/> must keep the break (P1, test 3).

    A non-text child at the split boundary is routed to one side (never
    dropped). Here the <w:br/> sits between "foo" and "bar", so it travels with
    the right half; the assertion is that it survives somewhere in the result.
    """
    dom = defusedxml.minidom.parseString(break_in_run_paragraph())
    run = dom.getElementsByTagName("w:r")[0]  # <w:t>foo</w:t><w:br/><w:t>bar</w:t>

    # Split between "foo" and "bar" (offset 3 over concatenated text "foobar").
    left, right = split_run_at(run, 3, dom)
    assert left is not None and right is not None
    combined = left.toxml() + right.toxml()
    assert "<w:br/>" in combined, "break must be preserved, not dropped"
    assert "foo" in left.toxml()
    assert "bar" in right.toxml()
    # The break must not be duplicated or lost.
    assert combined.count("<w:br/>") == 1


def test_split_keeps_trailing_field_end_on_left():
    """A <w:fldChar end/> sharing a run with text must stay with the left half."""
    dom = defusedxml.minidom.parseString(field_end_after_text_paragraph())
    run = dom.getElementsByTagName("w:r")[0]

    # Split right after the ". " prefix (offset 2 over ". However, ...").
    left, right = split_run_at(run, 2, dom)
    assert left is not None and right is not None
    assert 'w:fldCharType="end"' in left.toxml(), "field END must stay on the left"
    assert "However" in right.toxml()


# ---------------------------------------------------------------------------
# P1/P3: tracked_edit preserves fields (the two original failure cases)
# ---------------------------------------------------------------------------


def test_edit_preserves_trailing_field_end(tmp_path):
    """P1 test 1: replace text in a run whose BEGINNING is a field END.

    Reproduces the [16-21] corruption: an edit whose anchor did NOT include the
    citation still dropped the field's closing fldChar because it shared a run
    with the replaced text.
    """
    doc = _write(tmp_path, field_end_after_text_paragraph())
    before = count_fldchars(doc.read_text())

    ids, msg = tracked_edit(doc, "However, the model performed well.",
                            "Thus, the model performed well.", mode="replace")
    assert ids, msg
    after = count_fldchars(doc.read_text())

    # begin/separate/end counts must be unchanged -> field END survived.
    assert after == before, f"field counts changed: {before} -> {after}"
    xml = doc.read_text()
    assert 'w:fldCharType="end"' in xml
    assert "Thus, the model performed well." in xml


def test_insert_before_field_begin(tmp_path):
    """P1 test 2: insert anchored on text that shares a run with field BEGIN.

    Reproduces the [69] corruption: an insert collapsed field markers because
    they lived in the same run as the anchor text.
    """
    doc = _write(tmp_path, field_begin_after_anchor_paragraph())
    before = count_fldchars(doc.read_text())
    before_data = doc.read_text()

    ids, msg = tracked_edit(doc, "training data", "training data) ", mode="insert")
    assert ids, msg
    after = count_fldchars(doc.read_text())

    assert after == before, f"field counts changed: {before} -> {after}"
    xml = doc.read_text()
    # The field's begin/fldData/instrText/separate must all survive.
    assert 'w:fldCharType="begin"' in xml
    assert "<w:fldData>qk==</w:fldData>" in xml
    assert "ADDIN EN.JS.CITE" in xml
    assert 'w:fldCharType="separate"' in xml
    assert "[69]" in xml, "display glyph must survive"
    # And nothing was dropped vs the original.
    assert _field_fingerprint(defusedxml.minidom.parseString(xml)) == \
        _field_fingerprint(defusedxml.minidom.parseString(before_data))


# ---------------------------------------------------------------------------
# P3: preflight refuses field-spanning / overlapping edits
# ---------------------------------------------------------------------------


def test_edit_refuses_to_span_across_field(tmp_path):
    """An anchor that visually includes a citation's display text must be refused."""
    from field_fixtures import citation_in_paragraph

    doc = _write(tmp_path, citation_in_paragraph("[1]"))
    original = doc.read_text()

    # "see [1] for" spans across the field's begin/end markers.
    ids, msg = tracked_edit(doc, "see [1] for", "saw [1] for", mode="replace")
    assert not ids, "edit should have been refused"
    assert "Refusing" in msg or "field" in msg.lower()

    # Document must be left untouched.
    assert doc.read_text() == original


def test_edit_refuses_when_run_overlaps_field_marker(tmp_path):
    """A destructive edit over a run that contains field markers is refused."""
    # Build a paragraph where the whole field (begin..display..end) lives in a
    # single run with text, so the matched text overlaps field markers.
    body = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body><w:p>"
        '<w:r><w:t>note </w:t>'
        '<w:fldChar w:fldCharType="begin"/>'
        '<w:instrText xml:space="preserve"> ADDIN X </w:instrText>'
        '<w:fldChar w:fldCharType="separate"/>'
        "<w:t>[5]</w:t>"
        '<w:fldChar w:fldCharType="end"/>'
        "</w:r></w:p></w:body></w:document>"
    )
    doc = _write(tmp_path, body)
    original = doc.read_text()

    ids, msg = tracked_edit(doc, "[5]", "<deleted>", mode="replace")
    assert not ids
    assert "Refusing" in msg or "field" in msg.lower()
    assert doc.read_text() == original


def test_edit_allows_field_edit_with_flag(tmp_path):
    """With allow_field_edit=True, a field-spanning edit proceeds and warns."""
    from field_fixtures import citation_in_paragraph

    doc = _write(tmp_path, citation_in_paragraph("[1]"))

    # "see [1] for" spans the field; refused without the flag, allowed with it.
    ids_refused, _ = tracked_edit(doc, "see [1] for", "saw [1] for", mode="replace")
    assert not ids_refused

    ids, msg = tracked_edit(doc, "see [1] for", "saw [1] for", mode="replace",
                            allow_field_edit=True)
    assert ids, msg
    # Proceeding over a field region appends a warning.
    assert "WARNING" in msg
    # The replacement was applied.
    xml = doc.read_text()
    assert "saw [1] for" in xml
