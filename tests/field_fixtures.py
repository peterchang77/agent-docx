"""Shared helpers for field-aware tests.

Builds minimal DOCX / unpacked-dir fixtures containing real ``<w:fldChar>``
fields (EndNote-style citations) so the field-aware splitting, validation, and
diffing code paths can be exercised end to end.
"""

import zipfile
from pathlib import Path

import defusedxml.minidom

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""


def _wrap_document(body_xml: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document xmlns:w="{W_NS}">\n'
        f"  <w:body>\n{body_xml}\n  </w:body>\n</w:document>\n"
    )


def field_end_after_text_paragraph() -> str:
    """Paragraph whose run packs a field END into the same run as following text.

    Mirrors the EndNote pattern that broke ``[16-21]``::

        <w:r><w:fldChar end/><w:t>. However, ...</w:t></w:r>

    Used to verify split_run_at keeps the ``fldChar end`` on the left half.
    """
    return _wrap_document(
        "    <w:p>\n"
        '      <w:r><w:fldChar w:fldCharType="end"/><w:t>. However, the model'
        " performed well.</w:t></w:r>\n"
        "    </w:p>"
    )


def field_begin_after_anchor_paragraph() -> str:
    """Paragraph whose run packs field BEGIN+fldData+instrText+separate after a text node.

    Mirrors the EndNote pattern that broke ``[69]``::

        <w:r><w:t>...training data) </w:t>
             <w:fldChar begin><w:fldData>...</w:fldData></w:fldChar>
             <w:instrText> ADDIN EN.JS.CITE </w:instrText>
             <w:fldChar separate/></w:r>

    Used to verify an insert anchored on the leading text leaves the field
    markers intact in one run.
    """
    return _wrap_document(
        "    <w:p>\n"
        '      <w:r><w:t xml:space="preserve">(see training data) </w:t>'
        '<w:fldChar w:fldCharType="begin"><w:fldData>qk==</w:fldData></w:fldChar>'
        '<w:instrText xml:space="preserve"> ADDIN EN.JS.CITE </w:instrText>'
        '<w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>[69]</w:t>'
        '<w:fldChar w:fldCharType="end"/></w:r>\n'
        "    </w:p>"
    )


def break_in_run_paragraph() -> str:
    """Paragraph with a run that mixes text and a <w:br/>."""
    return _wrap_document(
        "    <w:p>\n"
        "      <w:r><w:t>foo</w:t><w:br/><w:t>bar</w:t></w:r>\n"
        "    </w:p>"
    )


def citation_in_paragraph(display: str = "[1]") -> str:
    """A paragraph containing a simple, well-formed citation field.

    The field spans three runs (begin/separate run, display run, end run), the
    normal Word layout.
    """
    return _wrap_document(
        "    <w:p>\n"
        "      <w:r><w:t xml:space=\"preserve\">see </w:t></w:r>"
        '<w:r><w:fldChar w:fldCharType="begin"/>'
        '<w:instrText xml:space="preserve"> ADDIN EN.CITE </w:instrText>'
        '<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:t>{display}</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r><w:t xml:space="preserve"> for details.</w:t></w:r>\n'
        "    </w:p>"
    )


def make_docx(path: Path, document_xml: str) -> Path:
    """Write a minimal .docx (just enough for validators/redlining) to ``path``."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("word/document.xml", document_xml)
    return path


def make_unpacked(path: Path, document_xml: str) -> Path:
    """Write a minimal unpacked dir (word/document.xml) to ``path``."""
    (path / "word").mkdir(parents=True, exist_ok=True)
    (path / "word" / "document.xml").write_text(document_xml, encoding="utf-8")
    return path


def count_fldchars(document_xml: str) -> dict:
    """Return {'begin': n, 'separate': n, 'end': n} for the given XML."""
    dom = defusedxml.minidom.parseString(document_xml)
    counts = {"begin": 0, "separate": 0, "end": 0}
    for elem in dom.getElementsByTagName("w:fldChar"):
        ftype = elem.getAttribute("w:fldCharType")
        if ftype in counts:
            counts[ftype] += 1
    return counts
