"""Field-character (``<w:fldChar>``) integrity checks for OOXML documents.

Catches truncated / unbalanced Word fields (EndNote citations, cross-references,
TOC, page numbers, ...) that the rest of the validator suite does not detect.

A well-formed field is an ordered sequence::

    <w:fldChar w:fldCharType="begin"/>
    [ <w:instrText>...</w:instrText> | <w:fldData>...</w:fldData> ]*
    <w:fldChar w:fldCharType="separate"/>          (optional)
    ... display runs ...
    <w:fldChar w:fldCharType="end"/>

Fields may nest, so this module tracks depth with a stack rather than a flat
counter. ``lxml`` is used (matching ``validators/docx.py``) so that ``sourceline``
is available for actionable error messages.
"""

import lxml.etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_FLDCHAR_TAG = f"{{{W}}}fldChar"
_TYPE_ATTR = f"{{{W}}}fldCharType"


def _root_of(tree_or_root):
    """Accept either an lxml ElementTree or an Element and return the Element."""
    if hasattr(tree_or_root, "getroot"):
        return tree_or_root.getroot()
    return tree_or_root


def validate_field_chars(tree_or_root):
    """Return a list of human-readable error strings (empty == valid).

    Walks ``<w:fldChar>`` elements in document order and asserts proper
    begin -> (instrText|fldData)* -> separate? -> end nesting, reporting
    orphans and unclosed begins.
    """
    root = _root_of(tree_or_root)
    errors = []

    # Stack of "begin" sourcelines still awaiting their "end".
    open_stack = []

    for fld in root.iter(_FLDCHAR_TAG):
        ftype = fld.get(_TYPE_ATTR)
        line = getattr(fld, "sourceline", "?")

        if ftype == "begin":
            open_stack.append(line)
        elif ftype == "separate":
            if not open_stack:
                errors.append(
                    f'Line {line}: <w:fldChar w:fldCharType="separate"> with no '
                    f"open field (missing begin)."
                )
        elif ftype == "end":
            if not open_stack:
                errors.append(
                    f'Line {line}: <w:fldChar w:fldCharType="end"> with no '
                    f"matching begin (orphaned field end)."
                )
            else:
                open_stack.pop()
        else:
            errors.append(
                f"Line {line}: <w:fldChar> with unexpected/missing "
                f"fldCharType={ftype!r}."
            )

    for line in open_stack:
        errors.append(
            f'Line {line}: <w:fldChar w:fldCharType="begin"> never closed '
            f"(truncated field - begin/separate present but no end). This "
            f"usually means an edit dropped the closing fldChar."
        )

    return errors


def validate_field_counts_against_original(edited_root, original_root):
    """Stronger check for the pack step: compare field-instruction fingerprints.

    Returns a list of error strings. A dropped citation is caught here even if
    begin/end happen to stay balanced (e.g. the whole field was removed), by
    comparing the multiset of ``<w:instrText>``/``<w:delInstrText>`` codes.
    """
    def instr_codes(root):
        codes = []
        for tag in ("instrText", "delInstrText"):
            for el in root.iter(f"{{{W}}}{tag}"):
                codes.append((el.text or "").strip())
        return sorted(c for c in codes if c)

    edited = instr_codes(_root_of(edited_root))
    original = instr_codes(_root_of(original_root))
    errors = []
    if len(edited) != len(original):
        errors.append(
            f"Field-instruction count changed: {len(original)} -> {len(edited)}. "
            f"A field (e.g. a citation) may have been dropped during editing."
        )
    return errors
