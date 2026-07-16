"""Core logic for inserting tracked changes into DOCX XML.

Finds target text across runs (even spanning multiple <w:r> elements),
splits runs at match boundaries, and wraps in <w:del>/<w:ins> markup.
"""

from datetime import datetime, timezone
from pathlib import Path

import defusedxml.minidom

WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def get_max_id(dom) -> int:
    """Scan all elements for w:id attributes, return the max integer value."""
    max_id = 0
    for elem in dom.getElementsByTagName("*"):
        for attr_name in ("w:id",):
            val = elem.getAttribute(attr_name)
            if val:
                try:
                    max_id = max(max_id, int(val))
                except ValueError:
                    pass
    return max_id


def _get_child(parent, local_name):
    for child in parent.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            name = child.localName or child.tagName
            if name == local_name or name.endswith(f":{local_name}"):
                return child
    return None


# Text-bearing run children whose lengths count toward the character offset.
_TEXT_TAGS = ("t", "delText")

# Run children that constitute Word "field" structure (EndNote citations, cross
# references, TOC, page numbers, ...). These must never be silently dropped by
# run splitting or wrapped in <w:del> by a text edit.
_FIELD_TAGS = ("fldChar", "fldData", "instrText", "delInstrText")


def _local(node):
    """Local name of an element, stripped of any namespace prefix."""
    name = node.localName or node.tagName
    return name.split(":")[-1]


def _run_content_children(run):
    """Ordered element children of a run, excluding <w:rPr>."""
    out = []
    for child in run.childNodes:
        if child.nodeType != child.ELEMENT_NODE:
            continue
        if _local(child) == "rPr":
            continue
        out.append(child)
    return out


def _node_text(node):
    """Text of a single <w:t>/<w:delText> element (empty if none)."""
    if node.firstChild and node.firstChild.nodeValue:
        return node.firstChild.nodeValue
    return ""


def _get_text(run):
    """Concatenate ALL <w:t>/<w:delText> text in the run, in document order.

    Previously this returned only the first text child, which (a) made offset
    math wrong for runs containing multiple text nodes or interleaved fields,
    and (b) hid field characters that Word routinely packs into the same run
    as adjacent text. Concatenating every text child keeps
    ``find_text_in_paragraph`` offsets correct for the field-aware splitter.
    """
    parts = []
    for child in _run_content_children(run):
        if _local(child) in _TEXT_TAGS:
            parts.append(_node_text(child))
    return "".join(parts)


def _run_text_length(run):
    """Total character length of all text children in the run."""
    return sum(
        len(_node_text(c))
        for c in _run_content_children(run)
        if _local(c) in _TEXT_TAGS
    )


def _run_has_field_markers(run):
    """True if the run carries any Word field structure (fldChar/fldData/instrText)."""
    return any(_local(c) in _FIELD_TAGS for c in _run_content_children(run))


def _is_run(node):
    if node.nodeType != node.ELEMENT_NODE:
        return False
    name = node.localName or node.tagName
    return name == "r" or name.endswith(":r")


def _is_in_tracked_change(run):
    """Check if a run is inside a w:ins or w:del element."""
    parent = run.parentNode
    if parent:
        name = parent.localName or parent.tagName
        if name in ("ins", "del") or name.endswith(":ins") or name.endswith(":del"):
            return True
    return False


def _collect_runs(paragraph):
    """Collect direct-child <w:r> elements and runs inside <w:ins>/<w:del>."""
    runs = []
    for child in paragraph.childNodes:
        if child.nodeType != child.ELEMENT_NODE:
            continue
        if _is_run(child):
            runs.append(child)
        else:
            name = child.localName or child.tagName
            if name in ("ins", "del") or name.endswith(":ins") or name.endswith(":del"):
                for grandchild in child.childNodes:
                    if grandchild.nodeType == grandchild.ELEMENT_NODE and _is_run(grandchild):
                        runs.append(grandchild)
    return runs


def find_text_in_paragraph(paragraph, target):
    """Find target text across runs in a paragraph.

    Returns list of (run, char_start_in_run, char_end_in_run) tuples,
    or None if not found.
    """
    runs = _collect_runs(paragraph)
    text_runs = []  # (run, text, global_start)
    pos = 0
    for run in runs:
        text = _get_text(run)
        if text:
            text_runs.append((run, text, pos))
            pos += len(text)

    if not text_runs:
        return None

    full_text = "".join(t for _, t, _ in text_runs)
    idx = full_text.find(target)
    if idx == -1:
        return None

    match_start = idx
    match_end = idx + len(target)

    matched = []
    for run, text, global_start in text_runs:
        run_end = global_start + len(text)
        if run_end <= match_start or global_start >= match_end:
            continue
        local_start = max(0, match_start - global_start)
        local_end = min(len(text), match_end - global_start)
        matched.append((run, local_start, local_end))

    return matched


def _clone_rpr(run, dom):
    """Clone the <w:rPr> from a run, or return None."""
    rpr = _get_child(run, "rPr")
    if rpr:
        return rpr.cloneNode(True)
    return None


def _make_t_element(dom, tag_name, text, ns_prefix="w"):
    """Create a <w:t> or <w:delText> element with proper xml:space."""
    t = dom.createElement(f"{ns_prefix}:{tag_name}")
    t.appendChild(dom.createTextNode(text))
    if text and (text[0] in " \t" or text[-1] in " \t"):
        t.setAttribute("xml:space", "preserve")
    return t


def _make_run(dom, rpr_node, t_element, ns_prefix="w"):
    """Create a <w:r> with optional <w:rPr> and a text element."""
    r = dom.createElement(f"{ns_prefix}:r")
    if rpr_node:
        r.appendChild(rpr_node)
    r.appendChild(t_element)
    return r


def split_run_at(run, offset, dom):
    """Split a <w:r> into two runs at a character offset.

    Returns (left_run, right_run). Either may be None if the offset lands on a
    boundary, or if that side ends up with no content children.

    The offset is measured over the run's concatenated text content (all
    <w:t>/<w:delText> children in document order). Non-text children — field
    characters (<w:fldChar>), field data (<w:fldData>), instructions
    (<w:instrText>), breaks (<w:br>), tabs, drawings, etc. — are **preserved**
    and routed to the correct side of the split based on their position. This
    fixes the EndNote-style corruption where a field marker shares a run with
    adjacent text and the old splitter discarded it.
    """
    text_len = _run_text_length(run)
    if offset <= 0:
        return None, run
    if offset >= text_len:
        return run, None

    ns_prefix = run.prefix or "w"
    left = dom.createElement(f"{ns_prefix}:r")
    right = dom.createElement(f"{ns_prefix}:r")

    rpr = _get_child(run, "rPr")
    if rpr is not None:
        left.appendChild(rpr.cloneNode(True))
        right.appendChild(rpr.cloneNode(True))

    consumed = 0
    split_done = False

    for child in _run_content_children(run):
        tag = _local(child)

        if split_done:
            # Everything after the split point goes right, unchanged.
            right.appendChild(child.cloneNode(True))
            continue

        if tag not in _TEXT_TAGS:
            # Non-text child before the split point stays on the left. This is
            # what keeps <w:fldChar w:fldCharType="end"/> attached to the text
            # that precedes it instead of being dropped.
            left.appendChild(child.cloneNode(True))
            continue

        # Text-bearing child.
        txt = _node_text(child)
        seg_len = len(txt)

        if consumed + seg_len <= offset:
            # Whole text segment is on the left.
            left.appendChild(child.cloneNode(True))
            consumed += seg_len
            if consumed == offset:
                split_done = True
        else:
            # The split falls inside this text segment — divide it.
            cut = offset - consumed
            left_txt, right_txt = txt[:cut], txt[cut:]
            if left_txt:
                left.appendChild(_make_t_element(dom, tag, left_txt, ns_prefix))
            if right_txt:
                right.appendChild(_make_t_element(dom, tag, right_txt, ns_prefix))
            consumed += seg_len
            split_done = True

    def _has_content(r):
        return any(
            _local(c) != "rPr"
            for c in r.childNodes
            if c.nodeType == c.ELEMENT_NODE
        )

    left = left if _has_content(left) else None
    right = right if _has_content(right) else None
    return left, right


def _find_paragraphs(dom):
    results = []
    for elem in dom.getElementsByTagName("*"):
        name = elem.localName or elem.tagName
        if name == "p" or name.endswith(":p"):
            results.append(elem)
    return results


def _field_fingerprint(dom):
    """Return a tuple capturing Word field structure across the document.

    ``(begin, separate, end, fldData, tuple(sorted instrText codes))``.

    Two documents with the same fingerprint have the same set of fields
    (EndNote citations, cross-references, TOC, page numbers, ...). A text edit
    that leaves this unchanged cannot have silently dropped or damaged a
    field. Used as a before/after self-check inside :func:`tracked_edit`.
    """
    counts = {"begin": 0, "separate": 0, "end": 0, "fldData": 0}
    instr = []
    for elem in dom.getElementsByTagName("*"):
        tag = _local(elem)
        if tag == "fldChar":
            ftype = (
                elem.getAttribute("w:fldCharType")
                or elem.getAttribute("fldCharType")
            )
            if ftype in counts:
                counts[ftype] += 1
        elif tag == "fldData":
            counts["fldData"] += 1
        elif tag in ("instrText", "delInstrText"):
            instr.append(_node_text(elem).strip())
    return (
        counts["begin"],
        counts["separate"],
        counts["end"],
        counts["fldData"],
        tuple(sorted(i for i in instr if i)),
    )


def _detect_field_conflict(para, matched_runs, isolated_runs, mode):
    """Return an error message if a text edit would collide with a Word field.

    Two cases are detected:
      1. An isolated run still carries field markers — deleting/wrapping it
         (delete/replace modes) would drop the markers.
      2. The match spans across a field: a field-marker run sits strictly
         between the first and last matched run in the paragraph, so the
         anchor visually includes a field's display text while its begin/end
         markers fall outside the match.

    Returns ``None`` when there is no conflict.
    """
    # Case 1 only matters for destructive modes (insert never wraps runs).
    if mode in ("delete", "replace"):
        for run in isolated_runs:
            if _run_has_field_markers(run):
                return (
                    "Refusing to edit: the target text shares a run with a Word "
                    "field marker (fldChar/fldData/instrText); deleting it would "
                    "damage the field. Edit around the field, or pass "
                    "--allow-field-edit."
                )

    # Case 2 applies to every mode: a field-marker run between matched runs.
    all_runs = _collect_runs(para)
    matched_ids = {id(r) for r, _, _ in matched_runs}
    indices = [i for i, r in enumerate(all_runs) if id(r) in matched_ids]
    if len(indices) >= 2:
        for i in range(indices[0], indices[-1] + 1):
            r = all_runs[i]
            if id(r) not in matched_ids and _run_has_field_markers(r):
                return (
                    "Refusing to edit: the target text spans a Word field (the "
                    "match crosses a field's begin/end markers). Edit around the "
                    "field, or pass --allow-field-edit."
                )
    return None


def tracked_edit(doc_xml_path, target, replacement, mode="replace",
                 author="Peter", date=None, paragraph_index=None,
                 allow_field_edit=False):
    """Apply a tracked change to document.xml.

    Args:
        doc_xml_path: Path to word/document.xml
        target: Text to find (for delete/replace) or anchor text (for insert)
        replacement: New text (for insert/replace), ignored for delete
        mode: "insert", "delete", or "replace"
        author: Author name for tracked change
        date: ISO timestamp, defaults to current UTC time
        paragraph_index: Optional 0-based paragraph index to restrict search
        allow_field_edit: If False (default), refuse edits that overlap or span
            a Word field (EndNote citations, cross-references, TOC, ...). Such
            edits are inherently ambiguous and can silently corrupt a field.
            Pass True to override the preflight guard; a post-edit field
            fingerprint check still runs as a safety net.

    Returns:
        (ids_used, message) tuple. ids_used is list of w:id values assigned.
        An empty ids_used list with an "Error:"/"Refusing" message means the
        edit was not applied.
    """
    path = Path(doc_xml_path)
    path = Path(doc_xml_path)
    dom = defusedxml.minidom.parseString(path.read_text(encoding="utf-8"))

    # Snapshot field structure before editing. Compared again after the edit as
    # a fail-safe: a text change must never silently drop/damage a Word field.
    before_fp = _field_fingerprint(dom)

    if date is None:
        date = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    paragraphs = _find_paragraphs(dom)

    # Find all matching paragraphs
    matches = []
    for i, para in enumerate(paragraphs):
        result = find_text_in_paragraph(para, target)
        if result:
            matches.append((i, para, result))

    if not matches:
        return [], f"Error: '{target}' not found in document"

    if paragraph_index is not None:
        matches = [(i, p, r) for i, p, r in matches if i == paragraph_index]
        if not matches:
            return [], f"Error: '{target}' not found in paragraph {paragraph_index}"

    if len(matches) > 1:
        indices = [str(i) for i, _, _ in matches]
        return [], (
            f"Error: '{target}' found in {len(matches)} paragraphs: {', '.join(indices)}. "
            f"Use --paragraph N to specify which one."
        )

    para_idx, para, matched_runs = matches[0]
    next_id = get_max_id(dom) + 1
    ids_used = []

    # Determine the ns prefix from the first matched run
    ns_prefix = matched_runs[0][0].prefix or "w"

    # Step 1: Split boundary runs to isolate the matched text
    # We need to handle the first and last runs which may be partially matched
    isolated_runs = []

    for i, (run, start, end) in enumerate(matched_runs):
        text = _get_text(run)
        parent = run.parentNode

        if start == 0 and end == len(text):
            # Entire run is matched — use as-is
            isolated_runs.append(run)
        elif start > 0 and end < len(text):
            # Match is in the middle — split twice
            left, remainder = split_run_at(run, start, dom)
            middle, right = split_run_at(remainder, end - start, dom)
            parent.insertBefore(left, run)
            parent.insertBefore(middle, run)
            parent.insertBefore(right, run)
            parent.removeChild(run)
            isolated_runs.append(middle)
        elif start > 0:
            # Match starts partway through
            left, right = split_run_at(run, start, dom)
            parent.insertBefore(left, run)
            parent.insertBefore(right, run)
            parent.removeChild(run)
            isolated_runs.append(right)
        else:
            # Match ends partway through (start == 0, end < len)
            left, right = split_run_at(run, end, dom)
            parent.insertBefore(left, run)
            parent.insertBefore(right, run)
            parent.removeChild(run)
            isolated_runs.append(left)

    # Preflight guard (P3): refuse to edit across or over a Word field unless
    # the caller explicitly opted in with allow_field_edit. This catches the
    # ambiguous case where the match includes a field's display text while its
    # begin/end markers fall outside (or share a run with) the matched region.
    conflict = _detect_field_conflict(para, matched_runs, isolated_runs, mode)
    field_conflict = conflict is not None
    if field_conflict and not allow_field_edit:
        return [], conflict
    # With allow_field_edit we proceed; the post-edit fingerprint check below
    # is the ultimate safety net, and a warning is appended to the result.

    # Step 2: Build tracked change elements
    ref_rpr = _clone_rpr(isolated_runs[0], dom)
    first_run = isolated_runs[0]
    insert_parent = first_run.parentNode

    # Capture a stable reference: the node after the last isolated run
    after_ref = isolated_runs[-1].nextSibling

    if mode in ("delete", "replace"):
        del_elem = dom.createElement(f"{ns_prefix}:del")
        del_elem.setAttribute(f"{ns_prefix}:id", str(next_id))
        del_elem.setAttribute(f"{ns_prefix}:author", author)
        del_elem.setAttribute(f"{ns_prefix}:date", date)
        ids_used.append(next_id)
        next_id += 1

        for run in isolated_runs:
            # Convert ALL <w:t> children to <w:delText>. After field-aware
            # splitting a run may contain more than one text node; converting
            # only the first would leave visible text behind inside <w:del>.
            t_children = [
                c for c in run.childNodes
                if c.nodeType == c.ELEMENT_NODE and _local(c) == "t"
            ]
            for t in t_children:
                run.replaceChild(
                    _make_t_element(dom, "delText", _node_text(t), ns_prefix), t
                )
            del_elem.appendChild(run)  # moves run out of parent

        # Insert del_elem where the runs used to be (before after_ref)
        insert_parent.insertBefore(del_elem, after_ref)

    if mode in ("insert", "replace"):
        ins_elem = dom.createElement(f"{ns_prefix}:ins")
        ins_elem.setAttribute(f"{ns_prefix}:id", str(next_id))
        ins_elem.setAttribute(f"{ns_prefix}:author", author)
        ins_elem.setAttribute(f"{ns_prefix}:date", date)
        ids_used.append(next_id)
        next_id += 1

        rpr_for_ins = _clone_rpr(isolated_runs[0], dom) if ref_rpr else None
        ins_run = _make_run(dom, rpr_for_ins,
                            _make_t_element(dom, "t", replacement, ns_prefix),
                            ns_prefix)
        ins_elem.appendChild(ins_run)

        if mode == "replace":
            # Insert after the <w:del>
            insert_parent.insertBefore(ins_elem, del_elem.nextSibling)
        elif mode == "insert":
            # Insert after the anchor runs
            insert_parent.insertBefore(ins_elem, after_ref)

    # Post-edit fail-safe: compare the field fingerprint before/after. A text
    # edit must not alter field structure (begin/separate/end balance, field
    # data, or instruction codes). If it did, the edit damaged a field.
    after_fp = _field_fingerprint(dom)
    field_note = ""
    if after_fp != before_fp:
        if not allow_field_edit:
            # Do not write; leave document.xml untouched.
            return [], (
                "Refusing to edit: this change would alter a Word field "
                "(fldChar/fldData/instrText). Field fingerprint changed: "
                f"begin/sep/end {before_fp[:3]} -> {after_fp[:3]}, "
                f"instruction count {len(before_fp[3])} -> {len(after_fp[3])}. "
                "Edit around the field, or re-run with --allow-field-edit if you "
                "intend to modify a field."
            )
        field_note = (
            " (WARNING: field fingerprint changed "
            f"{before_fp[:3]} -> {after_fp[:3]}; --allow-field-edit was set)"
        )
    elif field_conflict:
        # The edit proceeded over a field region even though field markers were
        # preserved (e.g. they were relocated into <w:del>). Flag it so the
        # caller knows a field was touched.
        field_note = (
            " (WARNING: edited over a Word field; --allow-field-edit was set. "
            "Verify the citation/field survived.)"
        )

    path.write_bytes(dom.toxml(encoding="UTF-8"))

    mode_verb = {"insert": "Inserted", "delete": "Deleted", "replace": "Replaced"}
    return ids_used, f"{mode_verb[mode]} in paragraph {para_idx}, w:id(s): {ids_used}{field_note}"
