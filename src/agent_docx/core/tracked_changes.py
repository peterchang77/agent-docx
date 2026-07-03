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


def _get_text(run):
    """Get text content from a <w:r> element (from w:t or w:delText)."""
    for tag in ("t", "delText"):
        t = _get_child(run, tag)
        if t and t.firstChild:
            return t.firstChild.nodeValue or ""
    return ""


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

    Returns (left_run, right_run). Either may be None if offset is at boundary.
    """
    text = _get_text(run)
    if offset <= 0:
        return None, run
    if offset >= len(text):
        return run, None

    left_text = text[:offset]
    right_text = text[offset:]

    # Determine if this is a delText run
    is_del = _get_child(run, "delText") is not None
    tag = "delText" if is_del else "t"
    ns_prefix = run.prefix or "w"

    rpr = _clone_rpr(run, dom)
    left = _make_run(dom, rpr, _make_t_element(dom, tag, left_text, ns_prefix), ns_prefix)

    rpr2 = _clone_rpr(run, dom)
    right = _make_run(dom, rpr2, _make_t_element(dom, tag, right_text, ns_prefix), ns_prefix)

    return left, right


def _find_paragraphs(dom):
    results = []
    for elem in dom.getElementsByTagName("*"):
        name = elem.localName or elem.tagName
        if name == "p" or name.endswith(":p"):
            results.append(elem)
    return results


def tracked_edit(doc_xml_path, target, replacement, mode="replace",
                 author="Peter", date=None, paragraph_index=None):
    """Apply a tracked change to document.xml.

    Args:
        doc_xml_path: Path to word/document.xml
        target: Text to find (for delete/replace) or anchor text (for insert)
        replacement: New text (for insert/replace), ignored for delete
        mode: "insert", "delete", or "replace"
        author: Author name for tracked change
        date: ISO timestamp, defaults to current UTC time
        paragraph_index: Optional 0-based paragraph index to restrict search

    Returns:
        (ids_used, message) tuple. ids_used is list of w:id values assigned.
    """
    path = Path(doc_xml_path)
    dom = defusedxml.minidom.parseString(path.read_text(encoding="utf-8"))

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
            t = _get_child(run, "t")
            if t:
                text_content = t.firstChild.nodeValue if t.firstChild else ""
                new_dt = _make_t_element(dom, "delText", text_content, ns_prefix)
                run.replaceChild(new_dt, t)
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

    path.write_bytes(dom.toxml(encoding="UTF-8"))

    mode_verb = {"insert": "Inserted", "delete": "Deleted", "replace": "Replaced"}
    return ids_used, f"{mode_verb[mode]} in paragraph {para_idx}, w:id(s): {ids_used}"
