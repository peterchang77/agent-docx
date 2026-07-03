"""Detect and optionally fix formatting artifacts in DOCX XML.

Scans for:
- br_only_run: <w:r> containing only <w:br/> (with optional <w:rPr>)
- empty_styled_run: <w:r> with <w:rPr> but no text content at all
- whitespace_only_text: <w:r> where <w:t> contains only whitespace
"""

from collections import namedtuple
from pathlib import Path

import defusedxml.minidom

Artifact = namedtuple("Artifact", ["type", "line_hint", "context"])


def _get_child(parent, local_name):
    for child in parent.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            name = child.localName or child.tagName
            if name == local_name or name.endswith(f":{local_name}"):
                return child
    return None


def _get_children(parent, local_name):
    results = []
    for child in parent.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            name = child.localName or child.tagName
            if name == local_name or name.endswith(f":{local_name}"):
                results.append(child)
    return results


def _element_children(parent):
    return [c for c in parent.childNodes if c.nodeType == c.ELEMENT_NODE]


def _local_name(node):
    return node.localName or node.tagName


def _is_run(node):
    name = _local_name(node)
    return name == "r" or name.endswith(":r")


def _preview(node, max_len=80):
    xml = node.toxml()
    return xml[:max_len] + "..." if len(xml) > max_len else xml


def _classify_run(run):
    """Classify a run as an artifact type, or None if it's normal."""
    children = _element_children(run)
    names = [_local_name(c) for c in children]

    has_rpr = any(n in ("rPr", "w:rPr") or n.endswith(":rPr") for n in names)
    content_children = [c for c, n in zip(children, names)
                        if not (n in ("rPr", "w:rPr") or n.endswith(":rPr"))]

    # br_only_run: only content is <w:br/>
    if content_children:
        content_names = [_local_name(c) for c in content_children]
        if all(n in ("br",) or n.endswith(":br") for n in content_names):
            return "br_only_run"

    # empty_styled_run: has rPr but no content children at all
    if has_rpr and not content_children:
        return "empty_styled_run"

    # whitespace_only_text: has <w:t> with only whitespace
    t_elem = _get_child(run, "t")
    if t_elem:
        text = t_elem.firstChild.nodeValue if t_elem.firstChild else ""
        if not text.strip():
            # Only flag if there's no other meaningful content besides rPr and this t
            non_rpr_non_t = [c for c, n in zip(children, names)
                             if not (n.endswith(":rPr") or n == "rPr"
                                     or n.endswith(":t") or n == "t")]
            if not non_rpr_non_t:
                return "whitespace_only_text"

    return None


def find_artifacts(doc_xml_path, fix=False):
    """Scan document.xml for formatting artifacts.

    Returns (artifacts_list, fix_count, message).
    """
    path = Path(doc_xml_path)
    dom = defusedxml.minidom.parseString(path.read_text(encoding="utf-8"))

    artifacts = []
    to_remove = []

    for elem in dom.getElementsByTagName("*"):
        if not _is_run(elem):
            continue
        # Skip runs inside tracked changes
        parent = elem.parentNode
        if parent:
            pname = _local_name(parent)
            if pname in ("ins", "del") or pname.endswith(":ins") or pname.endswith(":del"):
                continue

        artifact_type = _classify_run(elem)
        if artifact_type:
            artifacts.append(Artifact(artifact_type, None, _preview(elem)))
            if fix:
                to_remove.append(elem)

    fix_count = 0
    if fix and to_remove:
        for elem in to_remove:
            if elem.parentNode:
                elem.parentNode.removeChild(elem)
                fix_count += 1
        path.write_bytes(dom.toxml(encoding="UTF-8"))

    counts = {}
    for a in artifacts:
        counts[a.type] = counts.get(a.type, 0) + 1

    parts = [f"{v} {k}" for k, v in sorted(counts.items())]
    summary = ", ".join(parts) if parts else "none"
    msg = f"Found {len(artifacts)} artifact(s): {summary}"
    if fix_count:
        msg += f". Removed {fix_count}."

    return artifacts, fix_count, msg
