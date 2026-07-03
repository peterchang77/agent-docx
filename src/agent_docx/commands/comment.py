"""Add comments to DOCX documents.

Usage:
    docx comment unpacked/ 0 "Comment text"
    docx comment unpacked/ 1 "Reply text" --parent 0

Text should be pre-escaped XML (e.g., &amp; for &, &#x2019; for smart quotes).

After running, add markers to document.xml:
  <w:commentRangeStart w:id="0"/>
  ... commented content ...
  <w:commentRangeEnd w:id="0"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
"""

import argparse
import importlib.resources
import random
import shutil
from datetime import datetime, timezone
from pathlib import Path

import defusedxml.minidom

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
}

COMMENT_XML = """\
<w:comment w:id="{id}" w:author="{author}" w:date="{date}" w:initials="{initials}">
  <w:p w14:paraId="{para_id}" w14:textId="77777777">
    <w:r>
      <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
      <w:annotationRef/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:color w:val="000000"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
      <w:t>{text}</w:t>
    </w:r>
  </w:p>
</w:comment>"""

COMMENT_MARKER_TEMPLATE = """
Add to document.xml (markers must be direct children of w:p, never inside w:r):
  <w:commentRangeStart w:id="{cid}"/>
  <w:r>...</w:r>
  <w:commentRangeEnd w:id="{cid}"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>"""

REPLY_MARKER_TEMPLATE = """
Nest markers inside parent {pid}'s markers (markers must be direct children of w:p, never inside w:r):
  <w:commentRangeStart w:id="{pid}"/><w:commentRangeStart w:id="{cid}"/>
  <w:r>...</w:r>
  <w:commentRangeEnd w:id="{cid}"/><w:commentRangeEnd w:id="{pid}"/>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{pid}"/></w:r>
  <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="{cid}"/></w:r>"""


def _template_path(name: str) -> Path:
    return importlib.resources.files("agent_docx") / "templates" / name


def _generate_hex_id() -> str:
    return f"{random.randint(0, 0x7FFFFFFE):08X}"


SMART_QUOTE_ENTITIES = {
    "\u201c": "&#x201C;",
    "\u201d": "&#x201D;",
    "\u2018": "&#x2018;",
    "\u2019": "&#x2019;",
}


def _encode_smart_quotes(text: str) -> str:
    for char, entity in SMART_QUOTE_ENTITIES.items():
        text = text.replace(char, entity)
    return text


def _append_xml(xml_path: Path, root_tag: str, content: str) -> None:
    dom = defusedxml.minidom.parseString(xml_path.read_text(encoding="utf-8"))
    root = dom.getElementsByTagName(root_tag)[0]
    ns_attrs = " ".join(f'xmlns:{k}="{v}"' for k, v in NS.items())
    wrapper_dom = defusedxml.minidom.parseString(f"<root {ns_attrs}>{content}</root>")
    for child in wrapper_dom.documentElement.childNodes:
        if child.nodeType == child.ELEMENT_NODE:
            root.appendChild(dom.importNode(child, True))
    output = _encode_smart_quotes(dom.toxml(encoding="UTF-8").decode("utf-8"))
    xml_path.write_text(output, encoding="utf-8")


def _find_para_id(comments_path: Path, comment_id: int) -> str | None:
    dom = defusedxml.minidom.parseString(comments_path.read_text(encoding="utf-8"))
    for c in dom.getElementsByTagName("w:comment"):
        if c.getAttribute("w:id") == str(comment_id):
            for p in c.getElementsByTagName("w:p"):
                if pid := p.getAttribute("w14:paraId"):
                    return pid
    return None


def _get_next_rid(rels_path: Path) -> int:
    dom = defusedxml.minidom.parseString(rels_path.read_text(encoding="utf-8"))
    max_rid = 0
    for rel in dom.getElementsByTagName("Relationship"):
        rid = rel.getAttribute("Id")
        if rid and rid.startswith("rId"):
            try:
                max_rid = max(max_rid, int(rid[3:]))
            except ValueError:
                pass
    return max_rid + 1


def _has_relationship(rels_path: Path, target: str) -> bool:
    dom = defusedxml.minidom.parseString(rels_path.read_text(encoding="utf-8"))
    for rel in dom.getElementsByTagName("Relationship"):
        if rel.getAttribute("Target") == target:
            return True
    return False


def _has_content_type(ct_path: Path, part_name: str) -> bool:
    dom = defusedxml.minidom.parseString(ct_path.read_text(encoding="utf-8"))
    for override in dom.getElementsByTagName("Override"):
        if override.getAttribute("PartName") == part_name:
            return True
    return False


def _ensure_comment_relationships(unpacked_dir: Path) -> None:
    rels_path = unpacked_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return

    rels = [
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            "comments.xml",
        ),
        (
            "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
            "commentsExtended.xml",
        ),
        (
            "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
            "commentsIds.xml",
        ),
        (
            "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
            "commentsExtensible.xml",
        ),
    ]

    missing = [
        (rel_type, target)
        for rel_type, target in rels
        if not _has_relationship(rels_path, target)
    ]
    if not missing:
        return

    dom = defusedxml.minidom.parseString(rels_path.read_text(encoding="utf-8"))
    root = dom.documentElement
    next_rid = _get_next_rid(rels_path)

    for rel_type, target in missing:
        rel = dom.createElement("Relationship")
        rel.setAttribute("Id", f"rId{next_rid}")
        rel.setAttribute("Type", rel_type)
        rel.setAttribute("Target", target)
        root.appendChild(rel)
        next_rid += 1

    rels_path.write_bytes(dom.toxml(encoding="UTF-8"))


def _ensure_comment_content_types(unpacked_dir: Path) -> None:
    ct_path = unpacked_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return

    overrides = [
        (
            "/word/comments.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        ),
        (
            "/word/commentsExtended.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
        ),
        (
            "/word/commentsIds.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
        ),
        (
            "/word/commentsExtensible.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
        ),
    ]

    missing = [
        (part_name, content_type)
        for part_name, content_type in overrides
        if not _has_content_type(ct_path, part_name)
    ]
    if not missing:
        return

    dom = defusedxml.minidom.parseString(ct_path.read_text(encoding="utf-8"))
    root = dom.documentElement

    for part_name, content_type in missing:
        override = dom.createElement("Override")
        override.setAttribute("PartName", part_name)
        override.setAttribute("ContentType", content_type)
        root.appendChild(override)

    ct_path.write_bytes(dom.toxml(encoding="UTF-8"))


def add_comment(
    unpacked_dir: str,
    comment_id: int,
    text: str,
    author: str = "Peter",
    initials: str = "P",
    parent_id: int | None = None,
) -> tuple[str, str]:
    word = Path(unpacked_dir) / "word"
    if not word.exists():
        return "", f"Error: {word} not found"

    para_id, durable_id = _generate_hex_id(), _generate_hex_id()
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    comments = word / "comments.xml"
    first_comment = not comments.exists()
    if first_comment:
        shutil.copy(_template_path("comments.xml"), comments)
    _ensure_comment_relationships(Path(unpacked_dir))
    _ensure_comment_content_types(Path(unpacked_dir))
    _append_xml(
        comments,
        "w:comments",
        COMMENT_XML.format(
            id=comment_id,
            author=author,
            date=ts,
            initials=initials,
            para_id=para_id,
            text=text,
        ),
    )

    ext = word / "commentsExtended.xml"
    if not ext.exists():
        shutil.copy(_template_path("commentsExtended.xml"), ext)
    if parent_id is not None:
        parent_para = _find_para_id(comments, parent_id)
        if not parent_para:
            return "", f"Error: Parent comment {parent_id} not found"
        _append_xml(
            ext,
            "w15:commentsEx",
            f'<w15:commentEx w15:paraId="{para_id}" w15:paraIdParent="{parent_para}" w15:done="0"/>',
        )
    else:
        _append_xml(
            ext,
            "w15:commentsEx",
            f'<w15:commentEx w15:paraId="{para_id}" w15:done="0"/>',
        )

    ids = word / "commentsIds.xml"
    if not ids.exists():
        shutil.copy(_template_path("commentsIds.xml"), ids)
    _append_xml(
        ids,
        "w16cid:commentsIds",
        f'<w16cid:commentId w16cid:paraId="{para_id}" w16cid:durableId="{durable_id}"/>',
    )

    extensible = word / "commentsExtensible.xml"
    if not extensible.exists():
        shutil.copy(_template_path("commentsExtensible.xml"), extensible)
    _append_xml(
        extensible,
        "w16cex:commentsExtensible",
        f'<w16cex:commentExtensible w16cex:durableId="{durable_id}" w16cex:dateUtc="{ts}"/>',
    )

    action = "reply" if parent_id is not None else "comment"
    return para_id, f"Added {action} {comment_id} (para_id={para_id})"


def register(subparsers) -> None:
    parser = subparsers.add_parser("comment", help="Add comments to a DOCX document")
    parser.add_argument("unpacked_dir", help="Unpacked DOCX directory")
    parser.add_argument("comment_id", type=int, help="Comment ID (must be unique)")
    parser.add_argument("text", help="Comment text")
    parser.add_argument("--author", default="Peter", help="Author name")
    parser.add_argument("--initials", default="P", help="Author initials")
    parser.add_argument("--parent", type=int, help="Parent comment ID (for replies)")
    parser.set_defaults(func=_run)


def _run(args: argparse.Namespace) -> int:
    para_id, msg = add_comment(
        args.unpacked_dir,
        args.comment_id,
        args.text,
        args.author,
        args.initials,
        args.parent,
    )
    print(msg)
    if "Error" in msg:
        return 1

    if args.parent is not None:
        print(REPLY_MARKER_TEMPLATE.format(pid=args.parent, cid=args.comment_id))
    else:
        print(COMMENT_MARKER_TEMPLATE.format(cid=args.comment_id))
    return 0
