"""Apply tracked changes to unpacked DOCX XML.

Supports insert, delete, and replace operations with proper <w:ins>/<w:del> markup.
Handles text spanning multiple <w:r> elements.

Usage:
    docx edit unpacked/ --find "old text" --replace "new text"
    docx edit unpacked/ --find "delete me" --mode delete
    docx edit unpacked/ --find "anchor" --replace "inserted" --mode insert
    docx edit unpacked/ --find "ambiguous" --replace "new" --paragraph 3
"""

import argparse
from pathlib import Path

from agent_docx.core.tracked_changes import tracked_edit


def register(subparsers) -> None:
    parser = subparsers.add_parser(
        "edit", help="Apply a tracked change (insert/delete/replace) to unpacked DOCX XML"
    )
    parser.add_argument("unpacked_dir", help="Unpacked DOCX directory")
    parser.add_argument("--find", required=True, help="Text to find")
    parser.add_argument("--replace", help="Replacement text (required for insert/replace)")
    parser.add_argument(
        "--mode",
        choices=["insert", "delete", "replace"],
        default="replace",
        help="Edit mode (default: replace)",
    )
    parser.add_argument("--author", default="Peter", help="Author name (default: Peter)")
    parser.add_argument("--date", help="ISO timestamp (default: current UTC time)")
    parser.add_argument(
        "--paragraph", type=int, default=None,
        help="Restrict to Nth paragraph (0-indexed)",
    )
    parser.set_defaults(func=_run)


def _run(args: argparse.Namespace) -> int:
    if args.mode in ("insert", "replace") and not args.replace:
        print(f"Error: --replace is required for {args.mode} mode")
        return 1

    doc_xml = Path(args.unpacked_dir) / "word" / "document.xml"
    if not doc_xml.exists():
        print(f"Error: {doc_xml} not found")
        return 1

    ids_used, message = tracked_edit(
        doc_xml,
        args.find,
        args.replace or "",
        mode=args.mode,
        author=args.author,
        date=args.date,
        paragraph_index=args.paragraph,
    )

    print(message)
    return 0 if ids_used else 1
