"""Scan unpacked DOCX XML for formatting artifacts.

Detects stray <w:br/>-only runs, empty styled runs, and whitespace-only text runs.

Usage:
    docx find-artifacts unpacked/
    docx find-artifacts unpacked/ --fix
    docx find-artifacts unpacked/ --verbose
"""

import argparse
from pathlib import Path

from agent_docx.core.artifacts import find_artifacts


def register(subparsers) -> None:
    parser = subparsers.add_parser(
        "find-artifacts", help="Scan unpacked DOCX XML for formatting artifacts"
    )
    parser.add_argument("unpacked_dir", help="Unpacked DOCX directory")
    parser.add_argument("--fix", action="store_true", help="Auto-remove detected artifacts")
    parser.add_argument(
        "--verbose", action="store_true", help="Show XML context for each artifact"
    )
    parser.set_defaults(func=_run)


def _run(args: argparse.Namespace) -> int:
    doc_xml = Path(args.unpacked_dir) / "word" / "document.xml"
    if not doc_xml.exists():
        print(f"Error: {doc_xml} not found")
        return 1

    artifacts, fix_count, message = find_artifacts(doc_xml, fix=args.fix)
    print(message)

    if args.verbose and artifacts:
        for i, a in enumerate(artifacts, 1):
            print(f"\n  [{i}] {a.type}: {a.context}")

    if artifacts and not args.fix:
        return 1
    return 0
