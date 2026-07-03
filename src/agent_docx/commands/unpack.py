"""Unpack a .docx file for editing.

Extracts the ZIP archive, pretty-prints XML files, and:
- Merges adjacent runs with identical formatting
- Simplifies adjacent tracked changes from the same author
- Converts smart quotes to XML entities

Usage:
    docx unpack document.docx unpacked/
    docx unpack document.docx unpacked/ --merge-runs false
"""

import argparse
import zipfile
from pathlib import Path

import defusedxml.minidom

from agent_docx.core.merge_runs import merge_runs as do_merge_runs
from agent_docx.core.simplify_redlines import simplify_redlines as do_simplify_redlines

SMART_QUOTE_REPLACEMENTS = {
    "\u201c": "&#x201C;",
    "\u201d": "&#x201D;",
    "\u2018": "&#x2018;",
    "\u2019": "&#x2019;",
}


def unpack(
    input_file: str,
    output_directory: str,
    merge_runs: bool = True,
    simplify_redlines: bool = True,
) -> tuple[None, str]:
    input_path = Path(input_file)
    output_path = Path(output_directory)

    if not input_path.exists():
        return None, f"Error: {input_file} does not exist"

    if input_path.suffix.lower() != ".docx":
        return None, f"Error: {input_file} must be a .docx file"

    try:
        output_path.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(output_path)

        xml_files = list(output_path.rglob("*.xml")) + list(output_path.rglob("*.rels"))
        for xml_file in xml_files:
            _pretty_print_xml(xml_file)

        message = f"Unpacked {input_file} ({len(xml_files)} XML files)"

        if simplify_redlines:
            simplify_count, _ = do_simplify_redlines(str(output_path))
            message += f", simplified {simplify_count} tracked changes"

        if merge_runs:
            merge_count, _ = do_merge_runs(str(output_path))
            message += f", merged {merge_count} runs"

        for xml_file in xml_files:
            _escape_smart_quotes(xml_file)

        return None, message

    except zipfile.BadZipFile:
        return None, f"Error: {input_file} is not a valid Office file"
    except Exception as e:
        return None, f"Error unpacking: {e}"


def _pretty_print_xml(xml_file: Path) -> None:
    try:
        content = xml_file.read_text(encoding="utf-8")
        dom = defusedxml.minidom.parseString(content)
        xml_file.write_bytes(dom.toprettyxml(indent="  ", encoding="utf-8"))
    except Exception:
        pass


def _escape_smart_quotes(xml_file: Path) -> None:
    try:
        content = xml_file.read_text(encoding="utf-8")
        for char, entity in SMART_QUOTE_REPLACEMENTS.items():
            content = content.replace(char, entity)
        xml_file.write_text(content, encoding="utf-8")
    except Exception:
        pass


def register(subparsers) -> None:
    parser = subparsers.add_parser("unpack", help="Unpack a .docx file for editing")
    parser.add_argument("input_file", help=".docx file to unpack")
    parser.add_argument("output_directory", help="Output directory")
    parser.add_argument(
        "--merge-runs",
        type=lambda x: x.lower() == "true",
        default=True,
        metavar="true|false",
        help="Merge adjacent runs with identical formatting (default: true)",
    )
    parser.add_argument(
        "--simplify-redlines",
        type=lambda x: x.lower() == "true",
        default=True,
        metavar="true|false",
        help="Merge adjacent tracked changes from same author (default: true)",
    )
    parser.set_defaults(func=_run)


def _run(args: argparse.Namespace) -> int:
    _, message = unpack(
        args.input_file,
        args.output_directory,
        merge_runs=args.merge_runs,
        simplify_redlines=args.simplify_redlines,
    )
    print(message)
    return 1 if "Error" in message else 0
