"""Pack a directory into a .docx file.

Validates with auto-repair, condenses XML formatting, and creates the .docx file.

Usage:
    docx pack <input_directory> <output_file> [--original <file>] [--validate true|false]
    docx pack <input_directory> --check --original <file>
    docx pack <input_directory> --diff --original <file>
"""

import argparse
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import defusedxml.minidom

from agent_docx.core.simplify_redlines import infer_author
from agent_docx.validators import DOCXSchemaValidator, RedliningValidator


def pack(
    input_directory: str,
    output_file: str,
    original_file: str | None = None,
    validate: bool = True,
) -> tuple[None, str]:
    input_dir = Path(input_directory)
    output_path = Path(output_file)

    if not input_dir.is_dir():
        return None, f"Error: {input_dir} is not a directory"

    if output_path.suffix.lower() != ".docx":
        return None, f"Error: {output_file} must be a .docx file"

    if validate and original_file:
        original_path = Path(original_file)
        if original_path.exists():
            success, output = _run_validation(input_dir, original_path)
            if output:
                print(output)
            if not success:
                return None, f"Error: Validation failed for {input_dir}"

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_content_dir = Path(temp_dir) / "content"
        shutil.copytree(input_dir, temp_content_dir)

        for pattern in ["*.xml", "*.rels"]:
            for xml_file in temp_content_dir.rglob(pattern):
                _condense_xml(xml_file)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in temp_content_dir.rglob("*"):
                if f.is_file():
                    zf.write(f, f.relative_to(temp_content_dir))

    return None, f"Successfully packed {input_dir} to {output_file}"


def _run_validation(
    unpacked_dir: Path,
    original_file: Path,
) -> tuple[bool, str | None]:
    output_lines = []

    author = "Peter"
    try:
        author = infer_author(unpacked_dir, original_file)
    except ValueError as e:
        print(f"Warning: {e} Using default author 'Peter'.", file=sys.stderr)

    validators = [
        DOCXSchemaValidator(unpacked_dir, original_file),
        RedliningValidator(unpacked_dir, original_file, author=author),
    ]

    total_repairs = sum(v.repair() for v in validators)
    if total_repairs:
        output_lines.append(f"Auto-repaired {total_repairs} issue(s)")

    success = all(v.validate() for v in validators)

    if success:
        output_lines.append("All validations PASSED!")

    return success, "\n".join(output_lines) if output_lines else None


def _run_diff(unpacked_dir: Path, original_file: Path, author: str = "Peter") -> bool:
    """Lightweight diff between unpacked dir and original.

    Returns True if both text AND field structure are identical. The field
    fingerprint check is essential: dropping a citation's field wrapper while
    keeping its display glyph (e.g. "[69]") yields identical text but a changed
    field structure, which the text-only diff cannot see.
    """
    validator = RedliningValidator(unpacked_dir, original_file, author=author)

    modified_xml = unpacked_dir / "word" / "document.xml"
    if not modified_xml.exists():
        print(f"Error: {modified_xml} not found", file=sys.stderr)
        return False

    mod_tree = ET.parse(modified_xml)
    mod_root = mod_tree.getroot()
    validator._remove_author_tracked_changes(mod_root)
    modified_text = validator._extract_text_content(mod_root)
    modified_fields = validator._extract_field_fingerprint(mod_root)

    with tempfile.TemporaryDirectory() as tmp:
        with zipfile.ZipFile(original_file, "r") as z:
            z.extractall(tmp)
        orig_xml = Path(tmp) / "word" / "document.xml"
        if not orig_xml.exists():
            print(f"Error: document.xml not found in {original_file}", file=sys.stderr)
            return False
        orig_tree = ET.parse(orig_xml)
        orig_root = orig_tree.getroot()
        original_text = validator._extract_text_content(orig_root)
        original_fields = validator._extract_field_fingerprint(orig_root)

    ok = True

    if modified_text == original_text:
        print("No untracked text differences")
    else:
        diff = validator._get_git_word_diff(original_text, modified_text)
        if diff:
            print("Untracked text differences found:\n")
            print(diff)
        else:
            print("Texts differ but unable to generate diff (git not available)")
        ok = False

    if modified_fields != original_fields:
        print("\nField/citation fingerprint changed:")
        print(
            f"  fldChar begin/separate/end: "
            f"{original_fields[:3]} -> {modified_fields[:3]}"
        )
        print(
            f"  field-instruction count: "
            f"{len(original_fields[3])} -> {len(modified_fields[3])}"
        )
        print(
            "  A field (e.g. an EndNote citation) may have been dropped or "
            "damaged. --diff text comparison alone cannot detect this."
        )
        ok = False

    return ok


def _condense_xml(xml_file: Path) -> None:
    try:
        with open(xml_file, encoding="utf-8") as f:
            dom = defusedxml.minidom.parse(f)

        for element in dom.getElementsByTagName("*"):
            if element.tagName.endswith(":t"):
                continue

            for child in list(element.childNodes):
                if (
                    child.nodeType == child.TEXT_NODE
                    and child.nodeValue
                    and child.nodeValue.strip() == ""
                ) or child.nodeType == child.COMMENT_NODE:
                    element.removeChild(child)

        xml_file.write_bytes(dom.toxml(encoding="UTF-8"))
    except Exception as e:
        print(f"ERROR: Failed to parse {xml_file.name}: {e}", file=sys.stderr)
        raise


def register(subparsers) -> None:
    parser = subparsers.add_parser("pack", help="Pack a directory into a .docx file")
    parser.add_argument("input_directory", help="Unpacked .docx directory")
    parser.add_argument("output_file", nargs="?", help="Output .docx file")
    parser.add_argument("--original", help="Original file for validation comparison")
    parser.add_argument(
        "--validate",
        type=lambda x: x.lower() == "true",
        default=True,
        metavar="true|false",
        help="Run validation with auto-repair (default: true)",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="Dry-run: validate without packing (requires --original)",
    )
    parser.add_argument(
        "--diff",
        action="store_true",
        help="Show text diff against original (requires --original)",
    )
    parser.set_defaults(func=_run)


def _run(args: argparse.Namespace) -> int:
    if (args.check or args.diff) and not args.original:
        print("Error: --check and --diff require --original")
        return 1

    if not args.check and not args.diff and not args.output_file:
        print("Error: output_file is required unless using --check or --diff")
        return 1

    ok = True

    if args.diff:
        input_dir = Path(args.input_directory)
        original_path = Path(args.original)
        if not _run_diff(input_dir, original_path):
            ok = False

    if args.check:
        input_dir = Path(args.input_directory)
        original_path = Path(args.original)
        success, output = _run_validation(input_dir, original_path)
        if output:
            print(output)
        if not success:
            ok = False

    if args.check or args.diff:
        return 0 if ok else 1

    _, message = pack(
        args.input_directory,
        args.output_file,
        original_file=args.original,
        validate=args.validate,
    )
    print(message)

    return 1 if "Error" in message else 0
