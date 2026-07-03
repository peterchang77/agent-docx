"""Validate .docx document XML against XSD schemas and tracked-changes rules.

Usage:
    docx validate <path> [--original <original_file>] [--auto-repair] [--author NAME]

The first argument can be either:
- An unpacked directory containing the DOCX XML files
- A packed .docx file, which will be unpacked to a temp directory

Auto-repair fixes:
- paraId/durableId values that exceed OOXML limits
- Missing xml:space="preserve" on w:t elements with whitespace
"""

import argparse
import tempfile
import zipfile
from pathlib import Path

from agent_docx.validators import DOCXSchemaValidator, RedliningValidator


def register(subparsers) -> None:
    parser = subparsers.add_parser(
        "validate", help="Validate .docx XML against schemas and tracked-changes rules"
    )
    parser.add_argument(
        "path", help="Path to unpacked directory or packed .docx file"
    )
    parser.add_argument(
        "--original",
        default=None,
        help="Path to original .docx file. If omitted, all XSD errors are reported "
        "and redlining validation is skipped.",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose output")
    parser.add_argument(
        "--auto-repair",
        action="store_true",
        help="Automatically repair common issues (hex IDs, whitespace preservation)",
    )
    parser.add_argument(
        "--author", default="Peter", help="Author name for redlining validation"
    )
    parser.set_defaults(func=_run)


def _run(args: argparse.Namespace) -> int:
    path = Path(args.path)
    if not path.exists():
        print(f"Error: {path} does not exist")
        return 1

    original_file = None
    if args.original:
        original_file = Path(args.original)
        if not original_file.is_file():
            print(f"Error: {original_file} is not a file")
            return 1
        if original_file.suffix.lower() != ".docx":
            print(f"Error: {original_file} must be a .docx file")
            return 1

    if path.is_file() and path.suffix.lower() == ".docx":
        temp_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(path, "r") as zf:
            zf.extractall(temp_dir)
        unpacked_dir = Path(temp_dir)
    elif path.is_dir():
        unpacked_dir = path
    else:
        print(f"Error: {path} is not a directory or .docx file")
        return 1

    validators = [DOCXSchemaValidator(unpacked_dir, original_file, verbose=args.verbose)]
    if original_file:
        validators.append(
            RedliningValidator(
                unpacked_dir, original_file, verbose=args.verbose, author=args.author
            )
        )

    if args.auto_repair:
        total_repairs = sum(v.repair() for v in validators)
        if total_repairs:
            print(f"Auto-repaired {total_repairs} issue(s)")

    success = all(v.validate() for v in validators)

    if success:
        print("All validations PASSED!")

    return 0 if success else 1
