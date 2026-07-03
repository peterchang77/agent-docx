"""`docx` CLI dispatcher.

Subcommands are registered incrementally as they are implemented:
    unpack, pack, validate, edit, find-artifacts, comment, accept-changes
"""

import argparse
import sys

from agent_docx.commands import (
    accept_changes,
    comment,
    find_artifacts,
    pack,
    tracked_edit,
    unpack,
    validate,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx",
        description="Edit .docx files with tracked changes and comments.",
    )
    subparsers = parser.add_subparsers(dest="command")

    unpack.register(subparsers)
    validate.register(subparsers)
    pack.register(subparsers)
    tracked_edit.register(subparsers)
    find_artifacts.register(subparsers)
    comment.register(subparsers)
    accept_changes.register(subparsers)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command is None:
        parser.print_help()
        return 1

    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
