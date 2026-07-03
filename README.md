# agent-docx

CLI tools for editing `.docx` files with tracked changes and comments, built for use by AI agents (see `skills/docx/SKILL.md` for the agent-facing workflow).

A `.docx` file is a ZIP archive of XML files. This package provides a single `docx` command with subcommands to unpack, edit, comment on, validate, and repack that XML safely.

## Install

```bash
uv pip install -e .
```

This installs the `docx` console script on `PATH` inside the active environment.

### Third-Party Python Dependencies

| Package | Version | Used by |
|---------|---------|---------|
| defusedxml | 0.7.1 | unpack, pack, comment, edit, find-artifacts, core/*, validators/* |
| lxml | 6.1.1 | validators/* |

All other imports are Python stdlib. Dev-only: `pytest` (installed via the `dev` extra: `uv pip install -e ".[dev]"`).

### Non-Python Dependencies

| Tool | Required for | Notes |
|------|---------------|-------|
| git | `docx pack --diff`, redlining validation word-diff output | Falls back to a plain "texts differ" message if unavailable |
| LibreOffice (`soffice`) | `docx accept-changes` only | Optional — this command fails with a clear error if `soffice` is not on `PATH`. Install: `apt install libreoffice` |

## CLI Usage

```bash
docx unpack document.docx unpacked/
docx edit unpacked/ --find "old text" --replace "new text"
docx comment unpacked/ 0 "Comment text"
docx find-artifacts unpacked/
docx validate unpacked/ --original document.docx
docx pack unpacked/ output.docx --original document.docx
docx accept-changes tracked.docx clean.docx   # requires LibreOffice
```

Run `docx <subcommand> --help` for full options on any subcommand.

## Repository Structure

```
pyproject.toml
src/agent_docx/
├── cli.py                     # `docx` dispatcher (argparse subparsers)
├── commands/                  # thin CLI layer: arg parsing + wiring to core logic
│   ├── unpack.py
│   ├── validate.py
│   ├── pack.py
│   ├── tracked_edit.py        # subcommand: edit
│   ├── find_artifacts.py      # subcommand: find-artifacts
│   ├── comment.py
│   └── accept_changes.py
├── core/                       # pure logic, no CLI/argparse concerns
│   ├── merge_runs.py
│   ├── simplify_redlines.py
│   ├── tracked_changes.py
│   ├── artifacts.py
│   └── soffice.py              # LibreOffice wrapper (headless, sandboxed)
├── validators/
│   ├── base.py                 # schema validation (lxml + defusedxml)
│   ├── docx.py                 # DOCX-specific validation
│   └── redlining.py            # tracked-changes validation
├── schemas/                    # XML schema files (ECMA, ISO, Microsoft) used by validators
└── templates/                  # XML templates for comments.xml and sidecar parts
skills/docx/
├── SKILL.md                    # agent-facing workflow instructions
└── references/                 # supplementary reference docs (if any)
tests/
├── fixtures/                   # sample.docx test fixture + its generator script
└── test_*.py
```

## Testing

```bash
uv pip install -e ".[dev]"
pytest
```

Tests invoke the installed `docx` CLI against `tests/fixtures/sample.docx` and assert on output/exit codes — they are integration-style rather than deep unit tests of internals. The `accept-changes` LibreOffice-dependent test skips automatically if `soffice` is not on `PATH`.
