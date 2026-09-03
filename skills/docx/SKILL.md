---
name: docx
description: Edit, revise, and redline Word documents with tracked changes, comments, and validation.
---

# Word documents

Use the `docx` CLI. Verify the command before editing:

```bash
command -v docx
docx --help
```

## Installation

If `docx` is unavailable, install the immutable `v0.1.1` release. This requires `uv`, network access, and Python 3.12 or later:

```bash
uv tool install 'git+https://github.com/peterchang77/agent-docx@v0.1.1'
command -v docx
docx --help
```

To replace an existing installation with that release, add `--force`. Do not rely on a checkout at a machine-specific path. `docx accept-changes` additionally requires LibreOffice (`soffice`) on `PATH`.

Use `Peter` as the tracked-change and comment author unless the user requests another name. Do not use em dashes in inserted or edited text. Write comments as direct statements or questions without labels such as `TODO:` or `Note:`.

## Workflow

1. Preserve the original and unpack into a new, empty working directory:

   ```bash
   docx unpack document.docx unpacked/
   ```

   Read `unpacked/word/document.xml` directly. It is the ground truth for content, formatting, fields, and tracked changes. Unpacking defaults to merging runs and simplifying adjacent redlines. Use `--merge-runs false --simplify-redlines false` when original run boundaries or redline boundaries matter.

2. Edit with `docx edit` whenever possible:

   ```bash
   docx edit unpacked/ --find "old text" --replace "new text"
   docx edit unpacked/ --find "text" --mode delete
   docx edit unpacked/ --find "anchor" --replace " added" --mode insert
   ```

   Use `--paragraph N` for ambiguous matches, and `--author NAME` or `--date ISO_TIMESTAMP` when needed. Use `--allow-field-edit` only when deliberately editing a Word field. For comments, follow [the comment workflow](references/comments.md).

3. For structural changes, edit `unpacked/word/document.xml` directly. Replace complete `<w:r>...</w:r>` blocks, preserve each run's `<w:rPr>`, make matches unique, and recheck the surrounding XML after every edit. Preserve field markers (`w:fldChar`, `w:fldData`, and `w:instrText`) exactly. Use XML entities for quotes and apostrophes. See [XML and validation details](references/xml-and-validation.md).

4. Validate before packing:

   ```bash
   docx pack unpacked/ --check --diff --original document.docx
   docx validate unpacked/ --original document.docx
   docx find-artifacts unpacked/ --verbose
   ```

   `validate` without `--original` performs schema validation but skips redlining comparison. Review artifacts before using `find-artifacts --fix`, then validate again. `pack --check` is the dry-run validation path; `pack --diff` checks untracked text and field structure.

5. Pack the result:

   ```bash
   docx pack unpacked/ output.docx --original document.docx
   ```

   Keep `--validate` enabled unless explicitly appropriate. Accept all tracked changes only when requested and after review:

   ```bash
   docx accept-changes tracked.docx clean.docx
   ```

   Validate the final output when practical. Run `docx <subcommand> --help` for options not shown here.
