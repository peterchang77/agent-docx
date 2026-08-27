---
name: docx
description: Edit, revise, and redline Word documents with tracked changes, comments, and validation.
---

# Word documents

Use the `docx` CLI. It must be on `PATH`.

## Installation

If `docx` is unavailable, install the immutable `v0.1.1` release:

```bash
uv tool install 'git+https://github.com/peterchang77/agent-docx@v0.1.1'
docx --help
```

To replace an existing installation with that release, add `--force` to the `uv tool install` command. Do not rely on a checkout at a machine-specific path.

Use `Peter` as the tracked-change and comment author unless the user requests another name. Do not use em dashes in inserted or edited text. Write comments as direct statements or questions without labels such as `TODO:` or `Note:`.

## Workflow

1. Unpack the document.

   ```bash
   docx unpack document.docx unpacked/
   ```

   Read `unpacked/word/document.xml` directly. It is the ground truth for content, formatting, fields, and tracked changes. Use `--merge-runs false` when original run boundaries matter.

2. Edit with `docx edit` whenever possible.

   ```bash
   docx edit unpacked/ --find "old text" --replace "new text"
   docx edit unpacked/ --find "text" --mode delete
   docx edit unpacked/ --find "anchor" --replace " added" --mode insert
   ```

   Use `--paragraph N` for ambiguous matches, and `--author NAME` or `--date ISO_TIMESTAMP` when needed. Use `--allow-field-edit` only when deliberately editing a Word field.

3. For structural changes, edit `unpacked/word/document.xml` directly. Replace complete `<w:r>...</w:r>` blocks, preserve each run's `<w:rPr>`, make matches unique, and recheck the surrounding XML after every edit. Preserve field markers (`w:fldChar`, `w:fldData`, and `w:instrText`) exactly. Use XML entities for quotes and apostrophes.

   Comment range markers are siblings of `<w:r>`, never children. A full paragraph deletion must also mark the paragraph mark as deleted. Tracked-change and comment IDs must be unique document-wide. Preserve the original author's changes when rejecting or restoring them by nesting or placing the new change correctly.

4. Validate before packing.

   ```bash
   docx pack unpacked/ --check --diff --original document.docx
   docx validate unpacked/ --original document.docx
   docx find-artifacts unpacked/ --verbose
   ```

   Always run `docx validate` when fields or citations are present. Confirm that field `begin`, `separate`, and `end` markers remain balanced. Review artifact fixes before using `--fix`.

5. Pack the result.

   ```bash
   docx pack unpacked/ output.docx --original document.docx
   ```

   Use `--validate false` only when explicitly appropriate. Accept tracked changes only when requested.

   ```bash
   docx accept-changes tracked.docx clean.docx
   ```
