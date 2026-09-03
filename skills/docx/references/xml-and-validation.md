# XML and validation details

## Unpacking and source of truth

Unpack into a fresh directory. The command extracts and rewrites XML, merges adjacent identically formatted runs, simplifies adjacent tracked changes, and converts smart quotes to XML entities by default. Disable those transformations when their original boundaries matter:

```bash
docx unpack document.docx unpacked/ --merge-runs false --simplify-redlines false
```

Read `unpacked/word/document.xml` directly for content, formatting, fields, and tracked changes. Do not pack a directory containing stale files from another document.

## Direct XML edits

Prefer `docx edit`. For structural edits, replace complete `<w:r>` elements, preserve `<w:rPr>`, and make the match unique. Keep `w:fldChar` begin, separate, and end markers, `w:fldData`, and `w:instrText` intact. Do not edit across fields unless deliberately using `--allow-field-edit` and validating the result.

Preserve existing authors' tracked changes when adding or rejecting changes. Tracked-change and comment IDs must be unique in the document. Comment markers are paragraph children, not run children. A full paragraph deletion must include the paragraph mark.

## Validation modes

Use the original document whenever possible:

```bash
docx pack unpacked/ --check --original document.docx
docx pack unpacked/ --diff --original document.docx
docx validate unpacked/ --original document.docx
```

- `validate` without `--original` checks schemas but skips redlining comparison.
- `pack --check` validates without creating an output file.
- `pack --diff` compares untracked text and field structure; it does not prove that all formatting is unchanged.
- `pack` with `--original` validates with auto-repair by default. Auto-repair mutates the unpacked XML, so inspect the result and rerun checks.
- Do not disable validation unless there is a documented reason.

Inspect formatting artifacts:

```bash
docx find-artifacts unpacked/ --verbose
docx find-artifacts unpacked/ --fix
```

Review any fixes and validate again. `--fix` mutates the working directory. When fields or citations are present, confirm field markers remain balanced and always run validation with `--original`.

## Finalization

Pack only after validation:

```bash
docx pack unpacked/ output.docx --original document.docx
```

`docx accept-changes tracked.docx clean.docx` requires LibreOffice and permanently accepts all tracked changes in the output. Use it only after review and explicit request, then validate the resulting document.
