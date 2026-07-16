---
name: docx
description: "Use this skill when the user wants to edit, revise, or redline an existing Word document (.docx file) with tracked changes."
---

# Editing .docx Files with Tracked Changes

A .docx file is a ZIP archive of XML files. Editing workflow: unpack → edit XML → validate → repack.

This skill uses the `docx` CLI (install via `uv pip install -e .` from the package root; see the project README for setup). All commands below assume `docx` is on `PATH`.

## Step 1: Unpack

```bash
docx unpack document.docx unpacked/
```

Extracts XML, pretty-prints, merges adjacent runs, simplifies redlines, and converts smart quotes to XML entities. Use `--merge-runs false` to preserve original run boundaries.

**Reading content for planning edits:** read `unpacked/word/document.xml` directly — it is pretty-printed and is the ground truth for both content and formatting. Do not rely on a separate text-extraction pass; anything not visible in the XML (e.g. exact tracked-change state, run boundaries) won't be visible in a lossy text conversion either.

## Step 2: Edit

Use **"Peter" as the author** for tracked changes and comments, unless the user explicitly requests a different name.

**Writing conventions (author preferences):**

- **Do not use em-dashes** (`—` / `&#x2014;`) in inserted or edited text. Rephrase with commas, parentheses, or separate sentences instead. (En-dashes in ranges/compounds are fine.)
- **Do not prefix comments** with labels like "Open item:", "Note:", or "TODO:". Write the comment as a plain, direct statement or question.

### Approach A: `docx edit` (preferred for text edits)

Automates tracked change insertion — finds text across runs, splits at boundaries, wraps in proper `<w:del>`/`<w:ins>` markup with auto-incrementing `w:id` values. No orphaned tags.

```bash
# Replace text
docx edit unpacked/ --find "old text" --replace "new text"

# Delete text
docx edit unpacked/ --find "remove this" --mode delete

# Insert after anchor text
docx edit unpacked/ --find "anchor" --replace " (added)" --mode insert

# Disambiguate when text appears in multiple paragraphs
docx edit unpacked/ --find "ambiguous" --replace "new" --paragraph 5
```

Options: `--author NAME`, `--date ISO_TIMESTAMP`, `--paragraph N` (0-indexed), `--allow-field-edit`.

If the target text appears in multiple paragraphs and `--paragraph` is not specified, the tool lists matching paragraph indices and exits with an error.

#### Word fields (EndNote citations, cross-references, TOC) — important

A Word field is a sequence of `<w:fldChar w:fldCharType="begin">` ... `<w:instrText>` ... `<w:fldChar w:fldCharType="end">` markers. Word routinely packs these markers into the **same run as adjacent text**, which is why naive edits silently corrupt citations.

`docx edit` is now **field-aware**: it preserves field markers when splitting runs, and **refuses** any edit whose anchor overlaps or spans a field (for example, an anchor that visually includes a citation's `[69]` glyph). When a field is detected you will see an error like:

> Refusing to edit: the target text spans a Word field ... Edit around the field, or pass --allow-field-edit.

- **Edit *around* the field** (preferred): choose an anchor that does not include the citation glyph, so the field markers are never inside the matched text.
- **`--allow-field-edit`** overrides the guard when you genuinely intend to touch a field. The edit still runs a before/after field-integrity self-check and appends a `WARNING` if a field was touched — verify the result.
- For manual XML edits near a field, preserve every `<w:fldChar>`/`<w:fldData>`/`<w:instrText>` exactly; do not let a run split drop them.

### Approach B: Direct XML editing (for complex structural changes)

Edit `unpacked/word/document.xml` directly with the Edit tool. Required for: paragraph-level deletions, rejecting/restoring other authors' changes, adding comments, or any edit that `docx edit` can't express.

**CRITICAL rules for manual XML editing:**

1. **Replace entire `<w:r>` blocks.** Always include the full `<w:r>...</w:r>` open and close tags in both `old_str` and `new_str`. Never inject tracked-change tags inside a run — this leaves orphaned `</w:r>` closing tags that break validation.

2. **Preserve `<w:rPr>` formatting.** Copy the original run's `<w:rPr>` block into your tracked change runs to maintain bold, font size, etc.

3. **Ensure unique matches.** Many paragraphs share identical XML structure. Include enough surrounding context (full `<w:r>` block, neighboring runs, `<w:pPr>`) to ensure `old_str` matches exactly once. Use Grep to confirm before editing.

4. **Re-verify after each edit.** After modifying XML, nearby content may have shifted. Grep to confirm `old_str` still matches before the next edit in the same region.

5. **Use smart quotes.** When adding text with apostrophes or quotes, use XML entities:

| Entity | Character |
|--------|-----------|
| `&#x2018;` | ' (left single) |
| `&#x2019;` | ' (right single / apostrophe) |
| `&#x201C;` | " (left double) |
| `&#x201D;` | " (right double) |

```xml
<w:t>Here&#x2019;s a quote: &#x201C;Hello&#x201D;</w:t>
```

### Adding Comments

Use `docx comment` to create comment entries (text must be pre-escaped XML), then add markers to document.xml:

```bash
docx comment unpacked/ 0 "Comment text with &amp; and &#x2019;"
docx comment unpacked/ 1 "Reply text" --parent 0
docx comment unpacked/ 0 "Text" --author "Custom Author"
```

**CRITICAL: `<w:commentRangeStart>` and `<w:commentRangeEnd>` are siblings of `<w:r>`, never inside `<w:r>`.**

```xml
<!-- Comment wrapping content -->
<w:commentRangeStart w:id="0"/>
<w:r><w:t>commented text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>

<!-- Comment with nested reply -->
<w:commentRangeStart w:id="0"/>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t>text</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>
```

## Step 3: Validate & Pack

### Pre-pack checks

```bash
# Dry-run validation (schema + redlining, no ZIP created)
docx pack unpacked/ --check --original document.docx

# Lightweight text diff — catches untracked modifications
docx pack unpacked/ --diff --original document.docx

# Both together
docx pack unpacked/ --check --diff --original document.docx
```

`--diff` compares document text after stripping the editing author's tracked changes against the original. Catches accidental untracked edits before packing.

**`--diff` alone is NOT enough when the document contains Word fields** (EndNote citations, cross-references, TOC). Dropping a field's markers while keeping its display glyph (e.g. `[69]`) produces **identical text** but a corrupted citation, and the text diff is blind to it. `--diff` now also reports a *field/citation fingerprint change*, and `--check` runs a full field-integrity validator — but **always do the explicit field check below after editing near a field.**

#### Mandatory post-edit citation/field check

After any edit in a document that contains fields, verify the fields survived before packing:

```bash
# Field integrity is part of validation (begin/separate/end must balance):
docx validate unpacked/ --original document.docx

# Or a quick manual guard on raw counts — these must be UNCHANGED vs the original
# (compare the three numbers before and after your edits):
python -c "import re,sys; x=open('unpacked/word/document.xml').read(); \
  print({t: x.count(f'fldCharType=\"{t}\"') for t in ('begin','separate','end')})"
```

If `begin`/`separate`/`end` counts dropped, or `validate` reports a field-integrity violation, a citation was damaged — restore from the original and edit *around* the field.

### Scan for formatting artifacts

```bash
docx find-artifacts unpacked/              # report only
docx find-artifacts unpacked/ --verbose    # with XML context per artifact
docx find-artifacts unpacked/ --fix        # auto-remove
```

Detects: `<w:r>` with only `<w:br/>`, empty styled runs, whitespace-only `<w:t>` runs. These are common in Word documents and usually safe to remove. **Caveat:** the whitespace-only-text heuristic can occasionally flag a run that is a single meaningful space between words (not decorative cruft) — review `--fix` output before trusting it blindly on a document with unusual run splitting.

### Pack

```bash
docx pack unpacked/ output.docx --original document.docx
```

Validates with auto-repair, condenses XML formatting, and creates the DOCX. Use `--validate false` to skip validation.

**Auto-repair fixes:** `durableId` overflow, missing `xml:space="preserve"` on `<w:t>` with whitespace.

**Auto-repair does NOT fix:** malformed XML, invalid element nesting, missing relationships, schema violations.

### Optional: Accept all tracked changes

To produce a clean copy with all tracked changes accepted (requires LibreOffice installed and on `PATH`):

```bash
docx accept-changes tracked.docx clean.docx
```

This is outside the core edit/comment loop — use it only when the user explicitly wants a finalized, redline-free copy.

---

## Tracked Changes XML Reference

### Insertion

```xml
<w:ins w:id="1" w:author="Peter" w:date="2025-01-01T00:00:00Z">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

### Deletion

Inside `<w:del>`: use `<w:delText>` instead of `<w:t>`, and `<w:delInstrText>` instead of `<w:instrText>`.

```xml
<w:del w:id="2" w:author="Peter" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

### Replacement (delete + insert as siblings)

```xml
<w:r><w:t>before text </w:t></w:r>
<w:del w:id="1" w:author="Peter" w:date="...">
  <w:r><w:rPr><!-- original formatting --></w:rPr><w:delText>old</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Peter" w:date="...">
  <w:r><w:rPr><!-- original formatting --></w:rPr><w:t>new</w:t></w:r>
</w:ins>
<w:r><w:t> after text</w:t></w:r>
```

### Deleting an entire paragraph

When removing ALL content from a paragraph, also mark the paragraph mark as deleted so it merges with the next paragraph. Without this, accepting changes leaves an empty paragraph.

```xml
<w:p>
  <w:pPr>
    <w:rPr>
      <w:del w:id="1" w:author="Peter" w:date="2025-01-01T00:00:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:del w:id="2" w:author="Peter" w:date="2025-01-01T00:00:00Z">
    <w:r><w:delText>Entire paragraph content being deleted...</w:delText></w:r>
  </w:del>
</w:p>
```

### Rejecting another author's insertion

Nest deletion inside their insertion:

```xml
<w:ins w:author="Jane" w:id="5">
  <w:del w:author="Peter" w:id="10">
    <w:r><w:delText>their inserted text</w:delText></w:r>
  </w:del>
</w:ins>
```

### Restoring another author's deletion

Add insertion after (don't modify their deletion):

```xml
<w:del w:author="Jane" w:id="5">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
<w:ins w:author="Peter" w:id="10">
  <w:r><w:t>deleted text</w:t></w:r>
</w:ins>
```

### `w:id` values

`w:id` is `ST_DecimalNumber` (integer). IDs are shared across `<w:ins>`, `<w:del>`, `<w:comment>`, `<w:bookmarkStart>`, etc. — they must be unique within the document. `docx edit` auto-increments from the max existing ID.

## Schema Compliance

- **Element order in `<w:pPr>`**: `<w:pStyle>`, `<w:numPr>`, `<w:spacing>`, `<w:ind>`, `<w:jc>`, `<w:rPr>` last
- **Whitespace**: `xml:space="preserve"` required on `<w:t>` with leading/trailing spaces
- **RSIDs**: Must be 8-digit hex (e.g., `00AB1234`)
