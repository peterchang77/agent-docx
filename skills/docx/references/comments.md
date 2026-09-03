# Comments

`docx comment` creates the comment parts, but it does not select or anchor the commented document text. After running it, place the printed markers in `word/document.xml` around the intended content.

## Add a comment

```bash
docx comment unpacked/ 0 "Review this section" --author Peter --initials P
```

The comment ID must be unique. Escape text for XML before passing it to the command, especially `&`, `<`, and `>`. Use direct statements or questions, not labels such as `TODO:` or `Note:`.

The command prints a marker template. Insert the markers as direct children of the target `<w:p>`, never inside a `<w:r>`:

```xml
<w:commentRangeStart w:id="0"/>
<w:r>...</w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
```

Preserve the target runs and place the range around the exact text being commented. A reply uses a new unique ID and a valid parent:

```bash
docx comment unpacked/ 1 "Follow-up" --parent 0
```

Use the nested marker template printed by the command for replies. Comment range markers, comment references, and IDs must remain consistent. Validate after adding the markers and after packing.
