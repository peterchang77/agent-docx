"""Generate tests/fixtures/sample.docx: a minimal but valid DOCX containing
a plain paragraph, a tracked change (insertion + deletion), and a comment.

Run manually to (re)build the fixture:
    python tests/fixtures/generate_fixture.py
"""

import zipfile
from pathlib import Path

FIXTURE_PATH = Path(__file__).parent / "sample.docx"

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>
"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

DOCUMENT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>
"""

CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Sample</dc:title>
  <dc:creator>Test Fixture</dc:creator>
</cp:coreProperties>
"""

APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>agent-docx test fixture</Application>
</Properties>
"""

# Paragraph 1: plain text.
# Paragraph 2: "The quick fox jumps." with a tracked replace of "fox" -> "brown fox"
#   (del "fox", ins "brown fox") so both <w:del> and <w:ins> appear.
# A comment anchors on the word "quick" in paragraph 2.
DOCUMENT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>This is a sample document for testing.</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t xml:space="preserve">The </w:t></w:r>
      <w:commentRangeStart w:id="0"/>
      <w:r><w:t>quick</w:t></w:r>
      <w:commentRangeEnd w:id="0"/>
      <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
      <w:r><w:t xml:space="preserve"> </w:t></w:r>
      <w:del w:id="1" w:author="Peter" w:date="2025-01-01T00:00:00Z">
        <w:r><w:delText>fox</w:delText></w:r>
      </w:del>
      <w:ins w:id="2" w:author="Peter" w:date="2025-01-01T00:00:00Z">
        <w:r><w:t>brown fox</w:t></w:r>
      </w:ins>
      <w:r><w:t xml:space="preserve"> jumps.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>
"""

COMMENTS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="0" w:author="Peter" w:date="2025-01-01T00:00:00Z" w:initials="P">
    <w:p w14:paraId="10000000" w14:textId="77777777"><w:r><w:t>Why quick specifically?</w:t></w:r></w:p>
  </w:comment>
</w:comments>
"""


def generate() -> None:
    FIXTURE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(FIXTURE_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", RELS)
        zf.writestr("docProps/core.xml", CORE_XML)
        zf.writestr("docProps/app.xml", APP_XML)
        zf.writestr("word/document.xml", DOCUMENT_XML)
        zf.writestr("word/comments.xml", COMMENTS_XML)
        zf.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
    print(f"Wrote {FIXTURE_PATH}")


if __name__ == "__main__":
    generate()
