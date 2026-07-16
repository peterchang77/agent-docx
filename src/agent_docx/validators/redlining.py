"""
Validator for tracked changes in Word documents.
"""

import subprocess
import tempfile
import zipfile
from pathlib import Path


class RedliningValidator:

    def __init__(self, unpacked_dir, original_docx, verbose=False, author="Peter"):
        self.unpacked_dir = Path(unpacked_dir)
        self.original_docx = Path(original_docx)
        self.verbose = verbose
        self.author = author
        self.namespaces = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        }

    def repair(self) -> int:
        return 0

    def validate(self):
        modified_file = self.unpacked_dir / "word" / "document.xml"
        if not modified_file.exists():
            print(f"FAILED - Modified document.xml not found at {modified_file}")
            return False

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)

            try:
                with zipfile.ZipFile(self.original_docx, "r") as zip_ref:
                    zip_ref.extractall(temp_path)
            except Exception as e:
                print(f"FAILED - Error unpacking original docx: {e}")
                return False

            original_file = temp_path / "word" / "document.xml"
            if not original_file.exists():
                print(
                    f"FAILED - Original document.xml not found in {self.original_docx}"
                )
                return False

            try:
                import xml.etree.ElementTree as ET

                modified_tree = ET.parse(modified_file)
                modified_root = modified_tree.getroot()
                original_tree = ET.parse(original_file)
                original_root = original_tree.getroot()
            except ET.ParseError as e:
                print(f"FAILED - Error parsing XML files: {e}")
                return False

            self._remove_author_tracked_changes(original_root)
            self._remove_author_tracked_changes(modified_root)

            modified_text = self._extract_text_content(modified_root)
            original_text = self._extract_text_content(original_root)

            if modified_text != original_text:
                error_message = self._generate_detailed_diff(
                    original_text, modified_text
                )
                print(error_message)
                return False

            # Field fingerprint comparison (P4): display text can be identical
            # even when a field wrapper (and its instrText/fldData) was dropped.
            # This is the exact blind spot of the text-only diff above.
            field_ok, field_msg = self._compare_field_fingerprints(
                modified_root, original_root
            )
            if not field_ok:
                print(field_msg)
                return False

            if self.verbose:
                print(f"PASSED - All changes by {self.author} are properly tracked")
            return True

    def _generate_detailed_diff(self, original_text, modified_text):
        error_parts = [
            f"FAILED - Document text doesn't match after removing {self.author}'s tracked changes",
            "",
            "Likely causes:",
            "  1. Modified text inside another author's <w:ins> or <w:del> tags",
            "  2. Made edits without proper tracked changes",
            "  3. Didn't nest <w:del> inside <w:ins> when deleting another's insertion",
            "",
            "For pre-redlined documents, use correct patterns:",
            "  - To reject another's INSERTION: Nest <w:del> inside their <w:ins>",
            "  - To restore another's DELETION: Add new <w:ins> AFTER their <w:del>",
            "",
        ]

        git_diff = self._get_git_word_diff(original_text, modified_text)
        if git_diff:
            error_parts.extend(["Differences:", "============", git_diff])
        else:
            error_parts.append("Unable to generate word diff (git not available)")

        return "\n".join(error_parts)

    def _get_git_word_diff(self, original_text, modified_text):
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)

                original_file = temp_path / "original.txt"
                modified_file = temp_path / "modified.txt"

                original_file.write_text(original_text, encoding="utf-8")
                modified_file.write_text(modified_text, encoding="utf-8")

                result = subprocess.run(
                    [
                        "git",
                        "diff",
                        "--word-diff=plain",
                        "--word-diff-regex=.",
                        "-U0",
                        "--no-index",
                        str(original_file),
                        str(modified_file),
                    ],
                    capture_output=True,
                    text=True,
                )

                if result.stdout.strip():
                    lines = result.stdout.split("\n")
                    content_lines = []
                    in_content = False
                    for line in lines:
                        if line.startswith("@@"):
                            in_content = True
                            continue
                        if in_content and line.strip():
                            content_lines.append(line)

                    if content_lines:
                        return "\n".join(content_lines)

                result = subprocess.run(
                    [
                        "git",
                        "diff",
                        "--word-diff=plain",
                        "-U0",
                        "--no-index",
                        str(original_file),
                        str(modified_file),
                    ],
                    capture_output=True,
                    text=True,
                )

                if result.stdout.strip():
                    lines = result.stdout.split("\n")
                    content_lines = []
                    in_content = False
                    for line in lines:
                        if line.startswith("@@"):
                            in_content = True
                            continue
                        if in_content and line.strip():
                            content_lines.append(line)
                    return "\n".join(content_lines)

        except (subprocess.CalledProcessError, FileNotFoundError, Exception):
            pass

        return None

    def _remove_author_tracked_changes(self, root):
        ins_tag = f"{{{self.namespaces['w']}}}ins"
        del_tag = f"{{{self.namespaces['w']}}}del"
        author_attr = f"{{{self.namespaces['w']}}}author"

        for parent in root.iter():
            to_remove = []
            for child in parent:
                if child.tag == ins_tag and child.get(author_attr) == self.author:
                    to_remove.append(child)
            for elem in to_remove:
                parent.remove(elem)

        deltext_tag = f"{{{self.namespaces['w']}}}delText"
        t_tag = f"{{{self.namespaces['w']}}}t"

        for parent in root.iter():
            to_process = []
            for child in parent:
                if child.tag == del_tag and child.get(author_attr) == self.author:
                    to_process.append((child, list(parent).index(child)))

            for del_elem, del_index in reversed(to_process):
                for elem in del_elem.iter():
                    if elem.tag == deltext_tag:
                        elem.tag = t_tag

                for child in reversed(list(del_elem)):
                    parent.insert(del_index, child)
                parent.remove(del_elem)

    def _extract_field_fingerprint(self, root):
        """Field fingerprint: counts of fldChar types + sorted instrText codes.

        Captures Word field structure (citations, cross-refs, TOC, ...) so that
        a dropped field is detected even when its display text is unchanged -
        the blind spot of a text-only diff. Returns
        ``(begin, separate, end, tuple(sorted instrText codes))``.
        """
        w = self.namespaces["w"]
        fldchar_tag = f"{{{w}}}fldChar"
        type_attr = f"{{{w}}}fldCharType"
        counts = {"begin": 0, "separate": 0, "end": 0}
        instr = []
        for el in root.iter():
            if el.tag == fldchar_tag:
                ftype = el.get(type_attr)
                if ftype in counts:
                    counts[ftype] += 1
            elif el.tag in (f"{{{w}}}instrText", f"{{{w}}}delInstrText"):
                instr.append((el.text or "").strip())
        return (
            counts["begin"],
            counts["separate"],
            counts["end"],
            tuple(sorted(i for i in instr if i)),
        )

    def _compare_field_fingerprints(self, modified_root, original_root):
        """Compare field fingerprints; return (ok, message_or_None)."""
        modified = self._extract_field_fingerprint(modified_root)
        original = self._extract_field_fingerprint(original_root)
        if modified == original:
            return True, None
        msg = (
            f"FAILED - Field/citation fingerprint changed after removing "
            f"{self.author}'s tracked changes.\n"
            f"  fldChar begin/separate/end: {original[:3]} -> {modified[:3]}\n"
            f"  field-instruction count: {len(original[3])} -> {len(modified[3])}\n"
            f"  A field (e.g. an EndNote citation or cross-reference) was likely "
            f"dropped or damaged during editing. Note: --diff alone does not "
            f"guarantee citations survived."
        )
        return False, msg

    def _extract_text_content(self, root):
        p_tag = f"{{{self.namespaces['w']}}}p"
        t_tag = f"{{{self.namespaces['w']}}}t"

        paragraphs = []
        for p_elem in root.findall(f".//{p_tag}"):
            text_parts = []
            for t_elem in p_elem.findall(f".//{t_tag}"):
                if t_elem.text:
                    text_parts.append(t_elem.text)
            paragraph_text = "".join(text_parts)
            if paragraph_text:
                paragraphs.append(paragraph_text)

        return "\n".join(paragraphs)
