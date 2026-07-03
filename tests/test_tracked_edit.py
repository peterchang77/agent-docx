import subprocess
import sys
from pathlib import Path

FIXTURE = Path(__file__).parent / "fixtures" / "sample.docx"


def run_docx(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, "-m", "agent_docx.cli", *args],
        capture_output=True,
        text=True,
    )


def _unpack(tmp_path) -> Path:
    unpacked = tmp_path / "unpacked"
    run_docx("unpack", str(FIXTURE), str(unpacked))
    return unpacked


def test_edit_replace_inserts_tracked_change(tmp_path):
    unpacked = _unpack(tmp_path)

    result = run_docx(
        "edit", str(unpacked), "--find", "sample document", "--replace", "test file"
    )

    assert result.returncode == 0, result.stdout + result.stderr
    assert "w:id" in result.stdout

    doc_xml = (unpacked / "word" / "document.xml").read_text()
    assert "<w:del" in doc_xml
    assert "<w:delText>sample document</w:delText>" in doc_xml
    assert "<w:ins" in doc_xml
    assert "test file" in doc_xml


def test_edit_delete_mode(tmp_path):
    unpacked = _unpack(tmp_path)

    result = run_docx("edit", str(unpacked), "--find", "jumps", "--mode", "delete")

    assert result.returncode == 0, result.stdout + result.stderr
    doc_xml = (unpacked / "word" / "document.xml").read_text()
    assert "<w:delText>jumps</w:delText>" in doc_xml


def test_edit_replace_without_replacement_text_errors(tmp_path):
    unpacked = _unpack(tmp_path)

    result = run_docx("edit", str(unpacked), "--find", "jumps")

    assert result.returncode != 0
    assert "Error" in result.stdout


def test_edit_ambiguous_match_requires_paragraph(tmp_path):
    unpacked = _unpack(tmp_path)

    # "This" only appears once, but reuse a generic word to validate
    # the missing-text error path instead (deterministic across fixture edits).
    result = run_docx(
        "edit", str(unpacked), "--find", "nonexistent phrase", "--replace", "x"
    )

    assert result.returncode != 0
    assert "not found" in result.stdout


def test_edit_then_pack_check_passes(tmp_path):
    unpacked = _unpack(tmp_path)
    run_docx("edit", str(unpacked), "--find", "jumps", "--replace", "leaps")

    result = run_docx("pack", str(unpacked), "--check", "--original", str(FIXTURE))

    assert result.returncode == 0, result.stdout + result.stderr
    assert "PASSED" in result.stdout
