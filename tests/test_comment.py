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


def test_comment_adds_entry_to_comments_xml(tmp_path):
    unpacked = _unpack(tmp_path)

    result = run_docx("comment", str(unpacked), "1", "New comment text")

    assert result.returncode == 0, result.stdout + result.stderr
    assert "Added comment 1" in result.stdout

    comments_xml = (unpacked / "word" / "comments.xml").read_text()
    assert 'w:id="1"' in comments_xml
    assert "New comment text" in comments_xml


def test_comment_reply_nests_under_parent(tmp_path):
    unpacked = _unpack(tmp_path)
    run_docx("comment", str(unpacked), "1", "Parent comment")

    result = run_docx("comment", str(unpacked), "2", "Reply text", "--parent", "1")

    assert result.returncode == 0, result.stdout + result.stderr
    assert "Added reply 2" in result.stdout

    ext_xml = (unpacked / "word" / "commentsExtended.xml").read_text()
    assert "paraIdParent" in ext_xml


def test_comment_then_pack_check_passes(tmp_path):
    unpacked = _unpack(tmp_path)
    run_docx("comment", str(unpacked), "1", "New comment text")

    result = run_docx("pack", str(unpacked), "--check", "--original", str(FIXTURE))

    assert result.returncode == 0, result.stdout + result.stderr
    assert "PASSED" in result.stdout


def test_comment_missing_word_dir_errors(tmp_path):
    empty_dir = tmp_path / "not_unpacked"
    empty_dir.mkdir()

    result = run_docx("comment", str(empty_dir), "1", "text")

    assert result.returncode != 0
    assert "Error" in result.stdout


def test_comment_reply_to_nonexistent_parent_errors(tmp_path):
    unpacked = _unpack(tmp_path)

    result = run_docx("comment", str(unpacked), "1", "text", "--parent", "999")

    assert result.returncode != 0
    assert "Error" in result.stdout
