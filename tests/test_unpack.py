import subprocess
import sys
from pathlib import Path

import lxml.etree as ET

FIXTURE = Path(__file__).parent / "fixtures" / "sample.docx"


def run_docx(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, "-m", "agent_docx.cli", *args],
        capture_output=True,
        text=True,
    )


def test_unpack_extracts_and_pretty_prints(tmp_path):
    out_dir = tmp_path / "unpacked"
    result = run_docx("unpack", str(FIXTURE), str(out_dir))

    assert result.returncode == 0, result.stdout + result.stderr
    assert "Unpacked" in result.stdout

    doc_xml = out_dir / "word" / "document.xml"
    assert doc_xml.exists()

    # Pretty-printed XML should span multiple lines.
    assert len(doc_xml.read_text().splitlines()) > 1

    # Should still parse as valid XML.
    tree = ET.parse(str(doc_xml))
    assert tree.getroot() is not None


def test_unpack_merges_runs_and_reports_message(tmp_path):
    out_dir = tmp_path / "unpacked"
    result = run_docx("unpack", str(FIXTURE), str(out_dir))

    assert "merged" in result.stdout.lower()
    assert "simplified" in result.stdout.lower()


def test_unpack_missing_file_errors(tmp_path):
    out_dir = tmp_path / "unpacked"
    result = run_docx("unpack", str(tmp_path / "does_not_exist.docx"), str(out_dir))

    assert result.returncode != 0
    assert "Error" in result.stdout


def test_unpack_rejects_non_docx(tmp_path):
    bad_file = tmp_path / "not_a_docx.txt"
    bad_file.write_text("hello")
    out_dir = tmp_path / "unpacked"

    result = run_docx("unpack", str(bad_file), str(out_dir))

    assert result.returncode != 0
    assert "Error" in result.stdout
