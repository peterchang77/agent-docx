import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

FIXTURE = Path(__file__).parent / "fixtures" / "sample.docx"

SOFFICE_AVAILABLE = shutil.which("soffice") is not None


def run_docx(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, "-m", "agent_docx.cli", *args],
        capture_output=True,
        text=True,
    )


def test_accept_changes_missing_input_errors(tmp_path):
    result = run_docx(
        "accept-changes",
        str(tmp_path / "does_not_exist.docx"),
        str(tmp_path / "out.docx"),
    )

    assert result.returncode != 0
    assert "Error" in result.stdout


def test_accept_changes_rejects_non_docx_input(tmp_path):
    bad_file = tmp_path / "not_a_docx.txt"
    bad_file.write_text("hello")

    result = run_docx("accept-changes", str(bad_file), str(tmp_path / "out.docx"))

    assert result.returncode != 0
    assert "Error" in result.stdout


@pytest.mark.skipif(
    not SOFFICE_AVAILABLE, reason="LibreOffice (soffice) not installed"
)
def test_accept_changes_produces_clean_docx(tmp_path):
    output = tmp_path / "clean.docx"

    result = run_docx("accept-changes", str(FIXTURE), str(output))

    assert result.returncode == 0, result.stdout + result.stderr
    assert output.exists()
    assert zipfile.is_zipfile(output)
