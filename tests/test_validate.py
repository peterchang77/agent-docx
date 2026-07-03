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


def test_validate_unpacked_dir_passes(tmp_path):
    unpacked = tmp_path / "unpacked"
    run_docx("unpack", str(FIXTURE), str(unpacked))

    result = run_docx("validate", str(unpacked), "--original", str(FIXTURE))

    assert result.returncode == 0, result.stdout + result.stderr
    assert "All validations PASSED!" in result.stdout


def test_validate_packed_file_directly(tmp_path):
    # Validate can also accept a packed .docx directly (no unpack needed).
    result = run_docx("validate", str(FIXTURE))

    assert result.returncode == 0, result.stdout + result.stderr
    assert "PASSED" in result.stdout


def test_validate_missing_path_errors(tmp_path):
    result = run_docx("validate", str(tmp_path / "nonexistent"))

    assert result.returncode != 0
    assert "Error" in result.stdout
