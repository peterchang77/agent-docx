import subprocess
import sys
import zipfile
from pathlib import Path

FIXTURE = Path(__file__).parent / "fixtures" / "sample.docx"


def run_docx(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, "-m", "agent_docx.cli", *args],
        capture_output=True,
        text=True,
    )


def test_pack_check_passes_on_unmodified_unpack(tmp_path):
    unpacked = tmp_path / "unpacked"
    run_docx("unpack", str(FIXTURE), str(unpacked))

    result = run_docx("pack", str(unpacked), "--check", "--original", str(FIXTURE))

    assert result.returncode == 0, result.stdout + result.stderr
    assert "PASSED" in result.stdout


def test_pack_round_trip_produces_valid_docx(tmp_path):
    unpacked = tmp_path / "unpacked"
    output = tmp_path / "repacked.docx"
    run_docx("unpack", str(FIXTURE), str(unpacked))

    result = run_docx(
        "pack", str(unpacked), str(output), "--original", str(FIXTURE)
    )

    assert result.returncode == 0, result.stdout + result.stderr
    assert output.exists()
    assert zipfile.is_zipfile(output)

    # Repacked file should itself pass validation.
    validate_result = run_docx("validate", str(output))
    assert validate_result.returncode == 0, validate_result.stdout


def test_pack_requires_output_file_without_check_or_diff(tmp_path):
    unpacked = tmp_path / "unpacked"
    run_docx("unpack", str(FIXTURE), str(unpacked))

    result = run_docx("pack", str(unpacked))

    assert result.returncode != 0
    assert "Error" in result.stdout


def test_pack_rejects_non_docx_output(tmp_path):
    unpacked = tmp_path / "unpacked"
    run_docx("unpack", str(FIXTURE), str(unpacked))

    result = run_docx("pack", str(unpacked), str(tmp_path / "out.pptx"))

    assert result.returncode != 0
    assert "Error" in result.stdout
