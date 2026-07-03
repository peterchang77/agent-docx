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


def test_find_artifacts_runs_cleanly_and_reports_count(tmp_path):
    unpacked = _unpack(tmp_path)

    result = run_docx("find-artifacts", str(unpacked))

    # Report-only mode exits non-zero when artifacts are found, 0 when none.
    assert "Found" in result.stdout
    assert "artifact" in result.stdout


def test_find_artifacts_verbose_shows_context(tmp_path):
    unpacked = _unpack(tmp_path)

    result = run_docx("find-artifacts", str(unpacked), "--verbose")

    assert "Found" in result.stdout


def test_find_artifacts_fix_removes_and_exits_zero(tmp_path):
    unpacked = _unpack(tmp_path)

    result = run_docx("find-artifacts", str(unpacked), "--fix")

    assert result.returncode == 0, result.stdout + result.stderr


def test_find_artifacts_missing_doc_errors(tmp_path):
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()

    result = run_docx("find-artifacts", str(empty_dir))

    assert result.returncode != 0
    assert "Error" in result.stdout
