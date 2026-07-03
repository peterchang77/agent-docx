import subprocess
import sys


def test_help_exits_zero():
    result = subprocess.run(
        [sys.executable, "-m", "agent_docx.cli", "--help"],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0
    assert "docx" in result.stdout.lower() or "usage" in result.stdout.lower()


def test_docx_console_script_help():
    result = subprocess.run(["docx", "--help"], capture_output=True, text=True)
    assert result.returncode == 0
