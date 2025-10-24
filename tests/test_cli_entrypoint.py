from click.testing import CliRunner

from pdf_text_overlay.cli import main


def test_cli_version_command():
    runner = CliRunner()
    result = runner.invoke(main, ["--version"])
    assert result.exit_code == 0
    assert "pdf_text_overlay" in result.output


def test_cli_help():
    runner = CliRunner()
    result = runner.invoke(main, ["--help"])
    assert result.exit_code == 0
    assert "Usage" in result.output
