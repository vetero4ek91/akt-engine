"""Regression tests for the expert-review preflight (no Node.js or LibreOffice)."""
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import fast_build


class ExpertReviewTests(unittest.TestCase):
    def setUp(self):
        workspace = tempfile.TemporaryDirectory()
        self.addCleanup(workspace.cleanup)
        self.generator = Path(workspace.name) / "generator.js"
        self.generator.write_text("test input", encoding="utf-8")
        self.report_path = self.generator.with_name("review.json")
        root = Path(fast_build.__file__).resolve().parents[1]
        self.report = {
            "generator_sha256": fast_build.sha256_file(self.generator),
            "sources": {name: fast_build.sha256_file(root / name)
                        for name in fast_build.EXPERT_SOURCES},
            "checks": {name: f"Reviewed {name} using photo 2."
                       for name in fast_build.EXPERT_CHECKS},
            "open_errors": [],
        }

    def write_report(self):
        self.report_path.write_text(json.dumps(self.report), encoding="utf-8")

    def call_main(self, *flags):
        argv = ["fast_build.py", "--generator", str(self.generator),
                "--output", str(self.generator.with_suffix(".pdf")), *flags]
        with patch("sys.argv", argv):
            fast_build.main()

    def assert_blocked(self, message, *flags):
        with patch.object(fast_build, "read_js_data") as read_js, \
                patch.object(fast_build.argparse.ArgumentParser, "error",
                             side_effect=ValueError) as error:
            with self.assertRaisesRegex(ValueError, message):
                self.call_main(*flags)
            error.assert_called_once()
            read_js.assert_not_called()

    def test_production_requires_report(self):
        self.assert_blocked("--expert-review")

    def test_unreadable_or_invalid_report_blocks(self):
        flags = ("--expert-review", str(self.report_path))
        self.assert_blocked("JSON-отчёт", *flags)
        self.report_path.write_text("{", encoding="utf-8")
        self.assert_blocked("JSON-отчёт", *flags)

    def test_open_errors_block(self):
        self.report["open_errors"] = ["Подкрылок требует повторной сверки"]
        self.write_report()
        self.assert_blocked("незакрытые ошибки", "--expert-review", str(self.report_path))

    def test_missing_check_blocks(self):
        del self.report["checks"]["operational_defects"]
        self.write_report()
        self.assert_blocked("operational_defects", "--expert-review", str(self.report_path))

    def test_changed_input_blocks(self):
        self.write_report()
        self.generator.write_text("changed input", encoding="utf-8")
        self.assert_blocked("другой редакции", "--expert-review", str(self.report_path))

    def test_changed_knowledge_blocks(self):
        self.report["sources"]["CASES.md"] = "0" * 64
        self.write_report()
        self.assert_blocked("файлам экспертной базы", "--expert-review", str(self.report_path))

    def test_valid_report_reaches_js_preflight(self):
        self.write_report()
        self.assert_reaches_js_preflight("--expert-review", str(self.report_path))

    def test_explicit_test_mode_can_omit_report(self):
        self.assert_reaches_js_preflight("--test")

    def assert_reaches_js_preflight(self, *flags):
        with patch.object(fast_build, "read_js_data",
                          side_effect=RuntimeError("JS preflight")) as read_js:
            with self.assertRaisesRegex(RuntimeError, "JS preflight"):
                self.call_main(*flags)
            read_js.assert_called_once_with(self.generator.resolve())


if __name__ == "__main__":
    unittest.main()
