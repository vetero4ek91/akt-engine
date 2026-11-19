"""Regression tests for expert review and autonomous PDF pagination."""
from importlib.util import find_spec
import json
import os
from pathlib import Path
import shutil
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



@unittest.skipUnless(
    all(shutil.which(name) for name in ("node", "zip", "unzip", "soffice")) and find_spec("fitz"),
    "PDF pagination tests require Node.js, zip, unzip, LibreOffice and PyMuPDF",
)
class PaginationTests(unittest.TestCase):
    def render_act(self, count):
        workspace = tempfile.TemporaryDirectory()
        self.addCleanup(workspace.cleanup)
        workdir = Path(workspace.name)
        generator = Path(fast_build.__file__).resolve().parents[1] / "Favorit_AKT_autonomous.js"
        source = generator.read_text(encoding="utf-8")
        claim = "2097852"
        damages = [
            {
                "detail": f"Пробная деталь {i:02d}",
                "description": (
                    f"Контрольное описание позиции {i:02d}. "
                    "Деформация в средней части детали с повреждением лакокрасочного покрытия."
                ),
            }
            for i in range(1, count + 1)
        ]
        self.assertIn("const damages = [];", source)
        self.assertIn("claimNo:''", source)
        source = source.replace("claimNo:''", f"claimNo:'{claim}'", 1)
        source = source.replace(
            "const damages = [];",
            "const damages = " + json.dumps(damages, ensure_ascii=False) + ";",
            1,
        )
        test_generator = workdir / generator.name
        test_generator.write_text(source, encoding="utf-8")
        docx = workdir / "act.docx"
        fast_build.run(["node", str(test_generator)], 25,
                       env=dict(os.environ, AKT_TEST="1", AKT_OUT=str(docx)))
        profile = (workdir / "lo-profile").as_uri()
        fast_build.run(["soffice", f"-env:UserInstallation={profile}", "--headless",
                        "--convert-to", "pdf", "--outdir", str(workdir), str(docx)], 35)
        return workdir / "act.pdf", claim, damages

    def test_seventeen_damage_rows_and_conclusion_fit_four_pages(self):
        import fitz

        pdf, claim, damages = self.render_act(17)
        self.assertEqual(fast_build.check_pdf(pdf, claim, "", damages, 17), 4)
        with fitz.open(pdf) as document:
            for damage in damages:
                detail = fast_build.norm(damage["detail"])
                self.assertIn(detail, fast_build.norm(document[2].get_text()))
                self.assertNotIn(detail, fast_build.norm(document[3].get_text()))
            document.new_page()
            extra = pdf.with_name("extra-page.pdf")
            document.save(extra)
        with self.assertRaisesRegex(RuntimeError, "5 стр., ожидается 4"):
            fast_build.check_pdf(extra, claim, "", damages, 17)

    def test_eighteenth_damage_moves_to_continuation_page(self):
        import fitz

        pdf, claim, damages = self.render_act(18)
        self.assertEqual(fast_build.check_pdf(pdf, claim, "", damages, 18), 5)
        with fitz.open(pdf) as document:
            for damage in damages[:17]:
                self.assertIn(fast_build.norm(damage["detail"]),
                              fast_build.norm(document[2].get_text()))
            eighteenth = fast_build.norm(damages[17]["detail"])
            self.assertNotIn(eighteenth, fast_build.norm(document[2].get_text()))
            self.assertIn(eighteenth, fast_build.norm(document[3].get_text()))
            self.assertNotIn(eighteenth, fast_build.norm(document[4].get_text()))


if __name__ == "__main__":
    unittest.main()
