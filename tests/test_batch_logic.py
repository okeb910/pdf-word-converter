import tempfile
import threading
import unittest
from pathlib import Path

from batch_logic import deduplicate_paths, resolve_output_path, run_conversion_batch


class OutputPathTests(unittest.TestCase):
    def test_uses_source_directory_and_target_suffix(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "报告.pdf"
            source.touch()

            self.assertEqual(resolve_output_path(source, ".docx"), Path(tmp) / "报告.docx")

    def test_uses_custom_output_directory(self):
        with tempfile.TemporaryDirectory() as source_tmp, tempfile.TemporaryDirectory() as output_tmp:
            source = Path(source_tmp) / "example.docx"
            source.touch()

            self.assertEqual(
                resolve_output_path(source, "pdf", Path(output_tmp)),
                Path(output_tmp) / "example.pdf",
            )

    def test_adds_converted_and_incrementing_suffixes(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "sample.pdf"
            source.touch()
            (folder / "sample.docx").touch()

            self.assertEqual(resolve_output_path(source, ".docx"), folder / "sample_converted.docx")

            (folder / "sample_converted.docx").touch()
            (folder / "sample_converted_2.docx").touch()
            self.assertEqual(resolve_output_path(source, ".docx"), folder / "sample_converted_3.docx")

    def test_deduplicates_paths_without_changing_order(self):
        with tempfile.TemporaryDirectory() as tmp:
            first = Path(tmp) / "first.pdf"
            second = Path(tmp) / "second.pdf"
            self.assertEqual(
                deduplicate_paths([first, second, first]),
                [first.absolute(), second.absolute()],
            )


class BatchRunnerTests(unittest.TestCase):
    def _make_sources(self, folder, names):
        paths = []
        for name in names:
            path = folder / name
            path.touch()
            paths.append(path)
        return paths

    def test_runs_all_items_and_reports_overall_progress(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = self._make_sources(folder, ["a.pdf", "b.pdf"])
            progress_events = []
            status_events = []

            def converter(_source, output, progress):
                progress("处理中", 50)
                Path(output).write_bytes(b"docx")

            results = run_conversion_batch(
                sources,
                ".docx",
                None,
                converter,
                threading.Event(),
                lambda *args: progress_events.append(args),
                status_events.append,
            )

            self.assertEqual([result.status for result in results], ["success", "success"])
            self.assertEqual(progress_events[-1][3], 100)
            self.assertEqual([event.status for event in status_events].count("running"), 2)

    def test_failure_removes_partial_output_and_continues(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = self._make_sources(folder, ["bad.pdf", "good.pdf"])

            def converter(source, output, _progress):
                Path(output).write_bytes(b"partial")
                if Path(source).stem == "bad":
                    raise RuntimeError("模拟失败")

            results = run_conversion_batch(
                sources,
                ".docx",
                None,
                converter,
                threading.Event(),
                lambda *_args: None,
                lambda _result: None,
            )

            self.assertEqual([result.status for result in results], ["failed", "success"])
            self.assertIn("模拟失败", results[0].error)
            self.assertFalse((folder / "bad.docx").exists())
            self.assertTrue((folder / "good.docx").exists())

    def test_cancel_stops_after_current_item(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            sources = self._make_sources(folder, ["one.pdf", "two.pdf", "three.pdf"])
            cancel_event = threading.Event()
            converted = []

            def converter(source, output, _progress):
                converted.append(Path(source).name)
                Path(output).write_bytes(b"docx")
                cancel_event.set()

            results = run_conversion_batch(
                sources,
                ".docx",
                None,
                converter,
                cancel_event,
                lambda *_args: None,
                lambda _result: None,
            )

            self.assertEqual(converted, ["one.pdf"])
            self.assertEqual(
                [result.status for result in results],
                ["success", "cancelled", "cancelled"],
            )


if __name__ == "__main__":
    unittest.main()
