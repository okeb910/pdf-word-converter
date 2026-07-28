import tempfile
import threading
import unittest
import unicodedata
from pathlib import Path
from unittest.mock import patch

from batch_logic import (
    _cleanup_failed_output,
    deduplicate_paths,
    resolve_output_path,
    run_conversion_batch,
)


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
            first.touch()
            second.touch()
            self.assertEqual(
                deduplicate_paths([first, second, first]),
                [first.absolute(), second.absolute()],
            )

    def test_deduplicates_unicode_nfc_and_nfd_fallback_paths(self):
        with tempfile.TemporaryDirectory() as tmp:
            composed = Path(tmp) / "résumé.pdf"
            decomposed = Path(tmp) / unicodedata.normalize("NFD", composed.name)

            self.assertNotEqual(str(composed), str(decomposed))
            self.assertEqual(
                deduplicate_paths([decomposed, composed]),
                [decomposed.absolute()],
            )

    def test_deduplicates_symbolic_link_to_existing_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = folder / "source.pdf"
            alias = folder / "alias.pdf"
            source.touch()
            try:
                alias.symlink_to(source)
            except (NotImplementedError, OSError) as exc:
                self.skipTest(f"当前平台不允许创建符号链接: {exc}")

            self.assertEqual(
                deduplicate_paths([alias, source, alias]),
                [alias.absolute()],
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
            self.assertIsNone(results[0].output)
            self.assertFalse((folder / "bad.docx").exists())
            self.assertTrue((folder / "good.docx").exists())

    def test_cleanup_failed_output_deletes_rejected_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "rejected.docx"
            output.write_bytes(b"partial")

            message = _cleanup_failed_output(output, retry_delay=0)

            self.assertEqual(message, "")
            self.assertFalse(output.exists())

    def test_cleanup_failed_output_quarantines_locked_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "rejected.docx"
            output.write_bytes(b"partial")

            with patch.object(
                Path,
                "unlink",
                side_effect=PermissionError("locked"),
            ):
                message = _cleanup_failed_output(
                    output,
                    attempts=1,
                    retry_delay=0,
                )

            quarantine = Path(str(output) + ".failed")
            self.assertIn("已隔离", message)
            self.assertFalse(output.exists())
            self.assertTrue(quarantine.exists())

    def test_cleanup_failed_output_reports_unremovable_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "rejected.docx"
            output.write_bytes(b"partial")

            with patch.object(
                Path,
                "unlink",
                side_effect=PermissionError("locked"),
            ), patch.object(
                Path,
                "rename",
                side_effect=PermissionError("still locked"),
            ):
                message = _cleanup_failed_output(
                    output,
                    attempts=1,
                    retry_delay=0,
                )

            self.assertIn("清理失败", message)
            self.assertIn(str(output), message)
            self.assertTrue(output.exists())

    def test_batch_failure_appends_cleanup_message(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = self._make_sources(folder, ["bad.pdf"])[0]

            def converter(_source, output, _progress):
                Path(output).write_bytes(b"partial")
                raise RuntimeError("conversion rejected")

            with patch(
                "batch_logic._cleanup_failed_output",
                return_value="不完整输出仍被占用",
            ):
                results = run_conversion_batch(
                    [source],
                    ".docx",
                    None,
                    converter,
                    threading.Event(),
                    lambda *_args: None,
                    lambda _result: None,
                )

            self.assertEqual(results[0].status, "failed")
            self.assertIsNone(results[0].output)
            self.assertIn("conversion rejected", results[0].error)
            self.assertIn("不完整输出仍被占用", results[0].error)

    def test_cleanup_failure_never_exposes_formal_output_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            source = self._make_sources(folder, ["rejected.pdf"])[0]
            cleanup_targets = []

            def converter(_source, output, _progress):
                Path(output).write_bytes(b"invalid docx")
                raise RuntimeError("editable-table gate rejected output")

            def fail_cleanup(output, *args, **kwargs):
                del args, kwargs
                cleanup_targets.append(Path(output))
                return f"不完整输出清理失败，文件仍位于: {output}"

            with patch(
                "batch_logic._cleanup_failed_output",
                side_effect=fail_cleanup,
            ):
                results = run_conversion_batch(
                    [source],
                    ".docx",
                    None,
                    converter,
                    threading.Event(),
                    lambda *_args: None,
                    lambda _result: None,
                )

            formal_output = folder / "rejected.docx"
            self.assertEqual(results[0].status, "failed")
            self.assertIsNone(results[0].output)
            self.assertFalse(formal_output.exists())
            self.assertEqual(len(cleanup_targets), 1)
            isolated_output = cleanup_targets[0]
            self.assertEqual(isolated_output.parent, folder)
            self.assertTrue(isolated_output.name.startswith(".rejected."))
            self.assertTrue(isolated_output.name.endswith(".partial.docx"))
            self.assertTrue(isolated_output.exists())

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
