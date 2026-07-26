import queue
import tempfile
import threading
import tkinter as tk
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pdf_word_converter as converter
from drop_logic import MixedSourceKindsError, classify_dropped_paths


class ValueHolder:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class BlockingProbe:
    def __init__(self, result=True):
        self.result = result
        self.started = threading.Event()
        self.release = threading.Event()

    def __call__(self):
        self.started.set()
        if not self.release.wait(2):
            raise TimeoutError("probe was not released")
        return self.result


class DropLogicTests(unittest.TestCase):
    def test_classifies_case_insensitively_and_ignores_unsupported_paths(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            first = folder / "中文 文件.PDF"
            second = folder / "second.pdf"
            ignored = folder / "notes.txt"
            for path in (first, second, ignored):
                path.touch()

            source_kind, accepted, ignored_paths = classify_dropped_paths(
                (first, ignored, second, folder)
            )

            self.assertEqual(source_kind, "pdf")
            self.assertEqual(accepted, [first, second])
            self.assertEqual(ignored_paths, [ignored, folder])

    def test_ppt_and_pptx_share_one_source_kind(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            paths = [folder / "legacy.ppt", folder / "modern.pptx"]
            for path in paths:
                path.touch()

            source_kind, accepted, ignored = classify_dropped_paths(paths)

            self.assertEqual(source_kind, "powerpoint")
            self.assertEqual(accepted, paths)
            self.assertEqual(ignored, [])

    def test_mixed_supported_formats_are_rejected_as_one_drop(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            paths = [folder / "one.pdf", folder / "two.docx"]
            for path in paths:
                path.touch()

            with self.assertRaises(MixedSourceKindsError) as context:
                classify_dropped_paths(paths)

            self.assertEqual(context.exception.counts["pdf"], 1)
            self.assertEqual(context.exception.counts["word"], 1)

    def test_tcl_splitlist_preserves_chinese_and_space_paths(self):
        interpreter = tk.Tcl()
        raw = "{C:/测试 文件/一.pdf} {C:/other/two.pdf}"
        self.assertEqual(
            interpreter.splitlist(raw),
            ("C:/测试 文件/一.pdf", "C:/other/two.pdf"),
        )


class DropIntegrationTests(unittest.TestCase):
    def _make_app(self, dropped):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.root = Mock()
        app.root.tk.splitlist.return_value = tuple(str(path) for path in dropped)
        app._is_converting = False
        app._is_installing = False
        app._on_drop_hover = Mock()
        app._add_paths = Mock(return_value=len(dropped))
        app._current_target_kind = Mock(return_value="word")
        app.log = Mock()
        return app

    def test_pdf_target_zone_sets_explicit_target(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.pdf"
            source.touch()
            app = self._make_app([source])
            event = SimpleNamespace(data="ignored", action="copy")

            result = app._on_files_dropped(event, "powerpoint")

            app._add_paths.assert_called_once_with("pdf", [source], "powerpoint")
            self.assertEqual(result, converter.COPY)

    def test_move_drop_is_always_reported_as_copy(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.pdf"
            source.touch()
            app = self._make_app([source])
            event = SimpleNamespace(data="ignored", action="move")

            result = app._on_files_dropped(event, "word")

            self.assertEqual(result, converter.COPY)
            self.assertNotEqual(result, event.action)

    def test_generic_drop_makes_powerpoint_target_pdf(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "slides.pptx"
            source.touch()
            app = self._make_app([source])
            event = SimpleNamespace(data="ignored", action="copy")

            app._on_files_dropped(event)

            app._add_paths.assert_called_once_with("powerpoint", [source], "pdf")

    def test_non_pdf_is_rejected_from_pdf_specific_zone(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "document.docx"
            source.touch()
            app = self._make_app([source])
            event = SimpleNamespace(data="ignored", action="copy")

            with patch.object(converter.messagebox, "showinfo") as showinfo:
                result = app._on_files_dropped(event, "word")

            app._add_paths.assert_not_called()
            showinfo.assert_called_once()
            self.assertEqual(result, converter.REFUSE_DROP)

    def test_busy_drop_is_refused_before_paths_are_parsed(self):
        for busy_attribute in ("_is_converting", "_is_installing"):
            with self.subTest(busy_attribute=busy_attribute):
                app = self._make_app([])
                setattr(app, busy_attribute, True)
                event = SimpleNamespace(data="ignored", action="copy")

                result = app._on_files_dropped(event)

                self.assertEqual(result, converter.REFUSE_DROP)
                app.root.tk.splitlist.assert_not_called()
                app._add_paths.assert_not_called()

    def test_mixed_source_drop_is_refused_as_a_whole(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            pdf = folder / "one.pdf"
            docx = folder / "two.docx"
            pdf.touch()
            docx.touch()
            app = self._make_app([pdf, docx])
            event = SimpleNamespace(data="ignored", action="copy")

            with patch.object(converter.messagebox, "showwarning") as showwarning:
                result = app._on_files_dropped(event)

            self.assertEqual(result, converter.REFUSE_DROP)
            app._add_paths.assert_not_called()
            showwarning.assert_called_once()


class PdfTargetSelectionTests(unittest.TestCase):
    def _make_app(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.source_kind = "pdf"
        app.input_paths = [Path("one.pdf"), Path("two.pdf")]
        app.target_display_var = ValueHolder(app.TARGET_DISPLAY["word"])
        app._is_converting = False
        app._is_installing = False
        app._set_target_kind = Mock()
        app._update_method_panel = Mock()
        app._refresh_queue = Mock()
        return app

    def test_pdf_target_change_is_applied_after_confirmation(self):
        app = self._make_app()

        with patch.object(converter.messagebox, "askyesno", return_value=True) as ask:
            app._select_pdf_target("powerpoint")

        ask.assert_called_once()
        app._set_target_kind.assert_called_once_with("powerpoint")
        app._update_method_panel.assert_called_once()
        app._refresh_queue.assert_called_once()

    def test_pdf_target_change_is_unchanged_when_confirmation_is_rejected(self):
        app = self._make_app()

        with patch.object(converter.messagebox, "askyesno", return_value=False) as ask:
            app._select_pdf_target("powerpoint")

        ask.assert_called_once()
        app._set_target_kind.assert_not_called()
        app._update_method_panel.assert_not_called()
        app._refresh_queue.assert_not_called()


class ParallelDetectionTests(unittest.TestCase):
    def _make_detection_app(self, pending, generation=1):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._engine_generation = generation
        app._pending_engine_checks = set(pending)
        app._avail_methods = {
            "word_com": None,
            "powerpoint_com": None,
            "libreoffice": None,
            "images": True,
            "pptx": True,
        }
        app._engine_elapsed_ms = {}
        app._winget_path = None
        app._engine_detection_complete = False
        app._engine_started_at = converter.time.perf_counter()
        app._install_detection_generation = None
        app._is_installing = False
        app.method_var = ValueHolder("images")
        app.log = Mock()
        app._select_best_pdf_method = Mock()
        app._update_engine_indicators = Mock()
        app._refresh_engine_status_text = Mock()
        app._update_method_panel = Mock()
        app._update_action_states = Mock()
        return app

    def test_engine_cache_wrappers_do_not_serialize_different_engines(self):
        word = BlockingProbe()
        powerpoint = BlockingProbe()
        libreoffice = BlockingProbe()
        converter.reset_engine_cache()

        with patch.object(converter, "_check_word_com_available", word), patch.object(
            converter, "_check_powerpoint_com_available", powerpoint
        ), patch.object(converter, "_check_libreoffice_available", libreoffice):
            threads = [
                threading.Thread(target=converter.word_com_available),
                threading.Thread(target=converter.powerpoint_com_available),
                threading.Thread(target=converter.libreoffice_available),
            ]
            for thread in threads:
                thread.start()

            self.assertTrue(word.started.wait(1))
            self.assertTrue(powerpoint.started.wait(1))
            self.assertTrue(libreoffice.started.wait(1))
            word.release.set()
            powerpoint.release.set()
            libreoffice.release.set()
            for thread in threads:
                thread.join(1)
                self.assertFalse(thread.is_alive())

        converter.reset_engine_cache()

    def test_reset_engine_cache_does_not_wait_for_an_old_probe(self):
        word = BlockingProbe()
        converter.reset_engine_cache()

        with patch.object(converter, "_check_word_com_available", word):
            probe_thread = threading.Thread(target=converter.word_com_available)
            probe_thread.start()
            self.assertTrue(word.started.wait(1))

            reset_thread = threading.Thread(target=converter.reset_engine_cache)
            reset_thread.start()
            reset_thread.join(0.25)
            reset_was_blocked = reset_thread.is_alive()

            word.release.set()
            probe_thread.join(1)
            reset_thread.join(1)

        converter.reset_engine_cache()
        self.assertFalse(reset_was_blocked)
        self.assertFalse(probe_thread.is_alive())
        self.assertFalse(reset_thread.is_alive())

    def test_normal_detection_completion_does_not_clear_installing_state(self):
        app = self._make_detection_app({"word_com"}, generation=3)
        app._is_installing = True

        app._apply_engine_probe_result(3, "word_com", False, 20)

        self.assertTrue(app._engine_detection_complete)
        self.assertTrue(app._is_installing)
        self.assertIsNone(app._install_detection_generation)

    def test_installing_state_clears_only_when_its_detection_generation_finishes(self):
        app = self._make_detection_app(
            {"word_com", "powerpoint_com"}, generation=8,
        )
        app._is_installing = True
        app._install_detection_generation = 8

        app._apply_engine_probe_result(8, "word_com", False, 20)

        self.assertTrue(app._is_installing)
        self.assertEqual(app._install_detection_generation, 8)
        self.assertEqual(app._pending_engine_checks, {"powerpoint_com"})

        app._apply_engine_probe_result(8, "powerpoint_com", True, 30)

        self.assertFalse(app._is_installing)
        self.assertIsNone(app._install_detection_generation)
        self.assertTrue(app._engine_detection_complete)

    def test_timed_out_probe_is_removed_and_late_result_is_ignored(self):
        app = self._make_detection_app({"word_com"}, generation=11)

        app._expire_engine_probe(11, "word_com")

        self.assertEqual(app._pending_engine_checks, set())
        self.assertFalse(app._avail_methods["word_com"])
        self.assertEqual(app._engine_elapsed_ms["word_com"], 20000)
        timeout_log_count = app.log.call_count

        app._apply_engine_probe_result(11, "word_com", True, 25000)

        self.assertFalse(app._avail_methods["word_com"])
        self.assertEqual(app._engine_elapsed_ms["word_com"], 20000)
        self.assertEqual(app.log.call_count, timeout_log_count)

    def test_registered_word_com_failure_does_not_offer_reinstallation(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._avail_methods = {"word_com": False}
        app._is_installing = False
        app._word_registered = True
        app.log = Mock()
        app._start_install = Mock()

        with patch.object(converter.messagebox, "showwarning") as showwarning, patch.object(
            converter.messagebox, "askyesno"
        ) as askyesno:
            accepted = app._prompt_word_install()

        self.assertFalse(accepted)
        showwarning.assert_called_once()
        askyesno.assert_not_called()
        app._start_install.assert_not_called()

    def test_old_generation_result_is_ignored(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._engine_generation = 2
        app._pending_engine_checks = {"word_com"}
        app._avail_methods = {"word_com": None}
        app._engine_elapsed_ms = {}

        app._apply_engine_probe_result(1, "word_com", True, 10)

        self.assertIsNone(app._avail_methods["word_com"])
        self.assertEqual(app._pending_engine_checks, {"word_com"})

    def test_pdf_to_powerpoint_does_not_wait_for_external_engines(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.source_kind = "pdf"
        app.target_display_var = ValueHolder(app.TARGET_DISPLAY["powerpoint"])
        app.method_var = ValueHolder("word_com")
        app._avail_methods = {
            "word_com": None,
            "powerpoint_com": None,
            "libreoffice": None,
            "images": True,
            "pptx": True,
        }

        self.assertTrue(app._conversion_detection_ready())


if __name__ == "__main__":
    unittest.main()
