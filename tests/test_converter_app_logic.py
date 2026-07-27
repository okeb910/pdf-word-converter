import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import pdf_word_converter as converter
from engine_models import EngineState, EngineStatus


class ValueHolder:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class ConverterAppLogicTests(unittest.TestCase):
    def test_cancelled_source_switch_does_not_change_pdf_target(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.source_kind = "pdf"
        app.input_paths = [Path("current.pdf")]
        app.target_display_var = ValueHolder(app.TARGET_DISPLAY["word"])
        app._clear_queue = Mock()
        app._refresh_queue = Mock()
        app.log = Mock()

        with patch.object(
            converter.filedialog, "askopenfilenames", return_value=("slides.pptx",)
        ), patch.object(converter.messagebox, "askyesno", return_value=False):
            app.select_files("powerpoint")

        self.assertEqual(app._current_target_kind(), "word")
        self.assertEqual(app.source_kind, "pdf")
        app._clear_queue.assert_not_called()

    def test_powerpoint_selection_accepts_ppt_and_pptx(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.source_kind = None
        app.input_paths = []
        app.target_display_var = ValueHolder(app.TARGET_DISPLAY["word"])
        app._refresh_queue = Mock()
        app.log = Mock()

        with patch.object(
            converter.filedialog,
            "askopenfilenames",
            return_value=("legacy.ppt", "modern.pptx", "ignored.pdf"),
        ):
            app.select_files("powerpoint")

        self.assertEqual(app.source_kind, "powerpoint")
        self.assertEqual(app._current_target_kind(), "pdf")
        self.assertEqual([path.suffix for path in app.input_paths], [".ppt", ".pptx"])

    def test_engine_status_update_never_prompts_for_office_on_startup(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._method_labels = {
            "word_com": Mock(),
            "images": Mock(),
            "libreoffice": Mock(),
        }
        app.pptx_status_label = Mock()
        app.method_var = ValueHolder("images")
        app.engine_status_var = ValueHolder("")
        app._update_method_panel = Mock()
        app._update_action_states = Mock()
        app.log = Mock()
        app._prompt_word_install = Mock()

        app._update_engine_status(False, False, False, True, True, None)

        app._prompt_word_install.assert_not_called()
        self.assertIn("PowerPoint: 不可用", app.engine_status_var.get())

    def test_engine_status_update_does_not_treat_failure_objects_as_available(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._method_labels = {
            converter.WORD_NATIVE: Mock(),
            "images": Mock(),
            "libreoffice": Mock(),
        }
        app.pptx_status_label = Mock()
        app.method_var = ValueHolder("images")
        app.engine_status_var = ValueHolder("")
        app._update_method_panel = Mock()
        app._update_action_states = Mock()
        app.log = Mock()

        app._update_engine_status(
            EngineStatus(EngineState.MISSING),
            EngineStatus(EngineState.LAUNCH_FAILED),
            EngineStatus(EngineState.TIMEOUT),
            EngineStatus(EngineState.AVAILABLE),
            EngineStatus(EngineState.AVAILABLE),
            None,
        )

        self.assertFalse(app._method_availability(converter.WORD_NATIVE))
        self.assertFalse(app._method_availability(converter.POWERPOINT_NATIVE))
        self.assertFalse(app._method_availability("libreoffice"))
        self.assertIn("PowerPoint: 已安装/启动失败", app.engine_status_var.get())

    def test_windows_converter_rejects_failed_structured_engine_statuses(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.is_macos = False
        convert = app._create_converter(
            converter.resolve_conversion_spec("word", "pdf"),
            method=None,
        )

        with patch.object(
            converter,
            "word_com_available",
            return_value=EngineStatus(EngineState.LAUNCH_FAILED),
        ), patch.object(
            converter,
            "libreoffice_available",
            return_value=EngineStatus(EngineState.TIMEOUT),
        ), patch.object(converter, "docx_to_pdf_via_word") as word_convert, patch.object(
            converter, "docx_to_pdf_via_libreoffice"
        ) as libreoffice_convert, self.assertRaisesRegex(
            RuntimeError, "需要 Microsoft Word 或 LibreOffice"
        ):
            convert("source.docx", "output.pdf", Mock())

        word_convert.assert_not_called()
        libreoffice_convert.assert_not_called()

    def test_windows_native_failure_removes_partial_output_and_falls_back(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.is_macos = False
        convert = app._create_converter(
            converter.resolve_conversion_spec("word", "pdf"),
            method=None,
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "result.pdf"

            def fail_native(_source, target, _progress):
                Path(target).write_bytes(b"partial-native")
                raise RuntimeError("Word COM failed")

            def succeed_fallback(_source, target, _progress):
                self.assertFalse(Path(target).exists())
                Path(target).write_bytes(b"libreoffice")

            with patch.object(
                converter,
                "word_com_available",
                return_value=EngineStatus(EngineState.UNVERIFIED),
            ), patch.object(
                converter,
                "libreoffice_available",
                return_value=EngineStatus(EngineState.UNVERIFIED),
            ), patch.object(
                converter, "docx_to_pdf_via_word", side_effect=fail_native
            ), patch.object(
                converter,
                "docx_to_pdf_via_libreoffice",
                side_effect=succeed_fallback,
            ):
                progress = Mock()
                convert("source.docx", output, progress)

            self.assertEqual(output.read_bytes(), b"libreoffice")
            self.assertTrue(
                any(
                    "改用 LibreOffice" in call.args[0] and call.args[1] == 5
                    for call in progress.call_args_list
                )
            )


if __name__ == "__main__":
    unittest.main()
