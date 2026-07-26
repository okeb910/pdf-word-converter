import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import pdf_word_converter as converter


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


if __name__ == "__main__":
    unittest.main()
