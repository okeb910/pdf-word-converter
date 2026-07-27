import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pdf_word_converter as converter
from conversion_specs import POWERPOINT_TO_PDF, WORD_TO_PDF
from engine_models import EngineState, EngineStatus
from macos_office import MacOSOfficeError


class ValueHolder:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class CrossPlatformEngineUiTests(unittest.TestCase):
    def _make_status_app(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.is_macos = True
        app._avail_methods = {
            converter.WORD_NATIVE: None,
            converter.POWERPOINT_NATIVE: None,
            "word_com": None,
            "powerpoint_com": None,
            "libreoffice": True,
            "images": True,
            "pptx": True,
        }
        app._engine_statuses = {
            "libreoffice": EngineStatus(EngineState.AVAILABLE),
            "images": EngineStatus(EngineState.AVAILABLE),
            "pptx": EngineStatus(EngineState.AVAILABLE),
        }
        return app

    def test_unverified_native_engine_is_displayed_and_selectable(self):
        app = self._make_status_app()
        status = EngineStatus(
            EngineState.UNVERIFIED,
            "已检测到 Microsoft Word，尚未验证自动化权限",
        )

        app._set_engine_status(converter.WORD_NATIVE, status)

        self.assertEqual(
            app._availability_text(
                converter.WORD_NATIVE,
                app._method_availability(converter.WORD_NATIVE),
            ),
            "已安装/使用时验证",
        )
        self.assertTrue(app._method_availability(converter.WORD_NATIVE))
        self.assertTrue(app._method_availability("word_com"))

    def test_macos_engine_status_text_omits_winget(self):
        app = self._make_status_app()
        app._set_engine_status(
            converter.WORD_NATIVE,
            EngineStatus(EngineState.UNVERIFIED),
        )
        app._set_engine_status(
            converter.POWERPOINT_NATIVE,
            EngineStatus(EngineState.MISSING),
        )
        app.engine_status_var = ValueHolder("")
        app._pending_engine_checks = {"winget"}
        app._winget_path = "C:/Windows/System32/winget.exe"

        app._refresh_engine_status_text()

        status_text = app.engine_status_var.get()
        self.assertIn("Word: 已安装/使用时验证", status_text)
        self.assertNotIn("winget", status_text.casefold())

    def test_macos_best_pdf_to_word_method_defaults_to_images(self):
        app = self._make_status_app()
        app._set_engine_status(
            converter.WORD_NATIVE,
            EngineStatus(EngineState.UNVERIFIED),
        )
        app.method_var = ValueHolder(converter.WORD_NATIVE)

        app._select_best_pdf_method()

        self.assertEqual(app.method_var.get(), "images")


class MacOSConverterSelectionTests(unittest.TestCase):
    def _make_app(self, native_key, backend):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.is_macos = True
        app.root = Mock()
        app._avail_methods = {
            converter.WORD_NATIVE: native_key == converter.WORD_NATIVE,
            converter.POWERPOINT_NATIVE: native_key == converter.POWERPOINT_NATIVE,
            "libreoffice": False,
        }
        app._engine_statuses = {}
        app._mac_word_backend = backend if native_key == converter.WORD_NATIVE else None
        app._mac_powerpoint_backend = (
            backend if native_key == converter.POWERPOINT_NATIVE else None
        )
        app._record_native_status = Mock()
        return app

    def test_native_macos_converters_record_permission_denial(self):
        permission_status = EngineStatus(
            EngineState.PERMISSION_DENIED,
            "系统设置 → 隐私与安全性 → 自动化",
        )
        cases = (
            (converter.WORD_NATIVE, WORD_TO_PDF),
            (converter.POWERPOINT_NATIVE, POWERPOINT_TO_PDF),
        )

        for native_key, spec in cases:
            with self.subTest(engine=native_key):
                backend = Mock()
                backend.convert.side_effect = MacOSOfficeError(
                    "Microsoft Word"
                    if native_key == converter.WORD_NATIVE
                    else "Microsoft PowerPoint",
                    permission_status,
                )
                app = self._make_app(native_key, backend)
                convert = app._create_converter(spec, method=None)
                progress = Mock()

                with patch.object(
                    converter, "libreoffice_available", return_value=False
                ), self.assertRaises(MacOSOfficeError):
                    convert("C:/测试 文件/input.docx", "C:/输出/result.pdf", progress)

                backend.convert.assert_called_once_with(
                    "C:/测试 文件/input.docx",
                    "C:/输出/result.pdf",
                    progress,
                )
                app._record_native_status.assert_called_once_with(
                    native_key,
                    permission_status,
                )

    def test_native_permission_failure_falls_back_to_libreoffice_for_same_file(self):
        permission_status = EngineStatus(
            EngineState.PERMISSION_DENIED,
            "系统设置 → 隐私与安全性 → 自动化",
        )
        cases = (
            (
                converter.WORD_NATIVE,
                WORD_TO_PDF,
                "document.docx",
                "docx_to_pdf_via_libreoffice",
            ),
            (
                converter.POWERPOINT_NATIVE,
                POWERPOINT_TO_PDF,
                "slides.pptx",
                "presentation_to_pdf_via_libreoffice",
            ),
        )

        temporary_directories = []
        self.addCleanup(lambda: [directory.cleanup() for directory in temporary_directories])
        for native_key, spec, source, fallback_name in cases:
            with self.subTest(engine=native_key):
                temporary_directory = tempfile.TemporaryDirectory()
                temporary_directories.append(temporary_directory)
                output = Path(temporary_directory.name) / "result.pdf"
                backend = Mock()

                def fail_native(_source, target, _progress):
                    Path(target).write_bytes(b"partial-native")
                    raise MacOSOfficeError(
                        "Microsoft Word"
                        if native_key == converter.WORD_NATIVE
                        else "Microsoft PowerPoint",
                        permission_status,
                    )

                backend.convert.side_effect = fail_native
                app = self._make_app(native_key, backend)
                convert = app._create_converter(spec, method=None)
                progress = Mock()

                def succeed_fallback(_source, target, _progress):
                    self.assertFalse(Path(target).exists())
                    Path(target).write_bytes(b"libreoffice")

                with patch.object(
                    converter, "libreoffice_available", return_value=True
                ), patch.object(
                    converter, fallback_name, side_effect=succeed_fallback
                ) as fallback:
                    convert(source, output, progress)

                fallback.assert_called_once_with(source, output, progress)
                self.assertEqual(output.read_bytes(), b"libreoffice")
                app._record_native_status.assert_called_once_with(
                    native_key,
                    permission_status,
                )
                self.assertTrue(
                    any(
                        "LibreOffice" in call.args[0] and call.args[1] == 5
                        for call in progress.call_args_list
                    )
                )


class MacOSMissingEnginePromptTests(unittest.TestCase):
    def test_missing_native_and_libreoffice_offer_each_download_or_cancel(self):
        cases = (
            (True, True, "office"),
            (False, True, "libreoffice"),
            (None, False, None),
        )

        for choice, expected_result, expected_product in cases:
            with self.subTest(choice=choice):
                app = converter.ConverterApp.__new__(converter.ConverterApp)
                app.is_macos = True
                app._engine_statuses = {
                    converter.WORD_NATIVE: EngineStatus(EngineState.MISSING)
                }
                app._avail_methods = {
                    converter.WORD_NATIVE: False,
                    "libreoffice": False,
                }
                app._start_install = Mock()
                app.log = Mock()

                with patch.object(
                    converter.messagebox, "askyesnocancel", return_value=choice
                ) as prompt:
                    result = app._prompt_macos_conversion_engines(
                        "Microsoft Word",
                        converter.WORD_NATIVE,
                        registered=False,
                    )

                self.assertIs(result, expected_result)
                prompt.assert_called_once()
                if expected_product is None:
                    app._start_install.assert_not_called()
                    app.log.assert_called_once_with("用户取消打开转换引擎下载页")
                else:
                    app._start_install.assert_called_once_with(expected_product)


class SafeWindowCloseTests(unittest.TestCase):
    def test_conversion_close_requests_cancel_then_destroys_after_batch(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app.root = Mock()
        app._is_converting = True
        app._close_after_batch = False
        app.cancel_conversion = Mock()
        app.progress_text_var = ValueHolder("")
        app.log = Mock()

        with patch.object(converter.messagebox, "askyesno", return_value=True):
            app._on_window_close()

        self.assertTrue(app._close_after_batch)
        app.cancel_conversion.assert_called_once_with()
        app.root.destroy.assert_not_called()

        app._set_busy = Mock()
        app.output_paths = []
        app.queue_summary_var = ValueHolder("")
        app.progress = {"value": 0}
        app._on_batch_complete([])

        app.root.destroy.assert_called_once_with()


class TkDndFallbackTests(unittest.TestCase):
    def test_main_falls_back_to_standard_tk_when_tkdnd_root_fails(self):
        dnd_root_factory = Mock(side_effect=RuntimeError("TkDND load failed"))
        standard_root = Mock()
        dnd_module = SimpleNamespace(Tk=dnd_root_factory)

        with patch.object(converter, "TkinterDnD", dnd_module), patch.object(
            converter.tk, "Tk", return_value=standard_root
        ) as tk_factory, patch.object(converter, "ConverterApp") as app_class, patch.object(
            converter.logging, "exception"
        ):
            converter.main()

        dnd_root_factory.assert_called_once_with()
        tk_factory.assert_called_once_with()
        self.assertFalse(standard_root._pdf_converter_dnd_available)
        app_class.assert_called_once_with(standard_root)
        standard_root.mainloop.assert_called_once_with()


if __name__ == "__main__":
    unittest.main()
