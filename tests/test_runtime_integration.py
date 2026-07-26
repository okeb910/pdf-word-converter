import threading
import unittest
from unittest.mock import Mock, patch

import pdf_word_converter as converter


class RuntimeIntegrationTests(unittest.TestCase):
    def test_reset_engine_cache_clears_all_cached_values(self):
        converter._LO_PATH = "lo"
        converter._WORD_AVAILABLE = True
        converter._LIBREOFFICE_AVAILABLE = True
        converter._PYMUPDF_AVAILABLE = True

        converter.reset_engine_cache()

        self.assertIsNone(converter._LO_PATH)
        self.assertIsNone(converter._WORD_AVAILABLE)
        self.assertIsNone(converter._LIBREOFFICE_AVAILABLE)
        self.assertIsNone(converter._PYMUPDF_AVAILABLE)

    def test_word_install_refusal_never_starts_installer(self):
        app = converter.ConverterApp.__new__(converter.ConverterApp)
        app._avail_methods = {"word_com": False}
        app._is_installing = False
        app._start_install = Mock()
        app.log = Mock()

        with patch.object(converter.messagebox, "askyesno", return_value=False):
            accepted = app._prompt_word_install()

        self.assertFalse(accepted)
        app._start_install.assert_not_called()

    def test_libreoffice_filter_failure_is_recognized(self):
        error = RuntimeError("Error: no export filter for source format: pdf")
        self.assertTrue(converter.is_libreoffice_export_filter_error(error))

    def test_unrelated_libreoffice_error_is_not_retried(self):
        self.assertFalse(converter.is_libreoffice_export_filter_error(RuntimeError("timeout")))


if __name__ == "__main__":
    unittest.main()
