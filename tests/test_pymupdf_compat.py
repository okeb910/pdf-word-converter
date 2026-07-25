import builtins
import types
import unittest
from unittest.mock import patch

import pdf_word_converter as converter


class PyMuPdfCompatibilityTests(unittest.TestCase):
    def test_prefers_new_pymupdf_module_name(self):
        real_import = builtins.__import__
        fake_module = types.SimpleNamespace(open=object())

        def fake_import(name, *args, **kwargs):
            if name == "pymupdf":
                return fake_module
            return real_import(name, *args, **kwargs)

        with patch("builtins.__import__", side_effect=fake_import):
            self.assertIs(converter._load_pymupdf(), fake_module)

    def test_rejects_namespace_without_open_api(self):
        real_import = builtins.__import__

        def fake_import(name, *args, **kwargs):
            if name == "pymupdf":
                return types.SimpleNamespace()
            return real_import(name, *args, **kwargs)

        with patch("builtins.__import__", side_effect=fake_import):
            with self.assertRaises(ImportError):
                converter._load_pymupdf()


if __name__ == "__main__":
    unittest.main()
