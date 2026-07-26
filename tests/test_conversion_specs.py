import unittest

from conversion_specs import (
    PDF_TARGET_POWERPOINT,
    PDF_TARGET_WORD,
    resolve_conversion_spec,
    source_kind_for_extension,
)


class ConversionSpecTests(unittest.TestCase):
    def test_pdf_target_selects_word_or_powerpoint(self):
        word = resolve_conversion_spec("pdf", PDF_TARGET_WORD)
        powerpoint = resolve_conversion_spec("pdf", PDF_TARGET_POWERPOINT)

        self.assertEqual((word.key, word.target_suffix), ("pdf_to_word", ".docx"))
        self.assertEqual(
            (powerpoint.key, powerpoint.target_suffix),
            ("pdf_to_powerpoint", ".pptx"),
        )

    def test_office_sources_always_target_pdf(self):
        self.assertEqual(resolve_conversion_spec("word").key, "word_to_pdf")
        self.assertEqual(
            resolve_conversion_spec("powerpoint").key,
            "powerpoint_to_pdf",
        )

    def test_legacy_and_modern_powerpoint_extensions_share_source_kind(self):
        self.assertEqual(source_kind_for_extension(".ppt"), "powerpoint")
        self.assertEqual(source_kind_for_extension(".PPTX"), "powerpoint")

    def test_rejects_unsupported_direction(self):
        with self.assertRaises(ValueError):
            resolve_conversion_spec("pdf", "pdf")
        with self.assertRaises(ValueError):
            source_kind_for_extension(".doc")


if __name__ == "__main__":
    unittest.main()
