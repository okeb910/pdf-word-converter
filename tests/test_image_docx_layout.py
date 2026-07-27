import tempfile
import unittest
from pathlib import Path

import pymupdf
from docx import Document

from pdf_word_converter import pdf_to_word_via_images


class ImageDocxLayoutTests(unittest.TestCase):
    def test_mixed_page_images_are_placed_after_section_breaks(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "中文 mixed pages.pdf"
            output = folder / "中文 mixed pages.docx"

            pdf = pymupdf.open()
            pdf.new_page(width=595, height=842)
            pdf.new_page(width=842, height=595)
            pdf.save(source)
            pdf.close()

            pdf_to_word_via_images(source, output, lambda _message, _pct: None, dpi=72)

            document = Document(output)
            self.assertEqual(len(document.sections), 2)
            self.assertEqual(len(document.inline_shapes), 2)
            self.assertEqual(len(document.paragraphs), 3)

            section_break = document.paragraphs[1]._p
            second_image = document.paragraphs[2]._p
            self.assertTrue(section_break.xpath("./w:pPr/w:sectPr"))
            self.assertFalse(section_break.xpath(".//w:drawing"))
            self.assertTrue(second_image.xpath(".//w:drawing"))
            self.assertFalse(second_image.xpath("./w:pPr/w:sectPr"))

            self.assertLess(
                document.sections[0].page_width,
                document.sections[0].page_height,
            )
            self.assertGreater(
                document.sections[1].page_width,
                document.sections[1].page_height,
            )


if __name__ == "__main__":
    unittest.main()
