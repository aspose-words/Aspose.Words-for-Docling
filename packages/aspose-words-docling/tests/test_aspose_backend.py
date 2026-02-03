# Copyright (c) 2001-2026 Aspose Pty Ltd.

import pytest
import os, sys

from verify_utils import verify, ARTIFACTS_DIR 

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src', 'aspose_words_docling')))
from aspose_words_converter import LicenseManager

@pytest.fixture(scope="module", autouse=True)
def prepare_test_env():
    os.makedirs(ARTIFACTS_DIR, exist_ok=True)
    LicenseManager().apply_license()


@pytest.mark.parametrize("test_file, compare_with_docling", [
    ("docx\\simple.docx", True),
    ("docx\\equations-simple.docx", False),
    ("docx\\title.docx", True),
    ("docx\\title_bold.docx", True),
    ("docx\\heading.docx", True),
    ("docx\\heading_wo_title.docx", True),
    ("docx\\heading_with_title.docx", True),
    ("docx\\sections.docx", True),
    ("docx\\image.docx", True),
    ("docx\\hyperlink.docx", False),
    ("docx\\text-format.docx", True),
    ("docx\\table-simple.docx", True),
    ("docx\\table-with-picture.docx", False),
    ("docx\\table-h-merged-cells.docx", True),
    ("docx\\table-v-merged-cells.docx", False),
    ("docx\\table-hv-merged-cells.docx", False),
    ("docx\\list-ordered-simple.docx", False),
    ("docx\\list-unordered-simple.docx", False),
    ("docx\\list-levels.docx", False),
    ("docx\\table-complex-cells.docx", False),
    ("docx\\table-with-one-cell.docx", False),
    ("docx\\fields.docx", False),
    ("docx\\textbox-simple.docx", False),
    ("docx\\comment-simple.docx", True),
    ("docx\\footnotes.docx", True),
    ("docx\\bookmarks.docx", True),
    # Docling test documents
    ("docx\\docling-tests\\lorem_ipsum.docx", True),
    ("docx\\docling-tests\\unit_test_lists.docx", False),
    ("docx\\docling-tests\\tablecell.docx", False),
    ("docx\\docling-tests\\word_tables.docx", False),
    ("docx\\docling-tests\\unit_test_headers.docx", True),
    ("docx\\docling-tests\\unit_test_formatting.docx", False),
    ("docx\\docling-tests\\word_image_anchors.docx", True),
    ("docx\\docling-tests\\equations.docx", False),
    ("docx\\docling-tests\\textbox.docx", False),
    ("docx\\docling-tests\\word_sample.docx", False),
    # Other formats
    ("other\\word_sample.doc", False),
    ("other\\html_sample.html", False),
    ("other\\mhtml_sample.mhtml", False),
    ("other\\word_sample.odt", False),
    ("other\\word_sample.rtf", False),
    ("other\\pdf_sample.pdf", False),
    ("other\\epub_sample.epub", False),
    ("other\\md_sample.md", False)
])
def test_asposewordsconverter(test_file: str, compare_with_docling: bool):
    verify(test_file, compare_with_docling)


if __name__ == "__main__":
    pass