import pytest
import os, sys

from verify_utils import verify, ARTIFACTS_DIR 

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src', 'aspose_words_docling')))
from license_manager import LicenseManager

@pytest.fixture(scope="module", autouse=True)
def prepare_test_env():
    os.makedirs(ARTIFACTS_DIR, exist_ok=True)
    LicenseManager().apply_license()


@pytest.mark.parametrize("test_file, compare_with_docling", [
    ("docx\\simple.docx", True),
    ("docx\\equations-simple.docx", False),
    ("docx\\title.docx", True),
    ("docx\\title_bold.docx", True),
    ("docx\\heading.docx", False),
    ("docx\\heading_wo_title.docx", False),
    ("docx\\heading_with_title.docx", True),
    ("docx\\sections.docx", False),
    ("docx\\image.docx", True),
    ("docx\\hyperlink.docx", True),
    ("docx\\text-format.docx", True),
    ("docx\\table-simple.docx", True),
    ("docx\\table-with-picture.docx", True),
    ("docx\\table-h-merged-cells.docx", True),
    ("docx\\table-v-merged-cells.docx", True),
    ("docx\\table-hv-merged-cells.docx", True),
    ("docx\\list-ordered-simple.docx", False),
    ("docx\\list-unordered-simple.docx", True),
    ("docx\\list-levels.docx", False),
    ("docx\\table-complex-cells.docx", True),
    ("docx\\table-with-one-cell.docx", True),
    ("docx\\fields.docx", False),
    ("docx\\textbox-simple.docx", True),
    ("docx\\comment-simple.docx", True),
    ("docx\\footnotes.docx", True),
    ("docx\\bookmarks.docx", True),
    # Docling test documents
    ("docx\\docling-tests\\lorem_ipsum.docx", True),
    ("docx\\docling-tests\\unit_test_lists.docx", False),
    ("docx\\docling-tests\\tablecell.docx", True),
    ("docx\\docling-tests\\word_tables.docx", False),
    ("docx\\docling-tests\\unit_test_headers.docx", False),
    ("docx\\docling-tests\\unit_test_formatting.docx", False),
    ("docx\\docling-tests\\word_image_anchors.docx", True),
    ("docx\\docling-tests\\equations.docx", True),
    ("docx\\docling-tests\\textbox.docx", False),
    ("docx\\docling-tests\\word_sample.docx", False)
])
def test_asposewordsconverter(test_file: str, compare_with_docling: bool):
    verify(test_file, compare_with_docling)


if __name__ == "__main__":
    # TODO
    print("Tests OK")