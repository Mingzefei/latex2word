import unittest
import os  # Import the os module

from tex2docx import LatexToWordConverter

# Get the directory where the test file is located
TEST_DIR = os.path.dirname(os.path.abspath(__file__))
# Get the project root directory (assuming tests is one level down)
PROJECT_ROOT = os.path.dirname(TEST_DIR)


class TestLatexToWordConverter(unittest.TestCase):
    def setUp(self):
        # Construct absolute paths based on the test file's location
        # config for tests/en
        self.en_config = {
            "input_texfile": os.path.join(TEST_DIR, "en/main.tex"),
            "output_docxfile": os.path.join(TEST_DIR, "en/main.docx"),
            # Assuming reference/csl files are in the project root
            "reference_docfile": os.path.join(PROJECT_ROOT, "my_temp.docx"),
            "cslfile": os.path.join(PROJECT_ROOT, "ieee.csl"),
            # Assuming bib file is in the tests directory
            "bibfile": os.path.join(TEST_DIR, "ref.bib"),
            "debug": True,
        }
        self.en_chapter_config = {
            "input_texfile": os.path.join(TEST_DIR, "en_chapter/main.tex"),
            "output_docxfile": os.path.join(TEST_DIR, "en_chapter/main.docx"),
            "reference_docfile": os.path.join(PROJECT_ROOT, "my_temp.docx"),
            "cslfile": os.path.join(PROJECT_ROOT, "ieee.csl"),
            "bibfile": os.path.join(TEST_DIR, "ref.bib"),
            "debug": True,
        }
        self.en_include_config = {
            "input_texfile": os.path.join(TEST_DIR, "en_include/main.tex"),
            "output_docxfile": os.path.join(TEST_DIR, "en_include/main.docx"),
            "cslfile": os.path.join(PROJECT_ROOT, "ieee.csl"),
            "bibfile": os.path.join(TEST_DIR, "ref.bib"),
            "debug": True,
        }
        # config for tests/zh
        self.zh_config = {
            "input_texfile": os.path.join(TEST_DIR, "zh/main.tex"),
            "output_docxfile": os.path.join(TEST_DIR, "zh/main.docx"),
            "bibfile": os.path.join(TEST_DIR, "ref.bib"),
            "fix_table": True,
            "debug": False,
        }

    def test_convert_en(self):
        # test convert en
        converter = LatexToWordConverter(**self.en_config)
        converter.convert()

    def test_convert_en_chapter(self):
        # test convert en_chapter
        converter = LatexToWordConverter(**self.en_chapter_config)
        converter.convert()

    def test_convert_en_include(self):
        # test convert en_include
        converter = LatexToWordConverter(**self.en_include_config)
        converter.convert()

    def test_convert_zh(self):
        # test convert zh
        converter = LatexToWordConverter(**self.zh_config)
        converter.convert()


if __name__ == "__main__":
    unittest.main()
