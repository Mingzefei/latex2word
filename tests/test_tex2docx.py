import unittest

from tex2docx import LatexToWordConverter


class TestLatexToWordConverter(unittest.TestCase):
    def setUp(self):
        # config for tests/en
        self.en_config = {
            "input_texfile": "./en/main.tex",
            "output_docxfile": "./en/main.docx",
            "reference_docfile": "../my_temp.docx",
            "cslfile": "../ieee.csl",
            "bibfile": "./ref.bib",
            "debug": True,
        }
        self.en_chapter_config = {
            "input_texfile": "./en_chapter/main.tex",
            "output_docxfile": "./en_chapter/main.docx",
            "reference_docfile": "../my_temp.docx",
            "cslfile": "../ieee.csl",
            "bibfile": "./ref.bib",
            "debug": True,
        }
        self.en_include_config = {
            "input_texfile": "./en_include/main.tex",
            "output_docxfile": "./en_include/main.docx",
            "cslfile": "../ieee.csl",
            "bibfile": "./ref.bib",
            "debug": True,
        }
        # config for tests/zh
        self.zh_config = {
            "input_texfile": "./zh/main.tex",
            "output_docxfile": "./zh/main.docx",
            "bibfile": "./ref.bib",
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
