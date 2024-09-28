import unittest

from tex2docx  import LatexToWordConverter


class TestLatexToWordConverter(unittest.TestCase):
    def setUp(self):
        # config for tests/en
        self.en_config = {
            "input_texfile": "./en/main.tex",
            "output_docxfile": "./en/main.docx",
            "multifig_dir": "./en/multifigs",
            "reference_docfile": "../my_temp.docx",
            "cslfile": "../ieee.csl",
            "bibfile": "./ref.bib",
            "debug": True,
        }
        # config for tests/zh
        self.zh_config = {
            "input_texfile": "./zh/main.tex",
            "output_docxfile": "./zh/main.docx",
            "multifig_dir": "./zh/multifigs",
            "reference_docfile": "../my_temp.docx",
            "cslfile": "../ieee.csl",
            "bibfile": "./ref.bib",
            "debug": False,
        }

    def test_convert_en(self):
        # test convert en
        converter = LatexToWordConverter(**self.en_config)
        converter.convert()

    def test_convert_zh(self):
        # test convert zh
        converter = LatexToWordConverter(**self.zh_config)
        converter.convert()


if __name__ == "__main__":
    unittest.main()
