# LaTeX to Word Conversion Tool

[中文版本](README_zh.md)

This project provides a Python script that utilizes Pandoc and Pandoc-Crossref tools to automatically convert LaTeX files into Word documents in a specified format.
It should be noted that there is currently no perfect way to convert LaTeX to Word. The Word documents produced by this project can meet general review and editing needs, although about 5% of the content (such as author information) may need to be manually corrected after conversion.

## Features

- Supports the conversion of formulas;
- Supports automatic numbering and cross-referencing of images, tables, formulas, and references;
- Supports multi-figure images;
- Generally supports outputting Word in a specified format.

## Quick Start

Ensure that Pandoc, Pandoc-Crossref, and other dependencies are correctly installed, as detailed in [Installing Dependencies](#installing-dependencies). Execute the following command in the terminal:

```shell
python ./src/tex2docx.py --input_texfile <your_texfile> --multifig_dir <dir_saving_temporary_figs> --output_docxfile <your_docxfile> --reference_docfile <your_reference_docfile> --bibfile <your_bibfile> --cslfile <your_cslfile>
```

Replace `<...>` in the command with the appropriate file paths or folder names.

## Installing Dependencies

You need to install Pandoc, Pandoc-Crossref, and related Python libraries.

### Pandoc

Install Pandoc by referring to the [Pandoc Official Documentation](https://github.com/jgm/pandoc/blob/main/INSTALL.md). It is recommended to download the latest installation package from [Pandoc Releases](https://github.com/jgm/pandoc/releases).

### Pandoc-Crossref

Install Pandoc-Crossref as detailed in the [Pandoc-Crossref Official Documentation](https://github.com/lierdakil/pandoc-crossref). Ensure you download the version of Pandoc-Crossref that matches your Pandoc installation and configure the path appropriately.

### Related Python Libraries

Install Python dependencies:

```shell
pip install -e .
```

## Usage and Examples

The tool supports both command line and script usage, ensure all required dependencies are installed.

### Command Line Usage

Execute the following command in the terminal:

```shell
python ./src/tex2docx.py --input_texfile <your_texfile> --multifig_dir <dir_saving_temporary_figs> --output_docxfile <your_docxfile> --reference_docfile <your_reference_docfile> --bibfile <your_bibfile> --cslfile <your_cslfile>
```

Parameter explanations:
- `--input_texfile`: Specifies the path of the LaTeX file to convert.
- `--multifig_dir`: Specifies the directory for temporarily storing generated multi-figures.
- `--output_docxfile`: Specifies the path of the output Word document.
- `--reference_docfile`: Specifies a Word reference document to ensure consistency in document styling.
- `--bibfile`: Specifies the BibTeX file for document citations.
- `--cslfile`: Specifies the Citation Style Language file to control the formatting of references.
- `--debug`: Enables debug mode to output more run-time information, helpful for troubleshooting.

For example, using the `tests/en` test case, execute the following command in the repository directory:

```shell
python ./src/tex2docx.py --input_texfile ./tests/en/main.tex --multifig_dir ./tests/en/multifigs --output_docxfile ./tests/en/main_cli.docx --reference_docfile ./my_temp.docx --bibfile ./tests/ref.bib --cslfile ./ieee.csl
```
You will find the converted `main_cli.docx` file in the `tests/en` directory.

### Script Usage

Create the script `my_convert.py`, write the following code, and execute:

```python
# my_convert.py
from tex2docx import LatexToWordConverter

config = {
    'input_texfile': '<your_texfile>',
    'output_docxfile': '<your_docxfile>',
    'multifig_dir': '<dir_saving_temporary_figs>',
    'reference_docfile': '<your_reference_docfile>',
    'cslfile': '<your_cslfile>',
    'bibfile': '<your_bibfile>',
    'debug': False
}

converter = LatexToWordConverter(**config)
converter.convert()
```

Examples can be found in `tests/test_tex2docx.py`.

## Implementation Principles and References

The core of this project is the use of Pandoc and Pandoc-Crossref tools to convert LaTeX to Word, configured as follows:

```shell
pandoc texfile -o docxfile \
    --lua-filter resolve_equation_labels.lua \
    --filter pandoc-crossref \
    --reference-doc=temp.docx \
    --number-sections \
    -M autoEqnLabels \
    -M tableEqns \
    -M reference-section-title=Reference \
    --bibliography=ref.bib \
    --citeproc --csl ieee.csl
```

However, this method may encounter issues such as improper image importation and incorrect referencing when dealing with LaTeX files containing multi-figure images directly. To address this, the project extracts multi-figure image code from the LaTeX files and uses LaTeX's built-in `convert` and `pdftocairo` tools to automatically compile these images into a single large PNG format. These PNG files then replace the original image codes in the LaTeX document, ensuring smooth import of multi-figure images. For implementation details, see `tex2docx.py`.

## Outstanding Issues

1. Refer to subfigures uniformly using `\ref{<figure_lab>}(a)`, not `\ref{<subfigure_lab>}` (direct subfigure referencing will be supported in future updates);
2. The formatting of image captions and author information in the exported Word document needs manual adjustment.

## Other

There are two kinds of people in the world, those who use LaTeX and those who do not. The latter often request Word versions of documents from the former. Thus, the following command was born:

```bash
pandoc input.tex -o output.docx\
  --filter pandoc-crossref \
  --reference-doc=my_temp.docx \
  --number-sections \
  -M autoEqnLabels -M tableEqns \
  -M reference-section-title=Reference \
  --bibliography=my_ref.bib \
  --citeproc --csl ieee.csl
```