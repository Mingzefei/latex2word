# LaTeX to Word Conversion Tool

[简体中文](./README_zh.md)

This project provides a Python script that uses Pandoc and Pandoc-Crossref tools to automatically convert LaTeX files into Word documents in a specified format. It's important to note that there is no perfect method to convert LaTeX to Word, and the Word documents produced by this project are suitable for informal review purposes, with about 5% of the content (such as author information and other non-text elements) possibly requiring manual correction after conversion.

## Features

- Supports the conversion of equations
- Supports automatic numbering and cross-referencing of images, tables, equations, and references
- Supports the conversion of multi-figure images
- Outputs Word documents in a specified format
- Supports Chinese language

The effect is as follows, for more results please see `tests`:

<p align="center">
  <img src=".assets/en-word-1.jpg" width="200"/>
  <img src=".assets/en-word-2.jpg" width="200"/>
</p>

## Quick Start

Ensure all dependencies such as Pandoc and Pandoc-Crossref are properly installed, see [Installing Dependencies](#installing-dependencies). Execute the following command in the command line:

```shell
python ./tex2docx/tex2docx.py --input_texfile <your_texfile> --multifig_dir <dir_saving_temporary_figs> --output_docxfile <your_docxfile> --reference_docfile <your_reference_docfile> --bibfile <your_bibfile> --cslfile <your_cslfile>
```

Replace `<...>` in the command with the appropriate file paths or directory names.

## Installing Dependencies

You will need to install Pandoc, Pandoc-Crossref, and related Python libraries.

### Pandoc

Install Pandoc, see [Pandoc Official Documentation](https://github.com/jgm/pandoc/blob/main/INSTALL.md). It is recommended to download the latest package from [Pandoc Releases](https://github.com/jgm/pandoc/releases).

### Pandoc-Crossref

Install Pandoc-Crossref, see [Pandoc-Crossref Official Documentation](https://github.com/lierdakil/pandoc-crossref). Ensure you download the version that matches your Pandoc installation and configure the path appropriately.

### Related Python Libraries

Install Python dependencies:

```shell
pip install -e .
```

## Usage Instructions and Examples

Supports both command line and script usage methods, ensure required dependencies are installed.

### Command Line Method

Execute the following command in the terminal:

```shell
python ./tex2docx/tex2docx.py --input_texfile <your_texfile> --multifig_dir <dir_saving_temporary_figs> --output_docxfile <your_docxfile> --reference_docfile <your_reference_docfile> --bibfile <your_bibfile> --cslfile <your_cslfile>
```

Parameter explanations:
- `--input_texfile`: Specify the path to the LaTeX file to be converted.
- `--multifig_dir`: Specify the directory for temporarily storing generated multi-figure images.
- `--output_docxfile`: Specify the path for the output Word document.
- `--reference_docfile`: Specify a Word format reference document to ensure consistency in document style.
- `--bibfile`: Specify the BibTeX file for document citations.
- `--cslfile`: Specify the Citation Style Language file to control the formatting of references.
- `--debug`: Enable debug mode to output additional runtime information, helpful for troubleshooting.

For example, in the `tests/en` test case, execute the following command in the repository directory:

```shell
python ./tex2docx/tex2docx.py --input_texfile ./tests/en/main.tex --multifig_dir ./tests/en/multifigs --output_docxfile ./tests/en/main_cli.docx --reference_docfile ./my_temp.docx --bibfile ./tests/ref.bib --cslfile ./ieee.csl
```
You will find the converted `main_cli.docx` file in the `tests/en` directory.

### Script Method

```python
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

You can refer to the example in `tests/test_tex2docx.py`.

## Common Issues

1. The relative positions of multi-figures differ from the original tex file compilation results, as shown in the two images below:

![](.assets/raw_multifig_multi-L-charge-equalization.png)
![](.assets/modified_multifig_multi-L-charge-equalization.png)

This may be due to the original tex file redefining page size parameters; add the relevant tex code to the `MULTIFIG_TEXFILE_TEMPLATE` variable. Here is an example, modify according to actual needs:

```python
import tex2docx

my_multifig_texfile_template = r"""
\documentclass[preview,convert,convert={outext=.png,command=\unexpanded{pdftocairo -r 600 -png \infile}}]{standalone}
\usepackage{graphicx}
\usepackage{subfig}
\usepackage{xeCJK}
\usepackage{geometry}
\newgeometry{
    top=25.4mm, bottom=33.3mm, left=20mm, right=20mm,
    headsep=10.4mm, headheight=5mm, footskip=7.9mm,
}
\graphicspath{{%s}}

\begin{document}
\thispagestyle{empty}
%s
\end{document}
"""

config = {
    'input_texfile': 'tests/en/main.tex',
    'output_docxfile': 'tests/en/main.docx',
    'multifig_dir': 'tests/en/multifigs',
    'reference_docfile': 'my_temp.docx',
    'cslfile': 'ieee.csl',
    'bibfile': 'tests/ref.bib',
    'multifig_texfile_template': my_multifig_texfile_template,
}

converter = tex2docx.LatexToWordConverter(**config)
converter.convert()
```

2. The output Word document's format still does not meet requirements

Modify the styles in the `my_temp.docx` file using Word's style management.

## Implementation Principles

The core of this project is to use Pandoc and Pandoc-Crossref tools to convert LaTeX to Word, configured as follows:

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
    --citeproc --csl ieee.csl \
    -t docx+native_numbering
```

However, the method is not ideal for converting multi-figures. This project extracts the LaTeX file's multi-figure code and uses LaTeX's `convert` and `pdftocairo` tools to automatically compile these images into single large PNG files. Then, these PNG files replace the corresponding image codes in the original LaTeX document and update the references to ensure the multi-figure images are smoothly imported.

## Remaining Issues

1. Chinese figure and table captions still begin with "Figure" and "Table";
2. Author information is not fully converted.

## Other

There are two kinds of people in the world: those who can use LaTeX and those who cannot. The latter often ask the former for Word versions of documents. Thus, the following command line is provided:

```bash
pandoc input.tex -o output.docx\
  --filter pandoc-crossref \
  --reference-doc=my_temp.docx \
  --number-sections \
  -M autoEqnLabels -M tableEqns \
  -M reference-section-title=Reference \
  --bibliography=my_ref.bib \
  --citeproc --csl ieee.csl \
  -t docx+native_numbering
```