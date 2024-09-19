# README

There are two types of people in the world: those who use LaTeX and those who don't. The latter often ask the former for Word versions of their files. Therefore, the following command line is created:

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

The files used by this command can be found in this repository.
The following content will partially solve the above predicament, enabling you to comprehend and execute this command smoothly.

##  Installation

1. pandoc: Refer to the [official documentation](https://github.com/jgm/pandoc/blob/main/INSTALL.md) for instruction and installation. It is recommended to download the latest deb installation package from [Releases · jgm/pandoc (github.com)](https://github.com/jgm/pandoc/releases) and use `sudo dpkg -i /path/to/the/deb/file` to install it.
2. pandoc-crossref: Refer to the [official documentation](https://github.com/lierdakil/pandoc-crossref) for instruction and installation. **NOTE: Download the version that matches your pandoc version and move the executable file `pandoc-crossref` to `/usr/bin`, or specify the specific file when using the above command.**

## Usage

1. `--filter pandoc-crossref` processes cross-references.
2. `--reference-doc=my_temp.docx` processes the converted `output.docx` according to the style in `my_temp.docx`. There are two template files, `TIE-temp.docx` and `my_temp.docx`, in the repository [Mingzefei/latex2word](https://github.com/Mingzefei/latex2word). The former is the Word template for TIE journal submissions (two columns), and the latter is a Word template adjusted by the author (single column, large font, suitable for annotations).
3. `--number-sections` adds numerical numbering before (sub)chapter titles.
4. `-M autoEqnLabels`, `-M tableEqns` sets the numbering of equations, tables, etc.
5. `-M reference-sction-title=Reference` adds the chapter title "Reference" to the reference section.
6. `--biblipgraphy=my_ref.bib` generates the reference list using `my_ref.bib`.
7. The `--citeproc --csl ieee.csl` generates the references in the `ieee` format.

### Running Test 

Go to `./test` and run `bash ./run.sh`.

## Outstanding Issues

1. "Error" may raise when opening the converted docx document. This is likely due to the complexity of the tex file being converted. Try reducing the number of images and avoiding the use of tikz, for example.
2. Poor support for subfigures, particularly with regard to numbering. If the tex file does not involve subfigures, use the `-t docx+native_numbering` option to optimize numbering for images and tables.
3. References to equations appear in the form `[<label>]`. See [Equation numbering in MS Word · Issue](https://github.com/lierdakil/pandoc-crossref/issues/221) for more information. No solution has been found yet, but global replace commands can be used in Word to replace them.
4. The image size set in tex does not work in the converted docx file. A method for setting the caption style for images has not been found yet.
