# LaTeX 到 Word 文件转换工具

本项目提供一个 Python 脚本，利用 Pandoc 和 Pandoc-Crossref 工具，将 LaTeX 文件自动地按照指定格式转换为 Word 文件。
需要说明的是，目前仍没有能够将 LaTeX 转换为 Word 的完美方法，本项目生成的 Word 文件可满足一般的审阅和修改需求，其中约 5% 的内容（如作者信息等）可能需要在转换后手动更正。

## 特性

- 支持公式的转换；
- 支持图片、表格、公式和参考文献的自动编号及交叉引用；
- 支持多子图；
- 基本支持按照指定格式输出 Word。

## 快速使用

确保已正确安装 Pandoc 和 Pandoc-Crossref 等依赖，详见[安装依赖](#安装依赖)。在命令行中执行以下命令：

```shell
python ./src/tex2docx.py --input_texfile <your_texfile> --multifig_dir <dir_saving_temporary_figs> --output_docxfile <your_docxfile> --reference_docfile <your_reference_docfile> --bibfile <your_bibfile> --cslfile <your_cslfile>
```

将命令中的 `<...>` 替换为相应文件路径或文件夹名称即可。

## 安装依赖

需要安装 Pandoc、Pandoc-Crossref 和相关 Python 库。

### Pandoc

安装 Pandoc，详见 [Pandoc 官方文档](https://github.com/jgm/pandoc/blob/main/INSTALL.md)。建议从 [Pandoc Releases](https://github.com/jgm/pandoc/releases) 下载最新的安装包。

### Pandoc-Crossref

安装 Pandoc-Crossref，详见 [Pandoc-Crossref 官方文档](https://github.com/lierdakil/pandoc-crossref)。确保下载与 Pandoc 版本相匹配的 Pandoc-Crossref，并适当配置路径。

### 相关 Python 库

安装 Python 依赖：

```shell
pip install -e .
```

## 使用说明及案例

支持命令行和脚本两种使用方式，确保已安装所需依赖。

### 命令行使用

在终端执行以下命令：

```shell
python ./src/tex2docx.py --input_texfile <your_texfile> --multifig_dir <dir_saving_temporary_figs> --output_docxfile <your_docxfile> --reference_docfile <your_reference_docfile> --bibfile <your_bibfile> --cslfile <your_cslfile>
```

参数说明：
- `--input_texfile`：指定要转换的 LaTeX 文件的路径。
- `--multifig_dir`：指定临时存放生成的多图文件的目录。
- `--output_docxfile`：指定输出的 Word 文件的路径。
- `--reference_docfile`：指定 Word 输出格式的参考文档，这有助于确保文档样式的一致性。
- `--bibfile`：指定参考文献的 BibTeX 文件，用于文档中的引用。
- `--cslfile`：指定引用样式文件（Citation Style Language），控制参考文献的格式。
- `--debug`：开启调试模式以输出更多的运行信息，有助于排查问题。


以 `tests/en` 测试案例为例，在仓库目录下执行如下命令：

```shell
python ./src/tex2docx.py --input_texfile ./tests/en/main.tex --multifig_dir ./tests/en/multifigs --output_docxfile ./tests/en/main_cli.docx --reference_docfile ./my_temp.docx --bibfile ./tests/ref.bib --cslfile ./ieee.csl
```
则可以在 `tests/en` 目录下找到转换后的 `main_cli.docx` 文件。

### 脚本使用

创建脚本 `my_convert.py` ，写入以下代码，并执行：

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

案例可以参考`tests/test_tex2docx.py`。

## 实现原理及参考资料

该项目核心是使用 Pandoc 和 Pandoc-Crossref 工具实现 LaTeX 到 Word 的转换，具体配置如下：

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

其中，
1. `--lua-filter resolve_equation_labels.lua` 处理公式编号及公式交叉引用，受 Constantin Ahlmann-Eltze 的[脚本](https://gist.githubusercontent.com/const-ae/752ad85c43d92b72865453ea3a77e2dd/raw/28c1815979e5d03cd9ab3638f9befd354797a72b/resolve_equation_labels.lua)启发；
2. `--filter pandoc-crossref` 处理除公式以外的交叉引用；
3. `--reference-doc=my_temp.docx` 依照 `my_temp.docx` 中的样式生成 Word 文件。仓库 [Mingzefei/latex2word](https://github.com/Mingzefei/latex2word) 提供了两个模板文件 `TIE-temp.docx` 和 `my_temp.docx`，前者是 TIE 期刊的投稿 Word 模板（双栏），后者是个人调整出的 Word 模板（单栏，且便于批注）；
4. `--number-sections` 在（子）章节标题前添加数字编号；
5. `-M autoEqnLabels`， `-M tableEqns`设置公式、表格等的编号；
6. `-M reference-sction-title=Reference` 在参考文献部分添加章节标题 Reference；
7. `--biblipgraphy=my_ref.bib` 使用 `ref.bib` 生成参考文献；
8. `--citeproc --csl ieee.csl` 生成的参考文献格式为 `ieee` 。

然而，上述方法在直接处理包含多子图的 Latex 文件时可能遇到图片无法正常导入和引用编号错误等问题。为此，本项目通过提取 LaTeX 文件中的多子图代码，使用 LaTeX 自带的 `convert` 和 `pdftocairo` 工具自动化编译这些图片为单个大图形式的 PNG 文件；然后，这些 PNG 文件将替换原始 LaTeX 文档中的相应图片代码，从而确保多子图形式的图片被顺利导入。具体的实现代码见 `tex2docx.py`。

## 遗留问题

1. 子图引用请统一使用 `\ref{<figure_lab>}(a)` 形式，而非 `\ref{<subfigure_lab>}`（后续会支持直接引用子图）；
2. 导出 Word 文件的图片 caption 格式和作者信息需要手动调整。

## 其他

世界上有两种人，一种人会用 Latex，另一种人不会用 Latex。 后者常常向前者要 Word 版本文件。 因此有了如下一行命令。

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