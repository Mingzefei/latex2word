"""Constants and patterns used throughout the tex2docx package."""

from typing import Final


class TexPatterns:
    """Regular expression patterns for LaTeX parsing."""
    
    INCLUDE: Final[str] = r"\\include\{(.+?)\}"
    FIGURE: Final[str] = r"\\begin{figure}.*?\\end{figure}"
    TABLE: Final[str] = r"\\begin{table}.*?\\end{table}"
    CAPTION: Final[str] = r"\\caption\{([^{}]*(?:\{(?1)\}[^{}]*)*)\}"
    LABEL: Final[str] = r"\\label\{(.*?)\}"
    REF: Final[str] = r"\\ref\{(.*?)\}"
    GRAPHICSPATH: Final[str] = r"\\graphicspath\{\{(.+?)\}\}"
    INCLUDEGRAPHICS: Final[str] = r"\\includegraphics(?s:.*?)\}"
    COMMENT: Final[str] = r"((?<!\\)%.*\n)"
    SUBFIG_PACKAGE: Final[str] = r"\\usepackage\{subfig\}"
    SUBFIGURE_PACKAGE: Final[str] = r"\\usepackage\{subfigure\}"
    SUBFIG_ENV: Final[str] = r"\\begin\{subfig\}|\\subfloat"
    SUBFIGURE_ENV: Final[str] = r"\\begin\{subfigure\}|\\subfigure"
    CHINESE_CHAR: Final[str] = r"[\u4e00-\u9fff]"
    LINEWIDTH: Final[str] = r"\\linewidth|\\textwidth"
    CM_UNIT: Final[str] = r"\d+cm"
    CONTINUED_FLOAT: Final[str] = r"\\ContinuedFloat"


class TexTemplates:
    """LaTeX template strings."""
    
    BASE_MULTIFIG_TEXFILE: Final[str] = r"""
\documentclass[preview,convert,convert={outext=.png,command=\unexpanded{pdftocairo -r 500 -png \infile}}]{standalone}
\usepackage{graphicx}
% FIGURE_PACKAGE_PLACEHOLDER %
\usepackage{booktabs}
\usepackage{multirow}
\usepackage{makecell}
\usepackage{setspace}
\usepackage{siunitx}
% CJK_PACKAGE_PLACEHOLDER %
\graphicspath{{{GRAPHICSPATH_PLACEHOLDER}}}
\begin{document}
\thispagestyle{empty}
{FIGURE_CONTENT_PLACEHOLDER}
\end{document}
"""

    MULTIFIG_FIGENV: Final[str] = r"""
\begin{figure}[htbp]
    \centering
    \includegraphics[width=\linewidth]{%s}
    \caption{%s}
    \label{%s}
\end{figure}
"""

    MODIFIED_TABENV: Final[str] = r"""
\begin{table}[htbp]
    \centering
    \caption{%s}
    \label{%s}
    \begin{tabular}{l}
    \includegraphics[width=\linewidth]{%s}
    \end{tabular}
\end{table}
"""


class FilenamePatterns:
    """Patterns for filename sanitization."""
    
    INVALID_CHARS: Final[str] = r'[\\/*?:"<>|]'
    REPLACEMENT_CHAR: Final[str] = "_"


class PandocOptions:
    """Default options for Pandoc conversion."""
    
    BASIC_OPTIONS: Final[list[str]] = [
        "--number-sections",
        "-M", "autoEqnLabels",
        "-M", "tableEqns",
        "-t", "docx+native_numbering",
    ]
    
    FILTER_OPTIONS: Final[list[str]] = [
        "--filter", "pandoc-crossref",
    ]
    
    CITATION_OPTIONS: Final[list[str]] = [
        "-M", "reference-section-title=References",
        "--citeproc",
    ]


class CompilerOptions:
    """Options for LaTeX compilation."""
    
    XELATEX_OPTIONS: Final[list[str]] = [
        "-shell-escape",
        "-synctex=1", 
        "-interaction=nonstopmode",
    ]
    
    MAX_WORKERS: Final[int] = 4  # Default number of parallel compilation workers
