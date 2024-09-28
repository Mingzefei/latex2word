import concurrent.futures
import glob
import logging
import os
import shutil
import subprocess
import uuid

import regex
from tqdm import tqdm

# Templates and patterns
MULTIFIG_TEXFILE_TEMPLATE = r"""
\documentclass[preview,convert,convert={outext=.png,command=\unexpanded{pdftocairo -r 600 -png \infile}}]{standalone}
\usepackage{graphicx}
\usepackage{subfig}
\graphicspath{{%s}}
\begin{document}
\thispagestyle{empty}
%s
\end{document}
"""
MULTIFIG_FIGENV_TEMPLATE = r"""
\begin{figure}[htbp]
    \centering
    \includegraphics[width=0.8\linewidth]{%s}
    \caption{%s}
    \label{%s}
\end{figure}
"""
FIGURE_PATTERN = r"\\begin{figure}.*?\\end{figure}"
# CAPTION_PATTERN = r'\\caption(\{(?>[^{}]+|(?1))*\})' # this pattern contain {}
CAPTION_PATTERN = r"\\caption\{([^{}]*(?:\{(?1)\}[^{}]*)*)\}"
LABEL_PATTERN = r"\\label\{(.*?)\}"
REF_PATTERN = r"\\ref\{(.*?)\}"
GRAPHICSPATH_PATTERN = r"\\graphicspath\{\{(.+?)\}\}"
INCLUDEGRAPHICS_PATTERN = r"\\includegraphics(?s:.*?)}"


class LatexToWordConverter:
    def __init__(
        self,
        input_texfile,
        multifig_dir,
        output_docxfile,
        reference_docfile,
        bibfile,
        cslfile,
        debug=False,
        multifig_texfile_template=None,
        multifig_figenv_template=None,
    ):
        """
        Initializes the main class of the latex2word tool.

        Args:
            input_texfile (str): The path to the input LaTeX file.
            multifig_dir (str): The directory where the created multi-figure LaTeX files will be stored.
            output_docxfile (str): The path to the output Word document file.
            reference_docfile (str): The path to the reference Word document file.
            bibfile (str): The path to the BibTeX file.
            cslfile (str): The path to the CSL file.
            debug (bool, optional): Whether to enable debug mode. Defaults to False.
            multifig_texfile_template (str, optional): The template for generating multi-figure LaTeX files. Defaults to None.
            multifig_figenv_template (str, optional): The template for the figure environment in multi-figure LaTeX files. Defaults to None.
        """
        # Initialize file paths
        self.input_texfile = os.path.abspath(input_texfile)
        self.multifig_dir = os.path.abspath(multifig_dir)
        self.output_texfile = os.path.abspath(
            input_texfile.replace(".tex", "_modified.tex")
        )
        self.output_docxfile = os.path.abspath(output_docxfile)
        self.reference_docfile = os.path.abspath(reference_docfile)
        self.bibfile = os.path.abspath(bibfile)
        self.cslfile = os.path.abspath(cslfile)
        self.luafile = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "resolve_equation_labels.lua"
        )

        # Initialize other attributes
        self._raw_content = None
        self._modified_content = None
        self._raw_fig_contents = None
        self._raw_graphicspath = None
        self._created_multifig_texfiles = {}
        self._figurepackage = None
        self._multifig_texfile_template = (
            multifig_texfile_template
            if multifig_texfile_template
            else MULTIFIG_TEXFILE_TEMPLATE
        )
        self._multifig_figenv_template = (
            multifig_figenv_template
            if multifig_figenv_template
            else MULTIFIG_FIGENV_TEMPLATE
        )

        # Initialize logger
        self.logger = logging.getLogger(f"{__name__}_{uuid.uuid4().hex[:6]}")
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        )
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        self.logger.setLevel(logging.DEBUG if debug else logging.INFO)

        self.logger.debug(f"Init input texfile: {self.input_texfile}")
        self.logger.debug(f"Init multifig_dir: {self.multifig_dir}")
        self.logger.debug(f"Init output texfile: {self.output_texfile}")
        self.logger.debug(f"Init output docxfile: {self.output_docxfile}")
        self.logger.debug(f"Init reference docfile: {self.reference_docfile}")
        self.logger.debug(f"Init bibfile: {self.bibfile}")
        self.logger.debug(f"Init cslfile: {self.cslfile}")
        self.logger.debug(f"Init luafile: {self.luafile}")

    def _match_pattern(self, pattern, content, mode="last"):
        """
        Matches a pattern in the given content and returns the result.

        Args:
            pattern (str): The regular expression pattern to match.
            content (str): The content to search for the pattern.
            mode (str, optional): Specifies whether to find 'all' matches, 'first' match or 'last' match. Defaults to 'first'.

        Returns:
            str or list: The matched result(s) if found, otherwise None.

        Raises:
            ValueError: If mode is not 'all', 'first' or 'last'.

        """
        matches = regex.findall(pattern, content, regex.DOTALL)
        if mode == "all":
            return matches
        elif mode == "first":
            return matches[0] if matches else None
        elif mode == "last":
            return matches[-1] if matches else None
        else:
            raise ValueError("mode must be 'all', 'first' or 'last'")

    def analyze_texfile(self):
        """
        Analyzes the input LaTeX file and extracts relevant information.

        This method reads the content of the input LaTeX file, matches patterns to extract figure environments,
        and determines the path to the directory containing the figures.

        Returns:
            None
        """
        with open(self.input_texfile, "r") as file:
            self._raw_content = file.read()
        self._raw_fig_contents = self._match_pattern(
            FIGURE_PATTERN, self._raw_content, mode="all"
        )

        # Determine the figure package used in the LaTeX file (subfigure or subfig)
        if (
            r"\usepackage{subfig}" in self._raw_content
            or r"\begin{subfig}" in self._raw_content
            or r"\subfloat" in self._raw_content
        ):
            self._figurepackage = "subfig"
        elif (
            r"\usepackage{subfigure}" in self._raw_content
            or r"\begin{subfigure}" in self._raw_content
            or r"\subfigure" in self._raw_content
        ):
            self._figurepackage = "subfigure"
        else:
            pass
        self.logger.debug(f"Analyze figure package : {self._figurepackage}")
        # Change self._multifig_texfile_template based on the figure package used
        if self._figurepackage == "subfigure":
            self._multifig_texfile_template = self._multifig_texfile_template.replace(
                r"\usepackage{subfig}", r"\usepackage{subfigure}"
            )

        # Determine graphicspath
        graphicspath = self._match_pattern(
            GRAPHICSPATH_PATTERN, self._raw_content, mode="last"
        )
        if graphicspath:
            self._raw_graphicspath = os.path.abspath(
                os.path.join(os.path.dirname(self.input_texfile), graphicspath)
            )
        else:
            self._raw_graphicspath = os.path.abspath(
                os.path.dirname(self.input_texfile)
            )

            self.logger.debug(
                f"Init input figure directory(_raw_graphicspath): {self._raw_graphicspath}"
            )

        self.logger.info(
            f"Read {os.path.basename(self.input_texfile)} texfile and found {len(self._raw_fig_contents)} figenvs."
        )

        # Determine if _raw_fig_contents contains Chinese characters
        if any(
            regex.search(r"[\u4e00-\u9fff]", fig_content)
            for fig_content in self._raw_fig_contents
        ):
            # If Chinese characters are found, add \usepackage{xeCJK} to _multifig_texfile_template
            self._multifig_texfile_template = self._multifig_texfile_template.replace(
                r"\begin{document}", r"\usepackage{xeCJK}" + "\n" + r"\begin{document}"
            )

    def create_multifig_texfiles(self):
        """
        Create multiple tex files for each figure in the raw figure contents.

        This method takes the raw figure contents and creates separate tex files for each figure.
        It comments out the captions in the LaTeX code and uses the labels or default counters as filenames, with a prefix 'multifig_'.
        The created tex files are saved in the multifig_dir directory.

        Returns:
            None
        """
        default_counter = 0

        if os.path.exists(self.multifig_dir):
            shutil.rmtree(self.multifig_dir)
        os.makedirs(self.multifig_dir)

        for figure_content in self._raw_fig_contents:
            # Define a function to prepend a '%' character to each caption line
            # This effectively comments out the caption lines in LaTeX

            def comment_caption(match):
                # Add a '%' character before each caption line
                commented = "\n".join("%" + line for line in match.group(0).split("\n"))
                return commented

            # Apply the function to each caption in the figure content
            # This comments out all captions in the figure content
            processed_figure_content = regex.sub(
                CAPTION_PATTERN, comment_caption, figure_content
            )

            # Find the last label of the figure
            labels = self._match_pattern(
                LABEL_PATTERN, processed_figure_content, mode="all"
            )
            if labels:
                filename = labels[-1]
                # Remove prefix if it exists
                if (
                    filename.startswith("fig:")
                    or filename.startswith("fig-")
                    or filename.startswith("fig_")
                ):
                    filename = filename[4:]
                else:
                    pass
            else:
                # If no label is found, use the default counter as the filename
                filename = f"fig{default_counter}"

            filename = "multifig_" + filename + ".tex"

            # Check if the filename is already used
            while filename in self._created_multifig_texfiles.values():
                filename = filename.split(".")[0] + f"_{uuid.uuid4().hex[:6]}.tex"

            # Add the filename to the set of created filenames
            self._created_multifig_texfiles[default_counter] = filename
            default_counter += 1

            # Define the tex file content
            file_content = self._multifig_texfile_template % (
                os.path.abspath(os.path.join("..", self._raw_graphicspath)),
                processed_figure_content,
            )

            # Create the tex file
            file_path = os.path.join(self.multifig_dir, filename)
            with open(file_path, "w") as file:
                file.write(file_content)

            self.logger.info(
                f"Created texfile {os.path.basename(file_path)} under {os.path.basename(self.multifig_dir)}."
            )

    def compile_multifig_texfile(self, texfile):
        """
        Compiles a LaTeX file containing multiple figures.

        Args:
            texfile (str): The path to the LaTeX file to be compiled.

        Returns:
            str: The path to the compiled LaTeX file.

        Raises:
            subprocess.CalledProcessError: If the compilation fails.

        """
        self.logger.debug(
            f"Command: xelatex -shell-escape -synctex=1 -interaction=nonstopmode {texfile}"
        )
        if self.logger.level == logging.DEBUG:
            subprocess.run(
                [
                    "xelatex",
                    "-shell-escape",
                    "-synctex=1",
                    "-interaction=nonstopmode",
                    texfile,
                ],
                check=True,
            )
        else:
            subprocess.run(
                [
                    "xelatex",
                    "-shell-escape",
                    "-synctex=1",
                    "-interaction=nonstopmode",
                    texfile,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True,
            )
        return texfile

    def compile_multifig_texfiles(self):
        """
        Compiles multiple .tex files and creates corresponding .png files.

        This method changes the current working directory to the output directory specified by `self.multifig_dir`.
        It then uses a `ProcessPoolExecutor` to compile each .tex file in parallel. After each .tex file is compiled,
        it checks for the created .png files and performs the necessary checks and renaming operations.
        Finally, it changes back to the original working directory.

        Returns:
            None
        """
        # Save the current working directory
        cwd = os.getcwd()

        # Change to the output directory
        os.chdir(self.multifig_dir)

        try:
            # Compile each .tex file
            with concurrent.futures.ProcessPoolExecutor() as executor:
                futures = {
                    executor.submit(self.compile_multifig_texfile, texfile)
                    for texfile in self._created_multifig_texfiles.values()
                }
                for future in tqdm(
                    concurrent.futures.as_completed(futures),
                    total=len(futures),
                    desc="Compiling texfiles",
                ):
                    # Check and Rename
                    created_pngfiles = glob.glob(
                        f'{future.result().replace(".tex", "")}-*.png'
                    )
                    num_files = len(created_pngfiles)

                    if num_files == 0:
                        self.logger.error(f"No pngfile created for {future.result()}")
                    elif num_files > 1:
                        self.logger.warn(
                            f"Multiple pngfiles created for {future.result()}"
                        )

                    if num_files >= 1:
                        new_name = future.result().replace(".tex", ".png")
                        os.rename(created_pngfiles[0], new_name)
        finally:
            # Change back to the original working directory
            os.chdir(cwd)

        self.logger.info("Created and renamed pngfiles.")

    def create_modified_texfile(self):
        """
        creates a modified .tex file by replacing figure contents and updating \graphicspath.

        This method iterates over the figure contents in the raw content of the .tex file and replaces them with modified
        figure contents. It also updates the \graphicspath to the specified multifig_dir. Finally, it writes the modified
        content to a new .tex file.

        Returns:
            None
        """
        self._modified_content = self._raw_content

        # Redefine \graphicspath
        self._modified_content = regex.sub(
            GRAPHICSPATH_PATTERN,
            r"\\graphicspath{{%s}}" % self.multifig_dir,
            self._modified_content,
        )

        # Replace the figure contents with modified figure contents
        for fig_index, fig_content in enumerate(self._raw_fig_contents):
            # Replace the old figure content with the new figure content
            multifig_caption = self._match_pattern(
                CAPTION_PATTERN, fig_content, mode="last"
            )
            multifig_label = "multifig:" + os.path.basename(
                self._created_multifig_texfiles[fig_index]
            ).replace(".tex", "")
            png_file = os.path.basename(
                self._created_multifig_texfiles[fig_index]
            ).replace(".tex", ".png")
            self.logger.debug(
                f"Modify texfile with:\ncaption: {multifig_caption}\nlabel: {multifig_label}\npng_file: {png_file}"
            )
            modified_figure_content = self._multifig_figenv_template % (
                png_file,
                multifig_caption,
                multifig_label,
            )
            self._modified_content = self._modified_content.replace(
                fig_content, modified_figure_content
            )

            # Update the references to the subfigures
            subfig_labels = []

            # Find all occurrences of '\includegraphics' in the figure content
            includegraphics_occurrences = regex.finditer(
                INCLUDEGRAPHICS_PATTERN, fig_content
            )
            for occurrence in includegraphics_occurrences:
                # Get the content after the current '\includegraphics'
                content_after_includegraphics = fig_content[occurrence.end() :]
                # Search for '\label' in the content after the current '\includegraphics'
                label_occurrence = regex.search(
                    LABEL_PATTERN, content_after_includegraphics
                )
                # If '\label' is found and there is no other '\includegraphics' or '\caption' before it
                if (
                    label_occurrence
                    and "includegraphics"
                    not in content_after_includegraphics[: label_occurrence.end()]
                    and "caption"
                    not in content_after_includegraphics[: label_occurrence.end()]
                ):
                    # Add the label to the list of subfigure labels
                    subfig_labels.append(label_occurrence.group(1))
                else:
                    # If no '\label' is found, add an empty string to the list of subfigure labels
                    subfig_labels.append("")

            for subfig_index, subfig_label in enumerate(subfig_labels):
                if subfig_label:
                    raw_subfig_ref = r"\ref{%s}" % subfig_label
                    modified_subfig_ref = r"\ref{%s}(%s)" % (
                        multifig_label,
                        chr(ord("a") + subfig_index),
                    )
                    self._modified_content = self._modified_content.replace(
                        raw_subfig_ref, modified_subfig_ref
                    )

            # Update the references to the figure (multifig)
            fig_label = self._match_pattern(LABEL_PATTERN, fig_content, mode="last")
            if fig_label in subfig_labels:
                # this figure label is used in subfigure
                pass
            else:
                # this figure label is used as a whole figure
                raw_fig_ref = r"\ref{%s}" % fig_label
                modified_fig_ref = r"\ref{%s}" % multifig_label
                self._modified_content = self._modified_content.replace(
                    raw_fig_ref, modified_fig_ref
                )

        # Redefine \graphicspath
        self._modified_content = regex.sub(
            GRAPHICSPATH_PATTERN,
            r"\\graphicspath{{%s}}" % self.multifig_dir,
            self._modified_content,
        )

        # Write the modified text content to a new .tex file
        with open(self.output_texfile, "w") as f:
            f.write(self._modified_content)

        self.logger.info(f"Created {os.path.basename(self.output_texfile)} tex file.")

    def convert_modified_texfile(self):
        """
        Converts a modified TeX file to a DOCX file using pandoc.

        Raises:
            Exception: If pandoc or pandoc-crossref is not installed.

        Returns:
            None
        """
        # Check if pandoc and pandoc-crossref are installed
        if shutil.which("pandoc") is None:
            raise Exception(
                "pandoc is not installed. Please install it before running this script."
            )
        if shutil.which("pandoc-crossref") is None:
            raise Exception(
                "pandoc-crossref is not installed. Please install it before running this script."
            )

        # Define the command
        command = [
            "pandoc",
            self.output_texfile,
            "-o",
            self.output_docxfile,
            "--lua-filter",
            self.luafile,
            "--filter",
            "pandoc-crossref",
            "--reference-doc=" + self.reference_docfile,
            "--number-sections",
            "-M",
            "autoEqnLabels",
            "-M",
            "tableEqns",
            "-M",
            "reference-section-title=Reference",
            "--bibliography=" + self.bibfile,
            "--citeproc",
            "--csl",
            self.cslfile,
            "-t",
            "docx+native_numbering",
        ]

        # Save the current working directory
        cwd = os.getcwd()
        # Change to the output directory
        os.chdir(os.path.dirname(self.output_texfile))

        try:
            # Run the command
            subprocess.run(command, check=True)
        finally:
            # Change back to the original working directory
            os.chdir(cwd)

        self.logger.info(
            f"Converted {os.path.basename(self.output_texfile)} texfile to {os.path.basename(self.output_docxfile)} docxfile."
        )

    def convert(self):
        """
        Converts a LaTeX file to Word format.

        This method performs the following steps:
        1. Analyzes the LaTeX file.
        2. creates multiple figure LaTeX files.
        3. Compiles the multiple figure LaTeX files.
        4. creates the modified LaTeX file.
        5. Converts the modified LaTeX file to Word format.
        """
        self.analyze_texfile()
        self.create_multifig_texfiles()
        self.compile_multifig_texfiles()
        self.create_modified_texfile()
        self.convert_modified_texfile()