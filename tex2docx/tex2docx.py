import concurrent.futures
import glob
import logging
import os
import shutil
import subprocess
import uuid
from pathlib import Path

import regex
from tqdm import tqdm

class LatexToWordConverter:
    # --- Templates ---
    MULTIFIG_TEXFILE_TEMPLATE = r"""
\documentclass[preview,convert,convert={outext=.png,command=\unexpanded{pdftocairo -r 500 -png \infile}},varwidth=\maxdimen]{standalone}
\usepackage{graphicx}
\usepackage{subfig}
\usepackage{booktabs}
\usepackage{multirow}
\usepackage{makecell}
\usepackage{setspace}
\usepackage{siunitx}
\graphicspath{{%s}}
\begin{document}
\thispagestyle{empty}
%s
\end{document}
"""
    MULTIFIG_FIGENV_TEMPLATE = r"""
\begin{figure}[htbp]
    \centering
    \includegraphics[width=\linewidth]{%s}
    \caption{%s}
    \label{%s}
\end{figure}
"""
    MODIFIED_TABENV_TEMPLATE = r"""
\begin{table}[htbp]
    \centering
    \caption{%s}
    \label{%s}
    \begin{tabular}{l}
    \includegraphics[width=\linewidth]{%s}
    \end{tabular}
\end{table}
"""
    # --- Patterns ---
    INCLUD_PATTERN = r"\\include\{(.+?)\}"
    FIGURE_PATTERN = r"\\begin{figure}.*?\\end{figure}"
    TABLE_PATTERN = r"\\begin{table}.*?\\end{table}"
    CAPTION_PATTERN = r"\\caption\{([^{}]*(?:\{(?1)\}[^{}]*)*)\}"
    LABEL_PATTERN = r"\\label\{(.*?)\}"
    REF_PATTERN = r"\\ref\{(.*?)\}"
    GRAPHICSPATH_PATTERN = r"\\graphicspath\{\{(.+?)\}\}"
    INCLUDEGRAPHICS_PATTERN = r"\\includegraphics(?s:.*?)}"


    def __init__(
        self,
        input_texfile,
        output_docxfile,
        bibfile=None,
        cslfile=None,
        reference_docfile=None,
        debug=False,
        multifig_texfile_template=None,
        multifig_figenv_template=None,
        fix_table=True,
    ):
        """
        Initializes the main class of the latex2word tool.

        Args:
            input_texfile (str or Path): The path to the input LaTeX file.
            output_docxfile (str or Path): The path to the output Word document file.
            bibfile (str, optional): The path to the BibTeX file. Defaults to None (use the first .bib file found in the same directory as input_texfile).
            cslfile (str, optional): The path to the CSL file. Defaults to None (use the built-in ieee.csl file).
            reference_docfile (str, optional): The path to the reference Word document file. Defaults to None (use the built-in default_temp.docx file).
            debug (bool, optional): Whether to enable debug mode. Defaults to False.
            multifig_texfile_template (str, optional): The template for generating multi-figure LaTeX files. Defaults to None.
            multifig_figenv_template (str, optional): The template for the figure environment in multi-figure LaTeX files. Defaults to None.
            fix_table (bool, optional): Whether to fix tables. Defaults to True.
        """
        # Initialize file paths
        self.input_texfile = os.path.abspath(str(input_texfile))
        self.output_texfile = os.path.abspath(
            str(input_texfile).replace(".tex", "_modified.tex")
        )
        self.output_docxfile = os.path.abspath(str(output_docxfile))
        self.reference_docfile = (
            os.path.abspath(reference_docfile)
            if reference_docfile
            else os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "default_temp.docx"
            )
        )
        self.temp_subtexfile_dir = os.path.join(
            os.path.dirname(self.input_texfile), "temp_subtexfile_dir"
        )

        if bibfile:  # if bibfile is provided, use it
            self.bibfile = os.path.abspath(str(bibfile))
        else:  # if bibfile is not provided, search for bibfile in the same directory as input_texfile
            bibfile = glob.glob(
                os.path.join(os.path.dirname(self.input_texfile), "*.bib")
            )
            self.bibfile = bibfile[0] if bibfile else None

        self.cslfile = (
            os.path.abspath(cslfile)
            if cslfile
            else os.path.join(os.path.dirname(os.path.abspath(__file__)), "ieee.csl")
        )
        self.luafile = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "resolve_equation_labels.lua"
        )

        # Initialize other attributes
        self._raw_content = None
        self._clean_content = None
        self._modified_content = None
        self._clean_fig_contents = None
        self._raw_graphicspath = None
        self._created_multifig_texfiles = {}
        self._created_tab_texfiles = {}
        self._figurepackage = None
        self._multifig_texfile_template = (
            multifig_texfile_template
            if multifig_texfile_template
            else self.MULTIFIG_TEXFILE_TEMPLATE # Use class attribute
        )
        self._multifig_figenv_template = (
            multifig_figenv_template
            if multifig_figenv_template
            else self.MULTIFIG_FIGENV_TEMPLATE # Use class attribute
        )
        self._modified_tabenv_template = self.MODIFIED_TABENV_TEMPLATE # Use class attribute
        self.fix_table = fix_table

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
        self.logger.debug(f"Init multifig_dir: {self.temp_subtexfile_dir}")
        self.logger.debug(f"Init output texfile: {self.output_texfile}")
        self.logger.debug(f"Init output docxfile: {self.output_docxfile}")
        self.logger.debug(f"Init reference docfile: {self.reference_docfile}")
        self.logger.debug(f"Init bibfile: {self.bibfile}")
        self.logger.debug(f"Init cslfile: {self.cslfile}")
        self.logger.debug(f"Init luafile: {self.luafile}")
        self.logger.debug(f"Fix tabel: {self.fix_table}")

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

        This method reads the content of the input LaTeX file, matches patterns to extract table and figure environments,
        and determines the path to the directory containing the figures.

        Returns:
            None
        """
        with open(self.input_texfile, "r") as file:
            self._raw_content = file.read()
        self.logger.info(f"Read {os.path.basename(self.input_texfile)} texfile.")

        # Remove all LaTeX comments
        self._clean_content = regex.sub(r"((?<!\\)%.*\n)", "", self._raw_content)
        self.logger.debug("Removed all LaTeX comments.")

        # Replace all \include{...} with the content of the included file
        include_files = regex.findall(self.INCLUD_PATTERN, self._clean_content) # Use class attribute
        for include_file in include_files:
            # if include_file is not end with .tex, add .tex
            include_filename = (
                include_file + ".tex"
                if not include_file.endswith(".tex")
                else include_file
            )
            include_file_path = os.path.join(
                os.path.dirname(self.input_texfile), include_filename
            )
            with open(include_file_path, "r") as file:
                include_content = file.read()
            self._clean_content = self._clean_content.replace(
                r"\include{" + include_file + "}", include_content
            )
        self.logger.debug(
            "Replaced all \\include{...} with the content of the included file."
        )
        # Remove all LaTeX comments
        self._clean_content = regex.sub(r"((?<!\\)%.*\n)", "", self._clean_content)
        self.logger.debug("Removed all LaTeX comments of included files.")

        # Get all figure environments in the LaTeX file
        self._clean_fig_contents = self._match_pattern(
            self.FIGURE_PATTERN, self._clean_content, mode="all" # Use class attribute
        )
        self.logger.info(f"Found {len(self._clean_fig_contents)} figenvs.")

        # Get all table environments in the LaTeX file
        self._clean_tab_contents = self._match_pattern(
            self.TABLE_PATTERN, self._clean_content, mode="all" # Use class attribute
        )
        self.logger.info(f"Found {len(self._clean_tab_contents)} tabenvs.")

        # Determine the figure package used in the LaTeX file (subfigure or subfig)
        if (
            r"\usepackage{subfig}" in self._clean_content
            or r"\begin{subfig}" in self._clean_content
            or r"\subfloat" in self._clean_content
        ):
            self._figurepackage = "subfig"
        elif (
            r"\usepackage{subfigure}" in self._clean_content
            or r"\begin{subfigure}" in self._clean_content
            or r"\subfigure" in self._clean_content
        ):
            self._figurepackage = "subfigure"
        else:
            pass
        # Change self._multifig_texfile_template based on the figure package used
        if self._figurepackage == "subfigure":
            self._multifig_texfile_template = self._multifig_texfile_template.replace(
                r"\usepackage{subfig}", r"\usepackage{subfigure}"
            )
        self.logger.debug(f"Analyze figure package : {self._figurepackage}")

        # Determine graphicspath
        graphicspath = self._match_pattern(
            self.GRAPHICSPATH_PATTERN, self._clean_content, mode="last" # Use class attribute
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

        # Determine if _clean_fig_contents contains Chinese characters
        if any(
            regex.search(r"[\u4e00-\u9fff]", fig_content)
            for fig_content in self._clean_fig_contents
        ):
            # If Chinese characters are found, add \usepackage{xeCJK} to _multifig_texfile_template
            self._multifig_texfile_template = self._multifig_texfile_template.replace(
                r"\begin{document}", r"\usepackage{xeCJK}" + "\n" + r"\begin{document}"
            )

    def _comment_caption_in_content(self, content):
        """Helper function to comment out captions in a given content string."""
        def comment_caption_match(match):
            return "\n".join("%" + line for line in match.group(0).split("\n"))
        return regex.sub(self.CAPTION_PATTERN, comment_caption_match, content) # Use class attribute

    def create_multifig_texfiles(self):
        """
        Create multiple tex files for each figure in the raw figure contents.

        This method takes the raw figure contents and creates separate tex files for each figure.
        It comments out the captions in the LaTeX code and uses the labels or default counters as filenames, with a prefix 'multifig_'.
        The created tex files are saved in the self.temp_subtexfile_dir directory.

        Returns:
            None
        """
        default_counter = 0

        if os.path.exists(self.temp_subtexfile_dir):
            shutil.rmtree(self.temp_subtexfile_dir)
        os.makedirs(self.temp_subtexfile_dir)

        for figure_content in self._clean_fig_contents:
            # Comment out captions
            processed_figure_content = self._comment_caption_in_content(figure_content)

            # Find the last label of the figure
            labels = self._match_pattern(
                self.LABEL_PATTERN, processed_figure_content, mode="all" # Use class attribute
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

            # if the figure content contains \linewidth or \textwidth, change varwidth to 21cm
            # to fix the bug of too large pdf and png
            if (
                r"\linewidth" in processed_figure_content
                or r"\textwidth" in processed_figure_content
            ):
                if self._figurepackage == "subfig":
                    file_content = file_content.replace("varwidth=\\maxdimen", "")
                elif self._figurepackage == "subfigure":
                    file_content = file_content.replace(
                        "varwidth=\\maxdimen", "varwidth=21cm"
                    )
                else:
                    pass

            # Create the tex file
            file_path = os.path.join(self.temp_subtexfile_dir, filename)
            with open(file_path, "w") as file:
                file.write(file_content)

            self.logger.info(
                f"Created texfile {os.path.basename(file_path)} under {os.path.basename(self.temp_subtexfile_dir)}."
            )

    def compile_png_texfile(self, texfile):
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
        result = subprocess.run(
            [
                "xelatex",
                "-shell-escape",
                "-synctex=1",
                "-interaction=nonstopmode",
                texfile,
            ],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        with open(texfile.replace(".tex", ".out"), "w") as f:
            f.write(result.stdout)
        with open(texfile.replace(".tex", ".err"), "w") as f:
            f.write(result.stderr)
        return texfile

    def compile_png_texfiles(self):
        """
        Compiles multiple .tex files and creates corresponding .png files.

        This method changes the current working directory to the output directory specified by `self.temp_subtexfile_dir`.
        It then uses a `ProcessPoolExecutor` to compile each .tex file in parallel. After each .tex file is compiled,
        it checks for the created .png files and performs the necessary checks and renaming operations.
        Finally, it changes back to the original working directory.

        Returns:
            None
        """
        # Save the current working directory
        cwd = os.getcwd()

        # Change to the output directory
        os.chdir(self.temp_subtexfile_dir)

        try:
            # Compile each .tex file
            with concurrent.futures.ProcessPoolExecutor() as executor:
                if self.fix_table:
                    _input_subtexfiles = list(
                        self._created_multifig_texfiles.values()
                    ) + list(self._created_tab_texfiles.values())
                else:
                    _input_subtexfiles = list(self._created_multifig_texfiles.values())
                # Randomize the order of the tasks to balance the compilation time
                # random.shuffle(_input_subtexfiles)
                futures = {
                    executor.submit(self.compile_png_texfile, texfile)
                    for texfile in _input_subtexfiles
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

    def create_table_texfiles(self):
        """
        Create multiple tex files for each table in the raw figure contents.

        This method takes the raw table contents and creates separate tex files for each table.
        It comments out the captions in the LaTeX code and uses the labels or default counters as filenames, with a prefix 'table_'.
        The created tex files are saved in the self.temp_subtexfile_dir directory.

        Returns:
            None
        """
        default_counter = 0

        for table_content in self._clean_tab_contents:
            # Comment out captions
            processed_table_content = self._comment_caption_in_content(table_content)

            # Find the last label of the table
            labels = self._match_pattern(
                self.LABEL_PATTERN, processed_table_content, mode="all" # Use class attribute
            )
            if labels:
                filename = labels[-1]
                # Remove prefix if it exists
                if (
                    filename.startswith("tab:")
                    or filename.startswith("tab-")
                    or filename.startswith("tab_")
                ):
                    filename = filename[4:]
                else:
                    pass
            else:
                # If no label is found, use the default counter as the filename
                filename = f"tab{default_counter}"

            filename = "tab_" + filename + ".tex"

            # Check if the filename is already used
            while filename in self._created_tab_texfiles.values():
                filename = filename.split(".")[0] + f"_{uuid.uuid4().hex[:6]}.tex"

            # Add the filename to the set of created filenames
            self._created_tab_texfiles[default_counter] = filename
            default_counter += 1

            # Define the tex file content
            file_content = self._multifig_texfile_template % (
                os.path.abspath(os.path.join("..", self._raw_graphicspath)),
                processed_table_content,
            )

            # 如果出现 数字+cm 的形式
            if regex.search(r"\d+cm", processed_table_content):
                file_content = file_content.replace(
                    "varwidth=\\maxdimen", "varwidth=21cm"
                )
                
            file_content = file_content.replace("\ContinuedFloat", "")

            # Create the tex file
            file_path = os.path.join(self.temp_subtexfile_dir, filename)
            with open(file_path, "w") as file:
                file.write(file_content)

            self.logger.info(
                f"Created texfile {os.path.basename(file_path)} under {os.path.basename(self.temp_subtexfile_dir)}."
            )

    def _update_figure_references(self, original_fig_content, multifig_label):
        """Updates references for a single figure environment."""
        subfig_labels = []
        # Find all occurrences of '\includegraphics'
        includegraphics_occurrences = regex.finditer(
            self.INCLUDEGRAPHICS_PATTERN, original_fig_content # Use class attribute
        )
        for occurrence in includegraphics_occurrences:
            content_after_includegraphics = original_fig_content[occurrence.end() :]
            label_occurrence = regex.search(
                self.LABEL_PATTERN, content_after_includegraphics # Use class attribute
            )
            if (
                label_occurrence
                and "includegraphics" not in content_after_includegraphics[: label_occurrence.end()]
                and "caption" not in content_after_includegraphics[: label_occurrence.end()]
            ):
                subfig_labels.append(label_occurrence.group(1))
            else:
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

        fig_label = self._match_pattern(self.LABEL_PATTERN, original_fig_content, mode="last") # Use class attribute
        if fig_label and fig_label not in subfig_labels:
            raw_fig_ref = r"\ref{%s}" % fig_label
            modified_fig_ref = r"\ref{%s}" % multifig_label
            self._modified_content = self._modified_content.replace(
                raw_fig_ref, modified_fig_ref
            )

    def _replace_figure_environments(self):
        """Replaces original figure environments with image includes and updates references."""
        for fig_index, fig_content in enumerate(self._clean_fig_contents):
            multifig_caption = self._match_pattern(
                self.CAPTION_PATTERN, fig_content, mode="last" # Use class attribute
            )
            multifig_label = "multifig:" + os.path.basename(
                self._created_multifig_texfiles[fig_index]
            ).replace(".tex", "")
            png_file = os.path.basename(
                self._created_multifig_texfiles[fig_index]
            ).replace(".tex", ".png")

            self.logger.debug(
                f"Modify figure in texfile with:\ncaption: {multifig_caption}\nlabel: {multifig_label}\npng_file: {png_file}"
            )
            modified_figure_content = self._multifig_figenv_template % (
                png_file,
                multifig_caption,
                multifig_label,
            )
            self._modified_content = self._modified_content.replace(
                fig_content, modified_figure_content
            )

            # Update references related to this figure
            self._update_figure_references(fig_content, multifig_label)

    def _replace_table_environments(self):
        """Replaces original table environments with image includes."""
        for tab_index, tab_content in enumerate(self._clean_tab_contents):
            tab_caption = self._match_pattern(
                self.CAPTION_PATTERN, tab_content, mode="last" # Use class attribute
            )
            tab_label = self._match_pattern(self.LABEL_PATTERN, tab_content, mode="last") # Use class attribute
            png_file = os.path.basename(
                self._created_tab_texfiles[tab_index]
            ).replace(".tex", ".png")

            self.logger.debug(
                f"Modify table in texfile with:\ncaption: {tab_caption}\nlabel: {tab_label}\npng_file: {png_file}"
            )
            modified_table_content = self._modified_tabenv_template % (
                tab_caption,
                tab_label,
                png_file,
            )
            self._modified_content = self._modified_content.replace(
                tab_content, modified_table_content
            )

    def create_modified_texfile(self):
        """
        Creates a modified .tex file by replacing figure/table contents and updating \graphicspath.
        """
        self._modified_content = self._clean_content

        # Replace figure environments and update references
        self._replace_figure_environments()

        # Replace table environments if fix_table is True
        if self.fix_table:
            self._replace_table_environments()

        # Redefine \graphicspath
        self._modified_content = regex.sub(
            self.GRAPHICSPATH_PATTERN, # Use class attribute
            r"\\graphicspath{{%s}}" % self.temp_subtexfile_dir,
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
            "-t",
            "docx+native_numbering",
        ]

        # Only add citation-related options if bibfile is provided
        if self.bibfile is not None:
            command.extend([
                "-M",
                "reference-section-title=Reference",
                "--bibliography=" + self.bibfile,
                "--citeproc",
                "--csl",
                self.cslfile,
            ])

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
        self.logger.debug(f"Command: {' '.join(command)}")

    def clean_temp_files(self):
        """
        Clean the temporary files created during the conversion process.

        This method deletes the temporary files created during the conversion process, including the multifig_dir directory.

        Returns:
            None
        """
        if os.path.exists(self.temp_subtexfile_dir):
            shutil.rmtree(self.temp_subtexfile_dir)
            self.logger.info(f"Removed temporary directory: {self.temp_subtexfile_dir}")
        else:
            self.logger.info("Temporary directory not found, skipping removal.")


        if os.path.exists(self.output_texfile):
            try:
                os.remove(self.output_texfile)
                self.logger.info(f"Removed {os.path.basename(self.output_texfile)}.")
            except OSError as e:
                self.logger.error(f"Error removing {self.output_texfile}: {e}")
        else:
             self.logger.info(f"{os.path.basename(self.output_texfile)} not found, skipping removal.")

    def convert(self):
        """
        Converts a LaTeX file to Word format.

        This method performs the following steps:
        1. Analyzes the LaTeX file.
        2. Creates multiple figure (and table) LaTeX files.
        3. Compiles the multiple LaTeX files.
        4. Creates the modified LaTeX file.
        5. Converts the modified LaTeX file to Word format.

        Note: If the logger level is not set to DEBUG, temporary files will be cleaned up after conversion.
        """
        try: # Add try...finally block for cleanup
            self.analyze_texfile()
            # Consider moving temp dir creation here using _setup_temp_dir
            self.create_multifig_texfiles()
            if self.fix_table:
                self.create_table_texfiles()
            self.compile_png_texfiles()
            self.create_modified_texfile()
            self.convert_modified_texfile()
        finally: # Ensure cleanup happens even if errors occur
            if not self.logger.level == logging.DEBUG:
                self.clean_temp_files()
