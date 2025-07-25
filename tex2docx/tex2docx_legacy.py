import concurrent.futures
import logging
import os
import shutil
import subprocess
import uuid
from pathlib import Path

import regex
from tqdm import tqdm

class LatexToWordConverter:
    # --- Patterns ---
    INCLUD_PATTERN = r"\\include\{(.+?)\}"
    FIGURE_PATTERN = r"\\begin{figure}.*?\\end{figure}"
    TABLE_PATTERN = r"\\begin{table}.*?\\end{table}"
    CAPTION_PATTERN = r"\\caption\{([^{}]*(?:\{(?1)\}[^{}]*)*)\}"
    LABEL_PATTERN = r"\\label\{(.*?)\}"
    REF_PATTERN = r"\\ref\{(.*?)\}" # Keep for potential future use
    GRAPHICSPATH_PATTERN = r"\\graphicspath\{\{(.+?)\}\}"
    INCLUDEGRAPHICS_PATTERN = r"\\includegraphics(?s:.*?)\}"
    COMMENT_PATTERN = r"((?<!\\)%.*\n)"
    SUBFIG_PACKAGE_PATTERN = r"\\usepackage\{subfig\}"
    SUBFIGURE_PACKAGE_PATTERN = r"\\usepackage\{subfigure\}"
    SUBFIG_ENV_PATTERN = r"\\begin\{subfig\}|\\subfloat"
    SUBFIGURE_ENV_PATTERN = r"\\begin\{subfigure\}|\\subfigure"
    CHINESE_CHAR_PATTERN = r"[\u4e00-\u9fff]"
    LINEWIDTH_PATTERN = r"\\linewidth|\\textwidth"
    CM_UNIT_PATTERN = r"\d+cm"
    CONTINUED_FLOAT_PATTERN = r"\\ContinuedFloat"

    # --- Templates ---
    # Base template, figure package will be adjusted later
    BASE_MULTIFIG_TEXFILE_TEMPLATE = r"""
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

    def __init__(
        self,
        input_texfile: str | Path,
        output_docxfile: str | Path,
        bibfile: str | Path | None = None,
        cslfile: str | Path | None = None,
        reference_docfile: str | Path | None = None,
        debug: bool = False,
        multifig_texfile_template: str | None = None, # Deprecated
        multifig_figenv_template: str | None = None,
        fix_table: bool = True,
    ) -> None:
        """
        Initializes the main class of the latex2word tool.

        Args:
            input_texfile: The path to the input LaTeX file.
            output_docxfile: The path to the output Word document file.
            bibfile: The path to the BibTeX file. Defaults to None
                (use the first .bib file found in the same directory
                 as input_texfile).
            cslfile: The path to the CSL file. Defaults to None
                (use the built-in ieee.csl file).
            reference_docfile: The path to the reference Word document
                file. Defaults to None (use the built-in
                default_temp.docx file).
            debug: Whether to enable debug mode. Defaults to False.
            multifig_texfile_template: Deprecated. Template is now
                generated dynamically.
            multifig_figenv_template: The template for the figure
                environment in multi-figure LaTeX files. Defaults to
                class attribute.
            fix_table: Whether to fix tables by converting them to
                images. Defaults to True.
        """
        self._setup_paths(
            input_texfile, output_docxfile, bibfile, cslfile, reference_docfile
        )
        self._setup_logger(debug)
        self._setup_options(fix_table, multifig_figenv_template)

        # Internal state variables initialized
        self._raw_content: str | None = None
        self._clean_content: str | None = None
        self._modified_content: str | None = None
        self._clean_fig_contents: list[str] = []
        self._clean_tab_contents: list[str] = []
        self._raw_graphicspath: Path | None = None
        self._figurepackage: str | None = None # 'subfig', 'subfigure', or None
        self._contains_chinese: bool = False
        self._created_multifig_texfiles: dict[int, str] = {} # {index: filename}
        self._created_tab_texfiles: dict[int, str] = {} # {index: filename}

        self.logger.debug("LatexToWordConverter initialized.")
        self._log_initial_paths()

    def _setup_paths(
        self,
        input_texfile: str | Path,
        output_docxfile: str | Path,
        bibfile: str | Path | None,
        cslfile: str | Path | None,
        reference_docfile: str | Path | None,
    ) -> None:
        """Sets up all the necessary file and directory paths."""
        self.input_texfile = Path(input_texfile).resolve()
        self.output_docxfile = Path(output_docxfile).resolve()
        self.output_texfile = self.input_texfile.with_name(
            f"{self.input_texfile.stem}_modified.tex"
        )
        self.temp_subtexfile_dir = self.input_texfile.parent / "temp_subtexfile_dir"

        # Determine bibfile path
        if bibfile:
            self.bibfile = Path(bibfile).resolve()
        else:
            bib_files = list(self.input_texfile.parent.glob("*.bib"))
            self.bibfile = bib_files[0].resolve() if bib_files else None

        # Determine CSL file path
        default_csl_path = Path(__file__).parent / "ieee.csl"
        self.cslfile = Path(cslfile).resolve() if cslfile else default_csl_path.resolve()

        # Determine reference doc path
        default_ref_doc_path = Path(__file__).parent / "default_temp.docx"
        self.reference_docfile = (
            Path(reference_docfile).resolve()
            if reference_docfile
            else default_ref_doc_path.resolve()
        )

        # Lua filter path
        self.luafile = (Path(__file__).parent / "resolve_equation_labels.lua").resolve()

    def _setup_logger(self, debug: bool) -> None:
        """Configures the logger."""
        self.logger = logging.getLogger(f"{__name__}_{uuid.uuid4().hex[:6]}")
        # Avoid adding handlers multiple times if instantiated repeatedly
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter(
                "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
            )
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
        self.logger.setLevel(logging.DEBUG if debug else logging.INFO)

    def _setup_options(
        self, fix_table: bool, multifig_figenv_template: str | None
    ) -> None:
        """Sets up conversion options."""
        self.fix_table = fix_table
        self._multifig_figenv_template = (
            multifig_figenv_template
            if multifig_figenv_template
            else self.MULTIFIG_FIGENV_TEMPLATE
        )
        self._modified_tabenv_template = self.MODIFIED_TABENV_TEMPLATE

    def _log_initial_paths(self) -> None:
        """Logs the initial paths for debugging."""
        self.logger.debug(f"Input texfile: {self.input_texfile}")
        self.logger.debug(f"Temp subfile dir: {self.temp_subtexfile_dir}")
        self.logger.debug(f"Output texfile: {self.output_texfile}")
        self.logger.debug(f"Output docxfile: {self.output_docxfile}")
        self.logger.debug(f"Reference docfile: {self.reference_docfile}")
        self.logger.debug(f"Bibfile: {self.bibfile}")
        self.logger.debug(f"CSL file: {self.cslfile}")
        self.logger.debug(f"Lua filter: {self.luafile}")
        self.logger.debug(f"Fix table: {self.fix_table}")

    @staticmethod
    def _match_pattern(
        pattern: str, content: str, mode: str = "last"
    ) -> str | list[str] | None:
        """
        Matches a pattern in the given content and returns the result.

        Args:
            pattern: The regular expression pattern to match.
            content: The content to search for the pattern.
            mode: Specifies whether to find 'all' matches, 'first'
                match or 'last' match. Defaults to 'last'.

        Returns:
            The matched result(s) if found, otherwise None.

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

    def _read_and_preprocess_tex(self) -> None:
        """Reads the main tex file, handles includes, and removes comments."""
        try:
            with open(self.input_texfile, "r", encoding="utf-8") as file:
                self._raw_content = file.read()
            self.logger.info(f"Read {self.input_texfile.name}.")
        except Exception as e:
            self.logger.error(f"Error reading input file {self.input_texfile}: {e}")
            raise

        # Remove comments first
        clean_content = regex.sub(self.COMMENT_PATTERN, "", self._raw_content)
        self.logger.debug("Removed comments from main file.")

        # Handle includes iteratively to support nested includes
        while True:
            includes_found = regex.findall(self.INCLUD_PATTERN, clean_content)
            if not includes_found:
                break # No more includes found

            made_replacement_in_pass = False
            current_pass_processed = set() # Track includes processed in this specific pass

            for include_name in includes_found:
                # Construct the full directive to check against processed_includes
                include_directive = r"\include{" + include_name + "}"

                # Skip if already processed in *this pass* to avoid infinite loop on same level
                if include_directive in current_pass_processed:
                    continue

                include_filename = Path(
                    f"{include_name}.tex"
                    if not include_name.lower().endswith(".tex")
                    else include_name
                )
                include_file_path = self.input_texfile.parent / include_filename

                if include_file_path.exists():
                    try:
                        with open(include_file_path, "r", encoding="utf-8") as infile:
                            include_content = infile.read()
                        # Remove comments from included file before inserting
                        clean_include_content = regex.sub(
                            self.COMMENT_PATTERN, "", include_content
                        )
                        # Replace only the first occurrence found in this pass
                        if include_directive in clean_content:
                            clean_content = clean_content.replace(
                                include_directive, clean_include_content, 1
                            )
                            self.logger.debug(f"Included content from {include_filename}.")
                            made_replacement_in_pass = True
                            current_pass_processed.add(include_directive)
                        else:
                             # This might happen if nested includes modified the content
                             self.logger.warning(f"Directive '{include_directive}' not found for replacement, possibly due to nesting.")

                    except Exception as e:
                        self.logger.warning(
                            f"Could not read or process include file "
                            f"{include_file_path}: {e}"
                        )
                        # Replace with error comment to avoid breaking parsing
                        clean_content = clean_content.replace(
                            include_directive, f"% Error including {include_filename} %", 1
                        )
                        current_pass_processed.add(include_directive)
                else:
                    self.logger.warning(f"Include file not found: {include_file_path}")
                    # Replace with error comment
                    clean_content = clean_content.replace(
                        include_directive, f"% Include file not found: {include_filename} %", 1
                    )
                    current_pass_processed.add(include_directive)

            if not made_replacement_in_pass:
                break # Exit loop if no replacements were made in a full pass

        self._clean_content = clean_content
        self.logger.debug("Finished processing includes and comments.")


    def _analyze_tex_structure(self) -> None:
        """Analyzes the cleaned content for figures, tables, graphicspath, etc."""
        if not self._clean_content:
            self.logger.error("Clean content is not available for analysis.")
            return

        # Extract figures and tables
        self._clean_fig_contents = self._match_pattern(
            self.FIGURE_PATTERN, self._clean_content, mode="all"
        ) or [] # Ensure list even if no matches
        self.logger.info(f"Found {len(self._clean_fig_contents)} figure environments.")
        self._clean_tab_contents = self._match_pattern(
            self.TABLE_PATTERN, self._clean_content, mode="all"
        ) or [] # Ensure list even if no matches
        self.logger.info(f"Found {len(self._clean_tab_contents)} table environments.")

        # Determine figure package
        if regex.search(self.SUBFIG_PACKAGE_PATTERN, self._clean_content) or \
           regex.search(self.SUBFIG_ENV_PATTERN, self._clean_content):
            self._figurepackage = "subfig"
        elif regex.search(self.SUBFIGURE_PACKAGE_PATTERN, self._clean_content) or \
             regex.search(self.SUBFIGURE_ENV_PATTERN, self._clean_content):
            self._figurepackage = "subfigure"
        else:
            self._figurepackage = None # No specific subfigure package detected
        self.logger.debug(f"Detected figure package: {self._figurepackage}")

        # Determine graphicspath
        graphicspath_match = self._match_pattern(
            self.GRAPHICSPATH_PATTERN, self._clean_content, mode="last"
        )
        if graphicspath_match:
            # Handle multiple paths if present, taking the first one
            first_path = graphicspath_match.split('}{')[0]
            # Resolve relative to input tex file parent
            self._raw_graphicspath = (self.input_texfile.parent / first_path).resolve()
        else:
            self._raw_graphicspath = self.input_texfile.parent.resolve()
        self.logger.debug(f"Determined graphics path: {self._raw_graphicspath}")

        # Check for Chinese characters in figure/table content
        combined_content = "".join(self._clean_fig_contents + self._clean_tab_contents)
        if regex.search(self.CHINESE_CHAR_PATTERN, combined_content):
            self._contains_chinese = True
            self.logger.debug("Detected Chinese characters in figures/tables.")
        else:
            self._contains_chinese = False

    def _prepare_temp_directory(self) -> None:
        """Creates or cleans the temporary directory."""
        try:
            if self.temp_subtexfile_dir.exists():
                shutil.rmtree(self.temp_subtexfile_dir)
                self.logger.debug(f"Removed existing temp directory: {self.temp_subtexfile_dir}")
            self.temp_subtexfile_dir.mkdir(parents=True)
            self.logger.debug(f"Created temp directory: {self.temp_subtexfile_dir}")
        except OSError as e:
            self.logger.error(f"Error managing temp directory {self.temp_subtexfile_dir}: {e}")
            raise

    @staticmethod
    def _comment_out_captions(content: str) -> str:
        """Helper function to comment out captions in a given content string."""
        def comment_caption_match(match: regex.Match) -> str:
            # Ensure each line of the caption block starts with %
            caption_block = match.group(0).strip()
            return "\n".join("% " + line for line in caption_block.split("\n"))
        # Use CAPTION_PATTERN from the class
        return regex.sub(LatexToWordConverter.CAPTION_PATTERN, comment_caption_match, content)

    def _generate_subfile_content(
        self, original_content: str, graphicspath_rel_to_temp: Path
    ) -> str:
        """Generates the content for a single subfile (.tex)."""
        processed_content = self._comment_out_captions(original_content)
        # Remove \ContinuedFloat as it breaks standalone compilation
        processed_content = regex.sub(self.CONTINUED_FLOAT_PATTERN, "", processed_content)

        # Start with base template
        file_content = self.BASE_MULTIFIG_TEXFILE_TEMPLATE

        # --- Remove logic related to varwidth=\maxdimen ---
        # The option is now removed directly from the template.

        # Set figure package placeholder
        figure_package_lines = ""
        captionsetup_line = "" # Initialize captionsetup line
        if self._figurepackage == "subfig":
            # Load caption package explicitly before subfig
            figure_package_lines = "\\usepackage{caption}\n\\usepackage{subfig}"
            # Add captionsetup to potentially improve compatibility with standalone
            captionsetup_line = "\\captionsetup{list=false}" # Try disabling list generation
        elif self._figurepackage == "subfigure":
            figure_package_lines = f"\\usepackage{{{self._figurepackage}}}"
        # else: no specific package needed

        # Replace package placeholder
        if figure_package_lines:
            file_content = file_content.replace(
                "% FIGURE_PACKAGE_PLACEHOLDER %", figure_package_lines
            )
        else:
            # Remove the placeholder line if no specific package is needed
            file_content = file_content.replace("% FIGURE_PACKAGE_PLACEHOLDER %\n", "")

        # Replace captionsetup placeholder
        if captionsetup_line:
            file_content = file_content.replace(
                "% CAPTIONSETUP_PLACEHOLDER %", captionsetup_line
            )
        else:
            # Remove the placeholder line if no captionsetup is needed
            file_content = file_content.replace("% CAPTIONSETUP_PLACEHOLDER %\n", "")


        # Set CJK package placeholder if needed
        if self._contains_chinese:
            cjk_line = "\\usepackage{xeCJK}"
            file_content = file_content.replace("% CJK_PACKAGE_PLACEHOLDER %", cjk_line)
        else:
            # Remove the CJK placeholder line if not needed
            file_content = file_content.replace("% CJK_PACKAGE_PLACEHOLDER %\n", "")

        # Set graphicspath relative to the temp directory
        # Ensure path format is suitable for LaTeX (forward slashes)
        latex_graphicspath = str(graphicspath_rel_to_temp.as_posix())
        file_content = file_content.replace(
            "{GRAPHICSPATH_PLACEHOLDER}", latex_graphicspath
        )

        # Insert the actual figure/table content
        file_content = file_content.replace(
            "{FIGURE_CONTENT_PLACEHOLDER}", processed_content
        )

        return file_content

    def _create_subfiles(
        self, content_list: list[str], prefix: str, storage_dict: dict[int, str]
    ) -> None:
        """Creates multiple tex files for figures or tables."""
        default_counter = 0
        created_filenames = set(storage_dict.values()) # Track filenames used

        # Calculate relative graphicspath from temp dir to original graphicspath
        try:
            # Use os.path.relpath for robust relative path calculation
            graphicspath_rel_str = os.path.relpath(
                self._raw_graphicspath, self.temp_subtexfile_dir
            )
            graphicspath_rel_to_temp = Path(graphicspath_rel_str)
        except ValueError:
            # If paths are on different drives (Windows), use absolute path
            graphicspath_rel_to_temp = self._raw_graphicspath.resolve()
            self.logger.warning(
                "Graphics path and temp directory are on different drives. "
                "Using absolute graphics path in subfiles."
            )

        for index, item_content in enumerate(content_list):
            # Find the last label to use as a base for the filename
            labels = self._match_pattern(self.LABEL_PATTERN, item_content, mode="all")
            if labels:
                base_filename = labels[-1]
                # Clean common prefixes (fig:, tab:, etc.)
                for pfx in ["fig:", "fig-", "fig_", "tab:", "tab-", "tab_"]:
                    if base_filename.startswith(pfx):
                        base_filename = base_filename[len(pfx):]
                        break
            else:
                # Use counter if no label is found
                base_filename = f"{prefix}{default_counter}"

            # Sanitize filename (replace invalid characters)
            safe_base_filename = regex.sub(r'[\\/*?:"<>|]', "_", base_filename)
            filename = f"{prefix}_{safe_base_filename}.tex"

            # Ensure unique filename by adding suffix if needed
            unique_suffix = ""
            original_filename_stem = f"{prefix}_{safe_base_filename}"
            while filename in created_filenames:
                unique_suffix = f"_{uuid.uuid4().hex[:4]}" # Shorter suffix
                filename = f"{original_filename_stem}{unique_suffix}.tex"

            storage_dict[index] = filename
            created_filenames.add(filename)
            default_counter += 1 # Increment counter regardless of label usage

            # Generate content and write file
            try:
                file_content = self._generate_subfile_content(
                    item_content, graphicspath_rel_to_temp
                )
                file_path = self.temp_subtexfile_dir / filename
                with open(file_path, "w", encoding="utf-8") as file:
                    file.write(file_content)
                self.logger.info(f"Created subfile: {filename}")
            except Exception as e:
                self.logger.error(f"Error writing subfile {filename}: {e}")
                # Remove from dict if creation failed to avoid compilation attempt
                if index in storage_dict:
                    del storage_dict[index]
                    created_filenames.remove(filename) # Also remove from set

    def _create_figure_subfiles(self) -> None:
        """Creates temporary tex files for each figure environment."""
        self._create_subfiles(
            self._clean_fig_contents, "multifig", self._created_multifig_texfiles
        )

    def _create_table_subfiles(self) -> None:
        """Creates temporary tex files for each table environment."""
        self._create_subfiles(
            self._clean_tab_contents, "tab", self._created_tab_texfiles
        )

    def _compile_single_subfile(self, texfile_path: Path) -> tuple[Path, bool]:
        """
        Compiles a single subfile .tex to .png using xelatex.

        Args:
            texfile_path: Absolute path to the .tex subfile.

        Returns:
            A tuple containing the original texfile path and a boolean
            indicating success (True) or failure (False).
        """
        texfile_name = texfile_path.name
        self.logger.debug(f"Compiling: {texfile_name}")
        # Command to run xelatex with necessary options
        command = [
            "xelatex",
            "-shell-escape", # Needed for standalone's convert option
            "-synctex=1",
            "-interaction=nonstopmode", # Prevent stopping on errors
            texfile_name, # Run command relative to the temp dir
        ]
        self.logger.debug(f"Running command: {' '.join(command)}")
        try:
            # Run compilation within the temporary directory
            result = subprocess.run(
                command,
                check=True, # Raise exception on non-zero exit code
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='replace', # Handle potential encoding errors in output
                cwd=self.temp_subtexfile_dir # IMPORTANT: Execute in temp dir
            )
            # Write compilation logs for debugging purposes
            with open(texfile_path.with_suffix(".out"), "w", encoding="utf-8") as f_out:
                f_out.write(result.stdout)
            with open(texfile_path.with_suffix(".err"), "w", encoding="utf-8") as f_err:
                f_err.write(result.stderr)

            # --- PNG File Handling ---
            # The standalone class with 'convert' option should create one PNG
            # named exactly like the tex file but with .png extension.
            expected_png_path = texfile_path.with_suffix(".png")

            if not expected_png_path.exists():
                # Check if pdftocairo produced a different name (e.g., with page number)
                base_name = texfile_path.stem
                png_pattern = f"{base_name}*.png"
                created_pngs = list(self.temp_subtexfile_dir.glob(png_pattern))

                if not created_pngs:
                     self.logger.error(
                         f"No png file found for {texfile_name}. Check "
                         f"{texfile_path.with_suffix('.log')} and "
                         f"{texfile_path.with_suffix('.err')}"
                     )
                     return texfile_path, False # Indicate failure

                # Rename the first found PNG to the expected name
                target_png_name = expected_png_path
                try:
                     # If target exists (e.g., from previous run), remove it
                     if target_png_name.exists():
                         target_png_name.unlink()
                     created_pngs[0].rename(target_png_name)
                     self.logger.debug(f"Renamed {created_pngs[0].name} to {target_png_name.name}")
                     # Log warning if multiple PNGs were created unexpectedly
                     if len(created_pngs) > 1:
                         self.logger.warning(
                             f"Multiple PNGs found for {texfile_name}; used "
                             f"{created_pngs[0].name}. Others: "
                             f"{[p.name for p in created_pngs[1:]]}"
                         )
                         # Optionally remove extra PNGs here
                except Exception as e:
                     self.logger.error(
                         f"Error renaming {created_pngs[0].name} for "
                         f"{texfile_name}: {e}"
                     )
                     return texfile_path, False # Indicate failure
            else:
                self.logger.debug(f"Found expected PNG: {expected_png_path.name}")

            # Log successful compilation at debug level
            self.logger.debug(f"Successfully compiled {texfile_name}.")
            return texfile_path, True # Indicate success

        except subprocess.CalledProcessError as e:
            self.logger.error(f"Compilation failed for {texfile_name}. Return code: {e.returncode}")
            # Log stdout/stderr from the failed process
            stdout = e.stdout or ""
            stderr = e.stderr or ""
            self.logger.error(f"Stdout:\n{stdout[-1000:]}") # Log last 1000 chars
            self.logger.error(f"Stderr:\n{stderr[-1000:]}") # Log last 1000 chars
            # Write full logs even on failure
            try:
                with open(texfile_path.with_suffix(".out"), "w", encoding="utf-8") as f_out:
                    f_out.write(stdout)
                with open(texfile_path.with_suffix(".err"), "w", encoding="utf-8") as f_err:
                    f_err.write(stderr)
                # Also try to save the .log file if it exists
                log_file = texfile_path.with_suffix(".log")
                if log_file.exists():
                     self.logger.error(f"Check {log_file.name} for detailed LaTeX errors.")
            except Exception as log_e:
                 self.logger.error(f"Could not write log files for failed compilation: {log_e}")
            return texfile_path, False # Indicate failure
        except Exception as e:
            # Catch other potential errors (e.g., file system issues)
            self.logger.error(f"An unexpected error occurred during compilation of {texfile_name}: {e}")
            return texfile_path, False # Indicate failure


    def _compile_all_subfiles(self) -> None:
        """Compiles all generated .tex subfiles into .png files in parallel."""
        subfiles_to_compile = list(self._created_multifig_texfiles.values())
        if self.fix_table:
            subfiles_to_compile.extend(list(self._created_tab_texfiles.values()))

        if not subfiles_to_compile:
            self.logger.info("No subfiles generated for compilation.")
            return

        # Get full paths for the compilation function
        full_path_subfiles = [self.temp_subtexfile_dir / fname for fname in subfiles_to_compile]

        successful_compilations = 0
        failed_compilations: list[str] = []

        # Use ProcessPoolExecutor for parallel compilation
        # Adjust max_workers based on available CPU cores if needed
        with concurrent.futures.ProcessPoolExecutor() as executor:
            # Create a mapping from future to tex path for easier tracking
            futures = {
                executor.submit(self._compile_single_subfile, tex_path): tex_path
                for tex_path in full_path_subfiles
            }
            progress = tqdm(
                concurrent.futures.as_completed(futures),
                total=len(futures),
                desc="Compiling subfiles",
                unit="file"
            )

            for future in progress:
                tex_path = futures[future] # Get the path associated with the future
                try:
                    # Result is a tuple: (original_path, success_flag)
                    original_path, success = future.result()
                    if success:
                        successful_compilations += 1
                    else:
                        failed_compilations.append(original_path.name)
                except Exception as exc:
                    # Catch exceptions raised during the task execution itself
                    self.logger.error(
                        f"Subfile {tex_path.name} generated an exception "
                        f"during processing: {exc}"
                    )
                    failed_compilations.append(tex_path.name)

        self.logger.info(
            f"Subfile compilation finished. Success: {successful_compilations}, "
            f"Failed: {len(failed_compilations)}."
        )
        if failed_compilations:
            self.logger.warning(f"Failed compilations: {', '.join(failed_compilations)}")
            # Decide if processing should stop or continue with available PNGs.
            # Current behavior: Continue.

    def _update_references(
        self, original_content: str, new_label_base: str
    ) -> None:
        """
        Updates references (\ref) in the main content based on subfigure labels.

        This needs to modify self._modified_content directly.

        Args:
            original_content: The original LaTeX content of the
                figure/table environment being replaced.
            new_label_base: The base label assigned to the generated
                single image (e.g., "multifig:fig_xyz").
        """
        subfig_labels: list[str | None] = []
        # Find all \includegraphics commands within the original content
        includegraphics_matches = list(
            regex.finditer(self.INCLUDEGRAPHICS_PATTERN, original_content)
        )

        # For each includegraphics, find the *next* label that isn't
        # preceded by another includegraphics or caption within the scope
        # between the current includegraphics and the potential label.
        for i, img_match in enumerate(includegraphics_matches):
            # Define the search area starting after the current includegraphics
            start_pos = img_match.end()
            # Limit search area to before the next includegraphics, if any
            end_pos = (
                includegraphics_matches[i+1].start()
                if i + 1 < len(includegraphics_matches)
                else len(original_content)
            )
            search_area = original_content[start_pos:end_pos]

            label_match = regex.search(self.LABEL_PATTERN, search_area)
            found_label = None
            if label_match:
                # Check if a caption appears *before* this label within the search area
                content_before_label = search_area[:label_match.start()]
                if not regex.search(self.CAPTION_PATTERN, content_before_label):
                    found_label = label_match.group(1)

            subfig_labels.append(found_label) # Append label or None

        # Update references for found subfigure labels
        for subfig_index, subfig_label in enumerate(subfig_labels):
            if subfig_label:
                # Pattern to find \ref{subfig_label}, ensuring it's a whole ref
                # Use regex.escape to handle special characters in labels
                raw_subfig_ref_pattern = r"\\ref\{" + regex.escape(subfig_label) + r"\}"
                # Replacement: \ref{new_label_base}(a), \ref{new_label_base}(b), etc.
                subfig_char = chr(ord('a') + subfig_index)
                modified_subfig_ref = r"\\ref{%s}(%s)" % (new_label_base, subfig_char)

                # Use regex.sub for safer replacement on self._modified_content
                count_before = self._modified_content.count(f"\\ref{{{subfig_label}}}")
                self._modified_content = regex.sub(
                    raw_subfig_ref_pattern, modified_subfig_ref, self._modified_content
                )
                count_after = self._modified_content.count(modified_subfig_ref)
                if count_before > 0: # Log only if replacements were likely made
                    self.logger.debug(
                        f"Replaced ref '{subfig_label}' with "
                        f"'{new_label_base}({subfig_char})' "
                        f"({count_after} instance(s) potentially updated)."
                    )


        # Update reference for the main figure/table label (if it exists and
        # is different from the subfig labels found)
        main_label_match = regex.search(self.LABEL_PATTERN, original_content)
        if main_label_match:
             main_label = main_label_match.group(1)
             # Check if this main label was already handled as a subfig label
             is_subfig_label = any(lbl == main_label for lbl in subfig_labels if lbl)

             if not is_subfig_label:
                 # Pattern for the main label reference
                 raw_fig_ref_pattern = r"\\ref\{" + regex.escape(main_label) + r"\}"
                 modified_fig_ref = r"\\ref{%s}" % new_label_base
                 # Replace in the main modified content
                 count_before = self._modified_content.count(f"\\ref{{{main_label}}}")
                 self._modified_content = regex.sub(
                     raw_fig_ref_pattern, modified_fig_ref, self._modified_content
                 )
                 count_after = self._modified_content.count(modified_fig_ref)
                 if count_before > 0:
                     self.logger.debug(
                         f"Replaced main ref '{main_label}' with "
                         f"'{new_label_base}' "
                         f"({count_after} instance(s) potentially updated)."
                     )


    def _replace_environments(
        self,
        original_contents: list[str],
        created_files_dict: dict[int, str],
        env_template: str,
        label_prefix: str,
    ) -> None:
        """
        Replaces original figure or table environments with image includes
        and updates references. Modifies self._modified_content.
        """
        if self._modified_content is None:
             self.logger.error("Cannot replace environments, modified content is not initialized.")
             return

        processed_indices = set() # Track indices already processed

        # Iterate through original contents and their indices
        for index, original_content in enumerate(original_contents):
            if index in processed_indices:
                continue # Skip if already processed (e.g., identical content)

            if index not in created_files_dict:
                self.logger.warning(
                    f"Skipping replacement for item {index} as its subfile "
                    f"({label_prefix}) was not created or compiled successfully."
                )
                continue

            subfile_basename = Path(created_files_dict[index]).stem
            png_filename = f"{subfile_basename}.png"
            # The graphicspath in the modified tex file points to the temp dir.
            # Therefore, here we only need the filename for includegraphics.
            png_path_in_tex_str = png_filename
            # Ensure forward slashes for LaTeX path consistency (though just filename here)
            # png_path_in_tex_str = Path(png_filename).as_posix()

            # Extract caption and create a new label
            caption = self._match_pattern(self.CAPTION_PATTERN, original_content, mode="last") or ""
            # New label based on prefix and subfile name (without .tex)
            # Sanitize label base similar to filename sanitization
            safe_label_base = regex.sub(r'[\\/*?:"<>|]', "_", subfile_basename)
            new_label = f"{label_prefix}:{safe_label_base}"

            self.logger.debug(
                f"Replacing environment {index} ({label_prefix}) with:\n"
                f"  Caption: {caption[:50]}...\n" # Log truncated caption
                f"  Label: {new_label}\n"
                f"  PNG File: {png_path_in_tex_str}"
            )

            # Create the new environment content (figure or table wrapping the image)
            # Template arguments depend on the specific template structure
            if env_template == self.MULTIFIG_FIGENV_TEMPLATE:
                 # Template: image_path, caption, label
                 modified_env_content = env_template % (
                     png_path_in_tex_str,
                     caption,
                     new_label,
                 )
            elif env_template == self.MODIFIED_TABENV_TEMPLATE:
                 # Template: caption, label, image_path
                 modified_env_content = env_template % (
                     caption,
                     new_label,
                     png_path_in_tex_str,
                 )
            else:
                 self.logger.error(f"Unknown environment template for {label_prefix}.")
                 continue # Skip this replacement

            # Replace the original environment block with the new one
            # Use count=1 to replace only the first match for this specific content
            # This helps if identical figure/table environments exist
            if original_content in self._modified_content:
                self._modified_content = self._modified_content.replace(
                    original_content, modified_env_content, 1
                )
                processed_indices.add(index) # Mark as processed
            else:
                self.logger.warning(
                    f"Original content for item {index} ({label_prefix}) not found "
                    f"in modified content. Skipping replacement."
                )
                continue # Skip reference update if replacement failed

            # Update references related to this environment *after* replacement
            self._update_references(original_content, new_label)


    def _create_modified_texfile(self) -> None:
        """
        Creates the modified .tex file by replacing environments,
        updating references, adjusting graphicspath, and writing the result.
        """
        if self._clean_content is None:
            self.logger.error("Clean content not available. Cannot create modified tex file.")
            return

        # Start with the cleaned content (includes resolved, comments removed)
        self._modified_content = self._clean_content

        # Replace figure environments and update associated references
        self._replace_environments(
            self._clean_fig_contents,
            self._created_multifig_texfiles,
            self._multifig_figenv_template,
            "multifig" # Label prefix for figures
        )

        # Replace table environments if fix_table is True
        if self.fix_table:
            self._replace_environments(
                self._clean_tab_contents,
                self._created_tab_texfiles,
                self._modified_tabenv_template,
                "tab" # Label prefix for tables
            )

        # Adjust graphicspath for the modified file
        # Remove any existing \graphicspath lines first
        self._modified_content = regex.sub(self.GRAPHICSPATH_PATTERN, "", self._modified_content)

        # Add a new graphicspath pointing to the temp directory.
        # Pandoc runs relative to the modified tex file's directory,
        # so the path should be just the name of the temp directory.
        # Ensure it ends with a slash for LaTeX.
        temp_dir_name = self.temp_subtexfile_dir.name
        new_graphicspath_line = f"\\graphicspath{{{{{temp_dir_name}/}}}}"

        # Try to insert the new graphicspath after \documentclass or common packages
        insert_point_pattern = r"(\\documentclass.*?\}\s*)|(\\usepackage.*?\}\s*)"
        last_match_end = 0
        for match in regex.finditer(insert_point_pattern, self._modified_content, regex.DOTALL):
            last_match_end = match.end()

        if last_match_end > 0:
            # Insert after the last match of documentclass or usepackage
            insert_pos = last_match_end
            self._modified_content = (
                self._modified_content[:insert_pos] +
                new_graphicspath_line + "\n" +
                self._modified_content[insert_pos:]
            )
            self.logger.debug("Inserted graphicspath after preamble.")
        else:
            # Fallback: insert at the very beginning
            self._modified_content = new_graphicspath_line + "\n" + self._modified_content
            self.logger.debug("Inserted graphicspath at the beginning (fallback).")
        self.logger.debug(f"Set graphicspath in modified tex to: {new_graphicspath_line}")


        # Write the final modified content to the output .tex file
        try:
            with open(self.output_texfile, "w", encoding="utf-8") as f:
                f.write(self._modified_content)
            self.logger.info(f"Created modified tex file: {self.output_texfile.name}")
        except Exception as e:
            self.logger.error(f"Error writing modified tex file {self.output_texfile}: {e}")
            raise # Re-raise the exception

    def _convert_to_docx(self) -> None:
        """Converts the modified TeX file to DOCX using pandoc."""
        # --- Prerequisite Checks ---
        if shutil.which("pandoc") is None:
            self.logger.error("pandoc not found in PATH. Please install pandoc.")
            raise FileNotFoundError("pandoc is not installed or not in PATH.")
        # Check for pandoc-crossref only if needed (it's used in the command)
        if shutil.which("pandoc-crossref") is None:
            self.logger.warning(
                "pandoc-crossref not found in PATH. Cross-referencing "
                "(figures, tables, equations) might not work correctly. "
                "Please install pandoc-crossref."
            )
            # Decide whether to raise an error or just warn. Warning seems better.
            # raise FileNotFoundError("pandoc-crossref is not installed or not in PATH.")

        if not self.output_texfile.exists():
             self.logger.error(f"Modified tex file not found: {self.output_texfile}. Cannot convert.")
             raise FileNotFoundError(f"Modified tex file not found: {self.output_texfile}")
        if not self.reference_docfile.exists():
             self.logger.error(f"Reference docx file not found: {self.reference_docfile}. Cannot convert.")
             raise FileNotFoundError(f"Reference docx file not found: {self.reference_docfile}")
        if not self.luafile.exists():
             self.logger.error(f"Lua filter file not found: {self.luafile}. Cannot convert.")
             raise FileNotFoundError(f"Lua filter file not found: {self.luafile}")


        # --- Construct Pandoc Command ---
        command = [
            "pandoc",
            # Input file (relative path from CWD)
            str(self.output_texfile.name),
            # Output file (relative path from CWD)
            "-o", str(self.output_docxfile.name),
            # Lua filter for equation labels (use absolute path)
            "--lua-filter", str(self.luafile.resolve()),
            # Pandoc-crossref filter (relies on it being in PATH)
            "--filter", "pandoc-crossref",
            # Reference document for styling (use absolute path)
            "--reference-doc", str(self.reference_docfile.resolve()),
            # Enable section numbering
            "--number-sections",
            # Metadata for pandoc-crossref (equation numbering)
            "-M", "autoEqnLabels",
            "-M", "tableEqns", # Enable table equation numbering if used
            # Output format with native numbering for figures/tables
            "-t", "docx+native_numbering",
        ]

        # Add citation options only if bibfile exists and is valid
        if self.bibfile and self.bibfile.exists() and self.bibfile.is_file():
            if not self.cslfile.exists() or not self.cslfile.is_file():
                 self.logger.error(f"CSL file not found or invalid: {self.cslfile}. Cannot process citations.")
                 raise FileNotFoundError(f"CSL file not found or invalid: {self.cslfile}")

            command.extend([
                # Title for the bibliography section
                "-M", "reference-section-title=References",
                # Bibliography file (use absolute path)
                "--bibliography", str(self.bibfile.resolve()),
                # Enable citation processing using citeproc
                "--citeproc",
                # CSL style file (use absolute path)
                "--csl", str(self.cslfile.resolve()),
            ])
        elif self.bibfile:
            # Log a warning if bibfile was specified but not found/valid
            self.logger.warning(
                f"Specified bibfile not found or invalid: {self.bibfile}. "
                "Skipping citation processing."
            )

        self.logger.debug(f"Pandoc command: {' '.join(command)}")

        # --- Run Pandoc ---
        # Execute pandoc in the directory containing the modified tex file
        # and the temp subfile directory (where the images are).
        cwd = self.output_texfile.parent
        try:
            result = subprocess.run(
                command,
                check=True, # Raise exception on non-zero exit code
                cwd=cwd, # Run pandoc where the modified tex file is located
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='replace' # Handle potential encoding errors in output
                )
            self.logger.info(
                f"Successfully converted {self.output_texfile.name} to "
                f"{self.output_docxfile.name}."
            )
            # Pandoc often prints warnings or info messages to stderr
            if result.stderr:
                 self.logger.warning(f"Pandoc stderr output:\n{result.stderr}")
            if result.stdout: # Log stdout as well, might contain info
                 self.logger.debug(f"Pandoc stdout output:\n{result.stdout}")

        except subprocess.CalledProcessError as e:
            self.logger.error(f"Pandoc conversion failed. Return code: {e.returncode}")
            # Log the stdout and stderr from the failed pandoc command
            self.logger.error(f"Pandoc stdout:\n{e.stdout}")
            self.logger.error(f"Pandoc stderr:\n{e.stderr}")
            raise # Re-raise the exception to signal failure
        except Exception as e:
            # Catch other potential errors during subprocess execution
            self.logger.error(f"An unexpected error occurred during pandoc conversion: {e}")
            raise

    def _clean_temp_files(self) -> None:
        """Cleans up temporary files and directories."""
        # Remove the temporary directory containing subfiles and images
        if self.temp_subtexfile_dir.exists():
            try:
                shutil.rmtree(self.temp_subtexfile_dir)
                self.logger.info(f"Removed temporary directory: {self.temp_subtexfile_dir}")
            except Exception as e:
                self.logger.error(
                    f"Error removing temporary directory "
                    f"{self.temp_subtexfile_dir}: {e}"
                )
        else:
            self.logger.debug("Temporary directory not found, skipping removal.")

        # Remove the intermediate modified tex file
        if self.output_texfile.exists():
            try:
                self.output_texfile.unlink()
                self.logger.info(f"Removed modified tex file: {self.output_texfile.name}")
            except Exception as e:
                self.logger.error(
                    f"Error removing modified tex file "
                    f"{self.output_texfile}: {e}"
                )
        else:
            self.logger.debug("Modified tex file not found, skipping removal.")

    def convert(self) -> None:
        """
        Executes the full LaTeX to Word conversion workflow.
        """
        try:
            # Step 1: Read and analyze the input LaTeX file
            self._read_and_preprocess_tex()
            self._analyze_tex_structure()

            # Step 2: Prepare the temporary directory for subfiles
            self._prepare_temp_directory()

            # Step 3: Create and compile subfiles for figures (and tables if enabled)
            self._create_figure_subfiles()
            if self.fix_table:
                self._create_table_subfiles()
            # Compile all created subfiles (figures and tables) into PNGs
            self._compile_all_subfiles()

            # Step 4: Create the modified main TeX file with images included
            self._create_modified_texfile()

            # Step 5: Convert the modified TeX file to DOCX using Pandoc
            self._convert_to_docx()

            self.logger.info("Conversion process completed successfully.")

        except FileNotFoundError as e:
             # Specific handling for missing files/executables
             self.logger.error(f"Conversion failed: {e}")
             # No cleanup here, user might want to inspect temp files
        except subprocess.CalledProcessError as e:
             # Specific handling for failed external commands (xelatex, pandoc)
             self.logger.error(f"Conversion failed due to external command error: {e}")
             # No cleanup here
        except Exception as e:
            # Catch any other unexpected errors during the process
            self.logger.error(f"Conversion failed due to an unexpected error: {e}", exc_info=True)
            # Re-raise the exception if the caller needs to handle it
            # raise
        finally:
            # Step 6: Cleanup temporary files (unless in debug mode)
            if self.logger.level != logging.DEBUG:
                self._clean_temp_files()
            else:
                self.logger.info("Debug mode enabled, skipping cleanup of temporary files.")
                self.logger.info(f"Temporary files are in: {self.temp_subtexfile_dir}")
                self.logger.info(f"Modified tex file: {self.output_texfile}")
