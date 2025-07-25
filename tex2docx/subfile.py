"""LaTeX subfile generation and compilation."""

import concurrent.futures
import os
import subprocess
import uuid
from pathlib import Path
from typing import Dict, List, Tuple

from tqdm import tqdm

from .config import ConversionConfig
from .constants import TexTemplates, CompilerOptions
from .utils import PatternMatcher, TextProcessor


class SubfileGenerator:
    """Handles generation of LaTeX subfiles for figures and tables."""
    
    def __init__(self, config: ConversionConfig, parser_results: dict) -> None:
        """
        Initialize the subfile generator.
        
        Args:
            config: Configuration object.
            parser_results: Results from LaTeX parsing.
        """
        self.config = config
        self.logger = config.setup_logger()
        self.parser_results = parser_results
        
        # Storage for created files
        self.created_figure_files: Dict[int, str] = {}
        self.created_table_files: Dict[int, str] = {}
    
    def generate_figure_subfiles(self, figure_contents: List[str]) -> None:
        """
        Generate subfiles for figure environments.
        
        Args:
            figure_contents: List of figure environment strings.
        """
        self._generate_subfiles(
            figure_contents, 
            "multifig", 
            self.created_figure_files
        )
    
    def generate_table_subfiles(self, table_contents: List[str]) -> None:
        """
        Generate subfiles for table environments.
        
        Args:
            table_contents: List of table environment strings.
        """
        self._generate_subfiles(
            table_contents,
            "tab",
            self.created_table_files
        )
    
    def _generate_subfiles(
        self, 
        content_list: List[str], 
        prefix: str, 
        storage_dict: Dict[int, str]
    ) -> None:
        """
        Generate subfiles for a list of LaTeX environments.
        
        Args:
            content_list: List of environment content strings.
            prefix: Filename prefix.
            storage_dict: Dictionary to store created filenames.
        """
        if not content_list:
            self.logger.info(f"No {prefix} environments to process")
            return
        
        default_counter = 0
        created_filenames = set(storage_dict.values())
        
        # Calculate relative graphics path
        graphicspath_rel = self._get_relative_graphicspath()
        
        for index, item_content in enumerate(content_list):
            filename = self._generate_filename(
                item_content, prefix, default_counter, created_filenames
            )
            
            storage_dict[index] = filename
            created_filenames.add(filename)
            default_counter += 1
            
            # Generate and write file content
            self._write_subfile(item_content, filename, graphicspath_rel)
    
    def _get_relative_graphicspath(self) -> Path:
        """Calculate relative graphics path from temp dir to original path."""
        try:
            graphicspath = Path(self.parser_results.get("graphicspath", "."))
            rel_path_str = os.path.relpath(graphicspath, self.config.temp_subtexfile_dir)
            return Path(rel_path_str)
        except ValueError:
            # Different drives on Windows
            self.logger.warning(
                "Graphics path and temp directory on different drives. "
                "Using absolute graphics path."
            )
            return graphicspath.resolve()
    
    def _generate_filename(
        self, 
        content: str, 
        prefix: str, 
        counter: int, 
        existing_names: set
    ) -> str:
        """
        Generate a unique filename for a subfile.
        
        Args:
            content: The LaTeX content.
            prefix: Filename prefix.
            counter: Default counter for fallback naming.
            existing_names: Set of already used filenames.
            
        Returns:
            A unique filename.
        """
        # Extract label for filename
        labels = PatternMatcher.match_pattern(
            r"\\label\{(.*?)\}", content, mode="all"
        )
        
        if labels:
            base_name = labels[-1]
            # Clean common prefixes
            for pfx in ["fig:", "fig-", "fig_", "tab:", "tab-", "tab_"]:
                if base_name.startswith(pfx):
                    base_name = base_name[len(pfx):]
                    break
        else:
            base_name = f"{prefix}{counter}"
        
        # Sanitize filename
        safe_name = TextProcessor.sanitize_filename(base_name)
        filename = f"{prefix}_{safe_name}.tex"
        
        # Ensure uniqueness
        original_stem = f"{prefix}_{safe_name}"
        while filename in existing_names:
            unique_suffix = f"_{uuid.uuid4().hex[:4]}"
            filename = f"{original_stem}{unique_suffix}.tex"
        
        return filename
    
    def _write_subfile(
        self, 
        content: str, 
        filename: str, 
        graphicspath_rel: Path
    ) -> None:
        """
        Write a subfile to disk.
        
        Args:
            content: The LaTeX content.
            filename: The filename to write to.
            graphicspath_rel: Relative graphics path.
        """
        try:
            file_content = self._generate_file_content(content, graphicspath_rel)
            file_path = self.config.temp_subtexfile_dir / filename
            
            with open(file_path, "w", encoding="utf-8") as file:
                file.write(file_content)
            
            self.logger.info(f"Created subfile: {filename}")
        except Exception as e:
            self.logger.error(f"Error writing subfile {filename}: {e}")
            # Remove from storage if creation failed
            for index, stored_filename in list(self.created_figure_files.items()):
                if stored_filename == filename:
                    del self.created_figure_files[index]
            for index, stored_filename in list(self.created_table_files.items()):
                if stored_filename == filename:
                    del self.created_table_files[index]
    
    def _generate_file_content(self, content: str, graphicspath_rel: Path) -> str:
        """
        Generate the complete LaTeX file content for a subfile.
        
        Args:
            content: The original environment content.
            graphicspath_rel: Relative graphics path.
            
        Returns:
            Complete LaTeX file content.
        """
        # Process content
        processed_content = TextProcessor.comment_out_captions(content)
        processed_content = TextProcessor.remove_continued_float(processed_content)
        
        # Start with base template
        file_content = TexTemplates.BASE_MULTIFIG_TEXFILE
        
        # Set figure package
        figure_package = self.parser_results.get("figure_package")
        if figure_package == "subfig":
            package_lines = "\\usepackage{caption}\n\\usepackage{subfig}"
        elif figure_package == "subfigure":
            package_lines = f"\\usepackage{{{figure_package}}}"
        else:
            package_lines = ""
        
        if package_lines:
            file_content = file_content.replace(
                "% FIGURE_PACKAGE_PLACEHOLDER %", package_lines
            )
        else:
            file_content = file_content.replace("% FIGURE_PACKAGE_PLACEHOLDER %\n", "")
        
        # Set CJK package if needed
        if self.parser_results.get("contains_chinese", False):
            cjk_line = "\\usepackage{xeCJK}"
            file_content = file_content.replace("% CJK_PACKAGE_PLACEHOLDER %", cjk_line)
        else:
            file_content = file_content.replace("% CJK_PACKAGE_PLACEHOLDER %\n", "")
        
        # Set graphics path
        latex_graphicspath = str(graphicspath_rel.as_posix())
        file_content = file_content.replace(
            "{GRAPHICSPATH_PLACEHOLDER}", latex_graphicspath
        )
        
        # Insert content
        file_content = file_content.replace(
            "{FIGURE_CONTENT_PLACEHOLDER}", processed_content
        )
        
        return file_content


class SubfileCompiler:
    """Handles compilation of LaTeX subfiles to PNG images."""
    
    def __init__(self, config: ConversionConfig) -> None:
        """
        Initialize the compiler.
        
        Args:
            config: Configuration object.
        """
        self.config = config
        self.logger = config.setup_logger()
    
    def compile_all_subfiles(
        self, 
        figure_files: Dict[int, str], 
        table_files: Dict[int, str]
    ) -> None:
        """
        Compile all subfiles to PNG images in parallel.
        
        Args:
            figure_files: Dictionary of figure subfiles.
            table_files: Dictionary of table subfiles.
        """
        all_files = list(figure_files.values())
        if self.config.fix_table:
            all_files.extend(list(table_files.values()))
        
        if not all_files:
            self.logger.info("No subfiles to compile")
            return
        
        # Get full paths
        full_paths = [self.config.temp_subtexfile_dir / fname for fname in all_files]
        
        # Compile in parallel
        successful, failed = self._compile_parallel(full_paths)
        
        self.logger.info(
            f"Compilation finished. Success: {successful}, Failed: {len(failed)}"
        )
        if failed:
            self.logger.warning(f"Failed compilations: {', '.join(failed)}")
    
    def _compile_parallel(self, file_paths: List[Path]) -> Tuple[int, List[str]]:
        """
        Compile files in parallel using ProcessPoolExecutor.
        
        Args:
            file_paths: List of TeX file paths to compile.
            
        Returns:
            Tuple of (successful_count, failed_filenames).
        """
        successful_count = 0
        failed_files = []
        
        max_workers = min(CompilerOptions.MAX_WORKERS, len(file_paths))
        
        with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self._compile_single_file, path): path
                for path in file_paths
            }
            
            progress = tqdm(
                concurrent.futures.as_completed(futures),
                total=len(futures),
                desc="Compiling subfiles",
                unit="file"
            )
            
            for future in progress:
                file_path = futures[future]
                try:
                    success = future.result()
                    if success:
                        successful_count += 1
                    else:
                        failed_files.append(file_path.name)
                except Exception as e:
                    self.logger.error(f"Compilation exception for {file_path.name}: {e}")
                    failed_files.append(file_path.name)
        
        return successful_count, failed_files
    
    @staticmethod
    def _compile_single_file(file_path: Path) -> bool:
        """
        Compile a single TeX file to PNG.
        
        Args:
            file_path: Path to the TeX file.
            
        Returns:
            True if compilation succeeded, False otherwise.
        """
        command = ["xelatex"] + CompilerOptions.XELATEX_OPTIONS + [file_path.name]
        
        try:
            result = subprocess.run(
                command,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=file_path.parent
            )
            
            # Write logs
            with open(file_path.with_suffix(".out"), "w", encoding="utf-8") as f:
                f.write(result.stdout)
            with open(file_path.with_suffix(".err"), "w", encoding="utf-8") as f:
                f.write(result.stderr)
            
            # Check for PNG file
            return SubfileCompiler._handle_png_output(file_path)
            
        except subprocess.CalledProcessError as e:
            # Log compilation failure
            SubfileCompiler._log_compilation_failure(file_path, e)
            return False
        except Exception:
            return False
    
    @staticmethod
    def _handle_png_output(tex_path: Path) -> bool:
        """
        Handle PNG file output from compilation.
        
        Args:
            tex_path: Path to the source TeX file.
            
        Returns:
            True if PNG was successfully created/renamed, False otherwise.
        """
        expected_png = tex_path.with_suffix(".png")
        
        if expected_png.exists():
            return True
        
        # Look for PNGs with page numbers
        pattern = f"{tex_path.stem}*.png"
        created_pngs = list(tex_path.parent.glob(pattern))
        
        if not created_pngs:
            return False
        
        try:
            # Rename first PNG to expected name
            if expected_png.exists():
                expected_png.unlink()
            created_pngs[0].rename(expected_png)
            return True
        except Exception:
            return False
    
    @staticmethod
    def _log_compilation_failure(file_path: Path, error: subprocess.CalledProcessError) -> None:
        """
        Log compilation failure details.
        
        Args:
            file_path: Path to the failed file.
            error: The subprocess error.
        """
        try:
            with open(file_path.with_suffix(".out"), "w", encoding="utf-8") as f:
                f.write(error.stdout or "")
            with open(file_path.with_suffix(".err"), "w", encoding="utf-8") as f:
                f.write(error.stderr or "")
        except Exception:
            pass  # Ignore logging failures
