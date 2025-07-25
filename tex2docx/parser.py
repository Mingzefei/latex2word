"""LaTeX content parsing and preprocessing."""

import regex
from pathlib import Path
from typing import List, Optional, Set

from .config import ConversionConfig
from .constants import TexPatterns
from .exceptions import ParseError
from .utils import PatternMatcher, TextProcessor


class LatexParser:
    """Handles LaTeX content parsing and preprocessing."""
    
    def __init__(self, config: ConversionConfig) -> None:
        """
        Initialize the LaTeX parser.
        
        Args:
            config: Configuration object.
        """
        self.config = config
        self.logger = config.setup_logger()
        
        # Content state
        self.raw_content: Optional[str] = None
        self.clean_content: Optional[str] = None
        self.figure_contents: List[str] = []
        self.table_contents: List[str] = []
        self.graphicspath: Optional[Path] = None
        self.figure_package: Optional[str] = None
        self.contains_chinese: bool = False
    
    def read_and_preprocess(self) -> None:
        """Read the main TeX file and handle includes and comments."""
        try:
            with open(self.config.input_texfile, "r", encoding="utf-8") as file:
                self.raw_content = file.read()
            self.logger.info(f"Read {self.config.input_texfile.name}")
        except Exception as e:
            self.logger.error(f"Error reading input file {self.config.input_texfile}: {e}")
            raise ParseError(f"Could not read input file: {e}")
        
        # Remove comments first
        clean_content = TextProcessor.remove_comments(self.raw_content)
        self.logger.debug("Removed comments from main file")
        
        # Process includes iteratively
        clean_content = self._process_includes(clean_content)
        
        self.clean_content = clean_content
        self.logger.debug("Finished processing includes and comments")
    
    def _process_includes(self, content: str) -> str:
        """
        Process \\include directives iteratively to support nested includes.
        
        Args:
            content: The LaTeX content to process.
            
        Returns:
            Content with includes resolved.
        """
        while True:
            includes_found = regex.findall(TexPatterns.INCLUDE, content)
            if not includes_found:
                break  # No more includes found
            
            made_replacement = False
            processed_in_pass: Set[str] = set()
            
            for include_name in includes_found:
                include_directive = f"\\include{{{include_name}}}"
                
                # Skip if already processed in this pass
                if include_directive in processed_in_pass:
                    continue
                
                include_filename = self._get_include_filename(include_name)
                include_file_path = self.config.input_texfile.parent / include_filename
                
                if include_file_path.exists():
                    include_content = self._read_include_file(include_file_path)
                    if include_content is not None:
                        content = content.replace(include_directive, include_content, 1)
                        self.logger.debug(f"Included content from {include_filename}")
                        made_replacement = True
                        processed_in_pass.add(include_directive)
                else:
                    self.logger.warning(f"Include file not found: {include_file_path}")
                    content = content.replace(
                        include_directive, 
                        f"% Include file not found: {include_filename} %", 
                        1
                    )
                    processed_in_pass.add(include_directive)
            
            if not made_replacement:
                break  # Exit if no replacements were made
        
        return content
    
    @staticmethod
    def _get_include_filename(include_name: str) -> str:
        """Get the full filename for an include directive."""
        if not include_name.lower().endswith(".tex"):
            return f"{include_name}.tex"
        return include_name
    
    def _read_include_file(self, file_path: Path) -> Optional[str]:
        """
        Read and process an include file.
        
        Args:
            file_path: Path to the include file.
            
        Returns:
            Processed content or None if reading failed.
        """
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                include_content = file.read()
            # Remove comments from included file before inserting
            return TextProcessor.remove_comments(include_content)
        except Exception as e:
            self.logger.warning(f"Could not read include file {file_path}: {e}")
            return f"% Error including {file_path.name} %"
    
    def analyze_structure(self) -> None:
        """Analyze the cleaned content for figures, tables, packages, etc."""
        if not self.clean_content:
            raise ParseError("Clean content is not available for analysis")
        
        # Extract figures and tables
        figure_matches = PatternMatcher.match_pattern(
            TexPatterns.FIGURE, self.clean_content, mode="all"
        )
        self.figure_contents = figure_matches if isinstance(figure_matches, list) else []
        
        table_matches = PatternMatcher.match_pattern(
            TexPatterns.TABLE, self.clean_content, mode="all"
        )
        self.table_contents = table_matches if isinstance(table_matches, list) else []
        
        self.logger.info(f"Found {len(self.figure_contents)} figure environments")
        self.logger.info(f"Found {len(self.table_contents)} table environments")
        
        # Determine figure package
        self.figure_package = PatternMatcher.find_figure_package(self.clean_content)
        self.logger.debug(f"Detected figure package: {self.figure_package}")
        
        # Determine graphics path
        self._determine_graphicspath()
        
        # Check for Chinese characters
        combined_content = "".join(self.figure_contents + self.table_contents)
        self.contains_chinese = PatternMatcher.has_chinese_characters(combined_content)
        if self.contains_chinese:
            self.logger.debug("Detected Chinese characters in figures/tables")
    
    def _determine_graphicspath(self) -> None:
        """Determine the graphics path from the LaTeX content."""
        if self.clean_content:
            graphicspath_str = PatternMatcher.extract_graphicspath(self.clean_content)
            if graphicspath_str:
                # Resolve relative to input TeX file parent
                self.graphicspath = (self.config.input_texfile.parent / graphicspath_str).resolve()
            else:
                self.graphicspath = self.config.input_texfile.parent.resolve()
        else:
            self.graphicspath = self.config.input_texfile.parent.resolve()
        
        self.logger.debug(f"Determined graphics path: {self.graphicspath}")
    
    def get_analysis_summary(self) -> dict:
        """
        Get a summary of the parsed content analysis.
        
        Returns:
            Dictionary containing analysis results.
        """
        return {
            "num_figures": len(self.figure_contents),
            "num_tables": len(self.table_contents),
            "figure_package": self.figure_package,
            "contains_chinese": self.contains_chinese,
            "graphicspath": str(self.graphicspath) if self.graphicspath else None,
            "has_clean_content": self.clean_content is not None,
        }
