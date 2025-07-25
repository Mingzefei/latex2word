"""Pandoc conversion functionality."""

import shutil
import subprocess
from typing import List

from .config import ConversionConfig
from .constants import PandocOptions
from .exceptions import ConversionError, DependencyError


class PandocConverter:
    """Handles conversion from LaTeX to DOCX using Pandoc."""
    
    def __init__(self, config: ConversionConfig) -> None:
        """
        Initialize the Pandoc converter.
        
        Args:
            config: Configuration object.
        """
        self.config = config
        self.logger = config.setup_logger()
    
    def convert_to_docx(self) -> None:
        """Convert the modified TeX file to DOCX using Pandoc."""
        # Check dependencies
        self._check_dependencies()
        
        # Validate files
        self._validate_files()
        
        # Build and run command
        command = self._build_pandoc_command()
        self._run_pandoc(command)
    
    def _check_dependencies(self) -> None:
        """Check that required external dependencies are available."""
        if not shutil.which("pandoc"):
            raise DependencyError("pandoc not found in PATH. Please install pandoc.")
        
        if not shutil.which("pandoc-crossref"):
            self.logger.warning(
                "pandoc-crossref not found in PATH. Cross-referencing "
                "(figures, tables, equations) might not work correctly."
            )
    
    def _validate_files(self) -> None:
        """Validate that required files exist."""
        if not self.config.output_texfile.exists():
            raise FileNotFoundError(f"Modified TeX file not found: {self.config.output_texfile}")
        
        if not self.config.reference_docfile.exists():
            raise FileNotFoundError(f"Reference document not found: {self.config.reference_docfile}")
        
        if not self.config.luafile.exists():
            raise FileNotFoundError(f"Lua filter not found: {self.config.luafile}")
    
    def _build_pandoc_command(self) -> List[str]:
        """
        Build the Pandoc command line.
        
        Returns:
            List of command arguments.
        """
        command = [
            "pandoc",
            str(self.config.output_texfile.name),  # Input file (relative to CWD)
            "-o", str(self.config.output_docxfile.name),  # Output file
        ]
        
        # Add Lua filter
        command.extend([
            "--lua-filter", str(self.config.luafile.resolve())
        ])
        
        # Add filters
        command.extend(PandocOptions.FILTER_OPTIONS)
        
        # Add reference document
        command.extend([
            "--reference-doc", str(self.config.reference_docfile.resolve())
        ])
        
        # Add basic options
        command.extend(PandocOptions.BASIC_OPTIONS)
        
        # Add citation options if bibliography is available
        if self._should_add_citations():
            command.extend(self._get_citation_options())
        
        return command
    
    def _should_add_citations(self) -> bool:
        """Check if citation processing should be enabled."""
        return (
            self.config.bibfile is not None and
            self.config.bibfile.exists() and
            self.config.bibfile.is_file()
        )
    
    def _get_citation_options(self) -> List[str]:
        """
        Get citation-related command line options.
        
        Returns:
            List of citation options.
        """
        if not self.config.cslfile.exists() or not self.config.cslfile.is_file():
            raise FileNotFoundError(f"CSL file not found: {self.config.cslfile}")
        
        return PandocOptions.CITATION_OPTIONS + [
            "--bibliography", str(self.config.bibfile.resolve()),
            "--csl", str(self.config.cslfile.resolve()),
        ]
    
    def _run_pandoc(self, command: List[str]) -> None:
        """
        Execute the Pandoc command.
        
        Args:
            command: Command line arguments.
        """
        self.logger.debug(f"Pandoc command: {' '.join(command)}")
        
        # Execute in the directory containing the modified TeX file
        cwd = self.config.output_texfile.parent
        
        try:
            result = subprocess.run(
                command,
                check=True,
                cwd=cwd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace"
            )
            
            self.logger.info(
                f"Successfully converted {self.config.output_texfile.name} to "
                f"{self.config.output_docxfile.name}"
            )
            
            # Log warnings if any
            if result.stderr:
                self.logger.warning(f"Pandoc stderr:\n{result.stderr}")
            if result.stdout:
                self.logger.debug(f"Pandoc stdout:\n{result.stdout}")
                
        except subprocess.CalledProcessError as e:
            error_msg = (
                f"Pandoc conversion failed (return code: {e.returncode})\n"
                f"stdout: {e.stdout}\n"
                f"stderr: {e.stderr}"
            )
            self.logger.error(error_msg)
            raise ConversionError(f"Pandoc conversion failed: {e}")
        except Exception as e:
            self.logger.error(f"Unexpected error during Pandoc conversion: {e}")
            raise ConversionError(f"Unexpected conversion error: {e}")
