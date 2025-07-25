"""Configuration management for tex2docx."""

import logging
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Union

from .constants import TexTemplates


@dataclass
class ConversionConfig:
    """Configuration for LaTeX to Word conversion."""
    
    # Required parameters
    input_texfile: Union[str, Path]
    output_docxfile: Union[str, Path]
    
    # Optional file paths
    bibfile: Optional[Union[str, Path]] = None
    cslfile: Optional[Union[str, Path]] = None
    reference_docfile: Optional[Union[str, Path]] = None
    
    # Conversion options
    debug: bool = False
    fix_table: bool = True
    
    # Template customization
    multifig_figenv_template: Optional[str] = None
    
    # Derived paths (set during initialization)
    output_texfile: Optional[Path] = field(default=None, init=False)
    temp_subtexfile_dir: Optional[Path] = field(default=None, init=False)
    luafile: Optional[Path] = field(default=None, init=False)
    
    def __post_init__(self) -> None:
        """Initialize derived paths and validate configuration."""
        # Convert string paths to Path objects
        self.input_texfile = Path(self.input_texfile).resolve()
        self.output_docxfile = Path(self.output_docxfile).resolve()
        
        if self.bibfile:
            self.bibfile = Path(self.bibfile).resolve()
        if self.cslfile:
            self.cslfile = Path(self.cslfile).resolve()
        if self.reference_docfile:
            self.reference_docfile = Path(self.reference_docfile).resolve()
        
        # Set derived paths
        self.output_texfile = self.input_texfile.with_name(
            f"{self.input_texfile.stem}_modified.tex"
        )
        self.temp_subtexfile_dir = self.input_texfile.parent / "temp_subtexfile_dir"
        
        # Set default file paths
        self._set_default_paths()
        
        # Validate required files exist
        self._validate_input_files()
    
    def _set_default_paths(self) -> None:
        """Set default paths for optional files."""
        package_dir = Path(__file__).parent
        
        # Find bibfile if not specified
        if not self.bibfile:
            bib_files = list(self.input_texfile.parent.glob("*.bib"))
            if bib_files:
                self.bibfile = bib_files[0].resolve()
        
        # Set default CSL file
        if not self.cslfile:
            self.cslfile = (package_dir / "ieee.csl").resolve()
        
        # Set default reference document
        if not self.reference_docfile:
            self.reference_docfile = (package_dir / "default_temp.docx").resolve()
        
        # Set Lua filter path
        self.luafile = (package_dir / "resolve_equation_labels.lua").resolve()
    
    def _validate_input_files(self) -> None:
        """Validate that required input files exist."""
        if not self.input_texfile.exists():
            raise FileNotFoundError(f"Input TeX file not found: {self.input_texfile}")
        
        if self.bibfile and not self.bibfile.exists():
            logging.warning(f"Bibliography file not found: {self.bibfile}")
        
        if self.cslfile and not self.cslfile.exists():
            raise FileNotFoundError(f"CSL file not found: {self.cslfile}")
        
        if self.reference_docfile and not self.reference_docfile.exists():
            raise FileNotFoundError(f"Reference document not found: {self.reference_docfile}")
        
        if self.luafile and not self.luafile.exists():
            raise FileNotFoundError(f"Lua filter not found: {self.luafile}")
    
    def get_multifig_template(self) -> str:
        """Get the multi-figure template to use."""
        return self.multifig_figenv_template or TexTemplates.MULTIFIG_FIGENV
    
    def setup_logger(self) -> logging.Logger:
        """Set up and return a logger instance."""
        logger = logging.getLogger(f"tex2docx_{uuid.uuid4().hex[:6]}")
        
        # Avoid adding handlers multiple times
        if not logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter(
                "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
        
        logger.setLevel(logging.DEBUG if self.debug else logging.INFO)
        return logger
    
    def log_paths(self, logger: logging.Logger) -> None:
        """Log configuration paths for debugging."""
        logger.debug(f"Input TeX file: {self.input_texfile}")
        logger.debug(f"Output DOCX file: {self.output_docxfile}")
        logger.debug(f"Output TeX file: {self.output_texfile}")
        logger.debug(f"Temp directory: {self.temp_subtexfile_dir}")
        logger.debug(f"Bibliography file: {self.bibfile}")
        logger.debug(f"CSL file: {self.cslfile}")
        logger.debug(f"Reference document: {self.reference_docfile}")
        logger.debug(f"Lua filter: {self.luafile}")
        logger.debug(f"Fix tables: {self.fix_table}")
        logger.debug(f"Debug mode: {self.debug}")
