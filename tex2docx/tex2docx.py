"""Refactored LaTeX to Word converter with improved modularity."""

import logging
from pathlib import Path
from typing import Union

from .config import ConversionConfig
from .converter import PandocConverter
from .exceptions import Tex2DocxError
from .file_manager import FileManager
from .modifier import ContentModifier
from .parser import LatexParser
from .subfile import SubfileGenerator, SubfileCompiler


class LatexToWordConverter:
    """
    Main class for converting LaTeX documents to Word documents.
    
    This is a refactored version with improved modularity, separation of concerns,
    and better error handling.
    """
    
    def __init__(
        self,
        input_texfile: Union[str, Path],
        output_docxfile: Union[str, Path],
        bibfile: Union[str, Path, None] = None,
        cslfile: Union[str, Path, None] = None,
        reference_docfile: Union[str, Path, None] = None,
        debug: bool = False,
        multifig_texfile_template: Union[str, None] = None,  # Deprecated
        multifig_figenv_template: Union[str, None] = None,
        fix_table: bool = True,
    ) -> None:
        """
        Initialize the LaTeX to Word converter.

        Args:
            input_texfile: Path to the input LaTeX file.
            output_docxfile: Path to the output Word document file.
            bibfile: Path to the BibTeX file. Defaults to None
                (use the first .bib file found in the same directory
                 as input_texfile).
            cslfile: Path to the CSL file. Defaults to None
                (use the built-in ieee.csl file).
            reference_docfile: Path to the reference Word document file.
                Defaults to None (use the built-in default_temp.docx file).
            debug: Whether to enable debug mode. Defaults to False.
            multifig_texfile_template: Deprecated parameter, ignored.
            multifig_figenv_template: Template for figure environments
                in multi-figure LaTeX files. Defaults to built-in template.
            fix_table: Whether to fix tables by converting them to images.
                Defaults to True.
        """
        # Issue deprecation warning for old parameter
        if multifig_texfile_template is not None:
            logging.warning(
                "multifig_texfile_template parameter is deprecated and ignored. "
                "Templates are now generated dynamically."
            )
        
        # Create configuration
        self.config = ConversionConfig(
            input_texfile=input_texfile,
            output_docxfile=output_docxfile,
            bibfile=bibfile,
            cslfile=cslfile,
            reference_docfile=reference_docfile,
            debug=debug,
            fix_table=fix_table,
            multifig_figenv_template=multifig_figenv_template,
        )
        
        # Set up logger
        self.logger = self.config.setup_logger()
        
        # Log initial configuration
        self.config.log_paths(self.logger)
        
        self.logger.debug("LatexToWordConverter initialized with modular architecture")
    
    def convert(self) -> None:
        """
        Execute the full LaTeX to Word conversion workflow.
        
        This method orchestrates the entire conversion process using
        specialized components for each stage.
        """
        file_manager = FileManager(self.config)
        
        try:
            self.logger.info("Starting LaTeX to Word conversion process")
            
            # Step 1: Parse and preprocess LaTeX content
            self.logger.info("Step 1: Parsing LaTeX content")
            parser = LatexParser(self.config)
            parser.read_and_preprocess()
            parser.analyze_structure()
            
            # Step 2: Prepare temporary directory
            self.logger.info("Step 2: Preparing temporary directory")
            file_manager.prepare_temp_directory()
            
            # Step 3: Generate and compile subfiles
            self.logger.info("Step 3: Generating and compiling subfiles")
            subfile_gen = SubfileGenerator(self.config, parser.get_analysis_summary())
            subfile_gen.generate_figure_subfiles(parser.figure_contents)
            
            if self.config.fix_table:
                subfile_gen.generate_table_subfiles(parser.table_contents)
            
            # Compile all subfiles
            compiler = SubfileCompiler(self.config)
            compiler.compile_all_subfiles(
                subfile_gen.created_figure_files,
                subfile_gen.created_table_files
            )
            
            # Step 4: Modify LaTeX content
            self.logger.info("Step 4: Creating modified LaTeX file")
            modifier = ContentModifier(self.config)
            modifier.create_modified_content(
                parser.clean_content,
                parser.figure_contents,
                parser.table_contents,
                subfile_gen.created_figure_files,
                subfile_gen.created_table_files
            )
            modifier.write_modified_file()
            
            # Step 5: Convert to DOCX
            self.logger.info("Step 5: Converting to DOCX using Pandoc")
            pandoc_converter = PandocConverter(self.config)
            pandoc_converter.convert_to_docx()
            
            self.logger.info("Conversion process completed successfully")
            
        except Tex2DocxError as e:
            # Handle known errors from our package
            self.logger.error(f"Conversion failed: {e}")
            # Don't cleanup on error so user can inspect files
        except Exception as e:
            # Handle unexpected errors
            self.logger.error(f"Conversion failed due to unexpected error: {e}", exc_info=True)
            # Don't cleanup on error
        finally:
            # Step 6: Cleanup (unless in debug mode or there was an error)
            if file_manager.should_cleanup():
                file_manager.cleanup_temp_files()
            else:
                file_manager.log_temp_file_locations()


# Maintain backward compatibility by exposing the original interface
__all__ = ["LatexToWordConverter"]
