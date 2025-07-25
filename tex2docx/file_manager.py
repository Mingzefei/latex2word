"""File management utilities for tex2docx."""

import shutil

from .config import ConversionConfig


class FileManager:
    """Handles file and directory management operations."""
    
    def __init__(self, config: ConversionConfig) -> None:
        """
        Initialize the file manager.
        
        Args:
            config: Configuration object.
        """
        self.config = config
        self.logger = config.setup_logger()
    
    def prepare_temp_directory(self) -> None:
        """Create or clean the temporary directory."""
        try:
            if self.config.temp_subtexfile_dir.exists():
                shutil.rmtree(self.config.temp_subtexfile_dir)
                self.logger.debug(f"Removed existing temp directory: {self.config.temp_subtexfile_dir}")
            
            self.config.temp_subtexfile_dir.mkdir(parents=True)
            self.logger.debug(f"Created temp directory: {self.config.temp_subtexfile_dir}")
        except OSError as e:
            self.logger.error(f"Error managing temp directory {self.config.temp_subtexfile_dir}: {e}")
            raise
    
    def cleanup_temp_files(self) -> None:
        """Clean up temporary files and directories."""
        # Remove temporary directory
        if self.config.temp_subtexfile_dir.exists():
            try:
                shutil.rmtree(self.config.temp_subtexfile_dir)
                self.logger.info(f"Removed temporary directory: {self.config.temp_subtexfile_dir}")
            except Exception as e:
                self.logger.error(f"Error removing temp directory: {e}")
        else:
            self.logger.debug("Temporary directory not found, skipping removal")
        
        # Remove modified TeX file
        if self.config.output_texfile.exists():
            try:
                self.config.output_texfile.unlink()
                self.logger.info(f"Removed modified TeX file: {self.config.output_texfile.name}")
            except Exception as e:
                self.logger.error(f"Error removing modified TeX file: {e}")
        else:
            self.logger.debug("Modified TeX file not found, skipping removal")
    
    def should_cleanup(self) -> bool:
        """
        Determine if cleanup should be performed based on debug mode.
        
        Returns:
            True if cleanup should be performed, False otherwise.
        """
        return not self.config.debug
    
    def log_temp_file_locations(self) -> None:
        """Log the locations of temporary files for debugging."""
        self.logger.info("Debug mode enabled, skipping cleanup of temporary files")
        self.logger.info(f"Temporary files are in: {self.config.temp_subtexfile_dir}")
        self.logger.info(f"Modified TeX file: {self.config.output_texfile}")
