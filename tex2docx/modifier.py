"""LaTeX content modification and reference updating."""

import regex
from pathlib import Path
from typing import Dict, List, Optional

from .config import ConversionConfig
from .constants import TexPatterns, TexTemplates
from .utils import PatternMatcher, TextProcessor


class ContentModifier:
    """Handles modification of LaTeX content and reference updates."""
    
    def __init__(self, config: ConversionConfig) -> None:
        """
        Initialize the content modifier.
        
        Args:
            config: Configuration object.
        """
        self.config = config
        self.logger = config.setup_logger()
        self.modified_content: Optional[str] = None
    
    def create_modified_content(
        self,
        clean_content: str,
        figure_contents: List[str],
        table_contents: List[str],
        figure_files: Dict[int, str],
        table_files: Dict[int, str]
    ) -> str:
        """
        Create the modified LaTeX content with replaced environments.
        
        Args:
            clean_content: The clean LaTeX content.
            figure_contents: List of figure environment strings.
            table_contents: List of table environment strings.
            figure_files: Dictionary mapping indices to figure filenames.
            table_files: Dictionary mapping indices to table filenames.
            
        Returns:
            Modified LaTeX content.
        """
        self.modified_content = clean_content
        
        # Replace figure environments
        self._replace_environments(
            figure_contents,
            figure_files,
            self.config.get_multifig_template(),
            "multifig"
        )
        
        # Replace table environments if enabled
        if self.config.fix_table:
            self._replace_environments(
                table_contents,
                table_files,
                TexTemplates.MODIFIED_TABENV,
                "tab"
            )
        
        # Update graphics path
        self._update_graphicspath()
        
        # Fix any broken \ref commands that may have been split across lines
        self._fix_broken_refs()
        
        return self.modified_content
    
    def write_modified_file(self) -> None:
        """Write the modified content to the output TeX file."""
        if not self.modified_content:
            raise ValueError("No modified content to write")
        
        try:
            with open(self.config.output_texfile, "w", encoding="utf-8") as f:
                f.write(self.modified_content)
            self.logger.info(f"Created modified TeX file: {self.config.output_texfile.name}")
        except Exception as e:
            self.logger.error(f"Error writing modified TeX file: {e}")
            raise
    
    def _replace_environments(
        self,
        original_contents: List[str],
        created_files: Dict[int, str],
        template: str,
        label_prefix: str
    ) -> None:
        """
        Replace original environments with image includes.
        
        Args:
            original_contents: List of original environment strings.
            created_files: Dictionary mapping indices to filenames.
            template: LaTeX template for new environment.
            label_prefix: Prefix for new labels.
        """
        processed_indices = set()
        
        for index, original_content in enumerate(original_contents):
            if index in processed_indices or index not in created_files:
                continue
            
            # Generate new environment
            new_env = self._create_new_environment(
                original_content, created_files[index], template, label_prefix, index
            )
            
            if new_env and original_content in self.modified_content:
                # Replace content
                self.modified_content = self.modified_content.replace(
                    original_content, new_env, 1
                )
                processed_indices.add(index)
                
                # Update references
                self._update_references(original_content, new_env)
            else:
                self.logger.warning(
                    f"Could not replace environment {index} ({label_prefix})"
                )
    
    def _create_new_environment(
        self,
        original_content: str,
        filename: str,
        template: str,
        label_prefix: str,
        index: int
    ) -> Optional[str]:
        """
        Create a new environment replacing the original.
        
        Args:
            original_content: Original environment content.
            filename: PNG filename to include.
            template: LaTeX template to use.
            label_prefix: Prefix for the new label.
            index: Environment index.
            
        Returns:
            New environment string or None if creation failed.
        """
        # Extract caption
        caption = PatternMatcher.match_pattern(
            TexPatterns.CAPTION, original_content, mode="last"
        ) or ""
        
        # Generate new label
        base_name = Path(filename).stem
        safe_label = TextProcessor.sanitize_filename(base_name)
        new_label = f"{label_prefix}:{safe_label}"
        
        # PNG filename (just the name, not the path)
        png_filename = f"{base_name}.png"
        
        self.logger.debug(
            f"Creating new environment {index} ({label_prefix}):\n"
            f"  Caption: {caption[:50]}...\n"
            f"  Label: {new_label}\n"
            f"  PNG: {png_filename}"
        )
        
        # Format template based on its structure
        if template == TexTemplates.MULTIFIG_FIGENV:
            # Template order: image_path, caption, label
            return template % (png_filename, caption, new_label)
        elif template == TexTemplates.MODIFIED_TABENV:
            # Template order: caption, label, image_path
            return template % (caption, new_label, png_filename)
        else:
            self.logger.error(f"Unknown template for {label_prefix}")
            return None
    
    def _update_references(self, original_content: str, new_env: str) -> None:
        """
        Update references in the modified content.
        
        Args:
            original_content: Original environment content.
            new_env: New environment content.
        """
        # Extract the new label from the new environment
        new_label_match = regex.search(TexPatterns.LABEL, new_env)
        if not new_label_match:
            return
        
        new_label = new_label_match.group(1)
        
        # Update subfigure references
        self._update_subfigure_references(original_content, new_label)
        
        # Update main figure/table reference
        self._update_main_reference(original_content, new_label)
    
    def _update_subfigure_references(self, original_content: str, new_label: str) -> None:
        """
        Update references to subfigures.
        
        Args:
            original_content: Original environment content.
            new_label: New base label for the environment.
        """
        # Find includegraphics commands and their associated labels
        includegraphics_matches = list(
            regex.finditer(TexPatterns.INCLUDEGRAPHICS, original_content)
        )
        
        for i, img_match in enumerate(includegraphics_matches):
            # Look for the next label after this includegraphics
            start_pos = img_match.end()
            end_pos = (
                includegraphics_matches[i + 1].start()
                if i + 1 < len(includegraphics_matches)
                else len(original_content)
            )
            
            search_area = original_content[start_pos:end_pos]
            label_match = regex.search(TexPatterns.LABEL, search_area)
            
            if label_match:
                # Check if caption appears before this label
                content_before_label = search_area[:label_match.start()]
                if not regex.search(TexPatterns.CAPTION, content_before_label):
                    # Update this subfigure reference
                    subfig_label = label_match.group(1)
                    subfig_char = chr(ord('a') + i)
                    
                    ref_pattern = r"\\ref\{" + regex.escape(subfig_label) + r"\}"
                    new_ref = f"\\\\ref{{{new_label}}}({subfig_char})"  # Double backslash for proper LaTeX escaping
                    
                    self.modified_content = regex.sub(
                        ref_pattern, new_ref, self.modified_content
                    )
                    
                    self.logger.debug(
                        f"Updated subfigure reference '{subfig_label}' -> '{new_label}({subfig_char})'"
                    )
    
    def _update_main_reference(self, original_content: str, new_label: str) -> None:
        """
        Update the main figure/table reference.
        
        Args:
            original_content: Original environment content.
            new_label: New label for the environment.
        """
        main_label_match = regex.search(TexPatterns.LABEL, original_content)
        if main_label_match:
            main_label = main_label_match.group(1)
            
            # Update references to this main label
            ref_pattern = r"\\ref\{" + regex.escape(main_label) + r"\}"
            new_ref = f"\\\\ref{{{new_label}}}"  # Double backslash for proper LaTeX escaping
            
            count_before = self.modified_content.count(f"\\ref{{{main_label}}}")
            self.modified_content = regex.sub(ref_pattern, new_ref, self.modified_content)
            
            if count_before > 0:
                self.logger.debug(
                    f"Updated main reference '{main_label}' -> '{new_label}'"
                )
    
    def _update_graphicspath(self) -> None:
        """Update the graphics path in the modified content."""
        # Remove existing graphicspath
        self.modified_content = regex.sub(
            TexPatterns.GRAPHICSPATH, "", self.modified_content
        )
        
        # Add new graphicspath pointing to temp directory
        temp_dir_name = self.config.temp_subtexfile_dir.name
        new_graphicspath = f"\\graphicspath{{{{{temp_dir_name}/}}}}"
        
        # Try to insert after documentclass or usepackage
        insert_pattern = r"(\\documentclass.*?\}\s*)|(\\usepackage.*?\}\s*)"
        last_match_end = 0
        
        for match in regex.finditer(insert_pattern, self.modified_content, regex.DOTALL):
            last_match_end = match.end()
        
        if last_match_end > 0:
            # Insert after last match
            self.modified_content = (
                self.modified_content[:last_match_end] +
                new_graphicspath + "\n" +
                self.modified_content[last_match_end:]
            )
            self.logger.debug("Inserted graphicspath after preamble")
        else:
            # Fallback: insert at beginning
            self.modified_content = new_graphicspath + "\n" + self.modified_content
            self.logger.debug("Inserted graphicspath at beginning (fallback)")
        
        self.logger.debug(f"Set graphicspath to: {new_graphicspath}")
    
    def _fix_broken_refs(self) -> None:
        """Fix broken \\ref commands that may have been split across lines."""
        import regex
        
        # Pattern to match broken \ref commands
        # This matches backslash followed by newline or carriage return followed by "ef{"
        broken_ref_pattern = r"\\[\r\n]+ef\{"
        
        # Count occurrences before fixing
        broken_count = len(regex.findall(broken_ref_pattern, self.modified_content))
        
        if broken_count > 0:
            self.logger.debug(f"Found {broken_count} broken \\ref commands")
            
            # Fix the broken references
            self.modified_content = regex.sub(broken_ref_pattern, r"\\ref{", self.modified_content)
            
            self.logger.debug(f"Fixed {broken_count} broken \\ref commands")
        
        # Also fix any other common broken LaTeX commands that might occur
        # Pattern for \label commands
        broken_label_pattern = r"\\[\r\n]+label\{"
        broken_label_count = len(regex.findall(broken_label_pattern, self.modified_content))
        
        if broken_label_count > 0:
            self.logger.debug(f"Found {broken_label_count} broken \\label commands")
            self.modified_content = regex.sub(broken_label_pattern, r"\\label{", self.modified_content)
            self.logger.debug(f"Fixed {broken_label_count} broken \\label commands")
