"""Utilities for LaTeX pattern matching and text processing."""

import regex
from typing import Union, List, Optional

from .constants import TexPatterns


class PatternMatcher:
    """Utility class for LaTeX pattern matching."""
    
    @staticmethod
    def match_pattern(
        pattern: str, 
        content: str, 
        mode: str = "last"
    ) -> Union[str, List[str], None]:
        """
        Match a pattern in the given content and return the result.

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
    
    @staticmethod
    def find_figure_package(content: str) -> Optional[str]:
        """
        Determine which figure package is used in the LaTeX content.
        
        Args:
            content: The LaTeX content to analyze.
            
        Returns:
            'subfig', 'subfigure', or None if no specific package is detected.
        """
        if (regex.search(TexPatterns.SUBFIG_PACKAGE, content) or 
            regex.search(TexPatterns.SUBFIG_ENV, content)):
            return "subfig"
        elif (regex.search(TexPatterns.SUBFIGURE_PACKAGE, content) or 
              regex.search(TexPatterns.SUBFIGURE_ENV, content)):
            return "subfigure"
        else:
            return None
    
    @staticmethod
    def has_chinese_characters(content: str) -> bool:
        """
        Check if the content contains Chinese characters.
        
        Args:
            content: The content to check.
            
        Returns:
            True if Chinese characters are found, False otherwise.
        """
        return bool(regex.search(TexPatterns.CHINESE_CHAR, content))
    
    @staticmethod
    def extract_graphicspath(content: str) -> Optional[str]:
        """
        Extract the graphics path from LaTeX content.
        
        Args:
            content: The LaTeX content to search.
            
        Returns:
            The graphics path if found, None otherwise.
        """
        match = PatternMatcher.match_pattern(
            TexPatterns.GRAPHICSPATH, content, mode="last"
        )
        if match and isinstance(match, str):
            # Handle multiple paths if present, taking the first one
            return match.split('}{')[0]
        return None


class TextProcessor:
    """Utility class for text processing operations."""
    
    @staticmethod
    def remove_comments(content: str) -> str:
        """
        Remove LaTeX comments from content.
        
        Args:
            content: The LaTeX content.
            
        Returns:
            Content with comments removed.
        """
        return regex.sub(TexPatterns.COMMENT, "", content)
    
    @staticmethod
    def comment_out_captions(content: str) -> str:
        """
        Comment out captions in LaTeX content.
        
        Args:
            content: The LaTeX content.
            
        Returns:
            Content with captions commented out.
        """
        def comment_caption_match(match: regex.Match) -> str:
            # Ensure each line of the caption block starts with %
            caption_block = match.group(0).strip()
            return "\n".join("% " + line for line in caption_block.split("\n"))
        
        return regex.sub(TexPatterns.CAPTION, comment_caption_match, content)
    
    @staticmethod
    def remove_continued_float(content: str) -> str:
        """
        Remove \\ContinuedFloat commands that break standalone compilation.
        
        Args:
            content: The LaTeX content.
            
        Returns:
            Content with \\ContinuedFloat commands removed.
        """
        return regex.sub(TexPatterns.CONTINUED_FLOAT, "", content)
    
    @staticmethod
    def sanitize_filename(filename: str) -> str:
        """
        Sanitize a filename by replacing invalid characters.
        
        Args:
            filename: The filename to sanitize.
            
        Returns:
            A sanitized filename.
        """
        from .constants import FilenamePatterns
        return regex.sub(
            FilenamePatterns.INVALID_CHARS, 
            FilenamePatterns.REPLACEMENT_CHAR, 
            filename
        )
