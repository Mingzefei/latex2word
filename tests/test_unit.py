"""Unit tests for the tex2docx package.

This module contains unit tests for individual functions and classes
in the tex2docx package. These tests verify that each component works
correctly in isolation and provide good test coverage for the codebase.

Test categories:
- ConversionConfig: Configuration validation and setup
- PatternMatcher: Text pattern matching utilities  
- TextProcessor: Text processing and manipulation
- LatexParser: LaTeX document parsing functionality
- Constants: Template and pattern definitions
- Integration: Component interaction testing
- Performance: Performance and scalability testing
"""

import logging
import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

from tex2docx.config import ConversionConfig
from tex2docx.constants import TexPatterns, TexTemplates
from tex2docx.parser import LatexParser
from tex2docx.utils import PatternMatcher, TextProcessor


class TestConversionConfig:
    """Test the ConversionConfig class."""
    
    def test_config_initialization(self, tmp_path):
        """Test basic config initialization."""
        input_file = tmp_path / "test.tex"
        output_file = tmp_path / "test.docx"
        input_file.write_text("\\documentclass{article}\\begin{document}Test\\end{document}")
        
        config = ConversionConfig(
            input_texfile=input_file,
            output_docxfile=output_file,
            debug=True
        )
        
        assert config.input_texfile == input_file.resolve()
        assert config.output_docxfile == output_file.resolve()
        assert config.debug is True
        assert config.fix_table is True
        assert config.output_texfile is not None
        assert config.temp_subtexfile_dir is not None
    
    def test_config_missing_input_file(self, tmp_path):
        """Test config with missing input file."""
        input_file = tmp_path / "nonexistent.tex"
        output_file = tmp_path / "test.docx"
        
        with pytest.raises(FileNotFoundError, match="Input TeX file not found"):
            ConversionConfig(
                input_texfile=input_file,
                output_docxfile=output_file
            )
    
    def test_config_logger_setup(self, tmp_path):
        """Test logger setup."""
        input_file = tmp_path / "test.tex"
        output_file = tmp_path / "test.docx"
        input_file.write_text("\\documentclass{article}\\begin{document}Test\\end{document}")
        
        config = ConversionConfig(
            input_texfile=input_file,
            output_docxfile=output_file,
            debug=True
        )
        
        logger = config.setup_logger()
        assert logger.level == logging.DEBUG
        
        config.debug = False
        logger2 = config.setup_logger()
        assert logger2.level == logging.INFO


class TestPatternMatcher:
    """Test the PatternMatcher utility class."""
    
    def test_match_pattern_all(self):
        """Test pattern matching with 'all' mode."""
        content = "\\ref{fig1} and \\ref{fig2} and \\ref{fig3}"
        result = PatternMatcher.match_pattern(TexPatterns.REF, content, "all")
        assert result == ["fig1", "fig2", "fig3"]
    
    def test_match_pattern_first(self):
        """Test pattern matching with 'first' mode."""
        content = "\\ref{fig1} and \\ref{fig2} and \\ref{fig3}"
        result = PatternMatcher.match_pattern(TexPatterns.REF, content, "first")
        assert result == "fig1"
    
    def test_match_pattern_last(self):
        """Test pattern matching with 'last' mode."""
        content = "\\ref{fig1} and \\ref{fig2} and \\ref{fig3}"
        result = PatternMatcher.match_pattern(TexPatterns.REF, content, "last")
        assert result == "fig3"
    
    def test_match_pattern_none(self):
        """Test pattern matching with no matches."""
        content = "No references here"
        result = PatternMatcher.match_pattern(TexPatterns.REF, content, "first")
        assert result is None
    
    def test_match_pattern_invalid_mode(self):
        """Test pattern matching with invalid mode."""
        with pytest.raises(ValueError, match="mode must be"):
            PatternMatcher.match_pattern(TexPatterns.REF, "content", "invalid")
    
    def test_find_figure_package_subfig(self):
        """Test detection of subfig package."""
        content = "\\usepackage{subfig}\\subfloat{content}"
        result = PatternMatcher.find_figure_package(content)
        assert result == "subfig"
    
    def test_find_figure_package_subfigure(self):
        """Test detection of subfigure package."""
        content = "\\usepackage{subfigure}\\subfigure{content}"
        result = PatternMatcher.find_figure_package(content)
        assert result == "subfigure"
    
    def test_find_figure_package_none(self):
        """Test no figure package detection."""
        content = "\\usepackage{graphicx}"
        result = PatternMatcher.find_figure_package(content)
        assert result is None
    
    def test_has_chinese_characters(self):
        """Test Chinese character detection."""
        assert PatternMatcher.has_chinese_characters("这是中文")
        assert not PatternMatcher.has_chinese_characters("This is English")
        assert PatternMatcher.has_chinese_characters("Mixed 中文 content")
    
    def test_extract_graphicspath(self):
        """Test graphics path extraction."""
        content = "\\graphicspath{{figures/}}"
        result = PatternMatcher.extract_graphicspath(content)
        assert result == "figures/"
        
        content_multi = "\\graphicspath{{figures/}{images/}}"
        result_multi = PatternMatcher.extract_graphicspath(content_multi)
        assert result_multi == "figures/"


class TestTextProcessor:
    """Test the TextProcessor utility class."""
    
    def test_remove_comments(self):
        """Test comment removal."""
        content = "Line 1\n% This is a comment\nLine 2"
        result = TextProcessor.remove_comments(content)
        assert "% This is a comment" not in result
        assert "Line 1" in result
        assert "Line 2" in result
    
    def test_comment_out_captions(self):
        """Test caption commenting."""
        content = "\\caption{Test caption}"
        result = TextProcessor.comment_out_captions(content)
        assert result.startswith("% ")
    
    def test_remove_continued_float(self):
        """Test ContinuedFloat removal."""
        content = "\\ContinuedFloat\\caption{Test}"
        result = TextProcessor.remove_continued_float(content)
        assert "\\ContinuedFloat" not in result
        assert "\\caption{Test}" in result
    
    def test_sanitize_filename(self):
        """Test filename sanitization."""
        filename = "test:file*name?.tex"
        result = TextProcessor.sanitize_filename(filename)
        assert result == "test_file_name_.tex"


class TestLatexParser:
    """Test the LatexParser class."""
    
    @pytest.fixture
    def sample_tex_content(self):
        """Sample LaTeX content for testing."""
        return """
        \\documentclass{article}
        \\usepackage{graphicx}
        \\usepackage{subfig}
        \\graphicspath{{figures/}}
        \\begin{document}
        \\begin{figure}
            \\centering
            \\includegraphics{test.png}
            \\caption{Test figure}
            \\label{fig:test}
        \\end{figure}
        \\begin{table}
            \\centering
            \\caption{Test table}
            \\label{tab:test}
            \\begin{tabular}{cc}
                A & B \\\\
                C & D
            \\end{tabular}
        \\end{table}
        \\end{document}
        """
    
    @pytest.fixture
    def config_with_temp_file(self, tmp_path, sample_tex_content):
        """Create a config with a temporary TeX file."""
        input_file = tmp_path / "test.tex"
        output_file = tmp_path / "test.docx"
        input_file.write_text(sample_tex_content)
        
        return ConversionConfig(
            input_texfile=input_file,
            output_docxfile=output_file,
            debug=True
        )
    
    def test_parser_initialization(self, config_with_temp_file):
        """Test parser initialization."""
        parser = LatexParser(config_with_temp_file)
        assert parser.config == config_with_temp_file
        assert parser.raw_content is None
        assert parser.clean_content is None
        assert parser.figure_contents == []
        assert parser.table_contents == []
    
    def test_read_and_preprocess(self, config_with_temp_file):
        """Test reading and preprocessing."""
        parser = LatexParser(config_with_temp_file)
        parser.read_and_preprocess()
        
        assert parser.raw_content is not None
        assert parser.clean_content is not None
        assert "\\documentclass{article}" in parser.clean_content
    
    def test_analyze_structure(self, config_with_temp_file):
        """Test structure analysis."""
        parser = LatexParser(config_with_temp_file)
        parser.read_and_preprocess()
        parser.analyze_structure()
        
        assert len(parser.figure_contents) == 1
        assert len(parser.table_contents) == 1
        assert parser.figure_package == "subfig"
        assert parser.graphicspath is not None
        assert not parser.contains_chinese
    
    def test_get_analysis_summary(self, config_with_temp_file):
        """Test analysis summary."""
        parser = LatexParser(config_with_temp_file)
        parser.read_and_preprocess()
        parser.analyze_structure()
        
        summary = parser.get_analysis_summary()
        assert summary["num_figures"] == 1
        assert summary["num_tables"] == 1
        assert summary["figure_package"] == "subfig"
        assert summary["contains_chinese"] is False
        assert summary["has_clean_content"] is True
    
    def test_missing_file_error(self, tmp_path):
        """Test error when input file is missing."""
        input_file = tmp_path / "nonexistent.tex"
        output_file = tmp_path / "test.docx"
        
        with pytest.raises(FileNotFoundError):
            ConversionConfig(
                input_texfile=input_file,
                output_docxfile=output_file
            )


class TestConstants:
    """Test constants and templates."""
    
    def test_tex_patterns_defined(self):
        """Test that all required patterns are defined."""
        assert hasattr(TexPatterns, 'FIGURE')
        assert hasattr(TexPatterns, 'TABLE')
        assert hasattr(TexPatterns, 'CAPTION')
        assert hasattr(TexPatterns, 'LABEL')
        assert hasattr(TexPatterns, 'REF')
        assert hasattr(TexPatterns, 'GRAPHICSPATH')
        assert hasattr(TexPatterns, 'INCLUDEGRAPHICS')
        assert hasattr(TexPatterns, 'COMMENT')
        assert hasattr(TexPatterns, 'CHINESE_CHAR')
    
    def test_tex_templates_defined(self):
        """Test that all required templates are defined."""
        assert hasattr(TexTemplates, 'BASE_MULTIFIG_TEXFILE')
        assert hasattr(TexTemplates, 'MULTIFIG_FIGENV')
        assert hasattr(TexTemplates, 'MODIFIED_TABENV')
        
        # Check that templates contain expected placeholders
        assert "FIGURE_CONTENT_PLACEHOLDER" in TexTemplates.BASE_MULTIFIG_TEXFILE
        assert "GRAPHICSPATH_PLACEHOLDER" in TexTemplates.BASE_MULTIFIG_TEXFILE
        assert "%s" in TexTemplates.MULTIFIG_FIGENV
        assert "%s" in TexTemplates.MODIFIED_TABENV


class TestIntegration:
    """Integration tests for the complete workflow."""
    
    @pytest.fixture
    def minimal_tex_file(self, tmp_path):
        """Create a minimal TeX file for testing."""
        content = """
        \\documentclass{article}
        \\usepackage{graphicx}
        \\begin{document}
        Hello World!
        \\begin{figure}
            \\centering
            \\includegraphics[width=0.5\\textwidth]{example.png}
            \\caption{Example figure}
            \\label{fig:example}
        \\end{figure}
        See Figure \\ref{fig:example}.
        \\end{document}
        """
        tex_file = tmp_path / "minimal.tex"
        tex_file.write_text(content)
        return tex_file
    
    def test_config_creation_and_validation(self, minimal_tex_file, tmp_path):
        """Test configuration creation and validation."""
        output_file = tmp_path / "output.docx"
        
        # This should work without raising exceptions
        config = ConversionConfig(
            input_texfile=minimal_tex_file,
            output_docxfile=output_file,
            debug=True
        )
        
        assert config.input_texfile.exists()
        assert config.output_texfile is not None
        assert config.temp_subtexfile_dir is not None
    
    @patch('shutil.which')
    def test_dependency_checking(self, mock_which, minimal_tex_file, tmp_path):
        """Test dependency checking."""
        from tex2docx.converter import PandocConverter
        
        config = ConversionConfig(
            input_texfile=minimal_tex_file,
            output_docxfile=tmp_path / "output.docx",
            debug=True
        )
        
        converter = PandocConverter(config)
        
        # Test when pandoc is missing
        mock_which.return_value = None
        with pytest.raises(Exception):  # DependencyError
            converter._check_dependencies()
        
        # Test when pandoc is available
        mock_which.return_value = "/usr/bin/pandoc"
        # Should not raise exception
        converter._check_dependencies()


# Test fixtures for common use cases
@pytest.fixture
def temp_dir():
    """Provide a temporary directory."""
    with tempfile.TemporaryDirectory() as tmp:
        yield Path(tmp)


@pytest.fixture
def sample_config(temp_dir):
    """Provide a sample configuration for testing."""
    input_file = temp_dir / "test.tex"
    output_file = temp_dir / "test.docx"
    
    # Create a minimal valid TeX file
    input_file.write_text("""
    \\documentclass{article}
    \\begin{document}
    Test document
    \\end{document}
    """)
    
    return ConversionConfig(
        input_texfile=input_file,
        output_docxfile=output_file,
        debug=True
    )


# Performance and stress tests
class TestPerformance:
    """Performance and stress tests."""
    
    def test_large_document_parsing(self, temp_dir):
        """Test parsing of a large document."""
        # Create a document with many figures
        content_parts = ["\\documentclass{article}", "\\begin{document}"]
        
        for i in range(50):  # 50 figures
            content_parts.append(f"""
            \\begin{{figure}}
                \\centering
                \\includegraphics{{fig{i}.png}}
                \\caption{{Figure {i}}}
                \\label{{fig:test{i}}}
            \\end{{figure}}
            """)
        
        content_parts.append("\\end{document}")
        content = "\n".join(content_parts)
        
        input_file = temp_dir / "large.tex"
        output_file = temp_dir / "large.docx"
        input_file.write_text(content)
        
        config = ConversionConfig(
            input_texfile=input_file,
            output_docxfile=output_file,
            debug=True
        )
        
        parser = LatexParser(config)
        parser.read_and_preprocess()
        parser.analyze_structure()
        
        # Should find all 50 figures
        assert len(parser.figure_contents) == 50
        
        summary = parser.get_analysis_summary()
        assert summary["num_figures"] == 50


if __name__ == "__main__":
    pytest.main([__file__])
