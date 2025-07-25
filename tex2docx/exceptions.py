"""Exception classes for tex2docx package."""


class Tex2DocxError(Exception):
    """Base exception for tex2docx package."""
    pass


class FileNotFoundError(Tex2DocxError):
    """Raised when a required file is not found."""
    pass


class CompilationError(Tex2DocxError):
    """Raised when LaTeX compilation fails."""
    pass


class ConversionError(Tex2DocxError):
    """Raised when Pandoc conversion fails."""
    pass


class DependencyError(Tex2DocxError):
    """Raised when required external dependencies are missing."""
    pass


class ParseError(Tex2DocxError):
    """Raised when LaTeX parsing fails."""
    pass


class ConfigurationError(Tex2DocxError):
    """Raised when configuration is invalid."""
    pass
