"""Main entry point for tex2docx package."""

if __name__ == "__main__":
    from tex2docx.cli import app
    app()
else:
    from .cli import app
