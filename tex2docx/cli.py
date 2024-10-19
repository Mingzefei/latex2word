import typer
from tex2docx import LatexToWordConverter
import shutil

app = typer.Typer()


# Subcommand for conversion
@app.command("convert")
def convert(
    input_texfile: str = typer.Option(..., help="The path to the input LaTeX file."),
    output_docxfile: str = typer.Option(
        ..., help="The path to the output Word document."
    ),
    reference_docfile: str = typer.Option(
        None,
        help="The path to the reference Word document. Defaults to None (use the built-in default_temp.docx file).",
    ),
    bibfile: str = typer.Option(
        None,
        help="The path to the BibTeX file. Defaults to None (use the first .bib file found in the same directory as input_texfile).",
    ),
    cslfile: str = typer.Option(
        None,
        help="The path to the CSL file. Defaults to None (use the built-in ieee.csl file).",
    ),
    fix_table: bool = typer.Option(
        True, help="Whether to fix tables with png. Defaults to True."
    ),
    debug: bool = typer.Option(False, help="Enable debug mode. Defaults to False."),
):
    """Convert LaTeX to Word with the given options."""
    converter = LatexToWordConverter(
        input_texfile,
        output_docxfile,
        reference_docfile=reference_docfile,
        bibfile=bibfile,
        cslfile=cslfile,
        fix_table=fix_table,
        debug=debug,
    )
    converter.convert()


# Subcommand for downloading dependencies #TODO(Hua)
@app.command("init")
def download():
    """Download dependencies for the tex2docx tool."""

    # Check pandoc and pandoc-crossref
    if not shutil.which("pandoc"):
        typer.echo("Pandoc is not installed. Please install Pandoc first.")
        raise typer.Exit(code=1)
    if not shutil.which("pandoc-crossref"):
        typer.echo(
            "Pandoc-crossref is not installed. Please install Pandoc-crossref first."
        )
        raise typer.Exit(code=1)

    typer.echo("Downloading dependencies...")
    # Add code to download and install dependencies here
    # For example, you could call external package managers, etc.
    # Example placeholder:
    typer.echo("Dependencies installed successfully.")


if __name__ == "__main__":
    app()
