import typer
from tex2docx import LatexToWordConverter

app = typer.Typer()


@app.command()
def convert(
    input_texfile: str = typer.Option(..., help="The path to the input LaTeX file."),
    multifig_dir: str = typer.Option(
        ..., help="The directory for multi-figure LaTeX files."
    ),
    output_docxfile: str = typer.Option(
        ..., help="The path to the output Word document."
    ),
    reference_docfile: str = typer.Option(
        ..., help="The path to the reference Word document."
    ),
    bibfile: str = typer.Option(..., help="The path to the BibTeX file."),
    cslfile: str = typer.Option(..., help="The path to the CSL file."),
    debug: bool = typer.Option(False, help="Enable debug mode."),
):
    """Convert LaTeX to Word with the given options."""
    converter = LatexToWordConverter(
        input_texfile,
        multifig_dir,
        output_docxfile,
        reference_docfile,
        bibfile,
        cslfile,
        debug,
    )
    converter.convert()


if __name__ == "__main__":
    app()
