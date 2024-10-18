import typer
from tex2docx import LatexToWordConverter

app = typer.Typer()


@app.command()
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


if __name__ == "__main__":
    app()
