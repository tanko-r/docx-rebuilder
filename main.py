#!/usr/bin/env python3
"""
DOCX Rebuilder - Command-line interface for fixing Word document formatting.

This tool analyzes Word documents, detects formatting conventions,
and rebuilds them with clean, consistent OOXML structure.
"""

import sys
import zipfile
from pathlib import Path

import click

from docx_rebuilder import DocxRebuilder
from docx_rebuilder.manifest import build_manifest


@click.command()
@click.argument('input_file', type=click.Path(exists=True, path_type=Path))
@click.option(
    '-o', '--output',
    type=click.Path(path_type=Path),
    help='Output file path. Defaults to input_rebuilt.docx'
)
@click.option(
    '--no-normalize-numbering',
    is_flag=True,
    help='Skip numbering normalization'
)
@click.option(
    '--no-normalize-styles',
    is_flag=True,
    help='Skip style normalization'
)
@click.option(
    '--no-preserve-refs',
    is_flag=True,
    help='Skip cross-reference preservation'
)
@click.option(
    '--analyze-only',
    is_flag=True,
    help='Only analyze the document, do not rebuild'
)
@click.option(
    '-v', '--verbose',
    is_flag=True,
    help='Show detailed output'
)
def main(input_file: Path, output: Path, no_normalize_numbering: bool,
         no_normalize_styles: bool, no_preserve_refs: bool,
         analyze_only: bool, verbose: bool):
    """
    Rebuild a Word document with clean, consistent formatting.

    INPUT_FILE: Path to the input .docx file
    """
    click.echo(f"Processing: {input_file}")

    try:
        # Initialize rebuilder
        rebuilder = DocxRebuilder(input_file, output)

        if analyze_only:
            # Just parse and analyze
            rebuilder._parse_document()
            rebuilder._analyze_formatting()

            report = rebuilder.get_analysis_report()
            click.echo("\nDocument Analysis Report:")
            click.echo("-" * 40)

            for key, value in report.items():
                label = key.replace('_', ' ').title()
                click.echo(f"  {label}: {value}")

            # Show detected issues
            suggestions = rebuilder.analyzer.suggest_corrections()
            if suggestions:
                click.echo(f"\nDetected {len(suggestions)} potential issues:")
                for suggestion in suggestions[:10]:  # Show first 10
                    click.echo(f"  - {suggestion['message']}")
                if len(suggestions) > 10:
                    click.echo(f"  ... and {len(suggestions) - 10} more")

            return

        # Perform full rebuild
        rebuild_options = {
            'normalize_numbering': not no_normalize_numbering,
            'normalize_styles': not no_normalize_styles,
            'preserve_cross_refs': not no_preserve_refs,
        }

        output_path = rebuilder.rebuild(**rebuild_options)

        # Get report and suggestions for manifest
        report = rebuilder.get_analysis_report()
        suggestions = []
        if rebuilder.analyzer:
            suggestions = rebuilder.analyzer.suggest_corrections()

        # Print the full manifest
        manifest = build_manifest(
            input_path=input_file,
            output_path=output_path,
            report=report,
            suggestions=suggestions,
            options=rebuild_options,
        )
        click.echo("")
        click.echo(manifest)

    except FileNotFoundError as e:
        click.echo(f"Error: File not found - {e}", err=True)
        sys.exit(1)
    except zipfile.BadZipFile:
        click.echo(f"Error: Invalid DOCX file - {input_file}", err=True)
        sys.exit(1)
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        if verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


@click.command()
@click.argument('input_file', type=click.Path(exists=True, path_type=Path))
@click.option(
    '-o', '--output',
    type=click.Path(path_type=Path),
    help='Output directory for extracted XML files'
)
def extract(input_file: Path, output: Path):
    """
    Extract and display the XML structure of a DOCX file.

    Useful for debugging and understanding document structure.
    """
    import zipfile
    from lxml import etree

    output = output or input_file.with_suffix('')

    click.echo(f"Extracting: {input_file}")
    click.echo(f"Output directory: {output}")

    output.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(input_file, 'r') as zf:
        for name in zf.namelist():
            if name.endswith('.xml') or name.endswith('.rels'):
                content = zf.read(name)

                # Pretty print XML
                try:
                    tree = etree.fromstring(content)
                    content = etree.tostring(tree, pretty_print=True, encoding='unicode')
                except Exception:
                    content = content.decode('utf-8', errors='replace')

                # Create output path
                out_path = output / name
                out_path.parent.mkdir(parents=True, exist_ok=True)

                with open(out_path, 'w', encoding='utf-8') as f:
                    f.write(content)

                click.echo(f"  Extracted: {name}")

    click.echo("\nExtraction complete!")


@click.command()
@click.argument('input_file', type=click.Path(path_type=Path), required=False)
def gui(input_file: Path):
    """
    Launch the graphical user interface.

    Optionally pass a .docx file to open it directly.
    """
    from docx_rebuilder.gui import run_gui
    run_gui(str(input_file) if input_file else None)


@click.group()
def cli():
    """DOCX Rebuilder - Fix Word document formatting for lawyers."""
    pass


cli.add_command(main, name='rebuild')
cli.add_command(extract)
cli.add_command(gui)


if __name__ == '__main__':
    # If called directly, use the rebuild command
    if len(sys.argv) > 1 and sys.argv[1] not in ['rebuild', 'extract', 'gui', '--help', '-h']:
        sys.argv.insert(1, 'rebuild')
    cli()
