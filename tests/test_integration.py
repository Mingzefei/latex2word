"""Integration tests for the LatexToWordConverter class.

This module contains end-to-end integration tests that verify the complete
conversion workflow from LaTeX documents to Word documents. These tests
exercise the full functionality of the tex2docx package by running actual
conversions and checking that output files are generated successfully.

Test scenarios:
- Basic English document conversion
- English document with chapters
- English document with includes  
- Chinese document conversion

Note: These tests require manual verification of the generated Word documents
to ensure proper formatting, references, and content layout.
"""

import pytest  # Use pytest instead of unittest
from pathlib import Path  # Use pathlib for better path handling
from typing import Dict, Any

from tex2docx import LatexToWordConverter

# Get the directory where the test file is located using pathlib
TEST_DIR: Path = Path(__file__).parent.resolve()
# Get the project root directory (assuming tests is one level down)
PROJECT_ROOT: Path = TEST_DIR.parent


# Define fixtures for test configurations using pytest.fixture
# This replaces the setUp method from unittest.
@pytest.fixture(scope="module")
def en_config() -> Dict[str, Any]:
    """Provides configuration for the English basic test."""
    return {
        "input_texfile": TEST_DIR / "en/main.tex",
        "output_docxfile": TEST_DIR / "en/main.docx",
        # Assuming reference/csl files are in the project root
        "reference_docfile": PROJECT_ROOT / "my_temp.docx",
        "cslfile": PROJECT_ROOT / "ieee.csl",
        # Assuming bib file is in the tests directory
        "bibfile": TEST_DIR / "ref.bib",
        "debug": True,
    }


@pytest.fixture(scope="module")
def en_chapter_config() -> Dict[str, Any]:
    """Provides configuration for the English chapter test."""
    return {
        "input_texfile": TEST_DIR / "en_chapter/main.tex",
        "output_docxfile": TEST_DIR / "en_chapter/main.docx",
        "reference_docfile": PROJECT_ROOT / "my_temp.docx",
        "cslfile": PROJECT_ROOT / "ieee.csl",
        "bibfile": TEST_DIR / "ref.bib",
        "debug": True,
    }


@pytest.fixture(scope="module")
def en_include_config() -> Dict[str, Any]:
    """Provides configuration for the English include test."""
    return {
        "input_texfile": TEST_DIR / "en_include/main.tex",
        "output_docxfile": TEST_DIR / "en_include/main.docx",
        "cslfile": PROJECT_ROOT / "ieee.csl",
        "bibfile": TEST_DIR / "ref.bib",
        "debug": True,
    }


@pytest.fixture(scope="module")
def zh_config() -> Dict[str, Any]:
    """Provides configuration for the Chinese test."""
    return {
        "input_texfile": TEST_DIR / "zh/main.tex",
        "output_docxfile": TEST_DIR / "zh/main.docx",
        "bibfile": TEST_DIR / "ref.bib",
        "fix_table": True,
        "debug": False,
    }


# Test functions now accept fixtures as arguments
def test_convert_en(en_config: Dict[str, Any]) -> None:
    """Tests conversion for the basic English document."""
    # Ensure output directory exists, or handle potential errors
    output_path: Path = Path(en_config["output_docxfile"])
    output_path.parent.mkdir(parents=True, exist_ok=True)

    converter = LatexToWordConverter(**en_config)
    converter.convert()

    # Assert that the output file exists and is not empty.
    assert output_path.exists(), f"Output file not found: {output_path}"
    assert (
        output_path.stat().st_size > 0
    ), f"Output file is empty: {output_path}"
    # Remind the user to manually check the generated file.
    print(f"\n[Manual Check] Please verify: {output_path}")


def test_convert_en_chapter(en_chapter_config: Dict[str, Any]) -> None:
    """Tests conversion for the English document with chapters."""
    output_path: Path = Path(en_chapter_config["output_docxfile"])
    output_path.parent.mkdir(parents=True, exist_ok=True)

    converter = LatexToWordConverter(**en_chapter_config)
    converter.convert()

    # Assert that the output file exists and is not empty.
    assert output_path.exists(), f"Output file not found: {output_path}"
    assert (
        output_path.stat().st_size > 0
    ), f"Output file is empty: {output_path}"
    # Remind the user to manually check the generated file.
    print(f"\n[Manual Check] Please verify: {output_path}")


def test_convert_en_include(en_include_config: Dict[str, Any]) -> None:
    """Tests conversion for the English document with includes."""
    output_path: Path = Path(en_include_config["output_docxfile"])
    output_path.parent.mkdir(parents=True, exist_ok=True)

    converter = LatexToWordConverter(**en_include_config)
    converter.convert()

    # Assert that the output file exists and is not empty.
    assert output_path.exists(), f"Output file not found: {output_path}"
    assert (
        output_path.stat().st_size > 0
    ), f"Output file is empty: {output_path}"
    # Remind the user to manually check the generated file.
    print(f"\n[Manual Check] Please verify: {output_path}")


def test_convert_zh(zh_config: Dict[str, Any]) -> None:
    """Tests conversion for the Chinese document."""
    output_path: Path = Path(zh_config["output_docxfile"])
    output_path.parent.mkdir(parents=True, exist_ok=True)

    converter = LatexToWordConverter(**zh_config)
    converter.convert()

    # Assert that the output file exists and is not empty.
    assert output_path.exists(), f"Output file not found: {output_path}"
    assert (
        output_path.stat().st_size > 0
    ), f"Output file is empty: {output_path}"
    # Remind the user to manually check the generated file.
    print(f"\n[Manual Check] Please verify: {output_path}")

# The if __name__ == "__main__": block is removed as pytest handles test discovery.
