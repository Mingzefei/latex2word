# Tests Directory

This directory contains the test suite for the tex2docx package, organized into different categories of tests.

## Test Files

### `test_integration.py` - Integration Tests
End-to-end tests that verify the complete LaTeX-to-Word conversion workflow.

**Purpose:**
- Test the full conversion pipeline from LaTeX input to Word output
- Verify that all document types (English, Chinese, with/without chapters) work correctly
- Ensure output files are generated successfully

**Test Scenarios:**
- `test_convert_en` - Basic English document conversion
- `test_convert_en_chapter` - English document with chapter structure
- `test_convert_en_include` - English document with include files
- `test_convert_zh` - Chinese document conversion

**Manual Verification Required:**
These tests generate Word documents that require manual inspection to verify:
- Proper formatting and layout
- Correct reference numbering (tables, figures, equations)
- Bibliography formatting
- Image and table placement

### `test_unit.py` - Unit Tests
Component-level tests that verify individual functions and classes work correctly in isolation.

**Purpose:**
- Test individual modules and functions
- Provide comprehensive code coverage
- Catch regressions in specific components
- Validate edge cases and error handling

**Test Categories:**
- `TestConversionConfig` - Configuration validation and setup
- `TestPatternMatcher` - Text pattern matching utilities
- `TestTextProcessor` - Text processing and manipulation functions
- `TestLatexParser` - LaTeX document parsing functionality
- `TestConstants` - Template and pattern definitions
- `TestIntegration` - Component interaction testing
- `TestPerformance` - Performance and scalability testing

## Test Data

### Sample Documents
- `en/` - English test documents and expected outputs
- `en_chapter/` - English documents with chapter structure
- `en_include/` - English documents with include statements
- `zh/` - Chinese test documents and expected outputs

### Supporting Files
- `ref.bib` - Bibliography file for citation testing
- `example-figure-maker.tex` - LaTeX script for generating test figures

## Running Tests

```bash
# Run all tests
pytest tests/

# Run only integration tests
pytest tests/test_integration.py

# Run only unit tests  
pytest tests/test_unit.py

# Run with verbose output
pytest tests/ -v

# Run specific test
pytest tests/test_integration.py::test_convert_en -v
```

## Test Coverage

The test suite provides comprehensive coverage of:
- ✅ Configuration validation
- ✅ Text pattern matching
- ✅ LaTeX parsing and preprocessing
- ✅ Content modification and reference updating
- ✅ Template application
- ✅ File I/O operations
- ✅ Error handling and edge cases
- ✅ End-to-end conversion workflows

## Development Guidelines

When adding new features:

1. **Add unit tests** in `test_unit.py` for new functions/classes
2. **Add integration tests** in `test_integration.py` for new conversion scenarios
3. **Update test data** if new LaTeX constructs are supported
4. **Document manual verification steps** for visual output validation

When fixing bugs:

1. **Add regression tests** to prevent the bug from reoccurring
2. **Test both unit and integration levels** where applicable
3. **Verify existing tests still pass** after changes
