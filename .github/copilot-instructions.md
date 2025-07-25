# GitHub Copilot Custom Instructions for Latex2word Python Project

This project aims to build a Python script tool that converts LaTeX documents to Word documents. The core functionality utilizes Pandoc and Pandoc-Crossref tools to automatically convert LaTeX files to Word files in specified formats, including the conversion and automatic numbering of formulas, images, tables, and references.

## Dependency Management
- This Python project uses **uv** to manage dependencies.
- Dependencies are recorded in `pyproject.toml`, not `requirements.txt`.
- Always follow this when generating or modifying dependency code.

## Testing
- Use **pytest** as the testing framework.
- Write clear and maintainable test functions.
- Use fixtures for setup and teardown.

## Commit Message Conventions
- Use conventional commit prefixes:
  - `feat:` for features
  - `fix:` for bug fixes
  - `docs:` for documentation
  - `style:` for formatting and whitespace only
  - `refactor:` for code restructuring without behavior change
  - `test:` for tests
  - `chore:` for miscellaneous tasks
- Format: `<type>(optional scope): <description>`
- Write commit messages in English.

## Python Code Style
- Follow PEP 8 guidelines.
- Use 4 spaces for indentation; no tabs.
- Use snake_case for functions and variables.
- Use CamelCase for classes.
- Constants in UPPER_SNAKE_CASE.
- Max line length: 79 characters for code, 72 for comments.
- Two blank lines before top-level functions/classes.
- Use double quotes for strings.
- Add meaningful docstrings following PEP 257.
- Use type annotations.
- Use descriptive names.
- Keep functions focused and simple.

## Logging
- Use the standard `logging` module for all logging purposes.
- Configure loggers with appropriate log levels (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`).
- Avoid using print statements for logging.
- Include contextual information in log messages to aid debugging.
- Follow consistent log message formatting.

## Comment Style
- Comments must be complete sentences starting with a capital letter.
- Limit comment lines to 72 characters.
- Use block comments indented to code level, starting with `# `.
- Separate paragraphs in block comments with a single `#` line.
- Use inline comments sparingly, separated by two spaces, starting with `# `.
- Avoid obvious comments; explain *why* not *what*.
- Write comments in English.
- Keep comments up to date.
- **Ensure comments are updated promptly to reflect any code changes.**

