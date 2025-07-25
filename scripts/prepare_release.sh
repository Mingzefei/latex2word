#!/bin/bash

# Release preparation script for tex2docx

set -e

echo "🚀 Preparing release for tex2docx..."

# Check if we're on the master branch
current_branch=$(git branch --show-current)
if [ "$current_branch" != "master" ]; then
    echo "❌ Error: Not on master branch. Current branch: $current_branch"
    exit 1
fi

# Check if working directory is clean
if [ -n "$(git status --porcelain)" ]; then
    echo "❌ Error: Working directory is not clean. Please commit or stash your changes."
    git status --short
    exit 1
fi

# Run tests
echo "🧪 Running tests..."
if ! .venv/bin/python -m pytest tests/ -v; then
    echo "❌ Tests failed. Please fix them before releasing."
    exit 1
fi

# Check code quality
echo "🔍 Checking code quality..."
if ! .venv/bin/python -m ruff check tex2docx/ tests/; then
    echo "❌ Linting failed. Please fix the issues before releasing."
    exit 1
fi

# Build package
echo "📦 Building package..."
if ! .venv/bin/python -m build; then
    echo "❌ Build failed."
    exit 1
fi

# Show current version
current_version=$(git describe --tags --abbrev=0 2>/dev/null || echo "No tags found")
echo "📋 Current version: $current_version"

# Suggest next version
echo "💡 Suggested next version: v1.3.0 (major refactoring release)"

echo ""
echo "✅ Release preparation complete!"
echo ""
echo "📝 Next steps:"
echo "1. Create and push a new tag: git tag v1.3.0 && git push origin v1.3.0"
echo "2. Create a GitHub release at: https://github.com/Mingzefei/latex2word/releases/new"
echo "3. Use the tag v1.3.0 and copy the changelog from README.md"
echo "4. The GitHub Action will automatically publish to PyPI when you publish the release"
echo ""
echo "🎉 Happy releasing!"
