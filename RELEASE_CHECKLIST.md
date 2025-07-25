# Release Checklist for tex2docx v1.3.0

## Pre-Release Verification

### ✅ Code Quality
- [x] All tests pass (31/31)
- [x] Code is properly formatted
- [x] Type hints are in place
- [x] No critical bugs remain

### ✅ Documentation  
- [x] README.md updated with v1.3.0 changelog
- [x] README_zh.md updated with v1.3.0 changelog  
- [x] Test documentation is complete
- [x] API documentation is accurate

### ✅ Project Structure
- [x] Code properly refactored into 8 modules
- [x] Tests reorganized (test_unit.py, test_integration.py)
- [x] Dependencies properly configured in pyproject.toml
- [x] GitHub Actions workflows updated

### ✅ Critical Fixes Verified
- [x] LaTeX `\ref{}` line break issue fixed
- [x] CLI import errors resolved
- [x] Reference numbering works correctly

## Release Process

### 1. Final Preparation
```bash
# Run the preparation script
./scripts/prepare_release.sh
```

### 2. Create and Push Tag
```bash
git tag v1.3.0
git push origin v1.3.0
```

### 3. Create GitHub Release
1. Go to https://github.com/Mingzefei/latex2word/releases/new
2. Use tag: `v1.3.0`
3. Release title: `v1.3.0 - Major Refactoring Release`
4. Copy the v1.3.0 changelog from README.md
5. Publish release

### 4. Verify Automatic PyPI Deployment
- GitHub Action should automatically trigger
- Package should appear on PyPI within 10 minutes
- Verify installation: `pip install tex2docx==1.3.0`

## Post-Release Tasks

### 1. Update Documentation
- [ ] Update project badges if needed
- [ ] Announce on relevant channels
- [ ] Update any external documentation

### 2. Monitor Release
- [ ] Check PyPI package page
- [ ] Monitor for user feedback
- [ ] Watch for any critical issues

## Release Highlights for Announcement

**tex2docx v1.3.0** brings major improvements to code maintainability and reliability:

🏗️ **Complete Code Refactoring**
- Modular architecture with 8 specialized components
- Better separation of concerns
- Enhanced type safety and error handling

🧪 **Improved Testing**
- Clear separation of unit vs integration tests
- Comprehensive test documentation
- Better CI/CD pipeline

🐛 **Critical Bug Fixes**
- Fixed LaTeX reference line break issues
- Resolved CLI import problems
- Enhanced reference numbering accuracy

⚡ **Developer Experience**
- Cleaner project structure
- Better documentation
- Easier to contribute and maintain

All existing functionality is preserved while significantly improving code quality and maintainability.

---

**Breaking Changes:** None - this is a fully backward-compatible release.

**Migration:** No migration needed - all APIs remain the same.
