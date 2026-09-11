# Changelog

All notable changes to QuickPhrase are recorded here. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- QuickPhrase ribbon tab with a Favorites group of 12 phrase buttons.
- All Phrases dynamic menu, holding the full list with no size limit.
- Manage Phrases dialog: add, edit, reorder and delete, with edits applied to
  the ribbon immediately.
- Phrase storage as UTF-8 JSON (no BOM) at `%APPDATA%\QuickPhrase\snippets.json`.
- Import and export, with a choice of replacing or appending on import.
- Reload button, to pick up edits made to the phrase file outside Word.
- Insertion grouped into a single undo step on Word 2010 and later.
- `build/build.ps1`, building the `.dotm` from text sources on Windows.
- `install/Install.bat` and `Uninstall.bat`, clearing mark-of-the-web on install.
- Persian and English starter phrase packs in `examples/`.

[Unreleased]: https://github.com/cenaav/QuickPhrase/commits/dev
