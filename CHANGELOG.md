# Changelog

All notable changes to QuickPhrase are recorded here. Format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Fixed
- Persian, Arabic and other non-Latin text in message boxes showed as question
  marks. VBA's `MsgBox` converts through the system ANSI codepage; dialogs now
  call the Unicode `MessageBoxW` API via the new `UnicodeUI` module.

### Added
- QuickPhrase ribbon tab with a Favorites group of 12 phrase buttons.
- All Phrases dynamic menu, holding the full list with no size limit.
- Manage Phrases dialog: add, edit, reorder and delete, with edits applied to
  the ribbon immediately.
- Phrase storage as UTF-8 JSON (no BOM) at `%APPDATA%\QuickPhrase\snippets.json`.
- Automatic spacing around inserted phrases, so clicking a button never runs the
  text into the previous word. Configurable as before, after, or off; defaults to
  before, and is suppressed where whitespace or a paragraph mark already sits.
- Per-phrase "start a new line after inserting" flag, stored as `newline` in
  `snippets.json`.
- Ribbon placement setting: the phrase buttons can live on their own tab, inside
  Word's built-in Home tab, or both. Applies immediately, no restart.
- Import and export, with a choice of replacing or appending on import.
- Reload button, to pick up edits made to the phrase file outside Word.
- Insertion grouped into a single undo step on Word 2010 and later.
- `build/build.ps1`, building the `.dotm` from text sources on Windows, with
  `-ExportVba` to emit the compiled VBA project for committing.
- `build/pack.py`, assembling the `.dotm` from `package/` on any OS with no Word
  installed, plus `explode` and `verify` subcommands.
- `build-and-push.cmd`, a double-clickable Windows script that pulls, rebuilds
  the VBA project with Word, restores the Trust Center setting, then commits and
  pushes only the compiled blob.
- `build/check_sources.py`, static checks on the links the build makes by name:
  ribbon callbacks, dialog control names, favourite count across both namespaces,
  and that a fresh install seeds exactly the two starter phrases.
- GitHub Actions: `ci.yml` validates and packs on every push; `release.yml`
  publishes `QuickPhrase.dotm` and an installer bundle on a version tag.
- `package/`, the template stored as its individual OOXML parts.
- `install/Install.bat` and `Uninstall.bat`, clearing mark-of-the-web on install.
- Persian and English starter phrase packs in `examples/`.

### Documentation
- README rewritten with complete English and Persian sections covering install,
  usage, spacing and line-break settings, the phrase file format, compatibility
  and troubleshooting.

### Changed
- Ships two starter phrases, `Hello` and `سلام`, instead of four.
- Manager dialog is shorter, and the inserted-text box is sized for the sentence
  or sign-off a phrase usually is rather than a full page.

[Unreleased]: https://github.com/cenaav/QuickPhrase/commits/dev
