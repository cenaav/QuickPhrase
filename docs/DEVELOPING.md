# Developing QuickPhrase

## Why there is no .dotm in the repo

A `.dotm` is an OOXML zip containing `vbaProject.bin` — a compiled, undocumented
binary that only Word can write. So:

- Source lives in `src/` as plain text and is diffable.
- The `.dotm` is a build artifact, gitignored, attached to Releases.
- CI cannot build it. GitHub's Windows runners have no Word. Releases are cut
  by hand on a machine that does.

## Layout

```
src/
  customUI/
    customUI14.xml         Ribbon definition for Word 2010+
    customUI.xml           Same tab, Word 2007 namespace
  modules/
    JsonLite.bas           Minimal JSON reader/writer
    SnippetStore.bas       Phrase list + UTF-8 file I/O
    QuickPhraseRibbon.bas  Ribbon callbacks
    QuickPhraseMain.bas    Insert, manager, import/export, about
  forms/
    frmManager.code.vb     Manager dialog behaviour (controls are generated)

build/build.ps1            Source -> dist/QuickPhrase.dotm
install/                   End-user install and uninstall
examples/                  Importable phrase packs
```

### Why the form has no .frm/.frx

VBA exports a UserForm as a text `.frm` plus a binary `.frx`, and the `.frm`
references the `.frx` by offset — so importing the text half alone fails. Rather
than commit a binary blob that no diff can review, `build.ps1` creates the form
at build time: it adds the controls from the `$controls` table, then pastes
`frmManager.code.vb` into the form's code module.

Consequence: **control names appear in two places.** Rename a control in
`$controls` and you must rename it in `frmManager.code.vb` too. A mismatch is
not caught at build time — the form compiles and the handler simply never
fires.

## Building

Needs Windows and Word. One-time setup:

*Word → File → Options → Trust Center → Trust Center Settings → Macro Settings*
→ tick **Trust access to the VBA project object model**.

The build automates the VBA editor, which is exactly what that setting gates.
It is safe to untick afterwards; the add-in itself never needs it.

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1

# watch Word work, useful when a stage fails
powershell -ExecutionPolicy Bypass -File build\build.ps1 -ShowWord
```

Two stages:

1. **Word automation.** New template → import `.bas` modules → generate
   `frmManager` → save as `wdFormatXMLTemplateMacroEnabled` (15).
2. **Zip surgery.** Insert both customUI parts and rewrite `_rels/.rels` to
   reference them, preserving Word's own relationships. No Word needed.

## How the ribbon manages a dynamic list

Ribbon XML is read once at load and cannot be regenerated, which is awkward for
a user-editable list. QuickPhrase works around it two ways at once:

- **Favorites** ships a fixed pool of 12 buttons (`qpBtn01`…`qpBtn12`). Each
  asks `GetFavLabel` / `GetFavVisible` what to show, so `Invalidate` is enough
  to relabel and hide them. Real flat buttons, capped count.
- **All Phrases** is a `dynamicMenu`. `GetMenuContent` returns menu XML built at
  click time, so it has no limit — at the cost of one extra click.

To change the favorites count, edit `FAV_COUNT` in `QuickPhraseRibbon.bas`
**and** the button count in both customUI files. They must agree.

`GetMenuContent` picks its XML namespace from `Application.Version`: 2007 wants
the 2006 namespace, 2010+ wants the 2009 one. A mismatch yields an empty menu
with no error — worth remembering if the menu ever comes up blank.

## Editing VBA in the VBE instead

Faster for debugging than rebuilding each time:

1. Build and install as usual.
2. In Word press <kbd>Alt</kbd>+<kbd>F11</kbd>, find the `QuickPhrase` project.
3. Edit, then **export the changed module back over `src/modules/`** —
   right-click → Export File. Otherwise the next build overwrites your work.

Ribbon XML cannot be edited this way. Change the XML and rebuild.

## Gotchas worth knowing

**The `IRibbonUI` reference is fragile.** An unhandled VBA error resets the
project and drops it, after which `Invalidate` silently does nothing until Word
rebuilds the ribbon. `RefreshRibbon` swallows this on purpose; the **Reload**
button is the user-facing escape hatch.

**UTF-8 without a BOM needs the two-stream dance.** `ADODB.Stream` always
writes a BOM in text mode, so `SnippetStore.WriteUtf8` stages the text, flips
the stream to binary, seeks past the three BOM bytes and copies into a second
stream. Skipping this leaves a BOM that shows up as `ï»¿` in other editors.
Also note `Type` may only be changed while `Position` is 0.

**Line breaks change form three times.** `vbLf` in the file, `vbCrLf` in the
dialog's text box, `vbCr` in a Word range. Each conversion is explicit in the
code; dropping one produces stray boxes or lost line breaks.

**Buttons need no image.** Ribbon buttons at `size="normal"` render label-only
quite happily, which is why there are no icon assets to maintain.

## Releasing

1. Bump `QP_VERSION` in `QuickPhraseMain.bas` and add a `CHANGELOG.md` entry.
2. `build\build.ps1`
3. Install the result and smoke-test: insert a phrase, add/edit/delete/reorder
   one, import and export, restart Word and confirm they persisted.
4. Tag, then create the release and attach `dist\QuickPhrase.dotm`.
