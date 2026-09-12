# Developing QuickPhrase

## How the build is split, and why

A `.dotm` is an OOXML zip. Every part inside it is text — except
`word/vbaProject.bin`, a compiled VBA project that only Microsoft Word can
produce. That one binary shapes the whole pipeline.

So the repository stores the template **exploded**:

- `package/` — every text part of the `.dotm`, diffable in a pull request.
- `package/word/vbaProject.bin` — the compiled VBA project. Committed on
  purpose, despite being a binary.
- `src/` — the VBA source that blob is compiled *from*, plus the ribbon XML.

Two commands assemble it:

| Command | Needs | Produces |
|---|---|---|
| `build/pack.py pack` | Python 3, any OS | `dist/QuickPhrase.dotm` |
| `build/build.ps1 -ExportVba` | Windows + Word | `package/word/vbaProject.bin` |

The first runs on every push and on release. The second runs only when the VBA
source changes — and a human runs it, because no CI runner has Word.

**Committing a compiled binary is a deliberate trade.** It is the only way
GitHub Actions can publish releases; the alternative is a self-hosted Windows
runner with Word installed, which is worth revisiting if the VBA starts changing
often. The mitigation is that `src/` holds the readable source, so the blob is
reproducible and reviewable by rebuilding rather than by reading.

## Layout

```
src/
  customUI/
    customUI14.xml         Ribbon definition, Word 2010+ namespace
    customUI.xml           Same tab, Word 2007 namespace
  modules/
    JsonLite.bas           Minimal JSON reader/writer
    UnicodeUI.bas          MessageBoxW wrapper, for non-Latin text in dialogs
    SnippetStore.bas       Phrase list + UTF-8 file I/O
    QuickPhraseRibbon.bas  Ribbon callbacks
    QuickPhraseMain.bas    Insert, manager, import/export, about
  forms/
    frmManager.code.vb     Manager dialog behaviour (controls are generated)

package/                   The .dotm, exploded
  [Content_Types].xml
  _rels/.rels              Includes both customUI relationships
  word/
    document.xml           Intentionally empty body
    styles.xml             Minimal; must not restyle user documents
    settings.xml
    _rels/document.xml.rels
    vbaProject.bin         Compiled VBA (binary, committed)
  docProps/

build.cmd                  Rebuild the VBA project (Windows, needs Word)

build/
  pack.py                  package/ + src/customUI -> .dotm  (any OS)
  build.ps1                VBA source -> vbaProject.bin      (Windows + Word)
  check_sources.py         Static cross-file consistency checks

install/                   End-user install and uninstall
examples/                  Importable phrase packs
.github/workflows/         ci.yml (checks + pack), release.yml (publish)
```

Note `src/customUI/` is the single source of truth for the ribbon XML: `pack.py`
maps those files into the archive at `customUI/`, and `package/` deliberately
does not contain a copy. `pack.py explode` skips them for the same reason.

## Everyday workflow

Changing ribbon XML, docs, examples or workflows — **no Word needed**:

```bash
python3 build/check_sources.py
python3 build/pack.py pack
python3 build/pack.py verify
```

Changing VBA under `src/` — needs Windows and Word once. Easiest route is to
double-click:

```
build.cmd
```

It refuses to run while Word is open, rebuilds, and restores the Trust Center
setting even if the build fails. It makes no git changes at all: publishing the
new blob is a separate, deliberate step, so a build can be inspected or thrown
away without anything leaving the machine.

By hand:

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1 -ExportVba
git add package/word/vbaProject.bin
```

One-time Word setup for that step. Either let the script do it:

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1 -ExportVba -EnableVbomTrust

# once the build has succeeded
powershell -ExecutionPolicy Bypass -File build\build.ps1 -DisableVbomTrust
```

or tick it by hand in *Word → File → Options → Trust Center → Trust Center
Settings → Macro Settings → **Trust access to the VBA project object model***.

The build automates the VBA editor, which is precisely what that setting gates.
It applies to the current user only and needs no admin rights. Safe to turn back
off afterwards; the add-in itself never needs it.

Without it, Word raises no error - `$doc.VBProject` simply returns nothing, and
the failure surfaces later as a misleading "property 'Name' cannot be found".
So `build.ps1` checks
`HKCU:\Software\Microsoft\Office\<version>\Word\Security\AccessVBOM`
*before* launching Word, and reports whether the value is unset, 0, or already 1
— which distinguishes "never enabled" from "enabled, but a Word window is still
holding the old setting".

## If Word rejects the packed template

`package/` was hand-authored, and a newer Word may expect parts it does not
contain. Rather than guess, regenerate the parts from a template Word wrote
itself:

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1   # full Word build
python build\pack.py explode dist\QuickPhrase.dotm         # refresh package/
```

Then review the diff before committing — Word may have added parts, and some of
them (`docProps/app.xml` timestamps, for example) are noise.

## Why the form has no .frm/.frx

VBA exports a UserForm as a text `.frm` plus a binary `.frx`, and the `.frm`
references the `.frx` by byte offset — so importing the text half alone fails.
Rather than commit a second binary blob that no diff can review, `build.ps1`
creates the form at build time: it adds the controls listed in its `$controls`
table, then pastes `frmManager.code.vb` into the form's code module.

Consequence: **control names live in two files.** Rename a control in
`$controls` and you must rename it in `frmManager.code.vb` too. A mismatch is
not a build error — the form compiles and the handler simply never fires, which
is why `check_sources.py` verifies the two agree.

## How the ribbon manages a user-editable list

Ribbon XML is read once when Word loads the template and cannot be regenerated.
That is awkward for a list the user edits at run time. QuickPhrase solves it two
ways at once:

- **Favorites** declares a fixed pool of 12 buttons (`qpBtn01`…`qpBtn12`). Each
  asks `GetFavLabel` / `GetFavVisible` what to display, so a single
  `IRibbonUI.Invalidate` relabels and hides them. Real flat buttons; capped
  count.
- **All Phrases** is a `dynamicMenu`. `GetMenuContent` returns menu XML built at
  click time, so it has no limit — at the cost of one extra click.

To change the favourites count, edit `FAV_COUNT` in `QuickPhraseRibbon.bas`
**and** the button list in both customUI files — in **both placements**, so four
pools in total. `check_sources.py` fails if any disagree.

**The same phrases appear in two places.** A second copy of the buttons is
declared inside Word's built-in `TabHome`, with ids prefixed `qpHomeBtn`, and
`SnippetStore.Location` decides which placement `getVisible` reveals. Both are
always present in the XML; only visibility changes, because the ribbon cannot be
rebuilt at run time. `IndexFromButtonId` strips either prefix so every callback
is shared rather than duplicated.

Because two `dynamicMenu` controls can now exist at once, `GetMenuContent`
namespaces the controls it generates with the id of the menu that asked for
them. Without that the two menus emit colliding ids and Word shows one of them
empty.

`GetMenuContent` picks its namespace from `Application.Version`: Word 2007 wants
the 2006 namespace, 2010+ the 2009 one. A mismatch produces an empty menu with
no error — worth remembering if the menu ever comes up blank.

## Editing VBA in the VBE instead

Faster than rebuilding for each experiment:

1. Build and install as usual.
2. In Word press <kbd>Alt</kbd>+<kbd>F11</kbd> and find the `QuickPhrase` project.
3. Edit, then **export the changed module back over `src/modules/`** —
   right-click → Export File — and re-run `build.ps1 -ExportVba`. Otherwise the
   next build overwrites your work.

Ribbon XML cannot be edited this way. Change the XML and repack.

## Gotchas worth knowing

**The `IRibbonUI` reference is fragile.** An unhandled VBA error resets the
project and drops it, after which `Invalidate` silently does nothing until Word
rebuilds the ribbon. `RefreshRibbon` swallows that on purpose; the **Reload**
button is the user-facing escape hatch.

**UTF-8 without a BOM needs the two-stream dance.** `ADODB.Stream` always writes
a BOM in text mode, so `SnippetStore.WriteUtf8` stages the text, flips the stream
to binary, seeks past the three BOM bytes and copies into a second stream.
Skipping it leaves a BOM that surfaces as `ï»¿` in other editors. Also note
`Type` may only be changed while `Position` is 0.

**`MsgBox` cannot display Persian.** VBA's built-in `MsgBox` converts its text
through the system ANSI codepage, so anything outside it becomes a question
mark — `Delete "سلام"?` renders as `Delete "??????"`. `UnicodeUI.MsgBoxW` calls
the Windows `MessageBoxW` API with UTF-16 pointers instead. **Use it for any
message that can contain text the user typed.** The same trap applies to
`InputBox` and to a control's `Caption`, which is why phrase text is edited in a
`TextBox` rather than shown in a label.

**Line breaks change form three times.** `vbLf` in the file, `vbCrLf` in the
dialog's text box, `vbCr` in a Word range. Each conversion is explicit in the
code; dropping one produces stray boxes or lost line breaks.

**`[Content_Types].xml` must be the first zip entry.** `pack.py` enforces the
order, and `verify` checks it. Word rejects the package otherwise.

**Buttons need no image.** Ribbon buttons at `size="normal"` render label-only
quite happily, which is why there are no icon assets to maintain.

## Releasing

1. Bump `QP_VERSION` in `QuickPhraseMain.bas` and add a `CHANGELOG.md` entry.
2. If the VBA changed: `build\build.ps1 -ExportVba` on Windows, and commit the blob.
3. Smoke-test a real install: insert a phrase, add/edit/delete/reorder one,
   import and export, restart Word and confirm everything persisted.
4. Tag and push:
   ```bash
   git tag v1.0.0
   git push origin v1.0.0
   ```
5. `release.yml` packs the template, bundles it with the installer and publishes
   the Release. It fails deliberately if `vbaProject.bin` is missing, since a
   release with no macros is worse than no release.
