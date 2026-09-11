# Screenshots

Drop the files here using **exactly these names**, then uncomment the
`Screenshots` block in the root `README.md` — the markdown is already written
and pointing at these paths.

| Filename | What it should show | Notes |
|---|---|---|
| `ribbon-tab.png` | The **QuickPhrase** tab open, with phrase buttons visible | The headline image. Crop to the ribbon only — no desktop, no taskbar. |
| `insert-demo.gif` | Cursor in a document, a click on a phrase button, text appearing | The single most convincing asset. 5–10 seconds, loop it. |
| `manage-dialog.png` | The **Manage Phrases** dialog with several phrases listed | Show a mix of Persian and English entries. |
| `home-tab.png` | The QuickPhrase group inside Word's **Home** tab | Demonstrates the placement setting. |
| `all-phrases-menu.png` | The **All Phrases** dropdown open | Shows the unlimited list. |

## Guidelines

**Crop tightly.** A full-screen capture of Word reduces the ribbon to an
unreadable strip on a phone. Crop to the ribbon, or to the dialog, and nothing
else.

**Use real phrases, not `test` and `asdf`.** The screenshot doubles as the
explanation of what the tool is for.

**Width 1200–1600px.** Wider wastes bandwidth; narrower looks soft on a HiDPI
screen. GitHub scales images down to the column width, never up.

**Light theme.** It matches the default GitHub view and the default Word theme.

**Mind what is in frame.** A screenshot of Word captures whatever document is
open — check for names, addresses, client details or file paths before
committing. Once pushed, it is in the git history even if you delete the file
later.

## Capturing

- **Still image:** <kbd>Win</kbd>+<kbd>Shift</kbd>+<kbd>S</kbd> opens Snipping
  Tool in region mode, which is enough for every PNG above.
- **GIF:** [ScreenToGif](https://www.screentogif.com/) is free, open source and
  records a selected region straight to `.gif`. Keep it under about 3 MB —
  GitHub serves it on every page view.

## Size

Keep each PNG under roughly 500 KB. If one is larger, it is almost certainly an
uncropped full-screen capture. Compress with [Squoosh](https://squoosh.app/) or
`pngquant` if needed.
