# QuickPhrase

Word add-in that puts your frequently used phrases as ribbon buttons — click to insert at the cursor. Add, edit, reorder and delete snippets from a dialog; import/export as JSON. Works in Word 2007–365. Full UTF-8 / RTL support.

---

## What it does

A **QuickPhrase** tab appears next to Home. Your phrases sit there as buttons. Type in your document, click a button, the text lands at the cursor. No dialog, no clipboard, no retyping.

- **Favorites group** — your first 12 phrases, each one its own button, one click away.
- **All Phrases menu** — every phrase, no limit, one click deeper.
- **Manage Phrases** — add, edit, reorder, delete. Changes show on the ribbon immediately, no restart.
- **Import / Export** — phrases live in a single JSON file you can back up, sync or share.

Persian, Arabic, Hebrew and emoji all work: the phrase file is UTF-8, and text is inserted through Word's own typing engine, so bidirectional text keeps its correct direction.

## Install

1. Download `QuickPhrase.dotm` from the [Releases page](https://github.com/cenaav/QuickPhrase/releases).
2. Download this repo's `install` folder (or clone it) and put `QuickPhrase.dotm` next to `Install.bat`.
3. Close Word.
4. Double-click **`Install.bat`**.
5. Start Word.

No admin rights needed. The installer copies the template into
`%APPDATA%\Microsoft\Word\STARTUP\` and clears the "downloaded from the internet"
flag that would otherwise make Word disable the macros without telling you.

Prefer doing it by hand, or hitting a problem? See **[docs/INSTALL.md](docs/INSTALL.md)**.

To remove it: double-click **`Uninstall.bat`**. Your phrases are kept.

### نصب (فارسی)

۱. فایل `QuickPhrase.dotm` را از صفحهٔ [Releases](https://github.com/cenaav/QuickPhrase/releases) دانلود کنید.
۲. آن را کنار `Install.bat` قرار دهید.
۳. Word را کامل ببندید.
۴. روی `Install.bat` دوبار کلیک کنید.
۵. Word را باز کنید — تب **QuickPhrase** کنار Home ظاهر می‌شود.

اگر تب ظاهر نشد، ماکروها غیرفعال هستند. مسیر:
`File > Options > Trust Center > Trust Center Settings > Macro Settings`
گزینهٔ **Disable all macros with notification** را انتخاب کنید.

عبارت‌های شما در این فایل ذخیره می‌شوند:
`%APPDATA%\QuickPhrase\snippets.json`

## Using it

**Insert:** put the cursor where you want the text, click the phrase button. The insertion is a single undo step — one <kbd>Ctrl</kbd>+<kbd>Z</kbd> removes it.

**Add or edit:** *QuickPhrase → Manage Phrases…*

| Field | Meaning |
|---|---|
| **Button label** | Short text shown on the ribbon button |
| **Text inserted at the cursor** | What actually goes into the document — can be multiple lines |

Label and text are separate on purpose: a button can read `Signature` while inserting three lines of contact details.

Order matters — the first 12 phrases become ribbon buttons, so use **Up** / **Down** to promote the ones you use most.

**Share a phrase set:** *Export…* writes a `.json` file. Someone else uses *Import…* to load it, choosing whether to replace their list or append to it. Two ready-made sets are in [examples/](examples/).

## Where your phrases live

```
%APPDATA%\QuickPhrase\snippets.json
```

Plain UTF-8 JSON, safe to edit by hand:

```json
[
  {
    "label": "Best regards",
    "text": "Best regards,\nYour Name"
  }
]
```

Use `\n` for line breaks. After editing the file outside Word, click **Reload** on the ribbon.

Putting this file in a cloud-synced folder and symlinking it is a simple way to share one phrase set across machines.

## Compatibility

| | |
|---|---|
| **Word** | 2007, 2010, 2013, 2016, 2019, 2021, 2024, Microsoft 365 |
| **OS** | Windows |
| **Word for Mac** | Not supported — VBA there cannot add ribbon tabs |
| **Word Online** | Not supported — no VBA |

## Building from source

The repo holds source only. `QuickPhrase.dotm` is a compiled OOXML package containing `vbaProject.bin`, which nothing but Word itself can produce — so it is **not** checked in, and GitHub Actions cannot build it either (no Word on the runners). Releases are built by hand.

On Windows with Word installed:

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1
```

Output lands in `dist\QuickPhrase.dotm`. One-time prerequisite: tick *Trust access to the VBA project object model* in Word's Trust Center. Details and the source layout are in **[docs/DEVELOPING.md](docs/DEVELOPING.md)**.

## Licence

[MIT](LICENSE). Free for personal **and commercial** use — modify it, ship it, sell it. The only condition is that you keep the copyright and licence notice, i.e. credit where it came from.
