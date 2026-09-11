<div align="center">

# QuickPhrase

### Insert your frequently used phrases into Microsoft Word with one click

**A free, open-source Word add-in that turns your most-typed sentences into ribbon buttons.**
Click a button, the text appears at your cursor. No clipboard, no retyping, no dialogs.

[![CI](https://github.com/cenaav/QuickPhrase/actions/workflows/ci.yml/badge.svg)](https://github.com/cenaav/QuickPhrase/actions/workflows/ci.yml)
[![Release](https://img.shields.io/github/v/release/cenaav/QuickPhrase?include_prereleases&sort=semver)](https://github.com/cenaav/QuickPhrase/releases)
[![Downloads](https://img.shields.io/github/downloads/cenaav/QuickPhrase/total)](https://github.com/cenaav/QuickPhrase/releases)
[![Licence: MIT](https://img.shields.io/badge/licence-MIT-blue.svg)](LICENSE)
[![Word 2007–365](https://img.shields.io/badge/Word-2007%20%E2%80%93%20365-2B579A?logo=microsoftword&logoColor=white)](#compatibility)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)](#compatibility)

[**Download**](https://github.com/cenaav/QuickPhrase/releases) ·
[Install guide](docs/INSTALL.md) ·
[FAQ](#faq) ·
[Build from source](docs/DEVELOPING.md) ·
[نصب فارسی](#نصب-فارسی)

</div>

---

## The problem

If your job involves Word, you retype the same text all day. Sign-offs. Standard clauses. Diagnoses. Disclaimers. A greeting in Persian. A patient instruction. A refund policy paragraph.

Word's own answers each ask something of you first:

- **AutoCorrect** needs you to memorise a trigger abbreviation.
- **AutoText / Quick Parts** needs you to remember a name, then navigate a gallery.
- **Copy-paste from a scratch document** needs a second window and destroys your clipboard.
- **Third-party text expanders** need a background process, an account, or a subscription.

Each one puts a lookup between you and the text. QuickPhrase removes it: your phrases are *visible*, sitting on the ribbon as labelled buttons. You point and click.

## What QuickPhrase does

Installing adds a **QuickPhrase** tab next to Home:

```
┌─ Home ─┬─ QuickPhrase ─┬─ Insert ─┬─ Design ─┬─ Layout ─────────────────────┐
│                                                                            │
│  ┌── Favorites ────────────────┐  ┌─ All Phrases ─┐  ┌── Manage ─────────┐  │
│  │  سلام        Best regards   │  │               │  │                   │  │
│  │  کجایی؟      Kind regards   │  │      ▼        │  │  Manage Phrases…  │  │
│  │  خوبی؟       Please find…   │  │  All Phrases  │  │  Import   Export  │  │
│  │  دوستت دارم  Signature      │  │               │  │  Reload   About   │  │
│  └─────────────────────────────┘  └───────────────┘  └───────────────────┘  │
└────────────────────────────────────────────────────────────────────────────┘
```

Type in your document. Click a button. The text lands exactly where the cursor is, as a single undo step.

## Features

|  | Feature | Detail |
|---|---|---|
| 🖱️ | **One-click insertion** | Text goes in at the cursor. One <kbd>Ctrl</kbd>+<kbd>Z</kbd> takes it back out. |
| 📌 | **12 phrases as real ribbon buttons** | Your most-used phrases stay permanently visible, no menu to open. |
| 📜 | **Unlimited phrases in a dropdown** | The **All Phrases** menu holds as many as you like, one click deeper. |
| ✏️ | **Visual editor** | Add, edit, reorder and delete in a dialog. The ribbon updates instantly — no restarting Word. |
| 🏷️ | **Short label, long text** | A button reading `Signature` can insert five lines of contact details. |
| 🌍 | **Persian, Arabic, Hebrew, emoji** | UTF-8 throughout, and insertion via Word's own typing engine, so RTL text keeps its direction. |
| ↕️ | **Reorderable** | Up/Down decides which phrases earn a ribbon button. |
| 💾 | **One plain JSON file** | Human-readable, hand-editable, easy to back up or put in a synced folder. |
| 📤 | **Import / export** | Share a phrase set with colleagues. Import replaces or appends, your choice. |
| 🚫 | **No background process** | It is a Word template. Nothing runs when Word is closed. |
| 🔌 | **No account, no network** | Zero telemetry. Zero network calls. Your phrases never leave your machine. |
| 🆓 | **MIT licensed** | Free for personal *and* commercial use. |

## Install

**[⬇ Download the latest release](https://github.com/cenaav/QuickPhrase/releases/latest)**

1. Download `QuickPhrase-vX.Y.Z.zip` and unzip it.
2. **Close Word** completely.
3. Double-click **`Install.bat`**.
4. Start Word. The **QuickPhrase** tab appears next to Home.

No admin rights. No installer bundle. The script copies one template into
`%APPDATA%\Microsoft\Word\STARTUP\` and clears the "downloaded from the internet"
flag that otherwise makes Word disable macros silently.

To remove it, double-click `Uninstall.bat`. Your phrases are kept.

> **Why can't it be zero-install?** Word only auto-loads add-ins from its `STARTUP`
> folder, so one file copy is unavoidable — for QuickPhrase and for every other
> Word add-in. `Install.bat` reduces that to one double-click.

Manual installation and troubleshooting: **[docs/INSTALL.md](docs/INSTALL.md)**

### نصب فارسی

۱. آخرین نسخه را از صفحهٔ [Releases](https://github.com/cenaav/QuickPhrase/releases/latest) دانلود و فایل zip را باز کنید.
۲. Word را کامل ببندید.
۳. روی **`Install.bat`** دوبار کلیک کنید.
۴. Word را باز کنید — تب **QuickPhrase** کنار Home ظاهر می‌شود.

دسترسی ادمین لازم نیست.

**اگر تب ظاهر نشد**، ماکروها غیرفعال هستند:
`File → Options → Trust Center → Trust Center Settings → Macro Settings`
گزینهٔ **Disable all macros with notification** را انتخاب کنید (نه `without notification`).

**عبارت‌های شما اینجا ذخیره می‌شوند:**
`%APPDATA%\QuickPhrase\snippets.json`

فایل با کدگذاری UTF-8 است، پس متن فارسی و عربی سالم می‌ماند. برای جهت راست‌به‌چپ، جهت پاراگراف را در تب Home تعیین کنید.

## How to use it

### Inserting text

Put the cursor where the text belongs, then click the phrase button. That is all.

Phrases past the first twelve live in the **All Phrases** dropdown, which has no limit.

### Adding and editing phrases

**QuickPhrase → Manage Phrases…**

| Field | What it is |
|---|---|
| **Button label** | The short text shown on the ribbon button |
| **Text inserted at the cursor** | What actually goes into your document — multiple lines allowed |

These are separate on purpose. A button labelled `Disclaimer` can insert an
entire paragraph.

Order matters: the **first 12 phrases become ribbon buttons**, so use **Up** and
**Down** to promote the ones you reach for most. Changes appear on the ribbon as
soon as you press **Save & Close**.

### Sharing a phrase set

**Export…** writes a `.json` file. A colleague uses **Import…** to load it,
choosing whether to replace their list or add to it.

Two ready-made sets ship in [`examples/`](examples/):

- [`persian-starter.json`](examples/persian-starter.json) — Persian greetings and sign-offs
- [`english-business.json`](examples/english-business.json) — English business correspondence

## Where your phrases are stored

```
%APPDATA%\QuickPhrase\snippets.json
```

UTF-8 JSON, no BOM, safe to edit in any text editor:

```json
[
  {
    "label": "Best regards",
    "text": "Best regards,\nJane Doe\nSenior Editor"
  },
  {
    "label": "سلام",
    "text": "سلام"
  }
]
```

Use `\n` for line breaks. After editing the file outside Word, click **Reload**
on the ribbon to pick up the changes.

Want the same phrases on several machines? Put this file in a cloud-synced
folder and symlink it.

## Who it is for

QuickPhrase pays off wherever the same wording is typed repeatedly:

- **Lawyers and paralegals** — standard clauses, disclaimers, case citations
- **Doctors and medical transcriptionists** — findings, patient instructions, referral wording
- **Translators** — stock phrases in both directions, especially Persian ⇄ English
- **Teachers and markers** — grading feedback, rubric comments, parent-letter paragraphs
- **HR and recruiting** — offer wording, rejection letters, policy paragraphs
- **Customer support and admin** — refund policies, escalation wording, sign-offs
- **Academic writers** — citation boilerplate, methods-section stock sentences
- **Anyone writing in Persian or Arabic** on an English keyboard layout, where retyping is slowest

## Compatibility

| | Supported |
|---|---|
| **Microsoft Word** | 2007, 2010, 2013, 2016, 2019, 2021, 2024, Microsoft 365 |
| **Operating system** | Windows (7 and later) |
| **Word bitness** | 32-bit and 64-bit |
| **Word for Mac** | ❌ Not supported — Mac VBA cannot add ribbon tabs |
| **Word Online / Web** | ❌ Not supported — no VBA |
| **Excel / PowerPoint / Outlook** | ❌ Word only, for now |

The template ships ribbon definitions in both the 2006 and 2009 customUI
namespaces, which is what lets a single file cover Word 2007 through Microsoft 365.

## FAQ

<details>
<summary><b>Is it free? Can I use it at work, or in a commercial product?</b></summary>

Yes to all three. QuickPhrase is [MIT licensed](LICENSE): use it, modify it,
redistribute it, bundle it into something you sell. The only condition is that
you keep the copyright and licence notice — credit where it came from.
</details>

<details>
<summary><b>Does it send my phrases anywhere?</b></summary>

No. There is no network code in the project at all. Your phrases live in one
file on your own disk. No account, no telemetry, no sync service.
</details>

<details>
<summary><b>Does it run in the background and slow my computer down?</b></summary>

No. QuickPhrase is a Word template, not an application. Nothing runs unless Word
is open, and inside Word it does nothing until you click a button.
</details>

<details>
<summary><b>Why does Word warn me about macros?</b></summary>

Because QuickPhrase *is* a macro — that is how a Word add-in adds a ribbon tab.
The whole source is in this repository for you to read before trusting it, which
is more than a closed-source alternative can offer.

Once installed in the `STARTUP` folder, Word treats it as a trusted location and
stops asking.
</details>

<details>
<summary><b>The QuickPhrase tab didn't appear. What now?</b></summary>

Almost always one of two things: the file is still marked as downloaded from the
internet, or macros are disabled. Both are covered step by step in
[docs/INSTALL.md](docs/INSTALL.md#troubleshooting).
</details>

<details>
<summary><b>How is this different from AutoCorrect or Quick Parts?</b></summary>

Those require you to recall something first — an abbreviation, or an entry name.
QuickPhrase makes your phrases *visible* as labelled buttons, so recall is
replaced by recognition. It also keeps everything in one portable JSON file you
can read, diff, back up and share; AutoText entries are buried inside
`Normal.dotm`.
</details>

<details>
<summary><b>Can I have more than 12 buttons on the ribbon?</b></summary>

The 12 is a deliberate limit, not laziness: ribbon XML is read once when Word
starts and cannot be regenerated, so the buttons have to be pre-declared. Any
number beyond 12 is available in the **All Phrases** dropdown.

If you want a different count, change `FAV_COUNT` in
`src/modules/QuickPhraseRibbon.bas` and the matching button list in both
customUI files, then rebuild. `build/check_sources.py` verifies they agree.
</details>

<details>
<summary><b>Can phrases contain formatting, tables or images?</b></summary>

Not yet — phrases are plain text, and they adopt the formatting at the cursor.
That is usually what you want. Rich-text phrases are on the [roadmap](#roadmap).
</details>

<details>
<summary><b>Can I assign keyboard shortcuts?</b></summary>

Not directly yet. As a workaround, Word's *Customize Ribbon → Keyboard Shortcuts*
can bind a key to a macro, and you can write a one-line macro calling
`QuickPhraseMain.InsertPhrase 0`.
</details>

<details>
<summary><b>Does it work with Persian and Arabic?</b></summary>

Yes, and that was a design goal. The phrase file is UTF-8, and text is inserted
through Word's own typing mechanism, so bidirectional text behaves exactly as if
you had typed it.

One honest limitation: the **Manage Phrases** dialog is left-to-right. VBA dialog
controls have no right-to-left mode, so Persian text displays correctly but the
cursor moves left-to-right while you edit it. This affects the editor only — your
document receives exactly the characters you saved.
</details>

<details>
<summary><b>Will this work on my work computer?</b></summary>

Usually. It needs no admin rights and installs entirely inside your user profile.
The exception is a managed machine with a Group Policy blocking macros from
outside the corporate network — no local setting overrides that, so you would
need whoever administers it to allow the file.
</details>

## Building from source

The repository contains source only. `QuickPhrase.dotm` is a build artifact,
attached to each [Release](https://github.com/cenaav/QuickPhrase/releases).

Anyone can build it, on any OS:

```bash
python3 build/check_sources.py   # validate the sources
python3 build/pack.py pack       # assemble dist/QuickPhrase.dotm
python3 build/pack.py verify     # structural checks
```

That works with no Word and no Windows, because the `.dotm` is stored **exploded
into its XML parts** under [`package/`](package/) — all diffable text — plus one
committed binary, `package/word/vbaProject.bin`, the compiled VBA project.

That single blob is the only thing Word itself must produce. Regenerate it only
when the VBA source changes, on a Windows machine with Word:

```powershell
powershell -ExecutionPolicy Bypass -File build\build.ps1 -ExportVba
```

This is also what lets GitHub Actions publish releases from a Linux runner.
Design notes, gotchas and the release checklist:
**[docs/DEVELOPING.md](docs/DEVELOPING.md)**

## Roadmap

Not promises — a list of what would most improve the tool, roughly in order.

- [ ] Rich-text phrases (bold, links, tables) via Word building blocks
- [ ] Per-phrase keyboard shortcuts
- [ ] Groups or categories, so one ribbon can hold several phrase sets
- [ ] Placeholder fields, e.g. `{date}` and `{recipient}`
- [ ] A searchable picker for very large phrase libraries
- [ ] Excel and PowerPoint versions
- [ ] Right-to-left manager dialog, if a workable approach exists in VBA

Ideas and votes are welcome in [Issues](https://github.com/cenaav/QuickPhrase/issues).

## Contributing

Contributions are welcome, including bug reports, phrase packs for other
languages, and documentation fixes.

Before opening a pull request, run:

```bash
python3 build/check_sources.py
```

These checks exist because the build wires files together by name — a ribbon
callback, a dialog control — and nothing else catches a mismatch. A typo would
produce a template that loads perfectly and then silently does nothing.

Read [docs/DEVELOPING.md](docs/DEVELOPING.md) first; it explains why there is no
`.frm` file and why the compiled blob is committed.

## Licence

[MIT](LICENSE) — free for personal **and commercial** use. Modify it, ship it,
sell it. The only requirement is that you keep the copyright and licence notice,
crediting where it came from.

---

<div align="center">

**Found this useful? A ⭐ helps other people find it.**

</div>

<!--
Keywords for search: Microsoft Word add-in, Word plugin, Word macro, VBA add-in,
text snippets, snippet manager, text expander, boilerplate text, canned
responses, AutoText alternative, Quick Parts alternative, AutoCorrect
alternative, custom ribbon tab, ribbon customisation, customUI, insert text at
cursor, frequently used phrases, phrase manager, productivity tool for writers,
Word automation, dotm template, global template, Word 2007, Word 2010, Word
2013, Word 2016, Word 2019, Word 2021, Word 2024, Microsoft 365, Persian Word
add-in, Farsi text insertion, Arabic Word add-in, RTL support, right-to-left,
UTF-8, free open source Word add-in, MIT licensed, افزونه ورد, افزونه وُرد فارسی,
درج متن سریع, عبارات پرکاربرد, ماکرو ورد, قالب ورد
-->
