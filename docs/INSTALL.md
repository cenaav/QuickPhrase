# Installing QuickPhrase

## The short version

Close Word, put `QuickPhrase.dotm` next to `Install.bat`, double-click `Install.bat`, start Word.

## Manual install

If you would rather not run a script:

1. Close Word completely (check Task Manager for a stray `WINWORD.EXE`).
2. Right-click `QuickPhrase.dotm` → **Properties**. If there is an **Unblock**
   checkbox at the bottom of the General tab, tick it and press OK. Skipping
   this is the single most common reason the add-in appears to do nothing.
3. Press <kbd>Win</kbd>+<kbd>R</kbd>, paste this and press Enter:
   ```
   %APPDATA%\Microsoft\Word\STARTUP
   ```
4. Copy `QuickPhrase.dotm` into that folder.
5. Start Word.

Word loads every template in `STARTUP` at launch, for every document, so the
tab is always there.

## Uninstall

Double-click `Uninstall.bat`, or delete `QuickPhrase.dotm` from the `STARTUP`
folder by hand.

Your phrases are **not** removed — they stay in `%APPDATA%\QuickPhrase\`, so
reinstalling picks up where you left off. To delete them too:

```powershell
powershell -ExecutionPolicy Bypass -File install\Uninstall.ps1 -RemoveData
```

## Troubleshooting

### No QuickPhrase tab after restarting Word

Work through these in order.

**1. Is the file in the right place?**
Paste `%APPDATA%\Microsoft\Word\STARTUP` into the Explorer address bar and
confirm `QuickPhrase.dotm` is there.

**2. Is the file still blocked?**
Right-click it → Properties → tick **Unblock** → OK. Restart Word.

**3. Are macros disabled?**
*File → Options → Trust Center → Trust Center Settings → Macro Settings.*

Pick **Disable all macros with notification**. QuickPhrase works with this
setting because the `STARTUP` folder is a trusted location, so no prompt
appears. What does not work is **Disable all macros without notification** —
that blocks everything unconditionally.

Leave *Trust access to the VBA project object model* **off**. Only the build
script needs it, never the add-in.

**4. Is the add-in loaded but the tab hidden?**
*File → Options → Add-ins → Manage: **Templates** → Go.* `QuickPhrase.dotm`
should be listed and ticked.

**5. Did your organisation block it?**
Managed machines often have a Group Policy that blocks macros from files
originating outside the network, which no local setting can override. Ask
whoever administers your machine to allow it, or to add it to a trusted
location. This is a policy decision, not a bug.

### The tab is there but buttons are blank

The phrase file is empty or unreadable. Click **Manage Phrases…** and add one,
or check the file:

```
%APPDATA%\QuickPhrase\snippets.json
```

If you edited it by hand and broke the JSON, QuickPhrase says so on startup
and leaves the file untouched so you can fix it.

### Clicking a button does nothing

- **No document open** — QuickPhrase says so; open one.
- **Document is protected or read-only** — Word rejects the insertion.
  *Review → Restrict Editing* to check.
- **Cursor is somewhere text cannot go**, e.g. a selected image. Click into
  the text first.

### Persian or Arabic text inserts in the wrong direction

The paragraph's direction, not QuickPhrase, controls this. Set it on the Home
tab with the right-to-left paragraph button (<kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Left
Shift</kbd> in most builds) before inserting.

Note that the **Manage Phrases** dialog itself is always left-to-right. VBA
dialog controls have no right-to-left mode, so Persian text shows correctly but
the cursor moves left-to-right while you edit. This affects the editor only;
the document gets exactly the characters you saved.

### Changes in the dialog do not show on the ribbon

Click **Reload** on the QuickPhrase tab. This normally happens automatically;
it can be missed if a macro error reset the VBA project earlier in the session.
Restarting Word always fixes it.
