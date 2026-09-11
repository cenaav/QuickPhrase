#!/usr/bin/env python3
"""Build dist/QuickPhrase.dotm from the exploded package in package/.

A .dotm is an OOXML zip. Every part is text and lives in package/ or
src/customUI/, except word/vbaProject.bin - a compiled VBA blob that only Word
can produce. So the split is:

    pack     zip package/ + src/customUI/ + the blob -> a .dotm.
             Pure stdlib, runs anywhere, which is what lets CI build releases
             on a Linux runner with no Word installed.

    explode  the reverse: take a .dotm Word wrote and refresh package/ from it.
             Use this if a hand-edited part ever makes Word reject the file, or
             to pull in parts a newer Word version expects.

The blob only changes when the VBA source changes. Regenerate it on Windows:

    powershell -ExecutionPolicy Bypass -File build\\build.ps1 -ExportVba

then commit package/word/vbaProject.bin.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
PACKAGE_DIR = REPO_ROOT / "package"
CUSTOMUI_DIR = REPO_ROOT / "src" / "customUI"
VBA_BLOB = PACKAGE_DIR / "word" / "vbaProject.bin"
DEFAULT_OUTPUT = REPO_ROOT / "dist" / "QuickPhrase.dotm"

# [Content_Types].xml must come first in the archive. The rest of the order is
# cosmetic, but a stable order keeps builds byte-comparable.
FIRST_ENTRY = "[Content_Types].xml"

VBA_MISSING_HELP = f"""
Missing: {VBA_BLOB.relative_to(REPO_ROOT)}

This is the compiled VBA project. Nothing but Microsoft Word can generate it,
so it is committed to the repo rather than built here.

To generate it, on Windows with Word installed:

    powershell -ExecutionPolicy Bypass -File build\\build.ps1 -ExportVba

then commit the resulting package/word/vbaProject.bin.

Only needed when the VBA source under src/ changes. Edits to the ribbon XML,
docs or workflows do not require a new blob.
"""


def collect_parts() -> list[tuple[str, Path]]:
    """Return (archive name, source path) for every part of the package."""
    parts: list[tuple[str, Path]] = []

    for path in sorted(PACKAGE_DIR.rglob("*")):
        if path.is_dir():
            continue
        parts.append((path.relative_to(PACKAGE_DIR).as_posix(), path))

    # The ribbon XML is shared with the Word-automation build, so src/customUI
    # stays the single source of truth and is mapped in at pack time.
    for path in sorted(CUSTOMUI_DIR.glob("*.xml")):
        parts.append((f"customUI/{path.name}", path))

    parts.sort(key=lambda item: (item[0] != FIRST_ENTRY, item[0]))
    return parts


def cmd_pack(output: Path, allow_missing_vba: bool) -> int:
    if not PACKAGE_DIR.is_dir():
        print(f"error: {PACKAGE_DIR} not found", file=sys.stderr)
        return 1

    if not VBA_BLOB.is_file():
        if not allow_missing_vba:
            print(VBA_MISSING_HELP.strip(), file=sys.stderr)
            return 1
        print("warning: packing WITHOUT vbaProject.bin - the result has no "
              "macros and no ribbon tab", file=sys.stderr)

    parts = collect_parts()
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, path in parts:
            # Fixed timestamp so the same sources always produce the same file.
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            zf.writestr(info, path.read_bytes())

    print(f"packed {len(parts)} parts -> {output}")
    for name, _ in parts:
        print(f"  {name}")
    return 0


def cmd_explode(source: Path) -> int:
    """Refresh package/ from a .dotm that Word produced."""
    if not source.is_file():
        print(f"error: {source} not found", file=sys.stderr)
        return 1

    keep_blob = VBA_BLOB.read_bytes() if VBA_BLOB.is_file() else None

    with zipfile.ZipFile(source) as zf:
        names = zf.namelist()
        if "word/document.xml" not in names:
            print(f"error: {source} does not look like a Word package",
                  file=sys.stderr)
            return 1

        if PACKAGE_DIR.exists():
            shutil.rmtree(PACKAGE_DIR)

        for name in names:
            if name.endswith("/"):
                continue
            # customUI belongs to src/, not package/ - skip it so the exploded
            # copy cannot drift from the version the Word build uses.
            if name.startswith("customUI/"):
                continue

            target = PACKAGE_DIR / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(zf.read(name))
            print(f"  {name}")

    if keep_blob is not None and not VBA_BLOB.is_file():
        VBA_BLOB.parent.mkdir(parents=True, exist_ok=True)
        VBA_BLOB.write_bytes(keep_blob)
        print("  word/vbaProject.bin (kept from previous package/)")

    print(f"exploded {source} -> {PACKAGE_DIR}")
    print("Review the diff before committing; Word may have added parts.")
    return 0


def cmd_verify(output: Path, allow_missing_vba: bool = False) -> int:
    """Structural checks on a built .dotm."""
    if not output.is_file():
        print(f"error: {output} not found - run 'pack' first", file=sys.stderr)
        return 1

    problems: list[str] = []
    with zipfile.ZipFile(output) as zf:
        names = zf.namelist()

        if names[0] != FIRST_ENTRY:
            problems.append(f"{FIRST_ENTRY} must be the first entry, found {names[0]}")

        required = [
            FIRST_ENTRY,
            "_rels/.rels",
            "word/document.xml",
            "customUI/customUI.xml",
            "customUI/customUI14.xml",
        ]
        for name in required:
            if name not in names:
                problems.append(f"missing part: {name}")

        if "word/vbaProject.bin" not in names and not allow_missing_vba:
            problems.append("missing part: word/vbaProject.bin (no macros)")

        rels = zf.read("_rels/.rels").decode("utf-8")
        for rel_type, target in [
            ("office/2006/relationships/ui/extensibility", "customUI/customUI.xml"),
            ("office/2007/relationships/ui/extensibility", "customUI/customUI14.xml"),
        ]:
            if rel_type not in rels:
                problems.append(f"_rels/.rels has no relationship of type .../{rel_type}")
            if target not in rels:
                problems.append(f"_rels/.rels does not target {target}")

        doc_rels = zf.read("word/_rels/document.xml.rels").decode("utf-8")
        if "vbaProject.bin" not in doc_rels:
            problems.append("word/_rels/document.xml.rels does not reference vbaProject.bin")

        types = zf.read(FIRST_ENTRY).decode("utf-8")
        if "macroEnabledTemplate.main+xml" not in types:
            problems.append("document.xml is not typed as a macro-enabled template")
        if 'Extension="bin"' not in types:
            problems.append('[Content_Types].xml has no Default for Extension="bin"')

        bad = zf.testzip()
        if bad:
            problems.append(f"corrupt entry: {bad}")

    if problems:
        print(f"FAIL  {output.name}", file=sys.stderr)
        for problem in problems:
            print(f"  - {problem}", file=sys.stderr)
        return 1

    size_kb = output.stat().st_size / 1024
    print(f"PASS  {output.name}  ({len(names)} parts, {size_kb:.1f} KiB)")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    p_pack = sub.add_parser("pack", help="build the .dotm")
    p_pack.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT)
    p_pack.add_argument("--allow-missing-vba", action="store_true",
                        help="pack anyway, for testing the package structure only")

    p_explode = sub.add_parser("explode", help="refresh package/ from a Word-built .dotm")
    p_explode.add_argument("source", type=Path)

    p_verify = sub.add_parser("verify", help="structural checks on a built .dotm")
    p_verify.add_argument("-o", "--output", type=Path, default=DEFAULT_OUTPUT)
    p_verify.add_argument("--allow-missing-vba", action="store_true",
                          help="do not require word/vbaProject.bin")

    args = parser.parse_args()

    if args.command == "pack":
        return cmd_pack(args.output.resolve(), args.allow_missing_vba)
    if args.command == "explode":
        return cmd_explode(args.source.resolve())
    if args.command == "verify":
        return cmd_verify(args.output.resolve(), args.allow_missing_vba)
    return 1


if __name__ == "__main__":
    sys.exit(main())
