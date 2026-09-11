#!/usr/bin/env python3
"""Static checks on the QuickPhrase sources.

The build wires several files together by name, and no compiler checks those
links. A mismatch produces a template that loads fine and then silently does
nothing - a ribbon button with no handler, or a dialog button that ignores
clicks. These checks catch that on every push instead of after a release.

    python3 build/check_sources.py
"""

from __future__ import annotations

import json
import re
import sys
import xml.dom.minidom
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

CUSTOMUI_DIR = REPO_ROOT / "src" / "customUI"
CUSTOMUI_14 = CUSTOMUI_DIR / "customUI14.xml"
CUSTOMUI_07 = CUSTOMUI_DIR / "customUI.xml"
RIBBON_BAS = REPO_ROOT / "src" / "modules" / "QuickPhraseRibbon.bas"
FORM_CODE = REPO_ROOT / "src" / "forms" / "frmManager.code.vb"
BUILD_PS1 = REPO_ROOT / "build" / "build.ps1"
PACKAGE_DIR = REPO_ROOT / "package"

MODULES = [
    "JsonLite.bas",
    "UnicodeUI.bas",
    "PhraseColors.bas",
    "SnippetStore.bas",
    "QuickPhraseRibbon.bas",
    "QuickPhraseMain.bas",
]

# Hungarian prefixes used for dialog controls. Any new prefix must be added
# here, or controls using it silently escape the name check.
CONTROL_PREFIX_RE = r"\b(?:lst|txt|btn|lbl|chk|cbo|opt|fra|img|spn)[A-Z]\w*"

# Callback attributes in the ribbon XML that must name a real VBA procedure.
CALLBACK_ATTRS = (
    "onLoad", "onAction", "getLabel", "getVisible",
    "getScreentip", "getSupertip", "getContent", "getEnabled", "getImage",
)

failures: list[str] = []
checks_run = 0


def check(name: str, ok: bool, detail: str = "") -> None:
    global checks_run
    checks_run += 1
    if ok:
        print(f"  ok    {name}")
    else:
        print(f"  FAIL  {name}" + (f" - {detail}" if detail else ""))
        failures.append(name)


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def section(title: str) -> None:
    print(f"\n{title}")


def check_xml_well_formed() -> None:
    section("XML is well-formed")

    paths = sorted(CUSTOMUI_DIR.glob("*.xml")) + sorted(PACKAGE_DIR.rglob("*.xml"))
    # .rels parts are XML too but do not match *.xml.
    paths += [
        PACKAGE_DIR / "_rels" / ".rels",
        PACKAGE_DIR / "word" / "_rels" / "document.xml.rels",
    ]

    seen: set[Path] = set()
    for path in paths:
        if path in seen:
            continue
        seen.add(path)
        if not path.is_file():
            continue
        try:
            xml.dom.minidom.parse(str(path))
            check(str(path.relative_to(REPO_ROOT)), True)
        except Exception as exc:
            check(str(path.relative_to(REPO_ROOT)), False, str(exc))


def check_json_valid() -> None:
    section("Example phrase packs parse, and match the stored schema")

    for path in sorted((REPO_ROOT / "examples").glob("*.json")):
        rel = str(path.relative_to(REPO_ROOT))
        try:
            data = json.loads(read(path))
        except Exception as exc:
            check(rel, False, str(exc))
            continue

        if not isinstance(data, list):
            check(rel, False, "top level must be an array")
            continue

        bad = [
            i for i, entry in enumerate(data)
            if not isinstance(entry, dict)
            or not isinstance(entry.get("label"), str)
            or not isinstance(entry.get("text"), str)
        ]
        check(rel, not bad, f"entries missing string label/text: {bad}")

        # "newline" is optional - older files predate it - but VBA reads it as a
        # boolean, so a string or number there would be silently dropped.
        bad_flag = [
            i for i, entry in enumerate(data)
            if isinstance(entry, dict)
            and "newline" in entry
            and not isinstance(entry["newline"], bool)
        ]
        check(f"{rel} newline flags are boolean", not bad_flag, f"entries: {bad_flag}")

        # Colours are "#RRGGBB" or absent. VBA reads anything else as no colour.
        bad_color = [
            i for i, entry in enumerate(data)
            if isinstance(entry, dict)
            and entry.get("color")
            and not re.fullmatch(r"#[0-9A-Fa-f]{6}", str(entry["color"]))
        ]
        check(f"{rel} colours are #RRGGBB", not bad_color, f"entries: {bad_color}")


def check_favourite_count() -> None:
    section("FAV_COUNT matches the ribbon button pool in both namespaces")

    match = re.search(r"FAV_COUNT As Long = (\d+)", read(RIBBON_BAS))
    if not match:
        check("FAV_COUNT is declared", False, "not found in QuickPhraseRibbon.bas")
        return

    fav = int(match.group(1))

    # Two placements exist - the dedicated tab and the group mirrored into Home -
    # and both pools must match FAV_COUNT or one of them silently loses buttons.
    for path in (CUSTOMUI_14, CUSTOMUI_07):
        src = read(path)
        for prefix in ("qpBtn", "qpHomeBtn"):
            found = len(re.findall(rf'id="{prefix}\d+"', src))
            check(f"{path.name} has {fav} {prefix} buttons", found == fav,
                  f"found {found}")

    # Buttons are numbered from 1, and IndexFromButtonId subtracts 1, so any gap
    # would leave a dead button.
    for prefix in ("qpBtn", "qpHomeBtn"):
        ids = sorted(int(n) for n in re.findall(rf'id="{prefix}(\d+)"', read(CUSTOMUI_14)))
        check(f"{prefix} ids are 1..N with no gaps", ids == list(range(1, fav + 1)),
              f"got {ids}")


def check_ribbon_callbacks() -> None:
    section("Every ribbon callback exists as a Public Sub")

    ribbon_src = read(RIBBON_BAS)
    defined = set(re.findall(r"^Public Sub (\w+)\(", ribbon_src, re.MULTILINE))

    attr_pattern = r'(?:%s)="(\w+)"' % "|".join(CALLBACK_ATTRS)
    for path in (CUSTOMUI_14, CUSTOMUI_07):
        referenced = set(re.findall(attr_pattern, read(path)))
        missing = sorted(referenced - defined)
        check(f"{path.name} ({len(referenced)} callbacks)", not missing,
              f"undefined: {missing}")

    # The dynamicMenu XML is assembled in VBA string literals, where quotes are
    # doubled, so those onAction names need checking too.
    dynamic = set(re.findall(r'onAction=""(\w+)""', ribbon_src))
    missing = sorted(dynamic - defined)
    check(f"dynamic menu ({len(dynamic)} callbacks)", not missing,
          f"undefined: {missing}")

    # Both namespaces must be present, one per customUI version. Getting this
    # wrong yields an empty menu with no error message.
    check("GetMenuContent emits the 2006 namespace",
          "office/2006/01/customui" in ribbon_src)
    check("GetMenuContent emits the 2009 namespace",
          "office/2009/07/customui" in ribbon_src)


def check_form_controls() -> None:
    section("Manager dialog controls match the ones build.ps1 creates")

    declared = set(re.findall(r"Name='(\w+)'", read(BUILD_PS1)))
    check("build.ps1 declares controls", bool(declared), "none found")

    form_src = read(FORM_CODE)

    # Strip handler signatures first, so "btnSave_Click" is not misread as a
    # control named btnSave_Click.
    body = re.sub(r"Private Sub \w+_(?:Click|Change|Initialize|QueryClose)\([^)]*\)",
                  "", form_src)
    referenced = set(re.findall(CONTROL_PREFIX_RE, body))
    missing = sorted(referenced - declared)
    check(f"{len(referenced)} referenced controls exist", not missing,
          f"not created by build.ps1: {missing}")

    handlers = {
        name for name in re.findall(r"Private Sub (\w+)_(?:Click|Change)\(", form_src)
        if name != "UserForm"
    }
    orphans = sorted(handlers - declared)
    check(f"{len(handlers)} event handlers bind to real controls", not orphans,
          f"no such control: {orphans}")


def check_module_list() -> None:
    section("Module list in build.ps1 matches src/modules")

    on_disk = {p.name for p in (REPO_ROOT / "src" / "modules").glob("*.bas")}
    check("every module on disk is in the build list",
          on_disk == set(MODULES), f"disk={sorted(on_disk)} expected={MODULES}")

    build_src = read(BUILD_PS1)
    for module in MODULES:
        check(f"build.ps1 imports {module}", f"'{module}'" in build_src)

    # QuickPhraseRibbon references QuickPhraseMain and vice versa, so an import
    # order that put Main first would still compile - but Option Explicit means
    # a missing module fails loudly. Just confirm each file has one.
    for module in MODULES:
        src = read(REPO_ROOT / "src" / "modules" / module)
        check(f"{module} has Option Explicit", "Option Explicit" in src)


def check_seed_defaults() -> None:
    """A fresh install must start with exactly two phrases: Hello and سلام.

    They are written with ChrW escapes so the .bas file stays pure ASCII and
    cannot be corrupted by an editor saving in the wrong encoding - which also
    means a typo there is invisible on inspection. Hence this check.
    """
    section("Starter phrases are exactly Hello and سلام")

    src = read(REPO_ROOT / "src" / "modules" / "SnippetStore.bas")
    match = re.search(r"Private Sub SeedDefaults\(\)(.*?)End Sub", src, re.DOTALL)
    if not match:
        check("SeedDefaults exists", False, "not found in SnippetStore.bas")
        return

    body = match.group(1)
    calls = re.findall(r"^\s*AddPhrase\b", body, re.MULTILINE)
    check("exactly 2 starter phrases", len(calls) == 2, f"found {len(calls)}")

    check('"Hello" is one of them', 'AddPhrase "Hello", "Hello"' in body)

    # سلام = U+0633 U+0644 U+0627 U+0645
    salam = ["&H633", "&H644", "&H627", "&H645"]
    chrw = re.findall(r"ChrW\$\((&H[0-9A-Fa-f]+)\)", body)
    check("the Persian phrase is سلام", chrw == salam,
          f"ChrW codes found: {chrw}, expected {salam}")


def check_version_consistency() -> None:
    section("Version string is present and well-formed")

    main_src = read(REPO_ROOT / "src" / "modules" / "QuickPhraseMain.bas")
    match = re.search(r'QP_VERSION As String = "([^"]+)"', main_src)
    check("QP_VERSION is declared", bool(match))
    if match:
        check(f"QP_VERSION '{match.group(1)}' is semver",
              bool(re.fullmatch(r"\d+\.\d+\.\d+", match.group(1))))


def main() -> int:
    print("QuickPhrase source checks")

    check_xml_well_formed()
    check_json_valid()
    check_favourite_count()
    check_ribbon_callbacks()
    check_form_controls()
    check_module_list()
    check_seed_defaults()
    check_version_consistency()

    print(f"\n{checks_run} checks, {len(failures)} failed")
    if failures:
        for name in failures:
            print(f"  failed: {name}")
        return 1

    print("All checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
