#!/usr/bin/env python3
"""Deterministic GOST 7.32 verifier for reports built by gost-lab-report.

Usage:
    python verify.py <report.docx>

Prints one PASS/FAIL line per check with the actual value found and exits
0 (all passed) / 1 (any failure or unreadable file).

Stdlib only (zipfile + xml.etree + re) — must run in a bare container with
no pandoc / python-docx / docxcompose installed.

Scope: the checks target the report BODY styles and the document-level page
geometry. Title page styles ("Title*" ids, e.g. TitleNormal transplanted by
the merge) are intentionally NOT held to GOST values — the title page
legitimately uses its own fonts/spacing. We only ever look up styles by the
specific GOST style ids that build.py transplants from reference.docx, so
title styles never enter any check.
"""

import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R_ID = "{%s}id" % R_NS


def w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


# ---------------------------------------------------------------------------
# Reference values. Source: templates/reference.docx of the skill
# (word/styles.xml and word/document.xml sectPr), inspected 2026-07-17.
# Units: w:sz — half-points; w:line — 240ths of a line for lineRule="auto";
# w:ind / w:pgMar — twips (1/20 pt, 1 cm = 566.93 twips, 1 mm = 56.69 twips).
# ---------------------------------------------------------------------------

# Normal style: <w:rFonts w:ascii="Times New Roman" .../> <w:sz w:val="28"/>
NORMAL_FONT = "Times New Roman"
NORMAL_SZ = 28  # half-points -> 14 pt (GOST 7.32: 12-14 pt, the skill fixes 14)

# Normal: <w:spacing w:lineRule="auto" w:line="360" .../> -> 360/240 = 1.5
LINE_ONE_AND_HALF = 360
LINE_RULE = "auto"

# Normal: <w:jc w:val="both"/> — body text is justified
NORMAL_JC = "both"

# BodyText: <w:ind w:firstLine="709"/> — 1.25 cm = 708.66 twips, rounded 709
FIRST_LINE_TWIPS = 709
FIRST_LINE_TOL = 3  # rounding slack (708..712 all render as 1.25 cm)

# reference.docx sectPr: <w:pgMar w:left="1701" w:right="851" w:top="1134"
# w:bottom="1134"/> -> 30 / 15 / 20 / 20 mm (1 mm = 56.69 twips).
# Tolerance 20 twips (~0.35 mm): the spbgeu title sectPr carries right="850"
# (independent rounding of the same 15 mm) and may win the merge.
MARGIN_LEFT = 1701
MARGIN_RIGHT = 851
MARGIN_TOP = 1134
MARGIN_BOTTOM = 1134
MARGIN_TOL = 20

# reference.docx sectPr: <w:pgSz w:w="11906" w:h="16838"/> — A4 portrait
PAGE_W = 11906
PAGE_H = 16838
PAGE_TOL = 20

# Source Code paragraph style in reference.docx uses PT Mono; any of these
# counts as monospace so a font-substituted environment still passes intent.
MONOSPACE_FONTS = {
    "PT Mono",
    "Courier New",
    "Courier",
    "Consolas",
    "Liberation Mono",
    "DejaVu Sans Mono",
    "JetBrains Mono",
}

# GOST style ids transplant_reference_styles() must deliver into the merged
# styles.xml (ids as serialized in reference.docx; pandoc binds body
# paragraphs to them, so a missing one breaks formatting silently).
REQUIRED_STYLE_IDS = [
    "Normal",
    "BodyText",
    "FirstParagraph1",  # name "First Paragraph" — pandoc's first para after headings
    "Compact1",         # name "Compact" — list items
    "Heading1",
    "Heading2",
    "Heading3",
    "SourceCode1",      # name "Source Code" — fenced code blocks
    "VerbatimChar",     # inline code character style
    "TableCaption1",
    "UnnumberedHeading1NoTOC1",  # СОДЕРЖАНИЕ heading used by add_toc
]

# ---------------------------------------------------------------------------
# Profile-aware expectations
#
# Values below are fallbacks. When the skill's profile package is reachable
# (templates/<profile>/reference.docx), the expected geometry and body
# typography are read FROM that file — the reference document is the single
# source of truth (ПГУ ИБСТ: right margin 10 mm, not 15).
# ---------------------------------------------------------------------------

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(SKILL_DIR, "templates")
DEFAULT_PROFILE = "my"
LEGACY_PROFILE = "from_the_repository"


class Expect:
    """Expected values, profiled by the reference.docx of the university."""

    def __init__(self):
        self.font = NORMAL_FONT
        self.sz = NORMAL_SZ
        self.line = LINE_ONE_AND_HALF
        self.line_rule = LINE_RULE
        self.jc = NORMAL_JC
        self.first_line = FIRST_LINE_TWIPS
        self.first_line_tol = FIRST_LINE_TOL
        self.margins = {
            "left": MARGIN_LEFT, "right": MARGIN_RIGHT,
            "top": MARGIN_TOP, "bottom": MARGIN_BOTTOM,
        }
        self.margin_tol = MARGIN_TOL
        self.page_w = PAGE_W
        self.page_h = PAGE_H
        self.page_tol = PAGE_TOL
        self.heading_bold = True
        self.source = "built-in fallback"

    def load(self, profile: str) -> "Expect":
        path = os.path.join(TEMPLATES_DIR, profile, "reference.docx")
        if not os.path.exists(path):
            path = os.path.join(TEMPLATES_DIR, LEGACY_PROFILE, "reference.docx")
        if not os.path.exists(path):
            return self
        try:
            zf = zipfile.ZipFile(path)
            styles_root = ET.fromstring(zf.read("word/styles.xml"))
            doc_root = ET.fromstring(zf.read("word/document.xml"))
        except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
            print(f"Note: cannot read {path} ({exc}); using fallback expectations")
            return self

        styles = Styles(styles_root)
        self.font = styles.resolve(normal_sid(styles), pick_rpr_font) or self.font
        self.sz = _int(styles.resolve(normal_sid(styles), pick_rpr_sz)) or self.sz
        sp = styles.resolve(normal_sid(styles), pick_ppr_spacing)
        if sp and _int(sp[0]):
            self.line = _int(sp[0])
            self.line_rule = sp[1] or self.line_rule
        self.jc = styles.resolve(normal_sid(styles), pick_ppr_jc) or self.jc
        hb = styles.resolve(style_by_name(styles, "heading 1") or "Heading1", pick_rpr_bold)
        self.heading_bold = hb is True
        for sid in (style_by_name(styles, "Body Text") or "BodyText",
                style_by_name(styles, "First Paragraph") or "FirstParagraph1"):
            fl = _int(styles.resolve(sid, pick_ppr_first_line))
            if fl is not None:
                self.first_line = fl
                break

        sect = None
        for sect in doc_root.iter(w("sectPr")):
            pass
        if sect is not None:
            mar = sect.find(w("pgMar"))
            if mar is not None:
                for side in self.margins:
                    val = _int(mar.get(w(side)))
                    if val is not None:
                        self.margins[side] = val
            pgsz = sect.find(w("pgSz"))
            if pgsz is not None:
                self.page_w = _int(pgsz.get(w("w"))) or self.page_w
                self.page_h = _int(pgsz.get(w("h"))) or self.page_h
        self.source = path
        return self


# Known artifact of the LibreOffice-authored templates: style ids "-1"/"-2"
# exist both as character/paragraph AND table styles. Word resolves them by
# type, so only SAME-type duplicates are treated as defects.


# ---------------------------------------------------------------------------
# Style model with basedOn inheritance
# ---------------------------------------------------------------------------

class Styles:
    def __init__(self, root: ET.Element):
        self.by_id = {}       # styleId -> element (paragraph/character only wins last)
        self.by_type_id = {}  # (type, styleId) -> element
        for st in root.findall(w("style")):
            sid = st.get(w("styleId"))
            stype = st.get(w("type"))
            if sid is None:
                continue
            self.by_type_id.setdefault((stype, sid), st)
            # Paragraph/character styles get priority in the flat index so a
            # table style sharing the id (known template artifact) can't
            # shadow them.
            if sid not in self.by_id or stype in ("paragraph", "character"):
                self.by_id[sid] = st

    def get(self, sid):
        return self.by_id.get(sid)

    def chain(self, sid, max_depth=20):
        """Yield the style element and its basedOn ancestors."""
        seen = set()
        cur = sid
        for _ in range(max_depth):
            if cur is None or cur in seen:
                return
            seen.add(cur)
            el = self.by_id.get(cur)
            if el is None:
                return
            yield el
            based = el.find(w("basedOn"))
            cur = based.get(w("val")) if based is not None else None

    def resolve(self, sid, picker):
        """Walk the basedOn chain, return the first non-None picker() result."""
        for el in self.chain(sid):
            val = picker(el)
            if val is not None:
                return val
        return None



# В reference.docx профиля "my" стиль Normal записан под авто-id "a"
# (источник — бланк кафедры), поэтому ищем стиль по имени, а не по id.
def style_by_name(styles, name):
    for el in styles.by_id.values():
        nm = el.find(w("name"))
        if nm is not None and nm.get(w("val")) == name:
            return el.get(w("styleId"))
    return None


def normal_sid(styles):
    sid = style_by_name(styles, "Normal")
    if sid:
        return sid
    for el in styles.by_id.values():
        if el.get(w("default")) in ("1", "true", "on"):
            return el.get(w("styleId"))
    return "Normal"


def _bool_attr(el):
    """OOXML on/off value of an element like <w:b/> / <w:b w:val="false"/>."""
    val = el.get(w("val"))
    if val is None:
        return True
    return val.lower() not in ("0", "false", "off", "none")


def pick_rpr_font(el):
    rpr = el.find(w("rPr"))
    if rpr is None:
        return None
    fonts = rpr.find(w("rFonts"))
    if fonts is None:
        return None
    return fonts.get(w("ascii")) or fonts.get(w("hAnsi"))


def pick_rpr_sz(el):
    rpr = el.find(w("rPr"))
    if rpr is None:
        return None
    sz = rpr.find(w("sz"))
    return sz.get(w("val")) if sz is not None else None


def pick_rpr_bold(el):
    rpr = el.find(w("rPr"))
    if rpr is None:
        return None
    b = rpr.find(w("b"))
    return _bool_attr(b) if b is not None else None


def pick_ppr_spacing(el):
    ppr = el.find(w("pPr"))
    if ppr is None:
        return None
    sp = ppr.find(w("spacing"))
    if sp is None or sp.get(w("line")) is None:
        return None
    return (sp.get(w("line")), sp.get(w("lineRule")))


def pick_ppr_jc(el):
    ppr = el.find(w("pPr"))
    if ppr is None:
        return None
    jc = ppr.find(w("jc"))
    return jc.get(w("val")) if jc is not None else None


def pick_ppr_first_line(el):
    ppr = el.find(w("pPr"))
    if ppr is None:
        return None
    ind = ppr.find(w("ind"))
    if ind is None:
        return None
    fl = ind.get(w("firstLine"))
    if fl is not None:
        return fl
    # An explicit hanging indent at this level cancels a first-line indent.
    if ind.get(w("hanging")) not in (None, "0"):
        return "hanging:" + ind.get(w("hanging"))
    return None  # ind present but silent about firstLine -> keep walking chain


# ---------------------------------------------------------------------------
# Check runner
# ---------------------------------------------------------------------------

class Runner:
    def __init__(self):
        self.failures = 0

    def check(self, name, ok, actual, expected):
        status = "PASS" if ok else "FAIL"
        if not ok:
            self.failures += 1
        print(f"[{status}] {name}: actual={actual} (expected {expected})")


def _int(val):
    try:
        return int(val)
    except (TypeError, ValueError):
        return None


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    profile = DEFAULT_PROFILE
    for arg in sys.argv[1:]:
        if arg.startswith("--profile="):
            profile = arg.split("=", 1)[1].strip() or DEFAULT_PROFILE
    if len(args) != 1:
        print("Usage: python verify.py <report.docx> [--profile=my|from_the_repository]")
        sys.exit(2)
    path = args[0]

    exp = Expect().load(profile)
    print(f"Profile: {profile} (expectations from {exp.source})")

    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as exc:
        print(f"[FAIL] package: cannot open {path!r} as a .docx zip ({exc})")
        sys.exit(1)

    try:
        styles_xml = zf.read("word/styles.xml")
        document_xml = zf.read("word/document.xml")
    except KeyError as exc:
        print(f"[FAIL] package: missing part in docx ({exc})")
        sys.exit(1)

    try:
        styles_root = ET.fromstring(styles_xml)
        doc_root = ET.fromstring(document_xml)
    except ET.ParseError as exc:
        print(f"[FAIL] package: malformed XML ({exc})")
        sys.exit(1)

    styles = Styles(styles_root)
    r = Runner()

    # -- 1. Transplanted GOST styles survived the merge --------------------
    REQ_NAMES = {"Normal": "Normal", "BodyText": "Body Text", "FirstParagraph1": "First Paragraph",
                 "Compact1": "Compact", "Heading1": "heading 1", "Heading2": "heading 2",
                 "Heading3": "heading 3", "SourceCode1": "Source Code", "VerbatimChar": "Verbatim Char",
                 "TableCaption1": "Table Caption", "UnnumberedHeading1NoTOC1": "UnnumberedHeading1NoTOC"}
    missing = [sid for sid in REQUIRED_STYLE_IDS
               if styles.get(sid) is None and style_by_name(styles, REQ_NAMES.get(sid, sid)) is None]
    r.check(
        "gost-styles-present",
        not missing,
        "all present" if not missing else "missing: " + ", ".join(missing),
        f"{len(REQUIRED_STYLE_IDS)} transplanted style ids from reference.docx",
    )

    # -- 2. Normal: font / size / line spacing / justification -------------
    font = styles.resolve(normal_sid(styles), pick_rpr_font)
    r.check("normal-font", font == exp.font, font, exp.font)

    sz = _int(styles.resolve(normal_sid(styles), pick_rpr_sz))
    r.check("normal-size", sz == exp.sz,
            f"{sz} half-points" if sz else sz, f"{exp.sz} half-points (14 pt)")

    spacing = styles.resolve(normal_sid(styles), pick_ppr_spacing)
    ok = (
        spacing is not None
        and _int(spacing[0]) == exp.line
        and (spacing[1] or "auto") == exp.line_rule
    )
    r.check("normal-line-spacing", ok,
            f"line={spacing[0]} rule={spacing[1]}" if spacing else None,
            f"line={exp.line} rule={exp.line_rule} (1.5 spacing)")

    jc = styles.resolve(normal_sid(styles), pick_ppr_jc)
    r.check("normal-justified", jc == exp.jc, jc, exp.jc)

    # -- 3. Red line: first-line indent 1.25 cm on body styles -------------
    for sid in (style_by_name(styles, "Body Text") or "BodyText",
                style_by_name(styles, "First Paragraph") or "FirstParagraph1"):
        fl = styles.resolve(sid, pick_ppr_first_line)
        fl_int = _int(fl)
        ok = (
            fl_int is not None
            and abs(fl_int - exp.first_line) <= exp.first_line_tol
        )
        r.check(f"first-line-indent[{sid}]", ok,
                f"{fl} twips" if fl is not None else "none",
                f"{exp.first_line} twips (1.25 cm)")

    # -- 4. Headings bold state matches the profile reference ---------------
    for sid in (style_by_name(styles, "heading 1") or "Heading1",
                style_by_name(styles, "heading 2") or "Heading2",
                style_by_name(styles, "heading 3") or "Heading3"):
        bold = styles.resolve(sid, pick_rpr_bold) if styles.get(sid) is not None else None
        want = exp.heading_bold
        r.check(f"heading-bold[{sid}]", (bold is True) == want,
                f"{'bold' if bold is True else 'not bold'}",
                f"{'bold' if want else 'not bold'} (per reference.docx)")

    # -- 5. Source Code / inline code are monospace -------------------------
    for sid, label in ((style_by_name(styles, "Source Code") or "SourceCode1", "Source Code"),
                       (style_by_name(styles, "Verbatim Char") or "VerbatimChar", "Verbatim Char")):
        f = styles.resolve(sid, pick_rpr_font) if styles.get(sid) is not None else None
        r.check(f"monospace[{label}]", f in MONOSPACE_FONTS, f,
                f"one of {sorted(MONOSPACE_FONTS)}")

    # -- 6. Page geometry: last sectPr = body section -----------------------
    sect_prs = doc_root.iter(w("sectPr"))
    sect = None
    for sect in sect_prs:
        pass  # keep last
    if sect is None:
        r.check("page-margins", False, "no sectPr found",
                "left/right/top/bottom = {}/{}... twips".format(
                    exp.margins["left"], exp.margins["right"]))
        r.check("page-size-a4", False, "no sectPr found",
                f"{exp.page_w}x{exp.page_h} twips")
    else:
        mar = sect.find(w("pgMar"))
        vals = {
            side: _int(mar.get(w(side))) if mar is not None else None
            for side in ("left", "right", "top", "bottom")
        }
        want = exp.margins
        ok = all(
            vals[s] is not None and abs(vals[s] - want[s]) <= exp.margin_tol
            for s in want
        )
        r.check(
            "page-margins", ok,
            "/".join(str(vals[s]) for s in ("left", "right", "top", "bottom")),
            "/".join(str(want[s]) for s in ("left", "right", "top", "bottom"))
            + " twips +-20",
        )
        pgsz = sect.find(w("pgSz"))
        pw = _int(pgsz.get(w("w"))) if pgsz is not None else None
        ph = _int(pgsz.get(w("h"))) if pgsz is not None else None
        ok = (
            pw is not None and ph is not None
            and abs(pw - exp.page_w) <= exp.page_tol
            and abs(ph - exp.page_h) <= exp.page_tol
        )
        r.check("page-size-a4", ok, f"{pw}x{ph}",
                f"{exp.page_w}x{exp.page_h} twips (A4 portrait)")

    # -- 6b. Page numbers: bottom-centre, no number on the title page -------
    # Title template supplies the title section (with its own titlePg), the
    # body section carries the PAGEFAULT footer copied by pandoc from
    # reference.docx: centered PAGE field + <w:titlePg/>.
    footer_parts = [n for n in zf.namelist()
                    if re.match(r"word/footer[^/]*\.xml$", n)]
    page_field = False
    centered = False
    for name in footer_parts:
        try:
            froot = ET.fromstring(zf.read(name))
        except ET.ParseError:
            continue
        has_page = any(
            "PAGE" in (el.text or "") for el in froot.iter(w("instrText"))
        )
        if not has_page:
            continue
        page_field = True
        for para in froot.iter(w("p")):
            jc_el = para.find(w("pPr") + "/" + w("jc"))
            jc_val = jc_el.get(w("val")) if jc_el is not None else None
            if jc_val is None:
                style_el = para.find(w("pPr") + "/" + w("pStyle"))
                if style_el is not None:
                    jc_val = styles.resolve(style_el.get(w("val")), pick_ppr_jc)
            if jc_val == "center":
                centered = True
    r.check("page-number-field", page_field,
            f"{len(footer_parts)} footer part(s)" if footer_parts else "no footers",
            "a footer with a PAGE field")
    r.check("page-number-centered", centered, centered, "PAGE field centered")

    # Первая страница (титул) должна быть без номера: либо единственная секция
    # с titlePg и пустым «first»-колонтитулом, либо отдельная первая секция с
    # пустыми колонтитулами.
    def footer_has_page(target: str) -> bool:
        try:
            froot = ET.fromstring(zf.read("word/" + target))
        except (KeyError, ET.ParseError):
            return False
        return any("PAGE" in (el.text or "") for el in froot.iter(w("instrText")))

    rel_map = {}
    try:
        rel_root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
        for rel in rel_root:
            rel_map[rel.get("Id")] = rel.get("Target")
    except (KeyError, ET.ParseError):
        rel_map = {}

    all_sects = list(doc_root.iter(w("sectPr")))
    first_page_numbered = None  # None = не удалось определить
    if all_sects:
        if len(all_sects) == 1:
            only = all_sects[0]
            has_title_pg = only.find(w("titlePg")) is not None
            refs = {r.get(w("type")): rel_map.get(r.get(R_ID)) for r in only.iter(w("footerReference"))}
            target = refs.get("first") if has_title_pg else refs.get("default")
            if target:
                first_page_numbered = footer_has_page(target)
        else:
            refs = {r.get(w("type")): rel_map.get(r.get(R_ID))
                    for r in all_sects[0].iter(w("footerReference"))}
            targets = [t for t in refs.values() if t]
            first_page_numbered = any(footer_has_page(t) for t in targets)
    r.check("title-page-unnumbered",
            first_page_numbered is False,
            "no number" if first_page_numbered is False else
            ("number present" if first_page_numbered else "undetermined"),
            "no page number on the first page (empty first footer / titlePg)")

    # -- 7. styles.xml hygiene after merge -----------------------------------
    # 7a. No same-type duplicate style ids (cross-type dups "-1"/"-2" are a
    #     known artifact of the LibreOffice-made templates and are harmless).
    seen, dups = {}, []
    for st in styles_root.findall(w("style")):
        key = (st.get(w("type")), st.get(w("styleId")))
        if key in seen:
            dups.append(f"{key[0]}:{key[1]}")
        seen[key] = True
    r.check("no-duplicate-style-ids", not dups,
            "none" if not dups else ", ".join(dups),
            "no two styles share (type, styleId)")

    # 7b. At most one w:default="1" per style type — a second default
    #     paragraph style is exactly the kind of merge debris that reflows
    #     the whole body.
    defaults = {}
    for st in styles_root.findall(w("style")):
        dflt = st.get(w("default"))
        if dflt and dflt.lower() in ("1", "true"):
            defaults.setdefault(st.get(w("type")), []).append(st.get(w("styleId")))
    extra = {t: ids for t, ids in defaults.items() if len(ids) > 1}
    r.check("single-default-per-type", not extra,
            "ok" if not extra else "; ".join(f"{t}: {ids}" for t, ids in extra.items()),
            'at most one w:default="1" style per type')

    # -- 8. document.xml sanity ----------------------------------------------
    # 8a. Every referenced paragraph style must exist (a dangling pStyle
    #     silently falls back to the default style = lost formatting).
    para_style_ids = {
        sid for (stype, sid) in styles.by_type_id if stype in ("paragraph", None)
    }
    refs = sorted({
        ps.get(w("val"))
        for ps in doc_root.iter(w("pStyle"))
        if ps.get(w("val"))
    })
    dangling = [ref for ref in refs if ref not in para_style_ids]
    r.check("no-dangling-style-refs", not dangling,
            "none" if not dangling else ", ".join(dangling),
            "every w:pStyle in document.xml exists in styles.xml")

    # 8b. The body is actually bound to the GOST styles (not re-tagged to the
    #     title document's defaults during the merge).
    body_bound = any(ref in (style_by_name(styles, "Body Text"), style_by_name(styles, "First Paragraph")) for ref in refs)
    r.check("body-uses-gost-styles", body_bound,
            f"pStyle refs: {', '.join(refs) or 'none'}",
            "at least one paragraph styled BodyText/FirstParagraph1")

    # -- Summary --------------------------------------------------------------
    total = r.failures
    if total:
        print(f"\nFAIL: {total} failed check(s)")
    else:
        print("\nPASS: all checks passed")
    sys.exit(1 if total else 0)


if __name__ == "__main__":
    main()
