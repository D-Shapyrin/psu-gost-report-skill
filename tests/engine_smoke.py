"""Engine smoke test: run pandoc + _postprocess_body + normalize on real tests.

python-docx/docxcompose are absent here, so they are stubbed before importing
build.py. The title-page/merge steps need them and are covered by CI; the XML
postprocessing (TOC title, dash lists, formula numbers, page geometry) is what
this test exercises.
"""
import sys, types, zipfile, re, subprocess
from pathlib import Path

docx_mod = types.ModuleType("docx"); docx_mod.Document = lambda *a, **k: None
sys.modules["docx"] = docx_mod
compose_mod = types.ModuleType("docxcompose"); compose_mod.__path__ = []
composer_mod = types.ModuleType("docxcompose.composer"); composer_mod.Composer = object
sys.modules["docxcompose"] = compose_mod
sys.modules["docxcompose.composer"] = composer_mod

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "gost-lab-report" / "scripts"))
import build  # noqa: E402

WORK = Path(__file__).resolve().parent / "_scratch" / "engine"
WORK.mkdir(parents=True, exist_ok=True)

def run_case(md: Path, label: str):
    meta = build.parse_yaml_front_matter(md)
    profile = build.profile_for_metadata(meta)
    ref = build.resolve_reference_doc(meta)
    body = WORK / f"{label}.docx"
    subprocess.run(
        ["pandoc", str(md), "-o", str(body), f"--reference-doc={ref}",
         f"--lua-filter={ROOT / 'gost-lab-report' / 'filters' / 'pagebreak.lua'}"],
        check=True, cwd=WORK, capture_output=True,
    )
    build._postprocess_body(body, meta)
    # emulate the merged document: normalize the body's own sections
    build.normalize_final_sections(body, ref)

    z = zipfile.ZipFile(body)
    doc = z.read("word/document.xml").decode("utf-8")
    num = z.read("word/numbering.xml").decode("utf-8")
    import xml.etree.ElementTree as ET
    ET.fromstring(doc); ET.fromstring(num)  # well-formed
    print(f"\n=== {label} (profile {profile}) ===")
    print("  title page:", build.resolve_title_page(meta).relative_to(ROOT))
    print("  TOC title:", "ОГЛАВЛЕНИЕ" if "ОГЛАВЛЕНИЕ" in doc else
          ("СОДЕРЖАНИЕ" if "СОДЕРЖАНИЕ" in doc else "—"))
    bullets = re.findall(r'<w:lvl w:ilvl="0">.*?</w:lvl>', num, re.S)
    dash = [b for b in bullets if 'w:val="bullet"' in b and '<w:lvlText w:val="–"' in b]
    print("  dash bullets:", len(dash), "| symbol bullets:",
          sum(1 for b in bullets if "Symbol" in b))
    labels = re.findall(r'<w:t xml:space="preserve">\((?:[А-Я]\.)?\d+\)</w:t>', doc)
    print("  formula labels:", labels)
    print("  margins:", re.search(r'<w:pgMar[^>]*/>', doc).group(0))
    print("  page footer refs:", len(re.findall(r'<w:footerReference', doc)))
    print("  font of marker:", "Times New Roman" in num)
    z.close()
    return doc, num


if __name__ == "__main__":
    for md, label in [
        (ROOT / "tests" / "lab_test.md", "lab"),
        (ROOT / "tests" / "course_test.md", "course"),
        (ROOT / "tests" / "legacy_lab_test.md", "legacy_lab"),
    ]:
        run_case(md, label)
    print("\nOK")
