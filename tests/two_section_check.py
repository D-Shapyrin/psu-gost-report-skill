"""Check normalize_final_sections on a two-section (title+body) document."""
import sys, types, zipfile, re, shutil
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

src = ROOT / "tests" / "_scratch" / "engine" / "lab.docx"
out = ROOT / "tests" / "_scratch" / "engine" / "two_sect.docx"
with zipfile.ZipFile(src) as z:
    files = {n: z.read(n) for n in z.namelist()}
doc = files["word/document.xml"].decode("utf-8")

# split the single sectPr into a section break after the first paragraph:
sect = re.search(r'<w:sectPr[^>]*>.*?</w:sectPr>', doc, re.S).group(0)
first_p = re.search(r'<w:p\b[^>]*>.*?</w:p>', doc, re.S).group(0)
section_break = '<w:p><w:pPr>' + sect + '</w:pPr></w:p>'
doc = doc.replace(first_p, first_p + section_break, 1)
files["word/document.xml"] = doc.encode("utf-8")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for n, data in files.items():
        z.writestr(n, data)

build.normalize_final_sections(out, ROOT / "gost-lab-report" / "templates" / "my" / "reference.docx")

with zipfile.ZipFile(out) as z:
    d = z.read("word/document.xml").decode("utf-8")
sects = re.findall(r'<w:sectPr[^>]*>.*?</w:sectPr>', d, re.S)
print("sections:", len(sects))
for i, s in enumerate(sects):
    refs = re.findall(r'w:type="(\w+)" r:id="(\w+)"', s)
    print(f"-- sect {i}: titlePg={'<w:titlePg' in s} "
          f"pgNumType={bool(re.search(r'<w:pgNumType', s))} footers={refs}")
