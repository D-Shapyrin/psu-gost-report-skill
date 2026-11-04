"""CI-раннер: собирает тестовые документы и проверяет результат.

Использование:
    python tests/ci_check.py --expect-toc       # LibreOffice есть: оглавление должно быть посчитано
    python tests/ci_check.py --expect-degraded  # LibreOffice нет: сборка обязана пройти с честным warning'ом

Собирает все tests/*_test.md (профиль вуза берётся из YAML каждого файла),
гоняет scripts/verify.py с тем же профилем, затем проверяет режим TOC.
Выход 0 — всё прошло, 1 — есть провалы.
"""

import re
import subprocess
import sys
import zipfile
from pathlib import Path

TESTS_DIR = Path(__file__).resolve().parent
SKILL_DIR = TESTS_DIR.parent / "gost-lab-report"
BUILD = SKILL_DIR / "scripts" / "build.py"
VERIFY = SKILL_DIR / "scripts" / "verify.py"
OUT_DIR = TESTS_DIR / "out"


def run(args, **kw):
    print(f"\n$ {' '.join(str(a) for a in args)}", flush=True)
    proc = subprocess.run(
        [sys.executable, *map(str, args)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=TESTS_DIR, **kw,
    )
    sys.stdout.write(proc.stdout)
    sys.stderr.write(proc.stderr)
    return proc


def front_matter(md_path):
    """Мини-парсер YAML front matter: key -> value (строки как есть)."""
    text = md_path.read_text(encoding="utf-8-sig")
    m = re.match(r"^---\s*\n(.*?)\n---", text, re.DOTALL)
    meta = {}
    if not m:
        return meta
    for line in m.group(1).splitlines():
        if ":" in line:
            key, _, value = line.partition(":")
            meta[key.strip()] = value.strip().strip('"').strip("'")
    return meta


def document_xml(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        return z.read("word/document.xml").decode("utf-8", errors="replace")


HEADING_PARA_RE = re.compile(r"<w:p\b[^/>]*?>(?:(?!</w:p>).)*?</w:p>", re.DOTALL)


def heading_texts(docx_path):
    """Тексты заголовков уровня 1-3 в порядке документа.

    Уровень определяется по имени стиля в styles.xml: профиль ПГУ хранит
    заголовки под авто-id («1», «2», «3») — именно на этом уже ломалась
    нумерация (регресс ловил только человек).
    """
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8", errors="replace")
        styles = (
            z.read("word/styles.xml").decode("utf-8", errors="replace")
            if "word/styles.xml" in z.namelist()
            else ""
        )
    levels = {}
    for block in re.findall(r"<w:style [^>]*>.*?</w:style>", styles, re.DOTALL):
        sid = re.search(r'w:styleId="([^"]*)"', block)
        name = re.search(r'<w:name w:val="heading\s*([1-9])"', block, re.I)
        if sid and name:
            levels[sid.group(1)] = int(name.group(1))
    texts = []
    for p in HEADING_PARA_RE.findall(xml):
        sid = re.search(r'w:pStyle w:val="([^"]*)"', p)
        if not sid:
            continue
        level = levels.get(sid.group(1))
        if level is None:
            m = re.fullmatch(r"Heading([1-9])", sid.group(1))
            level = int(m.group(1)) if m else None
        if level is not None:
            texts.append("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)))
    return texts


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in ("--expect-toc", "--expect-degraded"):
        print(__doc__)
        return 1

    expect_toc = sys.argv[1] == "--expect-toc"
    OUT_DIR.mkdir(exist_ok=True)
    failures = []

    tests = sorted(TESTS_DIR.glob("*_test.md"))
    if not tests:
        print("No tests/*_test.md found")
        return 1

    results = {}
    for md in tests:
        name = md.stem
        meta = front_matter(md)
        profile = meta.get("profile", "my") or "my"
        out = OUT_DIR / f"{name}.docx"
        proc = run([BUILD, md, out])
        results[name] = (proc, meta)
        if proc.returncode != 0:
            failures.append(f"build {name}: exit {proc.returncode}")
            continue
        vproc = run([VERIFY, out, f"--profile={profile}"])
        if vproc.returncode != 0:
            failures.append(f"verify {name}: exit {vproc.returncode}")

        # number_headings=true меняет только заголовки: первое слово каждого
        # должно стать номером «N» или «N.N» (структурные элементы вроде
        # «ВВЕДЕНИЕ» не нумеруются — их отсутствие номера не претензия).
        if str(meta.get("number_headings", "")).lower() in ("true", "yes", "1"):
            heads = heading_texts(out)
            numbered = [h for h in heads if re.match(r"^\d+(\.\d+)*\s", h)]
            if heads and not numbered:
                failures.append(
                    f"numbering {name}: ни один заголовок не пронумерован "
                    f"({heads[:3]}...) — проверь number_headings/_heading_style_ids"
                )
            print(f"numbering {name}: {len(numbered)}/{len(heads)} headings numbered")

    # Режим оглавления проверяем на документах с add_toc=true
    toc_docs = [
        (name, proc, meta)
        for name, (proc, meta) in results.items()
        if str(meta.get("add_toc", "")).lower() in ("true", "yes", "1")
    ]
    for name, proc, meta in toc_docs:
        out = OUT_DIR / f"{name}.docx"
        if not out.exists():
            continue  # провал сборки зафиксирован выше
        xml = document_xml(out)
        # LibreOffice пишет посчитанный TOC гиперссылками на якоря __RefHeading___Toc…;
        # PAGEREF — на случай, если оглавление посчитал Word
        toc_computed = "__RefHeading" in xml or "PAGEREF" in xml
        log = proc.stdout + proc.stderr
        if expect_toc:
            if not toc_computed:
                failures.append(
                    f"TOC {name}: маркеры не найдены — оглавление не посчитано, "
                    "хотя LibreOffice ожидался"
                )
        else:
            if toc_computed:
                failures.append(
                    f"TOC {name}: оглавление посчитано — ожидалась деградация без LibreOffice"
                )
            if "LibreOffice" not in log:
                failures.append(
                    f"degradation {name}: сборка без LibreOffice не предупредила "
                    "пользователя (нет warning'а)"
                )

    print("\n=== CI SUMMARY ===")
    if failures:
        for f in failures:
            print(f"[FAIL] {f}")
        return 1
    mode = "TOC computed" if expect_toc else "degraded honestly (warning shown)"
    print(f"[PASS] {len(tests)} build(s) + verify + mode: {mode}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
