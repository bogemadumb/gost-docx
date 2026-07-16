#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gost-docx — движок сборки документов по ГОСТ / Положению СПбГЭУ:
markdown-исходник + meta.json -> docx + pdf.

Текст ЖИВЁТ ОТДЕЛЬНО от кода (в .md). Оформление зашито по стандарту
(Times New Roman 14, интервал 1,5, поля 30/10/20/20 мм, красная строка 1,25,
выравнивание по ширине, нумерация страниц, содержание с точками-лидерами).

Использование:
    python3 engine.py <исходник.md> <meta.json>
    python3 engine.py <исходник.md> <meta.json> --once   # без двухпроходного содержания (быстрый черновик)

Двухпроходная сборка (по умолчанию): docx -> pdf -> определить номера страниц -> пересобрать
содержание с реальными номерами -> финальный docx+pdf. В конце печатает отчёт по объёму
(страниц всего vs целевой диапазон из meta).

Требования:
    pip install python-docx pypdf
    LibreOffice (soffice) — для конвертации docx -> pdf. Путь берётся из переменной
    окружения SOFFICE_PATH, либо ищется в стандартных местах и в PATH.

Формат исходника (.md) — см. example/rabota.md. Кратко:
    # ЗАГОЛОВОК РАЗДЕЛА        -> H1 (если начинается с цифры — станет ЗАГЛАВНЫМИ; каждый H1 с новой страницы)
    ## 1.1 Подраздел          -> H2
    обычная строка            -> абзац (красная строка, по ширине); [1; 2] — ссылки на источники
    - пункт                   -> маркированный список
    @table Таблица N, подпись -> начало таблицы; строки «ячейка | ячейка»; @endtable — конец
    @image путь | ширина_см | Рисунок N, подпись
    @pagebreak                -> принудительный разрыв страницы
    @sources                  -> вставить нумерованный список источников из meta.sources (порядок — как в meta)
Пустая строка разделяет абзацы. Заголовок СОДЕРЖАНИЕ строится движком, в .md его не пишем.
"""
import os, re, sys, json, shutil, subprocess

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ------------------------------------------------------------------ поиск LibreOffice (soffice)
def find_soffice():
    """Путь к soffice: env SOFFICE_PATH -> стандартные места -> PATH -> 'soffice'."""
    env = os.environ.get("SOFFICE_PATH") or os.environ.get("SOFFICE")
    if env and os.path.exists(env):
        return env
    candidates = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",  # macOS
        "/usr/bin/soffice",                                       # Linux (пакет)
        "/usr/local/bin/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",      # Windows
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    found = shutil.which("soffice") or shutil.which("libreoffice")
    return found or "soffice"


# ------------------------------------------------------------------ разбор исходника
def parse_source(path):
    """Возвращает список блоков: ('h1'|'h2'|'para'|'bullet'|'table'|'image'|'pagebreak'|'sources', data)."""
    blocks = []
    lines = open(path, encoding="utf-8").read().split("\n")
    i = 0
    while i < len(lines):
        ln = lines[i].rstrip()
        s = ln.strip()
        if not s:
            i += 1; continue
        if s.startswith("@table"):
            cap = s[len("@table"):].strip()
            rows = []
            i += 1
            while i < len(lines) and lines[i].strip() != "@endtable":
                r = lines[i].rstrip()
                if r.strip():
                    rows.append([c.strip() for c in r.split("|")])
                i += 1
            i += 1  # съесть @endtable
            blocks.append(("table", {"caption": cap, "rows": rows}))
            continue
        if s.startswith("@image"):
            parts = [p.strip() for p in s[len("@image"):].split("|")]
            path_img = parts[0]
            width = float(parts[1]) if len(parts) > 1 and parts[1] else 12.0
            cap = parts[2] if len(parts) > 2 else None
            blocks.append(("image", {"path": path_img, "width": width, "caption": cap}))
            i += 1; continue
        if s == "@pagebreak":
            blocks.append(("pagebreak", None)); i += 1; continue
        if s == "@sources":
            blocks.append(("sources", None)); i += 1; continue
        if s.startswith("## "):
            blocks.append(("h2", s[3:].strip())); i += 1; continue
        if s.startswith("# "):
            blocks.append(("h1", s[2:].strip())); i += 1; continue
        if s.startswith("- "):
            blocks.append(("bullet", s[2:].strip())); i += 1; continue
        # обычный абзац — склеиваем до пустой строки/директивы
        buf = [s]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if not nxt or nxt.startswith(("#", "-", "@")):
                break
            buf.append(nxt); i += 1
        blocks.append(("para", " ".join(buf)))
    return blocks

def is_numbered(t):
    return bool(re.match(r"^\d", t.strip()))

def num_dot(t):
    return re.sub(r"^(\d+(?:\.\d+)*)\s", r"\1. ", t)

# ------------------------------------------------------------------ стили ГОСТ / СПбГЭУ
def new_doc():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Times New Roman"; st.font.size = Pt(14)
    st._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    pf = st.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(0); pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.first_line_indent = Cm(1.25)
    for hname in ("Heading 1", "Heading 2"):
        h = doc.styles[hname]
        h.font.name = "Times New Roman"; h.font.size = Pt(14); h.font.bold = True
        h.font.color.rgb = RGBColor(0, 0, 0)
        h._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        h.paragraph_format.first_line_indent = Cm(0)
        h.paragraph_format.space_before = Pt(12); h.paragraph_format.space_after = Pt(6)
        h.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    sec = doc.sections[0]
    sec.page_height = Cm(29.7); sec.page_width = Cm(21)
    sec.left_margin = Cm(3); sec.right_margin = Cm(1); sec.top_margin = Cm(2); sec.bottom_margin = Cm(2)
    return doc, sec

def set_start_page(section, start):
    sectPr = section._sectPr
    pgNum = sectPr.find(qn("w:pgNumType"))
    if pgNum is None:
        pgNum = OxmlElement("w:pgNumType"); sectPr.append(pgNum)
    pgNum.set(qn("w:start"), str(start))

def add_page_field(p):
    run = p.add_run()
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end")
    run._r.append(b); run._r.append(instr); run._r.append(e)

def footer_page_numbers(sec):
    fpar = sec.footer.paragraphs[0]
    fpar.alignment = WD_ALIGN_PARAGRAPH.CENTER; fpar.paragraph_format.first_line_indent = Cm(0)
    add_page_field(fpar)
    for r in fpar.runs:
        r.font.name = "Times New Roman"; r.font.size = Pt(14)

def borders(table):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "4"); e.set(qn("w:space"), "0"); e.set(qn("w:color"), "000000")
        b.append(e)
    tblPr.append(b)

def set_cell(cell, text, bold=False, align="left", size=12):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    r = p.add_run(text); r.font.name = "Times New Roman"; r.font.size = Pt(size); r.font.bold = bold

def fmt_caption(t):
    m = re.match(r"^((?:Таблица|Рисунок)\s+[\dА-ЯA-Z.]+),\s*(.*)$", t)
    if m:
        rest = m.group(2); rest = rest[:1].upper() + rest[1:] if rest else rest
        return f"{m.group(1)} – {rest}"
    return t

# ------------------------------------------------------------------ сборка тела
def build(blocks, meta, toc_pages, out_docx):
    doc, sec = new_doc()
    toc_start = int(meta.get("toc_start_page", 3))
    set_start_page(sec, toc_start)
    footer_page_numbers(sec)

    # собрать оглавление из заголовков
    toc_entries = []
    for kind, data in blocks:
        if kind == "h1":
            toc_entries.append((num_dot(data).upper() if is_numbered(data) else data, 1))
        elif kind == "h2":
            toc_entries.append((num_dot(data), 2))

    def h1(t):
        disp = num_dot(t)
        disp = disp.upper() if is_numbered(t) else disp
        p = doc.add_paragraph(disp, style="Heading 1"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    def h2(t):
        p = doc.add_paragraph(num_dot(t), style="Heading 2"); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
    def para(t):
        doc.add_paragraph(t)
    def bullet(t):
        p = doc.add_paragraph(t, style="List Bullet")
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        for r in p.runs:
            r.font.name = "Times New Roman"; r.font.size = Pt(14)
    def table_caption(t):
        p = doc.add_paragraph(fmt_caption(t)); p.paragraph_format.first_line_indent = Cm(0); p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p.runs:
            r.font.size = Pt(12); r.font.italic = True
    def img_caption(t):
        p = doc.add_paragraph(fmt_caption(t)); p.paragraph_format.first_line_indent = Cm(0); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        for r in p.runs:
            r.font.size = Pt(12); r.font.italic = True
    def make_table(rows):
        if not rows:
            return
        ncol = max(len(r) for r in rows)
        t = doc.add_table(rows=0, cols=ncol); t.alignment = WD_ALIGN_PARAGRAPH.CENTER; borders(t)
        for ri, row in enumerate(rows):
            cells = t.add_row().cells
            for ci in range(ncol):
                val = row[ci] if ci < len(row) else ""
                if ri == 0:
                    set_cell(cells[ci], val, bold=True, align="center")
                else:
                    set_cell(cells[ci], val, align=("left" if ci == 0 else "center"))
    def add_image(path_img, width, cap):
        if not os.path.exists(path_img):
            para(f"[нет файла изображения: {path_img}]")
        else:
            p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent = Cm(0)
            p.add_run().add_picture(path_img, width=Cm(width))
        if cap:
            img_caption(cap)

    # ---- СОДЕРЖАНИЕ ----
    ct = doc.add_paragraph("СОДЕРЖАНИЕ"); ct.alignment = WD_ALIGN_PARAGRAPH.CENTER; ct.paragraph_format.first_line_indent = Cm(0)
    for r in ct.runs:
        r.font.bold = True
    for disp, lvl in toc_entries:
        page = toc_pages.get(disp, "")
        tp = doc.add_paragraph(); tp.paragraph_format.first_line_indent = Cm(0)
        tp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        tp.paragraph_format.left_indent = Cm(0.75 if lvl == 2 else 0)
        tp.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        tp.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        run = tp.add_run(f"{disp}\t{page}")
        run.font.name = "Times New Roman"; run.font.size = Pt(14); run.font.bold = (lvl == 1)
    doc.add_page_break()

    # ---- ТЕЛО ----
    first_h1 = True
    for kind, data in blocks:
        if kind == "h1":
            if not first_h1:
                doc.add_page_break()
            first_h1 = False
            h1(data)
        elif kind == "h2":
            h2(data)
        elif kind == "para":
            para(data)
        elif kind == "bullet":
            bullet(data)
        elif kind == "table":
            table_caption(data["caption"]); make_table(data["rows"])
        elif kind == "image":
            add_image(data["path"], data["width"], data["caption"])
        elif kind == "pagebreak":
            doc.add_page_break()
        elif kind == "sources":
            for idx, src in enumerate(meta.get("sources", []), 1):
                p = doc.add_paragraph(f"{idx}. {src}")
                p.paragraph_format.first_line_indent = Cm(1.25)
    doc.save(out_docx)
    return toc_entries

# ------------------------------------------------------------------ конвертация и подсчёт страниц
def to_pdf(docx_path, out_dir):
    subprocess.run([find_soffice(), "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                   check=True, capture_output=True)
    return os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")

def detect_pages(pdf_path, toc_entries, toc_start):
    from pypdf import PdfReader
    r = PdfReader(pdf_path)
    pages = [re.sub(r"\s+", " ", (p.extract_text() or "")) for p in r.pages]
    result = {}
    for disp, lvl in toc_entries:
        needle = re.sub(r"\s+", " ", disp)[:28]  # начало заголовка — как в содержании
        found = ""
        for i, t in enumerate(pages):
            if i == 0:
                continue
            if needle in t:
                found = i + toc_start
                break
        result[disp] = found
    return result, len(pages)

# ------------------------------------------------------------------ main
def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    once = "--once" in sys.argv
    if len(args) < 2:
        print(__doc__); sys.exit(1)
    src, meta_path = args[0], args[1]
    meta = json.load(open(meta_path, encoding="utf-8"))
    blocks = parse_source(src)

    out_dir = os.path.expanduser(meta["output_dir"])
    os.makedirs(out_dir, exist_ok=True)
    out_docx = os.path.join(out_dir, meta["output_name"])
    toc_start = int(meta.get("toc_start_page", 3))

    # проход 1: без номеров страниц
    toc_entries = build(blocks, meta, {}, out_docx)
    pdf = to_pdf(out_docx, out_dir)

    total = None
    if not once:
        toc_pages, total = detect_pages(pdf, toc_entries, toc_start)
        missed = [d for d, _ in toc_entries if not toc_pages.get(d)]
        # проход 2: с реальными номерами
        build(blocks, meta, toc_pages, out_docx)
        pdf = to_pdf(out_docx, out_dir)
        # финальный подсчёт
        _, total = detect_pages(pdf, toc_entries, toc_start)
        if missed:
            print(f"ВНИМАНИЕ: не найдены страницы для {len(missed)} заголовков (проверь совпадение текста): {missed[:5]}")

    print(f"Готово:\n  docx: {out_docx}\n  pdf:  {pdf}")
    if total is not None:
        tgt = meta.get("target_pages")
        note = ""
        if tgt:
            lo, hi = tgt
            if total < lo:
                note = f"  — МЕНЬШЕ целевого ({lo}-{hi}), добавить объём"
            elif total > hi:
                note = f"  — БОЛЬШЕ целевого ({lo}-{hi}), сократить"
            else:
                note = f"  — в целевом диапазоне ({lo}-{hi}) ОК"
        print(f"Объём: {total} стр.{note}")

if __name__ == "__main__":
    main()
