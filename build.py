#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build.py — 產生「北歐接力精確體系」靜態網站

來源：
  原始資料/一方塊開叫.xlsx, 一梅花開叫.xlsx, 高花開叫.xlsx   -> 網站/叫序資料/*.html
  原始資料/3基本準則/第三章.md
  原始資料/4一方塊開叫/一方塊開叫.md
  原始資料/5高花開叫/高花開叫.md
  原始資料/7一梅花開叫/一梅花開叫.md                          -> 網站/原始書籍檔案/*.html

重新執行本腳本即可依來源檔案的最新內容重新產生整個 網站/ 目錄。
"""
import os
import re
import html
import shutil

import openpyxl
import markdown as md

ROOT = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(ROOT, "原始資料")
OUT = os.path.join(ROOT, "網站")

SUIT_RED = set("♥♦")
SUIT_BLACK = set("♠♣")


# --------------------------------------------------------------------------
# 共用版型
# --------------------------------------------------------------------------

NAV_ITEMS = [
    ("叫序資料", [
        ("一方塊開叫", "叫序資料/一方塊開叫.html"),
        ("一梅花開叫", "叫序資料/一梅花開叫.html"),
        ("高花開叫", "叫序資料/高花開叫.html"),
    ]),
    ("原始書籍檔案", [
        ("第三章", "原始書籍檔案/第三章.html"),
        ("一方塊開叫", "原始書籍檔案/一方塊開叫.html"),
        ("高花開叫", "原始書籍檔案/高花開叫.html"),
        ("一梅花開叫", "原始書籍檔案/一梅花開叫.html"),
    ]),
]


def suit_colorize(text):
    """把文字中的 ♥♦ 轉紅、♠♣ 轉黑（沿用速查頁的 .r/.b 慣例），並保留換行。"""
    out = []
    for ch in text:
        if ch in SUIT_RED:
            out.append('<span class="s r">%s</span>' % ch)
        elif ch in SUIT_BLACK:
            out.append('<span class="s b">%s</span>' % ch)
        elif ch == "\n":
            out.append("<br>")
        else:
            out.append(html.escape(ch))
    return "".join(out)


def render_nav(root_prefix, current_href=None):
    parts = []
    parts.append('<nav class="site-nav">')
    parts.append('<div class="nav-inner">')
    parts.append('<a class="nav-brand" href="%sindex.html">北歐接力精確體系</a>' % root_prefix)
    parts.append('<button class="nav-toggle" aria-label="開啟選單" type="button">'
                  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">'
                  '<path d="M3 6h18M3 12h18M3 18h18"/></svg></button>')
    parts.append('<ul class="nav-links" role="menubar">')
    for label, children in NAV_ITEMS:
        parts.append('<li class="nav-item has-menu" role="none">')
        parts.append('<button class="nav-link" type="button">%s <span class="car">▾</span></button>' % label)
        parts.append('<ul class="nav-menu">')
        for sub_label, sub_href in children:
            active = " is-active" if sub_href == current_href else ""
            parts.append('<li><a class="%s" href="%s%s">%s</a></li>' %
                          (active.strip() or "", root_prefix + sub_href, "", sub_label))
        parts.append('</ul></li>')
    parts.append('</ul>')
    parts.append('</div></nav>')
    return "".join(parts)


PAGE_TEMPLATE = """<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<meta name="description" content="{description}">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@500;700&family=Noto+Sans+TC:wght@400;500;700&family=Noto+Serif+TC:wght@600;700&display=swap">
<link rel="stylesheet" href="{root}assets/style.css">
</head>
<body>
{nav}
<div class="wrap">
{body}
</div>
<script src="{root}assets/site.js"></script>
</body>
</html>
"""


def render_page(title, description, root_prefix, current_href, body_html):
    return PAGE_TEMPLATE.format(
        title=html.escape(title),
        description=html.escape(description),
        root=root_prefix,
        nav=render_nav(root_prefix, current_href),
        body=body_html,
    )


def page_head(eyebrow, h1, meta_lines, lede_html, crumb_html=""):
    meta = "<br>".join(meta_lines)
    return """
{crumb}
<header class="page-head">
  <div class="mast">
    <div>
      <div class="eyebrow">{eyebrow}</div>
      <h1>{h1}</h1>
    </div>
    <div class="mast-meta">{meta}</div>
  </div>
  <p class="lede">{lede}</p>
</header>
""".format(crumb=crumb_html, eyebrow=eyebrow, h1=h1, meta=meta, lede=lede_html)


def write_file(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    print("wrote", os.path.relpath(path, ROOT))


# --------------------------------------------------------------------------
# xlsx -> HTML（叫序資料）
# --------------------------------------------------------------------------

def argb_to_hex(argb):
    if not argb or len(argb) < 6:
        return None
    rgb = argb[-6:]
    if rgb.upper() == "000000" and len(argb) == 8 and argb.startswith("00"):
        # 00 開頭且顏色為全黑，openpyxl 對「無填色」也會回報這個值，視為無填色
        return None
    return "#" + rgb


def cell_style(cell, is_header_zone):
    styles = []
    classes = []
    font = cell.font
    fill = cell.fill

    bg = None
    if fill and fill.fgColor and fill.fill_type == "solid":
        fg = fill.fgColor
        if fg.type == "rgb":
            bg = argb_to_hex(fg.rgb)
    if bg:
        styles.append("background:%s" % bg)

    color = None
    if font and font.color:
        fc = font.color
        if fc.type == "rgb" and fc.rgb:
            color = argb_to_hex(fc.rgb)
        elif fc.type == "theme" and fc.theme == 1:
            color = "#FFFFFF"
    if color:
        styles.append("color:%s" % color)

    if font and font.bold:
        styles.append("font-weight:700")

    align = cell.alignment.horizontal if cell.alignment else None
    if align in ("center",):
        styles.append("text-align:center")
    elif align in ("right",):
        styles.append("text-align:right")

    return styles


def sheet_to_html(ws):
    merged = {}
    covered = set()
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.bounds
        merged[(min_row, min_col)] = (max_row - min_row + 1, max_col - min_col + 1)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                if (r, c) != (min_row, min_col):
                    covered.add((r, c))

    rows_html = []
    max_col = ws.max_column
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        cells_html = []
        any_content = False
        for cell in row:
            r, c = cell.row, cell.column
            if (r, c) in covered:
                continue
            value = cell.value
            text = "" if value is None else str(value)
            if text.strip():
                any_content = True
            span = merged.get((r, c))
            attrs = []
            if span:
                rowspan, colspan = span
                if rowspan > 1:
                    attrs.append('rowspan="%d"' % rowspan)
                if colspan > 1:
                    attrs.append('colspan="%d"' % colspan)
            styles = cell_style(cell, is_header_zone=(r == 1))
            css_class = "banner" if (span and span[1] >= max_col and max_col > 1) else ""
            style_attr = (' style="%s"' % ";".join(styles)) if styles else ""
            class_attr = (' class="%s"' % css_class) if css_class else ""
            content = suit_colorize(text) if text else "&nbsp;"
            cells_html.append("<td%s%s%s>%s</td>" % (class_attr, style_attr, " ".join([""] + attrs), content))
        if not any_content:
            continue
        rows_html.append("<tr>%s</tr>" % "".join(cells_html))

    return '<div class="table-scroll"><table class="xlsx-tbl"><tbody>%s</tbody></table></div>' % "".join(rows_html)


def build_xlsx_page(xlsx_name, title, eyebrow, lede):
    path = os.path.join(SRC, xlsx_name)
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = wb.worksheets

    tabs = []
    panels = []
    for i, ws in enumerate(sheets):
        sid = "sheet-%d" % (i + 1)
        active = " is-active" if i == 0 else ""
        tabs.append('<button class="sheet-tab%s" data-target="%s">%s</button>' %
                     (active, sid, html.escape(ws.title)))
        panels.append('<div class="sheet-panel%s" id="%s">%s</div>' %
                       (active, sid, sheet_to_html(ws)))

    body = page_head(
        eyebrow, title,
        ["%d 個工作表" % len(sheets), "來源：%s" % xlsx_name],
        lede,
        crumb_html='<div class="crumb"><a href="../index.html">首頁</a> ／ 叫序資料 ／ %s</div>' % title,
    )
    body += '<section>'
    body += '<div class="sheet-tabs" data-sheet-tabs="#%s-panels">%s</div>' % (xlsx_name, "".join(tabs))
    body += '<div id="%s-panels">%s</div>' % (xlsx_name, "".join(panels))
    body += '</section>'
    body += '<footer class="site-footer">資料整理自《北歐接力精確體系》，原始表格見「原始書籍檔案」分頁。</footer>'

    html_out = render_page(
        title="%s ｜ 叫序資料 ｜ 北歐接力精確體系" % title,
        description="%s 的完整叫序資料表" % title,
        root_prefix="../",
        current_href="叫序資料/%s.html" % title,
        body_html=body,
    )
    write_file(os.path.join(OUT, "叫序資料", "%s.html" % title), html_out)


# --------------------------------------------------------------------------
# markdown -> HTML（原始書籍檔案）
# --------------------------------------------------------------------------

FRONTMATTER_RE = re.compile(r"^---\s*\n.*?\n---\s*\n", re.S)


def strip_frontmatter(text):
    return FRONTMATTER_RE.sub("", text, count=1)


def md_to_html(md_path, image_dir_href):
    with open(md_path, "r", encoding="utf-8") as f:
        text = f.read()
    text = strip_frontmatter(text)

    md_conv = md.Markdown(extensions=["tables", "sane_lists"])
    body = md_conv.convert(text)

    # 圖片路徑改指到 assets/images/<章節>/
    def fix_img(m):
        pre, src, post = m.group(1), m.group(2), m.group(3)
        if src.startswith("http://") or src.startswith("https://"):
            return m.group(0)
        fname = os.path.basename(src)
        return '<figure><img%ssrc="%s%s"%s></figure>' % (pre, image_dir_href, fname, post)

    body = re.sub(r'<img([^>]*?)src="([^"]+)"([^>]*?)>', fix_img, body)
    # markdown 會把單獨成行的圖片包在 <p> 裡，figure 是區塊元素，拆掉外層 <p> 避免非法巢狀
    body = re.sub(r"<p>\s*(<figure>.*?</figure>)\s*</p>", r"\1", body, flags=re.S)

    # 幫表格加水平捲動包裹（RWD）
    body = body.replace("<table>", '<div class="table-scroll"><table>')
    body = body.replace("</table>", "</table></div>")

    body = colorize_suits_html(body)

    return body


def colorize_suits_html(html_body):
    """花色符號著色，只處理標籤之間的純文字節點，避免污染 alt= 等屬性值。"""
    parts = re.split(r"(<[^>]+>)", html_body)
    suit_map = {
        "♠": '<span class="s b">♠</span>',
        "♣": '<span class="s b">♣</span>',
        "♥": '<span class="s r">♥</span>',
        "♦": '<span class="s r">♦</span>',
    }
    for i, part in enumerate(parts):
        if part.startswith("<"):
            continue
        for suit, replacement in suit_map.items():
            part = part.replace(suit, replacement)
        parts[i] = part
    return "".join(parts)


BOOK_DOCS = [
    dict(
        md_file="3基本準則/第三章.md",
        slug="第三章",
        title="第三章 北歐梅花的問叫和基本準則",
        image_dir="第三章",
        lede="接力系統的正式定義：五種問叫工具（Sidestep／Splinter／54Pick-up／Six-shooter／CRASH）與四條硬規則的原始出處。",
    ),
    dict(
        md_file="4一方塊開叫/一方塊開叫.md",
        slug="一方塊開叫",
        title="第四章 一方塊開叫及其發展",
        image_dir="一方塊開叫",
        lede="1♦ 開叫的完整定義與 1♦–1♥ 後的接力叫牌發展，含原始截圖對照還原。",
    ),
    dict(
        md_file="5高花開叫/高花開叫.md",
        slug="高花開叫",
        title="第五章 一階高花開叫及其發展",
        image_dir="高花開叫",
        lede="1♠／1♥ 開叫後的第一應叫與 1NT 後的接力主幹發展，含原始截圖對照還原。",
    ),
    dict(
        md_file="7一梅花開叫/一梅花開叫.md",
        slug="一梅花開叫",
        title="第七章 一梅花開叫及其發展",
        image_dir="一梅花開叫",
        lede="1♣ 開叫後的應叫系統與各分支接力發展，含原始截圖對照還原。",
    ),
]


def build_book_pages():
    toc = "".join(
        '<a href="%s.html">%s</a>' % (d["slug"], d["title"].split(" ", 1)[-1] if " " in d["title"] else d["title"])
        for d in BOOK_DOCS
    )
    for d in BOOK_DOCS:
        md_path = os.path.join(SRC, d["md_file"])
        content = md_to_html(md_path, "../assets/images/%s/" % d["image_dir"])

        body = page_head(
            "原始書籍檔案", d["title"],
            ["Markdown 還原版", "來源：%s" % d["md_file"]],
            d["lede"],
            crumb_html='<div class="crumb"><a href="../index.html">首頁</a> ／ 原始書籍檔案 ／ %s</div>' % d["title"],
        )
        body += '<div class="toc-links">%s</div>' % toc
        body += '<article class="article">%s</article>' % content
        body += '<footer class="site-footer">整理自《北歐接力精確體系》原始剪藏，已比對截圖還原花色符號與表格結構。</footer>'

        html_out = render_page(
            title="%s ｜ 原始書籍檔案 ｜ 北歐接力精確體系" % d["title"],
            description=d["lede"],
            root_prefix="../",
            current_href="原始書籍檔案/%s.html" % d["slug"],
            body_html=body,
        )
        write_file(os.path.join(OUT, "原始書籍檔案", "%s.html" % d["slug"]), html_out)


# --------------------------------------------------------------------------
# 首頁
# --------------------------------------------------------------------------

def build_index():
    body = page_head(
        "北歐接力精確體系",
        "接力叫序資料庫",
        ["3 份叫序資料表", "4 篇原始書籍章節"],
        "整理自《北歐接力精確體系》一書，收錄一方塊、一梅花、高花三種開叫的完整接力叫序資料，"
        "並保留原始書籍章節的 Markdown 還原版供對照查閱。",
    )

    body += """
<section>
  <div class="sec-head"><h2>叫序資料</h2><span class="tag">依開叫花色整理的接力叫序表</span></div>
  <div class="card-grid">
    <a class="doc-card" href="叫序資料/一方塊開叫.html">
      <div class="k">Diamonds</div>
      <h3>一方塊開叫</h3>
      <p>1♦ 開叫、1♦–1♥ 接力主幹與各分支發展，共 12 個工作表。</p>
    </a>
    <a class="doc-card" href="叫序資料/一梅花開叫.html">
      <div class="k">Clubs</div>
      <h3>一梅花開叫</h3>
      <p>1♣ 開叫的應叫系統與各分支接力發展，共 13 個工作表。</p>
    </a>
    <a class="doc-card" href="叫序資料/高花開叫.html">
      <div class="k">Majors</div>
      <h3>高花開叫</h3>
      <p>一階高花開叫、1NT 後接力主幹與各分支，共 11 個工作表。</p>
    </a>
  </div>
</section>

<section>
  <div class="sec-head"><h2>原始書籍檔案</h2><span class="tag">Markdown 還原版，含原書截圖對照</span></div>
  <div class="card-grid">
    <a class="doc-card" href="原始書籍檔案/第三章.html">
      <div class="k">Chapter 03</div>
      <h3>第三章：問叫與基本準則</h3>
      <p>接力系統的正式定義：五種問叫工具與四條硬規則。</p>
    </a>
    <a class="doc-card" href="原始書籍檔案/一方塊開叫.html">
      <div class="k">Chapter 04</div>
      <h3>一方塊開叫及其發展</h3>
      <p>1♦ 開叫的完整定義與接力叫牌發展。</p>
    </a>
    <a class="doc-card" href="原始書籍檔案/高花開叫.html">
      <div class="k">Chapter 05</div>
      <h3>一階高花開叫及其發展</h3>
      <p>高花開叫的第一應叫與接力主幹發展。</p>
    </a>
    <a class="doc-card" href="原始書籍檔案/一梅花開叫.html">
      <div class="k">Chapter 07</div>
      <h3>一梅花開叫及其發展</h3>
      <p>1♣ 開叫後的應叫系統與各分支接力發展。</p>
    </a>
  </div>
</section>

<div class="external-card">
  <p><strong>接力工具速查</strong> — 第三章五種問叫工具（Sidestep／Splinter／54Pick-up／Six-shooter／CRASH）的濃縮速查頁。</p>
  <a href="../原始資料/北歐接力工具速查.html">開啟速查頁</a>
</div>
"""
    body += '<footer class="site-footer">整理自《北歐接力精確體系》。以上資料僅供個人叫牌系統學習與對照使用。</footer>'

    html_out = render_page(
        title="北歐接力精確體系 ｜ 接力叫序資料庫",
        description="整理自《北歐接力精確體系》的接力叫序資料庫與原始書籍檔案",
        root_prefix="",
        current_href=None,
        body_html=body,
    )
    write_file(os.path.join(OUT, "index.html"), html_out)


# --------------------------------------------------------------------------

def main():
    build_index()
    build_xlsx_page("一方塊開叫.xlsx", "一方塊開叫", "叫序資料 · 一方塊開叫",
                     "1♦ 開叫、1♦–1♥ 接力主幹（A–J 分支）與非接力叫牌、示例牌張、示例叫序。")
    build_xlsx_page("一梅花開叫.xlsx", "一梅花開叫", "叫序資料 · 一梅花開叫",
                     "1♣ 開叫後的應叫系統、1♦ 後開叫者再叫、1NT／2♣／2♦／2♥／2♠／2NT 各系統。")
    build_xlsx_page("高花開叫.xlsx", "高花開叫", "叫序資料 · 高花開叫",
                     "一階高花開叫的第一應叫、1NT 後接力主幹（A–E 分支）與示例牌張、示例叫序。")
    build_book_pages()


if __name__ == "__main__":
    main()
