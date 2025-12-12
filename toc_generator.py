# -*- coding: utf-8 -*-
"""
toc_generator.py
- 封面使用 cover.png
- 封面副標再左移、放大 5pt、粗體
- 右下角文字放大 5pt、粗體
"""

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.utils import ImageReader
from reportlab.lib.colors import HexColor
from pathlib import Path
import re

PAGE_WIDTH, PAGE_HEIGHT = A4
LEFT_MARGIN = 60
BOTTOM_MARGIN = 80

PAGE_NO_X = PAGE_WIDTH - 60
DOT_END_X = PAGE_NO_X - 10

BASE_DIR = Path(__file__).parent
COVER_IMAGE = BASE_DIR / "cover.png"
INFO_IMAGE = BASE_DIR / "additionalinfo.png"

FONT_PATH = r"C:\Windows\Fonts\msjh.ttc"
pdfmetrics.registerFont(TTFont("msjh", FONT_PATH))
pdfmetrics.registerFont(TTFont("msjh-bold", FONT_PATH))

GREEN = HexColor("#3A9D7C")
BLACK = HexColor("#000000")


def parse_compile_date(text: str):
    m = re.search(r"(\d{3})年(\d{1,2})月", text)
    if not m:
        raise ValueError("格式錯誤，請輸入如：114年11月編製")

    year = int(m.group(1))
    month = int(m.group(2))

    if month == 1:
        display_year = year - 1
        display_month = 12
    else:
        display_year = year
        display_month = month - 1

    return f"{display_year}年{display_month}月", f"{year}年{month}月編製"


def draw_toc_header(c):
    c.setFont("msjh", 18)
    c.setFillColor(BLACK)
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 60, "臺 北 市 衛 生 統 計 摘 要 速 報")
    c.setFont("msjh", 22)
    c.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 100, "目　次")


def generate_toc_pdf(output_pdf, toc_items, compile_date_text):
    top_year_month, bottom_text = parse_compile_date(compile_date_text)
    c = canvas.Canvas(str(output_pdf), pagesize=A4)

    # =========================
    # 封面
    # =========================
    if COVER_IMAGE.exists():
        img = ImageReader(str(COVER_IMAGE))
        c.drawImage(img, 0, 0, PAGE_WIDTH, PAGE_HEIGHT, preserveAspectRatio=True, anchor="c")

    # 參考主標題左緣
    main_title = "衛生統計摘要速報"
    main_title_font_size = 36
    main_title_width = c.stringWidth(main_title, "msjh", main_title_font_size)
    main_title_left_x = PAGE_WIDTH / 2 - main_title_width / 2

    # 🔧【副標：左移 + 放大 + 粗體】
    subtitle_font_size = 21  # 原 16 + 5
    subtitle_x = main_title_left_x - 45   # 再往左移
    subtitle_y = PAGE_HEIGHT - 180

    c.setFont("msjh-bold", subtitle_font_size)
    c.setFillColor(BLACK)
    c.drawString(subtitle_x, subtitle_y, top_year_month)

    c.setFillColor(GREEN)
    c.drawString(
        subtitle_x + c.stringWidth(top_year_month + " ", "msjh-bold", subtitle_font_size),
        subtitle_y,
        "臺北市"
    )

    # 🔧【右下角：放大 + 粗體】
    footer_font_size = 17  # 原 12 + 5
    c.setFont("msjh-bold", footer_font_size)
    c.setFillColor(GREEN)
    c.drawRightString(PAGE_WIDTH - 40, 75, "臺北市政府衛生局")
    c.drawRightString(PAGE_WIDTH - 40, 50, bottom_text)

    c.setFillColor(BLACK)
    c.showPage()

    # =========================
    # 目錄
    # =========================
    draw_toc_header(c)
    c.setFont("msjh", 12)

    y = PAGE_HEIGHT - 150
    for item in toc_items:
        left_text = f"{item['index']}. {item['title']}"
        c.drawString(LEFT_MARGIN, y, left_text)

        text_width = c.stringWidth(left_text, "msjh", 12)
        dot_start_x = LEFT_MARGIN + text_width + 8
        dots = "." * int((DOT_END_X - dot_start_x) / c.stringWidth(".", "msjh", 12))

        c.drawString(dot_start_x, y, dots)
        c.drawRightString(PAGE_NO_X, y, str(item["page"]))

        y -= 22
        if y < BOTTOM_MARGIN + 40:
            c.showPage()
            draw_toc_header(c)
            c.setFont("msjh", 12)
            y = PAGE_HEIGHT - 150

    # additionalinfo.png
    if INFO_IMAGE.exists():
        img = ImageReader(str(INFO_IMAGE))
        iw, ih = img.getSize()
        img_w = PAGE_WIDTH - 2 * LEFT_MARGIN
        img_h = ih * (img_w / iw)

        if y - img_h < BOTTOM_MARGIN:
            c.showPage()
            draw_toc_header(c)
            y = PAGE_HEIGHT - 150

        c.drawImage(img, LEFT_MARGIN, y - img_h, img_w, img_h, mask="auto")

    c.showPage()
    c.save()
