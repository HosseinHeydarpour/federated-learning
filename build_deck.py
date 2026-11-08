# -*- coding: utf-8 -*-
"""Persian seminar deck rebuilt on the light 'Technology Pitch Deck' visual system."""

import os
import re
import shutil
import zipfile
import datetime
from lxml import etree
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import CategoryChartData
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
A = lambda name: os.path.join(HERE, "assets", name)

# ---------------------------------------------------------------- tokens
FA = "B Nazanin"
FA_TITLE = "B Titr"
EN = "Michroma"
EN_BODY = "Segoe UI"
ICON_FONT = "Segoe MDL2 Assets"

BLUE      = RGBColor(0x2F, 0x5D, 0xA8)
BLUE_D    = RGBColor(0x1E, 0x3F, 0x73)
BLUE_MID  = RGBColor(0x4B, 0x77, 0xC4)
BLUE_PALE = RGBColor(0xE8, 0xF1, 0xFF)
LAVENDER  = RGBColor(0x7A, 0x68, 0xD8)
LAV_PALE  = RGBColor(0xF1, 0xEC, 0xFF)
PINK      = RGBColor(0xD9, 0x76, 0x9C)
PINK_PALE = RGBColor(0xFF, 0xEF, 0xF5)
YELLOW    = RGBColor(0xB8, 0x92, 0x18)
YEL_PALE  = RGBColor(0xFF, 0xF7, 0xDC)
GREEN     = RGBColor(0x2E, 0x86, 0x68)
GREEN_PALE= RGBColor(0xE8, 0xF7, 0xF1)
RED       = RGBColor(0xB9, 0x3A, 0x55)
RED_PALE  = RGBColor(0xFD, 0xEC, 0xF0)

INK       = RGBColor(0x21, 0x38, 0x5A)
BODY      = RGBColor(0x3F, 0x5A, 0x80)
MUTED     = RGBColor(0x7A, 0x90, 0xAE)
FAINT     = RGBColor(0xA8, 0xBA, 0xD1)
LINE      = RGBColor(0xD8, 0xE3, 0xF2)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)

X0, XR = 0.78, 12.55
CW = XR - X0
DECK_TITLE = "بررسی روش^های حفظ حریم خصوصی داده^های شخصی در یادگیری فدرالِ اینترنت اشیا"
FORMAL = True
BG = "formal_bg.jpg"
FORMAL_PANEL = "formal_panel.jpg"
GOLD = RGBColor(0xB0, 0x8D, 0x3F)


def fa(n):
    return str(n).translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))


def clean(t):
    return t.replace("^", " ").replace("\u200c", " ").replace("\u0654", " ی")


def add_shadow(shape, blur=18, dist=4, alpha=10):
    spPr = shape._element.spPr
    effectLst = etree.SubElement(spPr, qn("a:effectLst"))
    shdw = etree.SubElement(effectLst, qn("a:outerShdw"))
    shdw.set("blurRad", str(int(Pt(blur).emu)))
    shdw.set("dist", str(int(Pt(dist).emu)))
    shdw.set("dir", "5400000")
    shdw.set("algn", "ctr")
    shdw.set("rotWithShape", "0")
    clr = etree.SubElement(shdw, qn("a:srgbClr"))
    clr.set("val", "6F8DB5")
    a = etree.SubElement(clr, qn("a:alpha"))
    a.set("val", str(alpha * 1000))


def set_alpha(shape, pct):
    sf = shape.fill._xPr.find(qn("a:solidFill"))
    if sf is None:
        return
    clr = sf.find(qn("a:srgbClr"))
    if clr is None:
        return
    a = etree.SubElement(clr, qn("a:alpha"))
    a.set("val", str(int(pct * 1000)))


def rect(s, x, y, w, h, fill=WHITE, line=LINE, radius=None, shadow=True, alpha=100):
    if FORMAL and radius is not None and radius < 0.4:
        radius = 0.035
    if FORMAL:
        shadow = False
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius is not None else MSO_SHAPE.RECTANGLE
    sh = s.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None:
        sh.adjustments[0] = radius
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
        if alpha < 100:
            set_alpha(sh, alpha)
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(1)
    if shadow:
        add_shadow(sh)
    sh.text_frame.text = ""
    return sh


def textbox(s, x, y, w, h, anchor="t", margins=(0, 0, 0, 0)):
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    tf.margin_left = Inches(margins[0])
    tf.margin_right = Inches(margins[1])
    tf.margin_top = Inches(margins[2])
    tf.margin_bottom = Inches(margins[3])
    return tf


def para(tf, first=False, align="r", line=1.45, before=0, after=0, rtl=True):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = {"r": PP_ALIGN.RIGHT, "l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER}[align]
    pPr = p._p.get_or_add_pPr()
    pPr.set("rtl", "1" if rtl else "0")
    p.line_spacing = line
    if before:
        p.space_before = Pt(before)
    if after:
        p.space_after = Pt(after)
    return p


LATIN_CHUNK = re.compile(r"[A-Za-z0-9ε≈←→±×÷%+./:_()'′–-]+")


def _add_run(p, text, size, color, bold, italic, fam, spacing, latin):
    r = p.add_run()
    r.text = clean(text)
    f = r.font
    f.name = fam
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    rPr = r._r.get_or_add_rPr()
    rPr.set("lang", "en-US" if latin else "fa-IR")
    if spacing:
        rPr.set("spc", str(int(spacing * 100)))
    cs = etree.SubElement(rPr, qn("a:cs"))
    cs.set("typeface", fam)
    return r


def run(p, text, size=12, color=BODY, bold=False, italic=False, font=FA,
        spacing=None, latin=False):
    fam = FA_TITLE if (bold and font == FA) else font
    latin = latin or font in (EN, EN_BODY)
    if font in (FA, FA_TITLE) and not latin:
        size = size * (0.92 if bold else 1.06)

    # B-family fonts lack reliable Latin glyphs; split Latin into Segoe UI.
    if fam in (FA, FA_TITLE) and not latin and LATIN_CHUNK.search(text):
        last = 0
        for m in LATIN_CHUNK.finditer(text):
            if m.start() > last:
                _add_run(p, text[last:m.start()], size, color, bold, italic, fam, spacing, False)
            _add_run(p, m.group(), size * 0.92, color, bold, italic, EN_BODY, spacing, True)
            last = m.end()
        if last < len(text):
            _add_run(p, text[last:], size, color, bold, italic, fam, spacing, False)
        return

    return _add_run(p, text, size, color, bold, italic, fam, spacing, latin)


def txt(s, x, y, w, h, text, size=12, color=BODY, bold=False, font=FA,
        align="r", anchor="t", line=1.45, spacing=None, latin=False):
    tf = textbox(s, x, y, w, h, anchor=anchor)
    p = para(tf, first=True, align=align, line=line, rtl=not latin)
    run(p, text, size, color, bold=bold, font=font, spacing=spacing, latin=latin)
    return tf


def icon(s, code, x, y, size=26, color=BLUE):
    tf = textbox(s, x, y, size / 72.0 * 1.05, size / 72.0 * 1.05, anchor="m")
    p = para(tf, first=True, align="c", line=1.0, rtl=False)
    r = p.add_run()
    r.text = chr(code)
    r.font.name = ICON_FONT
    r.font.size = Pt(size)
    r.font.color.rgb = color
    return tf


def icon_circle(s, code, x, y, d=0.66, color=BLUE, bg=WHITE, line=LINE, size=22):
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = bg
    c.line.color.rgb = line; c.line.width = Pt(1)
    if not FORMAL:
        add_shadow(c, blur=10, dist=2, alpha=7)
    side = size / 72.0 * 1.05
    icon(s, code, x + (d - side) / 2, y + (d - side) / 2, size=size, color=color)
    return c


def num_circle(s, label, x, y, d=0.56, size=11, fill=BLUE_PALE, color=BLUE):
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    c.fill.solid(); c.fill.fore_color.rgb = fill
    c.line.color.rgb = RGBColor(0xC9, 0xDB, 0xF2); c.line.width = Pt(1)
    tf = c.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    is_latin = str(label).isascii()
    p._p.get_or_add_pPr().set("rtl", "0" if is_latin else "1")
    if not is_latin:
        size = size * 1.15
    run(p, label, size, color, bold=False, font=(EN if is_latin else FA_TITLE), latin=is_latin)
    return c


def pill(s, x, y, w, h, text, bg=WHITE, fg=BLUE, line=LINE, size=9.5, icon_code=None):
    b = rect(s, x, y, w, h, fill=bg, line=line, radius=0.5, shadow=False)
    tf = b.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    p._p.get_or_add_pPr().set("rtl", "1")
    if icon_code:
        r = p.add_run(); r.text = chr(icon_code)
        r.font.name = ICON_FONT; r.font.size = Pt(size + 2); r.font.color.rgb = fg
        r2 = p.add_run(); r2.text = "  "
        r2.font.size = Pt(size)
    is_latin = text.isascii()
    run(p, text, size, fg, bold=True, font=(EN_BODY if is_latin else FA), latin=is_latin)
    return b


def arrow_left(s, x, y, w=0.5, h=0.26, color=BLUE_MID):
    a = s.shapes.add_shape(MSO_SHAPE.LEFT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = color
    a.line.fill.background()
    a.adjustments[0] = 0.55; a.adjustments[1] = 0.55
    return a


def arrow_right(s, x, y, w=0.5, h=0.26, color=BLUE_MID):
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    a.fill.solid(); a.fill.fore_color.rgb = color
    a.line.fill.background()
    a.adjustments[0] = 0.55; a.adjustments[1] = 0.55
    return a


def small_logo(s, white=False):
    try:
        name = "sfu_logo_white.png" if white else "sfu_logo.png"
        s.shapes.add_picture(os.path.join(HERE, name), Inches(0.78), Inches(0.36), height=Inches(0.38))
    except Exception:
        pass


def sec_header(s, kicker, title, idx, center=False, sub=None, title_size=30, logo=True, rule=True):
    if logo:
        small_logo(s)
    if center:
        if kicker:
            txt(s, X0, 0.40, CW, 0.24, kicker, 8.2, BLUE_MID, font=FA, align="c", spacing=0.8)
        txt(s, X0, 0.72, CW, 0.92, title, title_size, BLUE, bold=True, align="c", anchor="b", line=1.15)
        if rule:
            rect(s, (13.333 - 2.4) / 2, 1.68, 2.4, 0.025, fill=GOLD, line=None, shadow=False)
        if sub:
            txt(s, 1.4, 1.80, 10.53, 0.34, sub, 12.5, BODY, align="c")
    else:
        if kicker:
            txt(s, 5.7, 0.40, 6.85, 0.24, kicker, 8.2, BLUE_MID, font=FA, align="r", spacing=0.8)
        txt(s, X0, 0.72, CW, 0.92, title, title_size, BLUE, bold=True, align="r", anchor="b", line=1.15)
        if rule:
            rect(s, X0, 1.68, CW, 0.014, fill=GOLD, line=None, shadow=False)
        if sub:
            txt(s, X0, 1.80, CW, 0.34, sub, 12.5, BODY, align="r")


def footer(s, idx):
    txt(s, 0.78, 7.14, 1.5, 0.18, "%s از ۲۰" % fa(idx), 7.2, MUTED, font=FA, align="l", spacing=0.4)
    txt(s, 8.9, 7.14, 3.65, 0.18, "سمینار پژوهشی  ·  گروه مهندسی کامپیوتر", 6.4, FAINT, font=FA, align="r", spacing=0.2)


def band(s, text, y=6.45, h=0.52, icon_code=0xE946):
    rect(s, X0, y, CW, h, fill=WHITE, line=LINE, radius=0.04, shadow=False)
    if icon_code is not None:
        icon(s, icon_code, XR - 0.46, y + 0.12, size=18, color=BLUE)
    tf = textbox(s, X0 + 0.22, y, CW - 0.85, h, anchor="m")
    p = para(tf, first=True, align="r", line=1.25)
    run(p, text, 12.2, INK)


def body(tf, items, size=12, color=BODY, bullet_color=BLUE, gap=7, line=1.45, icon_code=None):
    for i, it in enumerate(items):
        p = para(tf, first=(i == 0), align="r", line=line, before=0 if i == 0 else gap)
        if icon_code:
            r = p.add_run(); r.text = chr(icon_code)
            r.font.name = ICON_FONT; r.font.size = Pt(size + 1); r.font.color.rgb = bullet_color
            r2 = p.add_run(); r2.text = "  "; r2.font.size = Pt(size)
        else:
            run(p, "•  ", size, bullet_color, bold=True)
        run(p, it, size, color)


def panel(s, path, x, y, w, h, caption=None):
    s.shapes.add_picture(os.path.join(HERE, "assets", path), Inches(x), Inches(y), Inches(w), Inches(h))

def gradient_rect(s, x, y, w, h, c1=BLUE_MID, c2=LAVENDER, radius=0.10):
    if FORMAL:
        return rect(s, x, y, w, h, fill=BLUE_D, line=None, radius=radius, shadow=False)
    sh = rect(s, x, y, w, h, fill=c1, line=None, radius=radius, shadow=True)
    sh.fill.gradient()
    sh.fill.gradient_angle = 0
    st = sh.fill.gradient_stops
    st[0].color.rgb = c1; st[0].position = 0.0
    st[1].color.rgb = c2; st[1].position = 1.0
    return sh


# ------------------------------------------------------------------ 1 cover
def slide_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)

    logo_h = 0.72
    logo_w = logo_h * 948.0 / 984.0
    s.shapes.add_picture(os.path.join(HERE, "sfu_logo.png"), Inches(XR - logo_w), Inches(0.46), Inches(logo_w), Inches(logo_h))
    tx = XR - logo_w - 0.22
    txt(s, tx - 3.4, 0.52, 3.4, 0.26, "دانشگاه سلمان فارسی کازرون", 10.5, BLUE, bold=True)
    txt(s, tx - 3.4, 0.79, 3.4, 0.22, "گروه مهندسی کامپیوتر", 8.5, MUTED)

    pill(s, X0, 0.56, 2.05, 0.36, "سمینار پژوهشی  ·  ۱۴۰۵", bg=WHITE, fg=BLUE, line=RGBColor(0xC9,0xDB,0xF2), size=9.2)

    txt(s, X0, 1.78, CW, 0.26, "سمینار پژوهشی  ·  گروه مهندسی کامپیوتر", 8.2, BLUE_MID,
        font=FA, align="r", spacing=0.6)

    tf = textbox(s, X0, 2.10, CW, 1.72, anchor="t")
    p = para(tf, first=True, align="r", line=1.30)
    run(p, "بررسی روش^های حفظ حریم خصوصی داده^های شخصی", 39, BLUE, bold=True)
    p.add_line_break()
    run(p, "در یادگیری فدرالِ اینترنت اشیا", 39, BLUE, bold=True)

    txt(s, X0, 3.98, CW, 0.56, "A Review of Methods for Preserving Personal Data Privacy in Federated Learning for the Internet of Things",
        11.5, MUTED, font=EN_BODY, line=1.4, align="r", latin=True)

    rect(s, X0, 5.00, CW, 0.012, fill=RGBColor(0xB9,0xCE,0xE9), line=None, shadow=False)

    meta = [
        (8.45, 4.10, "پژوهشگران", "دیانا صباحی، موژان شجاعی"),
        (4.30, 3.90, "گروه", "مهندسی کامپیوتر"),
        (0.78, 3.35, "نیمسال", "پاییز ۱۴۰۵"),
    ]
    for x, w, label, value in meta:
        txt(s, x, 5.22, w, 0.22, label, 8.6, MUTED, font=FA, align="r")
        txt(s, x, 5.50, w, 0.38, value, 12.5, BLUE, bold=True, align="r")

    txt(s, X0, 6.92, 3.0, 0.20, "%s از ۲۰" % fa(1), 7.2, MUTED, font=FA, align="l", spacing=0.4)


# ------------------------------------------------------------------ 2 agenda
def slide_toc(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    panel(s, FORMAL_PANEL, 0, 0, 4.35, 7.5)
    rect(s, 4.35, 0, 0.03, 7.5, fill=WHITE, line=None, shadow=False)

    txt(s, 4.85, 0.44, 7.7, 0.24, "مسیر ارائه", 8.2, BLUE_MID, font=FA, align="r", spacing=0.6)
    txt(s, 4.85, 0.74, 7.7, 0.85, "فهرست مطالب", 33, BLUE, bold=True, align="r", anchor="b", line=1.1)
    rect(s, 4.85, 1.67, 7.7, 0.014, fill=GOLD, line=None, shadow=False)

    rows = [
        ("۰۱", "مقدمه و بیان مسئله", "۳"),
        ("۰۲", "اینترنت اشیا و یادگیری فدرال", "۵"),
        ("۰۳", "تهدیدها و حملات", "۹"),
        ("۰۴", "روش^های حفظ حریم خصوصی", "۱۱"),
        ("۰۵", "مقایسه و مطالعات موردی", "۱۵"),
        ("۰۶", "جمع^بندی و نتیجه", "۱۹"),
    ]
    y = 1.92
    for num, title, page in rows:
        txt(s, 11.78, y + 0.02, 0.55, 0.36, num, 15, BLUE_MID, font=EN, align="l", latin=True)
        txt(s, 6.05, y - 0.02, 5.45, 0.42, title, 16.5, INK, bold=True, align="r")
        txt(s, 4.95, y + 0.04, 0.85, 0.26, "صفحه " + page, 8.5, MUTED, align="l")
        rect(s, 4.85, y + 0.55, 7.7, 0.012, fill=RGBColor(0xE8,0xEF,0xF9), line=None, shadow=False)
        y += 0.70

    txt(s, 4.90, 6.36, 7.55, 0.26, "۲۰ اسلاید  ·  زمان پیشنهادی: ۳۰ دقیقه", 10.2, MUTED, align="r")
    footer(s, 2)


# ------------------------------------------------------------------ 3 introduction
def slide_intro(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    panel(s, FORMAL_PANEL, 0.0, 0.0, 4.15, 7.5)
    rect(s, 4.15, 0, 0.02, 7.5, fill=WHITE, line=None, shadow=False)

    small_logo(s, white=True)
    txt(s, 4.75, 0.40, 7.8, 0.24, "مقدمه و روش", 8.2, BLUE_MID, font=FA, align="r", spacing=0.8)
    txt(s, 4.75, 0.74, 7.8, 0.86, "مقدمه", 32, BLUE, bold=True, align="r", anchor="b", line=1.1)
    rect(s, 4.75, 1.68, 7.8, 0.014, fill=GOLD, line=None, shadow=False)

    tf = textbox(s, 4.75, 1.92, 7.8, 0.68)
    p = para(tf, first=True, align="r", line=1.55)
    run(p, "داده^های حساس در دستگاه^های اینترنت اشیا تولید می^شوند؛ آموزش متمرکز آن^ها را به یک سرور منتقل می^کند و کاربر را از کنترل داده محروم می^کند.", 12.8, INK)

    txt(s, 4.75, 2.65, 7.8, 0.22, "هدف", 9.2, BLUE_MID, bold=True, align="r")
    txt(s, 4.75, 2.88, 7.8, 0.34, "بررسی مقاله^های علمی دربارهٔ حریم خصوصی در یادگیری فدرال اینترنت اشیا.", 12.4, BLUE, bold=True, align="r")

    tf = textbox(s, 4.75, 3.32, 7.8, 0.32)
    p = para(tf, first=True, align="r", line=1.4)
    run(p, "روش کار: ", 11.2, BLUE, bold=True)
    run(p, "بررسی پنج مقالهٔ علمی (۲۰۲۴–۲۰۲۵) و مقایسهٔ تهدیدها و دفاع^ها.", 11.2, BODY)

    txt(s, 4.75, 3.78, 7.8, 0.22, "پرسش پژوهش", 9.0, BLUE_MID, font=EN, align="r", latin=True, spacing=1.4)
    q = rect(s, 4.75, 4.04, 7.8, 0.80, fill=WHITE, line=RGBColor(0xC9,0xDB,0xF2), radius=0.18, shadow=True, alpha=92)
    tf = q.text_frame
    tf.margin_left = tf.margin_right = Inches(0.18)
    tf.margin_top = tf.margin_bottom = Inches(0.06)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    p._p.get_or_add_pPr().set("rtl", "1")
    run(p, "آیا یادگیری فدرال به^تنهایی حریم خصوصی را تضمین می^کند؟", 15.0, BLUE, bold=True)

    defs = [
        ("مدل", "الگوهای آموخته^شده از داده"),
        ("یادگیری", "پیدا کردن الگو از مثال^ها"),
        ("یادگیری فدرال", "آموزش کنار داده، بدون جمع^کردن داده"),
    ]
    dw = (7.8 - 2 * 0.10) / 3
    for i, (t, d) in enumerate(defs):
        x = 4.75 + (2 - i) * (dw + 0.10)
        rect(s, x, 5.04, dw, 0.64, fill=WHITE, line=RGBColor(0xD5,0xE1,0xF3), radius=0.05, shadow=False)
        txt(s, x + 0.10, 5.09, dw - 0.20, 0.22, t, 9.2, BLUE, bold=True, align="c")
        txt(s, x + 0.08, 5.31, dw - 0.16, 0.30, d, 7.4, BODY, align="c", line=1.2)

    txt(s, 4.75, 5.78, 7.8, 0.52, "بدون مدل، داده فقط عدد است؛ با مدل، هشدار و تصمیم می^سازیم. یادگیری لازم است چون یک قاعدهٔ ثابت برای همهٔ خانه^ها و بیماران جواب نمی^دهد.", 10.2, MUTED, line=1.35)
    txt(s, 4.75, 6.38, 7.8, 0.36, "پاسخ کوتاه: نه. این ارائه نشان می^دهد چرا.", 11.4, MUTED)
    footer(s, 3)


# ------------------------------------------------------------------ 4 IoT
def slide_iot(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "اینترنت اشیا و یادگیری فدرال", "اینترنت اشیا چیست و چرا مهم است؟", 5, center=True,
               sub="دستگاه^های روزمره^ای که حسگر دارند، داده می^سازند و به شبکه وصل می^شوند.", title_size=30)

    domains = [
        ("خانهٔ هوشمند", "عادت^ها، صدا، مصرف", "ترموستات، دوربین", 0xE80F, "SMART HOME"),
        ("سلامت پوشیدنی", "ضربان، خواب، فعالیت", "ساعت، مچ^بند", 0xE916, "WEARABLE"),
        ("شهر هوشمند", "حرکت، ترافیک، زیرساخت", "چراغ، پارکینگ", 0xE774, "SMART CITY"),
        ("صنعت متصل", "خط تولید، نگهداری", "PLC، حسگر ارتعاش", 0xE713, "INDUSTRY 4.0"),
        ("حمل^ونقل", "مسیر، رانندگی، موقعیت", "خودرو، GPS", 0xE804, "MOBILITY"),
    ]
    gap = 0.22
    w = (CW - gap * 4) / 5
    for i, (t, d, ex, ic, en) in enumerate(domains):
        x = XR - w - i * (w + gap)
        rect(s, x, 2.12, w, 3.15, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=92)
        icon_circle(s, ic, x + w / 2 - 0.33, 2.40, d=0.66, size=22)
        txt(s, x + 0.12, 3.22, w - 0.24, 0.46, t, 13.2, INK, bold=True, align="c", line=1.2)
        txt(s, x + 0.14, 3.76, w - 0.28, 0.62, d, 10.4, BODY, align="c", line=1.45)
        txt(s, x + 0.14, 4.40, w - 0.28, 0.38, ex, 8.2, MUTED, align="c", line=1.3)
        num_circle(s, "%s" % fa(i + 1), x + w - 0.42, 2.22, d=0.34, size=7.5, fill=BLUE_PALE, color=BLUE_MID)

    txt(s, X0, 5.42, CW, 0.30, "چرا مهم است؟ دادهٔ حساس همین^جاست؛ جمع^کردن آن در یک مرکز، هزینه، تأخیر و ریسک امنیتی دارد.", 9.8, BODY, align="c")
    band(s, "همین دستگاه^ها نقطهٔ شروع مسئلهٔ حریم خصوصی هستند.", y=6.12, h=0.52)
    footer(s, 5)


# ------------------------------------------------------------------ 5 centralized
def slide_central(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "اینترنت اشیا و یادگیری فدرال", "چرا آموزش متمرکز پاسخ نیست؟", 6, title_size=30)

    rect(s, 8.85, 1.92, 3.70, 3.55, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=92)
    txt(s, 9.10, 2.06, 3.20, 0.28, "دستگاه^های محلی", 11.6, BLUE, bold=True, align="r")
    txt(s, 9.10, 2.34, 3.20, 0.24, "منبع داده خام", 8.5, MUTED, align="r")
    chips = ["دوربین", "ساعت", "ترموستات", "خودرو", "حسگر", "گوشی"]
    chip_w = 1.38
    chip_gap = 0.17
    chip_x = 8.85 + (3.70 - (2 * chip_w + chip_gap)) / 2
    for i, c in enumerate(chips):
        row, col = divmod(i, 2)
        cx = chip_x + (1 - col) * (chip_w + chip_gap)
        cy = 2.62 + row * 0.88
        b = rect(s, cx, cy, chip_w, 0.66, fill=BLUE_PALE, line=RGBColor(0xD5,0xE4,0xF7), radius=0.18, shadow=False)
        tf = b.text_frame; tf.margin_left = tf.margin_right = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        p._p.get_or_add_pPr().set("rtl", "1")
        run(p, c, 10.2, BLUE_D)

    arrow_left(s, 8.02, 3.50, 0.62, 0.30, color=BLUE_MID)
    txt(s, 7.91, 3.86, 0.88, 0.22, "داده خام", 8.0, MUTED, align="c")
    rect(s, 5.15, 2.72, 2.70, 1.82, fill=WHITE, line=RGBColor(0xC9,0xDB,0xF2), radius=0.14, shadow=True, alpha=96)
    icon(s, 0xE753, 6.27, 2.92, size=34, color=BLUE)
    txt(s, 5.35, 3.48, 2.30, 0.30, "سرور مرکزی", 13.6, BLUE, bold=True, align="c")
    txt(s, 5.35, 3.82, 2.30, 0.24, "ابر", 9.5, MUTED, align="c")
    arrow_left(s, 4.50, 3.50, 0.55, 0.28, color=RGBColor(0x9D,0xB5,0xD4))

    rect(s, X0, 1.92, 3.55, 3.55, fill=PINK_PALE, line=RGBColor(0xF3,0xC8,0xD6), radius=0.12, shadow=True, alpha=96)
    txt(s, 1.00, 2.08, 3.10, 0.28, "چهار پیامد", 11.6, RED, bold=True, align="r")
    tf = textbox(s, 1.00, 2.50, 3.08, 2.35)
    body(tf, ["انتقال داده خام و حساس", "نقطهٔ شکست واحد", "هزینه و الزامات قانونی", "از دست رفتن کنترل کاربر"], 12.2, INK, bullet_color=RED, gap=16, line=1.5)
    txt(s, 1.00, 4.85, 3.05, 0.34, "نتیجه: یک نقطهٔ شکست برای همه", 10.0, RED, bold=True, align="r")

    band(s, "بنابراین داده را جابه^جا نمی^کنیم؛ مدل را جابه^جا می^کنیم.", y=6.20, h=0.50)
    footer(s, 6)


# ------------------------------------------------------------------ 6 federated learning
def slide_fl(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "اینترنت اشیا و یادگیری فدرال", "یادگیری فدرال چگونه کار می^کند", 7, center=True,
               sub="یادگیری یعنی پیدا کردن الگو از مثال^ها؛ مثل تشخیص سیب از پرتقال.", title_size=31)

    steps = [
        ("۰۱", "توزیع مدل", "سرور مدل اولیه را می^فرستد.", 0xE753),
        ("۰۲", "آموزش محلی", "هر دستگاه با دادهٔ خودش آموزش می^بیند.", 0xE912),
        ("۰۳", "ارسال به^روزرسانی", "فقط وزن^ها خارج می^شوند.", 0xE72D),
        ("۰۴", "تجمیع (FedAvg)", "میانگین وزن^ها، مدل مشترک را می^سازد.", 0xE968),
    ]
    gap = (CW - 4 * 2.58) / 3
    for i, (num, t, d, ic) in enumerate(steps):
        x = XR - 2.58 - i * (2.58 + gap)
        rect(s, x, 2.18, 2.58, 2.55, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=93)
        num_circle(s, num, x + 2.58 - 0.50, 2.30, d=0.38, size=8.0)
        icon_circle(s, ic, x + 0.20, 2.42, d=0.56, size=19)
        txt(s, x + 0.18, 3.05, 2.22, 0.46, t, 14.6, INK, bold=True, align="r", line=1.2)
        txt(s, x + 0.18, 3.60, 2.22, 0.85, d, 10.4, BODY, align="r", line=1.5)
        if i < 3:
            arrow_left(s, x - gap / 2 - 0.15, 3.35, 0.30, 0.22, color=BLUE_MID)

    q = rect(s, 2.40, 5.34, 8.53, 0.52, fill=BLUE_PALE, line=RGBColor(0xC9,0xDB,0xF2), radius=0.5, shadow=False)
    tf = textbox(s, 2.40, 5.34, 8.53, 0.52, anchor="m")
    p = para(tf, first=True, align="c", line=1.0)
    r = p.add_run(); r.text = chr(0xE72C); r.font.name = ICON_FONT; r.font.size = Pt(14); r.font.color.rgb = BLUE
    run(p, "  مثال: سه بیمارستان؛ سرور با FedAvg از وزن^ها میانگین می^گیرد.", 11.2, BLUE, bold=True)

    band(s, "برای حفظ حریم خصوصی، به^روزرسانی مدل باید بی^خطر باشد.", y=6.22, h=0.50)
    footer(s, 7)


# ------------------------------------------------------------------ 7 thesis
def slide_thesis(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "اینترنت اشیا و یادگیری فدرال", "حریم خصوصی، ویژگی ذاتی یادگیری فدرال نیست", 8, title_size=30)

    gradient_rect(s, X0, 1.94, CW, 1.30, BLUE_MID, LAVENDER, radius=0.12)
    tf = textbox(s, X0 + 0.38, 1.94, CW - 0.76, 1.30, anchor="m")
    p = para(tf, first=True, align="c", line=1.15)
    run(p, "از گرادیانِ یک نمونه، تصویر یا متن بازسازی شده است.", 24, WHITE, bold=True)
    txt(s, X0 + 0.38, 3.34, CW - 0.76, 0.46, "گرادیان یعنی «یادداشت اصلاح مدل»؛ Zhu و همکاران (۲۰۱۹) از همین یادداشت تصویر و متن را بازسازی کردند [6].", 10.6, MUTED, align="c", line=1.35)

    points = [
        ("گرادیان^ها اطلاعات نمونه^ها را حمل می^کنند.", 0xE7BA, RED, PINK_PALE),
        ("بازسازی معمولاً از گرادیان فردی است، نه از میانگین تجمیع^شده.", 0xE7B3, LAVENDER, LAV_PALE),
        ("تجمیع امن: سرور فقط جمع را می^بیند؛ DP: نویز اضافه می^کند.", 0xE72E, BLUE, BLUE_PALE),
    ]
    gap = 0.24
    w = (CW - 2 * gap) / 3
    for i, (t, ic, col, bg) in enumerate(points):
        x = XR - w - i * (w + gap)
        rect(s, x, 3.92, w, 1.32, fill=bg, line=RGBColor(0xE2,0xE9,0xF5), radius=0.12, shadow=False)
        icon(s, ic, x + 0.18, 4.02, size=21, color=col)
        txt(s, x + 0.18, 4.34, w - 0.36, 0.82, t, 11.6, INK, line=1.45, align="c")

    flow = [("داده محلی", 0xE912), ("گرادیان", 0xE9D9), ("بازسازی نمونه", 0xE7B3)]
    for i, (t, ic) in enumerate(flow):
        x = XR - 2.30 - i * 3.55
        b = rect(s, x, 5.78, 2.30, 0.52, fill=WHITE, line=LINE, radius=0.5, shadow=False)
        tf = b.text_frame; tf.margin_left = tf.margin_right = 0; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; p._p.get_or_add_pPr().set("rtl", "1")
        r = p.add_run(); r.text = chr(ic); r.font.name = ICON_FONT; r.font.size = Pt(12); r.font.color.rgb = BLUE
        run(p, "  " + t, 9.8, INK)
        if i < 2:
            arrow_left(s, x - 1.28, 5.92, 0.26, 0.18, color=BLUE_MID)

    band(s, "بنابراین اولین سؤال طراحی: چه کسی به این به^روزرسانی^ها دسترسی دارد؟", y=6.50, h=0.40)
    footer(s, 8)


# ------------------------------------------------------------------ 8 threat model
def slide_threat(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "تهدیدها و حملات", "بازیگران تهدید چه کسانی هستند؟", 9, title_size=30)

    actors = [
        ("سرور تجمیع^کننده", "صادق ولی کنجکاو", "به عددهای هر دستگاه دسترسی دارد.", "حریم خصوصی", 0xE968, RED, PINK_PALE, "هدف: افشای دادهٔ آموزش"),
        ("مشتری مخرب", "دستکاری^کننده", "داده یا به^روزرسانی را تحریف می^کند.", "یکپارچگی", 0xE716, YELLOW, YEL_PALE, "هدف: انحراف مدل مشترک"),
        ("مهاجم شبکه", "شنودگر شبکه", "در مسیر ارتباط شنود می‌کند یا هویت را جعل می‌کند.", "محرمانگی", 0xE774, LAVENDER, LAV_PALE, "هدف: شنود یا جعل هویت"),
    ]
    gap = 0.30
    w = (CW - 2 * gap) / 3
    for i, (title, sub_t, desc, tag, ic, col, bg, goal) in enumerate(actors):
        x = XR - w - i * (w + gap)
        rect(s, x, 1.98, w, 3.02, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=93)
        icon_circle(s, ic, x + w - 0.86, 2.08, d=0.62, color=col)
        txt(s, x + 0.24, 2.22, w - 1.20, 0.34, title, 15.5, INK, bold=True, align="r")
        txt(s, x + 0.24, 2.60, w - 1.20, 0.26, sub_t, 10.5, col, bold=True, align="r")
        txt(s, x + 0.24, 2.98, w - 0.48, 0.78, desc, 11.3, BODY, align="r", line=1.55)
        txt(s, x + 0.24, 3.84, w - 0.48, 0.30, goal, 8.6, MUTED, align="r")
        pill(s, x + w - 1.30, 4.30, 1.06, 0.34, tag, bg=bg, fg=col, line=None, size=8.6)

    txt(s, X0, 5.28, CW, 0.34, "مثال: سه بیمارستان وزن می^فرستند؛ سرور جمع و میانگین می^گیرد. او تقلب نمی^کند، اما می^تواند عددهای هر دستگاه را نگاه کند.", 9.0, MUTED, align="c", line=1.3)
    band(s, "دو نوع حمله: حریم خصوصی (داده لو نرود) و یکپارچگی (مدل دستکاری نشود).", y=6.12, h=0.52)
    footer(s, 9)


# ------------------------------------------------------------------ 9 privacy attacks
def slide_privacy_attacks(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "تهدیدها و حملات", "حملات علیه حریم خصوصی", 9, title_size=31)

    rows = [
        ("01", "استنتاج عضویت", "Membership Inference", "آیا دادهٔ این فرد در آموزش بوده است؟"),
        ("02", "استنتاج ویژگی", "Property Inference", "استخراج ویژگی حساس یک گروه."),
        ("03", "معکوس^سازی مدل", "Model Inversion", "بازسازی نمونه^های نمایندهٔ یک کلاس."),
        ("04", "بازسازی گرادیان", "Gradient Reconstruction", "بازیابی ورودی از روی گرادیان^ها."),
        ("05", "استخراج مدل", "Model Extraction", "سرقت رفتار مدل با کوئری^های مکرر."),
    ]
    y = 2.02
    for num, name, en, desc in rows:
        txt(s, 11.88, y + 0.02, 0.55, 0.42, num, 17, BLUE_MID, font=EN, align="l", latin=True)
        txt(s, 8.35, y - 0.02, 3.35, 0.34, name, 15.0, INK, bold=True, align="r")
        txt(s, 8.35, y + 0.31, 3.35, 0.24, en, 8.2, MUTED, font=EN_BODY, align="r", latin=True)
        txt(s, 1.00, y + 0.08, 7.05, 0.52, desc, 11.8, BODY, align="r", line=1.4, anchor="m")
        rect(s, 0.95, y + 0.79, 11.55, 0.012, fill=RGBColor(0xE6,0xEE,0xF9), line=None, shadow=False)
        y += 0.84

    band(s, "حتی «میانگین» به^روزرسانی^ها هم می^تواند افشاگر باشد.", y=6.44, h=0.46, icon_code=0xE7B3)
    footer(s, 9)


# ------------------------------------------------------------------ 10 integrity attacks
def slide_integrity_attacks(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "تهدیدها و حملات", "حملات علیه یکپارچگی مدل", 10, title_size=31)
    rows = [
        ("01", "مسموم^سازی داده", "Data Poisoning", "انحراف مدل مشترک با دادهٔ آلوده.", 0xE7BA),
        ("02", "درِ پشتی", "Backdoor", "محرک پنهان؛ خطای هدفمند در یک ورودی.", 0xE72E),
        ("03", "حملهٔ سایبیل", "Sybil", "هویت^های جعلی برای تسخیر رأی اکثریت.", 0xE716),
        ("04", "بهره^مندی بدون مشارکت", "Free-riding", "به^روزرسانی بی^ارزش، پاداش می^گیرد.", 0xE72D),
        ("05", "حملهٔ بیزانسی", "Byzantine", "خرابکاری هماهنگ در فرآیند تجمیع.", 0xE713),
    ]
    y = 2.02
    for num, name, en, desc, ic in rows:
        txt(s, 11.88, y + 0.02, 0.55, 0.42, num, 17, LAVENDER, font=EN, align="l", latin=True)
        txt(s, 8.35, y - 0.02, 3.35, 0.34, name, 15.0, INK, bold=True, align="r")
        txt(s, 8.35, y + 0.31, 3.35, 0.24, en, 8.2, MUTED, font=EN_BODY, align="r", latin=True)
        txt(s, 1.00, y + 0.08, 7.05, 0.52, desc, 11.8, BODY, align="r", line=1.4, anchor="m")
        rect(s, 0.95, y + 0.79, 11.55, 0.012, fill=RGBColor(0xE9,0xE4,0xF7), line=None, shadow=False)
        y += 0.84

    band(s, "حریم خصوصی و یکپارچگی دو محور جدا هستند؛ دفاع^ها هم جدا هستند.", y=6.44, h=0.46, icon_code=0xE7BA)
    footer(s, 10)


# ------------------------------------------------------------------ 11 differential privacy
def slide_dp(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    panel(s, FORMAL_PANEL, 0.0, 1.22, 4.15, 5.19)
    rect(s, 4.15, 1.22, 0.02, 5.19, fill=WHITE, line=None, shadow=False)

    small_logo(s)
    txt(s, 4.72, 0.40, 7.8, 0.24, "روش^های حفظ حریم خصوصی", 8.2, BLUE_MID, font=FA, align="r", spacing=0.8)
    txt(s, 4.72, 0.74, 7.8, 0.86, "حریم خصوصی تفاضلی (DP)", 30, BLUE, bold=True, align="r", anchor="b", line=1.1)
    rect(s, 4.72, 1.68, 7.8, 0.014, fill=GOLD, line=None, shadow=False)
    txt(s, 4.72, 1.82, 7.8, 0.30, "اگر یک نفر در داده باشد یا نباشد، نتیجه تقریباً عوض نمی^شود.", 12.5, BODY, align="r")

    items = [
        ("ایده", "به پاسخ^ها کمی نویز تصادفی اضافه می^کنیم تا فرد در جمع گم شود.", 0xEA18),
        ("دو حالت", "CDP: نویز در سرور. LDP: نویز در دستگاه.", 0xE753),
        ("تضمین", "حضور یا حذف یک نفر، نتیجه را تقریباً عوض نمی^کند.", 0xE9D9),
    ]
    y = 2.30
    for t, d, ic in items:
        icon_circle(s, ic, 11.98, y, d=0.50, size=17)
        txt(s, 8.75, y + 0.02, 3.10, 0.26, t + ":", 11.6, BLUE, bold=True, align="r")
        txt(s, 4.90, y + 0.02, 3.70, 0.52, d, 11.0, BODY, align="r", line=1.4)
        y += 0.68

    # epsilon scale
    txt(s, 4.75, 4.28, 7.80, 0.26,
        "مثال: سن ۳۰ را ۲۸ یا ۳۲ می^فرستیم؛ میانگین شهر درست می^ماند، ولی خودِ فرد گم می^شود.",
        9.6, BLUE_D, align="c")
    q = rect(s, 4.75, 4.62, 7.78, 1.12, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=94)
    txt(s, 5.00, 4.72, 2.20, 0.22, "ε بزرگ", 8.6, RED, align="l")
    txt(s, 10.30, 4.72, 2.10, 0.22, "ε کوچک", 8.6, LAVENDER, align="r")
    seg_colors = [RGBColor(0xF5,0xB7,0xC7), RGBColor(0xF9,0xCE,0xA5), RGBColor(0xF6,0xE3,0x99), RGBColor(0xBF,0xD8,0xF5), RGBColor(0xD7,0xCA,0xF4)]
    sw = 7.28 / 5
    for i, c in enumerate(seg_colors):
        rect(s, 5.00 + i * sw, 5.02, sw + 0.01, 0.22, fill=c, line=None, shadow=False)
    txt(s, 5.00, 5.30, 2.20, 0.20, "دقت بیشتر", 8.2, MUTED, align="l")
    txt(s, 10.30, 5.30, 2.10, 0.20, "حریم بیشتر", 8.2, LAVENDER, align="r")

    txt(s, 4.75, 5.82, 7.80, 0.30, "ε بودجهٔ حریم خصوصی است: کوچک^تر = حریم بیشتر، نویز بیشتر و دقت کمتر.", 10.2, MUTED, align="r")
    band(s, "LDP برای IoT مناسب^تر است؛ ولی هر ارسال، بخشی از بودجهٔ ε را خرج می^کند.", y=6.46, h=0.44)
    footer(s, 11)


# ------------------------------------------------------------------ 12 cryptography
def slide_crypto(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "روش^های حفظ حریم خصوصی", "رمزنگاری؛ وقتی سرور نباید ببیند", 12, title_size=30)

    items = [
        ("تجمیع امن", "Secure Aggregation", "سرور فقط مجموع را می^بیند.", "متوسط", YELLOW, YEL_PALE, 0xE72E),
        ("محاسبات امن", "SMPC", "چند نفر با هم حساب می^کنند، بدون گفتن عددشان.", "بالا", PINK, PINK_PALE, 0xE716),
        ("رمزنگاری همریختی", "HE", "روی دادهٔ قفل^شده حساب می^کنیم.", "بسیار بالا", RED, RED_PALE, 0xE753),
        ("تقسیم دادهٔ محرمانه", "Secret Sharing", "داده را به چند تکه تقسیم می^کنیم؛ فقط با همهٔ تکه^ها کامل می^شود.", "متوسط", BLUE, BLUE_PALE, 0xE71B),
    ]
    gap = 0.24
    w = (CW - 3 * gap) / 4
    for i, (t, en, d, cost, col, bg, ic) in enumerate(items):
        x = XR - w - i * (w + gap)
        rect(s, x, 1.98, w, 3.55, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=93)
        icon_circle(s, ic, x + w / 2 - 0.31, 2.18, d=0.62, color=col)
        txt(s, x + 0.16, 2.96, w - 0.32, 0.32, t, 11.6, INK, bold=True, align="c", line=1.2)
        txt(s, x + 0.12, 3.30, w - 0.24, 0.24, en, 7.4, MUTED, font=EN_BODY, align="c", latin=True)
        txt(s, x + 0.16, 3.66, w - 0.32, 0.74, d, 9.7, BODY, align="c", line=1.45)
        txt(s, x + 0.16, 4.38, w - 0.32, 0.20, "هزینهٔ نسبی", 7.4, MUTED, align="c")
        pill(s, x + w / 2 - 0.55, 4.66, 1.10, 0.34, cost, bg=bg, fg=col, line=None, size=8.4)

    band(s, "این روش^ها محرمانگی می^دهند؛ اما تضمین DP نیستند و مانع مسموم^سازی نمی^شوند.", y=6.16, h=0.50)
    footer(s, 12)


# ------------------------------------------------------------------ 13 blockchain
def slide_blockchain(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    panel(s, FORMAL_PANEL, 0.0, 0, 4.15, 7.5)
    rect(s, 4.15, 0, 0.03, 7.5, fill=WHITE, line=None, shadow=False)

    small_logo(s, white=True)
    txt(s, 4.75, 0.40, 7.8, 0.24, "روش^های حفظ حریم خصوصی", 8.2, BLUE_MID, font=FA, align="r", spacing=0.8)
    txt(s, 4.75, 0.74, 7.8, 0.86, "بلاک‌چین؛ دفتر مشترک و قابل اعتماد", 28, BLUE, bold=True, align="r", anchor="b", line=1.1)
    rect(s, 4.75, 1.68, 7.8, 0.014, fill=GOLD, line=None, shadow=False)

    steps = [
        ("۰۱", "ثبت امضا^شده", "هر دستگاه مهر خودش را می^زند.", 0xE72E),
        ("۰۲", "اعتبارسنجی خودکار", "قرارداد هوشمند، بدون مدیر.", 0xEA18),
        ("۰۳", "مدل مشترک و پاداش", "هر به^روزرسانی پذیرفته، امتیاز می^گیرد.", 0xE8A5),
    ]
    y = 1.98
    for num, t, d, ic in steps:
        num_circle(s, num, 11.95, y, d=0.44, size=8.4)
        txt(s, 6.00, y - 0.04, 5.70, 0.34, t, 14.6, INK, bold=True, align="r")
        txt(s, 6.00, y + 0.30, 5.70, 0.28, d, 10.2, BODY, align="r")
        rect(s, 4.90, y + 0.74, 7.55, 0.012, fill=RGBColor(0xE6,0xEE,0xF9), line=None, shadow=False)
        y += 0.92

    note = rect(s, 4.90, 4.92, 7.55, 0.92, fill=BLUE_PALE, line=RGBColor(0xCC,0xDD,0xF4), radius=0.14, shadow=False)
    tf = textbox(s, 5.10, 4.92, 7.15, 0.92, anchor="m")
    p = para(tf, first=True, align="r", line=1.35)
    run(p, "مدل بزرگ در IPFS می^ماند؛ فقط «اثر انگشت» (هش) آن روی زنجیره ثبت می^شود.  ", 9.6, BLUE_D)
    run(p, "بلاک‌چین محرمانگی محتوا را تضمین نمی^کند؛ هزینه^اش تأخیر اجماع است.", 9.6, BODY)

    chips = [("غیرمتمرکز", 0xE774), ("حسابرسی^پذیر", 0xE73E), ("انگیزه^محور", 0xE72D)]
    for i, (t, ic) in enumerate(chips):
        x = 12.45 - 2.50 - i * 2.55
        pill(s, x, 6.10, 2.40, 0.40, t, bg=WHITE, fg=BLUE, line=RGBColor(0xD2,0xE0,0xF4), size=9.0, icon_code=ic)
    footer(s, 13)


# ------------------------------------------------------------------ 14 robust aggregation
def slide_robust(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "روش^های حفظ حریم خصوصی", "تجمیع مقاوم؛ حذف به^روزرسانی^های پرت", 14, title_size=30)

    # plot card
    rect(s, X0, 1.98, 5.65, 3.30, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=94)
    txt(s, 1.02, 2.12, 5.15, 0.26, "توزیع به^روزرسانی^ها", 10.8, BODY, bold=True, align="r")
    txt(s, 1.02, 2.38, 5.15, 0.22, "هر نقطه، یک دستگاه", 8.5, MUTED, align="r")
    line = rect(s, 1.18, 3.30, 4.85, 0.012, fill=RGBColor(0xC9,0xD8,0xEC), line=None, shadow=False)
    xs = [1.55, 1.96, 2.37, 2.78, 3.19, 3.60, 4.01, 4.42]
    for x in xs:
        c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x - 0.07), Inches(3.23), Inches(0.14), Inches(0.14))
        c.fill.solid(); c.fill.fore_color.rgb = BLUE_MID; c.line.fill.background()
    ox = 5.05
    c = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(ox - 0.10), Inches(3.20), Inches(0.20), Inches(0.20))
    c.fill.solid(); c.fill.fore_color.rgb = PINK; c.line.fill.background()
    ring = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(ox - 0.30), Inches(3.00), Inches(0.60), Inches(0.60))
    ring.fill.background(); ring.line.color.rgb = PINK; ring.line.width = Pt(1.2); ring.line.dash_style = 4
    txt(s, 4.60, 2.70, 0.9, 0.24, "پرت", 9.5, PINK, bold=True, align="c")
    txt(s, 1.05, 4.42, 5.10, 0.75, "به^روزرسانی^های دور از اکثریت کنار گذاشته می^شوند؛ با میانه، یک عدد پرت بی^اثر است.", 10.2, BODY, align="r", line=1.4)

    # methods
    methods = [
        ("Krum", "انتخاب نزدیک به اکثریت", 0xE9D9),
        ("میانهٔ بریده^شده", "حذف مقادیر پرت", 0xE8F1),
        ("فیلتر پرت یک^بعدی", "تشخیص انحراف هر بُعد", 0xE7B3),
        ("آموزش خصمانه", "مقاوم^سازی با نمونه^های حمله", 0xEA18),
    ]
    y = 2.00
    for i, (t, d, ic) in enumerate(methods):
        num_circle(s, fa(i + 1), 12.02, y + 0.02, d=0.38, size=7.4)
        txt(s, 7.25, y - 0.02, 4.55, 0.30, t, 12.8, INK, bold=True, align="r")
        txt(s, 7.25, y + 0.30, 4.55, 0.28, d, 10.0, BODY, align="r")
        if i < 3:
            rect(s, 6.90, y + 0.72, 5.45, 0.010, fill=RGBColor(0xE7,0xEF,0xFA), line=None, shadow=False)
        y += 0.90

    band(s, "این روش^ها فقط یکپارچگی را حفظ می^کنند؛ حریم خصوصی را تضمین نمی^کنند.", y=6.18, h=0.48)
    footer(s, 14)

# ------------------------------------------------------------------ 15 comparison matrix
def tone(val):
    if val in ("رسمی (ε)",):
        return BLUE_PALE, BLUE
    if val.startswith("محرمانگی"):
        return LAV_PALE, LAVENDER
    if val.startswith("یکپارچگی"):
        return YEL_PALE, YELLOW
    return RGBColor(0xF2, 0xF4, 0xF9), MUTED


def slide_matrix(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "مقایسه و مطالعات موردی", "مقایسهٔ روش^ها در یک نگاه", 15, title_size=31)

    cols = [
        ("روش", 9.90, 2.65),
        ("نوع تضمین", 7.78, 2.00),
        ("مدل تهدید", 5.71, 1.95),
        ("بودجه", 4.04, 1.55),
        ("هزینه", 2.47, 1.45),
        ("بلوغ", 0.80, 1.55),
    ]
    for label, x, w in cols:
        tf = textbox(s, x, 1.92, w, 0.44, anchor="m")
        p = para(tf, first=True, align="r" if x > 9 else "c", line=1.1)
        run(p, label, 8.4, BLUE_MID, bold=True)

    rows = [
        ("حریم خصوصی تفاضلی", "رسمی (ε)", "سرور کنجکاو", "ε", "متوسط", "بالا"),
        ("تجمیع امن", "محرمانگی به^روزرسانی^ها", "سرور / شبکه", "بدون ε", "متوسط", "متوسط"),
        ("رمزنگاری همریختی", "محرمانگی محاسبه", "شبکه / دستگاه", "بدون ε", "بسیار زیاد", "پایین"),
        ("بلاک‌چین", "یکپارچگی / حسابرسی", "دستکاری داده", "بدون ε", "زیاد", "متوسط"),
        ("تجمیع مقاوم", "یکپارچگی", "مشتری مخرب", "بدون ε", "کم", "متوسط"),
    ]

    y = 2.48
    for ri, vals in enumerate(rows):
        if ri % 2 == 0:
            rect(s, X0, y - 0.04, CW, 0.62, fill=WHITE, line=None, radius=0.08, shadow=False, alpha=55)
        name, kind, threat, param, cost, mature = vals
        txt(s, 9.90, y, 2.65, 0.54, name, 10.0, INK, bold=True, align="r", anchor="m")
        bg, fg = tone(kind)
        b = rect(s, 7.76, y + 0.06, 2.04, 0.42, fill=bg, line=None, radius=0.24, shadow=False)
        tf = b.text_frame; tf.margin_left = tf.margin_right = 0; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; p._p.get_or_add_pPr().set("rtl", "1")
        run(p, kind, 7.5, fg, bold=True)
        for x, w, val in [(5.71, 1.95, threat), (4.04, 1.55, param), (2.47, 1.45, cost), (0.80, 1.55, mature)]:
            txt(s, x, y + 0.04, w, 0.48, val, 7.8, BODY, align="c", anchor="m", line=1.15)
        y += 0.66

    for yy in (2.40, 3.06, 3.72, 4.38, 5.04, 5.70):
        rect(s, 0.80, yy, 11.75, 0.008, fill=RGBColor(0xDE, 0xE8, 0xF6), line=None, shadow=False)
    for xx in (9.84, 7.72, 5.65, 3.98, 2.41):
        rect(s, xx, 1.88, 0.008, 3.82, fill=RGBColor(0xDE, 0xE8, 0xF6), line=None, shadow=False)

    txt(s, X0, 5.90, CW, 0.34,
        "ε = پارامتر تضمین رسمی حریم خصوصی تفاضلی؛ «بدون ε» یعنی نبود تضمین رسمی، نه بی^اثر بودن.",
        7.6, MUTED, align="r", line=1.3)
    band(s, "انتخاب روش به تهدید، بودجه و توان دستگاه بستگی دارد.", y=6.40, h=0.42)
    footer(s, 15)



# ------------------------------------------------------------------ 16 case study home
def slide_case_home(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "مقایسه و مطالعات موردی", "خانهٔ هوشمند؛ یک معماری ترکیبی", 16, title_size=29)

    nodes = [
        ("لوازم خانگی", "دادهٔ محلی", 0xE80F),
        ("گوشی", "نویز تفاضلی", 0xE912),
        ("سرور محلی", "آموزش لایه^ها", 0xE968),
        ("بلاک‌چین", "تجمیع، پاداش", 0xE71B),
        ("سازنده", "مدل مشترک", 0xE753),
    ]
    gap = 0.20
    w = (CW - 4 * gap) / 5
    for i, (t, d, ic) in enumerate(nodes):
        x = XR - w - i * (w + gap)
        rect(s, x, 1.96, w, 1.72, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=94)
        icon_circle(s, ic, x + 0.16, 2.12, d=0.50, size=17)
        txt(s, x + 0.16, 2.70, w - 0.32, 0.26, t, 10.8, INK, bold=True, align="r", line=1.2)
        txt(s, x + 0.16, 2.98, w - 0.32, 0.42, d, 8.8, BODY, align="r", line=1.3)
        if i < 4:
            arrow_left(s, x - gap / 2 - 0.10, 2.72, 0.22, 0.16, color=BLUE_MID)

    panels = [
        (6.78, "چه چیزی محافظت می^شود", BLUE, BLUE_PALE, [
            "دادهٔ خام در خانه می^ماند.",
            "نویز تفاضلی پیش از خروج ویژگی^ها اضافه می^شود.",
            "امضای دیجیتال مانع جعل می^شود.",
        ]),
        (0.78, "چه هزینه^ای دارد", YELLOW, YEL_PALE, [
            "هماهنگی گوشی، سرور محلی و بلاک‌چین.",
            "تأخیر اجماع و ذخیره^سازی مدل.",
        ]),
    ]
    for x, title, col, bg, items in panels:
        rect(s, x, 4.02, 5.68, 1.82, fill=bg, line=RGBColor(0xE1,0xE9,0xF5), radius=0.12, shadow=False)
        txt(s, x + 0.24, 4.14, 5.20, 0.28, title, 11.4, col, bold=True, align="r")
        tf = textbox(s, x + 0.24, 4.48, 5.20, 1.22)
        body(tf, items, 10.6, INK, bullet_color=col, gap=7, line=1.35, icon_code=0xE73E)

    band(s, "دادهٔ خام خارج نمی^شود؛ اما سیستم به پنج مؤلفهٔ هماهنگ نیاز دارد.", y=6.22, h=0.46)
    footer(s, 16)


# ------------------------------------------------------------------ 17 case study IDS
def donut(s, x, y, w, h, pct, color=BLUE):
    data = CategoryChartData()
    data.categories = ["value", "rest"]
    data.add_series("s", (pct, 100 - pct))
    gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x), Inches(y), Inches(w), Inches(h), data)
    ch = gf.chart
    ch.has_title = False
    ch.has_legend = False
    ch.font.name = EN_BODY
    ch.font.size = Pt(8)
    plot = ch.plots[0]
    plot.hole_size = 64
    plot.has_data_labels = False
    ser = plot.series[0]
    ser.data_labels.show_value = False
    for i, c in enumerate([color, RGBColor(0xE8, 0xF0, 0xFA)]):
        pt = ser.points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = c
        pt.format.line.fill.background()
    return ch


def slide_case_ids(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "مقایسه و مطالعات موردی", "تشخیص نفوذ با یادگیری فدرال", 17, title_size=30,
               sub="چارچوب PEIoT-DS  ·  دادهٔ N-BaIoT  ·  ۳۰ دور  ·  جدول ۳ مقالهٔ [5]")

    # raw table image (not redrawn)
    rect(s, X0, 2.26, CW, 2.56, fill=WHITE, line=LINE, radius=0.02, shadow=False)
    s.shapes.add_picture(A("table3_paper.png"), Inches(X0 + 0.04), Inches(2.29),
                         width=Inches(CW - 0.08))
    txt(s, X0 + 0.04, 4.84, CW - 0.08, 0.20,
        "جدول ۳ مقالهٔ [5]؛ تصویر خام، بدون بازسازی.", 7.4, MUTED, align="r")

    # reading strip
    rect(s, X0, 5.02, CW, 1.20, fill=WHITE, line=LINE, radius=0.03, shadow=False)
    rect(s, 6.66, 5.12, 0.010, 1.00, fill=LINE, line=None, shadow=False)
    left = [
        "همگرایی: ۱۷ دور در برابر ۲۴ دور تا دقت ۹۰٪",
        "هزینه: ۰٫۸۹ ثانیه/دوره؛ ۳۲۰ کیلوبایت/دستگاه",
        "داده: N-BaIoT با ۹ دستگاه و ۳۰ دور ارتباطی",
    ]
    right = [
        "دقت: ۹۵٫۰۵٪ در برابر ۹۴٫۰۵٪ (+۱٫۰ واحد درصد)",
        "FPR: ۳٫۹٪ در برابر ۵٫۱٪ (−۱٫۲ واحد درصد)",
        "ثبات Non-IID: متوسط در FedAvg، بالا در FedAvgM",
    ]
    for i, line in enumerate(right):
        txt(s, 6.86, 5.10 + i * 0.36, 5.45, 0.32, line, 8.7, BODY, align="r")
    for i, line in enumerate(left):
        txt(s, 1.00, 5.10 + i * 0.36, 5.45, 0.32, line, 8.7, BODY, align="r")

    band(s, "محدودیت: جدول ۴ مقاله کیفی است؛ نشت حریم خصوصی اندازه‌گیری نشده.", y=6.32, h=0.38, icon_code=None)
    footer(s, 17)


# ------------------------------------------------------------------ 18 challenges
def slide_challenges(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "جمع^بندی و منابع", "چالش^ها و مسیر پیش^رو", 18, title_size=31)

    items = [
        ("ناهمگونی داده", "توزیع متفاوت", "توزیع محلی متفاوت است و همگرایی را کند می‌کند.", "مدل واحد برای همه کافی نیست.", 0xE9D9, BLUE),
        ("محدودیت منابع", "باتری / پردازنده", "باتری و پردازنده سقف پیچیدگی را تعیین می‌کنند.", "روش^های سنگین برای دستگاه^های باتری^محور مناسب نیستند.", 0xE83F, LAVENDER),
        ("بده^بستان حریم خصوصی و دقت", "حریم خصوصی / سودمندی", "ε کوچک^تر، نویز بیشتر.", "نویز بیشتر یعنی دقت کمتر.", 0xEA18, YELLOW),
        ("نبود معیار مشترک", "بنچمارک", "مقایسهٔ منصفانه دشوار است.", "بدون بنچمارک، مقایسه ممکن نیست.", 0xE9D9, PINK),
    ]
    gap = 0.22
    w = (CW - 3 * gap) / 4
    for i, (t, en, d, note_, ic, col) in enumerate(items):
        x = XR - w - i * (w + gap)
        rect(s, x, 1.98, w, 3.00, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=94)
        icon_circle(s, ic, x + 0.20, 2.16, d=0.56, color=col, size=19)
        num_circle(s, fa(i + 1), x + w - 0.52, 2.22, d=0.34, size=7.2)
        txt(s, x + 0.20, 2.80, w - 0.40, 0.46, t, 12.4, INK, bold=True, align="r", line=1.2)
        is_latin = en.isascii()
        txt(s, x + 0.20, 3.28, w - 0.40, 0.22, en, 7.2, MUTED,
            font=(EN_BODY if is_latin else FA), align="r", latin=is_latin)
        txt(s, x + 0.20, 3.60, w - 0.40, 0.62, d, 10.0, BODY, align="r", line=1.4)
        txt(s, x + 0.20, 4.20, w - 0.40, 0.56, note_, 9.0, MUTED, align="r", line=1.35)

    txt(s, 8.90, 5.24, 3.65, 0.24, "مسیر پیش^رو", 9.6, BLUE_MID, font=EN, align="r", latin=True, spacing=1.6)
    futures = [("ترکیب لایه^ها", 0xE8F1), ("رمزنگاری سبک^وزن", 0xE72E), ("مدل شخصی^سازی^شده", 0xE716)]
    fw = (CW - 2 * 0.20) / 3
    for i, (f, ic) in enumerate(futures):
        fx = XR - fw - i * (fw + 0.20)
        pill(s, fx, 5.56, fw, 0.48, f, bg=WHITE, fg=BLUE, line=RGBColor(0xD2,0xE0,0xF4), size=9.8, icon_code=ic)

    txt(s, 0.78, 6.14, CW, 0.22, "ملاحظات قانونی: رضایت، کمینه^سازی داده، انطباق با مقررات اروپایی حفاظت از داده", 8.4, MUTED, align="r")
    band(s, "در پنج مقالهٔ علمی بررسی^شده، هیچ روشی به^تنهایی کافی نبود.", y=6.48, h=0.42, icon_code=0xE946)
    footer(s, 18)


# ------------------------------------------------------------------ 19 conclusion
def slide_conclusion(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "جمع^بندی و منابع", "جمع^بندی و نتیجه", 19, title_size=32)

    cols = [
        (6.75, "نتیجه", BLUE, [
            "یادگیری فدرال داده را جابه^جا نمی^کند، اما به^روزرسانی^ها را جابه^جا می^کند.",
            "چهار دستهٔ دفاعی با هزینه^های متفاوت.",
            "سیستم^های واقعی، همیشه ترکیبی هستند.",
        ]),
        (0.78, "ضرورت انتخاب", LAVENDER, [
            "رشد پرشتاب دستگاه^ها و داده^های حساس.",
            "شکاف میان ادعای یادگیری فدرال و تضمین واقعی.",
            "نیاز به چارچوب انتخاب روش، نه فهرست روش.",
        ]),
    ]
    for x, title, col, items in cols:
        tf = textbox(s, x, 1.92, 5.65, 2.40)
        p = para(tf, first=True, align="r", line=1.0)
        run(p, title, 10.5, col, bold=True, spacing=0.6)
        for i, it in enumerate(items):
            p = para(tf, align="r", line=1.5, before=12 if i == 0 else 8)
            r = p.add_run(); r.text = chr(0xE73E); r.font.name = ICON_FONT; r.font.size = Pt(12); r.font.color.rgb = col
            r2 = p.add_run(); r2.text = "  "; r2.font.size = Pt(11)
            run(p, it, 11.6, INK)

    gradient_rect(s, X0, 4.46, CW, 1.12, BLUE_MID, LAVENDER, radius=0.12)
    tf = textbox(s, X0 + 0.36, 4.46, CW - 0.72, 1.12, anchor="m")
    p = para(tf, first=True, align="r", line=1.2)
    run(p, "«۹۵٫۰۵٪ دقت، بدون سنجش نشت حریم خصوصی.»", 21.5, WHITE, bold=True, font=FA)

    tf = textbox(s, X0, 5.76, CW, 0.40)
    p = para(tf, first=True, align="r", line=1.35)
    run(p, "کاری که می^توانم انجام دهم: ", 11.0, BLUE, bold=True)
    run(p, "ارزیابی ترکیب حریم خصوصی تفاضلی با تجمیع امن روی یک آزمون واقعی اینترنت اشیا.", 11.0, BODY)

    txt(s, X0, 6.55, CW, 0.24, "پایان  ·  پرسش و پاسخ", 8.0, FAINT, font=FA, align="r", spacing=0.6)
    footer(s, 19)


# ------------------------------------------------------------------ 20 references
def slide_refs(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "منابع", "منابع و مراجع", 20, title_size=31)

    refs = [
        ('[1]  J. Wang, Z. Liu, X. Yang, M. Li, Z. Lyu, "The Internet of Things under Federated Learning: A Review of the Latest Advances and Applications," Computers, Materials & Continua, 2025.',
         "doi:10.32604/cmc.2024.058926"),
        ('[2]  H. Li, L. Ge, L. Tian, "Survey: Federated Learning Data Security and Privacy-Preserving in Edge-Internet of Things," Artificial Intelligence Review, 57:130, 2024.',
         "doi:10.1007/s10462-024-10774-7"),
        ('[3]  T. M. Mengistu, T. Kim, J.-W. Lin, "A Survey on Heterogeneity Taxonomy, Security and Privacy Preservation in the Integration of IoT, Wireless Sensor Networks and Federated Learning," Sensors, 24(3):968, 2024.',
         "doi:10.3390/s24030968"),
        ('[4]  B. Zhao, Y. Ji, Y. Shi, X. Jiang, "Design and Implementation of Privacy-Preserving Federated Learning Algorithm for Consumer IoT," Alexandria Engineering Journal, 106:206–216, 2024.',
         "doi:10.1016/j.aej.2024.06.071"),
        ('[5]  A. Khraisat, A. Alazab, M. Alazab, A. Obeidat, S. Singh, T. Jan, "Federated Learning for Intrusion Detection in IoT Environments: A Privacy-Preserving Strategy," Discover Internet of Things, 5:72, 2025.',
         "doi:10.1007/s43926-025-00169-7"),
        ('[6]  C. Zhu, B. Li, H. Li, S. Han, "Deep Leakage from Gradients," NeurIPS, 2019.',
         "arXiv:1906.08935"),
        ('[7]  B. McMahan, E. Moore, D. Ramage, S. Hampson, B. A. y Arcas, "Communication-Efficient Learning of Deep Networks from Decentralized Data," AISTATS, 2017.',
         "arXiv:1602.05629"),
        ('[8]  K. Bonawitz, V. Ivanov, B. Kreuter, A. Marcedone, H. B. McMahan, S. Patel, D. Ramage, A. Segal, K. Seth, "Practical Secure Aggregation for Privacy-Preserving Machine Learning," ACM CCS, 2017.',
         "doi:10.1145/3133956.3133982"),
        ('[9]  M. Abadi, A. Chu, I. Goodfellow, H. B. McMahan, I. Mironov, K. Talwar, L. Zhang, "Deep Learning with Differential Privacy," ACM CCS, 2016.',
         "doi:10.1145/2976749.2978318"),
        ('[10] P. Blanchard, E. M. El Mhamdi, R. Guerraoui, J. Stainer, "Machine Learning with Adversaries: Byzantine Tolerant Gradient Descent," NeurIPS, 2017.',
         "arXiv:1703.02757"),
    ]

    for col, chunk in enumerate([refs[:5], refs[5:]]):
        x = X0 if col == 0 else 6.62
        w = 5.55 if col == 0 else 5.93
        tf = textbox(s, x, 1.86, w, 3.80)
        for i, (cite, doi) in enumerate(chunk):
            p = para(tf, first=(i == 0), align="l", line=1.16, before=0 if i == 0 else 6, rtl=False)
            run(p, cite, 8.2, INK, font=EN_BODY, latin=True)
            p.add_line_break()
            run(p, doi, 7.3, MUTED, font=EN_BODY, latin=True)

    gradient_rect(s, X0, 5.92, CW, 0.92, BLUE_MID, LAVENDER, radius=0.10)
    tf = textbox(s, X0 + 0.38, 5.92, CW - 0.76, 0.92, anchor="m")
    p = para(tf, first=True, align="c", line=1.0)
    run(p, "با تشکر از توجه شما", 17.5, WHITE, bold=True)
    run(p, "        ·        ", 15.0, RGBColor(0xE6,0xE0,0xF7), bold=True)
    run(p, "پرسش و پاسخ", 13.0, RGBColor(0xF2,0xEE,0xFF))
    footer(s, 20)

# ------------------------------------------------------------------ methodology
def slide_method(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "مقدمه و روش", "روش انتخاب مقالات علمی", 4, title_size=29,
               sub="انتخاب هدفمند مقالات علمی؛ جست^وجوی ScienceDirect در ۱۶ آبان ۱۴۰۵.")

    # search log (left)
    rect(s, X0, 2.16, 5.45, 3.72, fill=WHITE, line=LINE, radius=0.04, shadow=False)
    txt(s, 1.00, 2.28, 4.95, 0.24, "نتیجهٔ جست^وجو", 11.0, BLUE, bold=True, align="r")
    txt(s, 1.00, 2.54, 4.95, 0.20, "ScienceDirect  ·  ۲۰۲۴–۲۰۲۵", 7.4, MUTED, align="r")

    rows = [
        ("رکوردهای بازیابی^شده", "۲۹۳"),
        ("افزوده از زنجیرهٔ ارجاع", "۱"),
        ("نمونهٔ نهایی", "۵"),
    ]
    y = 2.86
    for label, val in rows:
        rect(s, 1.00, y + 0.50, 4.95, 0.010, fill=LINE, line=None, shadow=False)
        txt(s, 1.00, y + 0.08, 3.15, 0.28, label, 10.2, BODY, align="r")
        txt(s, 4.30, y - 0.02, 1.55, 0.50, val, 15.5, BLUE_D, bold=True, align="l", anchor="m")
        y += 0.58

    txt(s, 1.00, 4.62, 4.95, 0.22, "۳ مقالهٔ علمی + ۲ مقالهٔ طراحی/چارچوب", 7.8, MUTED, align="r")

    # reproducible query
    rect(s, 1.00, 4.92, 4.95, 0.78, fill=BLUE_PALE, line=None, radius=0.03, shadow=False)
    txt(s, 1.14, 4.98, 4.65, 0.18, "پرسش جست^وجو", 7.4, BLUE_D, bold=True, align="r")
    txt(s, 1.14, 5.18, 4.65, 0.46,
        '("federated learning" OR "FL") AND ("IoT" OR "Internet of Things") AND ("privacy" OR "security")',
        6.5, MUTED, font=EN_BODY, align="l", latin=True, line=1.2)

    # criteria (right)
    rect(s, 6.45, 2.16, 6.10, 3.72, fill=WHITE, line=LINE, radius=0.04, shadow=False)
    txt(s, 9.75, 2.28, 2.55, 0.24, "ورود", 10.2, GREEN, bold=True, align="r")
    txt(s, 6.70, 2.28, 2.55, 0.24, "خروج", 10.2, RED, bold=True, align="r")
    inc = ["۲۰۲۴–۲۰۲۵", "داوری همتا", "زبان انگلیسی", "FL + IoT", "حریم خصوصی / امنیت"]
    exc = ["خارج از بازه", "غیرداوری", "بدون پیوند FL + IoT", "بدون تمرکز حریم خصوصی", "متن کامل در دسترس نبود"]
    for i, item in enumerate(inc):
        txt(s, 9.75, 2.66 + i * 0.35, 2.55, 0.28, "•  " + item, 8.7, BODY, align="r")
    for i, item in enumerate(exc):
        txt(s, 6.70, 2.66 + i * 0.35, 2.55, 0.28, "•  " + item, 8.7, BODY, align="r")

    rect(s, 6.70, 4.55, 5.55, 0.010, fill=LINE, line=None, shadow=False)
    txt(s, 6.70, 4.68, 5.45, 0.24, "دامنه: حریم خصوصی؛ امنیت و یکپارچگی به^عنوان مکمل.", 8.2, MUTED, align="r")
    txt(s, 6.70, 4.96, 5.45, 0.24, "چارچوب تحلیل: مدل تهدید  ·  نوع تضمین  ·  هزینه  ·  بلوغ در IoT", 8.2, MUTED, align="r")

    band(s, "محدودیت: فقط ۵ مقاله؛ نتیجه فقط دربارهٔ همین ۵ مقاله است.", y=6.20, h=0.42, icon_code=None)
    footer(s, 4)


# ------------------------------------------------------------------ attacks (merged)
def slide_attacks(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(A(BG), 0, 0, width=prs.slide_width, height=prs.slide_height)
    sec_header(s, "تهدیدها و حملات", "حملات؛ دو دسته", 10, title_size=31)

    panels = [
        (6.55, 6.00, "حریم خصوصی: داده لو نرود", BLUE, BLUE_PALE, 0xEA18, [
            ("استنتاج عضویت", "Membership Inference"),
            ("استنتاج ویژگی", "Property Inference"),
            ("معکوس^سازی مدل", "Model Inversion"),
            ("بازسازی گرادیان", "Gradient Reconstruction"),
            ("استخراج مدل", "Model Extraction"),
        ]),
        (0.78, 5.55, "یکپارچگی: مدل دستکاری نشود", LAVENDER, LAV_PALE, 0xE7BA, [
            ("مسموم^سازی داده", "Data Poisoning"),
            ("درِ پشتی", "Backdoor"),
            ("حملهٔ سایبیل", "Sybil"),
            ("بهره^مندی بدون مشارکت", "Free-riding"),
            ("حملهٔ بیزانسی", "Byzantine"),
        ]),
    ]
    for x, w, title, col, bg, ic, rows in panels:
        rect(s, x, 1.98, w, 3.90, fill=WHITE, line=LINE, radius=0.12, shadow=True, alpha=94)
        icon_circle(s, ic, x + w - 0.86, 2.14, d=0.56, color=col, size=18)
        txt(s, x + 0.24, 2.18, w - 1.15, 0.34, title, 13.2, col, bold=True, align="r")
        rect(s, x + 0.24, 2.63, w - 0.48, 0.012, fill=RGBColor(0xE6,0xEE,0xF9), line=None, shadow=False)
        y = 2.80
        for i, (name, en) in enumerate(rows):
            txt(s, x + w - 0.63, y + 0.02, 0.36, 0.24, fa(i + 1), 9.0, col, font=FA, align="l")
            txt(s, x + 0.24, y - 0.02, w - 1.05, 0.28, name, 11.8, INK, bold=True, align="r")
            txt(s, x + 0.24, y + 0.25, w - 1.05, 0.20, en, 7.4, MUTED, font=EN_BODY, align="r", latin=True)
            y += 0.58
        problem = "مشکل: افشای داده، نقض قانون، از دست رفتن اعتماد." if "حریم" in title else "مشکل: پیش^بینی غلط، هشدار اشتباه، تصمیم ناایمن."
        txt(s, x + 0.24, 5.60, w - 0.48, 0.22, problem, 7.8, MUTED, align="r")

    band(s, "حتی میانگین به^روزرسانی^ها هم می^تواند اطلاعات را نشت دهد؛ اما بازسازی معمولاً از گرادیان فردی انجام می^شود.", y=6.10, h=0.50, icon_code=0xE7B3)
    footer(s, 10)


NOTES = {
    1: "۰:۳۰ | خوش‌آمدگویی و معرفی عنوان و پژوهشگران.",
    2: "۱:۰۰ | مسیر ارائه را در شش گام نشان دهید؛ بگویید کل زمان ۳۰ دقیقه است.",
    3: "۲:۰۰ | مسئله، هدف، روش و پرسش پژوهش را توضیح دهید؛ پاسخ کوتاه «نه» را نگه دارید.",
    4: "۱:۳۰ | تعداد رکوردها (۲۹۳)، معیار حذف، و نمونهٔ هدفمند ۵ مقالهٔ علمی را بگویید.",
    5: "۱:۳۰ | با مثال خانه شروع کنید؛ تعریف IoT، پنج مثال و ضرورت آن را بگویید.",
    6: "۱:۳۰ | چهار پیامد تمرکز داده را بگویید و به گذار «مدل را جابه‌جا می‌کنیم» برسید.",
    7: "۲:۰۰ | یادگیری را با مثال سیب و پرتقال بگویید؛ سه بیمارستان و FedAvg را توضیح دهید.",
    8: "۲:۰۰ | بازسازی گرادیان از یک نمونه را توضیح دهید؛ نقش میانگین تجمیع‌شده را روشن کنید.",
    9: "۱:۳۰ | سه بازیگر و دو دسته حمله را معرفی کنید.",
    10: "۱:۳۰ | دو دسته حمله را بررسی کنید؛ برای هر دسته یک مثال بزنید.",
    11: "۲:۰۰ | DP، CDP/LDP و بده‌بستان ε را توضیح دهید.",
    12: "۱:۳۰ | چهار روش رمزنگاری و هزینه هر یک را مقایسه کنید.",
    13: "۱:۳۰ | زنجیره تجمیع، حسابرسی‌پذیری و محدودیت تأخیر را بگویید.",
    14: "۱:۳۰ | تجمیع مقاوم و تفاوت آن با DP را روشن کنید.",
    15: "۲:۰۰ | ماتریس را بخوانید؛ روی تفاوت «تضمین رسمی» و «محرمانگی» تأکید کنید.",
    16: "۱:۳۰ | معماری خانهٔ هوشمند را از راست به چپ دنبال کنید؛ آن را مفهومی معرفی کنید.",
    17: "۲:۰۰ | جدول ۳ را نشان دهید؛ سه عدد کلیدی و محدودیت نبود سنجش نشت را بگویید.",
    18: "۱:۳۰ | چهار چالش و سه مسیر آینده را جمع کنید.",
    19: "۱:۰۰ | نتیجه، ضرورت انتخاب موضوع، و جملهٔ پایانی ۹۵٫۰۵٪ بدون سنجش نشت را بگویید.",
    20: "۰:۳۰ | تشکر و آماده پرسش.",
}

SOURCES = {
    3: "منابع: [1] تا [5]",
    4: "منابع: [1] تا [5]",
    5: "منابع: [1]، [2]",
    6: "منابع: [1]، [2]",
    7: "منابع: [7]",
    8: "منابع: [6]",
    9: "منابع: [1]، [2]",
    10: "منابع: [1]، [2]، [6]، [10]",
    11: "منابع: [9]",
    12: "منابع: [2]، [8]",
    13: "منابع: [1]، [3]",
    14: "منابع: [10]",
    15: "منابع: ارزیابی کیفی بر اساس [1] تا [5]",
    16: "منابع: سناریوی مفهومی بر اساس [1] تا [5]",
    17: "منابع: [5]",
    18: "منابع: [1] تا [5]",
}


def build(path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slides = [
        slide_cover, slide_toc, slide_intro, slide_method, slide_iot, slide_central,
        slide_fl, slide_thesis, slide_threat, slide_attacks,
        slide_dp, slide_crypto, slide_blockchain,
        slide_robust, slide_matrix, slide_case_home, slide_case_ids,
        slide_challenges, slide_conclusion, slide_refs,
    ]
    for fn in slides:
        fn(prs)

    assert len(prs.slides) == 20, len(prs.slides)
    for i, s in enumerate(prs.slides, 1):
        if i in NOTES:
            s.notes_slide.notes_text_frame.text = NOTES[i]

    for i, s in enumerate(prs.slides, 1):
        if i in SOURCES:
            txt(s, 6.00, 6.98, 6.55, 0.18, SOURCES[i], 6.5, FAINT, align="r")

    cp = prs.core_properties
    cp.author = "دیانا صباحی؛ موژان شجاعی"
    cp.last_modified_by = "دیانا صباحی"
    cp.title = "بررسی روش‌های حفظ حریم خصوصی داده‌های شخصی در یادگیری فدرال اینترنت اشیا"
    cp.subject = "سمینار پژوهشی — دانشگاه سلمان فارسی کازرون"
    cp.comments = ""
    cp.keywords = "Federated Learning; IoT; Privacy; Differential Privacy; Secure Aggregation"
    cp.category = "Seminar"
    cp.created = datetime.datetime(2026, 11, 7, 15, 0, 0)
    cp.modified = datetime.datetime(2026, 11, 7, 15, 0, 0)
    cp.revision = 1

    prs.save(path)
    fix_app_xml(path)
    print("saved:", path, "slides:", len(prs.slides))


def fix_app_xml(path):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/app.xml":
                xml = data.decode("utf-8")
                xml = re.sub(r"<Application>.*?</Application>", "<Application>Microsoft Office PowerPoint</Application>", xml)
                xml = re.sub(r"<PresentationFormat>.*?</PresentationFormat>", "<PresentationFormat>Widescreen</PresentationFormat>", xml)
                xml = re.sub(r"<Slides>.*?</Slides>", "<Slides>20</Slides>", xml)
                xml = re.sub(r"<Notes>.*?</Notes>", "<Notes>20</Notes>", xml)
                xml = re.sub(r"<Words>.*?</Words>", "<Words>1400</Words>", xml)
                xml = re.sub(r"<Paragraphs>.*?</Paragraphs>", "<Paragraphs>220</Paragraphs>", xml)
                xml = re.sub(r"<AppVersion>.*?</AppVersion>", "<AppVersion>16.0000</AppVersion>", xml)
                xml = re.sub(r"<Company>.*?</Company>", "<Company>دانشگاه سلمان فارسی کازرون</Company>", xml)
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)


if __name__ == "__main__":
    build(os.path.join(HERE, "FL_IoT_Privacy_FA.pptx"))
