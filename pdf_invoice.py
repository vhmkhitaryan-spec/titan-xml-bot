"""Renders a PDF that visually mirrors the official Armenian tax invoice form
("Ապրանքների մատակարարման հարկային հաշիվ"), filled with the same data the bot
puts into the XML. This is a preview/printout only -- the XML remains the
document actually submitted; nothing here feeds back into it.
"""
import os

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
pdfmetrics.registerFont(TTFont("Arm", os.path.join(FONT_DIR, "DejaVuSans.ttf")))
pdfmetrics.registerFont(TTFont("Arm-Bold", os.path.join(FONT_DIR, "DejaVuSans-Bold.ttf")))

GRID = TableStyle([
    ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
    ("FONTNAME", (0, 0), (-1, -1), "Arm"),
    ("FONTSIZE", (0, 0), (-1, -1), 7),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
])


def _digit_boxes(value, n):
    """A small grid of n bordered single-digit cells (the ՀՎՀՀ-style boxes)."""
    s = str(value or "").replace("/", "")
    digits = list(s.ljust(n)[:n])
    t = Table([digits], colWidths=[5.2 * mm] * n, rowHeights=[5.5 * mm])
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("FONTNAME", (0, 0), (-1, -1), "Arm-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


def _field_row(label, value, label_w=95 * mm, val_w=140 * mm, bold_label=True):
    style_label = ParagraphStyle("l", fontName="Arm-Bold" if bold_label else "Arm", fontSize=7.5, leading=9)
    style_val = ParagraphStyle("v", fontName="Arm", fontSize=8, leading=10)
    t = Table(
        [[Paragraph(label, style_label), Paragraph(str(value or ""), style_val)]],
        colWidths=[label_w, val_w],
    )
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return t


def _section_header(text, width):
    style = ParagraphStyle("h", fontName="Arm-Bold", fontSize=8.5, alignment=TA_CENTER)
    t = Table([[Paragraph(text, style)]], colWidths=[width])
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def build_invoice_pdf(parsed, constants, buyer_info, goods, out_path):
    page_w, page_h = landscape(A4)
    doc = SimpleDocTemplate(
        out_path, pagesize=landscape(A4),
        leftMargin=10 * mm, rightMargin=10 * mm, topMargin=8 * mm, bottomMargin=8 * mm,
    )
    content_w = page_w - 20 * mm

    title_style = ParagraphStyle("title", fontName="Arm-Bold", fontSize=12, alignment=TA_CENTER)
    small = ParagraphStyle("small", fontName="Arm", fontSize=7.5)
    small_b = ParagraphStyle("small_b", fontName="Arm-Bold", fontSize=7.5)

    story = []
    story.append(Paragraph("Ապրանքների մատակարարման հարկային հաշիվ", title_style))
    story.append(Spacer(1, 4))

    # [1]/[4] issue+supply date, [2]/[3] series/number
    top = Table(
        [[
            Paragraph("[1] Դուրս գրման ամսաթիվը", small_b),
            Paragraph("[2] Սերիա", small_b),
            Paragraph("[3] Համար", small_b),
        ], [
            Paragraph("", small),
            Paragraph("", small),
            Paragraph("", small),
        ], [
            Paragraph("[4] Մատակարարման ամսաթիվը՝ <b>%s</b>" % parsed["date"], small),
            "", "",
        ]],
        colWidths=[content_w * 0.5, content_w * 0.25, content_w * 0.25],
        rowHeights=[7 * mm, 7 * mm, 7 * mm],
    )
    top.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, 1), 0.6, colors.black),
        ("SPAN", (0, 2), (-1, 2)),
        ("BOX", (0, 2), (-1, 2), 0.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(top)
    story.append(Spacer(1, 4))

    # Supplier section -------------------------------------------------
    story.append(_section_header("Ապրանքներ մատակարարող (առաքող) անձի տվյալներ", content_w))
    story.append(Table([[
        Paragraph("[8] Հարկ վճարողի հաշվառման համարը (ՀՎՀՀ)", small_b),
        _digit_boxes(constants.get("SupplierTIN"), 8),
    ]], colWidths=[content_w - 50 * mm, 50 * mm],
        style=TableStyle([("GRID", (0, 0), (-1, -1), 0.6, colors.black), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 4)])))
    story.append(_field_row("[9] Անվանում", constants.get("SupplierName"), label_w=content_w * 0.28, val_w=content_w * 0.72))
    vat_num = str(constants.get("SupplierVATNumber") or "")
    story.append(Table([[
        Paragraph("[10] Ավելացված արժեքի հարկ վճարողի հաշվառման համարը", small_b),
        _digit_boxes(vat_num.split("/")[0], 8),
        Paragraph("/ " + (vat_num.split("/")[1] if "/" in vat_num else ""), small_b),
    ]], colWidths=[content_w - 65 * mm, 45 * mm, 20 * mm],
        style=TableStyle([("GRID", (0, 0), (-1, -1), 0.6, colors.black), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 4)])))
    story.append(_field_row(
        "[11] Բանկի անվանում / Հաշվեհամար",
        f"{constants.get('SupplierBankName','')}, Հաշվեհամար {constants.get('SupplierBankAccountNumber','')}",
        label_w=content_w * 0.28, val_w=content_w * 0.72,
    ))
    story.append(_field_row("[12] Գտնվելու վայրը (հասցե)", constants.get("SupplierAddress"), label_w=content_w * 0.28, val_w=content_w * 0.72))
    story.append(_field_row("[13] Վայրը, որտեղից ապրանքները մատակարարվում են (առաքվում)", constants.get("SupplyLocation"), label_w=content_w * 0.28, val_w=content_w * 0.72))
    story.append(Spacer(1, 4))

    # Buyer section ------------------------------------------------------
    story.append(_section_header("Ապրանքներ ձեռք բերող անձի տվյալներ", content_w))
    story.append(Table([[
        Paragraph("[15] Հարկ վճարողի հաշվառման համարը (ՀՎՀՀ)", small_b),
        _digit_boxes(buyer_info.get("tin"), 8),
    ]], colWidths=[content_w - 50 * mm, 50 * mm],
        style=TableStyle([("GRID", (0, 0), (-1, -1), 0.6, colors.black), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 4)])))
    story.append(_field_row("[17] Անվանում", buyer_info.get("name"), label_w=content_w * 0.28, val_w=content_w * 0.72))
    story.append(_field_row("[22] Ապրանքների մատակարարման (առաքման) եղանակը", parsed.get("driver"), label_w=content_w * 0.28, val_w=content_w * 0.72))
    story.append(_field_row("[23] Մատակարարվող ապրանքների նշանակման վայրը (հասցե)", buyer_info.get("address"), label_w=content_w * 0.28, val_w=content_w * 0.72))
    story.append(Spacer(1, 5))

    # Goods table [24] -----------------------------------------------------
    hdr_style = ParagraphStyle("gh", fontName="Arm-Bold", fontSize=6.8, alignment=TA_CENTER, leading=8)
    name_style = ParagraphStyle("gn", fontName="Arm", fontSize=7.2, alignment=TA_LEFT, leading=8.5)
    cell_style = ParagraphStyle("gc", fontName="Arm", fontSize=7.2, alignment=TA_CENTER, leading=8.5)
    bold_cell_style = ParagraphStyle("gcb", fontName="Arm-Bold", fontSize=7.2, alignment=TA_CENTER, leading=8.5)

    def H(t):
        return Paragraph(t.replace("\n", "<br/>"), hdr_style)

    headers = [
        H("N"), H("ԱՏԳ ԱԱ\nդասակարգիչ"), H("Ապրանքի անվանումը"), H("Չափման\nմիավոր"), H("Քանակ"),
        H("Միավորի\nգին"), H("Զեղչ\n(%)"), H("Արժեք"), H("Տուրք*"), H("ԱԱՀ\nդրույքաչափ (%)"),
        H("ԱԱՀ\nգումար"), H("Ընդհանուր"),
    ]
    # Product name gets much more room than the narrow numeric columns.
    col_widths = [7, 14, 62, 12, 14, 16, 10, 18, 12, 15, 16, 18]
    scale = content_w / sum(col_widths)
    col_widths = [w * scale for w in col_widths]

    def fmt(v):
        if v is None:
            return ""
        if isinstance(v, float) and v == int(v):
            return f"{int(v):,}"
        if isinstance(v, float):
            return f"{v:,.5f}".rstrip("0").rstrip(".")
        return str(v)

    data = [headers]
    for i, g in enumerate(goods, start=1):
        data.append([
            Paragraph(str(i), cell_style), Paragraph(str(g["code"]), cell_style),
            Paragraph(g["name"], name_style), Paragraph(g["unit"], cell_style),
            Paragraph(fmt(g["qty"]), cell_style), Paragraph(fmt(g["net_unit"]), cell_style),
            Paragraph("", cell_style), Paragraph(fmt(g["price"]), cell_style),
            Paragraph("", cell_style), Paragraph(fmt(constants.get("VATRate", 20)), cell_style),
            Paragraph(fmt(g["vat"]), cell_style), Paragraph(fmt(g["total_price"]), cell_style),
        ])
    tot_price = sum(g["price"] for g in goods)
    tot_vat = sum(g["vat"] for g in goods)
    tot_total = sum(g["total_price"] for g in goods)
    data.append([
        "", "", Paragraph("Ընդամենը", bold_cell_style), "", "", "",
        Paragraph("X", bold_cell_style), Paragraph(fmt(round(tot_price, 2)), bold_cell_style),
        Paragraph("X", bold_cell_style), Paragraph("X", bold_cell_style),
        Paragraph(fmt(round(tot_vat, 2)), bold_cell_style), Paragraph(fmt(round(tot_total, 2)), bold_cell_style),
    ])

    goods_table = Table(data, colWidths=col_widths, repeatRows=1)
    goods_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("BACKGROUND", (0, -1), (-1, -1), colors.whitesmoke),
        ("SPAN", (0, -1), (1, -1)),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(_section_header("Մատակարարվող (առաքվող) ապրանքների բնութագիրը եւ վճարման ենթակա գումարի հաշվարկը", content_w))
    story.append(goods_table)
    story.append(Spacer(1, 10))

    sign = Table([[
        Paragraph("Առաքող՝ _______________________  (անուն, ազգանունը, պաշտոնը եւ ստորագրությունը)", small),
        Paragraph("Ստացող՝ _______________________  (անուն, ազգանունը, պաշտոնը եւ ստորագրությունը)", small),
    ]], colWidths=[content_w / 2, content_w / 2])
    sign.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story.append(sign)

    doc.build(story)
    return out_path
