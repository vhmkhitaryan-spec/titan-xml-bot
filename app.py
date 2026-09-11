import os
import io
import json
import math
import re

import openpyxl
import requests
from flask import Flask, request, jsonify
from lxml import etree
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

from pdf_invoice import build_invoice_pdf

app = Flask(__name__)

BOT_TOKEN = os.environ["BOT_TOKEN"]
DRIVE_FILE_ID = os.environ["DRIVE_FILE_ID"]
TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"

NS = "http://www.taxservice.am/tp3/invoice/definitions"
NSMAP = {"ns1": NS}

INVOICE_HEADER = "Հաշիվ-ապրանքագիր"
TG_DIVIDER_CHAR = "⠀"

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


# ---------------------------------------------------------------------------
# 1. Parsing the forwarded Telegram invoice text
#    (mirrors build3.py's buildInvoiceTelegramText exactly)
# ---------------------------------------------------------------------------

def _num(s):
    return float(s.replace(",", ""))


def parse_invoice_text(text):
    """Returns a dict {date, buyer, driver, products:[{name,qty,unit,gross_price}], total}
    or None if this text is not a (non-cancelled) invoice message."""
    lines = text.split("\n")
    if not lines or lines[0].strip() != INVOICE_HEADER:
        return None

    idx = 1
    m = re.match(r"Ամսաթիվ՝\s*(\d{4}-\d{2}-\d{2})", lines[idx])
    if not m:
        raise ValueError("Ամսաթվի տողը չգտնվեց")
    date = m.group(1)
    idx += 1

    m = re.match(r"Գնորդ՝\s*(.+)", lines[idx])
    if not m:
        raise ValueError("Գնորդի տողը չգտնվեց")
    buyer = m.group(1).strip()
    idx += 1

    m = re.match(r"Վարորդ՝\s*(.+)", lines[idx])
    if not m:
        raise ValueError("Վարորդի տողը չգտնվեց")
    driver = re.sub(r"\s*⚠️\s*$", "", m.group(1).strip()).strip()
    idx += 1

    # blank line before the product list
    idx += 1

    products = []
    while idx < len(lines):
        line = lines[idx]
        if line != "" and set(line) == {TG_DIVIDER_CHAR}:
            break  # reached the width-forcing divider before the total
        if line.startswith("➤ "):
            name = line[2:].strip()
            idx += 1  # blank line
            idx += 1  # calc line
            calc_line = lines[idx]
            cm = re.match(
                r"\s*([\d,]+(?:\.\d+)?)(?:\s+(\S+))?\s+x\s+([\d,]+(?:\.\d+)?)\s*=\s*([\d,]+\.\d{2})\s*դր\.",
                calc_line,
            )
            if not cm:
                raise ValueError(f"Հաշվարկի տողը չհասկացվեց՝ {calc_line!r}")
            qty = _num(cm.group(1))
            unit = cm.group(2) or ""
            gross_price = _num(cm.group(3))
            products.append({"name": name, "qty": qty, "unit": unit, "gross_price": gross_price})
            idx += 1
            if idx < len(lines) and lines[idx] == "":
                idx += 1
        else:
            idx += 1

    while idx < len(lines) and set(lines[idx]) != {TG_DIVIDER_CHAR}:
        idx += 1
    idx += 1  # skip the divider line itself

    m = re.match(r"Ընդամենը՝\s*([\d,]+\.\d{2})\s*դր\.", lines[idx])
    if not m:
        raise ValueError("Ընդամենի տողը չգտնվեց")
    total = _num(m.group(1))

    if not products:
        raise ValueError("Ապրանքի ոչ մի տող չգտնվեց")

    return {"date": date, "buyer": buyer, "driver": driver, "products": products, "total": total}


# ---------------------------------------------------------------------------
# 2. Price conversion (mirrors Sheet6's AF/AG reference table exactly:
#    net = ROUNDUP(gross / 1.2, 5 decimal places))
# ---------------------------------------------------------------------------

def gross_to_net_unit_price(gross):
    factor = 10 ** 5
    return math.ceil(gross / 1.2 * factor - 1e-9) / factor


def compute_good(qty, gross_price):
    net_unit = gross_to_net_unit_price(gross_price)
    price = round(qty * net_unit, 2)
    vat = round(price * 0.20, 2)
    total_price = round(price + vat, 2)
    return net_unit, price, vat, total_price


# ---------------------------------------------------------------------------
# 3. Reference data (Google Drive: Products / Counterparties / Constants)
# ---------------------------------------------------------------------------

def get_drive_service():
    info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def download_reference_workbook():
    service = get_drive_service()
    req = service.files().get_media(fileId=DRIVE_FILE_ID)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return openpyxl.load_workbook(buf, data_only=True)


def read_constants(ws):
    d = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        key = str(row[0]).split(" (")[0].strip()
        d[key] = row[1]
    return d


# Telegram messages use the short, everyday company name (e.g. "ՊՐԵՄԻՈՒՄ ՊՐԻՆՏ
# ՍՊԸ"), while the reference file carries the full legal name as registered
# ("«ՊՐԵՄԻՈՒՄ ՊՐԻՆՏ» Սահմանափակ պատասխանատվությամբ ընկերություն (ՍՊԸ)"). Strip
# the quoting and the legal-form boilerplate from both sides before comparing,
# so the two forms of the same name match.
_LEGAL_FORMS = [
    "Սահմանափակ պատասխանատվությամբ ընկերություն",
    "Փակ բաժնետիրական ընկերություն",
    "Բաց բաժնետիրական ընկերություն",
    "Հասարակական կազմակերպություն",
    "Անհատ ձեռնարկատեր",
]


def normalize_name(name):
    n = str(name).strip()
    n = n.replace("«", "").replace("»", "")
    for lf in _LEGAL_FORMS:
        n = n.replace(lf, "")
    n = n.replace("(", "").replace(")", "")
    # Strip a standalone "ԱՁ" token wherever it appears. On the Telegram side
    # it's a prefix ("ԱՁ ԱՆՈՒՆ"), while on the reference-file side it's a
    # leftover suffix token left after the "Անհատ ձեռնարկատեր" phrase above
    # is removed ("ԱՆՈՒՆ ... (ԱՁ)" -> "ԱՆՈՒՆ ... ԱՁ"). Removing it from both
    # sides (rather than trying to reposition it) makes them match.
    n = re.sub(r"\bԱՁ\b", "", n, flags=re.UNICODE)
    n = re.sub(r"\s+", " ", n).strip()
    return n.upper()


def lookup_product(ws, name):
    name_norm = normalize_name(name)
    for row in ws.iter_rows(min_row=2, values_only=True):
        pname, unit, code = (row + (None, None, None))[:3]
        if pname and normalize_name(pname) == name_norm:
            return {"unit": unit, "code": code}
    return None


def lookup_counterparty(ws, name):
    name_norm = normalize_name(name)
    for row in ws.iter_rows(min_row=2, values_only=True):
        cname, _tin_raw, tin_padded, address = (row + (None, None, None, None))[:4]
        if cname and normalize_name(cname) == name_norm:
            return {"tin": tin_padded, "name": cname, "address": address}
    return None


# ---------------------------------------------------------------------------
# 4. XML building (namespace ns1, exact structure from xl/xmlMaps.xml)
# ---------------------------------------------------------------------------

def _fmt_num(value):
    """Render a whole-valued float as '525' rather than '525.0', while keeping
    genuine fractional values (e.g. quantities, the ROUNDUP'd unit price) as-is."""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value)


def _add(parent, tag, value):
    el = etree.SubElement(parent, "{%s}%s" % (NS, tag))
    if value is None:
        el.text = ""
    elif isinstance(value, float):
        el.text = _fmt_num(value)
    else:
        el.text = str(value)
    return el


def build_xml(parsed, constants, buyer_info, goods):
    root = etree.Element("{%s}ExportedData" % NS, nsmap=NSMAP)
    invoice = etree.SubElement(root, "{%s}Invoice" % NS)
    invoice.set("Version", str(constants.get("InvoiceVersion", 1)))

    _add(invoice, "Type", constants.get("InvoiceType", 1))

    general = etree.SubElement(invoice, "{%s}GeneralInfo" % NS)
    _add(general, "SupplyDate", parsed["date"])
    _add(general, "Procedure", constants.get("GeneralInfoProcedure", 2))
    _add(general, "AdjustmentAccount", str(bool(constants.get("AdjustmentAccount", False))).lower())

    supplier = etree.SubElement(invoice, "{%s}SupplierInfo" % NS)
    _add(supplier, "VATNumber", constants.get("SupplierVATNumber"))
    taxpayer = etree.SubElement(supplier, "{%s}Taxpayer" % NS)
    _add(taxpayer, "TIN", constants.get("SupplierTIN"))
    _add(taxpayer, "Name", constants.get("SupplierName"))
    _add(taxpayer, "Address", constants.get("SupplierAddress"))
    bank = etree.SubElement(taxpayer, "{%s}BankAccount" % NS)
    _add(bank, "BankName", constants.get("SupplierBankName"))
    _add(bank, "BankAccountNumber", constants.get("SupplierBankAccountNumber"))
    _add(supplier, "SupplyLocation", constants.get("SupplyLocation"))

    buyer = etree.SubElement(invoice, "{%s}BuyerInfo" % NS)
    buyer_taxpayer = etree.SubElement(buyer, "{%s}Taxpayer" % NS)
    _add(buyer_taxpayer, "TIN", buyer_info["tin"])
    _add(buyer_taxpayer, "Name", buyer_info["name"])
    _add(buyer_taxpayer, "TinNotRequired", str(bool(constants.get("BuyerTinNotRequired", False))).lower())
    _add(buyer, "DeliveryMethod", parsed["driver"])
    _add(buyer, "DeliveryLocation", buyer_info["address"])

    goods_info = etree.SubElement(invoice, "{%s}GoodsInfo" % NS)
    total_price = total_vat = total_total = 0.0
    vat_rate = constants.get("VATRate", 20)
    for g in goods:
        good = etree.SubElement(goods_info, "{%s}Good" % NS)
        _add(good, "Description", g["name"])
        _add(good, "ClassifierCode", g["code"])
        _add(good, "Unit", g["unit"])
        _add(good, "Amount", g["qty"])
        _add(good, "PricePerUnit", g["net_unit"])
        _add(good, "Price", g["price"])
        _add(good, "VATRate", vat_rate)
        _add(good, "VAT", g["vat"])
        _add(good, "TotalPrice", g["total_price"])
        total_price += g["price"]
        total_vat += g["vat"]
        total_total += g["total_price"]

    total_el = etree.SubElement(goods_info, "{%s}Total" % NS)
    _add(total_el, "Price", round(total_price, 2))
    _add(total_el, "VAT", round(total_vat, 2))
    _add(total_el, "TotalPrice", round(total_total, 2))

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", pretty_print=True)


# ---------------------------------------------------------------------------
# 5. Telegram helpers
# ---------------------------------------------------------------------------

def tg_send_message(chat_id, text):
    requests.post(f"{TELEGRAM_API}/sendMessage", json={"chat_id": chat_id, "text": text}, timeout=20)


def tg_send_document(chat_id, filename, content_bytes):
    requests.post(
        f"{TELEGRAM_API}/sendDocument",
        data={"chat_id": chat_id},
        files={"document": (filename, content_bytes, "application/xml")},
        timeout=30,
    )


# ---------------------------------------------------------------------------
# 6. Webhook
# ---------------------------------------------------------------------------

@app.route("/webhook", methods=["POST"])
def webhook():
    update = request.get_json(force=True, silent=True) or {}
    message = update.get("message") or update.get("channel_post")
    if not message:
        return jsonify(ok=True)

    chat_id = message["chat"]["id"]
    text = message.get("text", "")

    # Only ever act on genuinely FORWARDED messages (see design notes: a Reply
    # does not expose a bot-authored message's content cross-bot, a Forward does).
    if "forward_origin" not in message and "forward_from" not in message and "forward_date" not in message:
        return jsonify(ok=True)

    try:
        parsed = parse_invoice_text(text)
        if not parsed:
            return jsonify(ok=True)  # e.g. a cancellation message - nothing to do

        wb = download_reference_workbook()
        constants = read_constants(wb["Constants"])

        buyer_info = lookup_counterparty(wb["Counterparties"], parsed["buyer"])
        if not buyer_info:
            tg_send_message(chat_id, f"Չգտա գործընկերոջը reference ֆայլում՝ {parsed['buyer']}")
            return jsonify(ok=True)

        goods = []
        for p in parsed["products"]:
            info = lookup_product(wb["Products"], p["name"])
            if not info:
                tg_send_message(chat_id, f"Չգտա ապրանքը reference ֆայլում՝ {p['name']}")
                return jsonify(ok=True)
            net_unit, price, vat, total_price = compute_good(p["qty"], p["gross_price"])
            goods.append(
                {
                    "name": p["name"],
                    "unit": info["unit"],
                    "code": info["code"],
                    "qty": p["qty"],
                    "net_unit": net_unit,
                    "price": price,
                    "vat": vat,
                    "total_price": total_price,
                }
            )

        xml_bytes = build_xml(parsed, constants, buyer_info, goods)
        xml_filename = f"invoice-{parsed['date']}.xml"
        tg_send_document(chat_id, xml_filename, xml_bytes)

        pdf_buf = io.BytesIO()
        build_invoice_pdf(parsed, constants, buyer_info, goods, pdf_buf)
        pdf_filename = f"invoice-{parsed['date']}.pdf"
        tg_send_document(chat_id, pdf_filename, pdf_buf.getvalue())
    except Exception as e:
        tg_send_message(chat_id, f"Սխալ. {e}")

    return jsonify(ok=True)


@app.route("/", methods=["GET"])
def health():
    return "ok"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
