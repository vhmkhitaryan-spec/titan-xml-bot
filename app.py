"""
Titan XML Bot
-------------
A forwarded Titan invoice message becomes a DRAFT goods tax invoice in the
ՊԵԿ e-invoicing system, and the bot replies with a short summary + the
draft's official PDF.

Flow (per forwarded message, strictly in order):
  parse text -> Excel lookups -> build XML -> XSD check (invoice.xsd)
  -> token valid?  (read from the token's own `exp`, never by trying ՊԵԿ)
       no  -> e-mail "TITAN-TOKEN" to the iPhone (the Shortcut logs in and
              POSTs the token to /token), wait; after 5 min tell the chat,
              keep waiting
       yes -> goods-validate-model -> goods-create-draft -> goods-pdf-generate
  -> summary + PDF to the chat

Queue: Telegram itself. Updates are read with getUpdates and an update is
confirmed (offset advanced) only after it is fully handled, so nothing is
lost if Render restarts; Telegram keeps unconfirmed updates for 24 h.
"""
import base64
import datetime
import io
import json
import math
import os
import re
import smtplib
import threading
import time
import uuid
from email.message import EmailMessage

import openpyxl
import requests
from flask import Flask, request
from lxml import etree
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

app = Flask(__name__)

# ---------------------------------------------------------------------------
# 0. Configuration (Render -> Environment)
# ---------------------------------------------------------------------------
BOT_TOKEN = os.environ["BOT_TOKEN"]
DRIVE_FILE_ID = os.environ["DRIVE_FILE_ID"]
TELEGRAM_API = f"https://api.telegram.org/bot{BOT_TOKEN}"
TOKEN_SECRET = os.environ.get("TOKEN_SECRET", "")

# E-mail that wakes the iPhone. Either RESEND_API_KEY (HTTPS API, works on
# Render free) or SMTP (MAIL_FROM + MAIL_APP_PASSWORD, e.g. Gmail app password).
MAIL_TO = os.environ.get("MAIL_TO", "")
MAIL_FROM = os.environ.get("MAIL_FROM", "")
MAIL_APP_PASSWORD = os.environ.get("MAIL_APP_PASSWORD", "")
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))
RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
MAIL_SUBJECT = "TITAN-TOKEN"

PEK_API = "https://e-invoicing.taxservice.am/api"
XSD_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "invoice.xsd")

NS = "http://www.taxservice.am/tp3/invoice/definitions"
NSMAP = {"ns1": NS}
YEREVAN = datetime.timezone(datetime.timedelta(hours=4))

TOKEN_USE_MARGIN = 60          # use a token only if it has > 60 s left
TOKEN_EXPIRY_GRACE = 15        # ask for a new one only 15 s after exp has passed
MAIL_REPEAT_AFTER = 60 * 60    # at most one e-mail per hour while waiting
BLOCK_AFTER_0032 = 60 * 60     # ՊԵԿ says a token is still valid: wait up to 1 h

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
# 5. XSD check (official schema from the ՊԵԿ integration package)
# ---------------------------------------------------------------------------
_XSD = None


def xsd_errors(xml_bytes):
    """Returns a list of human-readable schema errors ([] = valid)."""
    global _XSD
    if _XSD is None:
        _XSD = etree.XMLSchema(etree.parse(XSD_PATH))
    doc = etree.fromstring(xml_bytes)
    if _XSD.validate(doc):
        return []
    return [f"տող {e.line}: {e.message}" for e in _XSD.error_log][:10]


# ---------------------------------------------------------------------------
# 6. Token store. Render decides validity from the token's own `exp`;
#    it never "tries" ՊԵԿ with a token it knows is expired.
# ---------------------------------------------------------------------------
STATE_LOCK = threading.Lock()
STATE = {
    "token": None,
    "received": 0.0,
    "exp": 0.0,
    "blocked_until": 0.0,     # set when ՊԵԿ answered 0032 (a token we lost is still valid)
    "requested_at": 0.0,      # when the TITAN-TOKEN mail was last sent
    "login_events": [],       # ՊԵԿ login answers relayed by /token, shown in the invoice log
}
TOKEN_EVENT = threading.Event()


def _jwt_exp(token):
    try:
        part = token.split(".")[1]
        part += "=" * (-len(part) % 4)
        return float(json.loads(base64.urlsafe_b64decode(part)).get("exp", 0))
    except Exception:
        return 0.0


def _fmt_time(ts):
    return datetime.datetime.fromtimestamp(ts, YEREVAN).strftime("%H:%M")


def usable_token():
    with STATE_LOCK:
        t, exp = STATE["token"], STATE["exp"]
    if t and exp - time.time() > TOKEN_USE_MARGIN:
        return t
    return None


def may_request_new_token():
    """True only when ՊԵԿ will accept a new login: the known token's exp has
    passed (plus grace) and we are not inside a 0032 block."""
    now = time.time()
    with STATE_LOCK:
        t, exp, blocked = STATE["token"], STATE["exp"], STATE["blocked_until"]
    if now < blocked:
        return False
    if t and now < exp + TOKEN_EXPIRY_GRACE:
        return False
    return True


def next_attempt_at():
    """When the bot will next ask the iPhone for a token."""
    with STATE_LOCK:
        t, exp, blocked, req = STATE["token"], STATE["exp"], STATE["blocked_until"], STATE["requested_at"]
    candidates = [time.time()]
    if blocked:
        candidates.append(blocked)
    if t:
        candidates.append(exp + TOKEN_EXPIRY_GRACE)
    if req:
        candidates.append(req + MAIL_REPEAT_AFTER)
    return max(candidates)


def drop_token():
    with STATE_LOCK:
        STATE["token"] = None


# ---------------------------------------------------------------------------
# 7. E-mail that triggers the iPhone automation
# ---------------------------------------------------------------------------
def send_token_mail():
    body = "Titan XML Bot-ը ՊԵԿ token է խնդրում։ Այս նամակը iPhone-ի automation-ի համար է։"
    if RESEND_API_KEY:
        r = requests.post(
            "https://api.resend.com/emails",
            headers={"Authorization": f"Bearer {RESEND_API_KEY}"},
            json={
                "from": MAIL_FROM or "Titan Bot <onboarding@resend.dev>",
                "to": [MAIL_TO],
                "subject": MAIL_SUBJECT,
                "text": body,
            },
            timeout=20,
        )
        r.raise_for_status()
        return
    msg = EmailMessage()
    msg["From"] = MAIL_FROM
    msg["To"] = MAIL_TO
    msg["Subject"] = MAIL_SUBJECT
    msg.set_content(body)
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=20) as s:
        s.login(MAIL_FROM, MAIL_APP_PASSWORD)
        s.send_message(msg)


# ---------------------------------------------------------------------------
# 8. ՊԵԿ e-invoicing API
# ---------------------------------------------------------------------------
class Progress:
    """ONE Telegram message per invoice, a reply to the forwarded invoice.
    It starts as a document message (a small placeholder file) whose caption
    is the step log; every step edits the caption, and at the end the
    placeholder is replaced by the draft's PDF IN THE SAME MESSAGE, so the
    PDF always stays right under its own invoice."""

    CAPTION_LIMIT = 1000  # Telegram allows 1024 UTF-16 units; emoji count double

    def __init__(self, chat_id, title, reply_to=None):
        self.chat_id, self.lines, self.message_id = chat_id, [title], None
        self.reply_to, self.summary = reply_to, None
        self._send_placeholder()

    def _caption(self):
        head, steps = self.lines[:1], self.lines[1:]
        tail = ["", self.summary] if self.summary else []
        while True:
            text = "\n".join(head + steps + tail)
            if len(text) <= self.CAPTION_LIMIT or not steps:
                return text[: self.CAPTION_LIMIT]
            steps = ["\u2026"] + steps[2:] if steps[0] == "\u2026" else ["\u2026"] + steps[1:]

    def _send_placeholder(self):
        data = {"chat_id": self.chat_id, "caption": self._caption()}
        if self.reply_to:
            data["reply_parameters"] = json.dumps({"message_id": self.reply_to,
                                                   "allow_sending_without_reply": True})
        try:
            r = requests.post(
                f"{TELEGRAM_API}/sendDocument",
                data=data,
                files={"document": ("մշակվում-է.txt", "Titan XML Bot՝ մշակվում է…".encode("utf-8"), "text/plain")},
                timeout=30,
            ).json()
            self.message_id = (r.get("result") or {}).get("message_id")
        except Exception:
            pass

    def _push(self):
        if self.message_id is None:
            return
        try:
            requests.post(f"{TELEGRAM_API}/editMessageCaption",
                          json={"chat_id": self.chat_id, "message_id": self.message_id,
                                "caption": self._caption()},
                          timeout=20)
        except Exception:
            pass

    def step(self, text):
        self.lines.append(f"{datetime.datetime.now(YEREVAN):%H:%M:%S}  {text}")
        self._push()

    def _replace_file(self, filename, content, mime):
        if self.message_id is None:
            return False
        media = {"type": "document", "media": "attach://file", "caption": self._caption()}
        try:
            r = requests.post(
                f"{TELEGRAM_API}/editMessageMedia",
                data={"chat_id": self.chat_id, "message_id": self.message_id, "media": json.dumps(media)},
                files={"file": (filename, content, mime)},
                timeout=60,
            ).json()
            return bool(r.get("ok"))
        except Exception:
            return False

    def finish_pdf(self, filename, pdf, summary):
        """Same message: placeholder file -> the draft PDF, log + summary as caption."""
        self.summary = summary
        return self._replace_file(filename, pdf, "application/pdf")

    def finish_without_pdf(self):
        """Same message: placeholder file -> the log itself, so no 'processing' file is left."""
        self._replace_file("log.txt", "\n".join(self.lines).encode("utf-8"), "text/plain")


class _NoProgress:
    def step(self, text):
        pass


class PekError(Exception):
    def __init__(self, code, message, auth=False):
        super().__init__(f"{code}: {message}")
        self.code, self.message, self.auth = code, message, auth


class PekTransient(Exception):
    """Network trouble / 5xx: retry later, the message stays in the queue."""


def pek_call(token, path, payload, timeout=40):
    try:
        r = requests.post(
            f"{PEK_API}/{path}",
            json={"payload": payload},
            headers={"accept": "application/json"},
            cookies={"jwt-auth-token": token},
            timeout=timeout,
        )
    except requests.RequestException as e:
        raise PekTransient(type(e).__name__)
    if r.status_code in (401, 403):
        raise PekError(str(r.status_code), "մուտքը մերժվեց", auth=True)
    if r.status_code >= 500:
        raise PekTransient(f"HTTP {r.status_code}")
    try:
        data = r.json()
    except ValueError:
        raise PekTransient(f"HTTP {r.status_code}, ոչ JSON պատասխան")
    if not data.get("ok", False):
        f = data.get("failure") or {}
        code, message = str(f.get("code", "?")), str(f.get("message", ""))
        auth = any(k in code.lower() for k in ("auth", "token", "session", "login"))
        raise PekError(code, message, auth=auth)
    return data.get("payload")


_CLASSIFIERS = {}


def _code_key(v):
    """Excel numbers come back as floats (4810.0); ՊԵԿ codes may contain
    spaces or dots. Normalise to bare characters: '4810'."""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"[\s.\-]", "", str(v))


_CLASSIFIER_LISTS = ("dictionaries/classifier-list",
                     "dictionaries/classifier-full-list",
                     "dictionaries/item-classifier-list")
_CLASSIFIER_SIZES = {}


def _load_classifiers(token):
    for i, path in enumerate(_CLASSIFIER_LISTS):
        try:
            data = pek_call(token, path, {}) or []
        except PekTransient:
            if i == 0:
                raise          # main list unavailable: retry the invoice later
            continue
        except PekError:
            continue
        if isinstance(data, dict):
            data = next((v for v in data.values() if isinstance(v, list)), [])
        _CLASSIFIER_SIZES[path.split("/")[-1]] = len(data)
        for c in data:
            if isinstance(c, dict) and c.get("code") and c.get("id"):
                _CLASSIFIERS.setdefault(_code_key(c["code"]), c["id"])


def classifier_id(token, code):
    if not _CLASSIFIERS:
        _load_classifiers(token)
    key = _code_key(code)
    if key in _CLASSIFIERS:
        return _CLASSIFIERS[key]
    # e.g. the dictionary keeps a longer code that starts with ours
    longer = sorted(k for k in _CLASSIFIERS if k.startswith(key))
    return _CLASSIFIERS[longer[0]] if longer else None


def classifier_hint(code):
    """What the dictionaries hold near a missing code (for the Telegram log)."""
    key = _code_key(code)
    near = sorted(k for k in _CLASSIFIERS if k[:2] == key[:2])[:8]
    sizes = ", ".join(f"{k}={v}" for k, v in _CLASSIFIER_SIZES.items()) or "ցուցակ չստացվեց"
    return f"ՊԵԿ ցուցակներ՝ {sizes}; մոտ կոդեր՝ {', '.join(near) or 'չկան'}"


def _money(v):
    return round(float(v), 2)


# XML <Procedure> (invoice.xsd InvoiceProcedureType) -> API behalfOf.
# The working XML uses Constants.GeneralInfoProcedure, so the API draft takes
# the very same value: 2 = "on behalf of the taxpayer" -> TAXPAYER.
_PROCEDURE_TO_BEHALF = {2: "TAXPAYER", 3: "SUPPLIER", 4: "PRINCIPAL", 6: "JOINT_CASE_MNG"}


def behalf_of(constants):
    explicit = constants.get("BehalfOf")
    if explicit:
        return str(explicit).strip()
    try:
        proc = int(constants.get("GeneralInfoProcedure", 2))
    except (TypeError, ValueError):
        proc = 2
    if proc not in _PROCEDURE_TO_BEHALF:
        raise PekError("procedure", f"GeneralInfoProcedure={proc}-ի համար API-ի համարժեք չկա")
    return _PROCEDURE_TO_BEHALF[proc]


def build_pek_draft(token, parsed, constants, buyer_info, goods, doc_id):
    """Maps our data onto the goods-create-draft model.
    Field choices marked (?) are best guesses from the package; the
    goods-validate-model step reports anything ՊԵԿ disagrees with."""
    items = []
    missing = []
    for i, g in enumerate(goods, start=1):
        cid = classifier_id(token, g["code"]) if g.get("code") else None
        if g.get("code") and not cid:
            missing.append(_code_key(g["code"]))
        items.append({
            "mode": "new",
            "id": str(uuid.uuid4()),
            "invoiceId": doc_id,
            "seqNo": i,
            "classifierId": cid,
            "name": g["name"],
            "unit": g["unit"],
            "quantity": g["qty"],
            "unitPrice": g["net_unit"],
            "totalValue": _money(g["price"]),
            "vatRate": "VAT_20",
            "vatAmount": _money(g["vat"]),
            "total": _money(g["total_price"]),
        })
    if missing:
        miss = sorted(set(missing))
        raise PekError("classifier", "ՊԵԿ-ի դասակարգչում չգտնվեցին կոդերը՝ " + ", ".join(miss)
                       + ". " + classifier_hint(miss[0]))

    total_value = _money(sum(g["price"] for g in goods))
    total_vat = _money(sum(g["vat"] for g in goods))
    total = _money(sum(g["total_price"] for g in goods))

    entity = {
        "id": doc_id,
        "status": "DRAFT",
        "createdAt": datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z"),
        "deliveredAt": f"{parsed['date']}T00:00:00.000Z",
        "behalfOf": behalf_of(constants),
        "supplierTin": str(constants.get("SupplierTIN") or ""),
        "supplierVatTin": str(constants.get("SupplierVATNumber") or "") or None,
        "supplierName": constants.get("SupplierName"),
        "supplierAddress": constants.get("SupplierAddress"),
        "supplierBank": constants.get("SupplierBankName"),
        "supplierAccNo": str(constants.get("SupplierBankAccountNumber") or "") or None,
        "sourceFullAddress": constants.get("SupplyLocation"),
        "sourceOtherAddress": constants.get("SupplyLocation"),
        "buyerHasNoTin": False,
        "buyerIsNatural": False,
        "buyerTin": str(buyer_info["tin"]),
        "buyerName": buyer_info["name"],
        "buyerAddress": buyer_info["address"],
        "destinationFullAddress": buyer_info["address"],
        "destinationOtherAddress": buyer_info["address"],
        "deliveryMethod": parsed["driver"],
        "totalValue": total_value,
        "totalVatAmount": total_vat,
        "total": total,
        "source": "API",
        "finalUse": False,
        "hasCodes": False,
        "traceable": False,
    }
    entity = {k: v for k, v in entity.items() if v is not None}
    return entity, items


def validation_problems(result):
    """goods-validate-model answer -> list of problems ([] = passed)."""
    if not isinstance(result, dict):
        return []
    status = str(result.get("status") or "").upper()
    details = result.get("details")
    if any(w in status for w in ("ERROR", "FAIL", "INVALID", "REJECT")):
        return [f"{status}: {details}" if details else status]
    return []


def pek_draft_exists(token, doc_id):
    try:
        return bool(pek_call(token, "goods/goods-by-id", {"id": doc_id}))
    except PekError:
        return False


def pek_create_draft(token, parsed, constants, buyer_info, goods, doc_id, retry=False, log=_NoProgress()):
    # A retry after a network error may follow a create that actually succeeded:
    # the id is fixed per message, so check before creating it again.
    if retry and pek_draft_exists(token, doc_id):
        log.step("\u2714 Սևագիրը արդեն ստեղծված էր (նախորդ փորձից)")
        return doc_id
    entity, items = build_pek_draft(token, parsed, constants, buyer_info, goods, doc_id)
    log.step(f"\u2714 ՊԵԿ դասակարգիչ՝ բոլոր կոդերը գտնվեցին, behalfOf={entity['behalfOf']}")
    tin = entity["supplierTin"]

    check = pek_call(token, "goods/goods-validate-model", {"entity": entity, "items": items, "tin": tin})
    problems = validation_problems(check)
    if problems:
        raise PekError("validate", "; ".join(problems))
    status = check.get("status") if isinstance(check, dict) else None
    log.step(f"\u2714 ՊԵԿ վալիդացիան անցավ ({status or 'OK'})")

    pek_call(token, "goods/goods-create-draft", dict(entity, items=items))
    log.step("\u2714 Սևագիրը ստեղծվեց ՊԵԿ-ում")
    return entity["id"]


def pek_draft_pdf(token, doc_id):
    """The ՊԵԿ web app downloads PDFs with POST /api/dispatcher/pdf-export
    {ids, sortCol, sortAsc}; the response body is the PDF itself.
    (goods-pdf-generate only returns a path on ՊԵԿ's own disk.)"""
    try:
        r = requests.post(
            f"{PEK_API}/dispatcher/pdf-export",
            json={"payload": {"ids": [doc_id], "sortCol": "createdAt", "sortAsc": False}},
            headers={"accept": "application/pdf,application/json,*/*"},
            cookies={"jwt-auth-token": token},
            timeout=90,
        )
    except requests.RequestException as e:
        raise PekTransient(f"PDF՝ {type(e).__name__}")
    if r.status_code == 200 and r.content.startswith(b"%PDF"):
        return r.content
    detail = ""
    if "json" in r.headers.get("content-type", ""):
        try:
            f = r.json().get("failure") or {}
            detail = f"{f.get('code', '')}: {f.get('message', '')}"
        except ValueError:
            pass
    raise PekTransient(f"pdf-export → HTTP {r.status_code}, {r.headers.get('content-type', '?')}, "
                       f"{len(r.content)} բ {detail}".strip())


# ---------------------------------------------------------------------------
# 9. Telegram helpers
# ---------------------------------------------------------------------------
def tg(method, **params):
    r = requests.post(f"{TELEGRAM_API}/{method}", json=params, timeout=params.get("timeout", 0) + 20)
    return r.json()


def tg_send_message(chat_id, text):
    try:
        requests.post(f"{TELEGRAM_API}/sendMessage", json={"chat_id": chat_id, "text": text}, timeout=20)
    except requests.RequestException:
        pass


def tg_send_document(chat_id, filename, content_bytes, mime="application/pdf", caption=None, reply_to=None):
    data = {"chat_id": chat_id}
    if caption:
        data["caption"] = caption
    if reply_to:
        data["reply_parameters"] = json.dumps({"message_id": reply_to, "allow_sending_without_reply": True})
    requests.post(
        f"{TELEGRAM_API}/sendDocument",
        data=data,
        files={"document": (filename, content_bytes, mime)},
        timeout=60,
    )


def _fmt_amount(v):
    return f"{v:,.2f}".replace(",", " ")


# ---------------------------------------------------------------------------
# 10. One forwarded message -> prepared invoice (no ՊԵԿ calls here)
# ---------------------------------------------------------------------------
class Skip(Exception):
    """Message can't become a draft; tell the chat and move on."""


def prepare(text, log=_NoProgress()):
    parsed = parse_invoice_text(text)
    if not parsed:
        return None  # e.g. a cancellation message
    log.step(f"\u2714 Տեքստը կարդացվեց՝ {len(parsed['products'])} ապրանք, {_fmt_amount(parsed['total'])} դր.")

    wb = download_reference_workbook()
    constants = read_constants(wb["Constants"])
    log.step("\u2714 Excel-ը բեռնվեց Drive-ից")

    buyer_info = lookup_counterparty(wb["Counterparties"], parsed["buyer"])
    if not buyer_info:
        raise Skip(f"Չգտա գործընկերոջը reference ֆայլում՝ {parsed['buyer']}")
    log.step(f"\u2714 Գնորդը գտնվեց՝ ՀՎՀՀ {buyer_info['tin']}")

    goods = []
    for p in parsed["products"]:
        info = lookup_product(wb["Products"], p["name"])
        if not info:
            raise Skip(f"Չգտա ապրանքը reference ֆայլում՝ {p['name']}")
        net_unit, price, vat, total_price = compute_good(p["qty"], p["gross_price"])
        goods.append({
            "name": p["name"], "unit": info["unit"], "code": info["code"], "qty": p["qty"],
            "net_unit": net_unit, "price": price, "vat": vat, "total_price": total_price,
        })

    log.step(f"\u2714 Ապրանքները գտնվեցին՝ կոդեր {', '.join(sorted({_code_key(g['code']) for g in goods}))}")

    xml_bytes = build_xml(parsed, constants, buyer_info, goods)
    errs = xsd_errors(xml_bytes)
    if errs:
        raise Skip("XSD ստուգումը չանցավ՝\n" + "\n".join(errs))
    log.step("\u2714 XML-ը կազմվեց, XSD ստուգումն անցավ")

    return {"parsed": parsed, "constants": constants, "buyer": buyer_info, "goods": goods,
            "doc_id": str(uuid.uuid4())}


# ---------------------------------------------------------------------------
# 11. Worker: reads Telegram updates in order, confirms each only when done
# ---------------------------------------------------------------------------
WORKER = {"offset": None, "waiting_chat": None, "queue_len": 0, "handled_cmds": set()}


def _is_forward(msg):
    return any(k in msg for k in ("forward_origin", "forward_from", "forward_date"))


def _confirm(update_id):
    """Tell Telegram we are done with everything up to update_id."""
    WORKER["offset"] = update_id + 1
    try:
        tg("getUpdates", offset=WORKER["offset"], limit=1, timeout=0)
    except Exception:
        pass


def status_text():
    with STATE_LOCK:
        t, rec, exp, blocked = STATE["token"], STATE["received"], STATE["exp"], STATE["blocked_until"]
    lines = []
    if t and exp - time.time() > TOKEN_USE_MARGIN:
        lines.append(f"\u2705 Token-ը վավեր է մինչև {_fmt_time(exp)} (ստացվել է {_fmt_time(rec)})")
    elif rec:
        lines.append(f"\u274c Token-ը ժամկետանց է. վերջինը ստացվել է {_fmt_time(rec)}-ին")
    else:
        lines.append("\u274c Token դեռ չկա")
    if blocked > time.time():
        lines.append(f"ՊԵԿ-ը նոր token կտա {_fmt_time(blocked)}-ից հետո")
    if WORKER["waiting_chat"]:
        lines.append(f"Սպասում է token-ի, հերթում՝ {WORKER['queue_len']}")
    return "\n".join(lines)


def handle_command(msg):
    text = (msg.get("text") or "").strip()
    if text.startswith("/tokenstatus") or text.startswith("/status"):
        tg_send_message(msg["chat"]["id"], status_text())
        return True
    return False


def wait_for_token(chat_id, queue_len, log=_NoProgress()):
    """Blocks until a usable token exists. Sends the e-mail only when ՊԵԿ will
    accept a new login, at most once per hour. Every ՊԵԿ login answer and the
    time of the next attempt go into the invoice log."""
    WORKER["waiting_chat"], WORKER["queue_len"] = chat_id, queue_len
    mail_error_told = False
    with STATE_LOCK:
        STATE["login_events"].clear()
    if may_request_new_token():
        log.step("\u23f3 Token չկա կամ ժամկետանց է")
    else:
        log.step(f"\u23f3 Token չկա. Հաջորդ փորձը՝ {_fmt_time(next_attempt_at())}")
    try:
        while True:
            with STATE_LOCK:
                events, STATE["login_events"] = STATE["login_events"], []
            for text in events:
                log.step(text)
                if not usable_token():
                    log.step(f"\u21bb Հաջորդ փորձը՝ {_fmt_time(next_attempt_at())}")
            if usable_token():
                with STATE_LOCK:
                    exp = STATE["exp"]
                log.step(f"\u2714 Token-ը ստացվեց, վավեր է մինչև {_fmt_time(exp)}")
                return
            now = time.time()
            if may_request_new_token():
                with STATE_LOCK:
                    last = STATE["requested_at"]
                if now - last > MAIL_REPEAT_AFTER:
                    try:
                        send_token_mail()
                        with STATE_LOCK:
                            STATE["requested_at"] = now
                        log.step("\u2709 Email-ը ուղարկվեց iPhone-ին")
                    except Exception as e:
                        if not mail_error_told:
                            log.step(f"\u26a0 Email-ը չուղարկվեց՝ {e}. Կփորձեմ 5 րոպեն մեկ")
                            mail_error_told = True
                        TOKEN_EVENT.wait(300)
                        TOKEN_EVENT.clear()
                        continue
            TOKEN_EVENT.wait(10)
            TOKEN_EVENT.clear()
            if not usable_token():
                yield_commands()
    finally:
        WORKER["waiting_chat"] = None


def yield_commands():
    """While waiting, still answer /tokenstatus sent after the blocked message
    (without confirming anything)."""
    try:
        params = {"timeout": 0, "allowed_updates": ["message", "channel_post"]}
        if WORKER["offset"] is not None:
            params["offset"] = WORKER["offset"]
        res = tg("getUpdates", **params)
    except Exception:
        return
    for u in res.get("result", []):
        msg = u.get("message") or u.get("channel_post")
        if not msg or u["update_id"] in WORKER["handled_cmds"]:
            continue
        if (msg.get("text") or "").startswith("/") and handle_command(msg):
            WORKER["handled_cmds"].add(u["update_id"])


def process_invoice(msg, prepared, pending_after, log=_NoProgress()):
    chat_id = msg["chat"]["id"]
    parsed = prepared["parsed"]
    retry = False
    while True:
        token = usable_token()
        if not token:
            wait_for_token(chat_id, pending_after + 1, log)
            continue
        try:
            doc_id = pek_create_draft(token, parsed, prepared["constants"], prepared["buyer"],
                                      prepared["goods"], prepared["doc_id"], retry=retry, log=log)
        except PekTransient as e:
            log.step(f"\u26a0 ՊԵԿ-ը չպատասխանեց ({e}), կփորձեմ 1 րոպեից")
            retry = True
            time.sleep(60)
            continue
        except PekError as e:
            if e.auth:
                log.step("\u26a0 ՊԵԿ-ը token-ը չընդունեց, նորն եմ խնդրում")
                drop_token()
                continue
            raise Skip(f"ՊԵԿ-ը սևագիրը չընդունեց՝ {e.code}: {e.message}")
        return chat_id, doc_id


def worker_loop():
    # Switch Telegram from webhook to getUpdates. Pending updates are kept.
    try:
        tg("deleteWebhook", drop_pending_updates=False)
    except Exception:
        pass

    while True:
        try:
            params = {"timeout": 50, "allowed_updates": ["message", "channel_post"]}
            if WORKER["offset"] is not None:
                params["offset"] = WORKER["offset"]
            res = tg("getUpdates", **params)
            updates = res.get("result", []) if res.get("ok") else []
        except Exception:
            time.sleep(5)
            continue

        for idx, u in enumerate(updates):
            uid = u["update_id"]
            msg = u.get("message") or u.get("channel_post")
            if not msg:
                _confirm(uid)
                continue

            text = msg.get("text") or ""
            if text.startswith("/"):
                if uid not in WORKER["handled_cmds"]:
                    handle_command(msg)
                WORKER["handled_cmds"].discard(uid)
                _confirm(uid)
                continue

            if not _is_forward(msg):
                _confirm(uid)
                continue

            chat_id = msg["chat"]["id"]
            log = Progress(chat_id, "\U0001f4e5 Հաշիվը ստացվեց, մշակում եմ", reply_to=msg.get("message_id"))
            try:
                prepared = prepare(text, log)
                if not prepared:
                    _confirm(uid)
                    continue
                pending_after = sum(
                    1 for x in updates[idx + 1:]
                    if _is_forward(x.get("message") or x.get("channel_post") or {})
                )
                chat_id, doc_id = process_invoice(msg, prepared, pending_after, log)
            except (Skip, ValueError) as e:
                log.step(f"\u274c {e}")
                log.step("Սևագիր չի ստեղծվել")
                log.finish_without_pdf()
                _confirm(uid)
                continue
            except Exception as e:
                # Unexpected (Drive down etc.): keep it in the queue, retry later.
                log.step(f"\u26a0 Ժամանակավոր սխալ՝ {type(e).__name__}: {e}. Կփորձեմ 1 րոպեից")
                time.sleep(60)
                break

            # Draft exists in ՊԵԿ: confirm first so a restart can never create it twice.
            _confirm(uid)

            p = prepared["parsed"]
            summary = (
                "\u2705 Սևագիրը ստեղծված է ՊԵԿ-ում\n"
                f"Ամսաթիվ՝ {p['date']}\n"
                f"Գնորդ՝ {p['buyer']}\n"
                f"Գումար՝ {_fmt_amount(p['total'])} դր."
            )
            try:
                pdf = pek_draft_pdf(usable_token() or STATE["token"], doc_id)
                log.step(f"\u2714 PDF-ը ստացվեց ՊԵԿ-ից ({max(1, len(pdf) // 1024)} ԿԲ)")
                if not log.finish_pdf(f"sevagir-{p['date']}.pdf", pdf, summary):
                    # Editing failed (very rare): fall back to a separate reply.
                    tg_send_document(chat_id, f"sevagir-{p['date']}.pdf", pdf, caption=summary,
                                     reply_to=msg.get("message_id"))
            except Exception as e:
                log.step(f"\u26a0 PDF-ը չստացվեց՝ {e}. Սևագիրը կա ՊԵԿ-ում, PDF-ը վերցրու կայքից")
                log.finish_without_pdf()


# ---------------------------------------------------------------------------
# 12. HTTP routes (the iPhone Shortcut posts here) + keep-alive
# ---------------------------------------------------------------------------
@app.route("/token", methods=["POST"])
def receive_token():
    """Body = raw ՊԵԿ login response. Replies in plain ASCII so the iPhone
    notification is always readable."""
    if not TOKEN_SECRET or request.headers.get("X-Secret") != TOKEN_SECRET:
        return "forbidden", 403
    raw = request.get_data(as_text=True) or ""
    m = re.search(r"<(?:\w+:)?AuthToken>\s*(.*?)\s*</(?:\w+:)?AuthToken>", raw, re.S)
    if m:
        token = m.group(1)
        exp = _jwt_exp(token) or (time.time() + 3600)
        with STATE_LOCK:
            STATE.update(token=token, received=time.time(), exp=exp, blocked_until=0.0, requested_at=0.0)
            STATE["login_events"].append("\u2714 ՊԵԿ login՝ հաջող")
        TOKEN_EVENT.set()
        return f"OK token valid until {_fmt_time(exp)}"
    code = re.search(r'Code="([^"]*)"', raw)
    code = code.group(1) if code else "?"
    text = re.search(r'Message="([^"]*)"', raw)
    text = text.group(1) if text else raw.strip()[:200]
    with STATE_LOCK:
        if code == "0032":
            # A token we no longer know (e.g. after a restart) is still valid.
            STATE["blocked_until"] = time.time() + BLOCK_AFTER_0032
        STATE["login_events"].append(f"\u2716 ՊԵԿ-ը մերժեց login-ը՝ {code}: {text}")
    TOKEN_EVENT.set()
    if code == "0032":
        return "WAIT previous token still valid (0032)", 409
    return f"LOGIN FAILED code={code}", 400


@app.route("/", methods=["GET"])
def health():
    return "ok"


def _keepalive():
    # Render free sleeps after ~15 min without inbound traffic; the worker
    # must keep running, so ping our own public URL every 10 min.
    url = os.environ.get("RENDER_EXTERNAL_URL")
    while True:
        time.sleep(600)
        if url:
            try:
                requests.get(url + "/", timeout=20)
            except Exception:
                pass


# The background threads must live in the SAME process that serves HTTP,
# otherwise /token updates a STATE the worker never sees (gunicorn may import
# the app in its master process and then fork the serving worker). So they are
# started lazily, per process id, on the first request that process handles
# (Render's health check hits "/" within seconds of boot).
_started_pid = None
_start_lock = threading.Lock()


def start_background():
    global _started_pid
    with _start_lock:
        if _started_pid == os.getpid():
            return
        _started_pid = os.getpid()
    threading.Thread(target=worker_loop, daemon=True, name="worker").start()
    threading.Thread(target=_keepalive, daemon=True, name="keepalive").start()


@app.before_request
def _ensure_background():
    if os.environ.get("TITAN_NO_WORKER") != "1" and _started_pid != os.getpid():
        start_background()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
