# ============================================================
# Signboard Report — Zenput 600319 (Classic + Lubda) + 671663 (Garatis)
# GitHub Actions. Today's submissions (KSA): signboard photo +
# "Is the signboard working properly?" per branch, plus branches that
# did not submit -> PDF (HTML cards -> Playwright screenshot -> fpdf2),
# emailed as attachment with an Arabic summary.
# Secrets: ZENPUT_API_TOKEN, GMAIL_APP_PASSWORD
# Env: MODE = auto (send) | test (no email), REPORT_DATE = "" or YYYY-MM-DD
# ============================================================
import os, sys, smtplib, asyncio, requests, re, io, base64, html
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import formataddr
from email import encoders
import pytz
from PIL import Image as PILImage, ImageOps
from fpdf import FPDF
from playwright.async_api import async_playwright
PILImage.MAX_IMAGE_PIXELS = None

API_TOKEN   = os.environ.get("ZENPUT_API_TOKEN", "")
GMAIL_USER  = "aof.group.auto@gmail.com"
GMAIL_PASS  = os.environ.get("GMAIL_APP_PASSWORD", "")
SENDER_NAME = "Business Intelligence"
TO_EMAILS   = [  "m.alsaghir@aofgroup.com","n.joshe@aofgroup.com","a.banafe@aofgroup.com",
    "i.mostafa@aofgroup.com","m.alghazali@aofgroup.com","n.alzahrani@aofgroup.com","s.poudel@aofgroup.com","m.emad@aofgroup.com",
    "a.suliman@aofgroup.com","a.alarabi@aofgroup.com","o.aljawhari@aofgroup.com","s.mansuri@aofgroup.com","m.suhail@aofgroup.com",
    "a.alghanimi@lubdasa.com","m.alfadhil@aofgroup.com"]
CC_EMAILS   = [ "o.salahaddin@aofgroup.com","a.alsalem@aofgroup.com","m.hejazi@aofgroup.com",
    "omar@aofgroup.com","m.alhuaydar@aofgroup.com",
    "a.omara@aofgroup.com","s.alharbi@aofgroup.com",]
MODE        = os.environ.get("MODE", "auto").strip().lower()
REPORT_DATE = os.environ.get("REPORT_DATE", "").strip()   # "" = today (KSA)
BASE        = "https://www.zenput.com"
HEADERS     = {"X-API-TOKEN": API_TOKEN}
TZ          = pytz.timezone("Asia/Riyadh")
PAGE        = 100
LINK_BASE   = "https://ops-execution.crunchtime.com/reports/#form_id/"
PHOTO_W     = 600         # px width photos are resized to before embedding
PAGE_W      = 1240        # px width of each rendered page
SCALE       = 2           # screenshot device scale factor (sharpness)

# template -> brand
TEMPLATES = {600319: "Shawarma Classic", 671663: "Garatis"}

# Expected branches per brand: code -> (Zenput location ID or None, name). B03 deprecated.
BRANCHES = {
    "Shawarma Classic": {
        "B01": (2155652, "NURUH"),  "B02": (2164013, "KHRUH"),  "B04": (2197297, "NSRUH"),
        "B05": (2164016, "RWRUH"),  "B06": (2164017, "DARUH"),  "B07": (2197299, "LBRUH"),
        "B08": (2164019, "SWRUH"),  "B09": (2164020, "AZRUH"),  "B10": (2164021, "SHRUH"),
        "B11": (2164022, "NRRUH"),  "B12": (2164023, "TWRUH"),  "B13": (2155654, "AQRUH"),
        "B14": (2164025, "RBRUH"),  "B15": (2164026, "NDRUH"),  "B16": (2164027, "BDRUH"),
        "B17": (2164028, "QRRUH"),  "B18": (2197298, "TKRUH"),  "B19": (2164030, "MURUH"),
        "B21": (2164032, "KRRUH"),  "B22": (2185452, "OBJED"),  "B23": (2190657, "SLAHS"),
        "B24": (2164031, "SFJED"),  "B25": (2199002, "RWAHS"),  "B26": (2199835, "HAJED"),
        "B27": (2203271, "SARUH"),  "B28": (2210205, "MAJED"),  "B30": (2211854, "QARUH"),
        "B31": (2235670, "ANRUH"),  "B32": (2239240, "FYJED"),  "B33": (2242934, "HIRJED"),
        "B34": (2243963, "URRUH"),  "B35": (2250799, "IRRUH"),  "B36": (2258220, "PSJED"),
        "B37": (2257790, "SHWMAK"), "B38": (2260889, "UHDMM"),  "B39": (2263062, "HSRUH"),
        "B40": (2281339, "MZDMM"),
        "LB01": (2169459, "Lubda Alaqeq"), "LB02": (2222802, "Lubda Alkhaleej"),
    },
    "Garatis": {
        "QB01": (2232755, "السويدي"), "QB02": (2235805, "النرجس"), "QB03": (2254072, "اليرموك"),
        "QB04": (2256386, "العقيق"),  "QB05": (2268360, "الخافجي"), "QB06": (2270650, "الخليج"),
        "QB07": (2274188, "الندوة"),  "QB08": (2276794, "المنصورة"), "QB09": (None, "النزهة"),
    },
}

# ---------------- helpers ----------------
def parse_dt(s):
    if not s: return None
    try: return datetime.fromisoformat(s.replace("Z", "+00:00")).astimezone(TZ)
    except Exception: return None

def resolve_branch(brand, loc):
    exp = BRANCHES[brand]
    try: lid = int(loc.get("id"))
    except Exception: lid = None
    for code, (i, _) in exp.items():
        if i is not None and i == lid: return code
    name = re.sub(r"\s+", " ", (loc.get("name") or loc.get("display_name") or "")).strip()
    if brand == "Garatis":
        m = re.search(r"\bQB\s?(\d{2})\b", name, re.I) or re.search(r"Garatis+\s*B(\d{2})", name, re.I)
        code = f"QB{m.group(1)}" if m else None
        return code if code in exp else None
    m = re.match(r"([A-Z]+)\s+(B\d{2})\b", name.upper())
    if m and m.group(2) in exp and exp[m.group(2)][1] == m.group(1):
        return m.group(2)
    for code, (_, n) in exp.items():
        if code.startswith("LB") and n.lower() in name.lower(): return code
    return None

def is_photo_q(t):   return "signboard" in t and "open" in t and "drive" not in t
def is_working_q(t): return "signboard" in t and "working" in t

def photo_keys(v):
    items = v if isinstance(v, list) else ([v] if isinstance(v, dict) else [])
    return [i["s3_key"] for i in items if isinstance(i, dict) and i.get("s3_key")]

def signed_url(s3_key):
    r = requests.get(f"{BASE}/api/v2/users/current/storage/", headers=HEADERS,
                     params={"path": s3_key}, timeout=60)
    if r.status_code != 200: return None
    return ((r.json() or {}).get("data") or {}).get("location")

def fetch_photo(s3_key):
    """-> (data URI or None, error)"""
    url = signed_url(s3_key)
    if not url: return None, "no signed url"
    r = requests.get(url, timeout=60)
    if r.status_code != 200: return None, f"HTTP {r.status_code}"
    try:
        img = PILImage.open(io.BytesIO(r.content))
        try: img = ImageOps.exif_transpose(img)
        except Exception: pass
        img = img.convert("RGB")
        if img.width > PHOTO_W:
            img = img.resize((PHOTO_W, int(img.height * PHOTO_W / img.width)))
        buf = io.BytesIO(); img.save(buf, "JPEG", quality=82)
        return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode(), None
    except Exception as e:
        return None, str(e)

def fetch_day(template_id, day):
    out, start = [], 0
    while True:
        r = requests.get(f"{BASE}/api/v3/submissions/", headers=HEADERS, params={
            "form_template_id": template_id, "limit": PAGE, "start": start}, timeout=60)
        print(f"[{template_id}] GET start={start} -> HTTP {r.status_code}")
        if r.status_code != 200:
            print(r.text[:500]); break
        data = r.json().get("data", [])
        if not data: break
        out.extend(data)
        dts = [d for d in (parse_dt((s.get("smetadata") or {}).get("date_submitted")) for s in data) if d]
        if len(data) < PAGE or (dts and min(dts).date() < day): break
        start += PAGE
    return out

if not API_TOKEN:
    sys.exit("ZENPUT_API_TOKEN is not set")

# ---------------- fetch + latest per branch ----------------
now = datetime.now(TZ)
day = datetime.strptime(REPORT_DATE, "%Y-%m-%d").date() if REPORT_DATE else now.date()
print(f"Report date: {day} (KSA)")

rows = []
for tid, brand in TEMPLATES.items():
    subs = fetch_day(tid, day)
    latest, unknown = {}, set()
    for s in subs:
        sm = s.get("smetadata") or {}
        dt = parse_dt(sm.get("date_submitted"))
        if not dt or dt.date() != day: continue
        loc = sm.get("location") or {}
        code = resolve_branch(brand, loc)
        if code is None:
            unknown.add(f"{loc.get('id')} {loc.get('name')}"); continue
        if code not in latest or dt > latest[code][0]:
            latest[code] = (dt, s)
    print(f"[{tid}] {brand}: fetched {len(subs)} | submitted on {day}: {len(latest)} branches")
    if unknown: print(f"!! [{tid}] unmapped locations (ignored): {unknown}")

    warned = False
    for code, (_, name) in sorted(BRANCHES[brand].items(), key=lambda x: (x[0].startswith("LB"), int(re.sub(r"\D", "", x[0])))):
        if code not in latest:
            rows.append({"brand": brand, "code": code, "name": name, "submitted": False}); continue
        dt, s = latest[code]
        working, keys, prepared, found = None, [], "", set()
        for a in s.get("answers", []):
            t = (a.get("title") or "").strip().lower()
            if a.get("field_type") == "image" and is_photo_q(t):
                keys = photo_keys(a.get("value")); found.add("photo")
            elif is_working_q(t):
                working = str(a.get("value")).lower() in ("true", "yes", "1"); found.add("working")
            elif t == "prepared by":
                prepared = str(a.get("value") or "").strip()
        if len(found) < 2 and not warned:
            warned = True
            print(f"\n!! [{tid}] signboard questions not found in {code} — answer titles:")
            for a in s.get("answers", []):
                print(f"   {a.get('title')} / {a.get('field_type')}")
        rows.append({"brand": brand, "code": code, "name": name, "submitted": True,
                     "time": dt.strftime("%H:%M"), "prepared": prepared, "working": working,
                     "keys": keys, "link": LINK_BASE + str(s.get("legacy_submission_id") or "")})

# ---------------- photos ----------------
photo_fail = 0
for r in rows:
    r["img"] = None
    if r.get("submitted") and r.get("keys"):
        r["img"], err = fetch_photo(r["keys"][0])
        if err:
            photo_fail += 1; print(f"  photo failed {r['code']}: {err}")

# ---------------- HTML ----------------
BRAND_AR = {"Shawarma Classic": "شاورما كلاسيك", "Garatis": "قراطيس"}
CSS = f"""
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700;800&display=swap');
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ width: {PAGE_W}px; font-family: 'Tajawal', Tahoma, Arial, sans-serif; background: #f3f5f9; color: #1f2937; direction: rtl; }}
.wrap {{ padding: 28px 30px 34px; }}
.head {{ background: #1F3864; color: #fff; border-radius: 14px; padding: 20px 26px; display: flex; justify-content: space-between; align-items: center; }}
.head h1 {{ font-size: 30px; font-weight: 800; }}
.head .sub {{ font-size: 17px; opacity: .85; margin-top: 4px; }}
.head .date {{ font-size: 20px; font-weight: 700; direction: ltr; }}
.stats {{ display: flex; gap: 14px; margin: 18px 0 22px; }}
.stat {{ flex: 1; background: #fff; border-radius: 12px; padding: 14px 16px; text-align: center; border: 1px solid #e3e7ef; }}
.stat .n {{ font-size: 30px; font-weight: 800; }}
.stat .l {{ font-size: 15px; color: #6b7280; margin-top: 2px; }}
.ok {{ color: #15803d; }} .bad {{ color: #c00000; }} .muted {{ color: #6b7280; }}
.grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; }}
.card {{ background: #fff; border-radius: 12px; overflow: hidden; border: 1px solid #e3e7ef; display: flex; flex-direction: column; }}
.card.miss {{ background: #eceef2; border: 2px dashed #c4c9d4; }}
.card.notok {{ border: 2px solid #c00000; }}
.ch {{ padding: 10px 12px 8px; display: flex; justify-content: space-between; align-items: center; gap: 6px; }}
.bn {{ font-size: 16px; font-weight: 800; direction: ltr; text-align: left; }}
.badge {{ font-size: 13px; font-weight: 700; padding: 3px 10px; border-radius: 20px; white-space: nowrap; }}
.b-ok {{ background: #dcfce7; color: #15803d; }} .b-bad {{ background: #fde2e2; color: #c00000; }}
.b-miss {{ background: #d1d5db; color: #374151; }}
.ph {{ height: 210px; background: #e5e7eb; display: flex; align-items: center; justify-content: center; }}
.ph img {{ width: 100%; height: 100%; object-fit: cover; }}
.ph .none {{ color: #9ca3af; font-size: 15px; }}
.cf {{ padding: 8px 12px 10px; font-size: 13.5px; color: #4b5563; display: flex; justify-content: space-between; }}
.cf span {{ direction: ltr; }}
"""

def card(r):
    label = html.escape(f"{r['code']} {r['name']}")
    if not r["submitted"]:
        return (f"<div class='card miss'><div class='ch'><div class='bn'>{label}</div>"
                f"<div class='badge b-miss'>لم يسلّم</div></div>"
                f"<div class='ph'><div class='none'>لا يوجد تسليم اليوم</div></div>"
                f"<div class='cf'>&nbsp;</div></div>")
    ok = r["working"]
    badge = ("<div class='badge b-ok'>اللوحة تعمل</div>" if ok else
             "<div class='badge b-bad'>اللوحة لا تعمل</div>" if ok is False else
             "<div class='badge b-miss'>غير محدد</div>")
    ph = f"<img src='{r['img']}'>" if r["img"] else "<div class='none'>لا توجد صورة</div>"
    return (f"<div class='card{' notok' if ok is False else ''}'><div class='ch'><div class='bn'>{label}</div>{badge}</div>"
            f"<div class='ph'>{ph}</div>"
            f"<div class='cf'><span>{html.escape(r['prepared'] or '-')}</span><span>{r['time']}</span></div></div>")

def page_html(brand):
    br = [r for r in rows if r["brand"] == brand]
    sub = sum(1 for r in br if r["submitted"])
    ok = sum(1 for r in br if r["submitted"] and r["working"])
    bad = sum(1 for r in br if r["submitted"] and r["working"] is False)
    miss = len(br) - sub
    # order: not working first, then not submitted, then working
    order = sorted(br, key=lambda r: (0 if r["submitted"] and r["working"] is False else 1 if not r["submitted"] else 2))
    return f"""<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><style>{CSS}</style></head>
<body><div class="wrap">
<div class="head"><div><h1>تقرير لوحات الفروع - {BRAND_AR.get(brand, brand)}</h1>
<div class="sub">صورة اللوحة أثناء التشغيل وحالة عملها لكل فرع</div></div><div class="date">{day}</div></div>
<div class="stats">
<div class="stat"><div class="n">{sub} / {len(br)}</div><div class="l">الفروع المسلّمة</div></div>
<div class="stat"><div class="n ok">{ok}</div><div class="l">اللوحات تعمل</div></div>
<div class="stat"><div class="n bad">{bad}</div><div class="l">اللوحات لا تعمل</div></div>
<div class="stat"><div class="n muted">{miss}</div><div class="l">فروع لم تسلّم</div></div>
</div>
<div class="grid">{''.join(card(r) for r in order)}</div>
</div></body></html>"""

# ---------------- render -> PDF ----------------
async def render_pngs():
    pngs = []
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={"width": PAGE_W, "height": 1000}, device_scale_factor=SCALE)
        for brand in TEMPLATES.values():
            await page.set_content(page_html(brand), wait_until="networkidle")
            await page.wait_for_timeout(1500)
            safe = re.sub(r"\W+", "_", brand)
            path = f"_signboard_{safe}.png"
            await page.screenshot(path=path, full_page=True)
            pngs.append(path)
        await browser.close()
    return pngs

def pngs_to_pdf(pngs, out):
    pdf = FPDF(unit="pt")
    pdf.set_auto_page_break(False)
    for pth in pngs:
        w, h = PILImage.open(pth).size
        pw, ph = w / SCALE * 0.75, h / SCALE * 0.75     # px -> pt at 96 dpi
        pdf.add_page(format=(pw, ph))
        pdf.image(pth, x=0, y=0, w=pw, h=ph)
    pdf.output(out)

fname = f"signboard_report_{day}.pdf"
pngs = asyncio.run(render_pngs())
pngs_to_pdf(pngs, fname)
for pth in pngs:
    try: os.remove(pth)
    except Exception: pass

print(f"\nSaved {fname} | photo failures: {photo_fail}")
for brand in TEMPLATES.values():
    br = [r for r in rows if r["brand"] == brand]
    print(f"\n{brand}: submitted {sum(1 for r in br if r['submitted'])}/{len(br)}")
    print("  Not working:", [f"{r['code']} {r['name']}" for r in br if r["submitted"] and r["working"] is False] or "-")
    print("  Not submitted:", [f"{r['code']} {r['name']}" for r in br if not r["submitted"]] or "-")


# ---------------- email ----------------
BRAND_AR = {"Shawarma Classic": "شاورما كلاسيك", "Garatis": "قراطيس"}
TD = "padding:6px 12px;border:1px solid #cccccc;text-align:center"
TH = "padding:8px 12px;background:#1F3864;color:#ffffff;text-align:center"

def list_table(title, items):
    if not items:
        return f"<p style='margin:6px 0'><b>{title}:</b> لا يوجد</p>"
    body = "".join(f"<tr><td style='{TD}'>{i}</td><td style='{TD};color:#C00000' dir='auto'>{b}</td></tr>"
                   for i, b in enumerate(items, 1))
    return (f"<p style='margin:10px 0 4px'><b>{title} ({len(items)})</b></p>"
            f"<table style='border-collapse:collapse'><tr><th style='{TH}'>#</th><th style='{TH}'>الفرع</th></tr>{body}</table>")

def build_html():
    summary, details = "", ""
    for brand in TEMPLATES.values():
        br = [r for r in rows if r["brand"] == brand]
        sub = sum(1 for r in br if r["submitted"])
        ok = [r for r in br if r["submitted"] and r["working"]]
        bad = [f"{r['code']} {r['name']}" for r in br if r["submitted"] and r["working"] is False]
        miss = [f"{r['code']} {r['name']}" for r in br if not r["submitted"]]
        ar = BRAND_AR.get(brand, brand)
        summary += (f"<tr><td style='{TD}'><b>{ar}</b></td><td style='{TD}'>{sub} / {len(br)}</td>"
                    f"<td style='{TD};color:#006100'><b>{len(ok)}</b></td>"
                    f"<td style='{TD};color:#C00000'><b>{len(bad)}</b></td>"
                    f"<td style='{TD};color:#C00000'><b>{len(miss)}</b></td></tr>")
        details += (f"<h3 style='color:#1F3864;margin:20px 0 6px'>{ar}</h3>"
                    + list_table("فروع اللوحة لا تعمل", bad)
                    + list_table("فروع لم تسلّم النموذج", miss))
    return f"""<html><body dir="rtl" style="font-family:Tahoma,Arial,sans-serif;font-size:14px;color:#222;text-align:right">
<p>السلام عليكم ورحمة الله وبركاته،</p>
<p>مرفق لكم تقرير لوحات الفروع ليوم <span dir="ltr">{day}</span>، ويشمل صورة اللوحة وحالة تشغيلها لكل فرع.</p>
<table style="border-collapse:collapse">
<tr><th style="{TH}">العلامة التجارية</th><th style="{TH}">الفروع المسلّمة</th><th style="{TH}">اللوحات تعمل</th>
<th style="{TH}">اللوحات لا تعمل</th><th style="{TH}">فروع لم تسلّم</th></tr>
{summary}
</table>
{details}
<p style="margin-top:18px">صور اللوحات لجميع الفروع في ملف PDF المرفق.</p>
<p>مع خالص التحية،<br>إدارة ذكاء الأعمال</p>
</body></html>"""

def send_email(path):
    if not GMAIL_PASS:
        sys.exit("GMAIL_APP_PASSWORD is not set")
    msg = MIMEMultipart()
    msg["From"] = formataddr((SENDER_NAME, GMAIL_USER))
    msg["To"] = ", ".join(TO_EMAILS)
    msg["Cc"] = ", ".join(CC_EMAILS)
    msg["Subject"] = f"تقرير لوحات الفروع - {day}"
    msg.attach(MIMEText(build_html(), "html", "utf-8"))
    with open(path, "rb") as f:
        part = MIMEBase("application", "pdf")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{os.path.basename(path)}"')
    msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(GMAIL_USER, GMAIL_PASS)
        s.sendmail(GMAIL_USER, TO_EMAILS + CC_EMAILS, msg.as_string())
    print(f"Email sent to {TO_EMAILS} cc {CC_EMAILS}")

if MODE == "test":
    print("MODE=test -> email skipped")
else:
    send_email(fname)
