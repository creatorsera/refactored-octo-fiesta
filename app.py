"""
app.py — CSV Email Validator (standalone)
Upload CSV -> pick Best Email + All Emails columns -> validate with early-stop
on first Deliverable -> styled 3-sheet XLSX export with full audit trail.
"""

import streamlit as st
import re, io, smtplib, time
import pandas as pd
from datetime import datetime
from email_validator import validate_email as ev_validate, EmailNotValidError
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import requests

try:
    import dns.resolver as _dns_resolver
    DNS_AVAILABLE = True
except ImportError:
    DNS_AVAILABLE = False

# ── CONSTANTS ─────────────────────────────────────────────────────────────────
EMAIL_REGEX = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", re.IGNORECASE)
TIER1 = re.compile(r"^(editor|admin|press|advert|contact)[a-z0-9._%+\-]*@", re.IGNORECASE)
TIER2 = re.compile(r"^(info|sales|hello|office|team|support|help)[a-z0-9._%+\-]*@", re.IGNORECASE)

BLOCKED_TLDS = {
    'png','jpg','jpeg','webp','gif','svg','ico','bmp','tiff','avif','mp4','mp3',
    'wav','ogg','mov','avi','webm','pdf','zip','rar','tar','gz','7z','js','css',
    'php','asp','aspx','xml','json','ts','jsx','tsx','woff','woff2','ttf','eot',
    'otf','map','exe','dmg','pkg','deb','apk',
}
PLACEHOLDER_DOMAINS = {
    'example.com','example.org','example.net','test.com','domain.com',
    'yoursite.com','yourwebsite.com','website.com','email.com','placeholder.com',
}
PLACEHOLDER_LOCALS = {
    'you','user','name','email','test','example','someone','username',
    'yourname','youremail','enter','address','sample',
}
SUPPRESS_PREFIXES = [
    'noreply','no-reply','donotreply','do-not-reply','mailer-daemon','bounce',
    'bounces','unsubscribe','notifications','notification','newsletter',
    'newsletters','postmaster','webmaster','auto-reply','autoreply','daemon',
    'robot','alerts','alert','system',
}
FREE_EMAIL_DOMAINS = {
    "gmail.com","yahoo.com","hotmail.com","outlook.com","aol.com",
    "icloud.com","protonmail.com","zoho.com","live.com","msn.com",
}
_DISPOSABLE_FALLBACK = {
    'mailinator.com','guerrillamail.com','tempmail.com','throwaway.email','yopmail.com',
    'sharklasers.com','spam4.me','trashmail.com','trashmail.me','maildrop.cc',
    '10minutemail.com','fakeinbox.com','discard.email','mailnesia.com',
    'tempr.email','trashmail.at','trashmail.io','wegwerfemail.de','meltmail.com',
}

@st.cache_data(ttl=86400, show_spinner=False)
def fetch_disposable_domains():
    try:
        r = requests.get(
            "https://raw.githubusercontent.com/disposable-email-domains/"
            "disposable-email-domains/main/disposable_email_blocklist.conf",
            timeout=8)
        if r.status_code == 200:
            return set(r.text.splitlines())
    except Exception:
        pass
    return _DISPOSABLE_FALLBACK

# ── EMAIL HELPERS ─────────────────────────────────────────────────────────────
def is_valid_email(email):
    e = email.strip()
    if not e or e.count('@') != 1: return False
    local, domain = e.split('@'); lo, do = local.lower(), domain.lower()
    if not local or not domain: return False
    if local.startswith('.') or local.endswith('.') or local.startswith('-'): return False
    if len(local) > 64 or len(domain) > 255: return False
    if '.' not in domain: return False
    tld = do.rsplit('.', 1)[-1]
    if len(tld) < 2 or tld in BLOCKED_TLDS: return False
    if re.search(r'@\d+x[\-\d]', '@'+do): return False
    if re.match(r'^\d+x', do): return False
    if do in PLACEHOLDER_DOMAINS: return False
    if lo in PLACEHOLDER_LOCALS: return False
    if any(lo == p or lo.startswith(p) for p in SUPPRESS_PREFIXES): return False
    if re.search(r'\d+x\d+', lo): return False
    return True

def tier_key(e):
    if TIER1.match(e): return "1"
    if TIER2.match(e): return "2"
    return "3"

def tier_short(e): return {"1":"Tier 1","2":"Tier 2","3":"Tier 3"}[tier_key(e)]
def sort_by_tier(emails): return sorted(emails, key=tier_key)

def confidence_score(email, val):
    if not val: return None
    s = 100; t = tier_key(email)
    if t == "2": s -= 10
    if t == "3": s -= 25
    if not val.get("spf"):       s -= 15
    if val.get("catch_all"):     s -= 20
    if val.get("free"):          s -= 8
    st_ = val.get("status", "")
    if st_ == "Risky":           s -= 30
    if st_ == "Not Deliverable": s -= 65
    return max(0, s)

def conf_color(sc):
    if sc is None: return "#ccc"
    if sc >= 75:   return "#16a34a"
    if sc >= 45:   return "#d97706"
    return "#dc2626"

def parse_email_cell(cell_value):
    if pd.isna(cell_value) or not str(cell_value).strip():
        return []
    text = str(cell_value).strip()
    for delim in [';', ',', '|', '\n']:
        if delim in text:
            parts = text.split(delim)
            break
    else:
        parts = [text]
    emails = []
    for part in parts:
        part = part.strip().strip('"').strip("'")
        found = EMAIL_REGEX.findall(part)
        if found:
            emails.extend(found)
        elif is_valid_email(part):
            emails.append(part)
    seen = set(); result = []
    for e in emails:
        el = e.lower()
        if el not in seen:
            seen.add(el); result.append(e)
    return result

# ── VALIDATION ENGINE ─────────────────────────────────────────────────────────
def _val_syntax(email):
    try: ev_validate(email); return True
    except EmailNotValidError: return False

def _val_mx(domain):
    try:
        recs = _dns_resolver.resolve(domain, "MX")
        return True, [str(r.exchange) for r in recs]
    except: return False, []

def _val_spf(domain):
    try:
        for rd in _dns_resolver.resolve(domain, "TXT"):
            if "v=spf1" in str(rd): return True
    except: pass
    return False

def _val_dmarc(domain):
    try:
        for rd in _dns_resolver.resolve(f"_dmarc.{domain}", "TXT"):
            if "v=DMARC1" in str(rd): return True
    except: pass
    return False

def _val_mailbox(email, mx_records):
    try:
        mx = mx_records[0].rstrip(".")
        with smtplib.SMTP(mx, timeout=6) as s:
            s.helo("example.com"); s.mail("test@example.com")
            code, _ = s.rcpt(email)
            return code == 250
    except: return False

def _val_catch_all(domain, mx_records):
    try:
        mx = mx_records[0].rstrip(".")
        with smtplib.SMTP(mx, timeout=6) as s:
            s.helo("example.com"); s.mail("test@example.com")
            code, _ = s.rcpt(f"randomaddress9x7z@{domain}")
            return code == 250
    except: return False

def _deliverability(syntax, mx_ok, mailbox_ok, disposable, free, catch_all, spf_ok):
    if not syntax:    return "Not Deliverable", "Invalid syntax"
    if disposable:    return "Not Deliverable", "Disposable domain"
    if not mx_ok:     return "Not Deliverable", "No MX records"
    if mailbox_ok:
        if free: return ("Risky", "Catch-all + free") if catch_all else ("Deliverable", "Free provider")
        if catch_all:  return "Risky", "Catch-all enabled"
        if not spf_ok: return "Risky", "Missing SPF"
        return "Deliverable", "—"
    else:
        if catch_all:  return "Risky", "Catch-all, mailbox unknown"
        if free:       return "Deliverable", "Free provider (unverified)"
        if not spf_ok: return "Risky", "No SPF — spam risk"
        return "Deliverable", "MX/SPF OK, mailbox unconfirmed"

def validate_email_full(email):
    disp = fetch_disposable_domains()
    domain = email.split("@")[-1].lower()
    syntax = _val_syntax(email)
    mx_ok, mx_h = _val_mx(domain) if DNS_AVAILABLE else (False, [])
    spf   = _val_spf(domain)   if DNS_AVAILABLE else False
    dmarc = _val_dmarc(domain) if DNS_AVAILABLE else False
    disp_ = domain in disp
    free  = domain in FREE_EMAIL_DOMAINS
    mbox  = _val_mailbox(email, mx_h) if (mx_ok and DNS_AVAILABLE) else False
    ca    = _val_catch_all(domain, mx_h) if (mx_ok and DNS_AVAILABLE) else False
    status, reason = _deliverability(syntax, mx_ok, mbox, disp_, free, ca, spf)
    return {"status": status, "reason": reason, "syntax": syntax, "mx": mx_ok,
            "spf": spf, "dmarc": dmarc, "mailbox": mbox, "disposable": disp_,
            "free": free, "catch_all": ca}

def validate_with_early_stop(best_email, all_emails):
    log = []
    if not best_email or not is_valid_email(best_email):
        log.append((best_email or "(empty)", "skipped", "Invalid format"))
        for email in sort_by_tier(all_emails):
            if email == best_email or not is_valid_email(email): continue
            val = validate_email_full(email)
            log.append((email, val["status"], val["reason"]))
            if val["status"] == "Deliverable":
                return email, val, True, log
        return best_email or (all_emails[0] if all_emails else ""), None, False, log

    val = validate_email_full(best_email)
    log.append((best_email, val["status"], val["reason"]))

    if val["status"] == "Deliverable":
        return best_email, val, False, log

    best_risky_val = None; best_risky_email = None
    for email in sort_by_tier(all_emails):
        if email == best_email or not is_valid_email(email): continue
        v = validate_email_full(email)
        log.append((email, v["status"], v["reason"]))
        if v["status"] == "Deliverable":
            return email, v, True, log
        if v["status"] == "Risky" and best_risky_val is None:
            best_risky_val = v; best_risky_email = email

    if best_risky_val:
        return best_risky_email, best_risky_val, True, log
    return best_email, val, False, log

# ══════════════════════════════════════════════════════════════════════════════
#  XLSX BUILDER
# ══════════════════════════════════════════════════════════════════════════════
def _mf(h):  return PatternFill("solid", fgColor=h)
def _fn(b=False, c="111111", s=10, n="Calibri", i=False):
    return Font(bold=b, color=c, size=s, name=n, italic=i)
def _bd():
    t = Side(style="thin", color="E5E7EB")
    return Border(left=t, right=t, top=t, bottom=t)
def _ct(): return Alignment(horizontal="center", vertical="center")
def _lt(): return Alignment(horizontal="left", vertical="center", wrap_text=False)

RF_D = _mf("F0FDF4"); RF_R = _mf("FFFBEB"); RF_B = _mf("FFF1F2"); RF_N = _mf("F9FAFB")
EF_D = _mf("DCFCE7"); EF_R = _mf("FEF3C7"); EF_B = _mf("FECACA"); EF_F = _mf("E0F2FE")
TF1 = _mf("FEF9C3"); TF2 = _mf("EEF2FF"); TF3 = _mf("F1F5F9")
CF_H = _mf("D1FAE5"); CF_M = _mf("FEF3C7"); CF_L = _mf("FEE2E2")
SF = {"Deliverable": _mf("16A34A"), "Risky": _mf("D97706"), "Not Deliverable": _mf("DC2626")}
HDR = _mf("111111")

def _rf(s): return {"Deliverable": RF_D, "Risky": RF_R, "Not Deliverable": RF_B}.get(s, RF_N)
def _ef(s, fb): return EF_F if fb else {"Deliverable": EF_D, "Risky": EF_R, "Not Deliverable": EF_B}.get(s)
def _tf(t): return TF1 if "1" in t else (TF2 if "2" in t else TF3)
def _cf(sc): return None if sc is None else (CF_H if sc >= 75 else (CF_M if sc >= 45 else CF_L))

def _hdr(ws, r, c, v, w=None):
    cl = ws.cell(row=r, column=c, value=v)
    cl.fill = HDR; cl.font = _fn(b=True, c="FFFFFF"); cl.alignment = _ct(); cl.border = _bd()
    if w: ws.column_dimensions[get_column_letter(c)].width = w
    return cl

def _cl(ws, r, c, v, fl=None, fn_=None, al=None):
    cl = ws.cell(row=r, column=c, value=v)
    if fl: cl.fill = fl
    if fn_: cl.font = fn_
    if al: cl.alignment = al
    cl.border = _bd()
    return cl

def _stats_sheet(wb, name, rows, title, sub=""):
    ws = wb.create_sheet(name)
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 10; ws.column_dimensions["C"].width = 32
    t = ws.cell(row=1, column=1, value=title); t.font = _fn(b=True, s=15); t.fill = _mf("F9FAFB")
    ws.merge_cells("A1:C1"); ws.row_dimensions[1].height = 28; t.alignment = _lt()
    if sub:
        s = ws.cell(row=2, column=1, value=sub); s.font = _fn(c="999999", s=9, i=True); ws.merge_cells("A2:C2")
    ts = ws.cell(row=3, column=1, value=f"Generated: {datetime.now().strftime('%d %b %Y  %H:%M')}")
    ts.font = _fn(c="AAAAAA", s=9); ws.merge_cells("A3:C3")
    FG = {"total":"0C4A6E","deliverable":"14532D","risky":"78350F","fail":"881337",
          "fallback":"0C4A6E","none":"374151","avg":"14532D","default":"374151"}
    BG = {"total":"F0F9FF","deliverable":"F0FDF4","risky":"FFFBEB","fail":"FFF1F2",
          "fallback":"E0F2FE","none":"F9FAFB","avg":"F0FDF4","default":"F9FAFB"}
    total = max(1, next((v for _, v, _ in rows if "total" in _.lower()), 1))
    for i, (label, value, key) in enumerate(rows, 5):
        fg = FG.get(key, FG["default"]); bg = BG.get(key, BG["default"]); fl = _mf(bg)
        _cl(ws, i, 1, label, fl, _fn(c=fg, s=10), _lt())
        _cl(ws, i, 2, value, fl, _fn(b=True, c=fg, s=11), _ct())
        ws.row_dimensions[i].height = 21
        if isinstance(value, (int, float)) and key not in ("avg",):
            pct = min(float(value) / total, 1.0); n = int(pct * 22)
            _cl(ws, i, 3, "█"*n + "░"*(22-n) + f"  {round(pct*100)}%", fl, _fn(s=9, n="Courier New", c=fg), _lt())
        else:
            _cl(ws, i, 3, "", fl)
    return ws

def build_xlsx(results, original_columns):
    wb = Workbook()

    # ── Sheet 1: Results (Original CSV Data + Validation appended) ────
    ws = wb.active; ws.title = "Results"; ws.freeze_panes = "A2"; ws.row_dimensions[1].height = 26
    
    for ci, col_name in enumerate(original_columns, 1):
        w = min(max(len(str(col_name)) * 2, 15), 40)
        _hdr(ws, 1, ci, col_name, w=w)

    val_cols = [
        ("Validated Email",32), ("Status",16), ("Score",8), ("Tier",9),
        ("Reason",22), ("SPF",6), ("DMARC",7), ("Catch-all",10),
        ("Fallback?",10), ("Emails Checked",14)
    ]
    val_offset = len(original_columns)
    for ci, (n, w) in enumerate(val_cols, val_offset + 1):
        _hdr(ws, 1, ci, n, w=w)

    for ri, row in enumerate(results, 2):
        orig_data = row.get("original_row_data", {})
        
        for ci, col_name in enumerate(original_columns, 1):
            val = orig_data.get(col_name, "")
            try:
                if pd.isna(val): val = ""
            except TypeError:
                pass
            _cl(ws, ri, ci, val, RF_N, _fn(s=9), _lt())

        v = row.get("val") or {}; st_ = v.get("status", ""); fb = row.get("was_fallback")
        em = row.get("validated_email", ""); cf = row.get("confidence")
        rf = _rf(st_); ef = _ef(st_, fb)

        v_idx = val_offset + 1
        _cl(ws, ri, v_idx, em, ef or rf, _fn(b=True, n="Courier New", s=9), _lt())
        sf_ = SF.get(st_)
        w_col = "FFFFFF" if sf_ else "111111"
        _cl(ws, ri, v_idx+1, st_ or "—", sf_ or rf, _fn(b=bool(sf_), c=w_col, s=9), _ct())
        _cl(ws, ri, v_idx+2, cf if cf is not None else "—", _cf(cf) or rf, _fn(b=True, s=9), _ct())
        _cl(ws, ri, v_idx+3, tier_short(em) if em else "—", _tf(tier_short(em)) if em else rf, _fn(s=9), _ct())
        _cl(ws, ri, v_idx+4, v.get("reason","—") if v else "—", rf, _fn(s=9), _lt())
        
        for c_off, key in [(5,"spf"),(6,"dmarc"),(7,"catch_all")]:
            ok = v.get(key) if v else None
            c_val = "16A34A" if ok else "DC2626"
            f_val = _fn(c=c_val, s=11) if ok is not None else _fn(c="AAAAAA", s=11)
            _cl(ws, ri, v_idx+c_off, "Yes" if ok else "No", rf, f_val, _ct())
            
        fb_c = "0891B2" if fb else "AAAAAA"
        _cl(ws, ri, v_idx+8, "Yes" if fb else "No", rf, _fn(b=fb, c=fb_c, s=9), _ct())
        _cl(ws, ri, v_idx+9, len(row.get("val_log",[])), rf, _fn(s=9), _ct())

    # ── Sheet 2: Validation Log ───────────────────────────────────────
    ws2 = wb.create_sheet("Validation Log"); ws2.freeze_panes = "A2"; ws2.row_dimensions[1].height = 26
    for ci, (n, w) in enumerate([("#",6),("Domain",22),("Original Best",32),("Email Checked",32),
                                  ("Status",16),("Reason",22),("Result",14)], 1):
        _hdr(ws2, 1, ci, n, w)
    r2 = 2
    for ri, row in enumerate(results, 1):
        dom = row.get("domain", ""); orig = row.get("original_email", ""); vl = row.get("val_log", [])
        chosen = row.get("validated_email", "")
        for li, (ce, cs, cr) in enumerate(vl):
            is_f = ce == chosen
            rf2 = _rf(cs) if cs in ("Deliverable","Risky","Not Deliverable") else RF_N
            if is_f: rf2 = EF_D if cs == "Deliverable" else (EF_R if cs == "Risky" else EF_B)
            _cl(ws2,r2,1,f"{ri}.{li+1}",rf2,_fn(s=9),_ct())
            _cl(ws2,r2,2,dom if li == 0 else "",rf2,_fn(s=9),_lt())
            _cl(ws2,r2,3,orig if li == 0 else "",rf2,_fn(n="Courier New",s=9,c="888888"),_lt())
            _cl(ws2,r2,4,ce,rf2,_fn(b=is_f,n="Courier New",s=9),_lt())
            sf2 = SF.get(cs)
            s_col = "FFFFFF" if sf2 else "111111"
            _cl(ws2,r2,5,cs or "skipped",sf2 or rf2,_fn(b=bool(sf2),c=s_col,s=9),_ct())
            _cl(ws2,r2,6,cr,rf2,_fn(s=9),_lt())
            f_font = _fn(b=True,c="0891B2",s=9) if is_f else _fn(s=9)
            _cl(ws2,r2,7,"CHOSEN" if is_f else "",EF_F if is_f else rf2,f_font,_ct())
            ws2.row_dimensions[r2].height = 15; r2 += 1

    # ── Sheet 3: Stats ────────────────────────────────────────────────
    nt = len(results)
    nd = sum(1 for r in results if (r.get("val",{}) or {}).get("status")=="Deliverable")
    nri = sum(1 for r in results if (r.get("val",{}) or {}).get("status")=="Risky")
    nb = sum(1 for r in results if (r.get("val",{}) or {}).get("status")=="Not Deliverable")
    nfb = sum(1 for r in results if r.get("was_fallback"))
    n_empty = sum(1 for r in results if not r.get("val"))
    tc = sum(len(r.get("val_log",[])) for r in results)
    n_val_rows = nt - n_empty
    ac = round(tc/n_val_rows, 1) if n_val_rows else 0
    confs = [r.get("confidence") for r in results if r.get("confidence") is not None]
    avgc = round(sum(confs)/len(confs), 1) if confs else "—"
    saved = tc - n_val_rows

    _stats_sheet(wb, "Stats", [
        ("Total rows in CSV", nt, "total"),
        ("Rows with no email (retained)", n_empty, "none"),
        ("Rows validated", n_val_rows, "total"),
        ("Deliverable", nd, "deliverable"),
        ("Risky", nri, "risky"),
        ("Not Deliverable", nb, "fail"),
        ("Fallback emails used", nfb, "fallback"),
        ("Total emails checked", tc, "total"),
        ("Extra checks (fallbacks)", saved, "fallback"),
        ("Avg checks per valid row", ac, "avg"),
        ("Avg confidence score", avgc, "avg"),
    ], "CSV Email Validator Results", f"{nt} total rows · {n_val_rows} validated · {n_empty} empty retained")

    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT APP
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="CSV Email Validator", page_icon="✅",
                   layout="wide", initial_sidebar_state="expanded")

ACCENT = "#16a34a"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');
*,html,body,[class*="css"] {{ font-family:'Inter',system-ui,sans-serif !important; }}
#MainMenu,footer,header {{ visibility:hidden; }}
.block-container {{ padding:1.2rem 2rem 4rem !important; max-width:100% !important; background:#f6f5f2 !important; }}
[data-testid="stSidebar"] {{ background:#111 !important; }}
[data-testid="stSidebar"] * {{ color:#ccc !important; }}
[data-testid="stSidebar"] .stDownloadButton > button {{
    background:{ACCENT} !important; border:none !important; color:#fff !important;
    border-radius:8px !important; font-size:12px !important; font-weight:700 !important; width:100% !important;
}}
[data-testid="stSidebar"] .stDownloadButton > button:hover {{ opacity:.88 !important; }}
.mh-ph {{ display:flex; align-items:center; gap:12px; padding:14px 20px; background:#fff;
    border:1px solid #e8e8e4; border-radius:12px; margin-bottom:16px; }}
.mh-pi {{ width:38px; height:38px; border-radius:10px; background:{ACCENT}; display:flex;
    align-items:center; justify-content:center; font-size:18px; color:#fff; flex-shrink:0; }}
.mh-pt {{ font-size:17px; font-weight:800; color:#111; letter-spacing:-.4px; }}
.mh-ps {{ font-size:11px; color:#aaa; margin-top:1px; }}
.mh-sec {{ font-size:9.5px; font-weight:700; letter-spacing:1.3px; text-transform:uppercase;
    color:#c0bfbb; display:block; margin-bottom:6px; }}
.stButton > button {{ font-family:'Inter',sans-serif !important; font-weight:600 !important;
    border-radius:8px !important; font-size:12.5px !important; height:36px !important; transition:all .13s ease !important; }}
.stButton > button[kind="primary"] {{ background:{ACCENT} !important; border:2px solid {ACCENT} !important;
    color:#fff !important; box-shadow:0 1px 3px rgba(0,0,0,.15) !important; }}
.stButton > button[kind="primary"]:hover {{ opacity:.88 !important; transform:translateY(-1px) !important;
    box-shadow:0 4px 12px rgba(0,0,0,.2) !important; }}
.stButton > button[kind="primary"]:disabled {{ background:#e6e6e4 !important; border-color:#e6e6e4 !important;
    color:#bbb !important; box-shadow:none !important; transform:none !important; opacity:1 !important; }}
.stButton > button[kind="secondary"] {{ background:#fff !important; border:1.5px solid #ddd !important; color:#555 !important; }}
.stButton > button[kind="secondary"]:hover {{ border-color:{ACCENT} !important; color:{ACCENT} !important; }}
.mh-big .stButton > button {{ height:46px !important; font-size:14px !important; font-weight:800 !important; }}
.stDownloadButton > button {{ font-family:'Inter',sans-serif !important; font-weight:600 !important;
    border-radius:8px !important; font-size:12.5px !important; height:36px !important;
    background:{ACCENT} !important; border:none !important; color:#fff !important; }}
[data-testid="stFileUploader"] {{ background:#fff !important; border:1.5px dashed #e4e4e0 !important; border-radius:8px !important; }}
[data-testid="stMetric"] {{ background:#fff; border:1px solid #e8e8e4; border-radius:10px; padding:.75rem .9rem !important; }}
[data-testid="stMetricLabel"] p {{ font-size:9.5px !important; font-weight:700 !important; color:#c0bfbb !important;
    text-transform:uppercase !important; letter-spacing:.6px !important; }}
[data-testid="stMetricValue"] {{ font-size:22px !important; font-weight:800 !important; color:#111 !important; letter-spacing:-.7px !important; }}
.vp {{ height:4px; background:#f0f0ee; border-radius:99px; overflow:hidden; margin:6px 0; }}
.vf {{ height:100%; border-radius:99px; background:{ACCENT}; transition:width .35s; }}
.mh-log {{ background:#18181b; border-radius:8px; padding:10px 12px;
    font-family:'JetBrains Mono','Courier New',monospace; font-size:10.5px; line-height:1.8;
    max-height:200px; overflow-y:auto; margin-top:6px; }}
.mh-log::-webkit-scrollbar {{ width:4px; }}
.mh-log::-webkit-scrollbar-thumb {{ background:#3f3f46; border-radius:2px; }}
.lr {{ color:#fff; font-weight:700; border-top:1px solid #27272a; margin-top:4px; padding-top:4px; }}
.lr:first-child {{ border-top:none; margin-top:0; padding-top:0; }}
.lo {{ color:#4ade80; font-weight:600; }}
.lf {{ color:#f87171; }}
.ls {{ color:#fb923c; }}
.li {{ color:#3f3f46; }}
.lx {{ color:#22d3ee; font-weight:700; }}
.mh-info {{ background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px; padding:8px 13px;
    font-size:12px; color:#15803d; font-weight:600; margin:4px 0; }}
.mh-warn {{ background:#fff1f2; border:1px solid #fecdd3; border-radius:8px; padding:8px 13px;
    font-size:12px; color:#be123c; font-weight:600; margin:4px 0; }}
.cp {{ background:#fafaf8; border:1px solid #e8e8e4; border-radius:8px; padding:10px 14px; margin:6px 0; font-size:11.5px; }}
.cp-l {{ font-size:9.5px; font-weight:700; color:#999; text-transform:uppercase; letter-spacing:1px; margin-bottom:4px; }}
.cp-v {{ font-family:'JetBrains Mono',monospace; font-size:11px; color:#333; line-height:1.6; }}
hr {{ border-color:#eee !important; margin:10px 0 !important; }}
</style>""", unsafe_allow_html=True)

# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div style="font-size:17px;font-weight:800;color:#fff;letter-spacing:-.3px;margin-bottom:4px">CSV Validator</div>', unsafe_allow_html=True)
    st.markdown('<div style="font-size:10px;color:#555;margin-bottom:16px">early-stop fallback validation</div>', unsafe_allow_html=True)
    st.divider()
    res = st.session_state.get("cv_results", [])
    if res:
        orig_cols = st.session_state.get("cv_original_cols", [])
        xlsx = build_xlsx(res, orig_cols)
        st.download_button("Export .xlsx", xlsx,
            f"validated_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="sx_xlsx", use_container_width=True)
        st.divider()
        nv = sum(1 for r in res if r.get("val"))
        nd = sum(1 for r in res if (r.get("val",{}) or {}).get("status")=="Deliverable")
        nf = sum(1 for r in res if r.get("was_fallback"))
        tc = sum(len(r.get("val_log",[])) for r in res)
        saved = tc - nv
        st.markdown(f'<div style="font-size:11px;color:#666;line-height:2.2">'
            f'Total rows: <strong style="color:#fff">{len(res)}</strong><br>'
            f'Checked: <strong style="color:#fff">{nv}</strong><br>'
            f'Deliverable: <strong style="color:#4ade80">{nd}</strong><br>'
            f'Fallbacks: <strong style="color:#22d3ee">{nf}</strong><br>'
            f'Checks saved: <strong style="color:#22d3ee">{saved}</strong></div>', unsafe_allow_html=True)
    st.divider()
    st.markdown('<div style="font-size:9px;color:#333;line-height:1.8">'
        'Upload CSV -> pick 2 columns<br>'
        'Best Email -> validated first<br>'
        'All Emails -> fallback pool<br>'
        'Empty rows kept in XLSX<br>'
        'Stops on first Deliverable</div>', unsafe_allow_html=True)

# ── Session State ─────────────────────────────────────────────────────────────
for k, v in {"cv_results":[],"cv_running":False,"cv_idx":0,"cv_log":[],"cv_df":None,
             "cv_queue":[],"cv_original_cols":[]}.items():
    if k not in st.session_state: st.session_state[k] = v

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="mh-ph">
  <div class="mh-pi">✅</div>
  <div>
    <div class="mh-pt">CSV Email Validator</div>
    <div class="mh-ps">upload CSV · pick columns · early-stop on first Deliverable · keeps all original data</div>
  </div>
</div>""", unsafe_allow_html=True)

# ── Upload ────────────────────────────────────────────────────────────────────
st.markdown('<span class="mh-sec">Upload CSV</span>', unsafe_allow_html=True)
uploaded = st.file_uploader("Choose a CSV file", type=["csv"], key="csv_up")
df = None
if uploaded:
    try:
        df = pd.read_csv(uploaded)
        st.session_state.cv_df = df
        cols = list(df.columns)
        st.markdown(f'<div class="mh-info">Loaded <strong>{len(df)}</strong> rows · <strong>{len(cols)}</strong> columns</div>', unsafe_allow_html=True)
        st.caption("Preview (first 5 rows)")
        st.dataframe(df.head(5), use_container_width=True, hide_index=True, height=160)
    except Exception as e:
        st.error(f"Failed to parse CSV: {e}")

# ── Column Selection ─────────────────────────────────────────────────────────
if df is not None:
    cols = list(df.columns)
    bh = ["best email","best_email","primary email","email","mail","contact email","contact_email"]
    ah = ["all emails","all_emails","emails","other emails","other_emails","additional emails","additional_emails","alt emails","fallback"]
    dh = ["domain","website","site","url","company"]
    db = next((c for c in cols if any(h in c.lower() for h in bh)), cols[0])
    da = next((c for c in cols if any(h in c.lower() for h in ah)), None)
    dd = next((c for c in cols if any(h in c.lower() for h in dh)), None)

    st.divider()
    st.markdown('<span class="mh-sec">Column Mapping</span>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3, gap="large")
    with c1:
        st.markdown('<div style="font-size:12px;font-weight:700;color:#111;margin-bottom:4px">Best Email *</div>', unsafe_allow_html=True)
        best_col = st.selectbox("b", cols, index=cols.index(db), key="s_b", label_visibility="collapsed")
    with c2:
        st.markdown('<div style="font-size:12px;font-weight:700;color:#111;margin-bottom:4px">All Emails (fallback pool)</div>', unsafe_allow_html=True)
        all_col = st.selectbox("a", ["— None —"] + cols, index=(cols.index(da)+1 if da else 0), key="s_a", label_visibility="collapsed")
    with c3:
        st.markdown('<div style="font-size:12px;font-weight:700;color:#111;margin-bottom:4px">Domain (optional)</div>', unsafe_allow_html=True)
        dom_col = st.selectbox("d", ["— Auto —"] + cols, index=(cols.index(dd)+1 if dd else 0), key="s_d", label_visibility="collapsed")

    st.markdown(f'<div class="cp"><div class="cp-l">Best Email column</div><div class="cp-v">' +
                "<br>".join(str(v)[:60] for v in df[best_col].head(3).values) + '</div></div>', unsafe_allow_html=True)
    if all_col != "— None —":
        st.markdown(f'<div class="cp"><div class="cp-l">All Emails column</div><div class="cp-v">' +
                    "<br>".join(str(v)[:80] for v in df[all_col].head(3).values) + '</div></div>', unsafe_allow_html=True)

    queue = []
    for i, row in df.iterrows():
        br = str(row[best_col]).strip() if pd.notna(row[best_col]) else ""
        be = br if is_valid_email(br) else ""
        ar = str(row[all_col]).strip() if (all_col != "— None —" and pd.notna(row.get(all_col))) else ""
        ae = parse_email_cell(ar) if ar else []
        ae = [e for e in ae if e.lower() != be.lower()]
        
        has_emails = bool(be or ae)

        if not has_emails:
            dom = f"row_{i+1}"
        elif dom_col != "— Auto —" and pd.notna(row.get(dom_col)):
            dom = str(row[dom_col]).strip()
        elif be: dom = be.split("@")[-1]
        else: dom = ae[0].split("@")[-1]
        
        queue.append({
            "row_idx": i+1, 
            "domain": dom, 
            "original_email": be, 
            "all_emails": ae,
            "has_emails": has_emails,
            "original_row_data": row.to_dict()
        })

    nv = sum(1 for q in queue if q["has_emails"])
    n_empty = len(queue) - nv
    st.divider()
    if nv:
        empty_txt = f" · <strong>{n_empty}</strong> empty rows (retained in export)" if n_empty else ""
        st.markdown(f'<div class="mh-info">Validatable: <strong>{nv}</strong> rows{empty_txt}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="mh-warn">No valid emails found in any row. Check your column mapping.</div>', unsafe_allow_html=True)

    running = st.session_state.cv_running
    vc1, vc2, vc3 = st.columns([3, 1, 2])
    with vc1:
        st.markdown('<div class="mh-big">', unsafe_allow_html=True)
        if not running:
            if st.button(f"Validate {nv} row(s)", type="primary", use_container_width=True, disabled=not nv, key="cv_go"):
                st.session_state.cv_results = []; st.session_state.cv_idx = 0
                st.session_state.cv_log = []; st.session_state.cv_running = True
                st.session_state.cv_queue = queue
                st.session_state.cv_original_cols = list(df.columns)
                st.rerun()
        else:
            if st.button("Stop", type="secondary", use_container_width=True, key="cv_stop"):
                st.session_state.cv_running = False; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with vc2:
        if st.session_state.cv_results:
            if st.button("Clear", type="secondary", use_container_width=True, key="cv_clr"):
                st.session_state.cv_results = []; st.session_state.cv_log = []; st.rerun()
    with vc3:
        st.markdown('<div style="font-size:10.5px;color:#aaa;padding-top:12px">~3-8s/email · stops on first Deliverable</div>', unsafe_allow_html=True)

    res = st.session_state.cv_results; cq = st.session_state.get("cv_queue", queue); ci = st.session_state.cv_idx; tot = len(cq)
    if running and tot > 0:
        # Only count valid emails towards progress percentage visually
        valid_done = sum(1 for r in res if r.get("has_emails"))
        pct = round(valid_done/nv*100, 1) if nv else 0
        cur = cq[ci] if ci < tot else None
        ce = cur.get("original_email") or (cur["all_emails"][0] if cur and cur.get("all_emails") else "—")
        st.markdown(f'<div style="font-size:12px;font-weight:700;color:#111;margin:6px 0 2px">'
            f'Validating {valid_done}/{nv} — <code style="color:{ACCENT}">{ce[:40]}</code></div>'
            f'<div class="vp"><div class="vf" style="width:{pct}%"></div></div>'
            f'<div style="font-size:20px;font-weight:800;color:{ACCENT};text-align:right;margin-top:-4px">{pct}%</div>',
            unsafe_allow_html=True)

    ll = st.session_state.cv_log
    if ll:
        h = ""
        for kind, text in ll[-80:]:
            if   kind == "row":  h += f'<div class="lr">[ {text} ]</div>'
            elif kind == "try":  h += f'<div class="li">  -> {text}</div>'
            elif kind == "ok":   h += f'<div class="lo">  OK {text}</div>'
            elif kind == "fail": h += f'<div class="lf">  FAIL {text}</div>'
            elif kind == "skip": h += f'<div class="ls">  SKIP {text}</div>'
            elif kind == "stop": h += f'<div class="lx">  STOP - {text}</div>'
        st.markdown(f'<div class="mh-log">{h}</div>', unsafe_allow_html=True)

    if res:
        nvc = sum(1 for r in res if r.get("val"))
        nd = sum(1 for r in res if (r.get("val",{}) or {}).get("status")=="Deliverable")
        nri = sum(1 for r in res if (r.get("val",{}) or {}).get("status")=="Risky")
        nb = sum(1 for r in res if (r.get("val",{}) or {}).get("status")=="Not Deliverable")
        nfb = sum(1 for r in res if r.get("was_fallback"))
        tc = sum(len(r.get("val_log",[])) for r in res)
        m1,m2,m3,m4,m5,m6 = st.columns(6)
        m1.metric("Checked", nvc); m2.metric("Deliverable", nd)
        m3.metric("Risky", nri); m4.metric("Failed", nb)
        m5.metric("Fallback", nfb); m6.metric("Emails Checked", tc)

    if res:
        srch = st.text_input("Filter...", placeholder="domain or email...", label_visibility="collapsed", key="cv_sr")
        rows = []
        for r in res:
            v = r.get("val") or {}; s = v.get("status",""); em = r.get("validated_email","")
            orig = r.get("original_email",""); fb = r.get("was_fallback"); cf = r.get("confidence")
            
            if not r.get("has_emails"):
                rows.append({"#":r["row_idx"],"Domain":r["domain"],"Original":"(empty)","Validated":"—",
                    "Status":"Skipped","Tier":"—","Score":"—","Reason":"—",
                    "SPF":"—","DMARC":"—","CA":"—","Pool":0,"Checked":0})
                continue
                
            ed = f"{em} (fallback)" if fb and orig != em else em
            od = orig if orig else "(pool)"
            rows.append({"#":r["row_idx"],"Domain":r["domain"],"Original":od,"Validated":ed,
                "Status": s if s else "Pending","Tier":tier_short(em) if em else "—",
                "Score":cf if cf is not None else "—","Reason":v.get("reason","—") if v else "—",
                "SPF":("Yes" if v.get("spf") else "No") if v else "—",
                "DMARC":("Yes" if v.get("dmarc") else "No") if v else "—",
                "CA":("Yes" if v.get("catch_all") else "No") if v else "—",
                "Pool":len(r["all_emails"]),"Checked":len(r.get("val_log",[]))})
        dr = pd.DataFrame(rows)
        if srch:
            m = (dr["Domain"].str.contains(srch,case=False,na=False)|dr["Original"].str.contains(srch,case=False,na=False)|dr["Validated"].str.contains(srch,case=False,na=False))
            dr = dr[m]
        st.caption(f'**{len(dr)}** of {len(res)}  |  (fallback) = switched email  |  Pool = fallback size  |  Checked = emails validated before stop')
        st.dataframe(dr, use_container_width=True, hide_index=True,
            height=min(560, 44+max(len(dr),1)*36),
            column_config={"#":st.column_config.NumberColumn("#",width=45),
                "Domain":st.column_config.TextColumn("Domain",width=140),
                "Original":st.column_config.TextColumn("Original Email",width=200),
                "Validated":st.column_config.TextColumn("Validated Email",width=220),
                "Status":st.column_config.TextColumn("Status",width=140),
                "Tier":st.column_config.TextColumn("Tier",width=60),
                "Score":st.column_config.NumberColumn("Score",width=50),
                "Reason":st.column_config.TextColumn("Reason",width=150),
                "SPF":st.column_config.TextColumn("SPF",width=45),
                "DMARC":st.column_config.TextColumn("DMARC",width=50),
                "CA":st.column_config.TextColumn("CA",width=45),
                "Pool":st.column_config.NumberColumn("Pool",width=45),
                "Checked":st.column_config.NumberColumn("Checked",width=60)})

    if st.session_state.cv_running:
        q = st.session_state.cv_queue; idx = st.session_state.cv_idx; tot = len(q)
        if idx >= tot:
            st.session_state.cv_running = False; st.rerun()
        else:
            item = q[idx]; rn = item["row_idx"]; dom = item["domain"]
            orig_data = item.get("original_row_data", {})
            
            # BATCH SKIP EMPTY ROWS TO PREVENT UI STUTTERING
            if not item.get("has_emails"):
                st.session_state.cv_results.append({
                    "row_idx":rn,"domain":dom,"original_email":"",
                    "validated_email":"","all_emails":[],"val":None,"was_fallback":False,
                    "confidence":None,"val_log":[], "original_row_data": orig_data,
                    "has_emails": False
                })
                next_idx = idx + 1
                # Look ahead and grab all consecutive empty rows so we skip them in 1 frame
                while next_idx < tot and not q[next_idx].get("has_emails"):
                    n_item = q[next_idx]
                    st.session_state.cv_results.append({
                        "row_idx": n_item["row_idx"], "domain": n_item["domain"], "original_email":"",
                        "validated_email":"","all_emails":[],"val":None,"was_fallback":False,
                        "confidence":None,"val_log":[], "original_row_data": n_item.get("original_row_data", {}),
                        "has_emails": False
                    })
                    next_idx += 1
                
                st.session_state.cv_idx = next_idx
                if st.session_state.cv_idx >= tot: st.session_state.cv_running = False
                st.rerun()

            # NORMAL VALIDATION FLOW
            else:
                best = item["original_email"]; ae = item["all_emails"]
                st.session_state.cv_log.append(("row", f"Row {rn} - {dom}"))
                
                st.session_state.cv_log.append(("try", f"Best: {best or '(none)'}"))
                val_em, val_res, was_fb, vlog = validate_with_early_stop(best, ae)
                for ce, cs, cr in vlog:
                    if cs == "Deliverable":
                        st.session_state.cv_log.append(("ok", f"{ce} - DELIVERABLE"))
                        st.session_state.cv_log.append(("stop", f"Found deliverable, stopping"))
                    elif cs == "Risky":
                        st.session_state.cv_log.append(("try", f"{ce} - Risky ({cr})"))
                    elif cs == "Not Deliverable":
                        st.session_state.cv_log.append(("fail", f"{ce} - {cr}"))
                    else:
                        st.session_state.cv_log.append(("skip", f"{ce} - {cr}"))
                if was_fb:
                    st.session_state.cv_log.append(("ok", f"Fallback: {best} -> {val_em}"))
                cf = confidence_score(val_em, val_res) if val_res else None
                st.session_state.cv_results.append({"row_idx":rn,"domain":dom,"original_email":best,
                    "validated_email":val_em,"all_emails":ae,"val":val_res,"was_fallback":was_fb,
                    "confidence":cf,"val_log":vlog, "original_row_data": orig_data, "has_emails": True})
                st.session_state.cv_idx = idx + 1
                if st.session_state.cv_idx >= tot: st.session_state.cv_running = False
                st.rerun()

if df is None and not st.session_state.cv_results:
    st.markdown("""
    <div style="text-align:center;padding:60px 0">
        <div style="font-size:48px;opacity:.08;margin-bottom:16px">✅</div>
        <div style="font-size:18px;font-weight:800;color:#111;margin-bottom:10px">Upload a CSV to start</div>
        <div style="font-size:12.5px;color:#aaa;line-height:2;max-width:420px;margin:0 auto">
            Your CSV should have at least one column with emails.<br>
            Optionally a second column with <strong>additional emails</strong><br>
            (semicolon, comma, or pipe separated) as fallback pool.<br><br>
            <strong style="color:#16a34a">How it works:</strong><br>
            1. Empty rows are <strong>kept in export</strong> but skipped during validation<br>
            2. Validates the <strong>Best Email</strong> first<br>
            3. If not deliverable -> tries <strong>All Emails</strong> in tier order<br>
            4. <strong>Stops immediately</strong> when first Deliverable is found<br>
            5. Exports XLSX with <strong>all your original data kept intact</strong> + validation appended
        </div>
    </div>""", unsafe_allow_html=True)
