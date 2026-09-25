"""
Kentucky Wheat Yield Contest - Digital Entry Form
University of Kentucky Cooperative Extension
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import smtplib, base64, json, time, datetime, requests
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path

# ─────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────
EXCEL_FILE      = "wheat_contest_entries.xlsx"
RECIPIENT_EMAIL = "mshamim11@uky.edu"
RECIPIENT_NAME  = "Mohammad Jan Shamim"

# ── UK-registered Azure app (registered in UK's own tenant) ──────────
OAUTH_CLIENT_ID  = "546960b8-978a-4776-ad35-dcb8a8bd20f4"
OAUTH_TENANT     = "2b30530b-69b6-4457-b818-481cb53d42ae"
OAUTH_SCOPE      = "https://graph.microsoft.com/Mail.Send offline_access"
OAUTH_DEVICE_URL = f"https://login.microsoftonline.com/{OAUTH_TENANT}/oauth2/v2.0/devicecode"
OAUTH_TOKEN_URL  = f"https://login.microsoftonline.com/{OAUTH_TENANT}/oauth2/v2.0/token"
GRAPH_SEND_URL   = "https://graph.microsoft.com/v1.0/me/sendMail"

KY_COUNTIES = sorted([
    "Adair","Allen","Anderson","Ballard","Barren","Bath","Bell","Boone",
    "Bourbon","Boyd","Boyle","Bracken","Breathitt","Breckinridge","Bullitt",
    "Butler","Caldwell","Calloway","Campbell","Carlisle","Carroll","Carter",
    "Casey","Christian","Clark","Clay","Clinton","Crittenden","Cumberland",
    "Daviess","Edmonson","Elliott","Estill","Fayette","Fleming","Floyd",
    "Franklin","Fulton","Gallatin","Garrard","Grant","Graves","Grayson",
    "Green","Greenup","Hancock","Hardin","Harlan","Harrison","Hart",
    "Henderson","Henry","Hickman","Hopkins","Jackson","Jefferson","Jessamine",
    "Johnson","Kenton","Knott","Knox","LaRue","Laurel","Lawrence","Lee",
    "Leslie","Letcher","Lewis","Lincoln","Livingston","Logan","Lyon",
    "Madison","Magoffin","Marion","Marshall","Martin","Mason","McCracken",
    "McCreary","McLean","Meade","Menifee","Mercer","Metcalfe","Monroe",
    "Montgomery","Morgan","Muhlenberg","Nelson","Nicholas","Ohio","Oldham",
    "Owen","Owsley","Pendleton","Perry","Pike","Powell","Pulaski",
    "Robertson","Rockcastle","Rowan","Russell","Scott","Shelby","Simpson",
    "Spencer","Taylor","Todd","Trigg","Trimble","Union","Warren","Washington",
    "Wayne","Webster","Whitley","Wolfe","Woodford"
])

COUNTY_AREA = {
    "Ballard":1,"Calloway":1,"Carlisle":1,"Fulton":1,"Graves":1,"Hickman":1,
    "Marshall":1,"McCracken":1,"Caldwell":1,"Christian":1,"Crittenden":1,
    "Hopkins":1,"Lyon":1,"Todd":1,"Trigg":1,
    "Daviess":2,"Hancock":2,"Henderson":2,"McLean":2,"Muhlenberg":2,
    "Ohio":2,"Union":2,"Webster":2,
    "Adair":3,"Allen":3,"Barren":3,"Butler":3,"Edmonson":3,"Hart":3,
    "Logan":3,"Metcalfe":3,"Monroe":3,"Simpson":3,"Warren":3,
}

HEADER_FILL = "2E4057"

# ─────────────────────────────────────────────────────────
# OAUTH2  —  Device Code Flow
# ─────────────────────────────────────────────────────────

def oauth_start_device_flow() -> dict:
    resp = requests.post(OAUTH_DEVICE_URL, data={
        "client_id": OAUTH_CLIENT_ID,
        "scope":     OAUTH_SCOPE,
    }, timeout=10)
    resp.raise_for_status()
    return resp.json()


def oauth_poll_for_token(device_code: str, interval: int = 5, max_wait: int = 120) -> dict | None:
    deadline = time.time() + max_wait
    while time.time() < deadline:
        time.sleep(interval)
        resp = requests.post(OAUTH_TOKEN_URL, data={
            "client_id":   OAUTH_CLIENT_ID,
            "grant_type":  "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device_code,
        }, timeout=10)
        data = resp.json()
        if "access_token" in data:
            return data
        if data.get("error") == "authorization_pending":
            continue
        if data.get("error") == "expired_token":
            return None
        raise RuntimeError(data.get("error_description", data.get("error", "Unknown OAuth error")))
    return None


def oauth_refresh_token(refresh_token: str) -> dict | None:
    resp = requests.post(OAUTH_TOKEN_URL, data={
        "client_id":     OAUTH_CLIENT_ID,
        "grant_type":    "refresh_token",
        "refresh_token": refresh_token,
        "scope":         OAUTH_SCOPE,
    }, timeout=10)
    data = resp.json()
    return data if "access_token" in data else None


def oauth_send_email(access_token: str, sender_email: str,
                     recipient_email: str, subject: str, body_text: str,
                     filepath: str, cc: str | None = None) -> None:
    with open(filepath, "rb") as f:
        attachment_b64 = base64.b64encode(f.read()).decode()

    cc_recipients = [{"emailAddress": {"address": cc}}] if cc else []

    payload = {
        "message": {
            "subject": subject,
            "body":    {"contentType": "Text", "content": body_text},
            "toRecipients": [{"emailAddress": {"address": recipient_email}}],
            "ccRecipients":  cc_recipients,
            "attachments": [{
                "@odata.type":  "#microsoft.graph.fileAttachment",
                "name":         Path(filepath).name,
                "contentType":  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "contentBytes": attachment_b64,
            }],
        },
        "saveToSentItems": True,
    }

    resp = requests.post(
        GRAPH_SEND_URL,
        headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
        json=payload, timeout=30,
    )
    if resp.status_code not in (200, 202):
        raise RuntimeError(f"Graph API error {resp.status_code}: {resp.text}")


# ─────────────────────────────────────────────────────────
# SMTP FALLBACK
# ─────────────────────────────────────────────────────────

def smtp_send_email(smtp_server, smtp_port, sender_email, sender_password,
                    recipient_email, subject, body_text, filepath, cc=None):
    msg = MIMEMultipart()
    msg["From"]    = sender_email
    msg["To"]      = recipient_email
    msg["Subject"] = subject
    if cc:
        msg["Cc"] = cc
    msg.attach(MIMEText(body_text, "plain"))
    with open(filepath, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{Path(filepath).name}"')
    msg.attach(part)
    recipients = [recipient_email] + ([cc] if cc else [])
    with smtplib.SMTP(smtp_server, int(smtp_port), timeout=15) as server:
        server.ehlo(); server.starttls(); server.ehlo()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_string())


# ─────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────

def _today():
    return datetime.date.today()

def _blank_defaults():
    today = _today()
    return {
        "county": "— Select —",
        "producer_name": "", "producer_address": "", "producer_town": "",
        "producer_zip": "", "producer_phone": "", "producer_mobile": "",
        "profession": "", "harvest_date": today,
        "supervisor_name": "", "supervisor_sig_date": today,
        "division": "Division I - Tillage (conv./min.)",
        "previous_crop": "Corn",
        "planting_date": datetime.date(today.year - 1, 10, 1),
        "wheat_variety": "", "row_width": "", "seeding_rate": "",
        "fall_n": "0", "fall_p": "0", "fall_k": "0", "fall_other": "",
        "ws_n1_date": datetime.date(today.year, 3, 1), "ws_n1_rate": "0",
        "ws_n2_date": datetime.date(today.year, 4, 1), "ws_n2_rate": "0",
        "manure_used": "No", "manure_type": "", "manure_tons": "0", "manure_date": today,
        "growth_reg": "", "fall_pest": "", "spring_pest": "",
        "heading_pest": "", "biologicals": "", "tillage_used": "",
        "h_length": 0.0, "h_width": 0.0,
        "moisture_readings": [], "_moisture_gen": 0,
        "test_weight": 60.0, "grain_weight": 0.0,
        "agent_notes": "",
        "h_area_ft2": 0.0, "h_acres": 0.0, "gm_avg": 0.0, "official_yield": 0.0,
        "agreement_checked": False, "_agree_gen": 0,
    }

def _init_state():
    for k, v in _blank_defaults().items():
        if k not in st.session_state:
            st.session_state[k] = v
    for k, v in {
        "oauth_access_token": None, "oauth_refresh_token": None,
        "oauth_token_expiry": 0,    "oauth_user_email": None,
        "oauth_device_code": None,  "oauth_polling": False,
        "oauth_interval": 5,        "auth_mode": "oauth",
        "_agree_gen": 0,
    }.items():
        if k not in st.session_state:
            st.session_state[k] = v

def _clear_form():
    for k, v in _blank_defaults().items():
        st.session_state[k] = v
    st.rerun()

# ─────────────────────────────────────────────────────────
# LIVE RECALCULATION
# ─────────────────────────────────────────────────────────

def _recompute():
    length = float(st.session_state.get("h_length", 0.0) or 0.0)
    width  = float(st.session_state.get("h_width",  0.0) or 0.0)
    ft2 = length * width
    acres = ft2 / 43560.0
    st.session_state["h_area_ft2"] = round(ft2, 1)
    st.session_state["h_acres"]    = round(acres, 4)
    readings = [r for r in st.session_state.get("moisture_readings", []) if r > 0]
    gm_avg = sum(readings) / len(readings) if readings else 0.0
    st.session_state["gm_avg"] = round(gm_avg, 2)
    gw = float(st.session_state.get("grain_weight", 0.0) or 0.0)
    yld = gw * ((100 - gm_avg) / 86.5) / 60 / acres if (gw > 0 and gm_avg > 0 and acres > 0) else 0.0
    st.session_state["official_yield"] = round(yld, 2)

def _add_reading_cb():
    gen = st.session_state.get("_moisture_gen", 0)
    val = float(st.session_state.get(f"moisture_input_{gen}", 0.0) or 0.0)
    if val > 0:
        readings = st.session_state.get("moisture_readings", [])
        if len(readings) < 3:
            readings.append(round(val, 1))
            st.session_state["moisture_readings"] = readings
            st.session_state["_moisture_gen"] = gen + 1
            _recompute()
    else:
        st.session_state["_moisture_warn"] = True

def _clear_readings_cb():
    st.session_state["moisture_readings"] = []
    st.session_state["_moisture_gen"] = st.session_state.get("_moisture_gen", 0) + 1
    _recompute()

# ─────────────────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────────────────

def _thin_border():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr(ws, row, col, text, bg="2E4057", fg="FFFFFF", bold=True, size=10):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(bold=bold, color=fg, name="Arial", size=size)
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = _thin_border()
    return c

def _dat(ws, row, col, value="", bg="FFFFFF"):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Arial", size=10)
    c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = _thin_border()
    return c

COLUMNS = [
    "Entry_ID","Submission_Date","County","Area",
    "Producer_Name","Producer_Address","Producer_Town","Producer_Zip",
    "Producer_Phone","Producer_Mobile","Profession",
    "Harvest_Date","Supervisor_Name","Supervisor_Signature_Date",
    "Division","Previous_Crop","Planting_Date","Wheat_Variety",
    "Row_Width_inches","Seeding_Rate",
    "Fall_N_lbA","Fall_P2O5_lbA","Fall_K2O_lbA","Fall_Other_Fertilizer",
    "Winter_Spring_N1_Date","Winter_Spring_N1_lbA",
    "Winter_Spring_N2_Date","Winter_Spring_N2_lbA",
    "Manure_Used","Manure_Type","Manure_TonsA","Manure_Date",
    "Growth_Regulator","Fall_Pest_Products","Spring_Pest_Products",
    "Heading_Flowering_Pest","Biologicals_Other","Tillage_Used",
    "Harvest_Length_ft","Harvest_Width_ft","Harvest_Area_ft2","Harvest_Acres",
    "Grain_Moisture_1","Grain_Moisture_2","Grain_Moisture_3","Grain_Moisture_Avg",
    "Test_Weight_lbbu","Grain_Weight_lbs","Official_Yield_BuAcre","Agent_Notes",
]

SECTION_SPANS = [
    ("Producer / Agent Info",  1, 14, "1F4E79"),
    ("Agronomic Data",        15, 20, "375623"),
    ("Fertilizer",            21, 32, "7B3F00"),
    ("Pest Management",       33, 38, "6B2737"),
    ("Harvest Area",          39, 42, "4A235A"),
    ("Grain Characteristics", 43, 47, "7E5109"),
    ("Yield Calculation",     48, 49, "1A5276"),
    ("Notes",                 50, 50, "555555"),
]

def build_excel_with_entry(data: dict, filepath: str) -> int:
    area = COUNTY_AREA.get(data.get("County", ""), 4)
    data["Area"] = area
    if Path(filepath).exists():
        wb = openpyxl.load_workbook(filepath)
        ws = wb.active
        try:
            last_id = int(ws.cell(row=ws.max_row, column=1).value or 0)
        except Exception:
            last_id = ws.max_row - 3
        entry_id = last_id + 1
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Wheat Contest Entries"
        ws.merge_cells(f"A1:{get_column_letter(len(COLUMNS))}1")
        t = ws["A1"]
        t.value = "Kentucky Wheat Yield Contest - Master Entry Database"
        t.font = Font(bold=True, size=14, color="FFFFFF", name="Arial")
        t.fill = PatternFill("solid", fgColor=HEADER_FILL)
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28
        for label, c1, c2, color in SECTION_SPANS:
            ws.merge_cells(start_row=2, start_column=c1, end_row=2, end_column=c2)
            _hdr(ws, 2, c1, label, bg=color, size=9)
        ws.row_dimensions[2].height = 18
        for ci, col in enumerate(COLUMNS, 1):
            _hdr(ws, 3, ci, col.replace("_", " "), bg="4A4A4A", size=9)
        ws.row_dimensions[3].height = 42
        widths = [7,14,14,5,20,22,14,7,13,13,14,12,20,14,
                  14,14,12,18,8,16,8,8,8,20,12,8,12,8,8,14,8,12,
                  18,22,22,22,18,14,11,11,11,9,9,9,9,9,10,12,13,28]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A4"
        entry_id = 1
    data["Entry_ID"] = entry_id
    data["Submission_Date"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    row_num = ws.max_row + 1
    bg = "F7F9FC" if entry_id % 2 == 0 else "FFFFFF"
    for ci, col in enumerate(COLUMNS, 1):
        _dat(ws, row_num, ci, data.get(col, ""), bg=bg)
    wb.save(filepath)
    return entry_id

# ─────────────────────────────────────────────────────────
# OAUTH SIDEBAR PANEL
# ─────────────────────────────────────────────────────────

def _oauth_sidebar():
    st.markdown("**Microsoft Sign-In (Recommended)**")
    st.caption("Works with MFA. Signs in with your @uky.edu account. No password stored.")

    access_token  = st.session_state.get("oauth_access_token")
    refresh_token = st.session_state.get("oauth_refresh_token")
    token_expiry  = st.session_state.get("oauth_token_expiry", 0)
    user_email    = st.session_state.get("oauth_user_email")

    # Silent refresh
    if refresh_token and time.time() > token_expiry - 60:
        try:
            new_tokens = oauth_refresh_token(refresh_token)
            if new_tokens:
                st.session_state["oauth_access_token"] = new_tokens["access_token"]
                st.session_state["oauth_refresh_token"] = new_tokens.get("refresh_token", refresh_token)
                st.session_state["oauth_token_expiry"]  = time.time() + new_tokens.get("expires_in", 3600)
                access_token = new_tokens["access_token"]
        except Exception:
            pass

    # Already signed in
    if access_token and time.time() < token_expiry:
        st.success(f"Signed in as **{user_email or 'your UK account'}**")
        if st.button("Sign Out", use_container_width=True):
            for k in ["oauth_access_token","oauth_refresh_token","oauth_user_email","oauth_device_code"]:
                st.session_state[k] = None
            st.session_state["oauth_token_expiry"] = 0
            st.session_state["oauth_polling"] = False
            st.rerun()
        return True, user_email

    # Device code flow
    polling  = st.session_state.get("oauth_polling", False)
    dev_code = st.session_state.get("oauth_device_code")

    if not polling:
        if st.button("Sign in with Microsoft", use_container_width=True, type="primary"):
            try:
                flow = oauth_start_device_flow()
                st.session_state["oauth_device_code"] = flow["device_code"]
                st.session_state["oauth_interval"]    = flow.get("interval", 5)
                st.session_state["oauth_polling"]     = True
                st.session_state["_oauth_user_code"]  = flow["user_code"]
                st.session_state["_oauth_verify_url"] = flow["verification_uri"]
                st.rerun()
            except Exception as e:
                st.error(f"Could not start sign-in: {e}")
    else:
        user_code  = st.session_state.get("_oauth_user_code", "")
        verify_url = st.session_state.get("_oauth_verify_url", "https://microsoft.com/devicelogin")
        st.info(
            f"**Step 1** — Open this link in any browser:\n\n"
            f"**{verify_url}**\n\n"
            f"**Step 2** — Enter this code: **`{user_code}`**\n\n"
            f"**Step 3** — Sign in with your @uky.edu account,\n"
            f"then click the button below."
        )
        col1, col2 = st.columns(2)
        with col1:
            if st.button("I signed in — continue", use_container_width=True, type="primary"):
                try:
                    tokens = oauth_poll_for_token(
                        st.session_state["oauth_device_code"],
                        interval=st.session_state.get("oauth_interval", 5),
                        max_wait=15,
                    )
                    if tokens:
                        st.session_state["oauth_access_token"]  = tokens["access_token"]
                        st.session_state["oauth_refresh_token"] = tokens.get("refresh_token")
                        st.session_state["oauth_token_expiry"]  = time.time() + tokens.get("expires_in", 3600)
                        me = requests.get(
                            "https://graph.microsoft.com/v1.0/me",
                            headers={"Authorization": f"Bearer {tokens['access_token']}"},
                            timeout=10,
                        ).json()
                        st.session_state["oauth_user_email"] = me.get("mail") or me.get("userPrincipalName", "")
                        st.session_state["oauth_polling"] = False
                        st.rerun()
                    else:
                        st.warning("Not confirmed yet — finish sign-in in your browser first, then try again.")
                except Exception as e:
                    st.error(f"Sign-in error: {e}")
        with col2:
            if st.button("Cancel", use_container_width=True):
                st.session_state["oauth_polling"] = False
                st.rerun()

    return False, None

# ─────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────

def main():
    _init_state()

    st.set_page_config(page_title="KY Wheat Yield Contest", page_icon="🌾",
                       layout="wide", initial_sidebar_state="expanded")

    st.markdown("""
    <style>
    .main{background-color:#f0f4f8}
    .sec-hdr{font-size:1.05rem;font-weight:700;padding:7px 14px;border-radius:5px;
             margin:20px 0 8px 0;color:white}
    .s1{background:#1F4E79}.s2{background:#375623}.s3{background:#7B3F00}
    .s4{background:#6B2737}.s5{background:#4A235A}.s6{background:#7E5109}
    .s7{background:#1A5276}.s8{background:#555555}
    .metric-box{background:#f8f9fa;border:1px solid #dee2e6;border-radius:6px;
                padding:10px 14px;margin-top:4px}
    .metric-label{font-size:0.78rem;color:#6c757d;margin-bottom:2px}
    .metric-value{font-size:1.5rem;font-weight:700;color:#212529}
    .metric-ok{color:#198754}.metric-warn{color:#dc3545}
    .moisture-badge{background:#e9ecef;border-radius:4px;padding:3px 9px;
                    font-size:0.88rem;display:inline-block;margin:2px 3px}
    .agreement-box{background:#fff3cd;border:2px solid #ffc107;border-radius:8px;
                   padding:16px 20px;margin:24px 0 8px 0}
    .result-ok{background:#d1e7dd;border-left:5px solid #198754;border-radius:5px;
               padding:14px 18px;margin-top:12px}
    </style>
    """, unsafe_allow_html=True)

    # ── Header ──────────────────────────────────────────
    col_h, col_clr = st.columns([6, 1])
    with col_h:
        st.markdown("## 🌾 Kentucky Wheat Yield Contest")
        st.caption("University of Kentucky Cooperative Extension — Digital Entry Form")
    with col_clr:
        st.markdown("<div style='margin-top:18px'></div>", unsafe_allow_html=True)
        if st.button("🔄 Clear Form", use_container_width=True):
            _clear_form()
    st.divider()

    # ── Sidebar ─────────────────────────────────────────
    with st.sidebar:
        st.header("Settings")
        st.markdown("**Email Configuration**")
        auth_mode = st.radio("Sign-in method",
                             ["Microsoft Sign-In (OAuth2)", "Password (SMTP)"],
                             index=0 if st.session_state.get("auth_mode") == "oauth" else 1)
        st.session_state["auth_mode"] = "oauth" if "OAuth2" in auth_mode else "smtp"
        st.markdown("---")

        oauth_signed_in, oauth_user_email = False, None
        smtp_server = smtp_port = sender_email = sender_pass = None

        if st.session_state["auth_mode"] == "oauth":
            oauth_signed_in, oauth_user_email = _oauth_sidebar()
        else:
            st.markdown("**SMTP Password Login**")
            st.caption("Only works if your account does not require MFA.")
            smtp_server  = st.text_input("SMTP Server", value="smtp.office365.com")
            smtp_port    = st.number_input("SMTP Port", value=587, step=1)
            sender_email = st.text_input("Your UK Email", placeholder="yourname@uky.edu")
            sender_pass  = st.text_input("UK Email Password", type="password")

        st.markdown("---")
        send_flag = st.checkbox("Auto-send email after saving", value=True)
        cc_self   = st.checkbox("CC myself on submission", value=True)
        effective_sender = oauth_user_email if st.session_state["auth_mode"] == "oauth" else sender_email
        st.divider()
        excel_path = st.text_input("Excel File Path", value=EXCEL_FILE)
        st.divider()
        st.markdown("**Contest Rules**")
        st.info("- Min. 1.5 acres harvested\n- Deadline: July 31\n"
                "- Supervisor must witness harvest\n"
                "- Send grain sample to Colette Laurent, Princeton KY")

    # ════════════════════════════════════════════════════
    # SECTION 1 — PRODUCER / AGENT
    # ════════════════════════════════════════════════════
    st.markdown('<div class="sec-hdr s1">👤 Section 1 — Producer & Agent Information</div>',
                unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.selectbox("County *", ["— Select —"] + KY_COUNTIES, key="county")
        st.text_input("Producer Full Name *", key="producer_name")
        st.text_input("Producer Address", key="producer_address")
    with c2:
        st.text_input("Town", key="producer_town")
        st.text_input("Zip Code", key="producer_zip")
        st.text_input("Phone", key="producer_phone")
    with c3:
        st.text_input("Mobile", key="producer_mobile")
        st.text_input("Profession / Operation Type", key="profession")
    c1, c2 = st.columns(2)
    with c1:
        st.text_input("County Agent / Supervisor Name *", key="supervisor_name")
    with c2:
        st.date_input("Supervisor Sign Date", key="supervisor_sig_date")

    # ════════════════════════════════════════════════════
    # SECTION 2 — AGRONOMIC
    # ════════════════════════════════════════════════════
    st.markdown('<div class="sec-hdr s2">🌱 Section 2 — Agronomic Data</div>',
                unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.radio("Contest Division *",
                 ["Division I - Tillage (conv./min.)", "Division II - No-Tillage"],
                 horizontal=True, key="division")
        st.selectbox("Previous Crop", ["Corn", "Soybeans", "Other"], key="previous_crop")
    with c2:
        st.date_input("Planting Date *", key="planting_date")
        st.date_input("Harvest Date *", key="harvest_date")
        st.text_input("Wheat Variety *", key="wheat_variety")
    with c3:
        st.text_input("Row Width (inches)", key="row_width")
        st.text_input("Seeding Rate (seeds/A or lb/A)", key="seeding_rate")

    # ════════════════════════════════════════════════════
    # SECTION 3 — FERTILIZER
    # ════════════════════════════════════════════════════
    st.markdown('<div class="sec-hdr s3">🧪 Section 3 — Fertilizer</div>',
                unsafe_allow_html=True)
    st.caption("Fall Fertilizer (lb/acre)")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.text_input("N (lb/A)", key="fall_n")
    with c2: st.text_input("P2O5 (lb/A)", key="fall_p")
    with c3: st.text_input("K2O (lb/A)", key="fall_k")
    with c4: st.text_input("Other fertilizers (type, rate, timing)", key="fall_other")
    st.caption("Winter/Spring Nitrogen")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.date_input("Application 1 Date", key="ws_n1_date")
    with c2: st.text_input("N lb/A (App 1)", key="ws_n1_rate")
    with c3: st.date_input("Application 2 Date", key="ws_n2_date")
    with c4: st.text_input("N lb/A (App 2)", key="ws_n2_rate")
    st.caption("Manure")
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.selectbox("Manure (last 18 months)?", ["No", "Yes"], key="manure_used")
    manure_on = st.session_state.get("manure_used", "No") == "Yes"
    with c2: st.text_input("Type", key="manure_type", disabled=not manure_on)
    with c3: st.text_input("Tons/A", key="manure_tons", disabled=not manure_on)
    with c4: st.date_input("Manure Date", key="manure_date", disabled=not manure_on)

    # ════════════════════════════════════════════════════
    # SECTION 4 — PEST MANAGEMENT
    # ════════════════════════════════════════════════════
    st.markdown('<div class="sec-hdr s4">🛡️ Section 4 — Pest Management & Other Inputs</div>',
                unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.text_area("Growth Regulator(s) — product & timing", height=68, key="growth_reg")
        st.text_area("Fall Pest Products (herbicides, fungicides, insecticides)", height=68, key="fall_pest")
        st.text_area("Biologicals / Other", height=68, key="biologicals")
    with c2:
        st.text_area("Spring Pest Products", height=68, key="spring_pest")
        st.text_area("Heading/Flowering Pest Products", height=68, key="heading_pest")
        st.text_input("Tillage Equipment / Method Used", key="tillage_used")

    # ════════════════════════════════════════════════════
    # SECTION 5 — HARVEST AREA
    # ════════════════════════════════════════════════════
    st.markdown('<div class="sec-hdr s5">📐 Section 5 — Harvest Area Measurement</div>',
                unsafe_allow_html=True)
    st.info("Minimum harvest area: **1.50 acres** (65,340 sq ft). "
            "Measure each side with tape or measuring wheel.")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.number_input("Length (ft) *", min_value=0.0, step=1.0, format="%.1f",
                        key="h_length", on_change=_recompute)
    with c2:
        st.number_input("Width (ft) *", min_value=0.0, step=1.0, format="%.1f",
                        key="h_width", on_change=_recompute)
    h_area_ft2 = st.session_state["h_area_ft2"]
    h_acres    = st.session_state["h_acres"]
    area_ok    = h_acres >= 1.50
    with c3:
        st.markdown(f"""<div class="metric-box">
          <div class="metric-label">Area (ft²)</div>
          <div class="metric-value">{h_area_ft2:,.1f}</div></div>""",
                    unsafe_allow_html=True)
    with c4:
        col_cls = "metric-ok" if area_ok else "metric-warn"
        badge   = "OK" if area_ok else "Below 1.50 ac min"
        st.markdown(f"""<div class="metric-box">
          <div class="metric-label">Acres &mdash; {badge}</div>
          <div class="metric-value {col_cls}">{h_acres:.2f}</div></div>""",
                    unsafe_allow_html=True)

    # ════════════════════════════════════════════════════
    # SECTION 6 — GRAIN CHARACTERISTICS
    # ════════════════════════════════════════════════════
    st.markdown('<div class="sec-hdr s6">🌡️ Section 6 — Grain Characteristics</div>',
                unsafe_allow_html=True)
    readings_now = st.session_state.get("moisture_readings", [])
    n_done       = len(readings_now)
    slots_left   = 3 - n_done
    warn_empty   = st.session_state.pop("_moisture_warn", False)
    st.caption("**Grain Moisture:** Certified elevator tester → 1 reading. "
               "Handheld meter → 3 readings added one at a time, average is auto-calculated.")
    mc1, mc2, mc3 = st.columns([2, 1, 3])
    with mc1:
        gen   = st.session_state.get("_moisture_gen", 0)
        label = (f"Moisture Reading {n_done + 1} of 3 (%)"
                 if slots_left > 0 else "All 3 readings recorded")
        st.number_input(label, min_value=0.0, max_value=40.0, step=0.1, format="%.1f",
                        value=0.0, key=f"moisture_input_{gen}", disabled=(slots_left == 0),
                        help="Enter % from moisture meter, then press Add Reading.")
        st.session_state["moisture_input"] = st.session_state.get(f"moisture_input_{gen}", 0.0)
        if warn_empty:
            st.warning("Enter a value > 0 before adding.")
    with mc2:
        st.markdown("<div style='margin-top:28px'></div>", unsafe_allow_html=True)
        st.button("➕ Add Reading", disabled=(slots_left == 0),
                  use_container_width=True, on_click=_add_reading_cb)
        if readings_now:
            st.button("🗑️ Clear", use_container_width=True,
                      on_click=_clear_readings_cb)
    with mc3:
        gm_avg = st.session_state["gm_avg"]
        if readings_now:
            badges = " ".join(
                f'<span class="moisture-badge">#{i+1}: <b>{r}%</b></span>'
                for i, r in enumerate(readings_now))
            st.markdown(f"""<div class="metric-box">
              <div class="metric-label">Recorded readings &nbsp; {badges}</div>
              <div class="metric-value metric-ok">Avg: {gm_avg:.1f}%</div></div>""",
                        unsafe_allow_html=True)
        else:
            st.markdown("""<div class="metric-box">
              <div class="metric-label">Recorded readings</div>
              <div class="metric-value" style="color:#adb5bd;font-size:1rem">
                None yet &mdash; add up to 3</div></div>""", unsafe_allow_html=True)
    st.markdown("")
    st.number_input("Test Weight (lb/bu)", min_value=0.0, max_value=70.0,
                    step=0.1, format="%.1f", key="test_weight")

    # ════════════════════════════════════════════════════
    # SECTION 7 — YIELD CALCULATION
    # ════════════════════════════════════════════════════
    st.markdown('<div class="sec-hdr s7">📊 Section 7 — Yield Calculation</div>',
                unsafe_allow_html=True)
    st.caption("Formula: lbs x [(100 - %moisture) / 86.5] / 60 lb/bu / acres")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.number_input("Grain Weight from Scale (lbs) *", min_value=0.0,
                        step=10.0, format="%.1f", key="grain_weight", on_change=_recompute)
    with c2:
        official_yield = st.session_state["official_yield"]
        yld_cls = "metric-ok" if official_yield > 0 else ""
        st.markdown(f"""<div class="metric-box">
          <div class="metric-label">Official Yield (Bu/Acre)</div>
          <div class="metric-value {yld_cls}">{official_yield:.2f}</div></div>""",
                    unsafe_allow_html=True)
    with c3:
        gw = float(st.session_state.get("grain_weight", 0.0) or 0.0)
        st.markdown(f"""<div class="metric-box" style="font-size:0.82rem;color:#495057;line-height:1.7">
          <b>Step-by-step:</b><br>
          {gw:.0f} x [(100 - {gm_avg:.1f}) / 86.5]<br>
          &divide; 60 &divide; {h_acres:.2f} ac<br>
          = <b>{official_yield:.2f} bu/ac</b></div>""", unsafe_allow_html=True)

    # ════════════════════════════════════════════════════
    # SECTION 8 — NOTES
    # ════════════════════════════════════════════════════
    st.markdown('<div class="sec-hdr s8">📝 Section 8 — Agent Notes</div>',
                unsafe_allow_html=True)
    st.text_area("Additional notes, observations, or issues", height=80, key="agent_notes")

    # ════════════════════════════════════════════════════
    # AGENT CERTIFICATION + GATED SUBMIT
    # ════════════════════════════════════════════════════
    st.divider()
    st.markdown("""
    <div class="agreement-box">
    <b>📋 Agent Certification — Required Before Submission</b><br>
    By checking the box below, you certify that:
    <ul style="margin:8px 0 4px 20px;">
      <li>All information entered is <b>accurate and complete</b> to the best of your knowledge.</li>
      <li>The harvest area was <b>physically measured</b> and meets the 1.50-acre minimum.</li>
      <li>Grain moisture and weight were recorded using <b>certified or approved equipment</b>.</li>
      <li>You witnessed or verified the harvest as the <b>supervising county agent</b>.</li>
      <li>You understand this entry will be <b>submitted officially</b> to the UK Extension state office.</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)

    # Generation-counter pattern — bumping gen makes Streamlit create a fresh
    # unchecked widget, avoiding the "cannot modify after instantiation" error.
    agree_gen = st.session_state.get("_agree_gen", 0)
    st.checkbox("I certify that all information above is accurate and complete.",
                key=f"agreement_checked_{agree_gen}", value=False)
    agreement = st.session_state.get(f"agreement_checked_{agree_gen}", False)

    # ── Validation ──
    errors = []
    if st.session_state.get("county", "— Select —") == "— Select —":
        errors.append("County not selected")
    if not st.session_state.get("producer_name", "").strip():
        errors.append("Producer name missing")
    if not st.session_state.get("supervisor_name", "").strip():
        errors.append("Supervisor name missing")
    if not st.session_state.get("wheat_variety", "").strip():
        errors.append("Wheat variety missing")
    if st.session_state["h_acres"] < 1.50:
        errors.append(f"Harvest area {st.session_state['h_acres']:.2f} ac < 1.50 ac minimum")
    if float(st.session_state.get("grain_weight", 0.0) or 0.0) <= 0:
        errors.append("Grain weight not entered")
    if not st.session_state.get("moisture_readings"):
        errors.append("No moisture reading recorded — press Add Reading at least once")

    if errors:
        with st.expander("⚠️ Required fields not complete", expanded=True):
            for e in errors:
                st.markdown(f"- 🔴 {e}")

    form_ready = agreement and len(errors) == 0
    if not form_ready:
        st.caption("🔒 Check the certification box above to unlock." if not agreement
                   else "🔒 Fix the highlighted fields above to unlock.")

    submit_clicked = st.button("💾  Save Entry & Send to State Office",
                               disabled=not form_ready, type="primary",
                               use_container_width=True)

    # ════════════════════════════════════════════════════
    # ON SUBMIT
    # ════════════════════════════════════════════════════
    if submit_clicked and form_ready:
        r   = st.session_state.get("moisture_readings", [])
        gm1 = r[0] if len(r) > 0 else ""
        gm2 = r[1] if len(r) > 1 else ""
        gm3 = r[2] if len(r) > 2 else ""

        data = {
            "County":                    st.session_state["county"],
            "Producer_Name":             st.session_state["producer_name"],
            "Producer_Address":          st.session_state.get("producer_address", ""),
            "Producer_Town":             st.session_state.get("producer_town", ""),
            "Producer_Zip":              st.session_state.get("producer_zip", ""),
            "Producer_Phone":            st.session_state.get("producer_phone", ""),
            "Producer_Mobile":           st.session_state.get("producer_mobile", ""),
            "Profession":                st.session_state.get("profession", ""),
            "Harvest_Date":              str(st.session_state.get("harvest_date", "")),
            "Supervisor_Name":           st.session_state["supervisor_name"],
            "Supervisor_Signature_Date": str(st.session_state.get("supervisor_sig_date", "")),
            "Division":                  st.session_state.get("division", ""),
            "Previous_Crop":             st.session_state.get("previous_crop", ""),
            "Planting_Date":             str(st.session_state.get("planting_date", "")),
            "Wheat_Variety":             st.session_state["wheat_variety"],
            "Row_Width_inches":          st.session_state.get("row_width", ""),
            "Seeding_Rate":              st.session_state.get("seeding_rate", ""),
            "Fall_N_lbA":                st.session_state.get("fall_n", "0"),
            "Fall_P2O5_lbA":             st.session_state.get("fall_p", "0"),
            "Fall_K2O_lbA":              st.session_state.get("fall_k", "0"),
            "Fall_Other_Fertilizer":     st.session_state.get("fall_other", ""),
            "Winter_Spring_N1_Date":     str(st.session_state.get("ws_n1_date", "")),
            "Winter_Spring_N1_lbA":      st.session_state.get("ws_n1_rate", "0"),
            "Winter_Spring_N2_Date":     str(st.session_state.get("ws_n2_date", "")),
            "Winter_Spring_N2_lbA":      st.session_state.get("ws_n2_rate", "0"),
            "Manure_Used":               st.session_state.get("manure_used", "No"),
            "Manure_Type":               st.session_state.get("manure_type", "") if manure_on else "",
            "Manure_TonsA":              st.session_state.get("manure_tons", "") if manure_on else "",
            "Manure_Date":               str(st.session_state.get("manure_date", "")) if manure_on else "",
            "Growth_Regulator":          st.session_state.get("growth_reg", ""),
            "Fall_Pest_Products":        st.session_state.get("fall_pest", ""),
            "Spring_Pest_Products":      st.session_state.get("spring_pest", ""),
            "Heading_Flowering_Pest":    st.session_state.get("heading_pest", ""),
            "Biologicals_Other":         st.session_state.get("biologicals", ""),
            "Tillage_Used":              st.session_state.get("tillage_used", ""),
            "Harvest_Length_ft":         st.session_state["h_length"],
            "Harvest_Width_ft":          st.session_state["h_width"],
            "Harvest_Area_ft2":          st.session_state["h_area_ft2"],
            "Harvest_Acres":             st.session_state["h_acres"],
            "Grain_Moisture_1":          gm1, "Grain_Moisture_2": gm2, "Grain_Moisture_3": gm3,
            "Grain_Moisture_Avg":        st.session_state["gm_avg"],
            "Test_Weight_lbbu":          st.session_state.get("test_weight", 60.0),
            "Grain_Weight_lbs":          st.session_state["grain_weight"],
            "Official_Yield_BuAcre":     st.session_state["official_yield"],
            "Agent_Notes":               st.session_state.get("agent_notes", ""),
        }

        try:
            entry_id = build_excel_with_entry(data, excel_path)
            area = COUNTY_AREA.get(data["County"], 4)

            st.markdown(f"""<div class="result-ok">
              <h4>Entry #{entry_id} Saved Successfully!</h4>
              <b>Producer:</b> {data['Producer_Name']} &nbsp;|&nbsp;
              <b>County:</b> {data['County']} (Area {area}) &nbsp;|&nbsp;
              <b>Division:</b> {data['Division'].split('-')[0].strip()}<br>
              <b>Official Yield:</b> {data['Official_Yield_BuAcre']:.2f} bu/acre &nbsp;|&nbsp;
              <b>Harvest Area:</b> {data['Harvest_Acres']:.2f} acres &nbsp;|&nbsp;
              <b>Grain Moisture:</b> {data['Grain_Moisture_Avg']:.1f}%</div>""",
                        unsafe_allow_html=True)

            with open(excel_path, "rb") as f:
                st.download_button("Download Updated Master Excel", data=f,
                                   file_name=excel_path,
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            if send_flag:
                body = (
                    f"Dear {RECIPIENT_NAME},\n\nA new wheat yield contest entry has been submitted:\n\n"
                    f"  Entry ID:       #{entry_id}\n  Producer:       {data['Producer_Name']}\n"
                    f"  County:         {data['County']} (Area {area})\n  Division:       {data['Division']}\n"
                    f"  Harvest Date:   {data['Harvest_Date']}\n  Supervisor:     {data['Supervisor_Name']}\n"
                    f"  Official Yield: {data['Official_Yield_BuAcre']:.2f} bu/acre\n"
                    f"  Harvest Area:   {data['Harvest_Acres']:.2f} acres\n"
                    f"  Grain Moisture: {data['Grain_Moisture_Avg']:.1f}% "
                    f"(readings: {', '.join(str(x)+'%' for x in r)})\n\n"
                    f"Updated master Excel file attached.\n\nSubmitted via UK Extension Digital Contest Form.\n"
                )
                subject = f"KY Wheat Contest Entry #{entry_id} — {data['County']} County — {data['Producer_Name']}"
                cc_addr = effective_sender if cc_self else None
                try:
                    if st.session_state["auth_mode"] == "oauth":
                        tok = st.session_state.get("oauth_access_token")
                        if not tok:
                            st.warning("Not signed in via Microsoft. Sign in using the sidebar, then resubmit.")
                        else:
                            oauth_send_email(tok, effective_sender, RECIPIENT_EMAIL,
                                             subject, body, excel_path, cc=cc_addr)
                            st.success(f"Email sent via Microsoft to {RECIPIENT_EMAIL}" +
                                       (" + CC'd to you." if cc_self else "."))
                    else:
                        if sender_email and sender_pass:
                            smtp_send_email(smtp_server, smtp_port, sender_email, sender_pass,
                                            RECIPIENT_EMAIL, subject, body, excel_path, cc=cc_addr)
                            st.success(f"Email sent to {RECIPIENT_EMAIL}" +
                                       (" + CC'd to you." if cc_self else "."))
                        else:
                            st.info("Enter credentials in the sidebar to enable auto-email.")
                except Exception as email_err:
                    err_str = str(email_err)
                    if "535" in err_str or "Authentication unsuccessful" in err_str:
                        st.warning(
                            "**Email blocked — MFA is required on your UK account.**\n\n"
                            "Switch to **Microsoft Sign-In (OAuth2)** in the sidebar. "
                            "Sign in once per session and email will work automatically."
                        )
                    else:
                        st.warning(f"Entry saved but email failed: {email_err}\n\n"
                                   "Download the Excel above and email it manually.")

            # Reset: bump agree gen so checkbox re-renders unchecked
            st.session_state["_agree_gen"] = st.session_state.get("_agree_gen", 0) + 1

        except Exception as ex:
            st.error(f"Error saving entry: {ex}")

    # ════════════════════════════════════════════════════
    # ENTRIES TABLE
    # ════════════════════════════════════════════════════
    st.divider()
    st.subheader("📋 Current Season Entries")
    if Path(excel_path).exists():
        try:
            df = pd.read_excel(excel_path, header=2)
            show = ["Entry_ID","Submission_Date","County","Area","Producer_Name",
                    "Division","Wheat_Variety","Harvest_Acres","Official_Yield_BuAcre","Supervisor_Name"]
            cols = [c for c in show if c in df.columns]
            st.dataframe(df[cols].sort_values("Entry_ID", ascending=False),
                         use_container_width=True, hide_index=True)
            st.caption(f"Total entries: **{len(df)}**")
        except Exception as e:
            st.info(f"No entries yet: {e}")
    else:
        st.info("No entries saved yet. Submit the first entry above.")


if __name__ == "__main__":
    main()
