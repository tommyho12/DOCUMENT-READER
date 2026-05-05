import streamlit as st
import anthropic
import base64
import numpy as np
from PIL import Image, ImageEnhance, ImageFilter
import io
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable

st.set_page_config(page_title="DocuScan Pro", page_icon="📋", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #0d1117 !important; color: #e6edf3 !important; }
section[data-testid="stSidebar"] { background: #161b22 !important; border-right: 1px solid #21262d !important; }
section[data-testid="stSidebar"] * { color: #e6edf3 !important; }
.main .block-container { background: #0d1117 !important; padding-top: 1rem !important; }
input, textarea, select { background: #161b22 !important; border: 1px solid #30363d !important; color: #e6edf3 !important; border-radius: 8px !important; }
input:focus, textarea:focus { border-color: #2ea043 !important; box-shadow: 0 0 0 3px rgba(46,160,67,0.15) !important; }
.stButton > button { background: #238636 !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 600 !important; padding: 10px 20px !important; transition: all 0.2s !important; }
.stButton > button:hover { background: #2ea043 !important; transform: translateY(-1px); box-shadow: 0 4px 15px rgba(46,160,67,0.3) !important; }
.stButton > button[kind="secondary"] { background: #21262d !important; border: 1px solid #30363d !important; }
.stDownloadButton > button { background: #1f6feb !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 500 !important; }
.stTabs [data-baseweb="tab-list"] { background: #161b22 !important; border-radius: 10px !important; padding: 4px !important; gap: 4px !important; border: 1px solid #21262d !important; }
.stTabs [data-baseweb="tab"] { background: transparent !important; color: #8b949e !important; border-radius: 8px !important; font-weight: 500 !important; }
.stTabs [aria-selected="true"] { background: #238636 !important; color: white !important; }
.streamlit-expanderHeader { background: #161b22 !important; border: 1px solid #21262d !important; border-radius: 8px !important; color: #e6edf3 !important; }
.streamlit-expanderContent { background: #0d1117 !important; border: 1px solid #21262d !important; border-top: none !important; }
.stSelectbox > div > div { background: #161b22 !important; border: 1px solid #30363d !important; color: #e6edf3 !important; border-radius: 8px !important; }
[data-testid="metric-container"] { background: #161b22 !important; border: 1px solid #21262d !important; border-radius: 10px !important; padding: 12px !important; }
[data-testid="metric-container"] label { color: #8b949e !important; }
[data-testid="metric-container"] [data-testid="stMetricValue"] { color: #2ea043 !important; font-weight: 700 !important; }
.stSuccess { background: #0d2818 !important; border: 1px solid #238636 !important; color: #2ea043 !important; border-radius: 8px !important; }
.stError { background: #2d1117 !important; border: 1px solid #f85149 !important; color: #f85149 !important; border-radius: 8px !important; }
.stWarning { background: #271d0b !important; border: 1px solid #d29922 !important; color: #d29922 !important; border-radius: 8px !important; }
.stInfo { background: #0d1b2e !important; border: 1px solid #1f6feb !important; color: #388bfd !important; border-radius: 8px !important; }
[data-testid="stFileUploader"] { background: #161b22 !important; border: 2px dashed #30363d !important; border-radius: 12px !important; }
.ds-header { background: linear-gradient(135deg, #0d2818 0%, #1a3a2a 50%, #0d1117 100%); border: 1px solid #238636; padding: 28px 32px; border-radius: 16px; margin-bottom: 28px; display: flex; align-items: center; justify-content: space-between; box-shadow: 0 0 40px rgba(46,160,67,0.15); }
.ds-logo { font-size: 28px; font-weight: 800; letter-spacing: -1px; color: #e6edf3; }
.ds-logo span { color: #2ea043; }
.ds-logo sub { font-size: 13px; font-weight: 400; color: #8b949e; letter-spacing: 0; margin-left: 4px; }
.ds-tagline { font-size: 13px; color: #8b949e; margin-top: 4px; }
.ds-badge { background: linear-gradient(135deg, #238636, #2ea043); color: white; font-size: 11px; font-weight: 600; padding: 5px 14px; border-radius: 20px; letter-spacing: 0.5px; box-shadow: 0 2px 10px rgba(46,160,67,0.4); }
.quality-fail { background: #1a0a00; border: 2px solid #d29922; border-radius: 12px; padding: 20px 24px; margin: 12px 0; }
.quality-fail h3 { color: #d29922; font-size: 16px; margin-bottom: 8px; }
.quality-fail p { color: #c9a227; font-size: 13px; line-height: 1.7; margin: 0; }
.quality-ok { background: #0d2818; border: 1px solid #238636; border-radius: 10px; padding: 10px 16px; margin: 8px 0; font-size: 13px; color: #7ee787; }
.how-to-box { background: #0d1b2e; border: 1px solid #1f6feb; border-radius: 12px; padding: 20px 24px; margin: 12px 0; }
.how-to-box h4 { color: #388bfd; font-size: 14px; margin-bottom: 10px; }
.how-to-box p { color: #8b949e; font-size: 13px; line-height: 1.8; margin: 0; }
.how-to-box b { color: #e6edf3; }
.result-box { background: #161b22; border: 1px solid #21262d; border-left: 3px solid #2ea043; padding: 20px; border-radius: 10px; font-size: 13px; line-height: 1.9; white-space: pre-wrap; font-family: 'Courier New', monospace; color: #e6edf3; max-height: 420px; overflow-y: auto; }
.badge-high { background:#0d2818; color:#2ea043; border:1px solid #238636; padding:4px 12px; border-radius:20px; font-size:12px; font-weight:600; }
.badge-med  { background:#271d0b; color:#d29922; border:1px solid #d29922; padding:4px 12px; border-radius:20px; font-size:12px; font-weight:600; }
.badge-low  { background:#2d1117; color:#f85149; border:1px solid #f85149; padding:4px 12px; border-radius:20px; font-size:12px; font-weight:600; }
.section-title { font-size: 15px; font-weight: 600; color: #e6edf3; margin-bottom: 12px; }
.empty-state { text-align:center; padding:50px; color:#30363d; font-size:14px; }
</style>
""", unsafe_allow_html=True)

if "records" not in st.session_state:
    st.session_state.records = []

MIN_SHORT_SIDE = 600
MIN_PIXELS = 500_000

def check_quality(img):
    w, h = img.size
    issues = []
    if min(w,h) < MIN_SHORT_SIDE:
        issues.append(f"Too small: {w}×{h}px. Need at least {MIN_SHORT_SIDE}px on shortest side.")
    if w*h < MIN_PIXELS:
        issues.append(f"Too few pixels: {w*h:,}. Need at least {MIN_PIXELS:,}.")
    return issues

def upscale(img, target=1800):
    w, h = img.size
    m = max(w, h)
    if m < target:
        scale = (target / m) * (2 if m < 300 else 1)
        img = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)
    return img

def auto_levels(img, lo=2, hi=98):
    arr = np.array(img.convert("RGB"), dtype=np.float32)
    for c in range(3):
        ch = arr[:,:,c]
        p_lo, p_hi = np.percentile(ch, lo), np.percentile(ch, hi)
        if p_hi > p_lo:
            arr[:,:,c] = np.clip((ch - p_lo) / (p_hi - p_lo) * 255, 0, 255)
    return Image.fromarray(arr.astype(np.uint8))

def enhance_v1(img):
    """Auto levels + moderate sharpen — best for dark/low contrast docs"""
    img = auto_levels(img, 2, 98)
    img = ImageEnhance.Contrast(img).enhance(2.5)
    img = ImageEnhance.Brightness(img).enhance(1.3)
    for _ in range(3):
        img = ImageEnhance.Sharpness(img).enhance(5.0)
        img = img.filter(ImageFilter.SHARPEN)
        img = img.filter(ImageFilter.EDGE_ENHANCE_MORE)
    img = img.filter(ImageFilter.MedianFilter(3))
    return img

def enhance_v2(img):
    """Extreme contrast + sharpness — best for blurry text"""
    img = auto_levels(img, 1, 99)
    img = ImageEnhance.Contrast(img).enhance(4.0)
    img = ImageEnhance.Brightness(img).enhance(1.5)
    img = ImageEnhance.Sharpness(img).enhance(8.0)
    img = img.filter(ImageFilter.SHARPEN)
    img = img.filter(ImageFilter.EDGE_ENHANCE_MORE)
    img = ImageEnhance.Sharpness(img).enhance(6.0)
    return img

def enhance_v3(img):
    """Adaptive B&W threshold — best for printed forms"""
    arr = np.array(img.convert("L"), dtype=np.float32)
    blur_pil = Image.fromarray(arr.astype(np.uint8)).filter(ImageFilter.GaussianBlur(radius=15))
    blurred = np.array(blur_pil, dtype=np.float32)
    thresh = np.where(arr > blurred - 10, 255, 0).astype(np.uint8)
    return Image.fromarray(thresh).convert("RGB")

def generate_versions(pil_img):
    img = upscale(pil_img)
    return [
        ("Auto Levels", enhance_v1(img)),
        ("High Contrast", enhance_v2(img)),
        ("B&W Text Mode", enhance_v3(img)),
    ]

def img_to_b64(img):
    buf = io.BytesIO(); img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode()

def img_to_bytes(img):
    buf = io.BytesIO(); img.save(buf, format="PNG"); return buf.getvalue()

def build_prompt(doc_type, client_name, notes):
    hints = {
        "Insurance Card": "Member Name, Member ID, Group Number, Plan Name, Insurance Company, Effective Date, Copays, Phone numbers, Payer ID.",
        "Driver's License / ID": "Full Name, DOB, Address, License/ID Number, Expiration Date, State.",
        "Tax Form (W-2 / 1099 / 1040)": "Form type, Tax Year, Taxpayer Name, SSN last 4, Employer/Payer, EIN, all box numbers and dollar amounts.",
        "Medical Record": "Patient Name, DOB, Date of Service, Provider, Diagnoses, Medications, Instructions.",
        "Pay Stub": "Employee Name, Pay Period, Gross Pay, Net Pay, Deductions, YTD, Employer name.",
        "Bank Statement": "Account holder, Account number last 4, Bank name, Statement period, Balances.",
        "Letter / Notice": "Sender, Date, Reference numbers, Recipient, Subject, Key dates, Action items, Contact info.",
        "Social Security Card": "Full name, Social Security Number.",
        "Medicare / Medicaid Card": "Name, Medicare/Medicaid number, effective dates, coverage type.",
    }
    hint = hints.get(doc_type, "Extract all visible text and data fields.")
    return f"""You are DocuScan Pro — an elite document reading AI for insurance and tax offices.

You are receiving 3 enhanced versions of the same document image. Each uses a different enhancement technique. Cross-reference ALL 3 to extract the most complete information possible.

RULES:
1. NEVER say you cannot read the document — always extract something
2. Mark uncertain text with [?] but always give your best attempt
3. Cross-reference all 3 versions — text in one may not be in another
4. Read ANY language without translating
5. Look for patterns: $ = amounts, / between numbers = dates, letter+numbers = IDs

CLIENT: {client_name}
DOCUMENT TYPE: {doc_type}
KEY FIELDS TO FIND: {hint}
NOTES: {notes if notes else 'None'}

FORMAT YOUR RESPONSE:

DOCUMENT TYPE: [identified type]

CLIENT INFORMATION:
[Names, DOB, addresses, phones]

KEY FIELDS:
[All labeled fields with values]

FULL EXTRACTED TEXT:
[Every visible line top to bottom]

CONFIDENCE: [High / Medium / Low] — [one line reason]"""

def export_excel(records):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "DocuScan Records"
    hfill = PatternFill("solid", fgColor="238636")
    hfont = Font(color="FFFFFF", bold=True, size=11)
    headers = ["#","Client Name","Document Type","Date Processed","Confidence","Extracted Text"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    for i, w in enumerate([5,25,22,22,12,90], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    alt = PatternFill("solid", fgColor="0D2818")
    for i, rec in enumerate(records, 1):
        row = i+1
        vals = [i, rec["client"], rec["doc_type"], rec["timestamp"], rec.get("confidence","—"),
                rec["extracted_text"][:800]+("…" if len(rec["extracted_text"])>800 else "")]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=val)
            if i%2==0: cell.fill = alt
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def export_pdf(records):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
        leftMargin=0.75*inch, rightMargin=0.75*inch,
        topMargin=0.75*inch, bottomMargin=0.75*inch)
    title_s = ParagraphStyle("t", fontSize=22, fontName="Helvetica-Bold", textColor=colors.HexColor("#238636"), spaceAfter=4)
    sub_s   = ParagraphStyle("s", fontSize=10, textColor=colors.HexColor("#8b949e"), spaceAfter=20)
    lbl_s   = ParagraphStyle("l", fontSize=11, fontName="Helvetica-Bold", textColor=colors.HexColor("#2ea043"), spaceBefore=18, spaceAfter=4)
    meta_s  = ParagraphStyle("m", fontSize=9, textColor=colors.HexColor("#8b949e"), spaceAfter=6)
    body_s  = ParagraphStyle("b", fontSize=9, fontName="Courier", leading=14, textColor=colors.HexColor("#333333"))
    story = [
        Paragraph("DocuScan Pro — Extracted Records", title_s),
        Paragraph(f"Generated: {datetime.datetime.now().strftime('%B %d, %Y %I:%M %p')}  |  Records: {len(records)}", sub_s),
        HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#238636")),
        Spacer(1, 14),
    ]
    for i, rec in enumerate(records, 1):
        story.append(Paragraph(f"#{i}  {rec['client']}", lbl_s))
        story.append(Paragraph(f"Type: {rec['doc_type']}  |  {rec['timestamp']}  |  Confidence: {rec.get('confidence','—')}", meta_s))
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#dddddd")))
        story.append(Spacer(1,6))
        safe = rec["extracted_text"].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        story.append(Paragraph(safe.replace("\n","<br/>"), body_s))
        story.append(Spacer(1,18))
    doc.build(story); return buf.getvalue()

# ── Sidebar ───────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    api_key = st.text_input("Anthropic API Key", type="password", placeholder="sk-ant-api03-...")
    if api_key:
        if not api_key.startswith("sk-ant-"):
            st.error("Must start with sk-ant-")
        else:
            st.success("✓ Key ready")
    st.caption("🔒 Never stored. Session only.")
    st.markdown("[Get a free key →](https://console.anthropic.com)")
    st.divider()
    st.markdown("### 📊 Session Stats")
    st.metric("Documents Processed", len(st.session_state.records))
    if st.session_state.records:
        types = {}
        for r in st.session_state.records:
            types[r["doc_type"]] = types.get(r["doc_type"],0)+1
        for t,c in sorted(types.items(), key=lambda x: -x[1]):
            st.caption(f"• {t}: {c}")
    st.divider()
    st.markdown("**DocuScan Pro** v2.1")
    st.caption("Built for Vietrust Insurance & Tax")

# ── Header ────────────────────────────────────────────────
st.markdown("""
<div class="ds-header">
  <div>
    <div class="ds-logo">Docu<span>Scan</span> Pro<sub>by Vietrust</sub></div>
    <div class="ds-tagline">AI document reading · Any language · Clear image required for best results</div>
  </div>
  <div class="ds-badge">⚡ v2.1 POWERED BY CLAUDE AI</div>
</div>
""", unsafe_allow_html=True)

tab_scan, tab_records, tab_export, tab_help = st.tabs(["📷  Scan Document","📁  Records","📤  Export","❓  How to Send Documents"])

# ── TAB 1: SCAN ───────────────────────────────────────────
with tab_scan:
    col_left, col_right = st.columns([1,2])
    with col_left:
        st.markdown('<div class="section-title">Client Info</div>', unsafe_allow_html=True)
        client_name = st.text_input("Client Name *", placeholder="e.g. Nguyen, John")
        doc_type = st.selectbox("Document Type", [
            "Auto-Detect","Insurance Card","Driver's License / ID",
            "Tax Form (W-2 / 1099 / 1040)","Medical Record","Pay Stub",
            "Bank Statement","Letter / Notice","Social Security Card",
            "Medicare / Medicaid Card","Other"
        ])
        notes = st.text_area("Notes (optional)", placeholder="Any context...", height=80)

    with col_right:
        st.markdown('<div class="section-title">Upload Document</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Upload document image",
            type=["jpg","jpeg","png","webp","bmp","tiff"], label_visibility="collapsed")

        if uploaded:
            img = Image.open(uploaded)
            w, h = img.size
            issues = check_quality(img)
            st.image(img, caption=f"{uploaded.name} ({w}×{h}px)", use_container_width=True)

            if issues:
                st.markdown(f"""
                <div class="quality-fail">
                    <h3>⚠️ Image Quality Too Low to Read</h3>
                    <p>This image is <b>{w}×{h} pixels</b> — too small for the AI to read accurately.<br><br>
                    <b>Minimum required:</b> 600px on the shortest side ({min(w,h)}px detected).<br><br>
                    Please ask your client to resend using the instructions in the <b>❓ How to Send Documents</b> tab.</p>
                </div>
                """, unsafe_allow_html=True)
                st.markdown("**📋 Copy & send this to your client:**")
                st.code("""Hi! We received your document but the image quality is too low for our system to read it.

Please resend by:
• Photo in camera roll: Open the original photo and send it directly — do NOT screenshot it
• PDF/file: Attach the actual file from your downloads or email — do NOT photo your screen  
• Physical document: Place flat in good lighting and take a fresh photo

Thank you! — Vietrust""", language=None)
            else:
                st.markdown(f'<div class="quality-ok">✅ Image quality good — {w}×{h}px · Ready to scan</div>', unsafe_allow_html=True)

    st.divider()
    c1, c2 = st.columns([3,1])
    with c1:
        img_ok = uploaded and not check_quality(Image.open(uploaded))
        run_ai = st.button("⚡ Read & Extract with AI", type="primary", disabled=not img_ok, use_container_width=True)
    with c2:
        enh_only = st.button("✨ Preview Enhancement", disabled=not uploaded, use_container_width=True)

    if enh_only and uploaded:
        img = Image.open(uploaded)
        if check_quality(img):
            st.error("Image is too low quality. Please ask the client for a better photo.")
        else:
            with st.spinner("Generating enhanced previews..."):
                versions = generate_versions(img)
            cols = st.columns(len(versions)+1)
            with cols[0]:
                st.image(img, caption="Original", use_container_width=True)
            for i,(label,enh) in enumerate(versions):
                with cols[i+1]:
                    st.image(enh, caption=label, use_container_width=True)
                    st.download_button(f"⬇️ {label}", img_to_bytes(enh),
                        file_name=f"enhanced_{label.replace(' ','_')}.png",
                        mime="image/png", key=f"prev_{i}")

    if run_ai and uploaded:
        if not api_key:
            st.error("⚠️ Enter your Anthropic API key in the sidebar."); st.stop()
        if not client_name.strip():
            st.warning("⚠️ Enter the client name."); st.stop()

        img = Image.open(uploaded)
        with st.spinner("⚡ Enhancing + reading with AI..."):
            progress = st.progress(0, text="Enhancing image...")
            versions = generate_versions(img)
            progress.progress(40, text="Sending to Claude AI...")
            try:
                ai_client = anthropic.Anthropic(api_key=api_key)
                content = []
                for label, enh in versions:
                    content.append({"type":"image","source":{"type":"base64","media_type":"image/png","data":img_to_b64(enh)}})
                content.append({"type":"text","text":build_prompt(doc_type, client_name.strip(), notes)})
                progress.progress(60, text="AI reading document...")
                message = ai_client.messages.create(
                    model="claude-opus-4-6", max_tokens=2000,
                    messages=[{"role":"user","content":content}]
                )
                progress.progress(100, text="Done!")
                extracted = message.content[0].text
                confidence = "Medium"
                if "CONFIDENCE: High" in extracted: confidence = "High"
                elif "CONFIDENCE: Low" in extracted: confidence = "Low"

                record = {
                    "id": len(st.session_state.records)+1,
                    "client": client_name.strip(), "doc_type": doc_type,
                    "filename": uploaded.name,
                    "timestamp": datetime.datetime.now().strftime("%m/%d/%Y %I:%M %p"),
                    "extracted_text": extracted, "confidence": confidence, "notes": notes,
                    "enhanced_img": img_to_bytes(versions[0][1]),
                    "original_size": f"{img.width}×{img.height}px",
                    "enhanced_size": f"{versions[0][1].width}×{versions[0][1].height}px",
                }
                st.session_state.records.insert(0, record)
                progress.empty()

                badge = {"High":"badge-high","Medium":"badge-med","Low":"badge-low"}.get(confidence,"badge-med")
                st.markdown(f'✅ &nbsp;<span class="{badge}">Confidence: {confidence}</span>', unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                r1, r2 = st.columns(2)
                with r1:
                    t1,t2,t3,t4 = st.tabs(["Original","Auto Levels","High Contrast","B&W Text"])
                    with t1: st.image(img, use_container_width=True)
                    with t2: st.image(versions[0][1], use_container_width=True)
                    with t3: st.image(versions[1][1], use_container_width=True)
                    with t4: st.image(versions[2][1], use_container_width=True)
                    st.caption(f"{record['original_size']} → {record['enhanced_size']}")
                with r2:
                    st.markdown("**Extracted Text & Data**")
                    st.markdown(f'<div class="result-box">{extracted.replace("<","&lt;")}</div>', unsafe_allow_html=True)
                    st.download_button("⬇️ Download TXT", extracted,
                        file_name=f"{client_name}_{doc_type}.txt", mime="text/plain")
                st.success(f"✅ Saved to Records! Total: {len(st.session_state.records)}")

            except anthropic.AuthenticationError:
                st.error("❌ Invalid API key.")
            except anthropic.RateLimitError:
                st.error("❌ Rate limit hit. Wait 30 seconds.")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

# ── TAB 2: RECORDS ────────────────────────────────────────
with tab_records:
    st.markdown('<div class="section-title">📁 Processed Documents</div>', unsafe_allow_html=True)
    if not st.session_state.records:
        st.markdown('<div class="empty-state">📄<br><br>No documents scanned yet.</div>', unsafe_allow_html=True)
    else:
        search = st.text_input("🔍 Search", placeholder="Client name or document type...")
        filtered = [r for r in st.session_state.records if
                    not search or search.lower() in r["client"].lower() or search.lower() in r["doc_type"].lower()]
        st.caption(f"Showing {len(filtered)} of {len(st.session_state.records)} records")
        for rec in filtered:
            icon = {"High":"🟢","Medium":"🟡","Low":"🔴"}.get(rec.get("confidence","Medium"),"🟡")
            with st.expander(f"{icon}  #{rec['id']}  ·  {rec['client']}  ·  {rec['doc_type']}  ·  {rec['timestamp']}"):
                l, r = st.columns([1,2])
                with l:
                    if rec.get("enhanced_img"):
                        st.image(rec["enhanced_img"], use_container_width=True)
                    st.caption(f"File: {rec['filename']}")
                    st.caption(f"Size: {rec['original_size']} → {rec['enhanced_size']}")
                    st.caption(f"Confidence: {rec.get('confidence','—')}")
                with r:
                    st.markdown(f'<div class="result-box">{rec["extracted_text"].replace("<","&lt;")}</div>', unsafe_allow_html=True)
                    st.download_button("⬇️ TXT", rec["extracted_text"],
                        file_name=f"{rec['client']}_{rec['doc_type']}.txt",
                        mime="text/plain", key=f"r_{rec['id']}")

# ── TAB 3: EXPORT ─────────────────────────────────────────
with tab_export:
    st.markdown('<div class="section-title">📤 Export All Records</div>', unsafe_allow_html=True)
    if not st.session_state.records:
        st.markdown('<div class="empty-state">📄<br><br>No records to export yet.</div>', unsafe_allow_html=True)
    else:
        st.info(f"Ready to export **{len(st.session_state.records)} records**.")
        c1,c2,c3 = st.columns(3)
        with c1:
            st.markdown("### 📊 Excel")
            st.markdown("Formatted spreadsheet for data entry.")
            if st.button("Generate Excel", use_container_width=True):
                st.download_button("⬇️ Download Excel", export_excel(st.session_state.records),
                    file_name=f"DocuScan_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
        with c2:
            st.markdown("### 📄 PDF")
            st.markdown("Professional PDF for printing or sharing.")
            if st.button("Generate PDF", use_container_width=True):
                st.download_button("⬇️ Download PDF", export_pdf(st.session_state.records),
                    file_name=f"DocuScan_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf", use_container_width=True)
        with c3:
            st.markdown("### 📋 Text")
            st.markdown("Plain text for any other system.")
            if st.button("Generate TXT", use_container_width=True):
                lines = []
                for rec in st.session_state.records:
                    lines += [f"{'='*60}",f"Client: {rec['client']}",f"Type: {rec['doc_type']}",
                              f"Date: {rec['timestamp']}",f"Confidence: {rec.get('confidence','—')}",
                              "─"*40,rec["extracted_text"],""]
                st.download_button("⬇️ Download TXT","\n".join(lines),
                    file_name=f"DocuScan_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                    mime="text/plain", use_container_width=True)
        st.divider()
        if st.button("🗑️ Clear All Records", type="secondary"):
            st.session_state.records = []; st.rerun()

# ── TAB 4: HOW TO SEND ────────────────────────────────────
with tab_help:
    st.markdown('<div class="section-title">❓ How to Send Documents Correctly</div>', unsafe_allow_html=True)
    st.markdown("Share these instructions with your clients so their documents come through clearly.")
    st.divider()

    st.markdown("""
    <div class="how-to-box">
        <h4>📱 Sending a photo from your camera roll</h4>
        <p><b>iPhone:</b> Open the photo → tap Share (box with arrow) → choose Messages or Email → send as <b>full-size original</b>. Do NOT screenshot it first.<br><br>
        <b>Android:</b> Open the photo → tap Share → choose your app → select <b>Original Quality</b> before sending.</p>
    </div>
    <div class="how-to-box">
        <h4>📄 Sending a PDF or file</h4>
        <p>Go to your email, downloads, or files app → find the document → tap Share or Attach → send the <b>actual file</b>.<br><br>
        <b>Do NOT</b> take a screenshot of the PDF on your screen — this loses almost all quality.</p>
    </div>
    <div class="how-to-box">
        <h4>📷 Photographing a physical document</h4>
        <p>Place the document flat on a table in good lighting → open your camera → <b>tap on the document to focus</b> → take the photo and send it directly from your camera app.<br><br>
        Make sure the full document is in frame and the image is sharp before sending.</p>
    </div>
    """, unsafe_allow_html=True)

    st.divider()
    st.markdown("### 📋 Copy & send to your client:")
    st.code("""Hi! We received your document but the image quality is too low for our system to read accurately.

Please resend using one of these methods:

📱 PHOTO IN CAMERA ROLL:
Open your camera roll → find the original photo → send it directly. Do NOT screenshot it.

📄 PDF OR FILE:
Go to your downloads or email → find the file → attach and send it directly. Do NOT photo your screen.

📷 PHYSICAL DOCUMENT:
Place it flat on a table with good lighting → open your camera → tap the document to focus → take a clear photo and send it.

✅ KEY RULE: Always send the original file or photo — never a screenshot.

Thank you! — Vietrust Insurance & Tax""", language=None)
