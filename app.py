import streamlit as st
import anthropic
import base64
from PIL import Image, ImageEnhance, ImageFilter, ImageOps
import io
import numpy as np

st.set_page_config(page_title="Vietrust Document Reader", page_icon="📄", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 26px; font-weight: 700; color: #185FA5; margin-bottom: 2px; }
    .sub { font-size: 13px; color: #999; margin-bottom: 20px; }
    .result-box {
        background: #f8f9fa; border-left: 4px solid #185FA5;
        padding: 16px; border-radius: 6px; font-size: 14px;
        line-height: 1.9; white-space: pre-wrap; font-family: monospace;
    }
    .badge-ok { background:#EAF3DE; color:#3B6D11; padding:3px 12px; border-radius:20px; font-size:12px; font-weight:600; }
    .section-head { font-size:13px; font-weight:600; color:#444; margin-bottom:6px; text-transform:uppercase; letter-spacing:.04em; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📄 Vietrust Document Reader</div>', unsafe_allow_html=True)
st.markdown('<div class="sub">Upload blurry client photos — app cleans them up AND extracts all text automatically</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Settings")
    api_key = st.text_input("Anthropic API Key", type="password", placeholder="sk-ant-api03-...",
                            help="Get free at console.anthropic.com → API Keys")
    st.caption("🔒 Your key is never stored or shared.")
    st.markdown("**Get a free key:** [console.anthropic.com](https://console.anthropic.com)")
    st.divider()

    doc_type = st.selectbox("Document Type", [
        "Auto-Detect",
        "Insurance Card",
        "Driver's License / ID",
        "Tax Form (W-2, 1099, 1040)",
        "Medical Record",
        "Letter / Notice",
        "General Document",
    ])

    st.divider()
    st.markdown("**Enhancement Level**")
    enhance_level = st.radio("", ["Standard", "Aggressive (very blurry)", "Max (almost unreadable)"], index=1)

    st.divider()
    st.markdown("**Tips:**\n- Screenshots from phones work great\n- Upload multiple files at once\n- Use 'Max' for really bad photos\n- 'Enhance Only' button skips AI — just cleans image")

def enhance_image(img, level):
    img = img.convert("RGB")
    img_array = np.array(img, dtype=np.float32)

    # Auto white balance / levels
    for c in range(3):
        ch = img_array[:, :, c]
        p2, p98 = np.percentile(ch, 2), np.percentile(ch, 98)
        if p98 > p2:
            img_array[:, :, c] = np.clip((ch - p2) / (p98 - p2) * 255, 0, 255)

    img = Image.fromarray(img_array.astype(np.uint8))

    if level == "Standard":
        contrast, brightness, sharpness, passes = 1.8, 1.2, 2.5, 1
    elif level == "Aggressive (very blurry)":
        contrast, brightness, sharpness, passes = 2.5, 1.4, 4.0, 2
    else:
        contrast, brightness, sharpness, passes = 3.5, 1.6, 6.0, 3

    img = ImageEnhance.Contrast(img).enhance(contrast)
    img = ImageEnhance.Brightness(img).enhance(brightness)

    for _ in range(passes):
        img = ImageEnhance.Sharpness(img).enhance(sharpness)
        img = img.filter(ImageFilter.SHARPEN)
        img = img.filter(ImageFilter.EDGE_ENHANCE_MORE)

    img = img.filter(ImageFilter.MedianFilter(size=3))
    img = ImageEnhance.Sharpness(img).enhance(sharpness)
    return img

def image_to_base64(img):
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode()

def image_to_bytes(img):
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return buffer.getvalue()

prompts = {
    "Auto-Detect": """You are an expert document scanner and OCR system. This image may be blurry, dark, a phone screenshot, or low quality.

First, identify what type of document this is.
Then extract EVERY piece of visible text and data you can read or reasonably infer.

Format your response like this:
DOCUMENT TYPE: [what this document is]

KEY INFORMATION:
[List all important fields with labels — names, dates, ID numbers, addresses, amounts, etc.]

FULL EXTRACTED TEXT:
[All text visible in the document, in order]

Mark anything you are unsure about with [?].""",

    "Insurance Card": """You are reading an insurance card or insurance document. The image may be blurry — do your best.

Extract every field:
- Member Name
- Member ID / Policy Number
- Group Number
- Plan Name / Plan Type
- Insurance Company Name
- Effective Date
- Copay amounts (Primary, Specialist, ER, etc.)
- Deductible
- Phone numbers (Member Services, Claims, etc.)
- Payer ID
- Any other visible information

Mark unclear fields with [?]. If not visible, write N/A.""",

    "Driver's License / ID": """You are reading a driver's license or government-issued ID. May be blurry.

Extract:
- Full Name
- Date of Birth
- Address
- License / ID Number
- Expiration Date
- Issue Date
- State / Country
- License Class / Type
- Any restrictions

Mark unclear text with [?].""",

    "Tax Form (W-2, 1099, 1040)": """You are reading a tax form. May be blurry.

Identify the exact form type first (W-2, 1099-MISC, 1099-NEC, 1040, etc.)

Extract:
- Tax Year
- Taxpayer / Employee Name
- SSN (last 4 digits only if visible)
- Employer / Payer Name and Address
- EIN / Payer ID
- Every box number and its dollar amount
- All other visible fields

Mark unclear fields with [?].""",

    "Medical Record": """You are reading a medical document. May be blurry.

Extract:
- Patient Name
- Date of Birth
- Date of Service
- Provider / Doctor Name
- Facility Name
- Diagnosis codes or descriptions
- Medications and dosages
- Instructions or notes
- Important values or measurements
- Contact information

Mark unclear text with [?].""",

    "Letter / Notice": """You are reading an official letter or notice. May be blurry.

Extract:
- Sender name and organization
- Date
- Reference or case numbers
- Recipient name
- Subject or purpose
- Key dates
- Required actions
- Contact information
- Summary of what this letter is about

Mark unclear text with [?].""",

    "General Document": """You are an expert OCR system. Image may be blurry or low quality.
Extract ALL visible text and information including document type, all key fields with labels, and complete text in order.
Mark uncertain text with [?]."""
}

uploaded_files = st.file_uploader(
    "📎 Upload blurry client documents (photos, screenshots, scans)",
    type=["jpg", "jpeg", "png", "webp", "bmp", "tiff"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) ready to process")

col1, col2 = st.columns([3, 1])
with col1:
    run = st.button("🔍 Clean Up & Extract All Documents with AI", type="primary",
                    disabled=not uploaded_files, use_container_width=True)
with col2:
    enhance_only = st.button("🖼️ Enhance Image Only", disabled=not uploaded_files, use_container_width=True)

# Enhance only — no AI needed
if enhance_only and uploaded_files:
    for uf in uploaded_files:
        st.divider()
        st.markdown(f"### 🖼️ {uf.name}")
        img = Image.open(uf)
        enhanced = enhance_image(img, enhance_level)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Original**")
            st.image(img, use_container_width=True)
        with c2:
            st.markdown("**Enhanced**")
            st.image(enhanced, use_container_width=True)
            st.download_button(
                "⬇️ Download Enhanced Image",
                data=image_to_bytes(enhanced),
                file_name=uf.name.rsplit(".", 1)[0] + "_enhanced.png",
                mime="image/png",
                key="enh_" + uf.name
            )

# Full AI mode
if run and uploaded_files:
    if not api_key:
        st.error("⚠️ Please enter your Anthropic API key in the left sidebar first.")
        st.stop()

    client = anthropic.Anthropic(api_key=api_key)

    for uf in uploaded_files:
        st.divider()
        st.markdown(f"### 📄 {uf.name}")

        img = Image.open(uf)
        enhanced = enhance_image(img, enhance_level)

        left, right = st.columns(2)

        with left:
            t1, t2 = st.tabs(["📷 Original", "✨ Enhanced"])
            with t1:
                st.image(img, use_container_width=True)
            with t2:
                st.image(enhanced, use_container_width=True)
                st.download_button(
                    "⬇️ Download Enhanced Image",
                    data=image_to_bytes(enhanced),
                    file_name=uf.name.rsplit(".", 1)[0] + "_enhanced.png",
                    mime="image/png",
                    key="img_" + uf.name
                )

        with right:
            st.markdown("**AI Extracted Text & Data**")
            with st.spinner("🤖 AI is reading and extracting your document..."):
                try:
                    b64 = image_to_base64(enhanced)
                    message = client.messages.create(
                        model="claude-opus-4-6",
                        max_tokens=1500,
                        messages=[{
                            "role": "user",
                            "content": [
                                {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
                                {"type": "text", "text": prompts.get(doc_type, prompts["Auto-Detect"])}
                            ]
                        }]
                    )
                    extracted = message.content[0].text
                    st.markdown('<span class="badge-ok">✓ Successfully Extracted</span>', unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown(f'<div class="result-box">{extracted}</div>', unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.download_button(
                        "⬇️ Download Extracted Text",
                        data=extracted,
                        file_name=uf.name.rsplit(".", 1)[0] + "_extracted.txt",
                        mime="text/plain",
                        key="txt_" + uf.name
                    )
                except anthropic.AuthenticationError:
                    st.error("❌ Invalid API key. Check your key in the sidebar.")
                except anthropic.RateLimitError:
                    st.error("❌ Rate limit hit. Wait 30 seconds and try again.")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
