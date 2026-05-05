import streamlit as st
import anthropic
import base64
from PIL import Image, ImageEnhance, ImageFilter
import io

st.set_page_config(page_title="Vietrust Document Reader", page_icon="📄", layout="wide")

st.markdown("""
<style>
    .main-title { font-size: 28px; font-weight: 700; color: #185FA5; margin-bottom: 4px; }
    .sub-title { font-size: 14px; color: #999; margin-bottom: 24px; }
    .result-box { background: #f8f9fa; border-left: 4px solid #185FA5; padding: 16px; border-radius: 6px; font-size: 14px; line-height: 1.8; white-space: pre-wrap; }
    .success-badge { background: #EAF3DE; color: #3B6D11; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📄 Vietrust Document Reader</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">AI-powered — reads blurry, dark, and unreadable client documents</div>', unsafe_allow_html=True)

# API Key
with st.sidebar:
    st.header("⚙️ Settings")
    api_key = st.text_input("Anthropic API Key", type="password", placeholder="sk-ant-api03-...", help="Get your free key at console.anthropic.com")
    st.markdown("**Get a free key:** [console.anthropic.com](https://console.anthropic.com) → API Keys → Create Key")
    st.divider()
    doc_type = st.selectbox("Document Type", [
        "General Document",
        "Insurance Card",
        "Driver's License / ID",
        "Tax Form (W-2, 1099, 1040)",
        "Medical Record",
        "Letter / Notice"
    ])
    st.divider()
    st.markdown("**How to use:**\n1. Enter your API key\n2. Pick document type\n3. Upload your images\n4. Click Read Documents")

prompts = {
    "General Document": "You are an expert OCR system. This image may be blurry, dark, or low quality. Extract ALL visible text and data. State the document type, list all key fields (names, dates, numbers, addresses), then provide the full extracted text. Mark uncertain text with [?].",
    "Insurance Card": "You are reading an insurance card or document. It may be blurry. Extract: Member name, Member ID, Group number, Plan name, Insurance company, Effective date, Copays, Phone numbers, Payer ID. Mark unclear fields with [?]. Format with clear labels.",
    "Driver's License / ID": "You are reading a government ID or driver's license. Extract: Full name, Date of birth, Address, ID/License number, Expiration date, State. Mark unclear text with [?].",
    "Tax Form (W-2, 1099, 1040)": "You are reading a tax form. Identify the form type. Extract: Taxpayer name, Tax year, Employer/Payer info, all box numbers and dollar amounts. Mark unclear fields with [?].",
    "Medical Record": "You are reading a medical document. Extract: Patient name, DOB, Date of service, Provider, Diagnoses, Medications, Instructions, all important values. Mark unclear text with [?].",
    "Letter / Notice": "You are reading an official letter or notice. Extract: Sender, Date, Reference numbers, Recipient, Subject, Key dates, Action items, Contact info. Summarize the main purpose."
}

def enhance_image(img):
    img = img.convert("RGB")
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.0)
    enhancer = ImageEnhance.Brightness(img)
    img = enhancer.enhance(1.3)
    enhancer = ImageEnhance.Sharpness(img)
    img = enhancer.enhance(3.0)
    img = img.filter(ImageFilter.SHARPEN)
    return img

def image_to_base64(img):
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode()

uploaded_files = st.file_uploader(
    "Drop your blurry client documents here",
    type=["jpg", "jpeg", "png", "webp", "bmp"],
    accept_multiple_files=True,
    help="Upload multiple files at once — screenshots, photos, scanned docs all work"
)

if uploaded_files:
    st.markdown(f"**{len(uploaded_files)} file(s) ready to process**")

if st.button("🔍 Read & Extract All Documents with AI", type="primary", disabled=not uploaded_files):
    if not api_key:
        st.error("Please enter your Anthropic API key in the left sidebar first.")
        st.stop()

    client = anthropic.Anthropic(api_key=api_key)

    for uploaded_file in uploaded_files:
        st.divider()
        st.markdown(f"### 📄 {uploaded_file.name}")

        col1, col2 = st.columns(2)

        img = Image.open(uploaded_file)
        enhanced = enhance_image(img)

        with col1:
            tab1, tab2 = st.tabs(["Original", "Enhanced View"])
            with tab1:
                st.image(img, use_container_width=True)
            with tab2:
                st.image(enhanced, use_container_width=True)

        with col2:
            with st.spinner("AI is reading your document..."):
                try:
                    b64 = image_to_base64(img)
                    message = client.messages.create(
                        model="claude-opus-4-6",
                        max_tokens=1500,
                        messages=[{
                            "role": "user",
                            "content": [
                                {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
                                {"type": "text", "text": prompts[doc_type]}
                            ]
                        }]
                    )
                    extracted = message.content[0].text
                    st.markdown('<span class="success-badge">✓ AI Extracted Successfully</span>', unsafe_allow_html=True)
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown(f'<div class="result-box">{extracted}</div>', unsafe_allow_html=True)
                    st.download_button(
                        label="⬇️ Download Extracted Text",
                        data=extracted,
                        file_name=uploaded_file.name.rsplit(".", 1)[0] + "_extracted.txt",
                        mime="text/plain"
                    )
                except anthropic.AuthenticationError:
                    st.error("❌ Invalid API key. Please check your key in the sidebar.")
                except anthropic.RateLimitError:
                    st.error("❌ Rate limit hit. Wait 30 seconds and try again.")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
