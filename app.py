import io
import time
import streamlit as st
from pypdf import PdfReader
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from google.genai import types
from google.genai.errors import APIError
from pydantic import BaseModel, Field
from typing import List

# =====================================================================
# 1. APPLICATION INITIALIZATION & CONFIGURATION
# =====================================================================
st.set_page_config(
    page_title="Document Layout AI Translator",
    page_icon="🌐",
    layout="centered"
)

st.title("PDF-to-Word Visual AI Translator")
st.write("Translates complex documents while maintaining layout blocks, headers, and page structures.")

# Initialize Persistent Session State for Downloads (Prevents Nested Button Trap)
if "translated_docx_bytes" not in st.session_state:
    st.session_state.translated_docx_bytes = None
if "current_file_name" not in st.session_state:
    st.session_state.current_file_name = ""

# =====================================================================
# 2. MASTER SECURITY ACCESS GATE
# =====================================================================
MASTER_PASSWORD = st.secrets.get("APP_PASSWORD")

if not MASTER_PASSWORD:
    st.error("No APP_PASSWORD configured in st.secrets. Set one before deploying.")
    st.stop()

user_password = st.text_input("Enter Cloud Access Password:", type="password")

if user_password != MASTER_PASSWORD:
    if user_password:
        st.error("Invalid Credentials. Please check your password configuration.")
    st.info("🔒 Authorization required to boot the programmatic translation pipeline.")
    st.stop()

# =====================================================================
# 3. USER INTERFACE CONTROLS
# =====================================================================
uploaded_file = st.file_uploader("Upload Source Document (PDF Format)", type="pdf")
target_lang = st.selectbox(
    "Translate to:",
    ["English", "Spanish", "French", "German", "Japanese", "Italian", "Portuguese", "Chinese"]
)

# Reset session state if a brand new file is uploaded
if uploaded_file and uploaded_file.name != st.session_state.current_file_name:
    st.session_state.translated_docx_bytes = None
    st.session_state.current_file_name = uploaded_file.name

# =====================================================================
# 4. STRUCTURED DATA SCHEMAS FOR GEMINI
# =====================================================================
class ContentBlock(BaseModel):
    block_type: str = Field(
        ...,
        description="The layout component type: 'heading_1', 'heading_2', 'paragraph', or 'page_number'."
    )
    text: str = Field(
        ...,
        description="The strictly translated text belonging to this layout block component."
    )

class PageLayoutTranslation(BaseModel):
    blocks: List[ContentBlock] = Field(
        ...,
        description="Sequential list of all structural layout blocks translated from the page."
    )

# =====================================================================
# 5. TRANSLATION PIPELINE CORE ENGINE
# =====================================================================
if uploaded_file and st.button("Execute Translation Blueprint"):
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

    reader = PdfReader(uploaded_file)
    total_pages = len(reader.pages)

    if total_pages == 0:
        st.error("This PDF appears to have no pages, or couldn't be read.")
        st.stop()

    doc = Document()

    progress_bar = st.progress(0)
    status_msg = st.empty()

    previous_page_context = "This is the first page. No prior context exists."
    pipeline_failed = False

    for idx in range(total_pages):
        status_msg.text(f"Processing & translating page {idx + 1} of {total_pages}...")

        current_page_text = reader.pages[idx].extract_text() or "[Blank Page Text]"

        prompt = f"""
        You are an elite expert document layout translator. Your task is to translate the current page text into {target_lang}.

        CONTEXT FROM PREVIOUS PAGE:
        \"\"\"{previous_page_context}\"\"\"

        Use the previous page context to maintain absolute pronoun consistency, narrative flow, and grammar logic across the page break boundaries.

        BLOCK ARCHITECTURE LAYOUT DIRECTIVES:
        1. heading_1: Used for top-level primary document title names.
        2. heading_2: Used for section subheadings or chapter dividers.
        3. paragraph: Standard continuous reading paragraphs or bullet records.
        4. page_number: Footers, headers, or isolated indicator numbers matching page coordinates.

        CRITICAL SAFETY WARNING:
        You are strictly forbidden from skipping, omitting, summarizing, or dropping page numbers or footers.
        Even if a page number or footer sits at the edge of the text array, you MUST capture it, translate any surrounding words, and classify it explicitly under the 'page_number' block type. Do not let it fade out.

        Current Page Content to Translate:
        \"\"\"{current_page_text}\"\"\"
        """

        # Call Gemini, with retry-once-then-skip handling so one bad page doesn't kill the run
        translation_data = None
        for attempt in range(2):
            try:
                response = client.models.generate_content(
                    model="gemini-2.5-flash",
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=PageLayoutTranslation,
                        temperature=0.2
                    )
                )
                translation_data = PageLayoutTranslation.model_validate_json(response.text)
                break
            except APIError as e:
                if attempt == 0:
                    status_msg.warning(f"Page {idx + 1}: API error ({e}). Retrying once...")
                    time.sleep(2)
                    continue
                st.error(f"Page {idx + 1} failed after retry: {e}. Stopping pipeline.")
                pipeline_failed = True
                break
            except Exception:
                # JSON didn't validate against the schema; fall back to raw text as a single paragraph
                status_msg.warning(f"Formatting anomaly on Page {idx + 1}. Using raw text fallback.")
                raw_text = getattr(response, "text", "") or "[Could not parse this page]"
                translation_data = PageLayoutTranslation(
                    blocks=[ContentBlock(block_type="paragraph", text=raw_text)]
                )
                break

        if pipeline_failed:
            break

        # Build next page's context from the last non-empty paragraph/heading blocks (skip page numbers)
        context_candidates = [
            b.text for b in translation_data.blocks
            if b.block_type != "page_number" and b.text.strip()
        ]
        if context_candidates:
            previous_page_context = (
                "The previous page ended with this translation text: "
                + " ".join(context_candidates[-2:])
            )

        # =====================================================================
        # 6. THE BLUE STYLING WORD ENGINE
        # =====================================================================
        for block in translation_data.blocks:
            b_type = block.block_type
            text = block.text.strip()

            if not text:
                continue

            if b_type == "heading_1":
                h = doc.add_heading(text, level=1)
                for run in h.runs:
                    run.font.name = 'Calibri'

            elif b_type == "heading_2":
                h = doc.add_heading(text, level=2)
                for run in h.runs:
                    run.font.name = 'Calibri'

            elif b_type == "page_number":
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                run = p.add_run(text)
                run.font.name = 'Calibri'
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0, 51, 204)  # Professional Crisp Blue
                run.bold = True

            else:  # Paragraph / Default
                p = doc.add_paragraph()
                run = p.add_run(text)
                run.font.name = 'Calibri'
                run.font.size = Pt(11)

        if idx < total_pages - 1:
            doc.add_page_break()

        progress_bar.progress((idx + 1) / total_pages)
        time.sleep(0.5)

    if not pipeline_failed:
        status_msg.success("All pages parsed and compiled successfully!")

        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)

        st.session_state.translated_docx_bytes = docx_buffer.getvalue()
    else:
        st.session_state.translated_docx_bytes = None

# =====================================================================
# 7. INDEPENDENT PERSISTENT DOWNLOAD ACCESS CONTROL
# =====================================================================
if st.session_state.translated_docx_bytes is not None:
    st.write("---")
    st.success("🎉 Document Translation Built Successfully!")

    st.download_button(
        label="📥 Download Translated Word Document (.docx)",
        data=st.session_state.translated_docx_bytes,
        file_name=f"translated_{st.session_state.current_file_name.replace('.pdf', '.docx')}",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )