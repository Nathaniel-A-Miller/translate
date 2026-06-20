import streamlit as st
import io
import time
import json
from pypdf import PdfReader, PdfWriter
from google import genai
from google.genai import types
from docx import Document  # Pure Python Word document framework
from docx.shared import RGBColor  # Handles custom color formatting

# 1. Page Configuration
st.set_page_config(page_title="Visual PDF to Word Translator", page_icon="🌐")
st.title("🎯 Secure PDF-to-Word Visual Translator")
st.write("Translates layout-heavy PDFs page-by-page directly into a continuous, downloadable Microsoft Word (.docx) file.")

# 2. Security Layer: Password Gate
MASTER_PASSWORD = st.secrets.get("APP_PASSWORD", "TranslateSecure2026")
user_password = st.text_input("Enter the access password to unlock this utility:", type="password")

if user_password != MASTER_PASSWORD:
    if user_password:
        st.error("Incorrect password. Please try again.")
    st.info("🔒 Secure tool. Access requires an authorization password.")
    st.stop()

st.success("🔓 Access Granted!")
st.divider()

# 3. User Controls
uploaded_file = st.file_uploader("Upload your source document (PDF format only):", type="pdf")
target_lang = st.selectbox(
    "Choose target language for translation:",
    ["English", "Spanish", "French", "German", "Japanese", "Simplified Chinese", "Italian", "Portuguese"]
)

# 4. Processing Pipeline
if uploaded_file and st.button("🚀 Start Page-by-Page Translation"):
    try:
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            st.error("Missing Gemini API Key. Please add 'GEMINI_API_KEY' to your secrets.")
            st.stop()

        client = genai.Client(api_key=api_key)
        reader = PdfReader(uploaded_file)
        total_pages = len(reader.pages)
        
        # Initialize a blank Microsoft Word document template
        doc = Document()
        
        # INITIALIZE SLIDING WINDOW CONTEXT
        previous_page_json = None
        
        progress_bar = st.progress(0)
        status_msg = st.empty()
        
        for idx in range(total_pages):
            status_msg.text(f"Processing and parsing page {idx + 1} of {total_pages}...")
            
            # Slice out exactly one page from the uploaded file into memory
            writer = PdfWriter()
            writer.add_page(reader.pages[idx])
            pdf_buffer = io.BytesIO()
            writer.write(pdf_buffer)
            pdf_bytes = pdf_buffer.getvalue()
            
            # Base prompt block
            prompt = """
            Analyze this page and return a valid JSON array of structural content blocks.
            """
            
            # DYNAMIC INJECTION: Feed previous page context into the prompt if it exists
            if idx > 0 and previous_page_json:
                prompt += f"""
                
                CONTEXT FROM THE IMMEDIATELY PRECEDING PAGE (FOR REFERENCE ONLY):
                {previous_page_json}
                
                CRITICAL INSTRUCTION: The JSON data provided above is the exact translation of the previous page. Use it solely to ensure seamless continuity, link trailing sentence fragments, and correctly resolve pronoun antecedents across the page boundary. 
                Do NOT include or re-translate any text from this context box in your final response. Only translate and output elements found on the new page slice.
                
                """
            
            # Layout Schema Mapping Block
            prompt += """
            Each block object in the array must follow one of these exact formatting models:
            1. Heading: {"type": "heading", "level": 1, "text": "Heading text"}
            2. Paragraph: {"type": "paragraph", "text": "Text content block."}
            3. Table: {"type": "table", "headers": ["Col 1", "Col 2"], "rows": [["Data A1", "Data B1"]]}
            4. Image Description: {"type": "image_description", "text": "Detailed visual description of chart or image"}
            5. Page Number: {"type": "page_number", "text": "Page X"}
            
            CRITICAL: Do not omit or ignore page numbers, headers, or footers found at the top or bottom of the page. You must capture them explicitly using the "page_number" block type.
            
            Provide only raw JSON. Do not include any leading/trailing conversational text or markdown wrappers.
            """
            
            # Force translation rules into the dominant System Instruction Layer
            system_rule = f"You are an expert document translation engine. Your absolute priority is to translate all text content directly into {target_lang}. Do not extract original language text; everything must be translated."
            
            # Request translation with native JSON structure enforcement
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[
                    types.Part.from_bytes(data=pdf_bytes, mime_type='application/pdf'),
                    prompt
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    system_instruction=system_rule
                )
            )
            
            # Parse the JSON response and map it to native Word properties
            try:
                blocks = json.loads(response.text)
                
                # UPDATE SLIDING WINDOW CONTEXT: Save this response text to feed into the next loop iteration
                previous_page_json = response.text
                
                for block in blocks:
                    b_type = block.get("type")
                    if b_type == "heading":
                        doc.add_heading(block.get("text", ""), level=block.get("level", 1))
                    elif b_type == "paragraph":
                        doc.add_paragraph(block.get("text", ""))
                    elif b_type == "table":
                        headers = block.get("headers", [])
                        rows = block.get("rows", [])
                        num_cols = max(len(headers), len(rows[0]) if rows else 1)
                        
                        table = doc.add_table(rows=0, cols=num_cols)
                        table.style = 'Table Grid'
                        
                        if headers:
                            hdr_cells = table.add_row().cells
                            for i, h_text in enumerate(headers):
                                if i < len(hdr_cells):
                                    hdr_cells[i].text = h_text
                        for row_data in rows:
                            row_cells = table.add_row().cells
                            for i, c_text in enumerate(row_data):
                                if i < len(row_cells):
                                    row_cells[i].text = c_text
                    elif b_type == "image_description":
                        p = doc.add_paragraph()
                        p.add_run(f"[ILLUSTRATION/CHART: {block.get('text', '')}]").italic = True
                    elif b_type == "page_number":
                        p = doc.add_paragraph()
                        run = p.add_run(block.get("text", ""))
                        run.font.color.rgb = RGBColor(0, 51, 204)  # Professional Crisp Blue (#0033CC)
                        run.font.bold = True
                        
            except Exception:
                # Safe fallback text if a page returns imperfect JSON layout data
                doc.add_paragraph(response.text)
                # Save it as context anyway so the chain doesn't break
                previous_page_json = response.text
                
            progress_bar.progress((idx + 1) / total_pages)
            time.sleep(1.5)
            
        status_msg.success("🎉 Document translation completed successfully!")
        
        # Save output document stream to memory
        docx_io = io.BytesIO()
        doc.save(docx_io)
        docx_io.seek(0)
        
        st.divider()
        st.balloons()
        st.download_button(
            label="📥 Download Translated Word Document (.docx)",
            data=docx_io,
            file_name="translated_document_output.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"An error occurred: {e}")