import io
import json
import streamlit as st
import docx
from pypdf import PdfReader
from google import genai
from google.genai import types

st.set_page_config(page_title="Cover Sheet Prefiller (Gemini Edition)", layout="wide")

def extract_text_from_pdf(pdf_bytes):
    """Extract text from uploaded PDF CV."""
    reader = PdfReader(io.BytesIO(pdf_bytes))
    text = ""
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def extract_text_from_docx(docx_bytes):
    """Extract text from uploaded Word Docx CV."""
    doc = docx.Document(io.BytesIO(docx_bytes))
    text = []
    for paragraph in doc.paragraphs:
        if paragraph.text:
            text.append(paragraph.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text:
                    text.append(cell.text)
    return "\n".join(text)

def parse_template_questions(template_bytes):
    """
    Parses the cover sheet Word template and extracts (table_idx, row_idx, question) mapping.
    Maps left-column label directly to right-column input box.
    """
    doc = docx.Document(io.BytesIO(template_bytes))
    questions = []
    
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            if len(row.cells) >= 2:
                q_text = row.cells[0].text.strip()
                # Ensure it's a valid label and not a full-width header
                if q_text and row.cells[0].text != row.cells[1].text:
                    questions.append({
                        "table_idx": table_idx,
                        "row_idx": row_idx,
                        "question": q_text
                    })
    return questions

def extract_answers_with_gemini(cv_text, additional_text, questions, api_key):
    """
    Uses Google Gemini API to extract exact questions from the candidate documents.
    """
    client = genai.Client(api_key=api_key)
    
    q_list = [q["question"] for q in questions]
    
    prompt = f"""You are an assistant tasked with accurately populating a candidate cover sheet.

CRITICAL INSTRUCTIONS:
1. Answer ONLY questions that can be definitively answered from the CV or additional notes provided below.
2. DO NOT GUESS OR INFER information that is not explicitly stated.
3. If an answer is missing or uncertain, output null for that key.

Candidate CV / Document Text:
{cv_text}

Additional Notes / Email Info:
{additional_text}

Questions to answer:
{json.dumps(q_list, indent=2)}

Return a JSON object where each key matches the EXACT question text string provided above, and the value is either the string answer or null if unknown.
"""

    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json"
        )
    )
    
    return json.loads(response.text)

def write_answers_to_template(template_bytes, filled_answers, questions):
    """
    Writes finalized answers into column 1 (right cell) of each corresponding table row in Word doc.
    """
    doc = docx.Document(io.BytesIO(template_bytes))
    
    for q_info in questions:
        q_text = q_info["question"]
        table_idx = q_info["table_idx"]
        row_idx = q_info["row_idx"]
        
        val = filled_answers.get(q_text, "")
        if val is None:
            val = ""
            
        table = doc.tables[table_idx]
        row = table.rows[row_idx]
        
        # Write to right cell (index 1)
        if len(row.cells) > 1:
            row.cells[1].text = str(val)

    out_buffer = io.BytesIO()
    doc.save(out_buffer)
    out_buffer.seek(0)
    return out_buffer

# --- UI APP INTERFACE ---

st.title("📄 Candidate Cover Sheet Builder (Gemini)")
st.write("Automatically prefill Word cover sheet templates using Google Gemini.")

# Sidebar API Key & Template Input
st.sidebar.header("Configuration")

# Allow team members to input their own Gemini API key or use a fallback environment secret
user_api_key = st.sidebar.text_input("Enter your Gemini API Key:", type="password")
api_key = user_api_key or st.secrets.get("GEMINI_API_KEY", "")

st.sidebar.markdown("[Get a free Gemini API Key here](https://aistudio.google.com/)")

template_file = st.sidebar.file_uploader("Upload Word Cover Sheet Template (.docx)", type=["docx"])

if not api_key:
    st.warning("Please enter your Gemini API key in the sidebar to proceed.")

elif template_file:
    template_bytes = template_file.read()
    questions = parse_template_questions(template_bytes)
    st.sidebar.success(f"Template loaded! Found {len(questions)} fields.")

    st.header("1. Candidate Documents")
    col1, col2 = st.columns(2)
    
    with col1:
        cv_file = st.file_uploader("Upload CV (PDF or DOCX)", type=["pdf", "docx"])
    with col2:
        additional_text = st.text_area("Additional Info / Email Body Text", height=150)

    if cv_file and st.button("Process Documents with Gemini"):
        with st.spinner("Analyzing candidate files with Gemini..."):
            if cv_file.name.endswith(".pdf"):
                cv_text = extract_text_from_pdf(cv_file.read())
            else:
                cv_text = extract_text_from_docx(cv_file.read())

            extracted = extract_answers_with_gemini(cv_text, additional_text, questions, api_key)
            st.session_state["extracted_data"] = extracted
            st.session_state["processed"] = True

    # Step 2: Verification Form
    if st.session_state.get("processed", False):
        st.subheader("2. Review & Fill Missing Fields")
        st.info("AI filled known fields. Review entries or fill in remaining blanks below.")

        final_answers = {}
        extracted_data = st.session_state["extracted_data"]

        with st.form("review_form"):
            for item in questions:
                q_key = item["question"]
                default_val = extracted_data.get(q_key, "")
                if default_val is None:
                    default_val = ""
                
                c_q, c_a = st.columns([1, 2])
                with c_q:
                    st.write(f"**{q_key}**")
                with c_a:
                    final_answers[q_key] = st.text_input(
                        label=q_key, 
                        value=default_val, 
                        label_visibility="collapsed"
                    )
            
            submitted = st.form_submit_button("Generate Finished Cover Sheet")

        if submitted:
            output_doc = write_answers_to_template(template_bytes, final_answers, questions)
            
            st.success("Cover sheet updated successfully!")
            st.download_button(
                label="📥 Download Completed Cover Sheet (.docx)",
                data=output_doc,
                file_name="Completed_Cover_Sheet.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.info("Upload your `.docx` template in the left sidebar to get started.")
