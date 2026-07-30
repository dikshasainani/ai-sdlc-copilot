"""AI SDLC Copilot - Streamlit app."""

import json

import google.generativeai as genai
import streamlit as st

from doc_generators import (
    create_brd_doc,
    create_sql_doc,
    create_technical_test_doc,
    create_uat_doc,
    create_uat_excel,
    create_user_story_doc,
)
from prompts import OUTPUT_EXAMPLES, PROMPT_TEMPLATES, fill

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ==========================================
# PAGE SETTINGS
# ==========================================
st.set_page_config(page_title="AI SDLC Copilot", page_icon="🏦", layout="wide")
st.title("🏦 AI SDLC Copilot")
st.markdown(
    "Accelerate the Software Development Lifecycle with AI-powered "
    "Requirements, Testing, SQL, and Documentation."
)


# ==========================================
# MODEL + GENERATION HELPERS
# ==========================================
@st.cache_resource
def get_model():
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    return genai.GenerativeModel("models/gemini-2.5-flash")


def generate_json(prompt: str):
    """Call Gemini in JSON mode. Returns (data, error_message, raw_text)."""
    raw = ""
    try:
        response = get_model().generate_content(
            prompt,
            generation_config={
                "temperature": 0.2,
                "response_mime_type": "application/json",
            },
        )
        raw = response.text
        return json.loads(raw), None, raw
    except json.JSONDecodeError as exc:
        return None, f"Model returned invalid JSON: {exc}", raw
    except Exception as exc:  # API/network errors
        return None, f"Generation failed: {exc}", raw


def generate_text(prompt: str):
    """Call Gemini for free-text output. Returns (text, error_message)."""
    try:
        return get_model().generate_content(prompt).text, None
    except Exception as exc:
        return None, f"Generation failed: {exc}"


def run_json_generation(state_key: str, template_name: str, requirement: str,
                        spinner_text: str) -> None:
    """Validate input, call the model, store result in session state."""
    if not requirement.strip():
        st.warning("Please enter a requirement first.")
        return
    with st.spinner(spinner_text):
        data, error, raw = generate_json(fill(PROMPT_TEMPLATES[template_name], requirement))
    if error:
        st.error(error)
        if raw:
            st.code(raw)
        return
    st.session_state[state_key] = data


# ==========================================
# TABS
# ==========================================
tab_brd, tab_story, tab_uat, tab_sql, tab_tech, tab_templates = st.tabs([
    "📄 Requirements Document Generator",
    "📖 User Story Generator",
    "🧪 UAT Test Case Generator",
    "💻 SQL Query Generator",
    "🧩 Technical Test Case Generator",
    "📝 Prompt Templates",
])

# ------------------------------------------
# BRD
# ------------------------------------------
with tab_brd:
    brd_input = st.text_area("Paste Requirement", height=250, key="brd_input")
    if st.button("Generate BRD"):
        run_json_generation("brd_data", "BRD Prompt", brd_input, "Generating BRD...")

    if "brd_data" in st.session_state:
        st.json(st.session_state.brd_data, expanded=False)
        st.download_button(
            "📥 Download BRD",
            data=create_brd_doc(st.session_state.brd_data),
            file_name="Business_Requirements_Document.docx",
            mime=DOCX_MIME,
        )

# ------------------------------------------
# USER STORIES
# ------------------------------------------
with tab_story:
    story_input = st.text_area("Requirement", height=250, key="story_input")
    if st.button("Generate User Stories"):
        run_json_generation("story_data", "User Story Prompt", story_input,
                            "Generating user stories...")

    if "story_data" in st.session_state:
        st.json(st.session_state.story_data, expanded=False)
        st.download_button(
            "📥 Download User Stories",
            data=create_user_story_doc(st.session_state.story_data),
            file_name="User_Stories.docx",
            mime=DOCX_MIME,
        )

# ------------------------------------------
# UAT
# ------------------------------------------
with tab_uat:
    uat_input = st.text_area("Requirement", height=250, key="uat_input")
    if st.button("Generate UAT"):
        run_json_generation("uat_data", "UAT Prompt", uat_input,
                            "Generating UAT test cases...")

    if "uat_data" in st.session_state:
        st.json(st.session_state.uat_data, expanded=False)
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "📥 Download UAT (Word)",
                data=create_uat_doc(st.session_state.uat_data),
                file_name="UAT_Test_Cases.docx",
                mime=DOCX_MIME,
            )
        with col2:
            st.download_button(
                "📥 Download UAT (Excel)",
                data=create_uat_excel(st.session_state.uat_data),
                file_name="UAT_Test_Cases.xlsx",
                mime=XLSX_MIME,
            )

# ------------------------------------------
# SQL
# ------------------------------------------
with tab_sql:
    sql_input = st.text_area("Requirement", height=250, key="sql_input")
    if st.button("Generate SQL"):
        if not sql_input.strip():
            st.warning("Please enter a requirement first.")
        else:
            with st.spinner("Generating SQL..."):
                output, error = generate_text(fill(PROMPT_TEMPLATES["SQL Prompt"], sql_input))
            if error:
                st.error(error)
            else:
                st.session_state.sql_output = output

    if "sql_output" in st.session_state:
        st.markdown(st.session_state.sql_output)
        st.download_button(
            "📥 Download SQL",
            data=create_sql_doc(st.session_state.sql_output),
            file_name="SQL_Output.docx",
            mime=DOCX_MIME,
        )

# ------------------------------------------
# TECHNICAL TEST CASES
# ------------------------------------------
with tab_tech:
    st.subheader("🧩 Technical Test Case Generator")
    tech_input = st.text_area("Enter Requirement", height=250, key="tech_input")
    if st.button("Generate Technical Test Cases"):
        run_json_generation("tech_data", "Technical Test Case Prompt", tech_input,
                            "Generating technical test cases...")

    if "tech_data" in st.session_state:
        st.json(st.session_state.tech_data, expanded=False)
        st.download_button(
            "📥 Download Technical Test Cases",
            data=create_technical_test_doc(st.session_state.tech_data),
            file_name="Technical_Test_Cases.docx",
            mime=DOCX_MIME,
        )

# ------------------------------------------
# PROMPT TEMPLATES
# ------------------------------------------
with tab_templates:
    st.subheader("📝 BA Prompt Templates Library")
    prompt_type = st.selectbox("Select Template", list(PROMPT_TEMPLATES.keys()))

    st.markdown("## 📌 Prompt Used by AI")
    st.code(PROMPT_TEMPLATES[prompt_type], language="text")

    st.markdown("## 📌 Expected Output Structure")
    st.code(OUTPUT_EXAMPLES[prompt_type], language="text")

    st.download_button(
        "📥 Download Prompt Template",
        data=PROMPT_TEMPLATES[prompt_type],
        file_name=f"{prompt_type.replace(' ', '_')}.txt",
        mime="text/plain",
    )
