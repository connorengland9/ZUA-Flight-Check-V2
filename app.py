import streamlit as st
import os
import glob
import main  

st.set_page_config(page_title="CE ZUA Flight Check Tool", layout="centered")

page_bg_css = """
<style>
/* Custom Black Background with Cross-Hatch and Diagonal "CE" */
.stApp {
    background-color: #0b0b0b !important;
    background-image: url("data:image/svg+xml,%3Csvg width='120' height='120' viewBox='0 0 120 120' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg stroke='rgba(255,255,255,0.04)' stroke-width='1'%3E%3Cpath d='M0 0l120 120M120 0L0 120'/%3E%3C/g%3E%3Ctext x='60' y='68' font-family='Georgia, serif' font-size='24' font-style='italic' font-weight='bold' fill='rgba(255,255,255,0.06)' text-anchor='middle' transform='rotate(-45 60 60)'%3ECE%3C/text%3E%3C/g%3E%3C/svg%3E") !important;
    background-repeat: repeat !important;
}

/* Center all main text and headers */
.block-container {
    text-align: center;
}

/* Force standard text to white so it doesn't turn black in Light Mode */
h1, h2, h3, p {
    color: #ffffff !important;
}

/* --- UPLOADER BOX OVERRIDE --- */
/* Force the drag-and-drop box to have a dark background and Light Blue dashed border */
[data-testid="stFileUploadDropzone"] {
    background-color: #1a1a1a !important;
    border: 2px dashed #87CEFA !important;
    border-radius: 10px !important;
}

/* Force ALL text inside the uploader (including the 200MB limit) to Light Blue */
[data-testid="stFileUploader"] * {
    color: #87CEFA !important;
}

/* Style the 'Browse files' button inside the uploader */
[data-testid="stFileUploader"] button {
    background-color: transparent !important;
    border: 1px solid #87CEFA !important;
}
[data-testid="stFileUploader"] button:hover {
    background-color: #87CEFA !important;
}
[data-testid="stFileUploader"] button:hover * {
    color: #0b0b0b !important;
}

/* --- DOWNLOAD BUTTONS OVERRIDE --- */
/* Center and style the download buttons */
[data-testid="stDownloadButton"] > button {
    margin: 0 auto;
    display: block;
    background-color: transparent !important;
    border: 2px solid #87CEFA !important;
}
[data-testid="stDownloadButton"] > button * {
    color: #87CEFA !important;
}

/* Hover effect for download buttons */
[data-testid="stDownloadButton"] > button:hover {
    background-color: #87CEFA !important;
}
[data-testid="stDownloadButton"] > button:hover * {
    color: #0b0b0b !important; /* Changes text to dark when hovered */
}

/* Center the primary Generate button */
.stButton > button {
    margin: 0 auto;
    display: block;
}
</style>
"""
st.markdown(page_bg_css, unsafe_allow_html=True)

st.title("CE ZUA Flight Check Tool")
st.write("Upload the raw FAA Flight Check PDF to instantly generate the Overview Memo and Floor Briefing.")

if 'files_ready' not in st.session_state:
    st.session_state.files_ready = False

_, col_upload, _ = st.columns([1, 4, 1])
with col_upload:
    uploaded_file = st.file_uploader("Upload Flight Check PDF", type=["pdf"])

if uploaded_file is not None:
    _, col_gen, _ = st.columns([1, 2, 1])
    with col_gen:
        generate_pressed = st.button("Generate Documents", type="primary", use_container_width=True)
        
    if generate_pressed:
        with st.spinner("Processing PDF and generating documents... This might take a minute."):
            
            os.makedirs(main.INPUT_DIR, exist_ok=True)
            os.makedirs(main.OUTPUT_DIR, exist_ok=True)
            
            for f in glob.glob(os.path.join(main.OUTPUT_DIR, "*")):
                os.remove(f)

            input_path = os.path.join(main.INPUT_DIR, uploaded_file.name)
            with open(input_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            try:
                main.extract_filter_and_translate(uploaded_file.name)
                
                output_files = os.listdir(main.OUTPUT_DIR)
                
                st.session_state.memo_file = next((f for f in output_files if "Memo" in f and f.endswith(".docx")), None)
                st.session_state.briefing_file = next((f for f in output_files if "Briefing" in f and f.endswith(".docx")), None)
                st.session_state.files_ready = True
                
            except Exception as e:
                st.error(f"An error occurred while processing: {e}")

if st.session_state.files_ready:
    st.markdown("<p style='color: #4CAF50; font-weight: bold; font-size: 1.1em; margin-top: 10px; margin-bottom: 10px;'>Documents generated successfully!</p>", unsafe_allow_html=True)
    
    _, col_dl1, col_dl2, _ = st.columns([1, 3, 3, 1])
    
    if st.session_state.memo_file:
        with col_dl1:
            with open(os.path.join(main.OUTPUT_DIR, st.session_state.memo_file), "rb") as f:
                st.download_button(
                    label="Download Overview Memo", 
                    data=f, 
                    file_name=st.session_state.memo_file, 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
    if st.session_state.briefing_file:
        with col_dl2:
            with open(os.path.join(main.OUTPUT_DIR, st.session_state.briefing_file), "rb") as f:
                st.download_button(
                    label="Download Floor Briefing", 
                    data=f, 
                    file_name=st.session_state.briefing_file, 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
