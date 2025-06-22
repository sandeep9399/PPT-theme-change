import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import io
import os

# Load Apollo theme presentation (must be in root directory)
THEME_FILE = "apollo_theme.pptx"
LOGO_FILE = "apollo_logo.png"

def apply_apollo_theme(uploaded_pptx):
    input_ppt = Presentation(uploaded_pptx)
    theme_ppt = Presentation(THEME_FILE)

    # Theme colors & layout from first slide
    theme_background = theme_ppt.slides[0].background

    for slide in input_ppt.slides:
        # Update background color
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = theme_background.fill.fore_color.rgb

        # Add footer if not already there
        has_footer = any("Powered by Apollo Knowledge" in shape.text for shape in slide.shapes if shape.has_text_frame)
        if not has_footer:
            textbox = slide.shapes.add_textbox(Inches(0.5), Inches(7.0), Inches(8), Inches(0.5))
            textbox.text = "Powered by Apollo Knowledge"
            textbox.text_frame.paragraphs[0].font.size = Inches(0.15)

        # Add Apollo logo top-right
        slide_width = input_ppt.slide_width
        logo_width = Inches(1.2)
        logo_height = Inches(0.6)
        slide.shapes.add_picture(LOGO_FILE, slide_width - logo_width - Inches(0.2), Inches(0.2), logo_width, logo_height)

    # Prepare final output
    output = io.BytesIO()
    input_ppt.save(output)
    output.seek(0)
    return output

st.set_page_config(page_title="Apollo PPT Themer", layout="centered")
st.title("🎓 Apollo PPT Themer")

uploaded_file = st.file_uploader("Upload your .pptx file", type=["pptx"])

if uploaded_file:
    st.success("Uploaded successfully. Applying Apollo Knowledge theme...")

    with st.spinner("Processing..."):
        themed_pptx = apply_apollo_theme(uploaded_file)

    base_name = uploaded_file.name.replace(".pptx", "")
    st.download_button(
        label="📥 Download Themed PPTX",
        data=themed_pptx,
        file_name=f"Apollo_Themed_{base_name}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
