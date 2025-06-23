
import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
import io

LOGO_FILE = "apollo_logo.png"
THEME_FILE = "apollo_theme.pptx"

def choose_best_layout(layouts, num_shapes):
    if num_shapes <= 2:
        return layouts[1]  # Title + Content
    elif num_shapes <= 4:
        return layouts[2]  # Two-Content layout
    else:
        return layouts[3] if len(layouts) > 3 else layouts[1]  # Fallback

def copy_shapes(source_slide, dest_slide):
    for shape in source_slide.shapes:
        if shape.shape_type == 1 and shape.has_text_frame:
            left = shape.left
            top = shape.top
            width = shape.width
            height = shape.height
            new_shape = dest_slide.shapes.add_textbox(left, top, width, height)
            new_shape.text = shape.text
        elif shape.shape_type == 13:  # picture
            image_stream = io.BytesIO(shape.image.blob)
            dest_slide.shapes.add_picture(image_stream, shape.left, shape.top, shape.width, shape.height)
        elif shape.shape_type == 6 and shape.chart:  # chart
            continue
        elif shape.shape_type == 19:  # table
            continue

def apply_apollo_theme(uploaded_pptx):
    source_ppt = Presentation(uploaded_pptx)
    theme_ppt = Presentation(THEME_FILE)
    layouts = theme_ppt.slide_layouts

    output_ppt = Presentation()
    output_ppt.slide_width = source_ppt.slide_width
    output_ppt.slide_height = source_ppt.slide_height

    for slide in source_ppt.slides:
        layout = choose_best_layout(layouts, len(slide.shapes))
        new_slide = output_ppt.slides.add_slide(layout)
        copy_shapes(slide, new_slide)

        slide_width = output_ppt.slide_width
        logo_width = Inches(1.2)
        logo_height = Inches(0.6)
        new_slide.shapes.add_picture(LOGO_FILE, slide_width - logo_width - Inches(0.2), Inches(0.2), logo_width, logo_height)

        textbox = new_slide.shapes.add_textbox(Inches(0.5), Inches(6.5), Inches(8), Inches(0.3))
        textbox.text = "Powered by Apollo Knowledge"
        textbox.text_frame.paragraphs[0].font.size = Inches(0.15)

    output = io.BytesIO()
    output_ppt.save(output)
    output.seek(0)
    return output

st.set_page_config(page_title="Apollo PPT Themer", layout="centered")
st.title("🎓 Apollo PPT Themer")

uploaded_file = st.file_uploader("📤 Upload a .pptx file", type=["pptx"])

if uploaded_file:
    st.success("Uploaded. Rebuilding with Apollo theme...")

    with st.spinner("Processing..."):
        result = apply_apollo_theme(uploaded_file)

    base_name = uploaded_file.name.replace(".pptx", "")
    st.download_button(
        label="📥 Download Themed PPTX",
        data=result,
        file_name=f"Apollo_Themed_{base_name}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
