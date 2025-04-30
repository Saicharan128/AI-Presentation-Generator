from pptx import Presentation
from pptx.util import Pt, Inches
from docx import Document
from fpdf import FPDF
import yaml
import io
from PIL import Image
import os
from services.image_service import (
    search_pexels_image,
    download_image, 
    fetch_consistent_background_image
)

def determine_optimal_font_size(slides_data, presentation):
    """Calculate optimal font size based on content amount."""
    font_sizes = [18, 16, 14, 12, 11, 10]
    content_width = presentation.slide_width * 0.85
    content_height = presentation.slide_height * 0.7

    max_lines = 0
    for slide_info in slides_data:
        bullets = slide_info.get('bullets', [])
        estimated_lines = 0
        for bullet in bullets:
            chars_per_line = int(content_width / Pt(12).inches * 2)
            bullet_lines = max(1, len(bullet) / chars_per_line)
            estimated_lines += bullet_lines + 0.5
        max_lines = max(max_lines, estimated_lines)

    for size in font_sizes:
        line_height = Pt(size).inches * 1.2
        max_possible = content_height / line_height
        if max_lines <= max_possible:
            return size
    return 10

def create_content_slides(prs, layout, title, bullets, font_size, background_img_data=None):
    """Create content slides with bullet points."""
    font_pt = Pt(font_size)
    
    content_height = prs.slide_height * 0.6
    content_top = Inches(1.5)
    content_left = Inches(1)
    content_width = prs.slide_width - Inches(2)

    line_height = font_pt.pt * 1.3
    max_lines = int(content_height / Pt(line_height))

    current, remaining = [], bullets[:]
    while remaining:
        current, lines = [], 0
        while remaining and lines < max_lines:
            bullet = remaining[0]
            chars_per_line = int(80 * (14 / font_size))
            est_lines = max(1, len(bullet) / chars_per_line) + 0.5
            if lines + est_lines <= max_lines:
                current.append(remaining.pop(0))
                lines += est_lines
            else:
                break

        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Use blank layout

        # Add background image (behind everything)
        if background_img_data:
            img_stream = io.BytesIO(background_img_data)
            slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

        # Add title on top
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), prs.slide_width - Inches(1.6), Inches(1))
        title_frame = title_box.text_frame
        title_frame.text = f"{title} (continued)" if remaining else title
        title_frame.paragraphs[0].font.size = Pt(font_size + 2)
        title_frame.paragraphs[0].font.bold = True

        # Add bullet points
        content_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
        tf = content_box.text_frame
        tf.word_wrap = True
        for i, bullet in enumerate(current):
            p = tf.add_paragraph()
            p.text = bullet
            p.level = 0
            p.font.size = font_pt

def add_image_slide(prs, layout, title, image_url, background_img_data=None):
    """Add a slide with an image related to the topic."""
    img_data = download_image(image_url)
    if img_data:
        slide = prs.slides.add_slide(layout)

        # Add background image
        if background_img_data:
            bg_stream = io.BytesIO(background_img_data)
            slide.shapes.add_picture(bg_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

        # Add main image
        img = Image.open(io.BytesIO(img_data))
        width, height = img.size
        max_w, max_h = Inches(10), Inches(5.5)
        scale = min(max_w / width, max_h / height)
        new_w, new_h = width * scale, height * scale
        left = (prs.slide_width - new_w) / 2
        top = (prs.slide_height - new_h + Inches(1)) / 2
        slide.shapes.title.text = f"{title} - Visual"
        slide.shapes.add_picture(io.BytesIO(img_data), left, top, width=new_w, height=new_h)

def create_pptx_from_yaml(yaml_content, output_path, topic):
    """Create PowerPoint presentation from YAML content."""
    try:
        data = yaml.safe_load(yaml_content)
        slides_data = data.get('presentation', {}).get('slides', [])
        if not slides_data:
            return {"success": False, "error": "No slides found"}

        prs = Presentation()
        layout = prs.slide_layouts[1]
        img_layout = prs.slide_layouts[5]
        font_size = determine_optimal_font_size(slides_data, prs)

        # Fetch consistent background image
        background_img_data = fetch_consistent_background_image()

        for slide in slides_data:
            title = slide.get('title', 'Untitled')
            bullets = slide.get('bullets', [])

            # Fetch image for the slide
            img_url = search_pexels_image(f"{topic} {title}")

            create_content_slides(prs, layout, title, bullets, font_size, background_img_data=background_img_data)

            # Optional: also add separate image-only slide
            if img_url:
                add_image_slide(prs, img_layout, title, img_url, background_img_data=background_img_data)

        prs.save(output_path)
        return {"success": True, "font_size": font_size}
    except Exception as e:
        return {"success": False, "error": str(e)}

def create_docx_from_yaml(yaml_content, output_path):
    """Create Word document from YAML content."""
    try:
        data = yaml.safe_load(yaml_content)
        slides = data.get('presentation', {}).get('slides', [])
        if not slides:
            return {"success": False, "error": "No slides found"}
        doc = Document()
        doc.add_heading('Presentation', 0)
        for slide in slides:
            doc.add_heading(slide.get('title', 'Untitled'), level=1)
            for bullet in slide.get('bullets', []):
                doc.add_paragraph(bullet, style='ListBullet')
        doc.save(output_path)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}

def create_pdf_from_yaml(yaml_content, output_path):
    """Create PDF document from YAML content."""
    try:
        data = yaml.safe_load(yaml_content)
        slides = data.get('presentation', {}).get('slides', [])
        if not slides:
            return {"success": False, "error": "No slides found"}
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=14)
        pdf.cell(0, 10, txt="Presentation", ln=True, align='C')
        pdf.ln(10)

        for slide in slides:
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(0, 10, txt=slide.get('title', 'Untitled Slide'), ln=True)
            pdf.set_font("Arial", size=11)
            for bullet in slide.get('bullets', []):
                pdf.set_x(20)
                pdf.multi_cell(0, 8, txt=f"- {bullet}")
                pdf.ln(2)
            pdf.ln(5)

        pdf.output(output_path)
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}