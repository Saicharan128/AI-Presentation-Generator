from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor as PPTXRGBColor
from docx import Document
from fpdf import FPDF
import yaml
import io
from PIL import Image
import os
import logging
import re
from services.image_service import (
    search_pexels_image,
    download_image,
    is_bright_image,
    fetch_consistent_background_image
)

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def preprocess_yaml_content(yaml_content):
    """Preprocess YAML content to fix bullet characters and indentation."""
    try:
        if not isinstance(yaml_content, str):
            logger.error(f"Expected string for yaml_content, got {type(yaml_content)}")
            return None

        # Replace bullet characters (•, *, etc.) with YAML-compatible hyphen (-)
        yaml_content = re.sub(r'^\s*[\•*]\s+', '  - ', yaml_content, flags=re.MULTILINE)
        
        # Ensure consistent indentation (2 spaces for YAML lists)
        lines = yaml_content.splitlines()
        cleaned_lines = []
        for line in lines:
            # Remove excessive whitespace and normalize indentation
            stripped = line.rstrip()
            if stripped:
                # Count leading spaces
                leading_spaces = len(line) - len(line.lstrip())
                # Adjust indentation to 2 spaces for list items starting with '-'
                if stripped.lstrip().startswith('-'):
                    cleaned_lines.append(' ' * (leading_spaces - (leading_spaces % 2)) + stripped.lstrip())
                else:
                    cleaned_lines.append(line)
            else:
                cleaned_lines.append(line)
        
        cleaned_yaml = '\n'.join(cleaned_lines)
        logger.debug(f"Preprocessed YAML content:\n{cleaned_yaml}")
        return cleaned_yaml
    except Exception as e:
        logger.error(f"Error preprocessing YAML content: {str(e)}")
        return None

def determine_optimal_font_size(slides_data, presentation):
    """Calculate optimal font size based on content amount."""
    try:
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
    except Exception as e:
        logger.error(f"Error in determine_optimal_font_size: {str(e)}")
        return 10

def create_content_slides(prs, layout, title, bullets, font_size, background_img_data=None):
    """Create content slides with bullet points."""
    try:
        font_pt = Pt(font_size)
        content_height = prs.slide_height * 0.6
        content_top = Inches(1.5)
        content_left = Inches(1)
        content_width = prs.slide_width - Inches(2)

        line_height = font_pt.pt * 1.3
        max_lines = int(content_height / line_height)

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

            slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

            # Add professional deep blue gradient background
            from pptx.enum.dml import MSO_FILL
            background = slide.background
            fill = background.fill
            fill.gradient()
            fill.gradient_angle = 45
            gradient_stops = fill.gradient_stops
            gradient_stops[0].color.rgb = PPTXRGBColor(11, 61, 145)  # Deep Blue
            gradient_stops[1].color.rgb = PPTXRGBColor(50, 100, 180)  # Lighter Blue

            # Add background image if provided
            if background_img_data:
                img_stream = io.BytesIO(background_img_data)
                slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

            # Add subtle gray border
            from pptx.enum.shapes import MSO_SHAPE
            border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.1), Inches(0.1), prs.slide_width - Inches(0.2), prs.slide_height - Inches(0.2))
            border.fill.solid()
            border.fill.fore_color.rgb = PPTXRGBColor(0, 0, 0)
            border.fill.fore_color.opacity = 0.0
            border.line.color.rgb = PPTXRGBColor(176, 196, 222)  # Soft Gray
            border.line.width = Pt(1)

            # Add title
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), prs.slide_width - Inches(1.6), Inches(1))
            title_frame = title_box.text_frame
            title_frame.text = f"{title} (continued)" if remaining else title
            title_p = title_frame.paragraphs[0]
            title_p.font.size = Pt(font_size + 2)
            title_p.font.bold = True
            title_p.font.color.rgb = PPTXRGBColor(245, 245, 220)  # Light Beige
            title_p.font.name = 'Arial'

            # Add bullet points
            content_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
            tf = content_box.text_frame
            tf.word_wrap = True
            for bullet in current:
                p = tf.add_paragraph()
                p.text = bullet
                p.level = 0
                p.font.size = font_pt
                p.font.color.rgb = PPTXRGBColor(245, 245, 220)  # Light Beige
                p.font.name = 'Arial'
    except Exception as e:
        logger.error(f"Error in create_content_slides: {str(e)}")
        raise

def add_image_slide(prs, layout, title, image_url, background_img_data=None):
    """Add a slide with an image related to the topic."""
    try:
        img_data = download_image(image_url)
        if not img_data:
            logger.warning(f"Failed to download image from {image_url}")
            return

        slide = prs.slides.add_slide(layout)

        # Add professional deep blue gradient background
        from pptx.enum.dml import MSO_FILL
        background = slide.background
        fill = background.fill
        fill.gradient()
        fill.gradient_angle = 45
        gradient_stops = fill.gradient_stops
        gradient_stops[0].color.rgb = PPTXRGBColor(11, 61, 145)
        gradient_stops[1].color.rgb = PPTXRGBColor(50, 100, 180)

        # Add background image
        if background_img_data:
            bg_stream = io.BytesIO(background_img_data)
            slide.shapes.add_picture(bg_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

        # Add subtle gray border
        from pptx.enum.shapes import MSO_SHAPE
        border = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.1), Inches(0.1), prs.slide_width - Inches(0.2), prs.slide_height - Inches(0.2))
        border.fill.solid()
        border.fill.fore_color.rgb = PPTXRGBColor(0, 0, 0)
        border.fill.fore_color.opacity = 0.0
        border.line.color.rgb = PPTXRGBColor(176, 196, 222)
        border.line.width = Pt(1)

        # Add main image
        img = Image.open(io.BytesIO(img_data))
        width, height = img.size
        max_w, max_h = Inches(10), Inches(5.5)
        scale = min(max_w / width, max_h / height)
        new_w, new_h = width * scale, height * scale
        left = (prs.slide_width - new_w) / 2
        top = (prs.slide_height - new_h + Inches(1)) / 2
        slide.shapes.title.text = f"{title} - Visual"
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = PPTXRGBColor(245, 245, 220)
        slide.shapes.title.text_frame.paragraphs[0].font.name = 'Arial'
        slide.shapes.add_picture(io.BytesIO(img_data), left, top, width=new_w, height=new_h)
    except Exception as e:
        logger.error(f"Error in add_image_slide: {str(e)}")
        raise

def create_pptx_from_yaml(yaml_content, output_path, topic):
    """Create PowerPoint presentation from YAML content."""
    try:
        # Validate inputs
        if not isinstance(yaml_content, str):
            logger.error(f"Expected string for yaml_content, got {type(yaml_content)}")
            return {"success": False, "error": f"Expected string for yaml_content, got {type(yaml_content)}"}
        if not yaml_content.strip():
            logger.error("yaml_content is empty")
            return {"success": False, "error": "yaml_content is empty"}
        if not output_path.endswith('.pptx'):
            logger.error("Output path must end with .pptx")
            return {"success": False, "error": "Output path must end with .pptx"}

        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # Preprocess YAML content
        cleaned_yaml = preprocess_yaml_content(yaml_content)
        if not cleaned_yaml:
            logger.error("Failed to preprocess YAML content")
            return {"success": False, "error": "Failed to preprocess YAML content"}

        # Parse YAML
        try:
            data = yaml.safe_load(cleaned_yaml)
        except yaml.YAMLError as e:
            logger.error(f"YAML parsing error: {str(e)}")
            return {"success": False, "error": f"YAML parsing error: {str(e)}"}

        if not data or 'presentation' not in data:
            logger.error("Invalid YAML structure: missing 'presentation' key")
            return {"success": False, "error": "Invalid YAML structure: missing 'presentation' key"}
        
        slides_data = data.get('presentation', {}).get('slides', [])
        if not slides_data:
            logger.error("No slides found in YAML")
            return {"success": False, "error": "No slides found"}

        prs = Presentation()
        layout = prs.slide_layouts[1]  # Bullet slide layout
        img_layout = prs.slide_layouts[5]  # Title only layout
        font_size = determine_optimal_font_size(slides_data, prs)

        # Fetch consistent background image
        background_img_data = fetch_consistent_background_image()

        for slide in slides_data:
            title = slide.get('title', 'Untitled').strip()
            bullets = [b.strip() for b in slide.get('bullets', []) if b.strip()]
            if not title and not bullets:
                logger.warning(f"Skipping empty slide")
                continue

            img_url = search_pexels_image(f"{topic} {title}")
            logger.info(f"Image URL for slide '{title}': {img_url}")

            create_content_slides(prs, layout, title, bullets, font_size, background_img_data=background_img_data)

            # Add image slide if URL is valid
            if img_url:
                add_image_slide(prs, img_layout, title, img_url, background_img_data=background_img_data)

        prs.save(output_path)
        logger.info(f"PowerPoint saved to {output_path}")
        return {"success": True, "font_size": font_size}
    except Exception as e:
        logger.error(f"Error in create_pptx_from_yaml: {str(e)}")
        return {"success": False, "error": str(e)}

def create_docx_from_yaml(yaml_content, output_path):
    """Create Word document from YAML content."""
    try:
        # Validate inputs
        if not isinstance(yaml_content, str):
            return {"success": False, "error": f"Expected string for yaml_content, got {type(yaml_content)}"}
        if not yaml_content.strip():
            return {"success": False, "error": "yaml_content is empty"}
        if not output_path.endswith('.docx'):
            return {"success": False, "error": "Output path must end with .docx"}

        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # Preprocess YAML content
        cleaned_yaml = preprocess_yaml_content(yaml_content)
        if not cleaned_yaml:
            logger.error("Failed to preprocess YAML content")
            return {"success": False, "error": "Failed to preprocess YAML content"}

        # Parse YAML
        try:
            data = yaml.safe_load(cleaned_yaml)
        except yaml.YAMLError as e:
            logger.error(f"YAML parsing error: {str(e)}")
            return {"success": False, "error": f"YAML parsing error: {str(e)}"}

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
        logger.info(f"Word document saved to {output_path}")
        return {"success": True}
    except Exception as e:
        logger.error(f"Error in create_docx_from_yaml: {str(e)}")
        return {"success": False, "error": str(e)}

def create_pdf_from_yaml(yaml_content, output_path):
    """Create PDF document from YAML content."""
    try:
        # Validate inputs
        if not isinstance(yaml_content, str):
            return {"success": False, "error": f"Expected string for yaml_content, got {type(yaml_content)}"}
        if not yaml_content.strip():
            return {"success": False, "error": "yaml_content is empty"}
        if not output_path.endswith('.pdf'):
            return {"success": False, "error": "Output path must end with .pdf"}

        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # Preprocess YAML content
        cleaned_yaml = preprocess_yaml_content(yaml_content)
        if not cleaned_yaml:
            logger.error("Failed to preprocess YAML content")
            return {"success": False, "error": "Failed to preprocess YAML content"}

        # Parse YAML
        try:
            data = yaml.safe_load(cleaned_yaml)
        except yaml.YAMSError as e:
            logger.error(f"YAML parsing error: {str(e)}")
            return {"success": False, "error": f"YAML parsing error: {str(e)}"}

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
        logger.info(f"PDF saved to {output_path}")
        return {"success": True}
    except Exception as e:
        logger.error(f"Error in create_pdf_from_yaml: {str(e)}")
        return {"success": False, "error": str(e)}

def create_html_from_yaml(yaml_content, output_path, html_presentation_type='minimalist'):
    try:
        if not isinstance(yaml_content, str):
            logger.error(f"Expected string for yaml_content, got {type(yaml_content)}")
            return {"success": False, "error": f"Expected string for yaml_content, got {type(yaml_content)}"}
        if not yaml_content.strip():
            logger.error("yaml_content is empty")
            return {"success": False, "error": "yaml_content is empty"}
        if not output_path.endswith('.html'):
            logger.error("Output path must end with .html")
            return {"success": False, "error": "Output path must end with .html"}

        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        cleaned_yaml = preprocess_yaml_content(yaml_content)
        if not cleaned_yaml:
            logger.error("Failed to preprocess YAML content")
            return {"success": False, "error": "Failed to preprocess YAML content"}

        try:
            data = yaml.safe_load(cleaned_yaml)
        except yaml.YAMLError as e:
            logger.error(f"YAML parsing error: {str(e)}")
            return {"success": False, "error": f"YAML parsing error: {str(e)}"}

        if not data or 'presentation' not in data:
            logger.error("Invalid YAML structure: missing 'presentation' key")
            return {"success": False, "error": "Invalid YAML structure: missing 'presentation' key"}

        presentation_title = data.get('presentation', {}).get('title', 'AI-Generated Presentation')
        slides = data.get('presentation', {}).get('slides', [])
        if not slides:
            logger.error("No slides found in YAML")
            return {"success": False, "error": "No slides found"}

        styles = {
            'professional': """
                @import url('https://fonts.googleapis.com/css2?family=Arial&display=swap');
                body { font-family: 'Arial', sans-serif; margin: 0; padding: 0; background: #1a1a2e; overflow: hidden; }
                .reveal .slides section { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); color: #ffffff; height: 100%; display: flex; flex-direction: column; justify-content: center; padding: 40px; box-sizing: border-box; transition: transform 0.6s ease-in-out; }
                .reveal .slides section.present { opacity: 1; transform: translateX(0); }
                .reveal .slides section:not(.present) { opacity: 0.4; }
                .slide-content { max-width: 800px; margin: 0 auto; animation: fadeIn 0.5s ease-in-out; }
                .slide-title { font-size: 2em; margin-bottom: 1em; color: #a78bfa; text-align: center; }
                .slide-bullets { text-align: left; margin-left: 2em; font-size: 0.875em; line-height: 1.5; color: #ffffff; }
                .slide-bullets li { margin-bottom: 0.5em; }
                .nav-bar { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); display: flex; gap: 10px; }
                .nav-dot { width: 12px; height: 12px; background: #ffffff; border-radius: 50%; cursor: pointer; transition: background 0.3s ease-in-out, transform 0.3s ease-in-out; }
                .nav-dot.active { background: #a78bfa; box-shadow: 0 0 10px #a78bfa; }
                .nav-dot:hover { background: #d1c4e9; transform: scale(1.2); }
                .progress-bar { position: fixed; bottom: 0; left: 0; height: 5px; background: #a78bfa; transition: width 0.6s ease-in-out; }
                .slide-indicator { position: fixed; bottom: 20px; right: 20px; background: rgba(0, 0, 0, 0.5); color: #ffffff; padding: 5px 10px; border-radius: 20px; font-size: 12px; }
                @keyframes fadeIn { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
            """,
            'modern': """
                @import url('https://fonts.googleapis.com/css2?family=Segoe+UI:wght@400;700&display=swap');
                body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; background: #f5f5f5; overflow-x: hidden; }
                .presentation { width: 100vw; height: 100vh; position: relative; background: #ffffff; overflow: hidden; }
                .header { background: #3498db; color: #ffffff; padding: 20px; text-align: center; position: absolute; top: 0; left: 0; right: 0; z-index: 10; }
                .header h1 { margin: 0; font-size: 24px; font-weight: 700; }
                .slides-container { height: calc(100vh - 120px); width: 100%; position: absolute; top: 80px; overflow: hidden; }
                .slides { display: flex; transition: transform 0.6s ease-in-out; height: 100%; width: 100%; }
                .slide { min-width: 100%; height: 100%; padding: 40px; box-sizing: border-box; display: flex; flex-direction: column; justify-content: center; }
                .slide-title { font-size: 2.5em; margin-bottom: 30px; color: #2c3e50; text-align: center; }
                .slide-content { font-size: 1.4em; margin-bottom: 30px; color: #333; text-align: center; max-width: 800px; margin-left: auto; margin-right: auto; }
                ul.bullets { max-width: 800px; margin-left: auto; margin-right: auto; padding-left: 30px; }
                ul.bullets li { margin-bottom: 15px; line-height: 1.6; font-size: 1.3em; color: #333; }
                .nav-bar { position: fixed; bottom: 20px; left: 50%; transform: translateX(-50%); display: flex; gap: 20px; z-index: 100; }
                .nav-dot { background: #3498db; color: #ffffff; border: none; border-radius: 50%; width: 50px; height: 50px; font-size: 24px; cursor: pointer; display: flex; align-items: center; justify-content: center; box-shadow: 0 2px 5px rgba(0,0,0,0.2); transition: background 0.3s ease-in-out, transform 0.3s ease-in-out; }
                .nav-dot.active { background: #2980b9; }
                .nav-dot:hover { background: #2980b9; transform: scale(1.1); }
                .nav-dot:disabled { background: #bdc3c7; cursor: not-allowed; }
                .slide-indicator { position: fixed; bottom: 20px; right: 20px; background: rgba(0,0,0,0.5); color: #ffffff; padding: 8px 12px; border-radius: 20px; font-size: 14px; }
                .progress-bar { position: fixed; bottom: 0; left: 0; height: 5px; background: #3498db; transition: width 0.6s ease-in-out; }
            """,
            'minimalist': """
                @import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@300;400&display=swap');
                body { font-family: 'Open Sans', sans-serif; margin: 0; padding: 0; background: #4A7043; overflow: hidden; }
                .reveal .slides section { background: transparent; color: #FDF6E3; height: 100%; display: flex; flex-direction: column; justify-content: center; padding: 40px; box-sizing: border-box; transition: transform 0.6s ease-in-out; }
                .reveal .slides section.present { opacity: 1; transform: translateX(0); }
                .reveal .slides section:not(.present) { opacity: 0; }
                .slide-content { max-width: 800px; margin: 0 auto; animation: fadeIn 0.5s ease-in-out; }
                .slide-title { font-size: 2em; margin-bottom: 20px; color: #E07A5F; text-align: center; font-weight: 300; }
                .slide-bullets { text-align: left; margin-left: 0; font-size: 1em; line-height: 1.6; color: #FDF6E3; }
                .slide-bullets li { margin-bottom: 0.6em; }
                .nav-bar { position: fixed; bottom: 10px; left: 50%; transform: translateX(-50%); display: flex; gap: 8px; }
                .nav-dot { width: 8px; height: 8px; background: #FDF6E3; border-radius: 50%; cursor: pointer; transition: background 0.3s ease-in-out, transform 0.3s ease-in-out; }
                .nav-dot.active { background: #E07A5F; }
                .nav-dot:hover { background: #FFFFFF; opacity: 0.7; transform: scale(1.2); }
                .progress-bar { position: fixed; bottom: 0; left: 0; height: 5px; background: #E07A5F; transition: width 0.6s ease-in-out; }
                .slide-indicator { position: fixed; bottom: 20px; right: 20px; background: rgba(0,0,0,0.5); color: #ffffff; padding: 5px 10px; border-radius: 20px; font-size: 12px; }
                @keyframes fadeIn { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
            """
        }

        if html_presentation_type not in styles:
            logger.warning(f"Invalid html_presentation_type: {html_presentation_type}, defaulting to minimalist")
            html_presentation_type = 'minimalist'

        custom_style = styles[html_presentation_type]

        slide_sections = ""
        for slide in slides:
            title = slide.get('title', 'Untitled Slide').strip()
            bullets_list = [b.strip() for b in slide.get('bullets', []) if b.strip()]
            if not title and not bullets_list:
                logger.warning(f"Skipping empty slide")
                continue

            bullets = "<ul class='slide-bullets'>" + "".join(f"<li>{b}</li>" for b in bullets_list) + "</ul>" if bullets_list else ""

            if html_presentation_type == 'modern':
                slide_sections += f'''
                <div class="slide">
                    <h2 class="slide-title">{title}</h2>
                    {bullets.replace('slide-bullets', 'bullets')}
                </div>
                '''
            else:
                slide_sections += f'''
                <section class="{'slide-background' if html_presentation_type == 'professional' else ''}">
                    <div class="slide-content">
                        <h2 class="slide-title">{title}</h2>
                        {bullets}
                    </div>
                </section>
                '''

        html_template = ""
        if html_presentation_type == 'modern':
            html_template = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{presentation_title}</title>
    <style>
        {custom_style}
    </style>
</head>
<body>
    <div class="presentation">
        <div class="header">
            <h1>{presentation_title}</h1>
        </div>
        <div class="slides-container">
            <div class="slides" id="slides">
                {slide_sections}
            </div>
        </div>
        <div class="nav-bar" id="navBar"></div>
        <div class="slide-indicator" id="slideIndicator">1 / {len(slides)}</div>
        <div class="progress-bar" id="progressBar"></div>
    </div>
    <script>
        document.addEventListener('DOMContentLoaded', function() {{
            const slides = document.getElementById('slides');
            const slideIndicator = document.getElementById('slideIndicator');
            const progressBar = document.getElementById('progressBar');
            const navBar = document.getElementById('navBar');
            const totalSlides = {len(slides)};
            let currentSlide = 0;

            // Create navigation dots
            for (let i = 0; i < totalSlides; i++) {{
                const dot = document.createElement('button');
                dot.className = 'nav-dot';
                dot.addEventListener('click', () => {{
                    currentSlide = i;
                    updateSlide();
                }});
                navBar.appendChild(dot);
            }}

            // Initialize
            updateSlide();

            // Keyboard navigation
            document.addEventListener('keydown', function(e) {{
                if (e.key === 'ArrowLeft' && currentSlide > 0) {{
                    currentSlide--;
                    updateSlide();
                }}
                if (e.key === 'ArrowRight' && currentSlide < totalSlides - 1) {{
                    currentSlide++;
                    updateSlide();
                }}
            }});

            function updateSlide() {{
                slides.style.transform = `translateX(-${{currentSlide * 100}}%)`;
                slideIndicator.textContent = `${{currentSlide + 1}} / ${{totalSlides}}`;
                const progress = ((currentSlide + 1) / totalSlides) * 100;
                progressBar.style.width = `${{progress}}%`;
                const dots = navBar.querySelectorAll('.nav-dot');
                dots.forEach((dot, index) => {{
                    dot.classList.toggle('active', index === currentSlide);
                    dot.disabled = index === currentSlide;
                }});
            }}
        }});
    </script>
</body>
</html>'''
        else:
            html_template = f'''<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{presentation_title}</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/reveal.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/theme/black.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/plugin/highlight/monokai.min.css">
    <style>
        {custom_style}
    </style>
</head>
<body>
    <div class="reveal">
        <div class="slides">
            {slide_sections}
        </div>
    </div>
    <div class="nav-bar" id="navBar"></div>
    <div class="slide-indicator" id="slideIndicator">1 / {len(slides)}</div>
    <div class="progress-bar" id="progressBar"></div>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/reveal.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/reveal.js/4.3.1/plugin/highlight/highlight.min.js"></script>
    <script>
        const deck = Reveal.initialize({{
            hash: true,
            plugins: [RevealHighlight],
            transition: 'slide',
            transitionSpeed: 'default',
            backgroundTransition: 'slide',
            progress: false,
            controls: false
        }});

        const slideIndicator = document.getElementById('slideIndicator');
        const progressBar = document.getElementById('progressBar');
        const navBar = document.getElementById('navBar');
        const totalSlides = Reveal.getTotalSlides();

        // Create navigation dots
        for (let i = 0; i < totalSlides; i++) {{
            const dot = document.createElement('div');
            dot.className = 'nav-dot';
            dot.addEventListener('click', () => Reveal.slide(i));
            navBar.appendChild(dot);
        }}

        function updateSlideInfo() {{
            const currentSlide = Reveal.getSlidePastCount() + 1;
            slideIndicator.textContent = `${{currentSlide}} / ${{totalSlides}}`;
            const progress = (currentSlide / totalSlides) * 100;
            progressBar.style.width = `${{progress}}%`;
            const dots = navBar.querySelectorAll('.nav-dot');
            dots.forEach((dot, index) => {{
                dot.classList.toggle('active', index === currentSlide - 1);
            }});
        }}

        deck.addEventListener('slidechanged', updateSlideInfo);
        deck.addEventListener('ready', updateSlideInfo);
    </script>
</body>
</html>'''

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_template)
        
        logger.info(f"HTML presentation saved to {output_path}")
        return {"success": True}
    except Exception as e:
        logger.error(f"Error in create_html_from_yaml: {str(e)}")
        return {"success": False, "error": str(e)}