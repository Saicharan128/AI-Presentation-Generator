from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor as PPTXRGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from docx import Document
from fpdf import FPDF
import yaml
import io
from PIL import Image
import os
import logging
import re
import requests
import base64
from services.image_service import (
    search_pexels_image,
    download_image,
    fetch_consistent_background_image
)
from config import Config
from jinja2 import Template

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Pexels API configuration
PEXELS_API_KEY = Config.PEXELS_API_KEY
PEXELS_API_URL = "https://api.pexels.com/v1/search"

# HTML templates for different presentation styles
HTML_TEMPLATES = {
    'minimalist': """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }}</title>
    <style>
        /* Reset and Base Styles */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Helvetica Neue', Arial, sans-serif;
            background-color: #f0f0f0;
            overflow-x: hidden;
        }

        /* Presentation Container */
        .presentation {
            width: 100vw;
            height: 100vh;
            background-color: #ffffff;
            position: relative;
        }

        /* Header Styles */
        .header {
            background-color: #ffffff;
            color: #333333;
            padding: 1.5rem;
            text-align: center;
            border-bottom: 1px solid #e0e0e0;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 10;
        }
        .header h1 {
            font-size: 2rem;
            font-weight: 300;
        }

        /* Slides Container */
        .slides-container {
            position: absolute;
            top: 5rem;
            width: 100%;
            height: calc(100vh - 5rem);
            overflow: hidden;
        }
        .slides {
            display: flex;
            height: 100%;
            transition: transform 0.5s ease;
        }

        /* Slide Styles */
        .slide {
            min-width: 100%;
            height: 100%;
            display: flex;
            flex-direction: row;
            align-items: center;
            padding: 2rem;
        }
        .slide-image-container {
            width: 40%;
            padding: 1rem;
        }
        .slide-image {
            max-width: 100%;
            max-height: 30vh;
            object-fit: cover;
            border-radius: 0.5rem;
        }
        .slide-content-container {
            width: 60%;
            padding: 1rem;
        }
        .slide-title {
            font-size: 1.8rem;
            font-weight: 500;
            color: #333333;
            margin-bottom: 1rem;
        }
        .slide-content {
            font-size: 1.2rem;
            color: #555555;
            margin-bottom: 1rem;
        }
        .bullets {
            list-style-type: disc;
            padding-left: 1.5rem;
        }
        .bullets li {
            font-size: 1.1rem;
            color: #555555;
            margin-bottom: 0.5rem;
        }

        /* Navigation Styles */
        .navigation {
            position: fixed;
            bottom: 1.5rem;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            gap: 1rem;
            z-index: 10;
        }
        .nav-btn {
            background-color: #333333;
            color: #ffffff;
            border: none;
            border-radius: 0.3rem;
            width: 2.5rem;
            height: 2.5rem;
            font-size: 1.2rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .nav-btn:hover {
            background-color: #555555;
        }
        .nav-btn:disabled {
            background-color: #cccccc;
            cursor: not-allowed;
        }

        /* Slide Indicator and Progress Bar */
        .slide-indicator {
            position: fixed;
            bottom: 1.5rem;
            right: 1.5rem;
            background-color: #333333;
            color: #ffffff;
            padding: 0.4rem 0.8rem;
            border-radius: 1rem;
            font-size: 0.9rem;
        }
        .progress-bar {
            position: fixed;
            bottom: 0;
            left: 0;
            height: 0.25rem;
            background-color: #333333;
            transition: width 0.3s;
        }

        /* Responsive Design */
        @media screen and (max-width: 768px) {
            .slide {
                flex-direction: column;
            }
            .slide-image-container,
            .slide-content-container {
                width: 100%;
            }
            .slide-image {
                max-height: 20vh;
            }
            .slide-title {
                font-size: 1.5rem;
            }
            .slide-content,
            .bullets li {
                font-size: 1rem;
            }
            .header h1 {
                font-size: 1.5rem;
            }
        }
    </style>
</head>
<body>
    <div class="presentation">
        <header class="header">
            <h1>{{ title }}</h1>
        </header>
        <div class="slides-container">
            <div class="slides" id="slides">
                {% for slide in slides %}
                <section class="slide">
                    {% if slide.image %}
                    <div class="slide-image-container">
                        <img class="slide-image" src="{{ slide.image }}" alt="{{ slide.title }}">
                    </div>
                    {% endif %}
                    <div class="slide-content-container">
                        <h2 class="slide-title">{{ slide.title }}</h2>
                        {% if slide.content %}
                        <p class="slide-content">{{ slide.content }}</p>
                        {% endif %}
                        {% if slide.bullets %}
                        <ul class="bullets">
                            {% for bullet in slide.bullets %}
                            <li>{{ bullet }}</li>
                            {% endfor %}
                        </ul>
                        {% endif %}
                    </div>
                </section>
                {% endfor %}
            </div>
        </div>
        <nav class="navigation">
            <button class="nav-btn" id="prevBtn" aria-label="Previous Slide">←</button>
            <button class="nav-btn" id="nextBtn" aria-label="Next Slide">→</button>
        </nav>
        <div class="slide-indicator" id="slideIndicator" aria-live="polite">1 / {{ slides|length }}</div>
        <div class="progress-bar" id="progressBar" role="progressbar"></div>
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', () => {
            const slides = document.getElementById('slides');
            const prevBtn = document.getElementById('prevBtn');
            const nextBtn = document.getElementById('nextBtn');
            const slideIndicator = document.getElementById('slideIndicator');
            const progressBar = document.getElementById('progressBar');
            const totalSlides = {{ slides|length }};
            let currentSlide = 0;

            const updateSlide = () => {
                slides.style.transform = `translateX(-${currentSlide * 100}%)`;
                slideIndicator.textContent = `${currentSlide + 1} / ${totalSlides}`;
                prevBtn.disabled = currentSlide === 0;
                nextBtn.disabled = currentSlide === totalSlides - 1;
                progressBar.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
            };

            const previousSlide = () => {
                if (currentSlide > 0) {
                    currentSlide--;
                    updateSlide();
                }
            };

            const nextSlide = () => {
                if (currentSlide < totalSlides - 1) {
                    currentSlide++;
                    updateSlide();
                }
            };

            prevBtn.addEventListener('click', previousSlide);
            nextBtn.addEventListener('click', nextSlide);
            document.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowLeft') previousSlide();
                if (e.key === 'ArrowRight') nextSlide();
            });
            window.addEventListener('resize', updateSlide);

            updateSlide();
        });
    </script>
</body>
</html>
""",
    'modern': """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }}</title>
    <style>
        /* Reset and Base Styles */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Roboto', sans-serif;
            background-color: #e3f2fd;
            overflow-x: hidden;
        }

        /* Presentation Container */
        .presentation {
            width: 100vw;
            height: 100vh;
            background-color: #ffffff;
            position: relative;
            box-shadow: 0 0.25rem 0.75rem rgba(0, 0, 0, 0.1);
        }

        /* Header Styles */
        .header {
            background: linear-gradient(90deg, #0288d1, #4fc3f7);
            color: #ffffff;
            padding: 1rem;
            text-align: center;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 10;
        }
        .header h1 {
            font-size: 2.2rem;
            font-weight: 500;
        }

        /* Slides Container */
        .slides-container {
            position: absolute;
            top: 4.5rem;
            width: 100%;
            height: calc(100vh - 4.5rem);
            overflow: hidden;
        }
        .slides {
            display: flex;
            height: 100%;
            transition: transform 0.5s ease;
        }

        /* Slide Styles */
        .slide {
            min-width: 100%;
            height: 100%;
            display: flex;
            flex-direction: row;
            align-items: center;
            padding: 2rem;
            background-color: #fafafa;
        }
        .slide-image-container {
            width: 40%;
            padding: 1rem;
        }
        .slide-image {
            max-width: 100%;
            max-height: 30vh;
            object-fit: cover;
            border-radius: 0.625rem;
            box-shadow: 0 0.125rem 0.5rem rgba(0, 0, 0, 0.15);
        }
        .slide-content-container {
            width: 60%;
            padding: 1rem;
        }
        .slide-title {
            font-size: 2rem;
            font-weight: 600;
            color: #0277bd;
            margin-bottom: 1.25rem;
        }
        .slide-content {
            font-size: 1.3rem;
            color: #424242;
            margin-bottom: 1.25rem;
        }
        .bullets {
            list-style-type: square;
            padding-left: 1.5rem;
        }
        .bullets li {
            font-size: 1.2rem;
            color: #424242;
            margin-bottom: 0.75rem;
        }

        /* Navigation Styles */
        .navigation {
            position: fixed;
            bottom: 1.5rem;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            gap: 1.25rem;
            z-index: 10;
        }
        .nav-btn {
            background-color: #0288d1;
            color: #ffffff;
            border: none;
            border-radius: 50%;
            width: 3rem;
            height: 3rem;
            font-size: 1.25rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .nav-btn:hover {
            background-color: #0277bd;
        }
        .nav-btn:disabled {
            background-color: #b0bec5;
            cursor: not-allowed;
        }

        /* Slide Indicator and Progress Bar */
        .slide-indicator {
            position: fixed;
            bottom: 1.5rem;
            right: 1.5rem;
            background-color: #0288d1;
            color: #ffffff;
            padding: 0.5rem 0.75rem;
            border-radius: 1.25rem;
            font-size: 0.875rem;
        }
        .progress-bar {
            position: fixed;
            bottom: 0;
            left: 0;
            height: 0.3125rem;
            background-color: #4fc3f7;
            transition: width 0.3s;
        }

        /* Responsive Design */
        @media screen and (max-width: 768px) {
            .slide {
                flex-direction: column;
            }
            .slide-image-container,
            .slide-content-container {
                width: 100%;
            }
            .slide-image {
                max-height: 20vh;
            }
            .slide-title {
                font-size: 1.6rem;
            }
            .slide-content,
            .bullets li {
                font-size: 1.1rem;
            }
            .header h1 {
                font-size: 1.8rem;
            }
        }
    </style>
</head>
<body>
    <div class="presentation">
        <header class="header">
            <h1>{{ title }}</h1>
        </header>
        <div class="slides-container">
            <div class="slides" id="slides">
                {% for slide in slides %}
                <section class="slide">
                    {% if slide.image %}
                    <div class="slide-image-container">
                        <img class="slide-image" src="{{ slide.image }}" alt="{{ slide.title }}">
                    </div>
                    {% endif %}
                    <div class="slide-content-container">
                        <h2 class="slide-title">{{ slide.title }}</h2>
                        {% if slide.content %}
                        <p class="slide-content">{{ slide.content }}</p>
                        {% endif %}
                        {% if slide.bullets %}
                        <ul class="bullets">
                            {% for bullet in slide.bullets %}
                            <li>{{ bullet }}</li>
                            {% endfor %}
                        </ul>
                        {% endif %}
                    </div>
                </section>
                {% endfor %}
            </div>
        </div>
        <nav class="navigation">
            <button class="nav-btn" id="prevBtn" aria-label="Previous Slide">←</button>
            <button class="nav-btn" id="nextBtn" aria-label="Next Slide">→</button>
        </nav>
        <div class="slide-indicator" id="slideIndicator" aria-live="polite">1 / {{ slides|length }}</div>
        <div class="progress-bar" id="progressBar" role="progressbar"></div>
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', () => {
            const slides = document.getElementById('slides');
            const prevBtn = document.getElementById('prevBtn');
            const nextBtn = document.getElementById('nextBtn');
            const slideIndicator = document.getElementById('slideIndicator');
            const progressBar = document.getElementById('progressBar');
            const totalSlides = {{ slides|length }};
            let currentSlide = 0;

            const updateSlide = () => {
                slides.style.transform = `translateX(-${currentSlide * 100}%)`;
                slideIndicator.textContent = `${currentSlide + 1} / ${totalSlides}`;
                prevBtn.disabled = currentSlide === 0;
                nextBtn.disabled = currentSlide === totalSlides - 1;
                progressBar.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
            };

            const previousSlide = () => {
                if (currentSlide > 0) {
                    currentSlide--;
                    updateSlide();
                }
            };

            const nextSlide = () => {
                if (currentSlide < totalSlides - 1) {
                    currentSlide++;
                    updateSlide();
                }
            };

            prevBtn.addEventListener('click', previousSlide);
            nextBtn.addEventListener('click', nextSlide);
            document.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowLeft') previousSlide();
                if (e.key === 'ArrowRight') nextSlide();
            });
            window.addEventListener('resize', updateSlide);

            updateSlide();
        });
    </script>
</body>
</html>
""",
    'professional': """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }}</title>
    <style>
        /* Reset and Base Styles */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Arial', sans-serif;
            background-color: #eceff1;
            overflow-x: hidden;
        }

        /* Presentation Container */
        .presentation {
            width: 100vw;
            height: 100vh;
            background-color: #ffffff;
            position: relative;
        }

        /* Header Styles */
        .header {
            background-color: #263238;
            color: #ffffff;
            padding: 1.5rem;
            text-align: center;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 10;
        }
        .header h1 {
            font-size: 2rem;
            font-weight: 400;
        }

        /* Slides Container */
        .slides-container {
            position: absolute;
            top: 5rem;
            width: 100%;
            height: calc(100vh - 5rem);
            overflow: hidden;
        }
        .slides {
            display: flex;
            height: 100%;
            transition: transform 0.5s ease;
        }

        /* Slide Styles */
        .slide {
            min-width: 100%;
            height: 100%;
            display: flex;
            flex-direction: row;
            align-items: center;
            padding: 2rem;
            background-color: #ffffff;
        }
        .slide-image-container {
            width: 40%;
            padding: 1rem;
        }
        .slide-image {
            max-width: 100%;
            max-height: 30vh;
            object-fit: cover;
            border: 1px solid #e0e0e0;
        }
        .slide-content-container {
            width: 60%;
            padding: 1rem;
        }
        .slide-title {
            font-size: 1.9rem;
            font-weight: 500;
            color: #263238;
            margin-bottom: 1rem;
        }
        .slide-content {
            font-size: 1.2rem;
            color: #37474f;
            margin-bottom: 1rem;
        }
        .bullets {
            list-style-type: circle;
            padding-left: 1.5rem;
        }
        .bullets li {
            font-size: 1.1rem;
            color: #37474f;
            margin-bottom: 0.5rem;
        }

        /* Navigation Styles */
        .navigation {
            position: fixed;
            bottom: 1.5rem;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            gap: 1rem;
            z-index: 10;
        }
        .nav-btn {
            background-color: #263238;
            color: #ffffff;
            border: none;
            border-radius: 0.3rem;
            width: 2.5rem;
            height: 2.5rem;
            font-size: 1.2rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .nav-btn:hover {
            background-color: #37474f;
        }
        .nav-btn:disabled {
            background-color: #b0bec5;
            cursor: not-allowed;
        }

        /* Slide Indicator and Progress Bar */
        .slide-indicator {
            position: fixed;
            bottom: 1.5rem;
            right: 1.5rem;
            background-color: #263238;
            color: #ffffff;
            padding: 0.4rem 0.8rem;
            border-radius: 1rem;
            font-size: 0.9rem;
        }
        .progress-bar {
            position: fixed;
            bottom: 0;
            left: 0;
            height: 0.25rem;
            background-color: #263238;
            transition: width 0.3s;
        }

        /* Responsive Design */
        @media screen and (max-width: 768px) {
            .slide {
                flex-direction: column;
            }
            .slide-image-container,
            .slide-content-container {
                width: 100%;
            }
            .slide-image {
                max-height: 20vh;
            }
            .slide-title {
                font-size: 1.5rem;
            }
            .slide-content,
            .bullets li {
                font-size: 1rem;
            }
            .header h1 {
                font-size: 1.5rem;
            }
        }
    </style>
</head>
<body>
    <div class="presentation">
        <header class="header">
            <h1>{{ title }}</h1>
        </header>
        <div class="slides-container">
            <div class="slides" id="slides">
                {% for slide in slides %}
                <section class="slide">
                    {% if slide.image %}
                    <div class="slide-image-container">
                        <img class="slide-image" src="{{ slide.image }}" alt="{{ slide.title }}">
                    </div>
                    {% endif %}
                    <div class="slide-content-container">
                        <h2 class="slide-title">{{ slide.title }}</h2>
                        {% if slide.content %}
                        <p class="slide-content">{{ slide.content }}</p>
                        {% endif %}
                        {% if slide.bullets %}
                        <ul class="bullets">
                            {% for bullet in slide.bullets %}
                            <li>{{ bullet }}</li>
                            {% endfor %}
                        </ul>
                        {% endif %}
                    </div>
                </section>
                {% endfor %}
            </div>
        </div>
        <nav class="navigation">
            <button class="nav-btn" id="prevBtn" aria-label="Previous Slide">←</button>
            <button class="nav-btn" id="nextBtn" aria-label="Next Slide">→</button>
        </nav>
        <div class="slide-indicator" id="slideIndicator" aria-live="polite">1 / {{ slides|length }}</div>
        <div class="progress-bar" id="progressBar" role="progressbar"></div>
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', () => {
            const slides = document.getElementById('slides');
            const prevBtn = document.getElementById('prevBtn');
            const nextBtn = document.getElementById('nextBtn');
            const slideIndicator = document.getElementById('slideIndicator');
            const progressBar = document.getElementById('progressBar');
            const totalSlides = {{ slides|length }};
            let currentSlide = 0;

            const updateSlide = () => {
                slides.style.transform = `translateX(-${currentSlide * 100}%)`;
                slideIndicator.textContent = `${currentSlide + 1} / ${totalSlides}`;
                prevBtn.disabled = currentSlide === 0;
                nextBtn.disabled = currentSlide === totalSlides - 1;
                progressBar.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
            };

            const previousSlide = () => {
                if (currentSlide > 0) {
                    currentSlide--;
                    updateSlide();
                }
            };

            const nextSlide = () => {
                if (currentSlide < totalSlides - 1) {
                    currentSlide++;
                    updateSlide();
                }
            };

            prevBtn.addEventListener('click', previousSlide);
            nextBtn.addEventListener('click', nextSlide);
            document.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowLeft') previousSlide();
                if (e.key === 'ArrowRight') nextSlide();
            });
            window.addEventListener('resize', updateSlide);

            updateSlide();
        });
    </script>
</body>
</html>
"""
}

def fetch_image_from_pexels(query):
    """Fetch an image from Pexels API."""
    try:
        headers = {"Authorization": PEXELS_API_KEY}
        params = {"query": query, "per_page": 1}
        response = requests.get(PEXELS_API_URL, headers=headers, params=params)
        if response.status_code == 200:
            data = response.json()
            if data['photos']:
                img_url = data['photos'][0]['src']['large']
                img_response = requests.get(img_url)
                return io.BytesIO(img_response.content)
        logger.warning(f"No image found for query '{query}'")
        return None
    except Exception as e:
        logger.error(f"Error fetching image from Pexels: {str(e)}")
        return None

def lighten_image(img_stream, factor=0.9):
    """Lighten an image for use as a background."""
    try:
        img = Image.open(img_stream).convert("RGB")
        enhancer = Image.new("RGB", img.size, (255, 255, 255))
        img = Image.blend(img, enhancer, factor)
        new_stream = io.BytesIO()
        img.save(new_stream, format='PNG')
        new_stream.seek(0)
        return new_stream
    except Exception as e:
        logger.error(f"Error lightening image: {str(e)}")
        return img_stream

def add_shadow_to_shape(shape):
    """Add a shadow effect to a shape."""
    try:
        sp = shape._element
        spPr = sp.find('{http://schemas.openxmlformats.org/drawingml/2006/main}spPr')
        if spPr is None:
            spPr = OxmlElement('a:spPr')
            sp.append(spPr)
        
        effect_lst = OxmlElement('a:effectLst')
        outer_shdw = OxmlElement('a:outerShdw')
        outer_shdw.set('dist', '20000')
        outer_shdw.set('dir', '2700000')
        outer_shdw.set('algn', 'ctr')
        srgb_clr = OxmlElement('a:srgbClr')
        srgb_clr.set('val', '000000')
        alpha = OxmlElement('a:alpha')
        alpha.set('val', '40000')
        srgb_clr.append(alpha)
        outer_shdw.append(srgb_clr)
        effect_lst.append(outer_shdw)
        spPr.append(effect_lst)
    except Exception as e:
        logger.warning(f"Failed to apply shadow to shape: {e}")

def apply_element_properties(shape, properties):
    """Apply formatting properties to a shape."""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = True
    for paragraph in tf.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Calibri'
            if 'font_size' in properties:
                run.font.size = Pt(properties['font_size'])
            if 'font_color' in properties:
                r, g, b = properties['font_color']
                run.font.color.rgb = PPTXRGBColor(r, g, b)
    if 'alignment' in properties:
        align_map = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}
        tf.paragraphs[0].alignment = align_map.get(properties['alignment'], PP_ALIGN.LEFT)
    if properties.get('shadow', False):
        try:
            add_shadow_to_shape(shape)
        except Exception as e:
            logger.warning(f"Failed to apply shadow: {e}")

def add_custom_image(slide, img_stream, properties):
    """Add an image to a slide with specified properties."""
    try:
        left = Inches(properties.get('position', [0, 0])[0])
        top = Inches(properties.get('position', [0, 0])[1])
        width = Inches(properties.get('size', [6, 4])[0])
        height = Inches(properties.get('size', [6, 4])[1])
        slide.shapes.add_picture(img_stream, left, top, width=width, height=height)
    except Exception as e:
        logger.error(f"Error adding custom image: {str(e)}")

def add_slide(prs, slide_data, slide_config, topic, background_img_data=None):
    """Add a formatted slide to the presentation with dynamic positioning."""
    try:
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)

        if background_img_data:
            light_stream = lighten_image(io.BytesIO(background_img_data))
            add_custom_image(slide, light_stream, {
                "position": [0, 0],
                "size": [10, 7.5]
            })

        header_shape = slide.shapes.add_shape(
            1,
            Inches(0), Inches(0),
            Inches(10), Inches(slide_config['header']['height'])
        )
        fill = header_shape.fill
        fill.solid()
        r, g, b = slide_config['header']['color']
        fill.fore_color.rgb = PPTXRGBColor(r, g, b)

        title_text = slide_data.get('title', 'Untitled')
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(1.0))
        tf = title_box.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.auto_size = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title_text
        apply_element_properties(title_box, slide_config['title'])

        content_y = 1.8
        max_bullets = min(len(slide_data.get('bullets', [])), 3)

        for i, item in enumerate(slide_data.get('bullets', [])[:max_bullets]):
            content_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(content_y), Inches(9.0), Inches(0.8)
            )
            tf = content_box.text_frame
            tf.clear()
            tf.word_wrap = True
            tf.auto_size = True
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = f"• {item}"
            apply_element_properties(content_box, slide_config['content'])
            content_y += 1.0 + (len(item) // 80) * 0.4

        content = slide_data.get('bullets', [])
        icon_keyword = (content[0].split()[0].lower() if content else topic.lower().split()[0])
        icon_stream = fetch_image_from_pexels(icon_keyword)
        if icon_stream:
            footer_y = max(6.0, content_y + 0.5)
            footer_config = slide_config['footer'].copy()
            footer_config['position'] = [6.5, footer_y]
            add_custom_image(slide, icon_stream, footer_config)

        return slide
    except Exception as e:
        logger.error(f"Error adding slide: {str(e)}")
        raise

def preprocess_yaml_content(yaml_content):
    """Preprocess YAML content to fix bullet characters and indentation."""
    try:
        if not isinstance(yaml_content, str):
            logger.error(f"Expected string for yaml_content, got {type(yaml_content)}")
            return None

        # Remove any extra document markers
        yaml_content = re.sub(r'^---\s*$', '', yaml_content, flags=re.MULTILINE)
        yaml_content = yaml_content.strip()
        
        # Fix bullet points
        yaml_content = re.sub(r'^\s*[\•*]\s+', '  - ', yaml_content, flags=re.MULTILINE)
        
        lines = yaml_content.splitlines()
        cleaned_lines = []
        for line in lines:
            stripped = line.rstrip()
            if stripped:
                leading_spaces = len(line) - len(line.lstrip())
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

def create_pptx_from_yaml(yaml_content, output_path, topic):
    """Create PowerPoint presentation from YAML content."""
    try:
        if not isinstance(yaml_content, str):
            logger.error(f"Expected string for yaml_content, got {type(yaml_content)}")
            return {"success": False, "error": f"Expected string for yaml_content, got {type(yaml_content)}"}
        if not yaml_content.strip():
            logger.error("yaml_content is empty")
            return {"success": False, "error": "yaml_content is empty"}
        if not output_path.endswith('.pptx'):
            logger.error("Output path must end with .pptx")
            return {"success": False, "error": "Output path must end with .pptx"}

        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        cleaned_yaml = preprocess_yaml_content(yaml_content)
        if not cleaned_yaml:
            logger.error("Failed to preprocess YAML content")
            return {"success": False, "error": "Failed to preprocess YAML content"}

        try:
            # First try to load as single document
            data = yaml.safe_load(cleaned_yaml)
            if data is None:
                # If empty, try loading all documents and take the first one
                documents = list(yaml.safe_load_all(cleaned_yaml))
                if documents:
                    data = documents[0]
                else:
                    raise yaml.YAMLError("Empty YAML content")
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

        include_images = data.get('presentation', {}).get('include_images', True)

        prs = Presentation()

        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title_shape = title_slide.shapes.title
        title_shape.text = data.get('presentation', {}).get('title', topic)
        title_shape.text_frame.paragraphs[0].font.size = Pt(44)
        title_shape.text_frame.paragraphs[0].font.color.rgb = PPTXRGBColor(0, 51, 102)
        subtitle = title_slide.placeholders[1]
        subtitle.text = f"Exploring {topic}"
        subtitle.text_frame.paragraphs[0].font.size = Pt(24)
        subtitle.text_frame.paragraphs[0].font.color.rgb = PPTXRGBColor(50, 50, 50)

        background_img_data = fetch_consistent_background_image(topic)

        slide_config = {
            "header": {
                "color": [0, 51, 102],
                "height": 1.2
            },
            "title": {
                "font_size": 32,
                "font_color": [255, 255, 255],
                "alignment": "center",
                "shadow": True
            },
            "content": {
                "font_size": 22,
                "font_color": [10, 10, 10],
                "position": [0.5, 2.0],
                "shadow": True
            },
            "footer": {
                "color": [100, 100, 100],
                "position": [6.5, 4.0],
                "size": [3.0, 3.0]
            }
        }

        for slide_data in slides_data:
            if include_images:
                img_url = search_pexels_image(topic, slide_title=slide_data.get('title'))
                if img_url:
                    img_data = download_image(img_url)
                    if img_data:
                        img_stream = io.BytesIO(img_data)
                        slide = prs.slides.add_slide(prs.slide_layouts[5])
                        slide.shapes.title.text = f"{slide_data.get('title', 'Untitled')} - Visual"
                        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = PPTXRGBColor(245, 245, 220)
                        slide.shapes.title.text_frame.paragraphs[0].font.name = 'Arial'
                        img = Image.open(img_stream)
                        width, height = img.size
                        max_w, max_h = Inches(10), Inches(5.5)
                        scale = min(max_w / width, max_h / height)
                        new_w, new_h = width * scale, height * scale
                        left = (prs.slide_width - new_w) / 2
                        top = (prs.slide_height - new_h + Inches(1)) / 2
                        slide.shapes.add_picture(io.BytesIO(img_data), left, top, width=new_w, height=new_h)
            add_slide(prs, slide_data, slide_config, topic, background_img_data)

        prs.save(output_path)
        logger.info(f"PowerPoint saved to {output_path}")
        return {"success": True}
    except Exception as e:
        logger.error(f"Error in create_pptx_from_yaml: {str(e)}")
        return {"success": False, "error": str(e)}

def create_docx_from_yaml(yaml_content, output_path):
    """Create Word document from YAML content."""
    try:
        if not isinstance(yaml_content, str):
            return {"success": False, "error": f"Expected string for yaml_content, got {type(yaml_content)}"}
        if not yaml_content.strip():
            return {"success": False, "error": "yaml_content is empty"}
        if not output_path.endswith('.docx'):
            return {"success": False, "error": "Output path must end with .docx"}

        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        cleaned_yaml = preprocess_yaml_content(yaml_content)
        if not cleaned_yaml:
            logger.error("Failed to preprocess YAML content")
            return {"success": False, "error": "Failed to preprocess YAML content"}

        try:
            # First try to load as single document
            data = yaml.safe_load(cleaned_yaml)
            if data is None:
                # If empty, try loading all documents and take the first one
                documents = list(yaml.safe_load_all(cleaned_yaml))
                if documents:
                    data = documents[0]
                else:
                    raise yaml.YAMLError("Empty YAML content")
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
        if not isinstance(yaml_content, str):
            return {"success": False, "error": f"Expected string for yaml_content, got {type(yaml_content)}"}
        if not yaml_content.strip():
            return {"success": False, "error": "yaml_content is empty"}
        if not output_path.endswith('.pdf'):
            return {"success": False, "error": "Output path must end with .pdf"}

        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        cleaned_yaml = preprocess_yaml_content(yaml_content)
        if not cleaned_yaml:
            logger.error("Failed to preprocess YAML content")
            return {"success": False, "error": "Failed to preprocess YAML content"}

        try:
            # First try to load as single document
            data = yaml.safe_load(cleaned_yaml)
            if data is None:
                # If empty, try loading all documents and take the first one
                documents = list(yaml.safe_load_all(cleaned_yaml))
                if documents:
                    data = documents[0]
                else:
                    raise yaml.YAMLError("Empty YAML content")
        except yaml.YAMLError as e:
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

def process_images_for_slides(yaml_data, topic):
    """Process images for slides, adding base64-encoded data URIs."""
    try:
        if not isinstance(yaml_data, dict) or 'presentation' not in yaml_data:
            logger.error("Invalid YAML data: missing 'presentation' key")
            return yaml_data

        include_images = yaml_data.get('presentation', {}).get('include_images', True)
        if not include_images:
            logger.info("Images disabled for presentation")
            return yaml_data

        slides = yaml_data['presentation'].get('slides', [])
        if not slides:
            logger.warning("No slides found in YAML data")
            return yaml_data

        for slide in slides:
            if slide.get('needs_image', True):
                # Use slide title as primary keyword, fall back to topic
                keyword = slide.get('title', topic)
                img_url = search_pexels_image(topic, slide_title=keyword)
                if img_url:
                    img_data = download_image(img_url)
                    if img_data:
                        # Determine image format using PIL
                        img = Image.open(io.BytesIO(img_data))
                        format_map = {'JPEG': 'image/jpeg', 'PNG': 'image/png'}
                        mime_type = format_map.get(img.format, 'image/jpeg')
                        # Convert image data to base64 data URI
                        img_base64 = base64.b64encode(img_data).decode('utf-8')
                        slide['image'] = f"data:{mime_type};base64,{img_base64}"
                        logger.info(f"Added image for slide '{slide.get('title', 'Untitled')}'")
                    else:
                        logger.warning(f"Failed to download image for keyword: {keyword}")
                else:
                    logger.warning(f"No image found for keyword: {keyword}")
        return yaml_data
    except Exception as e:
        logger.error(f"Error processing images for slides: {str(e)}")
        return yaml_data

def create_html_from_yaml(yaml_content, output_path, topic, html_presentation_type='minimalist'):
    """Create HTML presentation from YAML content using the specified template."""
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
            # First try to load as single document
            data = yaml.safe_load(cleaned_yaml)
            if data is None:
                # If empty, try loading all documents and take the first one
                documents = list(yaml.safe_load_all(cleaned_yaml))
                if documents:
                    data = documents[0]
                else:
                    raise yaml.YAMLError("Empty YAML content")
        except yaml.YAMLError as e:
            logger.error(f"YAML parsing error: {str(e)}")
            return {"success": False, "error": f"YAML parsing error: {str(e)}"}

        if not data or 'presentation' not in data:
            logger.error("Invalid YAML structure: missing 'presentation' key")
            return {"success": False, "error": "Invalid YAML structure: missing 'presentation' key"}

        presentation_data = data.get('presentation', {})
        slides = presentation_data.get('slides', [])
        if not slides:
            logger.error("No slides found in YAML")
            return {"success": False, "error": "No slides found"}

        # Process images for slides
        data = process_images_for_slides(data, topic)

        # Prepare data for Jinja template
        template_data = {
            'title': presentation_data.get('title', 'AI-Generated Presentation'),
            'slides': [
                {
                    'title': slide.get('title', 'Untitled Slide'),
                    'content': slide.get('content', ''),
                    'bullets': slide.get('bullets', []),
                    'image': slide.get('image', '')
                } for slide in slides if slide.get('title') or slide.get('content') or slide.get('bullets')
            ]
        }

        # Select the appropriate template
        jinja_template = Template(HTML_TEMPLATES.get(html_presentation_type, HTML_TEMPLATES['minimalist']))
        html_content = jinja_template.render(**template_data)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        logger.info(f"HTML presentation saved to {output_path}")
        return {"success": True}
    except Exception as e:
        logger.error(f"Error in create_html_from_yaml: {str(e)}")
        return {"success": False, "error": str(e)}