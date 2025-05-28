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
from models import db, FileRecord
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
""",
    'corporate': """<!DOCTYPE html>
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
            font-family: 'Inter', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
            overflow-x: hidden;
            line-height: 1.6;
        }

        /* Presentation Container */
        .presentation {
            width: 100vw;
            height: 100vh;
            background-color: #ffffff;
            position: relative;
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
        }

        /* Header Styles */
        .header {
            background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
            color: #ffffff;
            padding: 2rem 3rem;
            text-align: center;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 10;
            border-bottom: 4px solid #0ea5e9;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
        }
        .header h1 {
            font-size: 2.5rem;
            font-weight: 600;
            letter-spacing: -0.025em;
            color: #ffffff;
        }
        .header::after {
            content: '';
            position: absolute;
            bottom: -4px;
            left: 50%;
            transform: translateX(-50%);
            width: 100px;
            height: 4px;
            background: linear-gradient(90deg, #0ea5e9, #06b6d4);
        }

        /* Slides Container */
        .slides-container {
            position: absolute;
            top: 6rem;
            width: 100%;
            height: calc(100vh - 6rem);
            overflow: hidden;
        }
        .slides {
            display: flex;
            height: 100%;
            transition: transform 0.6s cubic-bezier(0.4, 0, 0.2, 1);
        }

        /* Slide Styles */
        .slide {
            min-width: 100%;
            height: 100%;
            display: flex;
            flex-direction: row;
            align-items: center;
            padding: 3rem 4rem;
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            position: relative;
        }
        .slide::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 2px;
            background: linear-gradient(90deg, #0ea5e9, #06b6d4, #8b5cf6);
        }
        .slide-image-container {
            width: 45%;
            padding: 2rem;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .slide-image {
            max-width: 100%;
            max-height: 50vh;
            object-fit: contain;
            border-radius: 12px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
            border: 1px solid #e2e8f0;
        }
        .slide-content-container {
            width: 55%;
            padding: 2rem 3rem;
        }
        .slide-title {
            font-size: 2.25rem;
            font-weight: 700;
            color: #1e293b;
            margin-bottom: 1.5rem;
            line-height: 1.2;
            position: relative;
        }
        .slide-title::after {
            content: '';
            position: absolute;
            bottom: -8px;
            left: 0;
            width: 60px;
            height: 3px;
            background: linear-gradient(90deg, #0ea5e9, #06b6d4);
            border-radius: 2px;
        }
        .slide-content {
            font-size: 1.2rem;
            color: #475569;
            margin-bottom: 2rem;
            line-height: 1.7;
        }
        .bullets {
            list-style: none;
            padding-left: 0;
        }
        .bullets li {
            font-size: 1.1rem;
            color: #475569;
            margin-bottom: 1.2rem;
            padding: 1rem 1.5rem;
            background: #f8fafc;
            border-radius: 8px;
            border-left: 4px solid #0ea5e9;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.05);
            transition: all 0.3s ease;
            position: relative;
        }
        .bullets li::before {
            content: '▶';
            position: absolute;
            left: 0.75rem;
            top: 50%;
            transform: translateY(-50%);
            color: #0ea5e9;
            font-size: 0.8rem;
        }
        .bullets li:hover {
            transform: translateX(8px);
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            background: #ffffff;
        }

        /* Navigation Styles */
        .navigation {
            position: fixed;
            bottom: 2rem;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            gap: 1rem;
            z-index: 10;
        }
        .nav-btn {
            background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
            color: #ffffff;
            border: none;
            border-radius: 8px;
            width: 3rem;
            height: 3rem;
            font-size: 1.2rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 4px 12px rgba(30, 41, 59, 0.3);
            transition: all 0.3s ease;
        }
        .nav-btn:hover {
            background: linear-gradient(135deg, #334155 0%, #475569 100%);
            transform: translateY(-2px);
            box-shadow: 0 6px 16px rgba(30, 41, 59, 0.4);
        }
        .nav-btn:disabled {
            background: #cbd5e1;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }

        /* Slide Indicator and Progress Bar */
        .slide-indicator {
            position: fixed;
            bottom: 2rem;
            right: 2rem;
            background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
            color: #ffffff;
            padding: 0.75rem 1.25rem;
            border-radius: 20px;
            font-size: 0.9rem;
            font-weight: 600;
            box-shadow: 0 4px 12px rgba(30, 41, 59, 0.3);
        }
        .progress-bar {
            position: fixed;
            bottom: 0;
            left: 0;
            height: 4px;
            background: linear-gradient(90deg, #0ea5e9, #06b6d4, #8b5cf6);
            transition: width 0.4s ease;
            box-shadow: 0 -2px 8px rgba(14, 165, 233, 0.3);
        }

        /* Responsive Design */
        @media screen and (max-width: 1024px) {
            .slide {
                padding: 2rem 3rem;
            }
            .slide-title {
                font-size: 2rem;
            }
            .slide-content {
                font-size: 1.1rem;
            }
        }

        @media screen and (max-width: 768px) {
            .header {
                padding: 1.5rem;
            }
            .header h1 {
                font-size: 1.8rem;
            }
            .slides-container {
                top: 5rem;
                height: calc(100vh - 5rem);
            }
            .slide {
                flex-direction: column;
                padding: 2rem 1.5rem;
                text-align: center;
            }
            .slide-image-container,
            .slide-content-container {
                width: 100%;
                padding: 1rem;
            }
            .slide-image {
                max-height: 30vh;
                margin-bottom: 1rem;
            }
            .slide-title {
                font-size: 1.75rem;
            }
            .slide-content,
            .bullets li {
                font-size: 1rem;
            }
            .bullets li {
                padding: 0.8rem 1rem;
                margin-bottom: 1rem;
            }
            .navigation {
                bottom: 1rem;
            }
            .nav-btn {
                width: 2.5rem;
                height: 2.5rem;
                font-size: 1rem;
            }
            .slide-indicator {
                bottom: 1rem;
                right: 1rem;
                padding: 0.5rem 1rem;
                font-size: 0.8rem;
            }
        }

        @media screen and (max-width: 480px) {
            .header {
                padding: 1rem;
            }
            .header h1 {
                font-size: 1.5rem;
            }
            .slide {
                padding: 1.5rem 1rem;
            }
            .slide-title {
                font-size: 1.5rem;
            }
            .slide-content {
                font-size: 0.95rem;
            }
            .bullets li {
                font-size: 0.9rem;
                padding: 0.7rem 0.8rem;
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
    'executive': """<!DOCTYPE html>
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
            font-family: 'Playfair Display', 'Georgia', serif;
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%);
            overflow-x: hidden;
            line-height: 1.6;
        }

        /* Presentation Container */
        .presentation {
            width: 100vw;
            height: 100vh;
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            position: relative;
            border: 2px solid #e2e8f0;
        }

        /* Header Styles */
        .header {
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            color: #f8fafc;
            padding: 3rem;
            text-align: center;
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            z-index: 10;
            border-bottom: 6px solid #d4af37;
            box-shadow: 0 8px 32px rgba(15, 23, 42, 0.4);
        }
        .header h1 {
            font-size: 3rem;
            font-weight: 700;
            letter-spacing: 2px;
            color: #f8fafc;
            position: relative;
            text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.3);
        }
        .header::after {
            content: '';
            position: absolute;
            bottom: -6px;
            left: 50%;
            transform: translateX(-50%);
            width: 120px;
            height: 6px;
            background: linear-gradient(90deg, #d4af37, #f7931e, #d4af37);
        }

        /* Slides Container */
        .slides-container {
            position: absolute;
            top: 8rem;
            width: 100%;
            height: calc(100vh - 8rem);
            overflow: hidden;
        }
        .slides {
            display: flex;
            height: 100%;
            transition: transform 0.8s cubic-bezier(0.25, 0.46, 0.45, 0.94);
        }

        /* Slide Styles */
        .slide {
            min-width: 100%;
            height: 100%;
            display: flex;
            flex-direction: row;
            align-items: center;
            padding: 4rem 5rem;
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            position: relative;
        }
        .slide::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: 
                linear-gradient(90deg, transparent 0%, rgba(212, 175, 55, 0.05) 50%, transparent 100%),
                linear-gradient(0deg, transparent 0%, rgba(212, 175, 55, 0.03) 50%, transparent 100%);
            pointer-events: none;
        }
        .slide::after {
            content: '';
            position: absolute;
            top: 2rem;
            left: 2rem;
            right: 2rem;
            bottom: 2rem;
            border: 2px solid rgba(212, 175, 55, 0.2);
            border-radius: 8px;
            pointer-events: none;
        }
        .slide-image-container {
            width: 40%;
            padding: 2rem;
            z-index: 2;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .slide-image {
            max-width: 100%;
            max-height: 50vh;
            object-fit: contain;
            border-radius: 8px;
            box-shadow: 0 12px 32px rgba(0, 0, 0, 0.15);
            border: 2px solid #e2e8f0;
        }
        .slide-content-container {
            width: 60%;
            padding: 2rem 3rem;
            z-index: 2;
        }
        .slide-title {
            font-size: 2.75rem;
            font-weight: 700;
            color: #0f172a;
            margin-bottom: 2rem;
            letter-spacing: 1px;
            line-height: 1.2;
            position: relative;
            font-style: italic;
        }
        .slide-title::after {
            content: '';
            position: absolute;
            bottom: -10px;
            left: 0;
            width: 80px;
            height: 4px;
            background: linear-gradient(90deg, #d4af37, #f7931e);
            border-radius: 2px;
        }
        .slide-content {
            font-size: 1.3rem;
            color: #475569;
            margin-bottom: 2rem;
            line-height: 1.8;
            font-family: 'Inter', 'Segoe UI', sans-serif;
        }
        .bullets {
            list-style: none;
            padding-left: 0;
        }
        .bullets li {
            font-size: 1.2rem;
            color: #334155;
            margin-bottom: 1.5rem;
            padding: 1.2rem 2rem;
            background: linear-gradient(135deg, #f8fafc 0%, #ffffff 100%);
            border-radius: 12px;
            border-left: 5px solid #d4af37;
            box-shadow: 0 4px 16px rgba(0, 0, 0, 0.08);
            position: relative;
            transition: all 0.4s ease;
            font-family: 'Inter', 'Segoe UI', sans-serif;
            line-height: 1.6;
        }
        .bullets li::before {
            content: '◆';
            position: absolute;
            left: 1rem;
            top: 50%;
            transform: translateY(-50%);
            color: #d4af37;
            font-size: 1rem;
        }
        .bullets li:hover {
            transform: translateX(12px);
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
            background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
            border-left-width: 6px;
        }

        /* Navigation Styles */
        .navigation {
            position: fixed;
            bottom: 2.5rem;
            left: 50%;
            transform: translateX(-50%);
            display: flex;
            gap: 1.5rem;
            z-index: 10;
        }
        .nav-btn {
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            color: #f8fafc;
            border: 2px solid #d4af37;
            border-radius: 12px;
            width: 3.5rem;
            height: 3.5rem;
            font-size: 1.4rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 6px 20px rgba(15, 23, 42, 0.4);
            transition: all 0.4s ease;
            font-weight: 600;
        }
        .nav-btn:hover {
            background: linear-gradient(135deg, #1e293b 0%, #334155 100%);
            transform: translateY(-3px) scale(1.05);
            box-shadow: 0 8px 28px rgba(15, 23, 42, 0.5);
            border-color: #f7931e;
        }
        .nav-btn:disabled {
            background: #cbd5e1;
            color: #94a3b8;
            border-color: #cbd5e1;
            cursor: not-allowed;
            transform: none;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
        }

        /* Slide Indicator and Progress Bar */
        .slide-indicator {
            position: fixed;
            bottom: 2.5rem;
            right: 2.5rem;
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
            color: #f8fafc;
            border: 2px solid #d4af37;
            padding: 1rem 1.5rem;
            border-radius: 20px;
            font-size: 1rem;
            font-weight: 700;
            box-shadow: 0 6px 20px rgba(15, 23, 42, 0.4);
            letter-spacing: 1px;
        }
        .progress-bar {
            position: fixed;
            bottom: 0;
            left: 0;
            height: 6px;
            background: linear-gradient(90deg, #d4af37 0%, #f7931e 50%, #d4af37 100%);
            transition: width 0.6s ease;
            box-shadow: 0 -3px 12px rgba(212, 175, 55, 0.4);
        }

        /* Responsive Design */
        @media screen and (max-width: 1024px) {
            .slide {
                padding: 3rem 4rem;
            }
            .slide-title {
                font-size: 2.25rem;
            }
            .slide-content {
                font-size: 1.2rem;
            }
            .bullets li {
                font-size: 1.1rem;
                padding: 1rem 1.8rem;
            }
        }

        @media screen and (max-width: 768px) {
            .header {
                padding: 2rem 1.5rem;
            }
            .header h1 {
                font-size: 2.2rem;
                letter-spacing: 1px;
            }
            .slides-container {
                top: 7rem;
                height: calc(100vh - 7rem);
            }
            .slide {
                flex-direction: column;
                padding: 2rem 1.5rem;
                text-align: center;
            }
            .slide-image-container,
            .slide-content-container {
                width: 100%;
                padding: 1rem;
            }
            .slide-image {
                max-height: 35vh;
                margin-bottom: 1.5rem;
            }
            .slide-title {
                font-size: 2rem;
            }
            .slide-content {
                font-size: 1.1rem;
            }
            .bullets li {
                font-size: 1rem;
                padding: 0.9rem 1.5rem;
                margin-bottom: 1.2rem;
            }
            .navigation {
                bottom: 1.5rem;
                gap: 1rem;
            }
            .nav-btn {
                width: 3rem;
                height: 3rem;
                font-size: 1.2rem;
            }
            .slide-indicator {
                bottom: 1.5rem;
                right: 1.5rem;
                padding: 0.8rem 1.2rem;
                font-size: 0.9rem;
            }
        }

        @media screen and (max-width: 480px) {
            .header {
                padding: 1.5rem 1rem;
            }
            .header h1 {
                font-size: 1.8rem;
                letter-spacing: 0.5px;
            }
            .slides-container {
                top: 6rem;
                height: calc(100vh - 6rem);
            }
            .slide {
                padding: 1.5rem 1rem;
            }
            .slide-title {
                font-size: 1.75rem;
            }
            .slide-content {
                font-size: 1rem;
            }
            .bullets li {
                font-size: 0.95rem;
                padding: 0.8rem 1.2rem;
            }
            .nav-btn {
                width: 2.8rem;
                height: 2.8rem;
                font-size: 1.1rem;
            }
            .slide-indicator {
                padding: 0.7rem 1rem;
                font-size: 0.8rem;
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
    "elegant" : """<html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{{ title }}</title>
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Space+Grotesk:wght@400;500;600;700&display=swap');
            
            /* Reset and Base Styles */
            * {
                margin: 0;
                padding: 0;
                box-sizing: border-box;
            }
            
            body {
                font-family: 'Inter', system-ui, -apple-system, sans-serif;
                background: linear-gradient(135deg, #0a0a0a 0%, #1e1e2e 50%, #2d1b69 100%);
                overflow-x: hidden;
                line-height: 1.6;
                color: #ffffff;
                min-height: 100vh;
                min-height: 100dvh;
            }

            /* Animated Background Particles */
            .bg-particles {
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                pointer-events: none;
                z-index: 1;
            }
            
            .particle {
                position: absolute;
                width: 2px;
                height: 2px;
                background: rgba(139, 92, 246, 0.6);
                border-radius: 50%;
                animation: float 6s ease-in-out infinite;
            }
            
            @keyframes float {
                0%, 100% { transform: translateY(0px) rotate(0deg); opacity: 0.4; }
                50% { transform: translateY(-20px) rotate(180deg); opacity: 0.8; }
            }

            /* Presentation Container */
            .presentation {
                width: 100vw;
                height: 100vh;
                height: 100dvh;
                background: linear-gradient(135deg, rgba(15, 15, 35, 0.95) 0%, rgba(30, 30, 60, 0.95) 100%);
                backdrop-filter: blur(20px);
                position: relative;
                border: 1px solid rgba(139, 92, 246, 0.2);
                overflow: hidden;
                display: flex;
                flex-direction: column;
            }

            /* Header Styles */
            .header {
                background: linear-gradient(135deg, rgba(15, 15, 35, 0.9) 0%, rgba(30, 30, 60, 0.9) 100%);
                backdrop-filter: blur(30px);
                color: #ffffff;
                padding: 2.5rem 3rem;
                text-align: center;
                border-bottom: 2px solid rgba(139, 92, 246, 0.3);
                box-shadow: 0 8px 32px rgba(0, 0, 0, 0.4);
                flex-shrink: 0;
                z-index: 10;
            }
            
            .header h1 {
                font-family: 'Space Grotesk', sans-serif;
                font-size: 2.8rem;
                font-weight: 700;
                letter-spacing: -0.5px;
                background: linear-gradient(135deg, #8b5cf6 0%, #06b6d4 50%, #10b981 100%);
                background-clip: text;
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                position: relative;
                text-shadow: none;
                animation: glow 3s ease-in-out infinite alternate;
                word-wrap: break-word;
                hyphens: auto;
            }
            
            @keyframes glow {
                from { filter: drop-shadow(0 0 20px rgba(139, 92, 246, 0.3)); }
                to { filter: drop-shadow(0 0 30px rgba(6, 182, 212, 0.4)); }
            }
            
            .header::after {
                content: '';
                position: absolute;
                bottom: 0;
                left: 50%;
                transform: translateX(-50%);
                width: 200px;
                height: 2px;
                background: linear-gradient(90deg, #8b5cf6, #06b6d4, #10b981);
                border-radius: 2px;
                animation: pulse 2s ease-in-out infinite;
            }
            
            @keyframes pulse {
                0%, 100% { opacity: 0.8; transform: translateX(-50%) scaleX(1); }
                50% { opacity: 1; transform: translateX(-50%) scaleX(1.1); }
            }

            /* Main Content Area */
            .main-content {
                flex: 1;
                display: flex;
                flex-direction: column;
                overflow: hidden;
            }

            /* Slides Container */
            .slides-container {
                flex: 1;
                overflow: hidden;
                position: relative;
            }
            
            .slides {
                display: flex;
                height: 100%;
                transition: transform 0.8s cubic-bezier(0.25, 0.46, 0.45, 0.94);
            }

            /* Slide Styles */
            .slide {
                min-width: 100%;
                height: 100%;
                display: flex;
                flex-direction: row;
                align-items: flex-start;
                padding: 2rem 3rem;
                background: linear-gradient(135deg, rgba(15, 15, 35, 0.8) 0%, rgba(30, 30, 60, 0.6) 100%);
                backdrop-filter: blur(20px);
                position: relative;
                overflow-y: auto;
                scrollbar-width: thin;
                scrollbar-color: rgba(139, 92, 246, 0.5) transparent;
            }
            
            .slide::-webkit-scrollbar {
                width: 6px;
            }
            
            .slide::-webkit-scrollbar-track {
                background: transparent;
            }
            
            .slide::-webkit-scrollbar-thumb {
                background: rgba(139, 92, 246, 0.5);
                border-radius: 3px;
            }
            
            .slide::-webkit-scrollbar-thumb:hover {
                background: rgba(139, 92, 246, 0.7);
            }
            
            .slide::before {
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: 
                    radial-gradient(circle at 20% 80%, rgba(139, 92, 246, 0.1) 0%, transparent 50%),
                    radial-gradient(circle at 80% 20%, rgba(6, 182, 212, 0.1) 0%, transparent 50%),
                    radial-gradient(circle at 40% 40%, rgba(16, 185, 129, 0.05) 0%, transparent 50%);
                pointer-events: none;
            }
            
            .slide::after {
                content: '';
                position: absolute;
                top: 1rem;
                left: 1rem;
                right: 1rem;
                bottom: 1rem;
                border: 1px solid rgba(139, 92, 246, 0.2);
                border-radius: 16px;
                pointer-events: none;
                background: linear-gradient(135deg, 
                    rgba(139, 92, 246, 0.03) 0%, 
                    rgba(6, 182, 212, 0.03) 50%, 
                    rgba(16, 185, 129, 0.03) 100%);
            }
            
            .slide-image-container {
                width: 40%;
                padding: 1rem 2rem 1rem 1rem;
                z-index: 2;
                display: flex;
                justify-content: center;
                align-items: flex-start;
                flex-shrink: 0;
            }
            
            .slide-image {
                max-width: 100%;
                max-height: 60vh;
                object-fit: contain;
                border-radius: 16px;
                box-shadow: 
                    0 20px 40px rgba(0, 0, 0, 0.3),
                    0 0 0 1px rgba(139, 92, 246, 0.2);
                transition: transform 0.3s ease, box-shadow 0.3s ease;
            }
            
            .slide-image:hover {
                transform: scale(1.02);
                box-shadow: 
                    0 25px 50px rgba(0, 0, 0, 0.4),
                    0 0 0 1px rgba(139, 92, 246, 0.4),
                    0 0 30px rgba(139, 92, 246, 0.2);
            }
            
            .slide-content-container {
                width: 60%;
                padding: 1rem 1rem 1rem 2rem;
                z-index: 2;
                flex: 1;
                overflow-y: auto;
            }
            
            .slide-title {
                font-family: 'Space Grotesk', sans-serif;
                font-size: 2.5rem;
                font-weight: 700;
                background: linear-gradient(135deg, #ffffff 0%, #e2e8f0 100%);
                background-clip: text;
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                margin-bottom: 2rem;
                letter-spacing: -0.5px;
                line-height: 1.2;
                position: relative;
                word-wrap: break-word;
                hyphens: auto;
            }
            
            .slide-title::after {
                content: '';
                position: absolute;
                bottom: -12px;
                left: 0;
                width: 60px;
                height: 3px;
                background: linear-gradient(90deg, #8b5cf6, #06b6d4);
                border-radius: 3px;
                animation: expand 0.8s ease-out;
            }
            
            @keyframes expand {
                from { width: 0; opacity: 0; }
                to { width: 60px; opacity: 1; }
            }
            
            .slide-content {
                font-size: 1.2rem;
                color: #cbd5e1;
                margin-bottom: 2rem;
                line-height: 1.7;
                font-weight: 400;
                word-wrap: break-word;
                hyphens: auto;
            }
            
            .bullets {
                list-style: none;
                padding-left: 0;
            }
            
            .bullets li {
                font-size: 1.1rem;
                color: #e2e8f0;
                margin-bottom: 1.2rem;
                padding: 1.2rem 2rem 1.2rem 3rem;
                background: linear-gradient(135deg, 
                    rgba(139, 92, 246, 0.1) 0%, 
                    rgba(6, 182, 212, 0.05) 100%);
                backdrop-filter: blur(10px);
                border-radius: 12px;
                border: 1px solid rgba(139, 92, 246, 0.2);
                box-shadow: 0 4px 16px rgba(0, 0, 0, 0.2);
                position: relative;
                transition: all 0.3s ease;
                line-height: 1.6;
                word-wrap: break-word;
                hyphens: auto;
            }
            
            .bullets li::before {
                content: '→';
                position: absolute;
                left: 1.2rem;
                top: 50%;
                transform: translateY(-50%);
                color: #8b5cf6;
                font-size: 1.2rem;
                font-weight: 600;
            }
            
            .bullets li:hover {
                transform: translateX(8px);
                background: linear-gradient(135deg, 
                    rgba(139, 92, 246, 0.15) 0%, 
                    rgba(6, 182, 212, 0.1) 100%);
                border-color: rgba(139, 92, 246, 0.4);
                box-shadow: 0 8px 24px rgba(139, 92, 246, 0.2);
            }

            /* Bottom Navigation Bar */
            .bottom-navbar {
                flex-shrink: 0;
                background: linear-gradient(135deg, rgba(15, 15, 35, 0.95) 0%, rgba(30, 30, 60, 0.95) 100%);
                backdrop-filter: blur(30px);
                border-top: 2px solid rgba(139, 92, 246, 0.3);
                padding: 1.5rem 2rem;
                display: flex;
                justify-content: space-between;
                align-items: center;
                box-shadow: 0 -8px 32px rgba(0, 0, 0, 0.4);
                z-index: 10;
            }
            
            .navigation {
                display: flex;
                gap: 1rem;
            }
            
            .nav-btn {
                background: linear-gradient(135deg, 
                    rgba(139, 92, 246, 0.2) 0%, 
                    rgba(6, 182, 212, 0.2) 100%);
                backdrop-filter: blur(20px);
                color: #ffffff;
                border: 1px solid rgba(139, 92, 246, 0.4);
                border-radius: 12px;
                padding: 0.8rem 1.5rem;
                font-size: 1rem;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 0.5rem;
                box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3);
                transition: all 0.3s ease;
                font-weight: 600;
                min-width: 100px;
                touch-action: manipulation;
            }
            
            .nav-btn:hover {
                background: linear-gradient(135deg, 
                    rgba(139, 92, 246, 0.3) 0%, 
                    rgba(6, 182, 212, 0.3) 100%);
                transform: translateY(-2px);
                box-shadow: 0 10px 30px rgba(139, 92, 246, 0.3);
                border-color: rgba(139, 92, 246, 0.6);
            }
            
            .nav-btn:active {
                transform: translateY(0);
            }
            
            .nav-btn:disabled {
                background: rgba(71, 85, 105, 0.3);
                color: #64748b;
                border-color: rgba(71, 85, 105, 0.3);
                cursor: not-allowed;
                transform: none;
                box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
            }

            /* Slide Indicator and Progress Bar */
            .slide-indicator {
                background: linear-gradient(135deg, 
                    rgba(139, 92, 246, 0.2) 0%, 
                    rgba(6, 182, 212, 0.2) 100%);
                backdrop-filter: blur(20px);
                color: #ffffff;
                border: 1px solid rgba(139, 92, 246, 0.4);
                padding: 0.8rem 1.5rem;
                border-radius: 25px;
                font-size: 0.9rem;
                font-weight: 600;
                box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3);
                letter-spacing: 0.5px;
            }
            
            .progress-bar {
                position: absolute;
                bottom: 0;
                left: 0;
                height: 3px;
                background: linear-gradient(90deg, 
                    #8b5cf6 0%, 
                    #06b6d4 50%, 
                    #10b981 100%);
                transition: width 0.6s ease;
                box-shadow: 0 0 20px rgba(139, 92, 246, 0.5);
            }

            /* Enhanced Mobile Responsive Design */
            
            /* Large tablets and small desktops */
            @media screen and (max-width: 1200px) {
                .header {
                    padding: 2.2rem 2.5rem;
                }
                .header h1 {
                    font-size: 2.5rem;
                }
                .slide {
                    padding: 1.5rem 2.5rem;
                }
                .slide-title {
                    font-size: 2.2rem;
                }
                .slide-content {
                    font-size: 1.15rem;
                }
                .bottom-navbar {
                    padding: 1.2rem 2rem;
                }
            }
            
            /* Standard tablets */
            @media screen and (max-width: 1024px) {
                .header {
                    padding: 2rem 2rem;
                }
                .header h1 {
                    font-size: 2.3rem;
                }
                .slide {
                    padding: 1.5rem 2rem;
                }
                .slide-title {
                    font-size: 2rem;
                }
                .slide-content {
                    font-size: 1.1rem;
                }
                .bullets li {
                    font-size: 1rem;
                    padding: 1rem 1.5rem 1rem 2.5rem;
                    margin-bottom: 1rem;
                }
                .nav-btn {
                    padding: 0.7rem 1.2rem;
                    font-size: 0.95rem;
                    min-width: 90px;
                }
                .bottom-navbar {
                    padding: 1rem 1.5rem;
                }
            }

            /* Large mobile phones and small tablets */
            @media screen and (max-width: 768px) {
                .header {
                    padding: 1.8rem 1.5rem;
                }
                .header h1 {
                    font-size: 2rem;
                    line-height: 1.3;
                }
                .header::after {
                    width: 150px;
                }
                .slide {
                    flex-direction: column;
                    padding: 1.5rem;
                    align-items: center;
                }
                .slide::after {
                    top: 0.5rem;
                    left: 0.5rem;
                    right: 0.5rem;
                    bottom: 0.5rem;
                }
                .slide-image-container,
                .slide-content-container {
                    width: 100%;
                    padding: 1rem 0.5rem;
                }
                .slide-image-container {
                    order: 1;
                    flex-shrink: 0;
                }
                .slide-content-container {
                    order: 2;
                    flex: 1;
                    text-align: center;
                }
                .slide-image {
                    max-height: 30vh;
                    margin-bottom: 1rem;
                }
                .slide-title {
                    font-size: 1.7rem;
                    margin-bottom: 1.5rem;
                    line-height: 1.3;
                }
                .slide-content {
                    font-size: 1rem;
                    margin-bottom: 1.5rem;
                    line-height: 1.6;
                }
                .bullets li {
                    font-size: 0.95rem;
                    padding: 0.9rem 1.3rem 0.9rem 2.2rem;
                    margin-bottom: 0.8rem;
                    line-height: 1.5;
                }
                .bullets li::before {
                    left: 1rem;
                    font-size: 1.1rem;
                }
                .bottom-navbar {
                    padding: 1rem 1rem;
                }
                .nav-btn {
                    padding: 0.6rem 1rem;
                    font-size: 0.9rem;
                    min-width: 80px;
                }
                .slide-indicator {
                    padding: 0.6rem 1rem;
                    font-size: 0.8rem;
                }
            }

            /* Standard mobile phones */
            @media screen and (max-width: 480px) {
                .header {
                    padding: 1.5rem 1rem;
                }
                .header h1 {
                    font-size: 1.6rem;
                    line-height: 1.4;
                }
                .header::after {
                    width: 120px;
                }
                .slide {
                    padding: 1rem;
                }
                .slide::after {
                    top: 0.3rem;
                    left: 0.3rem;
                    right: 0.3rem;
                    bottom: 0.3rem;
                }
                .slide-image-container,
                .slide-content-container {
                    padding: 0.8rem 0.3rem;
                }
                .slide-image {
                    max-height: 25vh;
                }
                .slide-title {
                    font-size: 1.4rem;
                    margin-bottom: 1rem;
                    line-height: 1.3;
                }
                .slide-content {
                    font-size: 0.9rem;
                    margin-bottom: 1rem;
                    line-height: 1.5;
                }
                .bullets li {
                    font-size: 0.85rem;
                    padding: 0.7rem 1rem 0.7rem 1.8rem;
                    margin-bottom: 0.6rem;
                    line-height: 1.4;
                }
                .bullets li::before {
                    left: 0.8rem;
                    font-size: 1rem;
                }
                .bottom-navbar {
                    padding: 0.8rem 0.8rem;
                    flex-direction: column;
                    gap: 0.8rem;
                }
                .navigation {
                    gap: 0.8rem;
                }
                .nav-btn {
                    padding: 0.5rem 0.8rem;
                    font-size: 0.8rem;
                    min-width: 70px;
                }
                .slide-indicator {
                    padding: 0.5rem 0.8rem;
                    font-size: 0.75rem;
                }
                .progress-bar {
                    height: 2px;
                }
            }

            /* Small mobile phones */
            @media screen and (max-width: 360px) {
                .header {
                    padding: 1.2rem 0.8rem;
                }
                .header h1 {
                    font-size: 1.4rem;
                }
                .header::after {
                    width: 100px;
                }
                .slide {
                    padding: 0.8rem;
                }
                .slide-title {
                    font-size: 1.2rem;
                }
                .slide-content {
                    font-size: 0.85rem;
                }
                .bullets li {
                    font-size: 0.8rem;
                    padding: 0.6rem 0.8rem 0.6rem 1.6rem;
                }
                .bullets li::before {
                    left: 0.6rem;
                    font-size: 0.9rem;
                }
                .nav-btn {
                    padding: 0.4rem 0.6rem;
                    font-size: 0.75rem;
                    min-width: 60px;
                }
                .slide-indicator {
                    padding: 0.4rem 0.6rem;
                    font-size: 0.7rem;
                }
            }

            /* Landscape orientation optimizations for mobile */
            @media screen and (max-height: 500px) and (orientation: landscape) {
                .header {
                    padding: 0.8rem 1.5rem;
                }
                .header h1 {
                    font-size: 1.5rem;
                }
                .slide {
                    flex-direction: row;
                    padding: 0.8rem 1.5rem;
                }
                .slide-image-container {
                    width: 35%;
                    order: 1;
                }
                .slide-content-container {
                    width: 65%;
                    order: 2;
                    text-align: left;
                }
                .slide-image {
                    max-height: 40vh;
                }
                .slide-title {
                    font-size: 1.3rem;
                    text-align: left;
                }
                .slide-content {
                    font-size: 0.9rem;
                    text-align: left;
                }
                .bullets li {
                    font-size: 0.85rem;
                    padding: 0.6rem 1rem 0.6rem 1.8rem;
                }
                .bottom-navbar {
                    padding: 0.6rem 1rem;
                    flex-direction: row;
                    gap: 1rem;
                }
            }

            /* Touch-friendly enhancements */
            @media (hover: none) and (pointer: coarse) {
                .nav-btn:hover {
                    transform: none;
                    background: linear-gradient(135deg, 
                        rgba(139, 92, 246, 0.2) 0%, 
                        rgba(6, 182, 212, 0.2) 100%);
                    box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3);
                }
                
                .bullets li:hover {
                    transform: none;
                    background: linear-gradient(135deg, 
                        rgba(139, 92, 246, 0.1) 0%, 
                        rgba(6, 182, 212, 0.05) 100%);
                }
                
                .slide-image:hover {
                    transform: none;
                    box-shadow: 
                        0 20px 40px rgba(0, 0, 0, 0.3),
                        0 0 0 1px rgba(139, 92, 246, 0.2);
                }
            }

            /* High DPI displays */
            @media (-webkit-min-device-pixel-ratio: 2), (min-resolution: 192dpi) {
                .particle {
                    width: 1px;
                    height: 1px;
                }
            }

            /* Reduce motion for accessibility */
            @media (prefers-reduced-motion: reduce) {
                .particle {
                    animation: none;
                }
                .header h1 {
                    animation: none;
                }
                .header::after {
                    animation: none;
                }
                .slide-title::after {
                    animation: none;
                }
                .slides {
                    transition: transform 0.3s ease;
                }
                .nav-btn {
                    transition: none;
                }
                .bullets li {
                    transition: none;
                }
            }
        </style>
    </head>
    <body>
        <div class="bg-particles" id="particles"></div>
        
        <div class="presentation">
            <header class="header">
                <h1>{{ title }}</h1>
            </header>
            
            <div class="main-content">
                <div class="slides-container">
                    <div class="slides" id="slides">
                        {% for slide in slides %}
                        <section class="slide">
                            {% if slide.image %}
                            <div class="slide-image-container">
                                <img class="slide-image" src="{{ slide.image }}" alt="{{ slide.title }}" loading="lazy">
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
                
                <div class="bottom-navbar">
                    <nav class="navigation">
                        <button class="nav-btn" id="prevBtn" aria-label="Previous Slide">← Previous</button>
                        <button class="nav-btn" id="nextBtn" aria-label="Next Slide">Next →</button>
                    </nav>
                    <div class="slide-indicator" id="slideIndicator" aria-live="polite">1 / {{ slides|length }}</div>
                </div>
            </div>
            
            <div class="progress-bar" id="progressBar" role="progressbar"></div>
        </div>

        <script>
            // Particle animation with performance optimization
            function createParticles() {
                const particlesContainer = document.getElementById('particles');
                // Reduce particles on mobile for better performance
                const isMobile = window.innerWidth <= 768;
                const particleCount = isMobile ? 25 : 50;
                
                for (let i = 0; i < particleCount; i++) {
                    const particle = document.createElement('div');
                    particle.className = 'particle';
                    particle.style.left = Math.random() * 100 + '%';
                    particle.style.top = Math.random() * 100 + '%';
                    particle.style.animationDelay = Math.random() * 6 + 's';
                    particle.style.animationDuration = (Math.random() * 3 + 3) + 's';
                    particlesContainer.appendChild(particle);
                }
            }

            document.addEventListener('DOMContentLoaded', () => {
                createParticles();
                
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
    </html>""",
    "refined" : """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Playfair+Display:ital,wght@0,400;0,500;0,600;1,400&display=swap');
        
        /* Reset and Base Styles */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 30%, #0f3460 70%, #e94560 100%);
            overflow: hidden;
            line-height: 1.6;
            color: #2c3e50;
            position: relative;
            height: 100vh;
        }

        /* Animated geometric shapes */
        .bg-shapes {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            pointer-events: none;
            z-index: 1;
            overflow: hidden;
        }
        
        .shape {
            position: absolute;
            opacity: 0.03;
            animation: rotate 25s linear infinite;
        }
        
        .shape:nth-child(1) {
            top: 15%;
            left: 85%;
            width: 120px;
            height: 120px;
            background: linear-gradient(45deg, #667eea, #764ba2);
            clip-path: polygon(50% 0%, 0% 100%, 100% 100%);
            animation-delay: 0s;
        }
        
        .shape:nth-child(2) {
            top: 65%;
            left: 8%;
            width: 100px;
            height: 100px;
            background: linear-gradient(45deg, #f093fb, #f5576c);
            border-radius: 50%;
            animation-delay: -8s;
        }
        
        .shape:nth-child(3) {
            top: 25%;
            left: 15%;
            width: 80px;
            height: 80px;
            background: linear-gradient(45deg, #4facfe, #00f2fe);
            transform: rotate(45deg);
            animation-delay: -16s;
        }
        
        @keyframes rotate {
            from { transform: rotate(0deg) scale(1); }
            50% { transform: rotate(180deg) scale(1.1); }
            to { transform: rotate(360deg) scale(1); }
        }

        /* Presentation Container */
        .presentation {
            width: 100vw;
            height: 100vh;
            background: linear-gradient(135deg, 
                rgba(255, 255, 255, 0.98) 0%, 
                rgba(248, 250, 252, 0.96) 50%,
                rgba(241, 245, 249, 0.95) 100%);
            position: relative;
            box-shadow: inset 0 0 100px rgba(30, 41, 59, 0.08);
            display: flex;
            flex-direction: column;
        }

        /* Header Styles */
        .header {
            background: linear-gradient(135deg, 
                rgba(30, 41, 59, 0.95) 0%, 
                rgba(51, 65, 85, 0.95) 50%,
                rgba(71, 85, 105, 0.95) 100%);
            color: #ffffff;
            padding: 2rem 3rem;
            text-align: center;
            border-bottom: 3px solid #3b82f6;
            box-shadow: 0 8px 32px rgba(30, 41, 59, 0.12);
            backdrop-filter: blur(20px);
            flex-shrink: 0;
        }
        
        .header h1 {
            font-family: 'Playfair Display', serif;
            font-size: 2.5rem;
            font-weight: 600;
            letter-spacing: -0.5px;
            color: #ffffff;
            position: relative;
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        
        .header::before {
            content: '';
            position: absolute;
            top: 1rem;
            left: 1rem;
            right: 1rem;
            bottom: 1rem;
            border: 1px solid rgba(59, 130, 246, 0.3);
            border-radius: 12px;
            pointer-events: none;
        }
        
        .header::after {
            content: '';
            position: absolute;
            bottom: -3px;
            left: 50%;
            transform: translateX(-50%);
            width: 100px;
            height: 3px;
            background: linear-gradient(90deg, #3b82f6, #8b5cf6, #3b82f6);
            border-radius: 2px;
        }

        /* Slides Container - Now with proper scrolling area */
        .slides-container {
            flex: 1;
            overflow: hidden;
            position: relative;
        }
        
        .slides {
            display: flex;
            height: 100%;
            transition: transform 0.8s cubic-bezier(0.4, 0, 0.2, 1);
        }

        /* Slide Styles - Now properly contained */
        .slide {
            min-width: 100%;
            height: 100%;
            display: flex;
            flex-direction: row;
            align-items: flex-start;
            padding: 2rem 3rem;
            background: linear-gradient(135deg, 
                rgba(255, 255, 255, 0.95) 0%, 
                rgba(248, 250, 252, 0.9) 100%);
            position: relative;
            overflow-y: auto;
        }
        
        .slide::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: 
                radial-gradient(circle at 25% 75%, rgba(59, 130, 246, 0.04) 0%, transparent 50%),
                radial-gradient(circle at 75% 25%, rgba(139, 92, 246, 0.03) 0%, transparent 50%);
            pointer-events: none;
        }
        
        .slide::after {
            content: '';
            position: absolute;
            top: 1rem;
            left: 1rem;
            right: 1rem;
            bottom: 5rem;
            border: 1px solid rgba(59, 130, 246, 0.1);
            border-radius: 16px;
            pointer-events: none;
        }
        
        .slide-image-container {
            width: 45%;
            padding: 1.5rem;
            z-index: 2;
            display: flex;
            justify-content: center;
            align-items: flex-start;
            flex-shrink: 0;
        }
        
        .slide-image {
            max-width: 100%;
            max-height: 50vh;
            object-fit: contain;
            border-radius: 16px;
            box-shadow: 
                0 20px 40px rgba(30, 41, 59, 0.12),
                0 0 0 1px rgba(59, 130, 246, 0.1);
            transition: all 0.3s ease;
        }
        
        .slide-image:hover {
            transform: scale(1.02);
            box-shadow: 
                0 25px 50px rgba(30, 41, 59, 0.16),
                0 0 0 1px rgba(59, 130, 246, 0.2);
        }
        
        .slide-content-container {
            width: 55%;
            padding: 1.5rem 2rem;
            z-index: 2;
            overflow-y: auto;
            max-height: 100%;
        }
        
        .slide-title {
            font-family: 'Playfair Display', serif;
            font-size: 2.2rem;
            font-weight: 600;
            color: #1e293b;
            margin-bottom: 1.5rem;
            letter-spacing: -0.5px;
            line-height: 1.2;
            position: relative;
        }
        
        .slide-title::after {
            content: '';
            position: absolute;
            bottom: -12px;
            left: 0;
            width: 60px;
            height: 3px;
            background: linear-gradient(90deg, #3b82f6, #8b5cf6);
            border-radius: 2px;
        }
        
        .slide-content {
            font-size: 1.1rem;
            color: #475569;
            margin-bottom: 2rem;
            line-height: 1.7;
            font-weight: 400;
        }
        
        .bullets {
            list-style: none;
            padding-left: 0;
            margin-bottom: 2rem;
        }
        
        .bullets li {
            font-size: 1rem;
            color: #334155;
            margin-bottom: 1rem;
            padding: 1.25rem 1.5rem 1.25rem 2.5rem;
            background: linear-gradient(135deg, 
                rgba(59, 130, 246, 0.06) 0%, 
                rgba(255, 255, 255, 0.8) 100%);
            border-radius: 12px;
            border-left: 3px solid #3b82f6;
            box-shadow: 0 4px 12px rgba(30, 41, 59, 0.08);
            position: relative;
            transition: all 0.3s ease;
            line-height: 1.6;
            backdrop-filter: blur(10px);
        }
        
        .bullets li::before {
            content: '•';
            position: absolute;
            left: 1rem;
            top: 50%;
            transform: translateY(-50%);
            color: #3b82f6;
            font-size: 1.2rem;
            font-weight: bold;
        }
        
        .bullets li:hover {
            transform: translateX(8px);
            background: linear-gradient(135deg, 
                rgba(59, 130, 246, 0.1) 0%, 
                rgba(255, 255, 255, 0.9) 100%);
            border-left-color: #8b5cf6;
            box-shadow: 0 6px 20px rgba(30, 41, 59, 0.12);
        }

        /* Bottom Navigation Bar */
        .bottom-navbar {
            background: linear-gradient(135deg, 
                rgba(30, 41, 59, 0.95) 0%, 
                rgba(51, 65, 85, 0.95) 100%);
            border-top: 1px solid rgba(59, 130, 246, 0.3);
            padding: 1rem 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            backdrop-filter: blur(20px);
            box-shadow: 0 -4px 16px rgba(30, 41, 59, 0.1);
            flex-shrink: 0;
        }
        
        .navigation {
            display: flex;
            gap: 1rem;
            align-items: center;
        }
        
        .nav-btn {
            background: linear-gradient(135deg, 
                rgba(59, 130, 246, 0.9) 0%, 
                rgba(139, 92, 246, 0.9) 100%);
            color: #ffffff;
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 8px;
            padding: 0.75rem 1.5rem;
            font-size: 0.9rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 0.5rem;
            box-shadow: 0 2px 8px rgba(59, 130, 246, 0.3);
            transition: all 0.3s ease;
            font-weight: 500;
            backdrop-filter: blur(20px);
            min-width: 44px;
            min-height: 44px;
        }
        
        .nav-btn:hover:not(:disabled) {
            background: linear-gradient(135deg, 
                rgba(59, 130, 246, 1) 0%, 
                rgba(139, 92, 246, 1) 100%);
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(59, 130, 246, 0.4);
        }
        
        .nav-btn:active:not(:disabled) {
            transform: translateY(0);
        }
        
        .nav-btn:disabled {
            background: linear-gradient(135deg, 
                rgba(148, 163, 184, 0.6) 0%, 
                rgba(203, 213, 225, 0.6) 100%);
            color: #94a3b8;
            border-color: rgba(148, 163, 184, 0.3);
            cursor: not-allowed;
            transform: none;
            box-shadow: 0 1px 4px rgba(0, 0, 0, 0.1);
        }

        /* Slide Indicator and Progress Bar */
        .slide-indicator {
            background: linear-gradient(135deg, 
                rgba(59, 130, 246, 0.1) 0%, 
                rgba(139, 92, 246, 0.1) 100%);
            color: #ffffff;
            border: 1px solid rgba(59, 130, 246, 0.3);
            padding: 0.5rem 1rem;
            border-radius: 20px;
            font-size: 0.9rem;
            font-weight: 500;
            backdrop-filter: blur(20px);
        }
        
        .progress-bar {
            position: absolute;
            bottom: 0;
            left: 0;
            height: 3px;
            background: linear-gradient(90deg, 
                #3b82f6 0%, 
                #8b5cf6 50%, 
                #3b82f6 100%);
            transition: width 0.8s ease;
            box-shadow: 0 -1px 8px rgba(59, 130, 246, 0.3);
        }

        /* Enhanced Mobile Responsiveness */
        @media screen and (max-width: 1024px) {
            .header {
                padding: 1.5rem 2rem;
            }
            .header h1 {
                font-size: 2.2rem;
            }
            .slide {
                padding: 1.5rem 2rem;
            }
            .slide-title {
                font-size: 2rem;
            }
            .slide-content {
                font-size: 1rem;
            }
            .bullets li {
                font-size: 0.95rem;
                padding: 1.1rem 1.3rem 1.1rem 2.3rem;
            }
        }

        @media screen and (max-width: 768px) {
            .header {
                padding: 1.25rem 1.5rem;
            }
            .header h1 {
                font-size: 1.8rem;
                letter-spacing: -0.25px;
            }
            .slide {
                flex-direction: column;
                padding: 1.5rem;
                align-items: center;
            }
            .slide::after {
                bottom: 1rem;
            }
            .slide-image-container,
            .slide-content-container {
                width: 100%;
                padding: 1rem;
            }
            .slide-image-container {
                order: 1;
            }
            .slide-content-container {
                order: 2;
            }
            .slide-image {
                max-height: 30vh;
                margin-bottom: 1rem;
            }
            .slide-title {
                font-size: 1.6rem;
                text-align: center;
                margin-bottom: 1rem;
            }
            .slide-content {
                font-size: 0.95rem;
                text-align: left;
                margin-bottom: 1.5rem;
            }
            .bullets li {
                font-size: 0.9rem;
                padding: 1rem 1.2rem 1rem 2rem;
                margin-bottom: 0.8rem;
            }
            .bottom-navbar {
                padding: 0.75rem 1.5rem;
            }
            .nav-btn {
                padding: 0.6rem 1.2rem;
                font-size: 0.85rem;
            }
            .slide-indicator {
                padding: 0.4rem 0.8rem;
                font-size: 0.8rem;
            }
        }

        @media screen and (max-width: 640px) {
            .header {
                padding: 1rem;
            }
            .header h1 {
                font-size: 1.6rem;
            }
            .slide {
                padding: 1rem;
            }
            .slide-title {
                font-size: 1.4rem;
            }
            .slide-content {
                font-size: 0.9rem;
            }
            .bullets li {
                font-size: 0.85rem;
                padding: 0.9rem 1rem 0.9rem 1.8rem;
            }
            .bullets li::before {
                left: 0.6rem;
            }
            .bottom-navbar {
                padding: 0.6rem 1rem;
                flex-direction: column;
                gap: 0.75rem;
            }
            .navigation {
                gap: 0.75rem;
            }
            .nav-btn {
                padding: 0.5rem 1rem;
                font-size: 0.8rem;
            }
        }

        @media screen and (max-width: 480px) {
            body {
                font-size: 14px;
            }
            .header {
                padding: 0.8rem;
            }
            .header h1 {
                font-size: 1.4rem;
                letter-spacing: 0;
            }
            .slide {
                padding: 0.8rem;
            }
            .slide::after {
                top: 0.8rem;
                left: 0.8rem;
                right: 0.8rem;
                bottom: 0.8rem;
            }
            .slide-title {
                font-size: 1.2rem;
                margin-bottom: 0.8rem;
            }
            .slide-content {
                font-size: 0.85rem;
                margin-bottom: 1.2rem;
            }
            .bullets li {
                font-size: 0.8rem;
                padding: 0.8rem 0.9rem 0.8rem 1.6rem;
                margin-bottom: 0.7rem;
            }
            .bottom-navbar {
                padding: 0.5rem 0.8rem;
            }
            .nav-btn {
                padding: 0.4rem 0.8rem;
                font-size: 0.75rem;
            }
            .slide-indicator {
                padding: 0.3rem 0.6rem;
                font-size: 0.7rem;
            }
        }

        /* Touch-friendly enhancements for mobile */
        @media (hover: none) and (pointer: coarse) {
            .nav-btn {
                min-height: 44px;
                min-width: 44px;
            }
            .bullets li:hover {
                transform: none;
            }
            .slide-image:hover {
                transform: none;
            }
        }

        /* High DPI displays */
        @media (-webkit-min-device-pixel-ratio: 2), (min-resolution: 192dpi) {
            .slide-image {
                image-rendering: -webkit-optimize-contrast;
                image-rendering: crisp-edges;
            }
        }

        /* Landscape orientation on mobile */
        @media screen and (max-width: 768px) and (orientation: landscape) {
            .slide {
                flex-direction: row;
                padding: 1rem;
            }
            .slide-image-container,
            .slide-content-container {
                width: 50%;
                padding: 0.5rem;
            }
            .slide-image {
                max-height: 40vh;
            }
            .slide-title {
                font-size: 1.3rem;
                text-align: left;
            }
            .bullets li {
                padding: 0.7rem 1rem 0.7rem 1.8rem;
                margin-bottom: 0.6rem;
            }
            .bottom-navbar {
                flex-direction: row;
                padding: 0.5rem 1rem;
            }
        }

        /* Custom scrollbar for webkit browsers */
        .slide-content-container::-webkit-scrollbar {
            width: 6px;
        }
        
        .slide-content-container::-webkit-scrollbar-track {
            background: rgba(59, 130, 246, 0.1);
            border-radius: 3px;
        }
        
        .slide-content-container::-webkit-scrollbar-thumb {
            background: linear-gradient(135deg, #3b82f6, #8b5cf6);
            border-radius: 3px;
        }
        
        .slide-content-container::-webkit-scrollbar-thumb:hover {
            background: linear-gradient(135deg, #2563eb, #7c3aed);
        }
    </style>
</head>
<body>
    <div class="bg-shapes">
        <div class="shape"></div>
        <div class="shape"></div>
        <div class="shape"></div>
    </div>
    
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
        
        <div class="bottom-navbar">
            <nav class="navigation">
                <button class="nav-btn" id="prevBtn" aria-label="Previous Slide">
                    ← Previous
                </button>
                <button class="nav-btn" id="nextBtn" aria-label="Next Slide">
                    Next →
                </button>
            </nav>
            <div class="slide-indicator" id="slideIndicator" aria-live="polite">1 / {{ slides|length }}</div>
            <div class="progress-bar" id="progressBar" role="progressbar"></div>
        </div>
    </div>

    <script>
        // Animated geometric shapes with optimized performance
        function animateShapes() {
            const shapes = document.querySelectorAll('.shape');
            shapes.forEach((shape, index) => {
                let position = { x: 0, y: 0 };
                setInterval(() => {
                    const randomX = (Math.random() - 0.5) * 30;
                    const randomY = (Math.random() - 0.5) * 30;
                    position.x += randomX * 0.1;
                    position.y += randomY * 0.1;
                    
                    // Keep shapes within reasonable bounds
                    position.x = Math.max(-50, Math.min(50, position.x));
                    position.y = Math.max(-50, Math.min(50, position.y));
                    
                    shape.style.transform = `translate(${position.x}px, ${position.y}px) rotate(${Date.now() * 0.001 + index}rad)`;
                }, 4000 + index * 1500);
            });
        }

        document.addEventListener('DOMContentLoaded', () => {
            // Only animate shapes on non-mobile devices for better performance
            if (window.innerWidth > 768) {
                animateShapes();
            }
            
            const slides = document.getElementById('slides');
            const prevBtn = document.getElementById('prevBtn');
            const nextBtn = document.getElementById('nextBtn');
            const slideIndicator = document.getElementById('slideIndicator');
            const progressBar = document.getElementById('progressBar');
            const totalSlides = {{ slides|length }};
            let currentSlide = 0;
            let touchStartX = 0;
            let touchEndX = 0;

            const updateSlide = () => {
                slides.style.transform = `translateX(-${currentSlide * 100}%)`;
                slideIndicator.textContent = `${currentSlide + 1} / ${totalSlides}`;
                prevBtn.disabled = currentSlide === 0;
                nextBtn.disabled = currentSlide === totalSlides - 1;
                progressBar.style.width = `${((currentSlide + 1) / totalSlides) * 100}%`;
                
                // Update ARIA attributes for accessibility
                slides.setAttribute('aria-live', 'polite');
                progressBar.setAttribute('aria-valuenow', currentSlide + 1);
                progressBar.setAttribute('aria-valuemax', totalSlides);
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

            // Touch/swipe support for mobile
            const handleTouchStart = (e) => {
                touchStartX = e.changedTouches[0].screenX;
            };

            const handleTouchEnd = (e) => {
                touchEndX = e.changedTouches[0].screenX;
                const swipeThreshold = 50;
                const swipeDistance = touchStartX - touchEndX;
                
                if (Math.abs(swipeDistance) > swipeThreshold) {
                    if (swipeDistance > 0) {
                        nextSlide();
                    } else {
                        previousSlide();
                    }
                }
            };

            // Event listeners
            prevBtn.addEventListener('click', previousSlide);
            nextBtn.addEventListener('click', nextSlide);
            
            // Keyboard navigation
            document.addEventListener('keydown', (e) => {
                if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
                    e.preventDefault();
                    previousSlide();
                }
                if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
                    e.preventDefault();
                    nextSlide();
                }
            });
            
            // Touch events for mobile swiping
            slides.addEventListener('touchstart', handleTouchStart, { passive: true });
            slides.addEventListener('touchend', handleTouchEnd, { passive: true });
            
            // Handle window resize
            let resizeTimeout;
            window.addEventListener('resize', () => {
                clearTimeout(resizeTimeout);
                resizeTimeout = setTimeout(updateSlide, 150);
            });

            // Initialize
            updateSlide();
            
            // Set initial ARIA attributes
            progressBar.setAttribute('role', 'progressbar');
            progressBar.setAttribute('aria-label', 'Presentation progress');
            progressBar.setAttribute('aria-valuemin', '1');
        });
    </script>
</body>
</html>""",
    "glassy" : """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }}</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=Roboto:wght@300;400;500;600;700&display=swap');
        
        /* Reset and Base Styles */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Inter', sans-serif;
            background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 25%, #cbd5e1 100%);
            overflow: hidden;
            line-height: 1.6;
            color: #1e293b;
            position: relative;
            height: 100vh;
        }

        /* Subtle Background Elements */
        .bg-pattern {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            pointer-events: none;
            z-index: 1;
            opacity: 0.4;
        }
        
        .pattern-element {
            position: absolute;
            width: 2px;
            height: 2px;
            background: #64748b;
            border-radius: 50%;
            animation: subtleFloat 12s ease-in-out infinite;
        }
        
        .pattern-element:nth-child(1) {
            left: 20%;
            top: 30%;
            animation-delay: 0s;
        }
        
        .pattern-element:nth-child(2) {
            right: 25%;
            top: 60%;
            animation-delay: -4s;
        }
        
        .pattern-element:nth-child(3) {
            left: 70%;
            top: 20%;
            animation-delay: -8s;
        }
        
        @keyframes subtleFloat {
            0%, 100% { transform: translateY(0px); opacity: 0.3; }
            50% { transform: translateY(-10px); opacity: 0.6; }
        }

        /* Presentation Container */
        .presentation {
            width: 100vw;
            height: 100vh;
            background: rgba(255, 255, 255, 0.9);
            backdrop-filter: blur(10px);
            position: relative;
            border: 1px solid rgba(148, 163, 184, 0.2);
            overflow: hidden;
            display: flex;
            flex-direction: column;
            box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.1);
        }

        /* Header Styles */
        .header {
            background: rgba(30, 41, 59, 0.95);
            backdrop-filter: blur(20px);
            color: #ffffff;
            padding: 2.5rem 3rem;
            text-align: center;
            border-bottom: 3px solid #3b82f6;
            box-shadow: 0 4px 20px rgba(30, 41, 59, 0.15);
            flex-shrink: 0;
            z-index: 10;
            position: relative;
        }
        
        .header::before {
            content: '';
            position: absolute;
            top: 1rem;
            left: 2rem;
            right: 2rem;
            bottom: 1rem;
            border: 1px solid rgba(59, 130, 246, 0.3);
            border-radius: 8px;
            pointer-events: none;
        }
        
        .header h1 {
            font-family: 'Roboto', sans-serif;
            font-size: 2.5rem;
            font-weight: 600;
            letter-spacing: -0.5px;
            color: #ffffff;
            position: relative;
            margin: 0;
        }
        
        .header::after {
            content: '';
            position: absolute;
            bottom: -3px;
            left: 50%;
            transform: translateX(-50%);
            width: 120px;
            height: 3px;
            background: linear-gradient(90deg, #3b82f6, #1d4ed8);
            border-radius: 2px;
        }

        /* Main Content Area */
        .slides-container {
            flex: 1;
            overflow: hidden;
            position: relative;
        }
        
        .slides {
            display: flex;
            height: 100%;
            transition: transform 0.6s cubic-bezier(0.4, 0, 0.2, 1);
        }

        /* Slide Styles */
        .slide {
            min-width: 100%;
            height: 100%;
            display: flex;
            flex-direction: row;
            align-items: flex-start;
            padding: 2.5rem 3rem;
            background: rgba(255, 255, 255, 0.8);
            backdrop-filter: blur(8px);
            position: relative;
            overflow-y: auto;
        }
        
        .slide::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: 
                radial-gradient(circle at 20% 80%, rgba(59, 130, 246, 0.03) 0%, transparent 50%),
                radial-gradient(circle at 80% 20%, rgba(139, 92, 246, 0.02) 0%, transparent 50%);
            pointer-events: none;
        }
        
        .slide::after {
            content: '';
            position: absolute;
            top: 1.5rem;
            left: 1.5rem;
            right: 1.5rem;
            bottom: 1.5rem;
            border: 1px solid rgba(148, 163, 184, 0.2);
            border-radius: 12px;
            pointer-events: none;
        }
        
        .slide-image-container {
            width: 40%;
            padding: 1rem 2rem 1rem 1rem;
            z-index: 2;
            display: flex;
            justify-content: center;
            align-items: flex-start;
            flex-shrink: 0;
        }
        
        .slide-image {
            max-width: 100%;
            max-height: 55vh;
            object-fit: contain;
            border-radius: 12px;
            background: rgba(255, 255, 255, 0.9);
            border: 1px solid rgba(148, 163, 184, 0.3);
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.1);
            transition: all 0.3s ease;
            padding: 0.5rem;
        }
        
        .slide-image:hover {
            transform: translateY(-5px);
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.15);
            border-color: rgba(59, 130, 246, 0.4);
        }
        
        .slide-content-container {
            width: 60%;
            padding: 1rem 1rem 1rem 2rem;
            z-index: 2;
            flex: 1;
            overflow-y: auto;
        }
        
        .slide-title {
            font-family: 'Roboto', sans-serif;
            font-size: 2.3rem;
            font-weight: 600;
            color: #1e293b;
            margin-bottom: 1.5rem;
            letter-spacing: -0.3px;
            line-height: 1.2;
            position: relative;
        }
        
        .slide-title::after {
            content: '';
            position: absolute;
            bottom: -10px;
            left: 0;
            width: 60px;
            height: 3px;
            background: linear-gradient(90deg, #3b82f6, #1d4ed8);
            border-radius: 2px;
        }
        
        .slide-content {
            font-size: 1.1rem;
            color: #475569;
            margin-bottom: 2rem;
            line-height: 1.7;
            font-weight: 400;
        }
        
        .bullets {
            list-style: none;
            padding-left: 0;
        }
        
        .bullets li {
            font-size: 1rem;
            color: #334155;
            margin-bottom: 1rem;
            padding: 1.2rem 1.5rem 1.2rem 2.5rem;
            background: rgba(255, 255, 255, 0.7);
            backdrop-filter: blur(8px);
            border-radius: 8px;
            border-left: 3px solid #3b82f6;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
            position: relative;
            transition: all 0.3s ease;
            line-height: 1.6;
            border: 1px solid rgba(148, 163, 184, 0.1);
        }
        
        .bullets li::before {
            content: '▶';
            position: absolute;
            left: 1rem;
            top: 50%;
            transform: translateY(-50%);
            color: #3b82f6;
            font-size: 0.8rem;
            font-weight: 600;
        }
        
        .bullets li:hover {
            transform: translateX(8px);
            background: rgba(255, 255, 255, 0.9);
            border-left-color: #1d4ed8;
            box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1);
            border-color: rgba(59, 130, 246, 0.2);
        }
        
        .bullets li:hover::before {
            color: #1d4ed8;
            transform: translateY(-50%) scale(1.1);
        }

        /* Bottom Navigation Bar */
        .bottom-navbar {
            flex-shrink: 0;
            background: rgba(30, 41, 59, 0.95);
            backdrop-filter: blur(20px);
            border-top: 1px solid rgba(148, 163, 184, 0.3);
            padding: 1.5rem 2rem;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 -4px 20px rgba(30, 41, 59, 0.1);
            z-index: 10;
        }
        
        .navigation {
            display: flex;
            gap: 1rem;
        }
        
        .nav-btn {
            background: rgba(59, 130, 246, 0.9);
            backdrop-filter: blur(10px);
            color: #ffffff;
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 6px;
            padding: 0.8rem 1.5rem;
            font-size: 0.9rem;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 0.5rem;
            box-shadow: 0 2px 8px rgba(59, 130, 246, 0.2);
            transition: all 0.3s ease;
            font-weight: 500;
            min-width: 100px;
            font-family: 'Inter', sans-serif;
        }
        
        .nav-btn:hover {
            background: rgba(29, 78, 216, 0.95);
            transform: translateY(-2px);
            box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
            border-color: rgba(255, 255, 255, 0.3);
        }
        
        .nav-btn:active {
            transform: translateY(0);
        }
        
        .nav-btn:disabled {
            background: rgba(148, 163, 184, 0.5);
            color: rgba(255, 255, 255, 0.6);
            border-color: rgba(148, 163, 184, 0.3);
            cursor: not-allowed;
            transform: none;
            box-shadow: 0 1px 4px rgba(0, 0, 0, 0.1);
        }

        /* Slide Indicator and Progress Bar */
        .slide-indicator {
            background: rgba(255, 255, 255, 0.9);
            backdrop-filter: blur(10px);
            color: #1e293b;
            border: 1px solid rgba(148, 163, 184, 0.3);
            padding: 0.7rem 1.2rem;
            border-radius: 20px;
            font-size: 0.9rem;
            font-weight: 600;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
            font-family: 'Inter', sans-serif;
        }
        
        .progress-bar {
            position: absolute;
            bottom: 0;
            left: 0;
            height: 3px;
            background: linear-gradient(90deg, #3b82f6, #1d4ed8);
            transition: width 0.6s ease;
            box-shadow: 0 0 10px rgba(59, 130, 246, 0.3);
        }

        /* Mobile Responsive Design */
        @media screen and (max-width: 768px) {
            .header {
                padding: 2rem 1.5rem;
            }
            .header h1 {
                font-size: 2rem;
            }
            .slide {
                flex-direction: column;
                padding: 1.5rem;
                align-items: center;
            }
            .slide-image-container,
            .slide-content-container {
                width: 100%;
                padding: 1rem 0.5rem;
            }
            .slide-image {
                max-height: 30vh;
            }
            .slide-title {
                font-size: 1.8rem;
                text-align: center;
            }
            .slide-content {
                font-size: 1rem;
                text-align: center;
            }
            .bullets li {
                font-size: 0.95rem;
                padding: 1rem 1.3rem 1rem 2.2rem;
            }
            .bottom-navbar {
                padding: 1rem;
                flex-direction: column;
                gap: 1rem;
            }
            .nav-btn {
                padding: 0.6rem 1.2rem;
                font-size: 0.85rem;
                min-width: 90px;
            }
        }

        @media screen and (max-width: 480px) {
            .header h1 {
                font-size: 1.6rem;
            }
            .slide {
                padding: 1rem;
            }
            .slide-title {
                font-size: 1.4rem;
            }
            .slide-content {
                font-size: 0.9rem;
            }
            .bullets li {
                font-size: 0.85rem;
                padding: 0.9rem 1.1rem 0.9rem 2rem;
            }
            .nav-btn {
                padding: 0.5rem 1rem;
                font-size: 0.8rem;
                min-width: 80px;
            }
        }

        /* Touch-friendly enhancements */
        @media (hover: none) and (pointer: coarse) {
            .nav-btn:hover {
                transform: none;
                background: rgba(59, 130, 246, 0.9);
            }
            .bullets li:hover {
                transform: none;
            }
            .slide-image:hover {
                transform: none;
            }
        }
    </style>
</head>
<body>
    <div class="bg-pattern">
        <div class="pattern-element"></div>
        <div class="pattern-element"></div>
        <div class="pattern-element"></div>
    </div>
    
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
        
        <div class="bottom-navbar">
            <nav class="navigation">
                <button class="nav-btn" id="prevBtn" aria-label="Previous Slide">
                    ← Previous
                </button>
                <button class="nav-btn" id="nextBtn" aria-label="Next Slide">
                    Next →
                </button>
            </nav>
            <div class="slide-indicator" id="slideIndicator" aria-live="polite">1 / {{ slides|length }}</div>
        </div>
        
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
</html>""",
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

def create_pptx_from_yaml(yaml_content, output_path, topic, email):
    """Create PowerPoint presentation from YAML content and save both to disk and DB."""
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
            data = yaml.safe_load(cleaned_yaml)
            if data is None:
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

        # Your slide creation code here (title slide and others) ...
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

        # Save PPTX file to disk (existing functionality)
        prs.save(output_path)
        logger.info(f"PowerPoint saved to {output_path}")

        # Also save PPTX as base64 string in DB
        pptx_stream = io.BytesIO()
        prs.save(pptx_stream)
        pptx_stream.seek(0)
        base64_pptx = base64.b64encode(pptx_stream.read()).decode('utf-8')

        file_record = FileRecord(
            topic=topic,
            user_email=email,
            file_type='pptx',
            file_data=base64_pptx
        )
        db.session.add(file_record)
        db.session.commit()
        logger.info(f"PowerPoint stored in DB for email: {email}, topic: {topic}")

        return {"success": True, "file_id": file_record.id}

    except Exception as e:
        logger.error(f"Error in create_pptx_from_yaml: {str(e)}")
        return {"success": False, "error": str(e)}

def create_docx_from_yaml(yaml_content, output_path, email):
    """Create Word document from YAML content and store as base64 in DB."""
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
            data = yaml.safe_load(cleaned_yaml)
            if data is None:
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

        # Generate Word document
        doc = Document()
        doc.add_heading('Presentation', 0)
        for slide in slides:
            doc.add_heading(slide.get('title', 'Untitled'), level=1)
            for bullet in slide.get('bullets', []):
                doc.add_paragraph(bullet, style='ListBullet')

        # Save the file locally
        doc.save(output_path)
        logger.info(f"Word document saved to {output_path}")

        # Read file and encode in base64
        with open(output_path, "rb") as f:
            encoded_data = base64.b64encode(f.read()).decode('utf-8')

        # Extract topic from YAML (fallback to filename if not found)
        topic = data.get('presentation', {}).get('title') or os.path.basename(output_path).split('.')[0]

        # Store record in DB
        file_record = FileRecord(
            topic=topic,
            user_email=email,
            file_type='docx',
            file_data=encoded_data
        )
        db.session.add(file_record)
        db.session.commit()
        logger.info(f"FileRecord created in DB for {email}")

        return {"success": True, "file_record_id": file_record.id}
    
    except Exception as e:
        logger.error(f"Error in create_docx_from_yaml: {str(e)}")
        return {"success": False, "error": str(e)}

def create_pdf_from_yaml(yaml_content, output_path, email):
    """Create PDF document from YAML content and store as base64 in DB."""
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
            data = yaml.safe_load(cleaned_yaml)
            if data is None:
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

        # Save locally
        pdf.output(output_path)
        logger.info(f"PDF saved to {output_path}")

        # Read file as base64
        with open(output_path, "rb") as f:
            encoded_data = base64.b64encode(f.read()).decode('utf-8')

        topic = data.get('presentation', {}).get('title') or os.path.basename(output_path).split('.')[0]

        # Store in DB
        file_record = FileRecord(
            topic=topic,
            user_email=email,
            file_type='pdf',
            file_data=encoded_data
        )
        db.session.add(file_record)
        db.session.commit()
        logger.info(f"PDF FileRecord created for {email}")

        return {"success": True, "file_record_id": file_record.id}

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

def create_html_from_yaml(yaml_content, output_path, topic, email, html_presentation_type='minimalist'):
    """Create HTML presentation from YAML content and store base64 in DB."""
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
            if data is None:
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

        # Process images
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

        # Render HTML
        jinja_template = Template(HTML_TEMPLATES.get(html_presentation_type, HTML_TEMPLATES['minimalist']))
        html_content = jinja_template.render(**template_data)

        # Save to file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        logger.info(f"HTML presentation saved to {output_path}")

        # Encode content as base64
        encoded_data = base64.b64encode(html_content.encode('utf-8')).decode('utf-8')

        # Store in DB
        file_record = FileRecord(
            topic=topic,
            user_email=email,
            file_type='html',
            file_data=encoded_data
        )
        db.session.add(file_record)
        db.session.commit()
        logger.info(f"HTML FileRecord created for {email}")

        return {"success": True, "file_record_id": file_record.id}

    except Exception as e:
        logger.error(f"Error in create_html_from_yaml: {str(e)}")
        return {"success": False, "error": str(e)}