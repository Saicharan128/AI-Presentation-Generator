from playwright.async_api import async_playwright
import asyncio
import os
from config import Config
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

async def generate_slide_preview(html_content, output_path, presentation_type):
    """Generate a screenshot of an HTML slide for preview."""
    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch()
            page = await browser.new_page()
            
            # Set viewport size for consistent previews
            await page.set_viewport_size({"width": 800, "height": 600})
            
            # Write temporary HTML file
            temp_html_path = os.path.join(Config.UPLOAD_FOLDER, f"temp_preview_{presentation_type}.html")
            with open(temp_html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            # Load the HTML file
            await page.goto(f"file://{os.path.abspath(temp_html_path)}")
            
            # Wait for animations and images to load
            await page.wait_for_timeout(2000)
            
            # Take screenshot
            await page.screenshot(path=output_path, full_page=False)
            
            # Clean up
            await browser.close()
            os.remove(temp_html_path)
            
            logger.info(f"Generated preview image for {presentation_type} at {output_path}")
            return True
    except Exception as e:
        logger.error(f"Error generating preview image for {presentation_type}: {str(e)}")
        return False

def generate_preview_images(yaml_content, topic, upload_folder, email):
    """Generate preview images for all HTML presentation types."""
    from services.document_service import create_html_from_yaml
    
    preview_images = {}
    presentation_types = ['minimalist', 'modern', 'professional']
    
    for p_type in presentation_types:
        # Generate HTML content for the presentation type
        html_result = create_html_from_yaml(yaml_content, os.path.join(upload_folder, f"temp_{p_type}.html"), topic, email, html_presentation_type=p_type)
        if not html_result['success']:
            logger.error(f"Failed to generate HTML for {p_type} preview")
            continue
        
        # Read the generated HTML
        with open(os.path.join(upload_folder, f"temp_{p_type}.html"), 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # Generate preview image
        output_path = os.path.join(upload_folder, f"preview_{p_type}.png")
        success = asyncio.run(generate_slide_preview(html_content, output_path, p_type))
        
        if success:
            preview_images[p_type] = f"/download/preview_{p_type}.png"
        
        # Clean up temporary HTML
        if os.path.exists(os.path.join(upload_folder, f"temp_{p_type}.html")):
            os.remove(os.path.join(upload_folder, f"temp_{p_type}.html"))
    
    return preview_images