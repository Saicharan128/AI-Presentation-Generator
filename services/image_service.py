import requests
import io
from PIL import Image
import random
import re
from config import Config
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def generate_search_terms(query, slide_title=None):
    query = re.sub(r'[^\w\s]', '', query.lower()).strip()
    
    search_terms = [query]
    if slide_title:
        slide_title = re.sub(r'[^\w\s]', '', slide_title.lower()).strip()
        search_terms.append(f"{query} {slide_title}")
    
    adjectives = ['professional', 'modern', 'clean', 'vivid', 'bright']
    for adj in adjectives:
        search_terms.append(f"{adj} {query}")
    
    search_terms.append(f"{query} background")
    
    seen = set()
    search_terms = [t for t in search_terms if not (t in seen or seen.add(t))]
    
    logger.debug(f"Generated search terms for query '{query}': {search_terms}")
    return search_terms

def is_bright_image(img_data, threshold=180):
    """Check if an image is bright enough."""
    try:
        img = Image.open(io.BytesIO(img_data)).convert('L')  # Grayscale
        stat = img.resize((50, 50)).getdata()
        brightness = sum(stat) / len(stat)
        is_bright = brightness >= threshold
        logger.debug(f"Image brightness: {brightness:.2f}, is_bright: {is_bright}")
        return is_bright
    except Exception as e:
        logger.error(f"Error checking image brightness: {str(e)}")
        return False

def search_pexels_image(query, slide_title=None):
    """Search for bright images on Pexels API and randomly select one."""
    try:
        headers = {'Authorization': Config.PEXELS_API_KEY}
        search_terms = generate_search_terms(query, slide_title)
        bright_images = []
        per_page = 10
        
        for term in search_terms:
            url = f"https://api.pexels.com/v1/search?query={term}&per_page={per_page}&orientation=landscape"
            response = requests.get(url, headers=headers)
            if response.status_code != 200:
                logger.warning(f"Pexels API request failed for term '{term}': {response.status_code}")
                continue
            
            data = response.json()
            photos = data.get('photos', [])
            if not photos:
                logger.debug(f"No photos found for term '{term}'")
                continue
            
            for photo in photos:
                img_url = photo['src']['large']
                img_data = download_image(img_url)
                if img_data and is_bright_image(img_data, threshold=180):
                    bright_images.append(img_url)
            
            if len(bright_images) >= 3:
                break
        
        if bright_images:
            selected_image = random.choice(bright_images)
            logger.info(f"Randomly selected image for query '{query}': {selected_image}")
            return selected_image
        
        logger.warning(f"No bright images found for query '{query}'")
        return None
    except Exception as e:
        logger.error(f"Pexels search error: {str(e)}")
        return None

def download_image(url):
    """Download image from URL."""
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            return response.content
        logger.warning(f"Failed to download image from {url}: {response.status_code}")
        return None
    except Exception as e:
        logger.error(f"Image download error: {str(e)}")
        return None

def fetch_consistent_background_image(topic, count=1):
    """Fetch a consistent bright background image for slides based on topic."""
    try:
        headers = {'Authorization': Config.PEXELS_API_KEY}
        search_terms = generate_search_terms(topic)
        bright_images = []
        per_page = 5
        
        for term in search_terms:
            url = f"https://api.pexels.com/v1/search?query={term}&per_page={per_page}&orientation=landscape"
            response = requests.get(url, headers=headers)
            if response.status_code != 200:
                logger.warning(f"Pexels API request failed for term '{term}': {response.status_code}")
                continue
            
            photos = response.json().get('photos', [])
            if not photos:
                logger.debug(f"No photos found for term '{term}'")
                continue
            
            for photo in photos:
                img_url = photo['src']['large']
                img_data = download_image(img_url)
                if img_data and is_bright_image(img_data, threshold=200):
                    bright_images.append(img_data)
            
            if bright_images:
                break
        
        if bright_images:
            selected_image = random.choice(bright_images)
            logger.info(f"Randomly selected background image for topic '{topic}'")
            return selected_image
        
        # Fallback to generic bright backgrounds
        fallback_terms = [
            "white abstract background",
            "light texture background",
            "minimal white backdrop",
            "soft clean gradient",
            "light pastel abstract"
        ]
        random.shuffle(fallback_terms)
        
        for term in fallback_terms:
            url = f"https://api.pexels.com/v1/search?query={term}&per_page={per_page}&orientation=landscape"
            response = requests.get(url, headers=headers)
            if response.status_code != 200:
                continue
            
            photos = response.json().get('photos', [])
            for photo in photos:
                img_url = photo['src']['large']
                img_data = download_image(img_url)
                if img_data and is_bright_image(img_data, threshold=200):
                    bright_images.append(img_data)
            
            if bright_images:
                break
        
        if bright_images:
            selected_image = random.choice(bright_images)
            logger.info(f"Randomly selected fallback background image")
            return selected_image
        
        logger.warning(f"No suitable background image found for topic '{topic}'")
        return None
    except Exception as e:
        logger.error(f"Error fetching background image: {str(e)}")
        return None