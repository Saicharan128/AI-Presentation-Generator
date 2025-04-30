import requests
import io
from PIL import Image
import random
from config import Config

def search_pexels_image(query):
    """Search for an image on Pexels API based on query."""
    try:
        headers = {'Authorization': Config.PEXELS_API_KEY}
        url = f"https://api.pexels.com/v1/search?query={query}&per_page=1"
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            if data.get('photos'):
                return data['photos'][0]['src']['large']
        return None
    except Exception as e:
        print(f"Pexels search error: {e}")
        return None

def download_image(url):
    """Download image from URL."""
    try:
        response = requests.get(url)
        return response.content if response.status_code == 200 else None
    except Exception as e:
        print(f"Image download error: {e}")
        return None

def is_bright_image(img_data, threshold=180):
    """Check if an image is bright enough to use as background."""
    try:
        img = Image.open(io.BytesIO(img_data)).convert('L')  # grayscale
        stat = img.resize((50, 50)).getdata()
        avg = sum(stat) / len(stat)
        return avg > threshold
    except Exception as e:
        print(f"Brightness check error: {e}")
        return False

def fetch_consistent_background_image(count=5):
    """Fetch a consistent bright background image for slides."""
    search_terms = [
        "white abstract background",
        "light texture background",
        "minimal white backdrop",
        "soft clean gradient",
        "light pastel abstract",
        "white particle background",
        "abstract light particles",
        "minimal tech particles",
        "digital wave particles"
    ]
    
    random.shuffle(search_terms)
    images = []
    
    for term in search_terms:
        img_url = search_pexels_image(term)
        if img_url:
            img_data = download_image(img_url)
            if img_data and is_bright_image(img_data):
                images.append(img_data)
        if len(images) >= count:
            break

    return images[0] if images else None