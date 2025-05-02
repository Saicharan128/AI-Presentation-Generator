import re
import yaml
from llama_cpp import Llama
from config import Config

from llama_cpp import Llama

model = Llama(
    model_path=Config.MODEL_PATH,
    n_ctx=2048,
    n_gpu_layers=-1
)

def extract_yaml_block(text):
    """Extract YAML content from model output."""
    if "---" in text:
        yaml_parts = text.split("---", 1)
        return "---" + yaml_parts[1].split("\n---")[0]
    else:
        return "---\npresentation:\n  slides: []"

def fix_yaml_format(yaml_text):
    """Fix common YAML formatting issues."""
    fixed_text = re.sub(r"^\s*[\*\d]\.", "  -", yaml_text, flags=re.MULTILINE)
    fixed_text = re.sub(r"^\s*\*", "  -", fixed_text, flags=re.MULTILINE)
    return fixed_text

def extract_slide_count(topic):
    """Extract slide count from topic string like 'Topic Name (5 slides)'."""
    match = re.search(r"\((\d+)\s*slides?\)", topic, re.IGNORECASE)
    return int(match.group(1)) if match else 3

def clean_topic_name(topic):
    """Remove slide count part from topic for cleaner titles."""
    return re.sub(r"\s*\(\d+\s*slides?\)", "", topic, flags=re.IGNORECASE).strip()

def parse_specific_slides(prompt):
    """Parse prompt for specific slide details (e.g., Slide 1: Title - Bullets)."""
    slide_pattern = r"Slide\s*(\d+):\s*([^\n]+)\s*-\s*([^\n]+)"
    matches = re.findall(slide_pattern, prompt, re.MULTILINE)
    
    slides = []
    for match in matches:
        slide_num, title, bullets = match
        bullet_list = [b.strip() for b in bullets.split(";") if b.strip()]
        slides.append({
            "title": title.strip(),
            "bullets": bullet_list
        })
    return slides

def generate_yaml_from_topic(topic):
    """Generate presentation structure in YAML format based on a topic."""
    # Check if the prompt contains specific slide details
    specific_slides = parse_specific_slides(topic)
    clean_topic = clean_topic_name(topic)
    
    if specific_slides:
        # Use specific slides directly
        presentation = {
            "presentation": {
                "title": clean_topic,
                "slides": specific_slides
            }
        }
        yaml_content = yaml.dump(presentation, sort_keys=False)
        return f"---\n{yaml_content}\n---"
    
    # Fallback to model-based generation
    slide_count = extract_slide_count(topic)
    
    prompt = (
        f"Create a YAML for a detailed presentation on '{clean_topic}'.\n"
        f"Limit it to {slide_count} slides.\n"
        f"The bullets MUST be complete explanatory sentences.\n"
        f"Format:\n"
        f"---\n"
        f"presentation:\n"
        f"  slides:\n"
        f"    - title: Slide 1 Title\n"
        f"      bullets:\n"
        f"        - Bullet 1 full sentence.\n"
        f"        - Bullet 2 detailed info.\n"
        f"---\n"
        f"Only output valid YAML. No extra text."
    )

    chat_messages = [
        {"role": "system", "content": "You generate YAML content for informative presentations. Never output questions, only clear explanations."},
        {"role": "user", "content": prompt}
    ]

    response = model.create_chat_completion(
        messages=chat_messages,
        max_tokens=700,
        temperature=0.3,
        top_p=0.9,
        stop=["</s>"]
    )

    raw_output = response['choices'][0]['message']['content'].strip()
    yaml_only = extract_yaml_block(raw_output)
    yaml_fixed = fix_yaml_format(yaml_only)

    return yaml_fixed