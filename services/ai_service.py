import re
import yaml
import anthropic
from config import Config

# Initialize Claude client once
client = anthropic.Anthropic(api_key=Config.CLAUDE_API_KEY)

def extract_yaml_block(text):
    """Extract YAML content from model output."""
    try:
        if "---" in text:
            # Split on first --- and take everything until next --- or end
            parts = text.split("---", 1)
            yaml_content = "---" + parts[1].split("\n---", 1)[0]
            # Ensure we only have one document by removing any remaining ---
            yaml_content = yaml_content.split("\n---", 1)[0]
            return yaml_content.strip()
        else:
            return "---\npresentation:\n  slides: []"
    except Exception as e:
        print(f"Error extracting YAML block: {e}")
        return "---\npresentation:\n  slides: []"

def fix_yaml_format(yaml_text):
    """Fix common YAML formatting issues."""
    fixed_text = re.sub(r"^\s*[\*\d]\.", "  -", yaml_text, flags=re.MULTILINE)
    fixed_text = re.sub(r"^\s*\*", "  -", fixed_text, flags=re.MULTILINE)
    return fixed_text

def clean_topic_name(topic):
    """Remove slide count part from topic for cleaner titles."""
    return re.sub(r"\s*\(\d+\s*slides?\)", "", topic, flags=re.IGNORECASE).strip()

def extract_slide_count(topic):
    """Extract slide count from topic if specified."""
    match = re.search(r"(\d+)\s*slides?", topic, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return None

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

def generate_yaml_from_topic(topic, include_images=True):
    specific_slides = parse_specific_slides(topic)
    slide_count = extract_slide_count(topic)
    clean_topic = clean_topic_name(topic)
    
    if specific_slides:
        presentation = {
            "presentation": {
                "title": clean_topic,
                "slides": specific_slides,
                "include_images": include_images
            }
        }
        yaml_content = yaml.dump(presentation, sort_keys=False)
        return f"---\n{yaml_content}\n---"
    
    # Set default slide count if not specified
    slide_count_str = f" with exactly {slide_count} slides" if slide_count else " with 4 slides"
    
    system_prompt = "You generate YAML content for informative presentations. Never output questions, only clear explanations. Generate exactly the number of slides requested."
    
    user_prompt = (
        f"Create a YAML for a detailed presentation on '{clean_topic}'{slide_count_str}.\n"
        f"The bullets MUST be complete explanatory sentences.\n"
        f"Format:\n"
        f"---\n"
        f"presentation:\n"
        f"  slides:\n"
        f"    - title: Slide 1 Title\n"
        f"      bullets:\n"
        f"        - Bullet 1 full sentence.\n"
        f"        - Bullet 2 detailed info.\n"
        f"  include_images: {str(include_images).lower()}\n"
        f"---\n"
        f"Only output valid YAML. No extra text."
    )

    try:
        # Use CLAUDE_MODEL from config if available, otherwise default to claude-3-5-sonnet
        claude_model = getattr(Config, 'CLAUDE_MODEL', 'claude-3-5-sonnet-20241022')
        
        response = client.messages.create(
            model=claude_model,
            max_tokens=1500,
            temperature=0.3,
            system=system_prompt,
            messages=[
                {"role": "user", "content": user_prompt}
            ]
        )
        
        print("Response:", response, "\n\n\n\n\n")
        
        raw_output = response.content[0].text.strip()
        yaml_only = extract_yaml_block(raw_output)
        yaml_fixed = fix_yaml_format(yaml_only)

        # Ensure we have the correct number of slides if specified
        try:
            data = yaml.safe_load(yaml_fixed)
            slides = data.get('presentation', {}).get('slides', [])
            
            if slide_count and len(slides) != slide_count:
                # Adjust prompt to strongly emphasize slide count and retry
                adjusted_prompt = (
                    f"Create a YAML for a detailed presentation on '{clean_topic}' with EXACTLY {slide_count} slides.\n"
                    f"The bullets MUST be complete explanatory sentences.\n"
                    f"YOU MUST GENERATE EXACTLY {slide_count} SLIDES, NO MORE, NO LESS.\n"
                    f"Format:\n"
                    f"---\n"
                    f"presentation:\n"
                    f"  slides:\n"
                    f"    - title: Slide 1 Title\n"
                    f"      bullets:\n"
                    f"        - Bullet 1 full sentence.\n"
                    f"        - Bullet 2 detailed info.\n"
                    f"  include_images: {str(include_images).lower()}\n"
                    f"---\n"
                    f"Only output valid YAML. No extra text."
                )
                
                response = client.messages.create(
                    model=claude_model,
                    max_tokens=1500,
                    temperature=0.3,
                    system=system_prompt,
                    messages=[
                        {"role": "user", "content": adjusted_prompt}
                    ]
                )
                
                raw_output = response.content[0].text.strip()
                yaml_only = extract_yaml_block(raw_output)
                yaml_fixed = fix_yaml_format(yaml_only)
        except Exception as e:
            print(f"Error validating slide count: {e}")
        
        # Add include_images flag if not present
        try:
            data = yaml.safe_load(yaml_fixed)
            if 'presentation' in data and 'include_images' not in data['presentation']:
                data['presentation']['include_images'] = include_images
                yaml_fixed = yaml.dump(data, sort_keys=False)
                yaml_fixed = f"---\n{yaml_fixed}---"
        except Exception as e:
            print(f"Error adding include_images flag: {e}")
        
        return yaml_fixed
        
    except Exception as e:
        print(f"Error calling Claude API: {e}")
        return "---\npresentation:\n  slides: []"