from llama_cpp import Llama
import re
from config import Config

# Initialize model once
model = Llama(model_path=Config.MODEL_PATH, n_ctx=Config.MODEL_CONTEXT_LENGTH)

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

def generate_yaml_from_topic(topic):
    """Generate presentation structure in YAML format based on a topic."""
    prompt = (
        f"Create a YAML for a detailed presentation on '{topic}'.\n"
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