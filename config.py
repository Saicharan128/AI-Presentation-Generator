import os

class Config:
    # Application settings
    DEBUG = True
    
    # File upload settings
    UPLOAD_FOLDER = 'generated_files'
    
    # API keys
    # In production, these should be environment variables
    PEXELS_API_KEY = 'THTqiDKdJPqa0wGgJYQnbSdzcpfIBZnvLuswsf6ho3JbYxnyZQGqcdax'  # Replace with your real Pexels API key
    
    # Model settings
    MODEL_PATH = 'customweights.gguf'
    IMAGE_MODEL_PATH = 'customweights.gguf'
    MODEL_CONTEXT_LENGTH = 1024