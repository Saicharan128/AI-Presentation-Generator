"""
Utility functions for various tasks throughout the application.
This module contains helper functions that might be used across different services.
"""

def sanitize_filename(filename):
    """
    Sanitize a filename by removing invalid characters
    and replacing spaces with underscores.
    
    Args:
        filename (str): The original filename
        
    Returns:
        str: Sanitized filename
    """
    # Replace spaces with underscores
    sanitized = filename.replace(' ', '_')
    
    # Remove any characters that aren't allowed in filenames
    sanitized = ''.join(c for c in sanitized if c.isalnum() or c in '_-.')
    
    return sanitized

def truncate_text(text, max_length=100):
    """
    Truncate text to a maximum length, adding ellipsis if needed.
    
    Args:
        text (str): The text to truncate
        max_length (int): Maximum length
        
    Returns:
        str: Truncated text
    """
    if len(text) <= max_length:
        return text
    return text[:max_length - 3] + "..."