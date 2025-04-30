# AI Presentation Generator

An application that generates presentations (PowerPoint, Word, PDF) based on a given topic using AI.

## Features

- Generate comprehensive presentations from a topic
- Export to multiple formats (PPTX, DOCX, PDF)
- Automatic image search and inclusion
- Consistent formatting with attractive backgrounds

## Project Structure

```
presentation_generator/
│
├── app.py                  # Main entry point
├── config.py               # Configuration settings
├── requirements.txt        # Project dependencies
├── .gitignore              # Git ignore file
│
├── static/                 # Static files (add as needed)
│   ├── css/
│   └── js/
│
├── templates/              # HTML templates
│   └── index.html          # Main page template
│
├── generated_files/        # Directory for generated files
│
└── services/               # Application services
    ├── __init__.py
    ├── ai_service.py       # LLM interaction
    ├── image_service.py    # Image processing
    ├── document_service.py # Document generation
    └── utils.py            # Utility functions
```

## Setup and Installation

1. Clone the repository
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Add your Pexels API key to `config.py` (or use environment variables)
4. Download the LLaMA model weights and place in project root as `customweights.gguf`
5. Run the application:
   ```
   python app.py
   ```

## Configuration

Edit `config.py` to change:
- API keys
- File paths
- Model parameters
- Debug settings

## Adding New Features

To add new features:
1. Extend the relevant service modules
2. Update the main app routes as needed
3. Add new templates if necessary

## License

Copyright (c) 2016 TSK Engineers Private Limited. All rights reserved.
