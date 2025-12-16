import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# API Keys
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')
OCR_SPACE_API_KEY = os.getenv('OCR_SPACE_API_KEY')
OPENAI_API_KEY=os.getenv('OPENAI_API_KEY')
OPENAI_GPT_5_1_MODEL = "gpt-5.1"
OPENAI_GPT_5_MINI_MODEL = "gpt-5-mini"
# Gmail settings
GMAIL_CREDENTIALS_FILE = os.getenv('GMAIL_CREDENTIALS_FILE', 'credentials.json')

# Validate that required API keys are present
if not GEMINI_API_KEY:
    raise ValueError("GEMINI_API_KEY not found in .env file")

if not OCR_SPACE_API_KEY:
    raise ValueError("OCR_SPACE_API_KEY not found in .env file")

print("✓ Configuration loaded successfully!")
