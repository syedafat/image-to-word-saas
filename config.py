import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

class Config:
    # Storage Directories
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
    UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER', os.path.join(BASE_DIR, 'uploads'))
    OUTPUT_FOLDER = os.environ.get('OUTPUT_FOLDER', os.path.join(BASE_DIR, 'outputs'))
    ASSET_FOLDER = os.path.join(OUTPUT_FOLDER, 'assets')

    # Allowed File Extensions & Size Limits
    ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'pdf'}
    MAX_CONTENT_LENGTH = 32 * 1024 * 1024  # 32 MB limit

    # Google Gemini AI Settings (Strictly from environment)
    GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', '')
    GEMINI_MODEL = os.environ.get('GEMINI_MODEL', 'gemini-3.1-flash-lite')

    # Conversion Engine Settings
    DEFAULT_MODE = os.environ.get('DEFAULT_MODE', 'editable')  # 'editable' or 'high_fidelity'

    # Background Worker Settings
    MAX_WORKERS = int(os.environ.get('MAX_WORKERS', '4'))
    JOB_TIMEOUT_SECONDS = int(os.environ.get('JOB_TIMEOUT_SECONDS', '300'))

    # Automated Cleanup Settings
    CLEANUP_INTERVAL_MINUTES = 15
    FILE_MAX_AGE_SECONDS = 3600  # 1 hour
