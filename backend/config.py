import os
import json
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Get current directory (where this file is located)
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

# Try different possible locations for the frontend directory
FRONTEND_DIR_INSIDE = os.path.join(CURRENT_DIR, 'frontend')
FRONTEND_DIR_SIBLING = os.path.join(os.path.dirname(CURRENT_DIR), 'frontend')

# Check which location exists
if os.path.exists(FRONTEND_DIR_INSIDE) and os.path.isdir(FRONTEND_DIR_INSIDE):
    FRONTEND_DIR = FRONTEND_DIR_INSIDE
    PROJECT_ROOT = CURRENT_DIR
    logger.info(f"Frontend directory found inside backend: {FRONTEND_DIR}")
elif os.path.exists(FRONTEND_DIR_SIBLING) and os.path.isdir(FRONTEND_DIR_SIBLING):
    FRONTEND_DIR = FRONTEND_DIR_SIBLING
    PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
    logger.info(f"Frontend directory found as sibling to backend: {FRONTEND_DIR}")
else:
    # If neither location exists, use the sibling structure but log a warning
    FRONTEND_DIR = FRONTEND_DIR_SIBLING
    PROJECT_ROOT = os.path.dirname(CURRENT_DIR)
    logger.warning(f"Frontend directory not found! Using: {FRONTEND_DIR}")

# Log paths for debugging
logger.debug(f"Current directory: {CURRENT_DIR}")
logger.debug(f"Project root: {PROJECT_ROOT}")
logger.debug(f"Frontend directory: {FRONTEND_DIR}")

# Base directory for other operations
BASE_DIR = CURRENT_DIR

# Client secrets file
CLIENT_SECRETS_FILE = os.path.join(CURRENT_DIR, 'client_secret.json')

# Check if client_secret.json exists
if not os.path.exists(CLIENT_SECRETS_FILE):
    logger.error(f"未找到{CLIENT_SECRETS_FILE}文件。请从Google Cloud Console下载凭证。")
    exit(1)

# Read client_secret.json
with open(CLIENT_SECRETS_FILE, 'r') as f:
    CLIENT_CONFIG = json.load(f)

# OAuth2 callback address
REDIRECT_URI = 'https://localhost:5001/taskpane.html'

# API scopes
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# Allow OAuth scope changes
os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'
# Allow HTTP transfer of OAuth info (only for development)
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'