import os 
from dotenv import load_dotenv

# Load .env and require FMP_API_KEY from environment (no fallback)
load_dotenv()
API_KEY = os.getenv("FMP_API_KEY")
BASE_URL = "https://financialmodelingprep.com/stable"
