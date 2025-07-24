import os
from dotenv import load_dotenv
from openai import AzureOpenAI

load_dotenv()

AZURE_API_KEY = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_ENDPOINT = os.getenv("AZURE_ENDPOINT")
AZURE_DEPLOYMENT = os.getenv("AZURE_OPENAI_DEPLOYMENT")

client = AzureOpenAI(
    api_key=AZURE_API_KEY,
    azure_endpoint=AZURE_ENDPOINT,
    api_version="2025-01-01-preview"
)

from llama_cloud_services import LlamaParse

# Optimize LlamaParse for faster processing
parser = LlamaParse(
    api_key=os.getenv("LLAMA_API_KEY"),
    num_workers=4,  # Increased from 1 to 4 for parallel processing
    verbose=True,
    language="en",
    show_progress=True,  # Show progress for better user experience
    fast_mode=True,  # Enable fast mode for quicker processing
)

