import os
import pandas as pd
import requests
from tqdm import tqdm
from urllib.parse import urlparse, unquote

# Load Excel file
df = pd.read_excel("FinalLinks.xlsx")

# Change this if your column has a header like "Links"
pdf_links = df.iloc[:, 0].dropna().unique()

# Create output folder
os.makedirs("pdfs", exist_ok=True)

# Function to extract filename from URL


def get_filename_from_url(url):
    path = urlparse(url).path
    filename = os.path.basename(path)
    return unquote(filename) if filename else None


# Download PDFs
for url in tqdm(pdf_links, desc="Downloading PDFs"):
    try:
        filename = get_filename_from_url(url)
        if not filename or not filename.endswith(".pdf"):
            print(f"Invalid or missing filename for URL: {url}")
            continue

        output_path = os.path.join("pdfs", filename)

        if os.path.exists(output_path):
            print(f"Already downloaded: {filename}")
            continue

        response = requests.get(url, timeout=15)
        if response.status_code == 200 and 'application/pdf' in response.headers.get('Content-Type', ''):
            with open(output_path, 'wb') as f:
                f.write(response.content)
        else:
            print(f"Skipped (not a PDF): {url}")

    except Exception as e:
        print(f"Failed: {url} - {e}")
