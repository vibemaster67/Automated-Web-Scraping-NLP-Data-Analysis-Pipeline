import os
import requests
import pandas as pd
from bs4 import BeautifulSoup
import pandas as pd

df = pd.read_excel(r"C:\Blackcoffer Assignment\Input.xlsx", engine='openpyxl')

urls = list(df['URL'])

urls = list(df['URL'])  # List of URLs
url_ids = list(df['URL_ID'])  # Corresponding list of URL_IDs

# Function to create a valid filename from the URL_ID
def create_valid_filename(url_id):
    # Ensuring the filename is valid (in case URL_ID has any special characters)
    return str(url_id).replace(":", "_").replace("/", "_").replace("?", "").replace("&", "_").replace("=", "_")

# Function to fetch and save content from a URL
def fetch_and_save_content(url, url_id):
    try:
        # Send a GET request to fetch the raw HTML content
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')

        # Extract the title of the article
        title = soup.find('h1').text.strip() if soup.find('h1') else "Title not found"

        # Look for the content in the div with class 'td-post-content tagdiv-type'
        content_section = soup.find('div', class_='td-post-content tagdiv-type')
        if content_section:
            content = content_section.get_text(strip=True)  # Get the text from the content section
        else:
            content = "Content not found"

        # Create a valid filename from the URL_ID
        filename = create_valid_filename(url_id) + ".txt"

        # Save the title and content to a text file
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(f"Title: {title}\n\n")
            file.write(f"Content: {content}\n")
        
        print(f"Saved content from {url} into {filename}")
    
    except Exception as e:
        print(f"Error processing {url}: {e}")

# Iterate over the list of URLs and URL_IDs, and fetch content for each
for url, url_id in zip(urls, url_ids):
    fetch_and_save_content(url, url_id)
