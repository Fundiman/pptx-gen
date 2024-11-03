import requests
from bs4 import BeautifulSoup
import os
import sys
from urllib.parse import quote_plus
from PIL import Image
from io import BytesIO

def fetch_image_urls(query, num_results):
    query = quote_plus(query)  # Encode the query for URL
    search_url = f'https://www.google.com/search?q={query}&client=img&udm=2'
    headers = {'User-Agent': 'Mozilla/5.0'}
    response = requests.get(search_url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')

    image_urls = []
    for img_tag in soup.find_all('img'):
        if len(image_urls) >= num_results:
            break
        img_url = img_tag.get('src')
        if img_url and img_url.startswith('http'):
            image_urls.append(img_url)
    return image_urls

def download_image(url, query, index):
    response = requests.get(url)
    img = Image.open(BytesIO(response.content))
    file_extension = img.format.lower()  # Get image format (e.g., 'jpeg', 'png')
    file_name = f"{query}_{index+1}.{file_extension}"  # Name format: {query}_{image_number}.{extension}
    img.save(file_name)
    print(f"Saved {file_name}")

def main():
    if len(sys.argv) != 3:
        print("Usage: python script.py <search_query> <number_of_images>")
        sys.exit(1)

    query = sys.argv[1]
    try:
        num_images = int(sys.argv[2])
    except ValueError:
        print("Error: <number_of_images> must be an integer.")
        sys.exit(1)

    print(f"Searching for images of: {query}")
    
    image_urls = fetch_image_urls(query, num_images)
    for index, url in enumerate(image_urls):
        download_image(url, query, index)
    print(f"Downloaded {len(image_urls)} images.")

if __name__ == "__main__":
    main()

