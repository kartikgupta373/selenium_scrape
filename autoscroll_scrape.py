from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
from bs4 import BeautifulSoup
import re
import requests
from docx import Document
from docx.shared import Inches
from io import BytesIO

driver = webdriver.Chrome()

# URL of the webpage
url = "https://example.com"  # Replace with the actual URL

# Open the webpage
driver.get(url)

# Wait for 10 seconds to let the page fully load
time.sleep(10)

# Function to scroll down the page
def scroll_page(driver, pause_time=2):
    last_height = driver.execute_script("return document.body.scrollHeight")
    
    while True:
        # Scroll down to the bottom of the page
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        
        # Wait to load the page
        time.sleep(pause_time)
        
        # Calculate new scroll height and compare with the last scroll height
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break  # No more content to load, exit the loop
        last_height = new_height

# Scroll the page to load all images
scroll_page(driver)

# Get the page source and parse it with BeautifulSoup
soup = BeautifulSoup(driver.page_source, 'html.parser')

# Find all elements with style attribute containing 'background'
elements = soup.find_all(style=re.compile(r'background: url\((.*?)\)'))

# Close the browser as we no longer need it
driver.quit()

# Extract the image URLs
image_urls = []
for elem in elements:
    match = re.search(r'url\((.*?)\)', elem['style'])
    if match:
        image_url = match.group(1)
        full_image_url = requests.compat.urljoin(url, image_url)  # Handle relative URLs
        image_urls.append(full_image_url)

# Now, save the images to a Word document
doc = Document()

for image_url in image_urls:
    try:
        img_response = requests.get(image_url)
        img_stream = BytesIO(img_response.content)
        
        # Add image to the Word document
        doc.add_picture(img_stream, width=Inches(4))  # Adjust the size as needed
    except Exception as e:
        print(f"Failed to add image {image_url}: {e}")

# Save the document
doc.save("images.docx")

print("Images have been saved to images.docx")
