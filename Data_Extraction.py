##DATA EXTRACTION

import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

excel_file_path = 'Input.xlsx'
df = pd.read_excel(excel_file_path)

output_directory = 'output_file'
os.makedirs(output_directory, exist_ok=True)

def scrape_and_save(row):
    url = row['URL']
    url_id = row['URL_ID']
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
       
        # Extract the data using BeautifulSoup
        
        title_element = soup.find('h1', {'class': 'entry-title'})
        content_element = soup.find('div', {'class': 'td-post-content'})
        
        title = title_element.text if title_element else 'Title Not Found'
        content = content_element.text if content_element else 'Content Not Found'
        output_filename = os.path.join(output_directory, f'{url_id}.txt')
        with open(output_filename, 'w', encoding='utf-8') as file:
            file.write(f'Title: {title}\n\nContent: {content}')
        return output_filename
    else:
        print(f"Failed to retrieve data from {url}")
        return None
    
df['output_file'] = df.apply(scrape_and_save, axis=1)