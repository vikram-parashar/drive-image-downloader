import pandas as pd
import argparse
import requests
import os
import sys
import concurrent.futures

print("Starting the script...")

DOWNLOAD_DIR = 'downloads'
LOG_FILE = open('log.txt', 'w')

# Redirect stdout and stderr
print("All output will be logged to log.txt")
sys.stdout = LOG_FILE
sys.stderr = LOG_FILE

def parse_arguments():
    parser = argparse.ArgumentParser(description="Process an Excel file.")
    parser.add_argument('file', type=str, default=None, help='Path to the Excel file(. xlsx)')
    parser.add_argument('--sheet', type=str, default=None, help='Name of the sheet to read (optional)')
    parser.add_argument('--image-column',type=str, default=None, help='Name of the column to read')
    parser.add_argument('--filename-column',type=str, default=None, help='Name of the column to read')
    parser.add_argument('--workers', type=int, default=5, help='Number of worker threads for downloading files')

    args = parser.parse_args()

    if not args.file or not args.file.endswith('.xlsx'):
        raise ValueError("Please provide a valid Excel file with .xlsx extension.")
    if not args.image_column:
        raise ValueError("Please provide a column name for image URLs.")
    if not args.filename_column:
        raise ValueError("Please provide a column name for file names.")

    return args

def get_column_data(file_path, sheet_name=None, column_name=None):
    try:
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        if column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' does not exist in the sheet '{sheet_name}'.")

        return df[column_name]
    except Exception as e:
        raise ValueError(f"Error reading the Excel file: {e}")


def create_download_links(links):
    """
    public link:
    https://drive.google.com/file/d/163TBfAl7DvfG3EltKhmYLu3M9OIbRZnQ/view?usp=drive_link

    download link:
    https://drive.usercontent.google.com/download?id=163TBfAl7DvfG3EltKhmYLu3M9OIbRZnQ&export=download
    """

    download_links = []
    for link in links:
        if link.startswith('https://drive.google.com/file/d/'):
            id=link.split('/')[5]
        elif link.startswith('https://drive.google.com/uc?id='):
            id = link.split('=')[1]
        else:
            print(f"Invalid public link format: {link}. Skipping.")
            continue
        download_link = f"https://drive.usercontent.google.com/download?id={id}&export=download"
        download_links.append(download_link)

    return download_links

def create_download_dir(download_dir):
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

def get_file_ext(response):
    content_type = response.headers.get('Content-Type', '')
    if 'image' in content_type:
        return content_type.split('/')[-1]
    else:
        return None

def download_file(item):
    url = item['url']
    file_name = item['file_name']
    response = requests.get(url)
    if response.status_code == 200:
        ext = get_file_ext(response)
        if ext is None:
            print(f"Could not determine file extension for {url}. Skipping download.")
            return
        save_path = os.path.join(DOWNLOAD_DIR, f"{file_name}.{ext}")
        if os.path.exists(save_path):
            print(f"File {save_path} already exists. Randomizing name.")
            save_path = os.path.join(DOWNLOAD_DIR, f"{file_name}_{os.urandom(4).hex()}.{ext}")

        with open(save_path, 'wb') as file:
            file.write(response.content)
    else:
        print(f"Failed to download: {url} with status code {response.status_code}")


if __name__ == "__main__":
    args = parse_arguments()

    links = get_column_data(args.file, args.sheet, args.image_column)
    file_names = get_column_data(args.file, args.sheet, args.filename_column)

    download_links= create_download_links(links)

    create_download_dir(DOWNLOAD_DIR)
    
    items = []
    for i in range(len(download_links)):
        items.append({
            'url': download_links[i],
            'file_name': file_names[i]
        })
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        executor.map(download_file, items)
