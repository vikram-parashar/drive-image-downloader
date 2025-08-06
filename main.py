import pandas as pd
import argparse
import requests
import os

DOWNLOAD_DIR = 'downloads'

def get_column_data(file_path, sheet_name=None, column_name=None):
    """
    Reads a specific column from an Excel file.

    :param file_path: Path to the Excel file.
    :param sheet_name: Name of the sheet to read (optional).
    :param column_name: Name of the column to read.
    :return: Data from the specified column as a pandas Series.
    """
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

def parse_arguments():
    parser = argparse.ArgumentParser(description="Process an Excel file.")
    parser.add_argument('file', type=str, default=None, help='Path to the Excel file(. xlsx)')
    parser.add_argument('--sheet', type=str, default=None, help='Name of the sheet to read (optional)')
    parser.add_argument('--column',type=str, default=None, help='Name of the column to read')

    args = parser.parse_args()

    if not args.file or not args.file.endswith('.xlsx'):
        raise ValueError("Please provide a valid Excel file with .xlsx extension.")
    if not args.column:
        raise ValueError("Please provide a column name to read.")

    return args

def create_download_links(public_links):
    """
    public link:
    https://drive.google.com/file/d/163TBfAl7DvfG3EltKhmYLu3M9OIbRZnQ/view?usp=drive_link

    download link:
    https://drive.usercontent.google.com/download?id=163TBfAl7DvfG3EltKhmYLu3M9OIbRZnQ&export=download
    """

    download_links = []
    for public_link in public_links:
        id=public_link.split('/')[5]
        download_link = f"https://drive.usercontent.google.com/download?id={id}&export=download"
        download_links.append(download_link)

    return download_links

def create_download_dir(download_dir):
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

def get_save_path(response,index):
    file_name = response.headers.get('Content-Disposition')
    if file_name:
        file_name = file_name.split('filename=')[1].strip('"')
    else:
        file_name = f"file_{index}.jpg"
    return os.path.join(DOWNLOAD_DIR, file_name)

def download_file(url, index):
    response = requests.get(url)
    if response.status_code == 200:
        save_path = get_save_path(response, index)

        with open(save_path, 'wb') as file:
            file.write(response.content)
    else:
        print(f"Failed to download: {link} with status code {response.status_code}")


if __name__ == "__main__":
    args = parse_arguments()

    public_links = get_column_data(args.file, args.sheet, args.column)

    download_links= create_download_links(public_links)

    create_download_dir(DOWNLOAD_DIR)

    for index,link in enumerate(download_links):
        download_file(link, index)

