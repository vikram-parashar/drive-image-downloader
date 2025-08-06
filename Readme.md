# Google Drive Image Downloader from Excel

This Python script allows you to download images from **Google Drive public links** listed in an Excel file. It supports selecting specific sheets and columns and uses multithreading to speed up the downloading process.

## 🧩 Features

- Parses `.xlsx` Excel files.
- Reads image URLs and custom file names from given columns.
- Converts public Google Drive links to direct download links.
- Downloads images concurrently using multiple threads.
- Automatically detects image file type and saves them accordingly.

---

## 📦 Dependencies

Ensure you have Python 3.7+ installed. Then install the required packages:

```bash
pip install -r requirements.txt
``` 
Or install manually:
```bash
pip install pandas openpyxl requests
``` 

## 🛠 Setup
```bash
# 1. Clone the repository
git clone https://github.com/vikram-parashar/drive-image-downloader.git
cd drive-image-downloader

# 2. Set up a virtual environment
python3 -m venv .env

# 3. Activate the virtual environment
source .env/bin/activate

# On Windows (PowerShell):
.env\Scripts\Activate.ps1

# 4. Install dependencies
pip install -r requirements.txt
```

## 🧪 Example Usage
```bash
python main.py book.xlsx --image-column image --filename-column name

#or
python main.py your_file.xlsx\
--sheet Sheet1\
--image-column ImageURLs\
--filename-column FileNames\
--workers 10
```
| Argument            | Required | Description                                                      |
| ------------------- | -------- | ---------------------------------------------------------------- |
| `file`              | ✅        | Path to the `.xlsx` file containing the image URLs.              |
| `--sheet`           | ❌        | Name of the sheet in the Excel file (optional if only one).      |
| `--image-column`    | ✅        | Column name that contains the public Google Drive image links.   |
| `--filename-column` | ✅        | Column name that contains custom file names (without extension). |
| `--workers`         | ❌        | Number of concurrent threads to use (default: 5).                |

### 📁 Output
All downloaded images will be saved inside the downloads/ directory, automatically created if it doesn't exist.
