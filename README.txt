#PNG-to-Excel Digitization for Advent Energy

## Introduction
This Python script utilizes the pytesseract library to extract text from PNG images containing industry contact information. The extracted data is then organized and stored in an Excel spreadsheet for easy access and management.

## Prerequisites
Before running the script, ensure you have the following installed:
- Python 3.x: https://www.python.org/downloads/
- Tesseract OCR for Windows: https://github.com/UB-Mannheim/tesseract/wiki
- Simple installation instructions
    - Download installer: https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.3.20231005.
    - !!Save address when executing installer
    - !!Fill in installation address in following format on line 12 of final_loop.py
- For detailed instructions, visit github link: https://github.com/UB-Mannheim/tesseract/wiki

## Cloning Repository from GitHub
1.   


## Extracting files
1. Change directory to where zip files are by running following command in terminal: cd 'insert directory here'
2. In script unpacking_data.py, change line 2's ZIP file name
3. Run following line in terminal: unpacking_data.py

## Running script
1. In final_loop.py, change line 19's address to the image folder's address
2. Run the following line in terminal: final_loop.py
3. Expect 5-15 minutes of runtime, depending on the number of images
4. Output would be named database.xlsx
5. For future runs, change file name on line 135 from "database.xlsx" to something else to store different databases

