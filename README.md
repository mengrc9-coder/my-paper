# PDF Translation Tool

This project is a Python tool for batch translating English PDF files into Chinese PDF files.

It uses PDF24 Toolbox to convert PDF files to Word documents, then uses Argos Translate to translate the text in the Word files, and finally converts the translated Word files back to PDF.

## Features

* Batch process PDF files in a folder
* Convert PDF files to Word `.docx` files
* Translate Word document text from English to Chinese
* Convert translated Word files back to PDF
* Keep the original folder structure
* Use a translation cache to avoid repeated translation

## Requirements

This project requires:

* Windows
* Python 3
* PDF24 Toolbox
* Argos Translate English-to-Chinese offline model

Python packages:

```bash
pyautogui
pyperclip
PyGetWindow
argostranslate
lxml
```

You can install the Python packages with:

```bash
pip install -r requirements.txt
```

## Project Structure

```text
pdf-translation-tool/
├── main.py
├── README.md
└── requirements.txt
```

## Before Running

Before running the script, make sure the PDF24 path in the code is correct:

```python
PDF24_EXE = r"C:\Users\926\Desktop\PDF24\pdf24-Toolbox.exe"
```

You also need to set the correct input folder, output folder, and Word temporary folder:

```python
DEFAULT_INPUT = r"C:\Users\926\Desktop\T5A_Chapter 8.1 Building_TC4"
DEFAULT_OUTPUT = r"C:\Users\926\Desktop\trans"
WD_ROOT = r"C:\Users\926\Desktop\wd"
```

## How to Run

Run with the default settings:

```bash
python main.py
```

Or specify the input and output folders:

```bash
python main.py --input "C:\path\to\input_pdfs" --output "C:\path\to\output_pdfs"
```

Overwrite existing output files:

```bash
python main.py --input "C:\path\to\input_pdfs" --output "C:\path\to\output_pdfs" --overwrite
```

Only process a limited number of PDF files:

```bash
python main.py --input "C:\path\to\input_pdfs" --output "C:\path\to\output_pdfs" --limit 5
```

## Notes

This script controls PDF24 Toolbox through GUI automation. Therefore, the screen resolution, PDF24 window layout, and button coordinates may affect whether the script runs correctly.

If the program does not click the correct buttons, you may need to adjust the coordinate values in the script.

The Argos English-to-Chinese offline translation model must be installed before running this program.
