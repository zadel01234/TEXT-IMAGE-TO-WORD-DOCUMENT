# TEXT-IMAGE-TO-WORD-DOCUMENT
IT CONVERTS SCANNED IMAGES WITH TEXT TO A WORD DOCUMENT

How to Use:
1. Run the program
2. Click "Browse..." to select a scanned image containing text
3. Choose the appropriate language for the text in the image
4. Specify where to save the Word document (auto-populated based on image name)
5. Click "Convert to Word" to process the image
6. When complete, you can choose to open the document

Dependencies:
You'll need to install these packages:
pip install pillow pytesseract python-docx
You'll also need to install Tesseract OCR on your system:

On Windows: Download and install from https://github.com/UB-Mannheim/tesseract/wiki
On macOS: brew install tesseract
On Linux: sudo apt install tesseract-ocr

For languages other than English, you may need to install additional language packs.
