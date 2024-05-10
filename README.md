# 440-Project

# Member: Micah Moody, Bot Yi, Apram.

This application provides a user interface that allows users to select a text file and convert it into PDF, Word, or HTML format. The conversion preserves text formatting and provides additional document customization features.

## Presentation & Tutorial Link:
https://youtu.be/vIbzwthvJm4

https://docs.google.com/presentation/d/1-VZHwqcRcU1HTDdZ6ftY2WAdNy-JlIF0WhlCzO-2hg0/edit?usp=sharing

## Code Documentation Link:
https://docs.google.com/document/d/1Ni5S21mc3GQFdXF8zx1EQNk5rXSwqhwHt5IMnFTRjCQ/edit

## Notebook Link:
https://colab.research.google.com/drive/1z1JqeeNXJ3maCIA6dfhKRCBbMgrLqlap?usp=sharing

## Installation

Before running the application, ensure you have Python 3 installed on your system. You can download Python 3 from the official website: https://www.python.org/downloads/

After installing Python, you will need to install the spaCy library and download its English language model. You will also need the `python-docx` library for Word document conversion.

Follow these steps to set up your environment:

1. **Install spaCy and its English language model**:

   ```bash
   python3 -m pip install spacy
   python3 -m spacy download en_core_web_sm

2. **Install docx library**:

   ```bash
   python3 -m pip install python-docx

4. **Install matplotlib library**:
   
   ```bash
   python3 -m pip install matplotlib

6. **Install bs4 library**:

   ```bash
    pip3 install bs4

7. **Install pyPDF2**:

   ```bash
   pip3 install pyPDF2
   
**Running the application through VSCode/Console**:

    python3 FileConversion.py

