# Marp to PowerPoint Converter

A Python script that converts Marp markdown presentations to PowerPoint (PPTX) format while preserving styles and formatting.

## Features
- Converts Marp markdown files to PowerPoint presentations
- Preserves heading styles (h1-h6)
- Supports text formatting (bold, italic, strikethrough)
- Handles bullet lists with proper indentation
- Maintains slide layouts (title slide, section header, content)
- Processes HTML div elements and their content
- Configurable slide dimensions and margins
- Debug mode for troubleshooting

## Requirements
- Python 3.8 or higher
- python-pptx
- python-dotenv

## Usage

1. Create a `.env` file with your work folder path:
   ```
   WORK_FOLDER=/path/to/your/work/folder
   ```

2. Place your Marp markdown file as `main.md` in the work folder

3. Install required packages:
   ```bash
   pip install python-pptx python-dotenv
   ```

4. Run the converter:
   ```bash
   python src/convert.py
   ```

   To enable debug output:
   ```bash
   python src/convert.py --debug
   ```

The script will generate `presentation.pptx` in your work folder.
