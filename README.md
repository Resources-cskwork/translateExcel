# Excel Korean to English Translator

This application automatically translates Excel files from Korean to English while preserving the original formatting and styling.

## Features
- Translates Korean text to English in Excel files
- Preserves original Excel formatting and styling
- Supports multiple sheets in a single Excel file
- Simple web interface for easy use

## Installation

1. Create a virtual environment:
```bash
python3 -m venv venv
```

2. Activate the virtual environment:
```bash
source venv/bin/activate
```

3. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Make sure your virtual environment is activated:
```bash
source venv/bin/activate
```

2. Run the application:
```bash
python app.py
```

3. Open your web browser and go to:
```
http://localhost:5001
```

4. Upload your Excel file through the web interface and click "Translate" to get the English version.

5. To stop the application, press `Ctrl+C` in the terminal.

## Notes
- The application will preserve all formatting and styling from the original Excel file
- Make sure your Excel file is not currently open when uploading
- Large files may take longer to process
