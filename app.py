from flask import Flask, request, send_file, render_template_string
import io
import openpyxl
from openpyxl.cell.cell import MergedCell
import logging
import requests
import json
import copy

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

def translate_text(text):
    if not text or not isinstance(text, str):
        return text

    try:
        logger.info(f"Translating text: {text}")
        
        # Create a more detailed prompt with context and instructions
        prompt = f"""You are a professional Korean to English translator specializing in business documents and Excel spreadsheets.
Please translate the following Korean text to English:

Text to translate: {text}

Context: This text is from an Excel spreadsheet containing business schedules, budgets, and planning information.

Translation requirements:
1. Maintain all numbers, special characters, and formatting exactly as in the original
2. Excel formulas (starting with =) must remain unchanged
3. Keep all file paths, URLs, and email addresses in their original form
4. Preserve all technical terms, company names, and proper nouns
5. Maintain bullet points (•, -, □) and numbering formats
6. Keep date formats (YYYY.MM.DD, MM/DD) and time formats unchanged
7. For currency amounts, keep the original numbers and currency symbols
8. If there are abbreviations or technical terms you're uncertain about, keep them in Korean
9. Translate headers and labels clearly and professionally
10. Maintain cell references (A1, B2, etc.) exactly as they appear

Important: Provide ONLY the English translation without any explanations, notes, or alternative translations.
For Excel formulas, return them exactly as provided without translation.
"""

        # Make the API call to Ollama
        response = requests.post('http://localhost:11434/api/generate', 
            json={
                'model': 'mistral:7b-instruct',
                'prompt': prompt,
                'stream': False,
                'temperature': 0.3,  # Lower temperature for more consistent translations
                'top_p': 0.9
            }
        )
        
        if response.status_code == 200:
            result = response.json()
            translated_text = result['response'].strip()
            
            # Clean up the translation
            # Remove any quotes that might have been added by the model
            translated_text = translated_text.strip('"\'')
            
            # Preserve Excel formulas exactly
            if text.startswith('='):
                return text  # Return original formula without translation
            
            # Preserve numbers and dates in original format
            if text.replace('.', '').replace('/', '').replace('-', '').isdigit():
                return text
            
            logger.info(f"Translation completed: {translated_text}")
            return translated_text
        else:
            logger.error(f"Translation API error: {response.status_code}")
            return text
            
    except Exception as e:
        logger.error(f"Translation error: {str(e)}")
        return text

def process_excel(file):
    logger.info("Starting Excel processing")
    try:
        # Read the file
        logger.info("Loading workbook")
        wb = openpyxl.load_workbook(filename=io.BytesIO(file.read()))
        
        # Create a new workbook for the translated content
        new_wb = openpyxl.Workbook()
        
        # Remove the default sheet created
        new_wb.remove(new_wb.active)
        
        # Process each sheet
        for sheet_name in wb.sheetnames:
            logger.info(f"Processing sheet: {sheet_name}")
            # Get the original sheet
            sheet = wb[sheet_name]
            
            # Create a new sheet in the new workbook
            new_sheet = new_wb.create_sheet(title=sheet_name)
            
            # Copy sheet properties
            new_sheet.sheet_properties = copy.copy(sheet.sheet_properties)
            new_sheet.sheet_format = copy.copy(sheet.sheet_format)
            new_sheet.merged_cells = copy.copy(sheet.merged_cells)
            
            # Copy column dimensions
            for key, value in sheet.column_dimensions.items():
                new_sheet.column_dimensions[key] = copy.copy(value)
            
            # Copy row dimensions
            for key, value in sheet.row_dimensions.items():
                new_sheet.row_dimensions[key] = copy.copy(value)
            
            # Process cells
            for row in sheet.iter_rows():
                for cell in row:
                    # Create new cell
                    new_cell = new_sheet.cell(row=cell.row, column=cell.column)
                    
                    # Skip merged cells (they will be handled through the main cell)
                    if isinstance(cell, MergedCell):
                        continue
                    
                    # Copy basic formatting
                    if cell.has_style:
                        # Copy font properties
                        if cell.font:
                            font_props = {
                                'name': cell.font.name,
                                'size': cell.font.size,
                                'bold': cell.font.bold,
                                'italic': cell.font.italic,
                                'vertAlign': cell.font.vertAlign,
                                'underline': cell.font.underline,
                                'strike': cell.font.strike,
                                'color': copy.copy(cell.font.color) if cell.font.color else None
                            }
                            new_cell.font = openpyxl.styles.Font(**{k: v for k, v in font_props.items() if v is not None})
                        
                        # Copy border properties
                        if cell.border:
                            border_props = {
                                'left': copy.copy(cell.border.left),
                                'right': copy.copy(cell.border.right),
                                'top': copy.copy(cell.border.top),
                                'bottom': copy.copy(cell.border.bottom),
                                'diagonal': copy.copy(cell.border.diagonal),
                                'diagonal_direction': cell.border.diagonal_direction,
                                'outline': cell.border.outline,
                                'vertical': copy.copy(cell.border.vertical),
                                'horizontal': copy.copy(cell.border.horizontal)
                            }
                            new_cell.border = openpyxl.styles.Border(**border_props)
                        
                        # Copy fill properties
                        if cell.fill:
                            if cell.fill.fill_type == 'solid':
                                new_cell.fill = openpyxl.styles.PatternFill(
                                    fill_type='solid',
                                    start_color=copy.copy(cell.fill.start_color),
                                    end_color=copy.copy(cell.fill.end_color)
                                )
                            else:
                                new_cell.fill = copy.copy(cell.fill)
                        
                        # Copy number format
                        new_cell.number_format = cell.number_format
                        
                        # Copy alignment
                        if cell.alignment:
                            align_props = {
                                'horizontal': cell.alignment.horizontal,
                                'vertical': cell.alignment.vertical,
                                'text_rotation': cell.alignment.text_rotation,
                                'wrap_text': cell.alignment.wrap_text,
                                'shrink_to_fit': cell.alignment.shrink_to_fit,
                                'indent': cell.alignment.indent,
                                'justifyLastLine': cell.alignment.justifyLastLine,
                                'readingOrder': cell.alignment.readingOrder
                            }
                            new_cell.alignment = openpyxl.styles.Alignment(**{k: v for k, v in align_props.items() if v is not None})
                        
                        # Copy protection
                        if cell.protection:
                            new_cell.protection = copy.copy(cell.protection)
                    
                    # Translate content if it's text
                    if cell.value:
                        if isinstance(cell.value, str):
                            new_cell.value = translate_text(cell.value)
                        else:
                            new_cell.value = cell.value
                    else:
                        new_cell.value = cell.value
        
        # Save to buffer
        buffer = io.BytesIO()
        new_wb.save(buffer)
        buffer.seek(0)
        return buffer
    except Exception as e:
        logger.error(f"Error processing Excel file: {str(e)}", exc_info=True)
        raise

# HTML template for the upload page
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Excel Korean to English Translator</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            text-align: center;
        }
        .container {
            background-color: #f9f9f9;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h1 {
            color: #333;
        }
        .upload-form {
            margin: 20px 0;
        }
        .message {
            margin: 10px 0;
            padding: 10px;
            border-radius: 4px;
        }
        .success {
            background-color: #d4edda;
            color: #155724;
        }
        .error {
            background-color: #f8d7da;
            color: #721c24;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Excel Korean to English Translator</h1>
        <p>Upload an Excel file to translate its content from Korean to English</p>
        
        <form class="upload-form" action="/" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx" required>
            <br><br>
            <input type="submit" value="Translate">
        </form>
        
        {% if error %}
        <div class="message error">
            {{ error }}
        </div>
        {% endif %}
    </div>
</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        logger.info("Received POST request")
        if 'file' not in request.files:
            logger.warning("No file in request")
            return render_template_string(HTML_TEMPLATE, error='No file uploaded')
        
        file = request.files['file']
        if file.filename == '':
            logger.warning("Empty filename")
            return render_template_string(HTML_TEMPLATE, error='No file selected')
        
        if not file.filename.endswith('.xlsx'):
            logger.warning("Invalid file type")
            return render_template_string(HTML_TEMPLATE, error='Please upload an Excel (.xlsx) file')
        
        try:
            logger.info(f"Processing file: {file.filename}")
            output = process_excel(file)
            logger.info("File processing completed successfully")
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name=f"translated_{file.filename}"
            )
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}", exc_info=True)
            return render_template_string(HTML_TEMPLATE, error=f'An error occurred: {str(e)}')
    
    return render_template_string(HTML_TEMPLATE)

if __name__ == '__main__':
    logger.info("Starting Flask application...")
    app.run(debug=True, host='0.0.0.0', port=5001)
