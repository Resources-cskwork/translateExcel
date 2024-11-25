from flask import Flask, request, send_file, render_template_string
import io
import openpyxl
from openpyxl.cell.cell import MergedCell
import logging
import requests
import json
import copy
import random

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

def analyze_excel_structure(workbook, max_items=50):
    """Analyze Excel file structure and content to build context."""
    context = {
        'headers': set(),         # Column headers
        'row_headers': set(),     # Row headers/labels
        'dates': set(),          # Date formats found
        'times': set(),          # Time formats found
        'formulas': set(),       # Formula patterns
        'currencies': set(),      # Currency formats
        'bullet_types': set(),    # Types of bullets used
        'repeated_terms': {},     # Frequently occurring terms
        'cell_formats': set()     # Special cell formats
    }
    
    try:
        # Count total cells to process
        total_cells = sum(sheet.max_row * sheet.max_column for sheet in workbook)
        if total_cells > 10000:  # Limit for very large files
            logger.warning(f"Large file detected ({total_cells} cells). Using sampling for analysis.")
            sampling_rate = max(0.1, 10000 / total_cells)  # Sample at most 10000 cells
        else:
            sampling_rate = 1.0

        for sheet in workbook:
            # Analyze headers (first row only)
            for cell in list(sheet[1])[:max_items]:
                if cell.value:
                    context['headers'].add(str(cell.value))
            
            # Analyze first column (limited number of row headers)
            for row in list(sheet.iter_rows())[:max_items]:
                if row[0].value:
                    context['row_headers'].add(str(row[0].value))
            
            # Sample cells for pattern analysis
            rows = list(sheet.iter_rows())
            if sampling_rate < 1.0:
                rows = random.sample(rows, int(len(rows) * sampling_rate))
            
            for row in rows:
                for cell in row:
                    if cell.value and random.random() < sampling_rate:
                        try:
                            value = str(cell.value)
                            
                            # Limit collection sizes
                            if len(context['formulas']) < max_items and isinstance(value, str) and value.startswith('='):
                                context['formulas'].add(value[:100])  # Limit formula length
                            
                            if len(context['bullet_types']) < 10 and isinstance(value, str) and value.strip().startswith(('•', '-', '□', '○', '◆')):
                                context['bullet_types'].add(value[0])
                            
                            if cell.number_format:
                                if len(context['dates']) < 10 and ('y' in cell.number_format.lower() or 'm' in cell.number_format.lower() or 'd' in cell.number_format.lower()):
                                    context['dates'].add(cell.number_format)
                                
                                if len(context['times']) < 10 and ('h' in cell.number_format.lower() or ':' in cell.number_format):
                                    context['times'].add(cell.number_format)
                                
                                if len(context['currencies']) < 10 and any(curr in cell.number_format for curr in ['$', '￦', '¥', '€', 'RM']):
                                    context['currencies'].add(cell.number_format)
                                
                                if len(context['cell_formats']) < max_items and cell.number_format != 'General':
                                    context['cell_formats'].add(cell.number_format)
                            
                            # Count repeated terms (limit length and quantity)
                            if isinstance(value, str) and len(value) > 1 and len(value) <= 100:
                                if len(context['repeated_terms']) < max_items:
                                    context['repeated_terms'][value] = context['repeated_terms'].get(value, 0) + 1

                        except Exception as e:
                            logger.warning(f"Error processing cell: {str(e)}")
                            continue

        # Filter and limit repeated terms
        context['repeated_terms'] = dict(sorted(
            [(k, v) for k, v in context['repeated_terms'].items() if v > 1],
            key=lambda x: x[1],
            reverse=True
        )[:max_items])

        # Convert sets to lists and limit sizes
        for key in context:
            if isinstance(context[key], set):
                context[key] = list(context[key])[:max_items]

        return context

    except Exception as e:
        logger.error(f"Error in analyze_excel_structure: {str(e)}")
        return {key: [] if isinstance(value, set) else {} if isinstance(value, dict) else value 
                for key, value in context.items()}

def translate_text(text, context=None):
    if not text or not isinstance(text, str):
        return text

    try:
        logger.info(f"Translating text: {text}")
        
        # Skip translation for formulas, dates, and pure numbers
        if text.startswith('='):
            return text
        if text.replace('.', '').replace('/', '').replace('-', '').isdigit():
            return text
            
        # Check if the text matches any special patterns from context
        if context:
            # Preserve headers if they're being reused
            if text in context['headers']:
                logger.info(f"Found header pattern: {text}")
            
            # Preserve row labels if they're being reused
            if text in context['row_headers']:
                logger.info(f"Found row header pattern: {text}")
            
            # Preserve bullet points
            if any(text.startswith(bullet) for bullet in context['bullet_types']):
                # Only translate the text after the bullet
                bullet = text[0]
                remaining_text = text[1:].strip()
                translated_remaining = translate_with_context(remaining_text, context)
                return f"{bullet} {translated_remaining}"
        
        return translate_with_context(text, context)
            
    except Exception as e:
        logger.error(f"Translation error: {str(e)}")
        return text

def translate_with_context(text, context=None):
    """Translate text with context awareness."""
    try:
        context_str = ""
        if context:
            # Limit context size
            headers = ', '.join(context['headers'][:10])
            common_terms = ', '.join(k for k, v in list(context['repeated_terms'].items())[:10])
            date_formats = ', '.join(context['dates'][:5])
            time_formats = ', '.join(context['times'][:5])
            currency_formats = ', '.join(context['currencies'][:5])
            bullet_types = ', '.join(context['bullet_types'][:5])

            context_str = f"""Document Context:
- Common Headers: {headers[:200]}
- Frequent Terms: {common_terms[:200]}
- Date Formats: {date_formats[:100]}
- Time Formats: {time_formats[:100]}
- Currency Formats: {currency_formats[:100]}
- Bullet Types: {bullet_types[:50]}
"""

        # Limit prompt size
        prompt = f"""You are a professional Korean to English translator specializing in business documents and Excel spreadsheets.
Please translate the following Korean text to English:

Text to translate: {text[:500]}  # Limit input text size

{context_str[:1000]}  # Limit context size

Translation requirements:
1. Maintain all numbers, special characters, and formatting exactly as in the original
2. Keep all file paths, URLs, and email addresses in their original form
3. Preserve all technical terms, company names, and proper nouns
4. For currency amounts, keep the original numbers and currency symbols
5. If there are abbreviations or technical terms you're uncertain about, keep them in Korean
6. Translate consistently with other similar terms in the document

Important: Provide ONLY the English translation without any explanations or notes.
"""

        # Make the API call to Ollama with timeout
        response = requests.post(
            'http://localhost:11434/api/generate', 
            json={
                'model': 'mistral:7b-instruct',
                'prompt': prompt,
                'stream': False,
                'temperature': 0.3,
                'top_p': 0.9
            },
            timeout=30  # Add timeout
        )
        
        if response.status_code == 200:
            result = response.json()
            translated_text = result['response'].strip()
            translated_text = translated_text.strip('"\'')
            logger.info(f"Translation completed: {translated_text[:200]}")  # Limit log size
            return translated_text
        else:
            logger.error(f"Translation API error: {response.status_code}")
            return text
            
    except requests.Timeout:
        logger.error("Translation request timed out")
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
        
        # Analyze the Excel structure first
        context = analyze_excel_structure(wb)
        logger.info("Excel structure analyzed")
        
        # Create a new workbook for the translated content
        new_wb = openpyxl.Workbook()
        new_wb.remove(new_wb.active)
        
        # Process each sheet
        for sheet_name in wb.sheetnames:
            logger.info(f"Processing sheet: {sheet_name}")
            sheet = wb[sheet_name]
            new_sheet = new_wb.create_sheet(title=sheet_name)
            
            # Copy sheet properties
            new_sheet.sheet_properties = copy.copy(sheet.sheet_properties)
            new_sheet.sheet_format = copy.copy(sheet.sheet_format)
            new_sheet.merged_cells = copy.copy(sheet.merged_cells)
            
            # Copy dimensions
            for key, value in sheet.column_dimensions.items():
                new_sheet.column_dimensions[key] = copy.copy(value)
            for key, value in sheet.row_dimensions.items():
                new_sheet.row_dimensions[key] = copy.copy(value)
            
            # Process cells with context
            for row in sheet.iter_rows():
                for cell in row:
                    new_cell = new_sheet.cell(row=cell.row, column=cell.column)
                    
                    if isinstance(cell, MergedCell):
                        continue
                    
                    # Copy formatting
                    if cell.has_style:
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
                        
                        if cell.fill:
                            if cell.fill.fill_type == 'solid':
                                new_cell.fill = openpyxl.styles.PatternFill(
                                    fill_type='solid',
                                    start_color=copy.copy(cell.fill.start_color),
                                    end_color=copy.copy(cell.fill.end_color)
                                )
                            else:
                                new_cell.fill = copy.copy(cell.fill)
                        
                        new_cell.number_format = cell.number_format
                        
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
                        
                        if cell.protection:
                            new_cell.protection = copy.copy(cell.protection)
                    
                    # Translate content with context
                    if cell.value:
                        if isinstance(cell.value, str):
                            new_cell.value = translate_text(cell.value, context)
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
