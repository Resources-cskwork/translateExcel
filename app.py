from flask import Flask, request, send_file, render_template_string
from translate import Translator  # type: ignore
import io
import openpyxl
from openpyxl.cell.cell import MergedCell
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
translator = Translator(to_lang="en", from_lang="ko")

def translate_text(text):
    if not isinstance(text, str):
        return text
    try:
        logger.info(f"Translating text: {text}")
        translated = translator.translate(text)
        logger.info(f"Translation result: {translated}")
        return translated
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
            # Get the original sheet
            sheet = wb[sheet_name]
            
            # Create a new sheet in the new workbook
            new_sheet = new_wb.create_sheet(title=sheet_name)
            
            # Copy merged cells
            new_sheet.merged_cells = sheet.merged_cells
            
            # Copy sheet properties
            new_sheet.sheet_format = sheet.sheet_format
            new_sheet.sheet_properties = sheet.sheet_properties
            
            # Copy column dimensions
            for key, value in sheet.column_dimensions.items():
                new_sheet.column_dimensions[key] = value
                
            # Copy row dimensions
            for key, value in sheet.row_dimensions.items():
                new_sheet.row_dimensions[key] = value
            
            # Process cells
            for row in sheet.iter_rows():
                for cell in row:
                    # Get the cell's style
                    new_cell = new_sheet.cell(row=cell.row, column=cell.column)
                    new_cell._style = cell._style
                    
                    # Skip merged cells (they will be handled through the main cell)
                    if isinstance(cell, MergedCell):
                        continue
                    
                    # Translate content if it's text
                    if cell.value:
                        new_cell.value = translate_text(str(cell.value))
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
    app.run(debug=True, host='0.0.0.0', port=5001)
