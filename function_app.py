import azure.functions as func
import logging
import openpyxl
from io import BytesIO
import json
import base64
import re

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

@app.route(route="http_trigger1")
def http_trigger1(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    # Check if file is in request
    if not req.files:
        return func.HttpResponse("No file received", status_code=400)

    file = req.files['file']
    input_filename = file.filename  # Get the original file name

    # Get print areas from form data
    print_areas_json = req.form.get('print_areas', '[]')
    try:
        print_areas = json.loads(print_areas_json)
    except json.JSONDecodeError:
        return func.HttpResponse("Invalid print_areas JSON", status_code=400)

    # Check if print_areas is empty
    if not print_areas:
        return func.HttpResponse("print_areas cannot be empty", status_code=400)

    # Validate print areas format
    valid_format = True
    range_pattern = r'^[A-Z]+\d+:[A-Z]+\d+$'
    for item in print_areas:
        if not isinstance(item, dict) or 'sheet_name' not in item or 'print_area' not in item:
            valid_format = False
            break
        # Check for None values
        if item['sheet_name'] is None or item['print_area'] is None:
            valid_format = False
            break
        # Allow 'skip' as a valid print_area value
        if item['print_area'] == 'skip':
            continue
        # Check if print_area is in a valid Excel range format (e.g., A1:G66 or A1:E36,A37:E53)
        ranges = item['print_area'].split(',')
        for r in ranges:
            if not re.match(range_pattern, r.strip()):
                valid_format = False
                break
        if not valid_format:
            break

    if not valid_format:
        return func.HttpResponse("Invalid print_areas format", status_code=400)

    # Load workbook
    workbook = openpyxl.load_workbook(BytesIO(file.read()))
    logging.info(f"openpyxl {openpyxl.__version__}")
    # Prepare a dictionary to hold print area information
    print_areas_info = {}

    # Process each worksheet based on provided print areas
    for item in print_areas:
        sheet_name = item.get('sheet_name')
        print_area = item.get('print_area')
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            if print_area != 'skip':
                sheet.page_setup.orientation = sheet.ORIENTATION_PORTRAIT  # Set orientation to portrait
                sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
                
                # Clear existing page breaks
                sheet.row_breaks = openpyxl.worksheet.pagebreak.RowBreak()
                
                # Add a page break at the end of each range
                ranges = print_area.split(',')
                for r in ranges:
                    start_cell, end_cell = r.split(':')
                    end_row = int(end_cell[1:])
                    sheet.row_breaks.append(openpyxl.worksheet.pagebreak.Break(id=end_row))
                
                sheet.page_setup.fitToWidth = 1  # Fit content to one page width
                sheet.page_setup.fitToHeight = 0  # No limit on height (0 means as many pages as needed)
                sheet.page_setup.scale = 100
                sheet.sheet_properties.pageSetUpPr.fitToPage = True
                
                # Keep the print area information
                print_areas_info[sheet_name] = print_area
                sheet.print_area = print_area
                
                logging.info(f"Set print area for {sheet_name}: {print_area} and orientation to portrait")
            else:
                logging.info(f"Skipped setting print area for {sheet_name}")
        else:
            logging.warning(f"Sheet {sheet_name} not found in workbook.")

    # Save the modified workbook
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    # Create a JSON response with print areas info
    response_data = {
        "print_areas": print_areas_info  # This is a dictionary that will be converted to JSON
    }

    # Return the binary file and JSON metadata
    return func.HttpResponse(
        output.getvalue(),  # Return the binary content of the Excel file
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={
            'Content-Disposition': f'attachment; filename="{input_filename}"',
            'X-Print-Areas': json.dumps(response_data)  # Include print areas info in a custom header
        }
    )
