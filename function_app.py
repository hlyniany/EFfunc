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

    # Get print areas from form data
    print_areas_json = req.form.get('print_areas', '[]')
    try:
        print_areas = json.loads(print_areas_json)
    except json.JSONDecodeError:
        return func.HttpResponse("Invalid print_areas JSON", status_code=400)

    # Validate print areas format
    valid_format = True
    for item in print_areas:
        if not isinstance(item, dict) or 'sheet_name' not in item or 'print_area' not in item:
            valid_format = False
            break
        # Check if print_area is in a valid Excel range format (e.g., A1:G66)
        if not re.match(r'^[A-Z]+\d+:[A-Z]+\d+$', item['print_area']):
            valid_format = False
            break

    if not valid_format:
        return func.HttpResponse("Invalid print_areas format", status_code=400)

    # Load workbook
    workbook = openpyxl.load_workbook(BytesIO(file.read()))

    # Prepare a dictionary to hold print area information
    print_areas_info = {}

    # Process each worksheet based on provided print areas
    for item in print_areas:
        sheet_name = item.get('sheet_name')
        print_area = item.get('print_area')
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            sheet.print_area = print_area
            print_areas_info[sheet_name] = print_area
            logging.info(f"Set print area for {sheet_name}: {print_area}")
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
            'Content-Disposition': 'attachment; filename=updated_file.xlsx',
            'X-Print-Areas': json.dumps(response_data)  # Include print areas info in a custom header
        }
    )
