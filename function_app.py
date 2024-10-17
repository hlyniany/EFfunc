import azure.functions as func
import logging
import openpyxl
from io import BytesIO
import json

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

@app.route(route="http_trigger1")
def http_trigger1(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    # Check if file is in request
    if not req.files:
        return func.HttpResponse("No file received", status_code=400)

    file = req.files['file']

    # Error handling for x and y parameters
    try:
        x = int(req.form.get('x', 1))
        y = int(req.form.get('y', 1))
    except ValueError:
        return func.HttpResponse("Invalid x or y parameter", status_code=400)

    # Load workbook
    workbook = openpyxl.load_workbook(BytesIO(file.read()))

    # Prepare a dictionary to hold print area information
    print_areas_info = {}

    # Process each worksheet
    for sheet in workbook.worksheets:
        sheet.print_area = f'A1:{openpyxl.utils.get_column_letter(y)}{x}'
        print_areas_info[sheet.title] = sheet.print_area

    # Save the modified workbook
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    # Create a JSON response
    response_data = {
        "file": output.getvalue().decode('latin1'),  # Encode binary data to a string
        "print_areas": print_areas_info
    }

    return func.HttpResponse(
        json.dumps(response_data),
        mimetype='application/json',
        headers={
            'Content-Disposition': 'attachment; filename=updated_file.xlsx'
        }
    )
