import azure.functions as func
import logging
import openpyxl
from io import BytesIO
import json
import base64

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
        else:
            logging.warning(f"Sheet {sheet_name} not found in workbook.")

    # Save the modified workbook
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    # Encode the binary data to Base64
    encoded_file = base64.b64encode(output.getvalue()).decode('utf-8')

    # Create a JSON response
    response_data = {
        "file": encoded_file,  # Base64 encoded binary data
        "print_areas": print_areas_info
    }

    return func.HttpResponse(
        json.dumps(response_data),
        mimetype='application/json',
        headers={
            'Content-Disposition': 'attachment; filename=updated_file.xlsx'
        }
    )
