import azure.functions as func
import logging
import openpyxl
from io import BytesIO

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

@app.route(route="http_trigger1")
def http_trigger1(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

 # Check if file is in request
    if not req.files:
        return func.HttpResponse("No file received", status_code=400)

    file = req.files['file']
    x = int(req.params.get('x', 1))
    y = int(req.params.get('y', 1))

    # Load workbook
    workbook = openpyxl.load_workbook(BytesIO(file.read()))

    # Process each worksheet
    for sheet in workbook.worksheets:
        sheet.print_area = f'A1:{openpyxl.utils.get_column_letter(y)}{x}'

    # Save the modified workbook
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return func.HttpResponse(
        output.getvalue(),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        headers={
            'Content-Disposition': 'attachment; filename=updated_file.xlsx'
        }
    )