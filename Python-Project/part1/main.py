from flask import Flask, request, jsonify
from part1.excel_file import num_sheets, sum_column, average_column
from reportlab.pdfgen import canvas
from reportlab.lib import pagesizes
import json
import os

app = Flask(__name__)


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part in the request', 400

    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400

    if file:
        rout_file=os.path.join(os.getcwd(),'uploads',file.filename)
        file.save(rout_file)
        print(rout_file)

        numSheets = num_sheets(rout_file)
        object = {
            'file_path': rout_file,
            'num_sheets': numSheets


        }
        return object


@app.route('/calculation', methods=['POST'])
def calculation():
    data = request.get_json()
    report = []
    if 'path' not in data or 'sheets' not in data:
        return 'Invalid request data format', 400

    path = data['path']
    sheets = data['sheets']
    if path == '':
        return 'No selected path', 400

    if path:
        for sheet in sheets:
            sheet_name = sheet[0]
            sheet_data = sheet[1]
            operation = sheet_data['operation']
            columns = sheet_data['columns']
            if operation == 'sum':
                for column in columns:
                    sum = (sum_column(path, sheet_name, column))
                    report.append({"sheet": sheet_name, "column": column, "sum": sum})

            if operation == 'average':
                for column in columns:
                    avg = (average_column(path, sheet_name, column))
                    report.append({"sheet": sheet_name, "column": column, "average": avg})
        pdf_function(report)
        return jsonify(report)


def pdf_function(report):
    pdf_file = "output.pdf"
    width, height = pagesizes.A4

    c = canvas.Canvas(pdf_file, pagesize=pagesizes.A4)
    c.setFont("Helvetica", 12)

    formatted_data = json.dumps(report, indent=4)

    c.drawString(50, height - 50, "JSON Data:")
    c.drawString(50, height - 70, formatted_data)

    c.save()


if __name__ == '__main__':
    app.run()
