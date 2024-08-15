from flask import Flask, request, jsonify
import pandas as pd
from fpdf import FPDF

app = Flask(__name__)

def generate_report(file_path):
    xls = pd.ExcelFile(file_path)
    report_data = {}
    for sheet_name in xls.sheet_names:
        sheet_df = xls.parse(sheet_name)
        sheet_data = {
            'sheet_name': sheet_name,
            'quantity': len(sheet_df),
            'average': sheet_df.mean().mean()
        }
        report_data[sheet_name] = sheet_data
    return report_data

def create_pdf_report(report_data):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    for sheet_name, sheet_info in report_data.items():
        pdf.cell(200, 10, txt=f"Sheet: {sheet_name}", ln=True)
        pdf.cell(200, 10, txt=f"Quantity: {sheet_info['quantity']}", ln=True)
        pdf.cell(200, 10, txt=f"Average: {sheet_info['average']}", ln=True)
        pdf.cell(200, 10, txt="------------", ln=True)

    pdf.output('report.pdf')

@app.route('/generate_and_download_report', methods=['POST'])
def handle_generate_and_download_report():
    data = request.json
    file_path = data.get('file_path')

    if not file_path:
        return jsonify({'error': 'Missing file_path'}), 400

    report_data = generate_report(file_path)
    create_pdf_report(report_data)

    return jsonify({'message': 'PDF report generated and available for download'})

if __name__ == '__main__':
    app.run()

