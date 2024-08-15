from flask import Flask, request, jsonify
import pandas as pd
import matplotlib.pyplot as plt

app = Flask(__name__)

def count_sheets(file_path):
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    return len(sheet_names)

def calculate_sum_fields(file_path):
    df = pd.read_excel(file_path)
    sum_values = df.sum().to_dict()
    return sum_values

@app.route('/get_number_of_sheets', methods=['POST'])
def handle_get_number_of_sheets():
    data = request.json
    file_path = data.get('file_path')

    if not file_path:
        return jsonify({'error': 'Missing file_path'}), 400

    num_sheets = count_sheets(file_path)
    return jsonify({'number_of_sheets': num_sheets})

@app.route('/calculate_sum_fields', methods=['POST'])
def handle_calculate_sum_fields():
    data = request.json
    file_path = data.get('file_path')

    if not file_path:
        return jsonify({'error': 'Missing file_path'}), 400

    sum_values = calculate_sum_fields(file_path)
    return jsonify(sum_values)

# Assume the client will handle generating a column graph
# based on the sum of each sheet data

if __name__ == '__main__':
    app.run()


