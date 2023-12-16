from flask import Flask, render_template
import openpyxl

app = Flask(__name__)

# Function to read data from Excel file
def read_data_from_excel(file_path):
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        item = {'title': row[0], 'link': row[1], 'image': row[2], 'description': row[3]}
        data.append(item)

    return data

# Route to render a different template for the root route
@app.route('/')
def home():
    excel_file_path = 'home.xlsx'
    items_data = read_data_from_excel(excel_file_path)
    return render_template('home.html', items_data=items_data)

# Route to render the specific HTML file and import data from Excel
@app.route('/content/<filename_without_extension>')
def content(filename_without_extension):
    # Dynamically load the corresponding Excel file based on the filename parameter
    excel_file_path = f'{filename_without_extension}.xlsx'
    print(f"Attempting to open Excel file: {excel_file_path}")
    items_data = read_data_from_excel(excel_file_path)
    return render_template(f'content/{filename_without_extension}', items_data=items_data)

if __name__ == '__main__':
    app.run(debug=True)
