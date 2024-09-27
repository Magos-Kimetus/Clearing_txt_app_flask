from flask import Flask, redirect, url_for, render_template, send_file, send_from_directory, request, flash
import openpyxl as ox
import pandas as pd
import re
import os
from werkzeug.utils import secure_filename

upload_folder = "/Users/Дима/PycharmProjects/clear_txt"
allowed_extensions = set(["txt", "xlsx", "xlsm"])

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = upload_folder


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in allowed_extensions


def parse_txt_section(section):
    doc_number = re.search(r'Номер=(\d+)', section)
    amount = re.search(r'Сумма=([\d.]+)', section)
    purpose = re.search(r'НазначениеПлатежа=(.*)', section)
    if doc_number and amount and purpose:
        return {
            'doc_number': doc_number.group(1),
            'amount': float(amount.group(1)),
            'purpose': purpose.group(1)
        }
    return None


def load_excel_data(excel_path):
    df = pd.read_excel(excel_path, sheet_name=0, header=5) # header=5 поскольку названия колонок только на шестой строчке
    df.columns = df.columns.str.strip()
    relevant_data = df[['Вх. номер', 'Поступило', 'Назначение платежа']].dropna()
    return relevant_data


def filter_txt_file(txt_path, excel_path, output_path):
    with open(txt_path, 'r', encoding='cp1251') as file:
        txt_content = file.read()
    excel_data = load_excel_data(excel_path)
    sections = re.split(r'(СекцияДокумент=.*)', txt_content)
    filtered_sections = []
    total_amount = 0.0
    for i in range(1, len(sections), 2):
        header = sections[i] # заголовок
        content = sections[i + 1] # содержимое
        section_data = parse_txt_section(content)
        if section_data and section_matches(section_data, excel_data):
            filtered_sections.append(header + content)
            total_amount += section_data['amount']

    header_content = sections[0]
    header_content = re.sub(r'ВсегоСписано=\d+[\d,.]*', 'ВсегоСписано=0.00', header_content)
    header_content = re.sub(r'ВсегоПоступило=\d+[\d,.]*', f'ВсегоПоступило={total_amount:.2f}', header_content)

    # Новый тхт
    with open(output_path, 'w', encoding='cp1251') as file:
        file.write(header_content)  # Первая часть до "СекцияДокумент"
        for section in filtered_sections:
            file.write(section)


def section_matches(section_data, excel_data):
    doc_number = section_data['doc_number']
    amount = section_data['amount']
    purpose = section_data['purpose'].strip()  # Убираем лишние пробелы в назначении
    """match = excel_data[(excel_data['Вх.номер'] == doc_number) &
                (excel_data['Поступило'] == amount) &
                (excel_data['Назначение платежа'] == purpose)]
        return not match.empty"""
    # Проход по каждой строке
    for idx, row in excel_data.iterrows():
        excel_doc_number = str(row['Вх. номер']).strip()
        excel_amount = re.sub('[,]', '', str(row['Поступило']))
        excel_purpose = str(row['Назначение платежа']).strip()
        if doc_number == excel_doc_number and amount == float(excel_amount) and purpose == excel_purpose:
            print(f"Совпадение найдено для документа {doc_number}")
            return True


@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == 'POST':
        if 'txt_file' not in request.files and 'xlsx_file' not in request.files:
            flash('No txt part')
            return redirect(request.url)
        txt_file = request.files['txt_file']
        xlsx_file = request.files['xlsx_file']
        if txt_file.filename == '' or xlsx_file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        if txt_file and allowed_file(txt_file.filename) and xlsx_file and allowed_file(xlsx_file.filename):
            filename_txt = secure_filename(txt_file.filename)
            txt_file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename_txt))
            filename_xlsx = secure_filename(xlsx_file.filename)
            xlsx_file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename_xlsx))
        return redirect(url_for('download_results', name_xlsx=filename_xlsx, name_txt=filename_txt))
    return render_template("Start.html")


@app.route("/results/<name_txt>/<name_xlsx>", methods=["GET", "POST"])
def download_results(name_xlsx, name_txt):
    if request.method == 'POST':
        filter_txt_file(f"./{name_txt}", f"./{name_xlsx}", "./new_txt.txt")
        return send_file(f"./new_txt.txt", as_attachment=True)
    return render_template("results.html")


if __name__ == "__main__":
    app.run(debug=True)
