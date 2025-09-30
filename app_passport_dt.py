import os
import re
import json
import tempfile
from flask import Flask, request, render_template_string, send_file, send_from_directory
import pandas as pd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 МБ

HTML = '''
<!doctype html>
<html>
<head>
    <meta charset="utf-8">
    <title>Генератор паспорта данных</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 30px auto; padding: 20px; }
        input, button { padding: 8px; margin: 5px 0; width: 100%; box-sizing: border-box; }
        button { background: #4CAF50; color: white; border: none; cursor: pointer; }
        button:hover { background: #45a049; }
        .error { color: red; background: #ffe6e6; padding: 10px; border-radius: 4px; }
        .success { color: green; background: #e6ffe6; padding: 10px; border-radius: 4px; }
        .template-link {
            display: inline-block;
            margin: 10px 0;
            color: #1E88E5;
            text-decoration: none;
            font-weight: bold;
        }
        .template-link:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <h2>Генерация JSON-паспорта из Excel</h2>
    <p>
        <a href="/download-template" class="template-link">📥 Скачать шаблон Excel</a>
    </p>
    <form method="post" enctype="multipart/form-data">
        <label>Выберите Excel-файл (.xlsx):</label>
        <input type="file" name="file" accept=".xlsx" required>
        <label>Имя листа (по умолчанию "Лист1"):</label>
        <input type="text" name="sheet_name" value="Лист1">
        <button type="submit">Сгенерировать JSON</button>
    </form>
    {% if message %}
        <hr>
        <div class="{{ 'error' if error else 'success' }}">{{ message|safe }}</div>
    {% endif %}
</body>
</html>
'''

def clean_excel_value(val):
    if pd.isna(val) or val is None:
        return None
    s = str(val).strip()
    return s if s != "" else None


@app.route('/download-template')
def download_template():
    try:
        return send_from_directory(
            directory=os.path.dirname(__file__),
            path="example_excel_table_dataset_passport.xlsx",
            as_attachment=True,
            download_name="example_excel_table_dataset_passport.xlsx"
        )
    except FileNotFoundError:
        return "Шаблон не найден на сервере.", 404



@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file')
        sheet_name = request.form.get('sheet_name', 'Лист1').strip()

        if not file or not file.filename.endswith('.xlsx'):
            return render_template_string(HTML, message="Пожалуйста, загрузите файл .xlsx", error=True)

        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, "input.xlsx")
            file.save(input_path)

            try:
                df = pd.read_excel(input_path, sheet_name=sheet_name, skiprows=10)
                additional_data = pd.read_excel(input_path, sheet_name=sheet_name, header=None)

                output_name = clean_excel_value(additional_data.iloc[1, 1])
                full_name = clean_excel_value(additional_data.iloc[2, 1])
                description = clean_excel_value(additional_data.iloc[3, 1])
                geo_type = clean_excel_value(additional_data.iloc[4, 1])
                periodicity = clean_excel_value(additional_data.iloc[5, 1])
                schedule = additional_data.iloc[6, 1]

                external_name_id = clean_excel_value(additional_data.iloc[1, 4])
                name_external = clean_excel_value(additional_data.iloc[2, 4])
                period_nm = clean_excel_value(additional_data.iloc[3, 4])
                analytical_com = clean_excel_value(additional_data.iloc[4, 4])

                # Проверка обязательных полей
                required = {
                    "B2 (output_name)": output_name,
                    "B3 (full_name)": full_name,
                    "B4 (description)": description,
                    "B5 (geo_type)": geo_type,
                    "B6 (periodicity)": periodicity,
                }
                missing = [k for k, v in required.items() if v is None]
                if missing:
                    raise ValueError("Не заполнены обязательные поля: " + ", ".join(missing))

                if geo_type not in {"Точка", "Линия", "Полигон"}:
                    raise ValueError(f"Недопустимое значение geo_type: {geo_type}")

                # Обработка атрибутов
                allowed_data_types = {"string", "integer", "double", "boolean", "datetime"}
                result_lines = []

                for idx, row in df.iterrows():
                    number = row.get('number', idx + 11)
                    code = str(row.get('code', '')).strip()
                    name = str(row.get('name', '')).strip()
                    datatype = str(row.get('datatype', '')).strip()
                    isNullable = str(row.get('isNullable', '')).strip()
                    isUnique = str(row.get('isUnique', '')).strip()

                    if not code or not re.match(r'^[a-zA-Z0-9_]+$', code):
                        raise ValueError(f"Строка {number}: недопустимый 'code'")
                    if datatype.lower() not in allowed_data_types:
                        raise ValueError(f"Строка {number}: недопустимый 'datatype'")
                    if isNullable.lower() not in {"true", "false"} or isUnique.lower() not in {"true", "false"}:
                        raise ValueError(f"Строка {number}: isNullable/isUnique должны быть true/false")

                    result_lines.append({
                        "code": code,
                        "name": name,
                        "dataType": datatype.capitalize(),
                        "isNullable": isNullable.lower() == 'true',
                        "isUnique": isUnique.lower() == 'true'
                    })

                # Доп. атрибуты
                if period_nm == 'Да':
                    result_lines.append({"code": "period_nm", "name": "Наименование анализируемого периода года", "dataType": "String", "isNullable": True, "isUnique": False})
                result_lines.append({
                    "code": f"external_{external_name_id or 'ext'}_id",
                    "name": name_external or "Внешний идентификатор",
                    "dataType": "String",
                    "isNullable": True,
                    "isUnique": False
                })
                if analytical_com == 'Да':
                    result_lines.append({"code": "analytical_committee_num", "name": "Номер аналитического комитета", "dataType": "String", "isNullable": True, "isUnique": False})
                result_lines.append({"code": "create_dttm", "name": "Дата и время формирования новой версии данных", "dataType": "DateTime", "isNullable": True, "isUnique": False})

                # Schedule
                if hasattr(schedule, 'hour'):
                    schedule_str = f"{schedule.hour:02}:{schedule.minute:02}"
                elif isinstance(schedule, str) and len(schedule) >= 5:
                    schedule_str = schedule[:5]
                else:
                    schedule_str = "00:00"

                geo_map = {"Точка": "Point", "Линия": "MultiLineString", "Полигон": "MultiPolygon"}
                geometry_type = geo_map[geo_type]

                json_output = {
                    "datasetData": {
                        "mainData": {
                            "fullName": full_name,
                            "description": description,
                            "oiv": "Инновационный центр «Безопасный транспорт»",
                            "informationSystem": {
                                "fullName": "Автоматизированная система персональных коммуникаций на основе использования больших данных",
                                "shortName": "АС ПКБД",
                                "regNumber": "-",
                                "url": "curl --location -X GET --resolve dp-apigw.ic.mosmetro.ru:9080:10.204.0.243 'http://dp-apigw.ic.mosmetro.ru:9080/api/v1/personal_mobility_device/bike_park_route_geojson' --header 'Authorization: Bearer ${{API_TOKEN}}'",
                                "ip": ""
                            },
                            "responsiblePerson": {
                                "fio": "Петров В.В.",
                                "position": "Советник руководителя",
                                "email": "PetrovVV@transport.mos.ru",
                                "phone": "+7 926 206 8246"
                            },
                            "technicalSupport": {
                                "email": "ic_bd_support@transport.mos.ru",
                                "phone": "+7 926 206 8246"
                            }
                        },
                        "updateParams": {"periodicity": periodicity, "schedule": schedule_str},
                        "geoData": {"srid": "WGS 84", "type": geometry_type}
                    },
                    "datasetAttributes": result_lines
                }

                safe_name = re.sub(r'[^\w\-_]', '_', output_name)
                json_filename = f"DataHub_passport_CODD_{safe_name}.json"
                json_path = os.path.join(tmpdir, json_filename)

                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump(json_output, f, ensure_ascii=False, indent=4)

                return send_file(json_path, as_attachment=True, download_name=json_filename)

            except Exception as e:
                return render_template_string(HTML, message=f"❌ Ошибка:<br>{str(e)}", error=True)

    return render_template_string(HTML)

# --- ЗАПУСК НА RENDER ---
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)