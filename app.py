from flask import Flask, render_template, request, send_file, jsonify
from datetime import datetime, date
import io
import traceback

from analiz_core import run_analysis

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 МБ максимум


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/analyze', methods=['POST'])
def analyze():
    # Проверяем наличие файла
    if 'file' not in request.files or request.files['file'].filename == '':
        return jsonify({'error': 'Файл не выбран'}), 400

    file = request.files['file']
    if not file.filename.lower().endswith(('.xlsx', '.xlsm')):
        return jsonify({'error': 'Поддерживаются только файлы .xlsx и .xlsm'}), 400

    # Проверяем даты
    try:
        date_from = datetime.strptime(request.form['date_from'], '%Y-%m-%d').date()
        date_to   = datetime.strptime(request.form['date_to'],   '%Y-%m-%d').date()
    except (KeyError, ValueError):
        return jsonify({'error': 'Укажите корректные даты'}), 400

    # Запускаем анализ
    try:
        file_bytes   = file.read()
        excel_bytes, stats = run_analysis(file_bytes, date_from, date_to)
    except ValueError as e:
        return jsonify({'error': str(e)}), 422
    except Exception:
        return jsonify({'error': 'Ошибка обработки файла:\n' + traceback.format_exc()}), 500

    # Возвращаем Excel как скачиваемый файл
    filename = f"результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/preview', methods=['POST'])
def preview():
    """Возвращает только статистику без скачивания файла — для показа итогов на странице."""
    if 'file' not in request.files or request.files['file'].filename == '':
        return jsonify({'error': 'Файл не выбран'}), 400

    file = request.files['file']
    try:
        date_from = datetime.strptime(request.form['date_from'], '%Y-%m-%d').date()
        date_to   = datetime.strptime(request.form['date_to'],   '%Y-%m-%d').date()
        file_bytes = file.read()
        _, stats = run_analysis(file_bytes, date_from, date_to)
        return jsonify({'ok': True, 'stats': stats})
    except ValueError as e:
        return jsonify({'error': str(e)}), 422
    except Exception:
        return jsonify({'error': traceback.format_exc()}), 500


if __name__ == '__main__':
    app.run(debug=False)
