from flask import Flask, render_template, request, send_file, jsonify
from datetime import datetime
import io
import traceback

from analiz_core import run_analysis

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 МБ максимум


# ── Глобальные обработчики HTTP-ошибок ────────────────────────────────────────
@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'Маршрут не найден'}), 404

@app.errorhandler(405)
def method_not_allowed(e):
    return jsonify({'error': 'Метод не разрешён'}), 405

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'Файл слишком большой (максимум 50 МБ)'}), 413

@app.errorhandler(500)
def server_error(e):
    return jsonify({'error': f'Внутренняя ошибка сервера: {e}'}), 500


# ── Вспомогательные функции ───────────────────────────────────────────────────
def _parse_dates():
    """Читает date_from / date_to из form, возвращает (date, date) или бросает ValueError."""
    try:
        date_from = datetime.strptime(request.form['date_from'], '%Y-%m-%d').date()
        date_to   = datetime.strptime(request.form['date_to'],   '%Y-%m-%d').date()
    except (KeyError, ValueError):
        raise ValueError('Укажите корректные даты в формате ГГГГ-ММ-ДД')
    return date_from, date_to

def _get_file_bytes():
    """Проверяет наличие файла в запросе и возвращает bytes."""
    if 'file' not in request.files:
        raise ValueError('Поле file отсутствует в запросе')
    f = request.files['file']
    if not f or f.filename == '':
        raise ValueError('Файл не выбран')
    if not f.filename.lower().endswith(('.xlsx', '.xlsm')):
        raise ValueError('Поддерживаются только файлы .xlsx и .xlsm')
    return f.read()


# ── Маршруты ──────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/preview', methods=['POST'])
def preview():
    """Возвращает только статистику (JSON) — без генерации Excel."""
    try:
        file_bytes        = _get_file_bytes()
        date_from, date_to = _parse_dates()
        _, stats          = run_analysis(file_bytes, date_from, date_to)
        return jsonify({'ok': True, 'stats': stats})
    except ValueError as e:
        return jsonify({'error': str(e)}), 422
    except Exception:
        return jsonify({'error': traceback.format_exc()}), 500


@app.route('/analyze', methods=['POST'])
def analyze():
    """Запускает анализ и возвращает готовый Excel-файл."""
    try:
        file_bytes        = _get_file_bytes()
        date_from, date_to = _parse_dates()
        excel_bytes, _    = run_analysis(file_bytes, date_from, date_to)
    except ValueError as e:
        return jsonify({'error': str(e)}), 422
    except Exception:
        return jsonify({'error': traceback.format_exc()}), 500

    filename = f"результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        io.BytesIO(excel_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    app.run(debug=True)
