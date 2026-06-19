import os
import re
import tempfile
from flask import Flask, request, render_template, send_file
import pypandoc
import docx
from docx.shared import Pt

app = Flask(__name__)


def generate_reference_doc(font_name, font_size):
    """Динамически создает референс-файл с кастомным шрифтом и размером."""
    doc = docx.Document()

    # Настройка основного текста (стиль Normal)
    style_normal = doc.styles['Normal']
    style_normal.font.name = font_name
    style_normal.font.size = Pt(float(font_size))

    # Настройка шрифта для заголовков, чтобы они не выбивались из общего стиля
    for i in range(1, 5):
        heading_style = f'Heading {i}'
        if heading_style in doc.styles:
            doc.styles[heading_style].font.name = font_name

    # Создаем временный файл, который удалится после использования
    temp_dir = tempfile.gettempdir()
    template_path = os.path.join(temp_dir, f"template_{font_name}_{font_size}.docx")
    doc.save(template_path)
    return template_path


def process_text_to_word(raw_text, output_path, font_name, font_size):
    # Очистка синтаксиса формул для Pandoc
    clean_text = re.sub(r'\\\[\s*', '$$', raw_text)
    clean_text = re.sub(r'\s*\\\]', '$$', clean_text)
    clean_text = re.sub(r'\\\(\s*', '$', clean_text)
    clean_text = re.sub(r'\s*\\\)', '$', clean_text)

    # Генерируем шаблон под требования пользователя
    template_path = generate_reference_doc(font_name, font_size)
    extra_args = [f'--reference-doc={template_path}']

    try:
        pypandoc.convert_text(
            clean_text,
            'docx',
            format='markdown',
            outputfile=output_path,
            extra_args=extra_args
        )
    finally:
        # Удаляем временный шаблон
        if os.path.exists(template_path):
            os.remove(template_path)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        raw_text = ""

        # Получаем настройки форматирования из формы
        font_name = request.form.get('font_name', 'Calibri')
        font_size = request.form.get('font_size', '9')

        # 1. Проверяем загрузку файла
        if 'file' in request.files and request.files['file'].filename != '':
            uploaded_file = request.files['file']
            ext = uploaded_file.filename.rsplit('.', 1)[-1].lower()

            try:
                if ext in ['txt', 'md']:
                    raw_text = uploaded_file.read().decode('utf-8')
                elif ext == 'docx':
                    doc = docx.Document(uploaded_file)
                    raw_text = "\n".join([p.text for p in doc.paragraphs])
                else:
                    return "Формат файла не поддерживается. Разрешены .txt, .md, .docx", 400
            except Exception as e:
                return f"Ошибка чтения файла: {e}", 500

        # 2. Если файла нет, берем текст из textarea
        else:
            raw_text = request.form.get('text_input', '')

        if not raw_text.strip():
            return "Нет данных для обработки. Загрузите файл или вставьте текст.", 400

        output_filename = "converted_result.docx"

        try:
            process_text_to_word(raw_text, output_filename, font_name, font_size)
            return send_file(
                output_filename,
                as_attachment=True,
                download_name="result.docx"
            )
        except Exception as e:
            return f"Ошибка при конвертации: {e}", 500

    return render_template('index.html')


if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)