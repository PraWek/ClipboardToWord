import os
import re
import tempfile
import uuid
from flask import Flask, request, render_template, send_file
import pypandoc
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)


def generate_reference_doc(settings):
    """Создает референс-файл с расширенными метриками форматирования."""
    doc = docx.Document()

    # Настройка основного текста (стиль Normal)
    style_normal = doc.styles['Normal']

    # 1. Шрифт и размер
    style_normal.font.name = settings['font_name']
    style_normal.font.size = Pt(float(settings['font_size']))

    # 2. Абзац: межстрочный интервал и красная строка
    pf = style_normal.paragraph_format
    pf.line_spacing = float(settings['line_spacing'])
    pf.first_line_indent = Cm(float(settings['first_line_indent']))

    # 3. Выравнивание
    alignments = {
        'left': WD_ALIGN_PARAGRAPH.LEFT,
        'center': WD_ALIGN_PARAGRAPH.CENTER,
        'right': WD_ALIGN_PARAGRAPH.RIGHT,
        'justify': WD_ALIGN_PARAGRAPH.JUSTIFY
    }
    pf.alignment = alignments.get(settings['alignment'], WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 4. Поля страницы
    for section in doc.sections:
        section.left_margin = Cm(float(settings['margin_left']))
        section.right_margin = Cm(float(settings['margin_right']))
        section.top_margin = Cm(float(settings['margin_top']))
        section.bottom_margin = Cm(float(settings['margin_bottom']))

    # Синхронизация шрифта заголовков с основным текстом
    for i in range(1, 5):
        heading_style = f'Heading {i}'
        if heading_style in doc.styles:
            doc.styles[heading_style].font.name = settings['font_name']

    # Уникальное имя для временного файла (избегает конфликтов при одновременных запросах)
    temp_dir = tempfile.gettempdir()
    template_path = os.path.join(temp_dir, f"template_{uuid.uuid4().hex}.docx")
    doc.save(template_path)
    return template_path


def process_text_to_word(raw_text, output_path, settings):
    # Очистка синтаксиса формул для Pandoc
    clean_text = re.sub(r'\\\[\s*', '$$', raw_text)
    clean_text = re.sub(r'\s*\\\]', '$$', clean_text)
    clean_text = re.sub(r'\\\(\s*', '$', clean_text)
    clean_text = re.sub(r'\s*\\\)', '$', clean_text)

    # Генерация шаблона
    template_path = generate_reference_doc(settings)
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
        if os.path.exists(template_path):
            os.remove(template_path)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Сбор всех настроек форматирования из формы
        settings = {
            'font_name': request.form.get('font_name', 'Times New Roman'),
            'font_size': request.form.get('font_size', '14'),
            'line_spacing': request.form.get('line_spacing', '1.5'),
            'first_line_indent': request.form.get('first_line_indent', '1.25'),
            'alignment': request.form.get('alignment', 'justify'),
            'margin_left': request.form.get('margin_left', '3.0'),
            'margin_right': request.form.get('margin_right', '1.5'),
            'margin_top': request.form.get('margin_top', '2.0'),
            'margin_bottom': request.form.get('margin_bottom', '2.0')
        }

        raw_text = ""
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
                    return "Формат файла не поддерживается", 400
            except Exception as e:
                return f"Ошибка чтения файла: {e}", 500
        else:
            raw_text = request.form.get('text_input', '')

        if not raw_text.strip():
            return "Нет данных для обработки. Загрузите файл или вставьте текст.", 400

        output_filename = "converted_result.docx"
        try:
            process_text_to_word(raw_text, output_filename, settings)
            return send_file(
                output_filename,
                as_attachment=True,
                download_name="Clipboard_Result.docx"
            )
        except Exception as e:
            return f"Ошибка при конвертации: {e}", 500

    return render_template('index.html')


if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)