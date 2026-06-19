import os
import re  # Добавляем модуль для регулярных выражений
from flask import Flask, request, render_template, send_file
import pypandoc
import docx

app = Flask(__name__)


def process_text_to_word(raw_text, output_path):
    # Используем регулярные выражения, чтобы заменить скобки на доллары
    # И заодно удаляем пробелы СРАЗУ после открывающей скобки и ПЕРЕД закрывающей

    # Для блочных формул \[ ... \] -> $$...$$
    clean_text = re.sub(r'\\\[\s*', '$$', raw_text)
    clean_text = re.sub(r'\s*\\\]', '$$', clean_text)

    # Для строчных формул \( ... \) -> $...$
    clean_text = re.sub(r'\\\(\s*', '$', clean_text)
    clean_text = re.sub(r'\s*\\\)', '$', clean_text)

    template_path = "template.docx"
    extra_args = []
    if os.path.exists(template_path):
        extra_args = [f'--reference-doc={template_path}']

    # Конвертация через Pandoc
    pypandoc.convert_text(
        clean_text,
        'docx',
        format='markdown',
        outputfile=output_path,
        extra_args=extra_args
    )


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        raw_text = ""

        # 1. Проверяем, был ли загружен файл
        if 'file' in request.files and request.files['file'].filename != '':
            uploaded_file = request.files['file']
            ext = uploaded_file.filename.rsplit('.', 1)[-1].lower()

            try:
                if ext in ['txt', 'md']:
                    raw_text = uploaded_file.read().decode('utf-8')
                elif ext == 'docx':
                    # Вытаскиваем текст из загруженного Word-документа
                    doc = docx.Document(uploaded_file)
                    raw_text = "\n".join([p.text for p in doc.paragraphs])
                else:
                    return "Формат не поддерживается. Разрешены .txt, .md, .docx", 400
            except Exception as e:
                return f"Ошибка чтения файла: {e}", 500

        # 2. Если файла нет, берем текст из поля ввода
        else:
            raw_text = request.form.get('text_input', '')

        if not raw_text.strip():
            return "Нет текста для обработки. Загрузите файл или вставьте текст.", 400

        # 3. Обработка и генерация файла
        output_filename = "converted_result.docx"

        try:
            process_text_to_word(raw_text, output_filename)
            # Отправляем готовый файл пользователю (скачивание)
            return send_file(
                output_filename,
                as_attachment=True,
                download_name="Math_Result.docx"
            )
        except Exception as e:
            return f"Ошибка при конвертации Pandoc: {e}", 500

    return render_template('index.html')


if __name__ == "__main__":
    # Запуск сервера на порту 5000
    app.run(debug=True, host='0.0.0.0', port=5000)