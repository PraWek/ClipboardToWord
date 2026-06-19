import pyperclip
import pypandoc
import os


def clipboard_to_word(output_filename="solution.docx"):
    raw_text = pyperclip.paste()

    if not raw_text.strip():
        print("Буфер обмена пуст")
        return

    clean_text = raw_text.replace(r'\[', '$$').replace(r'\]', '$$')
    clean_text = clean_text.replace(r'\(', '$').replace(r'\)', '$')

    template_path = "../template.docx"

    if not os.path.exists(template_path):
        print(f"Внимание: файл шаблона {template_path} не найден. Использую стандартные настройки.")
        extra_args = []
    else:
        extra_args = [f'--reference-doc={template_path}']

    try:
        pypandoc.convert_text(
            clean_text,
            'docx',
            format='markdown',
            outputfile=output_filename,
            extra_args=extra_args
        )
        print(f"Успех! Файл '{output_filename}' создан с использованием шаблона '{template_path}'.")

    except Exception as e:
        print(f"Произошла ошибка: {e}")


if __name__ == "__main__":
    clipboard_to_word()