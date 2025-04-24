"""Модуль для извлечения задач из документа Word и сохранения их в Excel."""

import argparse
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import pandas as pd
from docx import Document

NUMBERED_VARIANT = r'(?:^|\s+)(\d+)\)\s*([^;.]+?)(?=(?:\s+\d+\)|$|\.|;))'
LETTER_VARIANT = r'(?:^|\s+)([а-яё])\)\s*([^;]*?)(?=(?:\s+[а-яё]\)|$|\.|;))'
INLINE_NUMBERED = r'^(\d+)\.\s*(\d+)\)\s*(.+)'
INLINE_LETTER = r'^(\d+)\.\s*([а-яё])\)\s*(.+)'
NEW_LINE_NUMBERED = r'^\s*(\d+)\)\s+(.+)'
NEW_LINE_LETTER = r'^\s*([а-яё])\)\s*(.+)'


def extract_toc(docx_path):
    """Извлекает содержание из документа Word."""
    toc = []
    with zipfile.ZipFile(docx_path, 'r') as docx:
        xml_content = docx.read('word/document.xml')

    tree = ElementTree.ElementTree(ElementTree.fromstring(xml_content))
    root = tree.getroot()

    inside_toc = False
    current_id = 1

    for para in root.iter():
        if para.tag.endswith('p'):
            text = ''.join(
                node.text if node.text else ''
                for node in para.iter()
                if node.tag.endswith('t')
            ).strip()

            if 'Оглавление' in text:
                inside_toc = True
                continue

            if inside_toc and re.match(r'^\d+(\.\d+)*[\.\s\t]+', text):
                section_match = re.match(
                    r'^(\d+(?:\.\d+)*)[\.\s\t]+(.+?)(?:\s+\d+)?$', text
                )
                if section_match:
                    section_num = section_match.group(1)
                    section_name = section_match.group(2).strip()
                    parent = (
                        next(
                            (
                                item['id']
                                for item in toc
                                if item.get('temp')
                                == section_num.rsplit('.', 1)[0]
                            ),
                            0,
                        )
                        if '.' in section_num
                        else 0
                    )

                    toc.append(
                        {
                            'id': current_id,
                            'name': section_name,
                            'parent': parent,
                            'temp': section_num,
                        }
                    )
                    current_id += 1

            elif inside_toc and not re.match(r'^\d+(\.\d+)*[\.\s\t]+', text):
                break

    for item in toc:
        item.pop('temp', None)

    return toc


def extract_answers_from_line(line):
    """Извлекает ответы."""
    answers = {}
    parts = re.split(r'(?<=\.)\s+(?=\d+\.)', line)

    for part in parts:
        task_match = re.match(r'(\d+)\.(.+)', part)
        if not task_match:
            continue

        task_num = task_match.group(1)
        answer_text = task_match.group(2).strip()

        for pattern in [NUMBERED_VARIANT, LETTER_VARIANT]:
            variants = list(re.finditer(pattern, answer_text, re.IGNORECASE))
            if variants:
                for var_match in variants:
                    variant = var_match.group(1)
                    answer = var_match.group(2).strip()
                    task_id = f'{task_num}.{variant}'
                    answers[task_id] = answer
                break
        else:
            if answer_text and not re.search(
                r'^\s*[а-яё0-9]\)', answer_text, re.IGNORECASE
            ):
                answers[task_num] = re.sub(r'\.$', '', answer_text).strip()

    return answers


def create_task(main_id, variant, task_text, is_subtask=True):
    """Создает словарь с задачей."""
    task_id = f'{main_id}.{variant}' if variant else main_id
    task_text = f'\t{task_text}' if is_subtask else task_text
    return {
        'id_tasks_book': task_id,
        'task': task_text,
        'answer': 'Отсутствует',
        'classes': '5;6',
        'topic_id': 1,
        'level': 1,
    }


def extract_tasks_from_docx(file_path):
    """Извлекает задачи и ответы из документа Word."""
    doc = Document(file_path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tasks = []
    current_main_id = None
    answers_section = False
    all_answers = {}

    for i, line in enumerate(lines):
        line = re.sub(r'\s+', ' ', line.strip())

        if line == 'Ответы и советы':
            answers_section = True
            continue

        if not answers_section:
            if re.match(r'^\d+\.\d+\.', line):
                continue

            main_match = re.match(r'^(\d+)\.\s*(.*)', line)
            if main_match:
                current_main_id = main_match.group(1)
                suffix = main_match.group(2).strip()

                is_header = i + 1 < len(lines) and re.match(
                    rf'^{current_main_id}\.\d+\.',
                    lines[i + 1].strip(),
                )
                if is_header:
                    continue

                for pattern, group_idx in [
                    (INLINE_LETTER, 2),
                    (INLINE_NUMBERED, 2),
                ]:
                    variant_match = re.match(pattern, line, re.IGNORECASE)
                    if variant_match:
                        variant = variant_match.group(group_idx)
                        task_text = variant_match.group(3).strip()
                        tasks.append(
                            create_task(current_main_id, variant, task_text)
                        )
                        break
                else:
                    if suffix:
                        tasks.append(
                            create_task(current_main_id, None, suffix, False)
                        )
                continue

            for pattern, group_idx in [
                (NEW_LINE_LETTER, 1),
                (NEW_LINE_NUMBERED, 1),
            ]:
                variant_match = re.match(pattern, line, re.IGNORECASE)
                if variant_match and current_main_id:
                    variant = variant_match.group(group_idx)
                    task_text = variant_match.group(2).strip()
                    tasks.append(
                        create_task(current_main_id, variant, task_text)
                    )
                    break
        else:
            answers = extract_answers_from_line(line)
            all_answers.update(answers)

    for task in tasks:
        task_id = task['id_tasks_book']
        if task_id in all_answers:
            task['answer'] = all_answers[task_id]

    return tasks


def create_author_table():
    """Создает таблицу author."""
    columns = ['name', 'author', 'description', 'topic_id', 'classes']
    return pd.DataFrame(columns=columns)


def save_tasks(tasks, toc, output_file):
    """Сохраняет задачи и оглавление."""
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_tasks = pd.DataFrame(tasks)
        columns = [
            'id_tasks_book',
            'task',
            'answer',
            'classes',
            'topic_id',
            'level',
        ]
        df_tasks = df_tasks[columns]
        df_tasks.to_excel(writer, sheet_name='tasks', index=False)

        df_author = create_author_table()
        df_author.to_excel(writer, sheet_name='author', index=False)

        df_toc = pd.DataFrame(toc)
        df_toc.to_excel(writer, sheet_name='table_of_contents', index=False)


def main():
    """Функция для обработки аргументов и запуска парсинга."""
    parser = argparse.ArgumentParser(
        description='Парсинг Word документа в Excel таблицу.'
    )
    parser.add_argument(
        'filename',
        type=str,
        help='Путь к Word документу (.docx) для обработки',
    )
    args = parser.parse_args()

    input_file = Path(args.filename)
    output_excel = 'output.xlsx'
    try:
        if not input_file.exists():
            raise FileNotFoundError(f'Файл {input_file} не существует!')

        if input_file.suffix.lower() != '.docx':
            raise ValueError(f'Файл {input_file} должен быть в формате .docx!')

        tasks = extract_tasks_from_docx(input_file)
        toc = extract_toc(input_file)

        save_tasks(tasks, toc, output_excel)

        print(
            'Извлечено {} задач(и) и {} пунктов оглавления,'.format(
                len(tasks), len(toc)
            ),
            'сохранено в {}'.format(output_excel),
        )

    except (FileNotFoundError, ValueError) as e:
        print(f'Ошибка: {e}')
    except Exception as e:
        print(f'Непредвиденная ошибка при обработке файла: {e}')


if __name__ == '__main__':
    main()
