# Третье тестовое задание для Яндекс Практикум.

## Описание

Скрипт для парсинга `.docx` документов и конвертации их структуры в таблицу Excel

## Установка

1. Клонируйте репозиторий:

    ```bash
    git clone https://github.com/AlexBrantt/YandexPracticumTZ3.git
    cd YandexPracticumTZ3
    ```

2. Создайте и активируйте виртуальное окружение:

    ```bash
    python -m venv venv
    source venv/bin/activate  # Для Linux/MacOS
    venv\Scripts\activate  # Для Windows
    ```

3. Установите зависимости:

    ```bash
    pip install -r requirements.txt
    ```

## Использование

Скрипт принимает путь к `.docx` файлу через аргументы командной строки и сохраняет результат в формате Excel.

Пример:

```bash
python parse_word_to_excel.py tekstovye_zadachi_po_matematike.docx
```

**Результат будет сохранён в файл output.xlsx**
