# Автоматизация обработки PDF-файлов

## Описание
Эта программа автоматически извлекает информацию о файлах (дата изменения, CRC32-хеш, размер) и вставляет эти данные в шаблон документа. Затем создаётся PDF-документ для каждого обработанного файла.

## Установка и использование

### 1. Подготовка
Создайте структуру папок:
```
PDF_automation/
│-- main.py
│-- pdf_files/      # Сюда помещаем файлы "Том1", "Том2", ...
│-- template.docx   # Шаблон документа (ИУЛ.docx)
```

### 2. Запуск программы
Запустите скрипт:
```sh
python main.py
```

После выполнения создастся папка `output/`, в которой будут готовые PDF-файлы.

## Сборка `.exe` (для Windows)
Если нужно запустить без Python, соберите исполняемый файл:
```sh
pyinstaller --onefile --add-data "template.docx;." main.py
```
После этого файл `main.exe` появится в папке `dist/`.

## Зависимости
Установите перед использованием:
```sh
pip install python-docx PyMuPDF zlib
```

## Контакты
Автор: [wordandlife](https://github.com/worldandlife)
