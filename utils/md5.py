import os
import hashlib
import datetime
from docx import Document
from docx.shared import Pt
import locale

# Установка русской локали для форматирования даты
try:
    locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Russian_Russia.1251')
    except:
        print("Предупреждение: Не удалось установить русскую локаль")


def calculate_md5(file_path):
    """Вычисляет MD5 хеш файла"""
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception as e:
        print(f"Ошибка при вычислении MD5 для {file_path}: {e}")
        return None


def format_file_size(size_bytes):
    """Форматирует размер файла в читаемый вид"""
    try:
        return f"{size_bytes:,}".replace(",", " ")
    except:
        return str(size_bytes)


def format_date(timestamp):
    """Форматирует дату в русском формате"""
    try:
        dt = datetime.datetime.fromtimestamp(timestamp)
        month_names = {
            1: "января", 2: "февраля", 3: "марта", 4: "апреля",
            5: "мая", 6: "июня", 7: "июля", 8: "августа",
            9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"
        }

        day = dt.day
        month = month_names[dt.month]
        year = dt.year
        time_str = dt.strftime("%H:%M")

        return f"{day:02d} {month} {year} г., {time_str}"
    except Exception as e:
        print(f"Ошибка форматирования даты: {e}")
        return dt.strftime("%d %B %Y г., %H:%M") if 'dt' in locals() else "Неизвестная дата"


def format_filename(filename):
    """Форматирует имя файла согласно требованиям"""
    return f"Раздел ПД №5. Подраздел №7. {filename}"


def find_pdf_files(root_dir):
    """Находит все PDF файлы в директории и поддиректориях, исключая файлы с '-УЛ'"""
    pdf_files = []

    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith('.pdf') and '-УЛ' not in file.upper():
                full_path = os.path.join(root, file)
                pdf_files.append(full_path)
                print(f"Найден PDF файл: {full_path}")

    return pdf_files


def process_pdf_files(root_dir, template_file):
    """Основная функция обработки PDF файлов"""
    # Поиск PDF файлов во всех подкаталогах
    pdf_files = find_pdf_files(root_dir)

    print(f"Найдено PDF файлов для обработки: {len(pdf_files)}")

    if not pdf_files:
        print("PDF файлы не найдены")
        return

    # Загрузка шаблона документа
    try:
        doc = Document(template_file)
        print("Шаблон документа загружен")
    except Exception as e:
        print(f"Ошибка загрузки шаблона: {e}")
        return

    # Собираем информацию о файлах
    file_info_list = []
    md5_hashes = []

    for pdf_file in pdf_files:
        try:
            # Получаем информацию о файле
            file_stats = os.stat(pdf_file)
            file_name = os.path.basename(pdf_file)

            # Вычисляем MD5
            md5_hash = calculate_md5(pdf_file)
            if md5_hash:
                md5_hashes.append(md5_hash)
                print(f"MD5 для {file_name}: {md5_hash}")

            # Форматируем данные
            formatted_name = format_filename(file_name)
            formatted_date = format_date(file_stats.st_ctime)
            formatted_size = format_file_size(file_stats.st_size)

            file_info_list.append({
                'formatted_name': formatted_name,
                'formatted_date': formatted_date,
                'formatted_size': formatted_size,
                'file_path': pdf_file
            })

            print(f"Обработан: {file_name}")

        except Exception as e:
            print(f"Ошибка обработки файла {pdf_file}: {e}")

    # Обновляем документ
    try:
        print("Начинаем обновление документа...")

        # Обновляем ячейку с MD5 хешами
        md5_updated = False
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if '<место для md5>' in cell.text or 'md5' in cell.text.lower():
                        # Очищаем ячейку и добавляем MD5 хеши
                        cell.text = ''
                        paragraph = cell.paragraphs[0]
                        for md5_hash in md5_hashes:
                            run = paragraph.add_run(md5_hash + '\n')
                            run.font.size = Pt(10)
                        md5_updated = True
                        print("Ячейка MD5 обновлена")

        # Обновляем строки с информацией о файлах
        for table in doc.tables:
            for i, row in enumerate(table.rows):
                for cell in row.cells:
                    if '<форматированное имя файла>' in cell.text:
                        print(f"Найдена строка для обновления в строке {i}")

                        # Находим индекс строки с placeholder
                        placeholder_row_index = i

                        # Обновляем информацию для каждого файла
                        for j, file_info in enumerate(file_info_list):
                            row_index = placeholder_row_index + j
                            if row_index < len(table.rows):
                                # Обновляем ячейки в строке
                                if len(table.rows[row_index].cells) >= 3:
                                    # Имя файла
                                    table.rows[row_index].cells[0].text = file_info['formatted_name']
                                    # Дата
                                    table.rows[row_index].cells[1].text = file_info['formatted_date']
                                    # Размер
                                    table.rows[row_index].cells[2].text = file_info['formatted_size']
                                    print(f"Обновлена строка {row_index} для файла {file_info['formatted_name']}")

        # Сохраняем обновленный документ
        doc.save(template_file)
        print(f"Документ успешно обновлен: {template_file}")

        # Выводим сводку
        print(f"\nСводка обработки:")
        print(f"Обработано файлов: {len(file_info_list)}")
        print(f"MD5 хешей: {len(md5_hashes)}")
        for info in file_info_list:
            print(f"  - {info['formatted_name']}")

    except Exception as e:
        print(f"Ошибка при обновлении документа: {e}")


def main(m_path):


    # Имя шаблонного файла
    template_file = "AGCC.0091-04-0000-ИОС7.12.2_0_RU-УЛ.DOCX"

    # Проверяем существование шаблонного файла
    if not os.path.exists(template_file):
        print(f"Ошибка: Файл шаблона {template_file} не найден в текущей директории")
        print(f"Текущая директория: {m_path}")
        return

    print(f"Начинаем поиск PDF файлов в: {m_path}")
    print("Исключаем файлы с '-УЛ' в имени")

    # Обрабатываем PDF файлы
    process_pdf_files(m_path, template_file)


if __name__ == "__main__":
    m_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\ПД\0091-04\WORK\v_05"
    main(m_path)