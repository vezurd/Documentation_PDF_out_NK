# Для считывания PDF
import PyPDF2
# Для анализа структуры PDF и извлечения текста
from pdfminer.high_level import extract_pages, extract_text
from pdfminer.layout import LTTextContainer, LTChar, LTRect, LTFigure
# Для извлечения текста из таблиц в PDF
import pdfplumber
# Для извлечения изображений из PDF
from PIL import Image
from pdf2image import convert_from_path
# Для выполнения OCR, чтобы извлекать тексты из изображений
import pytesseract
# Для удаления дополнительно созданных файлов
import os
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def text_extraction(element):
    # Извлекаем текст из вложенного текстового элемента
    line_text = element.get_text()

    # Находим форматы текста
    # Инициализируем список со всеми форматами, встречающимися в строке текста
    line_formats = []
    for text_line in element:
        if isinstance(text_line, LTTextContainer):
            # Итеративно обходим каждый символ в строке текста
            for character in text_line:
                if isinstance(character, LTChar):
                    # Добавляем к символу название шрифта
                    line_formats.append(character.fontname)
                    # Добавляем к символу размер шрифта
                    line_formats.append(character.size)
    # Находим уникальные размеры и названия шрифтов в строке
    format_per_line = list(set(line_formats))

    # Возвращаем кортеж с текстом в каждой строке вместе с его форматом
    return (line_text, format_per_line)

# Создаём функцию для вырезания элементов изображений из PDF
def crop_image(element, pageObj, file_name_save):
    # Получаем координаты для вырезания изображения из PDF
    [image_left, image_top, image_right, image_bottom] = [element.x0,element.y0,element.x1,element.y1]
    # Обрезаем страницу по координатам (left, bottom, right, top)
    pageObj.mediabox.lower_left = (image_left, image_bottom)
    pageObj.mediabox.upper_right = (image_right, image_top)
    # Сохраняем обрезанную страницу в новый PDF
    cropped_pdf_writer = PyPDF2.PdfWriter()
    cropped_pdf_writer.add_page(pageObj)
    # Сохраняем обрезанный PDF в новый файл
    with open(file_name_save, 'wb') as cropped_pdf_file:
        cropped_pdf_writer.write(cropped_pdf_file)

# Создаём функцию для преобразования PDF в изображения
def convert_to_images(input_file,):
    images = convert_from_path(input_file, 500 ,poppler_path=r'C:\Users\druze\AppData\Local\Programs\Python\Python312\poppler-24.02.0\Library\bin')
    image = images[0]
    output_file = "PDF_image.png"
    image.save(output_file, "PNG")

# Создаём функцию для считывания текста из изображений
def image_to_text(image_path):
    # Считываем изображение
    img = Image.open(image_path)
    # Извлекаем текст из изображения
    text = pytesseract.image_to_string(img)
    return text

# Извлечение таблиц из страницы

def extract_table(pdf_path, page_num, table_num):
    # Открываем файл pdf
    pdf = pdfplumber.open(pdf_path)
    # Находим исследуемую страницу
    table_page = pdf.pages[page_num]
    # Извлекаем соответствующую таблицу
    table = table_page.extract_tables()[table_num]
    return table

def pdf_get_text_from_file (pdf_path):
    reader11 = PyPDF2.PdfReader(pdf_path)
    #number_of_pages = len(reader.pages)
    page11 = reader11.pages[0]
    text11 = page11.extract_text()
    return text11

def table_converter(table):
    table_string = ''
    # Итеративно обходим каждую строку в таблице
    for row_num in range(len(table)):
        row = table[row_num]
        # Удаляем разрыв строки из текста с переносом
        cleaned_row = [item.replace('\n', ' ') if item is not None and '\n' in item else 'None' if item is None else item for item in row]
        # Преобразуем таблицу в строку
        table_string+=('|'+'|'.join(cleaned_row)+'|'+'\n')
    # Удаляем последний разрыв строки
    table_string = table_string[:-1]
    return table_string

def check_element_position (range, element):

    if (element.y0 >= range.y0 and element.y1 <= range.y1 and element.x1 <= range.x1 and element.x0 >= range.x0):
        return 1
    elif not (element.x1 < range.x0
                or element.x0 > range.x1
                or element.y1 < range.y0
                or element.y0 > range.y1):
        return 1
    else:
        return 0

def stop_until_enter():
    try:
        from msvcrt import getch
    except ImportError:
        import sys
        import tty, termios
        def getch():
            fd = sys.stdin.fileno()
            old_settings = termios.tcgetattr(fd)
            try:
                tty.setraw(sys.stdin.fileno())
                ch = sys.stdin.read(1)
            finally:
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
            return ch

    def stop(symbol, message):
        while True:
            print(message)
            if getch() != "":
                break

    stop(b'\r', 'Press Enter to exit')  # Первым параметром нужный вам символ