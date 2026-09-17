def xml_read_file(file_path):
    print("XML_READ_FILE start\n")
    import xml.etree.ElementTree as ET

    tree = ET.parse(file_path)
    root = tree.getroot()
    for elem in root.iter():
        print(elem.tag, elem.attrib, elem.text)

    for elem in root.findall('Value'):
        print(elem.attrib)


if __name__ in {"__main__"}:
    dir_path = r"C:\YandexDisk\темп\XML"
    file_path = r"C:\YandexDisk\темп\XML\Здание 1 Этаж 0.xml"

    open_path = ""
    if 1 == 0:
        print("Запуск проверки ТПК папки")
        # Окно выбора ПАПКИ
        dir_path = folder_select.getDirectory()
        print(dir_path, "<open_pdf_folder>")
        # Обработка исключения
        if is_file_path(dir_path) is False:
            exit(0)
    xml_read_file(file_path)
