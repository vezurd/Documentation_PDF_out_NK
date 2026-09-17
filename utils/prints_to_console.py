def print_file_list_to_console(curr_proj):
    from prettytable import PrettyTable
    table = PrettyTable()
    table.field_names = ["п/н", "Полный путь файла", "Имя файла"]
    table.border = 0
    table.align = "l"
    row_index = 0
    for document in curr_proj:
        row_index += 1
        table.add_row([row_index, document.file_full_path, document.file_name])
    print(table)
