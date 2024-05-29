import PDF_check_proc
import files_att_collector
import rules_check
import OD_tab_parsing
import folder_select

import time


if __name__ == '__main__':
    debug_print_excel_final_output = 1 #Выводит в эксель результат

    debug_print_time_of_execution = 1 #Выводит время выполнения программы

    debug_first_from_file = 0
    debug_second_from_file = 0

    debug_Curr_Proj_file = r"t__result_output_8525-SKUD.xlsx"
    debug_proj_OD_list_file = r"t__result_OD_output_8525-SKUD.xlsx"

    excel_template_file_name = "out_template.xlsx"




    #pdf_path = r'test_Project'

    if not (debug_first_from_file and debug_second_from_file):
        pdf_path = folder_select.getDirectory()

    if debug_print_time_of_execution:
        start_time = time.time()  ## точка отсчета времени

    ####################################################################
    # Первый этап  - парсим штампы всех страниц всех докуметов в папке #
    ####################################################################
    print ("Первый этап  - собираем информацию из штампов всех страниц всех докуметов в папке...")
    if not debug_first_from_file:
        Curr_Proj = files_att_collector.files_attribute_collector_f(pdf_path)
    else:
        import read_Curr_Proj_from_file
        print ("\t", "Читаем из файла: <"+ debug_Curr_Proj_file+">")
        Curr_Proj = read_Curr_Proj_from_file.read_Curr_Proj_from_excel(debug_Curr_Proj_file,1)
    print ("Первый этап  - окончен", "\n")

    ####################################################################
    # Второй этап  - парсим ведомости документов из общих данных       #
    ####################################################################
    print ("Второй этап  - ищем OD, собираем ведомости основных и прилагаемых доукументов...")
    if not debug_second_from_file:
        proj_OD_list = OD_tab_parsing.OD_table_parsing_f(pdf_path)
    else:
        print("\t", "Читаем из файла: <"+ debug_proj_OD_list_file+">")
        proj_OD_list = OD_tab_parsing.OD_table_from_file(debug_proj_OD_list_file)
    print ("Второй этап  - окончен", "\n")

    ####################################################################
    # Второй этап  - парсим ведомости документов из общих данных       #
    ####################################################################
    print ("Третий этап  - анализируем полученные данные...")
    check_list = rules_check.rules_check(Curr_Proj, proj_OD_list)
    print ("Третий этап  - окончен", "\n")

    #Выводим ошибки в консоль
        # print_check_mode == 0 - только ошибки, все проверки
        # print_check_mode == 1 - и ошибки и не ошибки, все проверки
        # print_check_mode == 2 - только ошибки, проверики по c_code
        # print_check_mode == 3 - и ошибки и не ошибки, проверики по c_code
    rules_check.print_check_list(check_list,2,1010)


    if debug_print_excel_final_output:
        from functions.excel_Cur_project_output import excel_Curr_Proj_output_f
        excel_Curr_Proj_output_f(Curr_Proj, excel_template_file_name)


    if debug_print_time_of_execution:
        end_time = time.time() - start_time ##время работы программы
        print("\nВремя работы программы: ", end_time)

    # print("press 'Esc' to quit...")
    # from msvcrt import getch
    # key=getch()
    # if ord(pressedKey) == 27:
    #    sys.exit()