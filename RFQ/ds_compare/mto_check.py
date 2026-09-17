from types import NoneType

from base.base_cheks import tag_vs_value_sub_check, check_position_row
import base.t_comm_initial_classes as t_com_init_cls
from base.base_classes import RowStd, PROHIBITION_LIST
from base.base_google import load_base
from base.tables_columns import *
from utils.colors import Color
from utils.prints_to_console import print_file_list_to_console
from utils.string_parsing import print_att_list_table
import utils.path
import RFQ.ds_compare.load_mto
from colorama import Fore, Back, Style, init

# Инициализация colorama для Windows
init(autoreset=True)

class MtoCheck:
    def __init__(self, input_base: dict[str:list[RowStd]]):
        self.input_base = input_base
        self.errors = []
        self.stats = {
            'total_specs': 0,
            'total_rows': 0,
            'checked_rows': 0,
            'empty_specs': 0
        }

    def check(self):
        print(f"\n{Fore.CYAN}{'='*60}")
        print(f"{Fore.CYAN}🔍 ПРОВЕРКА МТО НА ТЕГИ ЗАПУЩЕНА")
        print(f"{Fore.CYAN}{'='*60}")
        
        # Проверка на пустые данные
        if not self.input_base:
            print(f"\n{Fore.RED}❌ {Style.BRIGHT}ОШИБКА: Словарь спецификаций пуст!")
            print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
            return
        
        # Подсчет статистики
        self._calculate_stats()
        self._print_stats()
        
        #Проверка 1 на соответсвие тегам
        for spec_name, base in self.input_base.items():
            if not base:  # Проверка на пустую спецификацию
                self.stats['empty_specs'] += 1
                print(f"\n{Fore.YELLOW}⚠️  {Style.BRIGHT}ВНИМАНИЕ: Спецификация '{spec_name}' пуста!")
                continue
            self._check_tags_value(base, spec_name)
        # Проверка 2 на значения VALUES
        for spec_name, base in self.input_base.items():
            if not base:  # Пропускаем пустые спецификации
                continue
            self._check_value_to_float(base, spec_name)
        # Проверка 3 на Старые коды (Google-база один раз на все спецификации)
        google_base_std = load_base()
        for spec_name, base in self.input_base.items():
            if not base:  # Пропускаем пустые спецификации
                continue
            self._check_old_codes(base, spec_name, google_base_std)
        ##################
        # Вывод финальной статистики
        self._print_final_stats()
        
        if self.errors:
            print(f"\n{Fore.RED}❌ {Style.BRIGHT}НАЙДЕНЫ ОШИБКИ В ПРОВЕРКЕ!")
            print_att_list_table(self.errors, max_len=999, title_row=["Имя спецификации", "Описание ошибки"])
            # raise Exception(self.errors)
        else:
            print(f"\n{Fore.GREEN}✅ {Style.BRIGHT}ПРОВЕРКА МТО НА ТЕГИ ПРОШЛА УСПЕШНО!")
            print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")

    def _calculate_stats(self):
        """Подсчет статистики по спецификациям и строкам"""
        self.stats['total_specs'] = len(self.input_base)
        for spec_name, base in self.input_base.items():
            if base:  # Только непустые спецификации
                self.stats['total_rows'] += len(base)
            else:
                self.stats['empty_specs'] += 1

    def _print_stats(self):
        """Вывод краткой статистики"""
        print(f"\n{Fore.BLUE}📊 {Style.BRIGHT}СТАТИСТИКА:")
        print(f"{Fore.WHITE}   • Всего спецификаций: {Fore.CYAN}{self.stats['total_specs']}")
        print(f"{Fore.WHITE}   • Всего строк: {Fore.CYAN}{self.stats['total_rows']}")
        print(f"{Fore.WHITE}   • Пустых спецификаций: {Fore.YELLOW}{self.stats['empty_specs']}")
        print(f"{Fore.WHITE}   • Строк к проверке: {Fore.GREEN}{self.stats['total_rows']}")
        print(f"{Fore.CYAN}{'-' * 40}")

    def _print_final_stats(self):
        """Вывод финальной статистики после проверки"""
        print(f"\n{Fore.BLUE}📈 {Style.BRIGHT}ИТОГОВАЯ СТАТИСТИКА:")
        print(f"{Fore.WHITE}   • Проверено строк: {Fore.GREEN}{self.stats['checked_rows']}")
        
        if len(self.errors) > 0:
            print(f"{Fore.WHITE}   • Найдено ошибок: {Fore.RED}{len(self.errors)}")
        else:
            print(f"{Fore.WHITE}   • Найдено ошибок: {Fore.GREEN}{len(self.errors)}")
            
        if self.stats['checked_rows'] > 0:
            error_rate = (len(self.errors) / self.stats['checked_rows']) * 100
            if error_rate > 0:
                print(f"{Fore.WHITE}   • Процент ошибок: {Fore.RED}{error_rate:.2f}%")
            else:
                print(f"{Fore.WHITE}   • Процент ошибок: {Fore.GREEN}{error_rate:.2f}%")
        print(f"{Fore.CYAN}{'-' * 40}")

    def _check_tags_value(self, base: list[RowStd], spec_name: str):
        """
        Сравнение количества тегов у позиции с количеством материала
        """
        for row in base:
            self.stats['checked_rows'] += 1
            # Проверяем равно ли количество тегов значению количества позиций
            flag = tag_vs_value_sub_check(row)
            if flag is None:
                continue

            if not flag:
                tags_count = len(row.el[TAGS].value)
                comment = (f"Не совпадает кол-во тегов и кол-во оборудования. "
                           f"TAGS:{tags_count}; "
                           f"VALUES:{row.el[VALUES].value},"
                           f"NUMBERS:{row.el[NUMBERS].value},"
                           f"CODE:{row.el[CODE].value}")
                self.errors.append([spec_name, comment])

    def _check_value_to_float(self, base: list[RowStd], spec_name: str):
        """
        Проверка VALUES на преобразуемость к числу
        """
        for row in base:
            value = row.get_value(VALUES)
            try:
                value = float(value)
            except ValueError:
                comment = (f"Ошибка приведения кол-ва к числу. "                           
                           f"VALUES:{row.el[VALUES].value},"
                           f"NUMBERS:{row.el[NUMBERS].value},"
                           f"CODE:{row.el[CODE].value}")
                self.errors.append([spec_name, comment])
    def _check_old_codes(
        self, base: list[RowStd], spec_name: str, google_base_std: list[RowStd]
    ):
        """
        Проверка кодов на сравнение с гугл базой
        """
        g_base_t_com = google_base_std[0].t_com  # -> TableComments

        color = Color.red
        for row in base:
            """
            Проверка относится ли тип строки к позиции
            """
            row_text = ""
            if not check_position_row(row):
                continue
            if g_base_t_com.tabel_type != t_com_init_cls.GoogleBase.tabel_type:
                continue
            code = str(row.el[CODE].value).strip()
            # Ищем в гугл базе строку с таким же кодом
            google_base = RowStd.get_row_by_code(code, google_base_std)
            if isinstance(google_base, NoneType):  # Если код не найден - присваиваем текст "Не найден в гугл базе..."
                row_text = f"{code} - {PROHIBITION_LIST[1]}"
            else:
                prohibition = google_base.el[PROHIBITION].value  # Если код есть - значение поля К
                # Если в поле запретов и замен есть значение и оно входит в перечень обрабатываемы - обрабатываем
                if prohibition in PROHIBITION_LIST:
                    if not isinstance(google_base, NoneType) and prohibition != PROHIBITION_LIST[1]:
                        if (not isinstance(google_base.el[REPLACEMENT].value, NoneType)
                                and google_base.el[REPLACEMENT].value != ""):
                            # Ищем в столбце Q гугл базы ("Заменен на:")
                            code = str(google_base.el[REPLACEMENT].value).strip()
                            # Находим строку с кодом для замены
                            google_base = RowStd.get_row_by_code(code, google_base_std)
                            if google_base is None:
                                row_text = f"\"{code}\" ERROR check_prohibition: google_base row code - is NONE"
                            else:
                                row_text = f"{row.el[CODE].value}. Замена по ГУГЛ БАЗЕ: {google_base.el[CODE].value}"


            if row_text != "":
                comment = (f"{row_text}. "                           
                           f"VALUES:{row.el[VALUES].value},"
                           f"NUMBERS:{row.el[NUMBERS].value},"
                           )
                self.errors.append([spec_name, comment])

    @staticmethod
    def get_all_mto_from_path(mto_path: str, debug_print_files_list=True):
        # Массив для хранения всех Документов (файлов) текущего проекта
        curr_proj = utils.path.get_files_single(mto_path, endswith=(".xlsx", ".XLSX"))
        if curr_proj == -1:
            print(f"{Fore.RED}❌ {Style.BRIGHT}По пути <{mto_path}> файлы не найдены.")
            exit(0)
        ############################
        if debug_print_files_list:
            print_file_list_to_console(curr_proj)
        ############################
        spec_dict = {}
        for document in curr_proj:
                spec_dict[document.file_name] = document.file_full_path

        mto_dict, _load_audit = RFQ.ds_compare.load_mto.load_mto_by_dict(
            spec_dict, print_mto_to_console=False
        )

        return mto_dict

if __name__ in {"__main__"}:
    mto_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\АН_RFQ\МТО для закупки"
    mto_path = r"\\bcc\eng\PrDoc\377_НИПИГАЗ\АГХК\КСБ\РД\Таблички_графики\14_сравнение с ДС\МТО для закупки"
    mto_dict = MtoCheck.get_all_mto_from_path(mto_path)
    # Проверка МТО
    mto_check_obj = MtoCheck(mto_dict)
    mto_check_obj.check()


