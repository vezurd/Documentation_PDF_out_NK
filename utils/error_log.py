from utils.save_to_file import save_text_to_file
import logging


class ErrorLog:
    error_log = []
    work_log = []
    exit_path = ""
    logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s",
                        style="%",
                        datefmt="%Y-%m-%d %H:%M",
                        level=logging.WARNING)

    @staticmethod
    def clear():
        ErrorLog.error_log = []
        ErrorLog.work_log = []
        ErrorLog.exit_path = ""

    @staticmethod
    def start(exit_path=""):
        ErrorLog.error_log = []
        ErrorLog.work_log = []
        ErrorLog.exit_path = exit_path

    @staticmethod
    def add_error(text, print_flag=1):
        if print_flag:
            print(text)
        ErrorLog.error_log.append(text)

    @staticmethod
    def add_log(text, print_flag=1):
        if print_flag:
            print(text)
        ErrorLog.work_log.append(text)

    @staticmethod
    def print_err_log(out_dir="", open_file_flag=1):
        text = "Error_log\n"
        print("Error_log")
        if ErrorLog.error_log:
            for x in ErrorLog.error_log:
                print(x)
                text += str(x) + "\n"
            if out_dir != "":
                # print(out_dir)
                import datetime
                now = datetime.datetime.now()
                date = now.strftime("%Y.%m.%d_%HH-%MM")
                save_text_to_file(text, f"/ErrorLog_{date}.txt", out_dir, open_file_flag=open_file_flag)
                ErrorLog.exit_path = out_dir
        else:
            print("   Список ошибок пуст")

    @staticmethod
    def exit(open_file_flag=1):
        if ErrorLog.exit_path != "":
            err_text = (f"ErrorLog: exit\n"
                        f"    Файл ошибок сохранен по пути\n"
                        f"    {ErrorLog.exit_path}")
            ErrorLog.add_error(err_text)
            ErrorLog.print_err_log(ErrorLog.exit_path, open_file_flag=open_file_flag)
        exit(0)
