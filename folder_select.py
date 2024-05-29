import tkinter as tk
from tkinter import filedialog as fd

def getDirectory():
    root = tk.Tk()  # пустое родительское окно
    root.withdraw()  # прячем его
    result = fd.askdirectory(
        master=root,   # диалогу нужно родительское окно, путь даже невидимое.
        mustexist=True)  # только существующие каталоги
    root.destroy()  # уничтожаем родительское окно
    #print(type(result), repr(result))  # result будет содержать путь или пустую строку при отмене
    return result

if __name__ == "__main__":
    pdf_path = getDirectory()
    print(pdf_path)