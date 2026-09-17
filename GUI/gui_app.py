import TKinterModernThemes as TKMT
from functools import partial
import tkinter as tk
import json

import main


class App(TKMT.ThemedTKinterFrame):
    def __init__(self, theme, mode, usecommandlineargs=True, usethemeconfigfile=True):
        super().__init__("GUI АГХК проверка РД (2025.03.07)", theme, mode,
                         usecommandlineargs=usecommandlineargs, useconfigfile=usethemeconfigfile)
        # # self.Notebook(self.root)
        # self.checkbox1 = tk.BooleanVar()
        # self.checkbox2 = tk.BooleanVar(value=True)
        #
        # self.radiobuttonvar = tk.StringVar(value='button2')
        # self.togglebuttonvar = tk.BooleanVar()
        # #
        # self.textinputvar = tk.StringVar(value="Type text here.")
        # self.spinboxnumvar = tk.IntVar(value=25)
        # self.spinboxcolorvar = tk.StringVar(value="blue")
        # self.comboboxvar = tk.StringVar()
        # #
        # self.option_menu_list = ["a", "b", "c", "d"]
        # self.optionmenuvar = tk.StringVar(value=self.option_menu_list[0])
        #
        # self.slidervar = tk.IntVar(value=25)
        # #
        # self.check_frame = self.addLabelFrame("CheckButtons")
        # self.check_frame.Checkbutton("Unchecked", self.checkbox1, self.printcheckboxvars, (1,))
        # self.check_frame.Checkbutton("Unchecked", self.checkbox2, self.printcheckboxvars, (1,))
        # self.check_frame.Checkbutton("Disabled Unchecked", self.checkbox1, disabled=True)
        # self.check_frame.Checkbutton("Disabled Checked", self.checkbox2, disabled=True)
        # self.check_frame.SlideSwitch("Slide Switch", None)
        #
        # # # Separator
        # # self.Seperator()
        #
        # self.nextCol()
        # self.button_frame = self.addLabelFrame("Проверка РД")
        # self.button_frame.AccentButton("Открыть папку с PDF", self.handleButtonClick)
        # self.button_frame.AccentButton("Открыть папку с МТО", self.handleButtonClick)
        #
        #
        # self.button_frame = self.addLabelFrame("Сравнение МТО")
        # self.button_frame.Button("MTO vs MTO", self.handleButtonClick)
        # # self.button_frame.Button("MTO vs RFQ", self.handleButtonClick)
        # # self.button_frame.Button("Открыть файл MTO", self.handleButtonClick)
        # # self.button_frame.Button("Открыть файл RFQ", self.handleButtonClick)
        #
        # self.nextCol()
        self.panedWindow = self.PanedWindow("Paned Window Test", rowspan=3)
        # self.pane1 = self.panedWindow.addWindow()

        # Define treeview data

        self.pane1 = self.panedWindow.addWindow()
        self.notebook = self.pane1.Notebook("Test Notebook")
        self.tab_1 = self.notebook.addTab("Проверка PDF")
        self.button_frame1 = self.tab_1.addLabelFrame("Проверка папки с PDF комплекта")
        self.button_frame1.AccentButton("Открыть папку с PDF", self.handleButtonClick)
        # self.tab_1.button_frame.AccentButton("Открыть папку с PDF", self.handleButtonClick)
        # self.tab_1.button_frame.AccentButton("Открыть папку с МТО", self.handleButtonClick)

        self.tab_2 = self.notebook.addTab("Проверка MTO, RFQ, CO (.xlsx)")
        self.tab_2.SlideSwitch("Синхронизировать с ГуглБазой", None)
        self.button_frame_2_1 = self.tab_2.addLabelFrame("Проверить Excel MTO")
        self.button_frame_2_1.AccentButton("Открыть папку DWG с MTO", self.handleButtonClick)
        self.button_frame_2_1.Button("Открыть файл MTO", self.handleButtonClick)
        self.button_frame_2_1.Button("Открыть файл RFQ", self.handleButtonClick)

        self.button_frame_2_2 = self.tab_2.addLabelFrame("Проверить вывод NanoCAD (CO.xlsx)")
        self.button_frame_2_2.Button("Открыть файл CO.xlsx", self.handleButtonClick)
        self.button_frame_2_2.Button("Открыть файл MTO", self.handleButtonClick)

        self.tab_3 = self.notebook.addTab("Кабельный журнал")
        self.button_frame_3_1 = self.tab_3.addLabelFrame("Открыть папку с ТПК (при наличии со старым CJ.docx)")
        self.button_frame_3_1.AccentButton("Открыть папку с ТПК", self.handleButtonClick)
        self.tab_3.Label("Label text here.")

        self.tab_4 = self.notebook.addTab("Вспомогательные функции")
        self.button_frame_4_1 = self.tab_4.addLabelFrame("Сравнение двух МТО")
        self.button_frame_4_1.Button("MTO vs MTO", self.handleButtonClick)
        # self.tab_3 = self.notebook.addTab("Tab 3")
        # self.tab_3.Text("Normal text here.")

        # self.pane3 = self.panedWindow.addWindow()
        # self.pane3.Label(self.theme.capitalize() + " theme: " + self.mode)

        self.debugPrint()
        self.run()

    def printcheckboxvars(self, number):
        print("Checkbox number:", number, "was pressed")
        print("Checkboxes: ", self.checkbox1.get(), self.checkbox2.get())

    def printradiobuttons(self, _var, _indx, _mode):
        print("Radio button: ", self.radiobuttonvar.get(), "pressed.")

    def handleButtonClick(self):

        print("Button clicked. Current toggle button state: ", self.togglebuttonvar.get())

    def textupdate(self, _var, _indx, _mode):
        print("Current text status:", self.textinputvar.get())

    def menuprint(self, item):
        if self == self:
            pass
        print("Menu item chosen: ", item)

    def validateText(self, text):
        if self == self:
            pass
        if 'q' not in text:
            return True
        print("The letter q is not allowed.")
        return False


if __name__ == '__main__':
    app = App("sun-valley", "light")
