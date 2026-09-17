"""
CableLabelCls
Принимает на вход строку с коротким тегом кабеля.
Разбивает на 3 части:
    cable_code      : I, C, P и т.п.
    additional_code : EX или Is
    cable_number    : децимальный номер
"""


class CableLabelCls:
    def __init__(self, cable_label):
        self.cable_label = cable_label  # I.Ex-0105 ; E-0101 ; C.IS-0304

        t_split = self.cable_label.split("-")
        if len(t_split) == 2:
            self.cable_number = t_split[1]  # I.Ex; 0105 - берем 1050
        else:
            print(f"ERROR CableLabelCls: {cable_label} не корректный")
            exit(0)
        self.cable_code = None
        self.additional_code = ""
        left_split = t_split[0].split(".")
        if len(left_split) == 1:
            self.cable_code = left_split[0]
        elif len(left_split) == 2:
            self.cable_code = left_split[0]
            self.additional_code = left_split[1]
        else:
            print(f"ERROR CableLabelCls: {cable_label} не корректный")
            exit(0)
