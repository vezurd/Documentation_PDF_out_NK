from utils.colors import Color


class ElementCT:
    def __init__(self, value, color=Color.no):
        self.value = value  # Основное значение
        self.color = color  # Цвет для вывода в результат
        self.comment = ""

    def concatenate_str_values(self, str_value, separator=", "):
        if self.value == "":
            self.value = str_value
        else:
            if str_value == "":
                pass
            else:
                self.value += separator + str_value
