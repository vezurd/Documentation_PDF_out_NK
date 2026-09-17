import math


def integer_from_any(s):
    if "=" in s:
        s = s.replace("=", "")
        s = eval(s)
    return math.ceil(float(str(s).replace(",", ".")))