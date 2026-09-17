def get_Real_Page_Format(width, height):
    tolerance = 15
    if width > height:
        width, height = height, width
    formats = {
        "A4":   [210, 297],
        "A4x3": [297, 630],
        "A4x4": [297, 841],
        "A4x5": [297, 1051],
        "A4x6": [297, 1261],
        "A4x7": [297, 1471],
        "A4x8": [297, 1682],
        "A4x9": [297, 1987],
        "A3":   [297, 420],
        "A3x3": [420, 891],
        "A3x4": [420, 1189],
        "A3x5": [420, 1486],
        "A3x6": [420, 1783],
        "A3x7": [420, 2080],
        "A2":   [420, 594],
        "A2x3": [594, 1261],
        "A2x4": [594, 1682],
        "A2x5": [594, 2102],
        "A1":   [594, 841],
        "A1x3": [841, 1783],
        "A1x4": [841, 2378],
        "A0":   [841, 1189],
        "A0x2": [1189, 1682],
        "A0x3": [1189, 2523],
    }
    def check_size (size, f_size):
        min_size = size-tolerance
        max_size = size+tolerance
        bool_flag = False
        if f_size >= min_size and f_size<=max_size:
            bool_flag = True
        return bool_flag

    r_format = "неизвестный формат страницы"
    for key in formats:
        if check_size(width, formats[key][0]) and check_size(height, formats[key][1]):
            r_format = key
            break
        elif check_size(width, 309) and check_size(height, 437):
            r_format = "A3"
            break
        elif check_size(width, 340) and check_size(height, 477):
            r_format = "A3"
            break
    return r_format

def get_a4_count_in_format (curr_format: str):
    curr_format = curr_format.replace("х","x")
    curr_format = curr_format.replace("×", "x")
    formats = {
        "A4": 1,
        "A4x3": 3,
        "A4x4": 4,
        "A4x5": 5,
        "A4x6": 6,
        "A4x7": 7,
        "A4x8": 9,
        "A4x9": 9,
        "A3": 2,
        "A3x3": 2*3,
        "A3x4": 2*4,
        "A3x5": 2*5,
        "A3x6": 2*6,
        "A3x7": 2*7,
        "A2": 4*1,
        "A2x3": 4*3,
        "A2x4": 4*4,
        "A2x5": 4*5,
        "A1": 8,
        "A1x3": 8*3,
        "A1x4": 8*4,
        "A1x5": 8 * 5,
        "A1x6": 8 * 6,
        "A0": 16,
        "A0x2": 16*2,
        "A0x3": 16*3,
        "A0x4": 16 * 4,
        "A0x5": 16 * 5,
        "A0x6": 16 * 6,
    }
    return formats[curr_format]
if __name__ == "__main__":
    print(get_Real_Page_Format(437,309))