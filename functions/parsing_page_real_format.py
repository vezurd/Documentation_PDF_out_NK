def get_Real_Page_Format(width, height):
    tolerance = 10
    if width > height:
        width, height = height, width
    fomats = {
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


    for key in fomats:
        if check_size(width,fomats[key][0]) and check_size(height,fomats[key][1]):
            r_format = key
            break
        else:
            r_format = "неизвестный формат страницы"

    return r_format

if __name__ == "__main__":
    print(get_Real_Page_Format(420,2080))
    print(get_Real_Page_Format(2080, 420))