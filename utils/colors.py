class Color:
    red = 'ff6347'
    yellow = 'fbec5d'
    green = '00c179'
    no = "no"
    violet = 'ead1dc'
    gray = "c2c1c2"
    soft_pink = "e866b4"
    soft_cyan = "5edff5"
    dark_lime_green = "06b215"
    soft_yellow = "e5f54e"

    correction = "E9D113"
    correction_db = "c57ece"

    # Шаг 4: статусы сопоставления RFP / MTO / VO
    match_matched = "C6EFCE"        # Тег сопоставлен (зеленый)
    match_not_found = "FFEB9C"      # Не найден в МТО
    match_added_mto = "4FD0E7"      # Добавлен из МТО
    match_added_vo = "B7DEE8"       # Добавлен из VO
    match_replaced_mto = "FCE5CD"   # Тег в МТО заменен (светло красный)
    match_replaced_vo = "D9EAF7"    # Тег в VO заменен
    is_replacement_mto_vo = "EAF7D9"    # Тег в VO заменен (светло зеленый)
    unmatch_rfp_mto = "f3b2d9" # Светло роовый
    

    system_row = soft_yellow
    section_row = soft_cyan
    position_row = no
    empty_row = gray
    other_row = soft_pink
    cabinet_title_row = dark_lime_green
    head_row = "C0C0C0"

    mass = {
        correction_db: 210,
        correction: 200,
        red: 100,
        yellow: 50,
        green: 10,
        # Шаг 4: статусы сопоставления
        match_matched: 5,
        match_not_found: 5,
        match_added_mto: 5,
        match_added_vo: 5,
        match_replaced_mto: 5,
        match_replaced_vo: 5,
        is_replacement_mto_vo: 5,
        unmatch_rfp_mto: 5, 
        #
        no: 0,
        #
        violet: 1,
        gray: 1,
        soft_pink: 1,
        soft_cyan: 1,
        dark_lime_green: 1,
        soft_yellow: 1,
        head_row: 1
    }

    @staticmethod
    def set_el_color(el, color):
        """
        :param el:
        :param color:
        :return:
        """
        result = False
        color_mass_request = Color.mass[color]
        color_mass_current = Color.mass[el.color]
        if color_mass_request >= color_mass_current:
            el.color = color
            result = True

        return result

