#       К чему относится обор;  столбец
from __future__ import annotations

from tags.equipment_codes import *

d2_kword = {
    0: ["SOT", 1],
    1: ["SKUD", 1],
    2: ["POS", 1],
    3: ["SOS", 1],
    4: ["SPP+SOO", 1],
    5: ["Шкаф SOT (SK,SKT)", 2],
    6: ["Шкаф SOT (SX)", 2],
    7: ["Шкаф SKUD (SK,SKT)", 2],
    8: ["Шкаф POS (SK)", 2],
    9: ["Шкаф SOS (SK)", 2],
}
d_cj_kword = {
    0: ["SOT", 1],
    1: ["SKUD", 1],
    2: ["POS", 1],
    3: ["SOS", 1],
    4: ["SPP;SOO", 1],
    5: ["SOT", 2],
    6: ["SOT", 2],
    7: ["SKUD", 2],
    8: ["POS", 2],
    9: ["SOS", 2],
}
foreign_tag_row = 3  # Индекс столбца для тегов не относящихся к системе КСБ
hide_tags_list = ["XT"]

_UNCATEGORIZED_RU = "Не классифицировано"
_ROW_CATEGORY_NAME: dict[int, str] = {
    0: "КСБ (НЕ в шкафах)",
    1: "КСБ (оборудование в шкафах)",
    2: "КСБ (шкафы SOT SX)",
    3: "Не КСБ / иностранные",
}


def hide_tag(equipment) -> bool:
    """True if equipment is on ``hide_tags_list`` (XT).

    Used as ``hidden_by_rule`` on extra-in-sheets rows. Tag analysis still
    lists these tags in payload, GUI, and the txt report.
    """
    if equipment is None:
        return False
    eq = str(equipment).strip()
    if eq == "":
        return False
    return eq in hide_tags_list


def _coerce_tag(tag_or_tc: str | TagClass, *, strict: bool = True) -> TagClass:
    if isinstance(tag_or_tc, TagClass):
        return tag_or_tc
    return TagClass(tag_or_tc, strict=strict)


def check_system_code_94s_4_1(tag_or_tc: str | TagClass, *, strict: bool = True) -> bool:
    tag = _coerce_tag(tag_or_tc, strict=strict)
    if not tag.is_valid:
        return False
    return tag.system_code in tags_94s_4_1_sys_codes


def check_tag_94s_4_2(tag_or_tc: str | TagClass, *, strict: bool = True) -> bool:
    tag = _coerce_tag(tag_or_tc, strict=strict)
    if not tag.is_valid:
        return False
    return tag.equipment in tags_94s_4_2_eq


def check_tag_94s_4_4_additional_code(tag_or_tc: str | TagClass, *, strict: bool = True) -> bool:
    tag = _coerce_tag(tag_or_tc, strict=strict)
    if not tag.is_valid:
        return False
    if tag.additional_code is not None:
        return tag.additional_code in tags_94s_4_4_ac
    return True


class TagClass:
    """Parses a cable / KSB tag string.

    Use ``strict=True`` (default) for domain logic (Excel, CJ, RFQ): malformed
    input raises immediately. Use ``strict=False`` for bulk PDF tag scans and
    tag analysis payloads: sets ``is_valid`` / ``parse_error`` instead of raising.
    """

    def __init__(self, tag, *, strict: bool = True) -> None:
        self.tag = str(tag) if tag is not None else ""
        self.is_valid = True
        self.parse_error = ""
        try:
            self._parse_tag_core(self.tag)
            self.is_valid = True
            self.parse_error = ""
        except Exception as e:
            if strict:
                raise
            self.is_valid = False
            self.parse_error = str(e)
            self._apply_parse_failure_defaults()

    def _parse_tag_core(self, tag: str) -> None:
        def tag_get_sub_system(d):
            try:
                d = int(d)
                return d2_kword[d][0]
            except (ValueError, KeyError, IndexError) as e:
                print(f"Ошибка в tag_get_sub_system: {e}, входные данные: {d}")
                return "Unknown"

        def tag_get_sub_system_cj(d):
            try:
                d = int(d)
                return d_cj_kword[d][0]
            except (ValueError, KeyError, IndexError) as e:
                print(f"Ошибка в tag_get_sub_system_cj: {e}, входные данные: {d}")
                return "Unknown"

        def tag_get_out_row(sys_code, sub_sys_num):
            try:
                if sys_code == "S":
                    sub_sys_num = int(sub_sys_num)
                    return d2_kword[sub_sys_num][1]
                return foreign_tag_row
            except (ValueError, KeyError, IndexError) as e:
                print(f"Ошибка в tag_get_out_row: {e}, sys_code: {sys_code}, sub_sys_num: {sub_sys_num}")
                return "Unknown"

        t_split = tag.split("-")

        if len(t_split) < 3:
            raise ValueError(
                f"Тег должен содержать минимум 3 части, получено: {len(t_split)}. Тег - {tag}"
            )

        self.wbs = t_split[0]

        if len(t_split) == 6:
            self.building_symbol = t_split[1]
            self.building_number = t_split[2]
        else:
            self.building_symbol = None
            self.building_number = None

        if len(t_split) >= 3:
            self.system_code = t_split[-3]
        else:
            self.system_code = "Unknown"
            print(f"Предупреждение: недостаточно элементов для system_code в теге: {tag}")

        if len(t_split) >= 2:
            self.equipment = t_split[-2]
        else:
            self.equipment = "Unknown"
            print(f"Предупреждение: недостаточно элементов для equipment в теге: {tag}")

        suffix_split = t_split[-1].split("(")
        if len(suffix_split) > 1:
            self.sequence_number = suffix_split[0]
            self.additional_code = suffix_split[1].strip().strip(")")
        else:
            self.sequence_number = t_split[-1]
            self.additional_code = None

        if self.sequence_number and len(self.sequence_number) >= 2:
            self.floor = self.sequence_number[0]
            self.sum_system_number = self.sequence_number[1]
            self.sub_system = tag_get_sub_system(self.sum_system_number)
            self.sub_system_cj = tag_get_sub_system_cj(self.sum_system_number)
        else:
            self.floor = "Unknown"
            self.sum_system_number = "Unknown"
            self.sub_system = "Unknown"
            self.sub_system_cj = "Unknown"
            print(f"Предупреждение: sequence_number '{self.sequence_number}' слишком короткий")

        self.out_row = tag_get_out_row(self.system_code, self.sum_system_number)

    def _apply_parse_failure_defaults(self) -> None:
        self.wbs = ""
        self.system_code = ""
        self.equipment = ""
        self.building_symbol = None
        self.building_number = None
        self.sequence_number = ""
        self.additional_code = None
        self.floor = ""
        self.sum_system_number = ""
        self.sub_system = "Unknown"
        self.sub_system_cj = "Unknown"
        self.out_row = foreign_tag_row

    @property
    def category_name(self) -> str:
        if not self.is_valid:
            return _UNCATEGORIZED_RU
        orow = self.out_row
        if isinstance(orow, int) and 0 <= orow <= 3:
            return _ROW_CATEGORY_NAME[orow]
        return _UNCATEGORIZED_RU

    def is_hidden(self) -> bool:
        return hide_tag(self.equipment)

    def get_shield_number(self):
        if not self.is_valid:
            return None
        try:
            n = int(self.sum_system_number)
        except (TypeError, ValueError):
            return None
        try:
            if n in (5, 7, 8, 9):
                return self.sequence_number[:3]
            if n == 6:
                return self.sequence_number[:2]
        except (TypeError, ValueError, IndexError):
            return None
        return None


class GoogleSystemCode:
    def __init__(self, tag):
        self.tag = tag
        tag_split = tag.split("(")  # AEX(IS)
        if len(tag_split) > 1:
            self.equipment = tag_split[0]  # AEX
            self.additional_code = tag_split[1].strip(")")  # IS
        else:
            self.equipment = tag_split[-1]  # AEX
            self.additional_code = None


if __name__ == "__main__":
    t = ["2869-S-BGB-1305", "2612-PB-01-S-FV-1511", "2612-PB-01-S-BZ-1913"]
    for i in t:
        tag = TagClass(i)
        print(tag.get_shield_number())
