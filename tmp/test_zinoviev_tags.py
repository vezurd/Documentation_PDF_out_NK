# -*- coding: utf-8 -*-
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")

from RFQ.ds_compare.tsd_zinoviev_compare import (
    _format_tags,
    _format_tags_diff,
    compare_robot_vs_zinoviev,
    parse_tags,
)

messy = (
    "2245-SH-01-S-PG-1103(Ex); ['2245-SH-01-S-PG-1101(Ex)']; "
    "['2245-SH-01-S-PG-1103(Ex)']"
)
got = parse_tags(messy)
print("messy", sorted(got))
assert got == {"2245-SH-01-S-PG-1103(Ex)", "2245-SH-01-S-PG-1101(Ex)"}
assert parse_tags(["a", "b", "a"]) == {"a", "b"}
assert parse_tags("['x']") == {"x"}
print("format", _format_tags(got))
print("diff", _format_tags_diff({"a", "b"}, {"b", "c"}))
print("OK parse")

r = compare_robot_vs_zinoviev()
print(r.message)
s = r.stats
assert s is not None
print(
    "covered",
    s.zin_covered,
    "/",
    s.zin_keys,
    "tags_differ",
    s.tags_differ,
    "tags_missing",
    s.zin_tags_missing,
)
