import utils.file_name_converts as fn
import pdf_parsing_v2_engine.stamp_text.stamp_field_cleaners as sfc

text_list = ["AGCC.287-2879-SOT.OD-0001_01-AN02_RU.doc", 
"AGCC.287-4130-KSB1.WIR-0008_V_RU.dwg", 
"AGCC.287-2869-SOS.BOE-0001_01-AN02_RU.xlsx",
"AGCC.287-2869-SOS.BOE-0001_S_RU.xlsx"]

ctx = None
for text in text_list:
    outcome = sfc.pipeline_file_name_stamp(text, ctx, dbg=False)
    print(outcome.value)
    # fn.print_project_file_name_guide(text)
    # fn.print_parse_parts("text", fn.ProjectFileName.parse_strict(text))