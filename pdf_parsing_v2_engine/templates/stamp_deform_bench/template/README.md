Place the default template JSON here if you want to run the CLI without
passing `--template`.

Expected default file name:

```text
pdf_parsing_v2_engine/templates/stamp_deform_bench/template/stamp_template.json
```

You can also keep using project templates elsewhere and pass an explicit path:

```bash
python pdf_parsing_v2_engine/tools/generate_deformed_stamp_pdfs.py ^
  --template "pdf_parsing_v2_engine/templates/agcc_287/dwg_page1.json"
```
