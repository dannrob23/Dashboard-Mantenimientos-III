# -*- coding: utf-8 -*-
import openpyxl
wb = openpyxl.load_workbook(r"Dashboard_Mantenimientos_Preventivos.xlsx")
wsd = wb["Dashboard_KPIs"]
for i, ch in enumerate(wsd._charts):
    cats = ch.series[0].cat.numRef.f if ch.series[0].cat and ch.series[0].cat.numRef else None
    print("Chart", i, "| val:", ch.series[0].val.numRef.f, "| cats:", cats)
wsb = wb["Bitacora_Diaria"]
for rng, rules in wsb.conditional_formatting._cf_rules.items():
    print("CF range", rng.sqref, "->", len(rules), "rules")
print("Defined names:", [n.name for n in wb.defined_names._defined_name.values()])
wsi = wb["Instructivo_Tecnico"]
print("Instructivo max_row:", wsi.max_row)
print("OK VERIFY")
