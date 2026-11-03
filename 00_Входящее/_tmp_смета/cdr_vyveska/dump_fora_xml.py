# -*- coding: utf-8 -*-
import zipfile, re, os, sys
from xml.etree import ElementTree as ET

src = r"C:\Users\x6668\Desktop\ИИ\ARTEL\03_КП_и_сметы\Расчёты\Фора_шаблоны\Световые буквы без основы Артель.xlsx"
calc = r"C:\Users\x6668\Desktop\ИИ\ARTEL\00_Входящее\_tmp_смета\cdr_vyveska\calc_CakeHome_Artel.xlsx"
path = calc if os.path.exists(calc) else src
print("FILE", path)

NS = {"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r":"http://schemas.openxmlformats.org/officeDocument/2006/relationships"}

with zipfile.ZipFile(path) as z:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {}
    for rel in rels:
        rid_to_target[rel.attrib.get("Id")] = rel.attrib.get("Target")
    sheets = []
    for sh in wb.findall("m:sheets/m:sheet", NS):
        name = sh.attrib.get("name")
        rid = sh.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rid_to_target.get(rid)
        sheets.append((name, target))
        print("SHEET", name, "->", target)

    # shared strings
    sst = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in root.findall("m:si", NS):
            texts = [t.text or "" for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")]
            sst.append("".join(texts))

    def colrow(cell_ref):
        m = re.match(r"([A-Z]+)(\d+)", cell_ref)
        return m.group(1), int(m.group(2))

    def dump_sheet(name_substr, max_rows=120):
        for name, target in sheets:
            if name_substr not in name:
                continue
            xml = z.read("xl/" + target)
            root = ET.fromstring(xml)
            print("\n====", name, "====")
            rows = {}
            for c in root.findall(".//m:sheetData/m:row/m:c", NS):
                ref = c.attrib.get("r")
                if not ref: continue
                col, row = colrow(ref)
                t = c.attrib.get("t")
                f = c.find("m:f", NS)
                v = c.find("m:v", NS)
                val = None
                if f is not None:
                    val = "F:" + (f.text or "")
                    if v is not None and v.text is not None:
                        val += " = " + v.text
                elif v is not None:
                    if t == "s":
                        val = sst[int(v.text)] if v.text and int(v.text) < len(sst) else v.text
                    else:
                        val = v.text
                if val is not None:
                    rows.setdefault(row, {})[col] = val
            for r in sorted(rows)[:max_rows]:
                cols = rows[r]
                # print interesting cols A-H typically
                parts = []
                for col in sorted(cols, key=lambda x: (len(x), x)):
                    if col > "H" and col not in ("I","J","K","L"):
                        continue
                    parts.append(f"{col}:{cols[col]}")
                if parts:
                    print(f"R{r}\t" + " | ".join(parts))
            return
        print("NOT FOUND", name_substr)

    dump_sheet("Выходные", 40)
    dump_sheet("Параметры", 30)
    dump_sheet("Тех", 40)
    dump_sheet("Расходные", 30)
    dump_sheet("Дополнитель", 40)
    dump_sheet("Работа станков", 30)
    dump_sheet("Работа оператора", 30)
    dump_sheet("Налоги", 30)
    dump_sheet("Материалы", 100)