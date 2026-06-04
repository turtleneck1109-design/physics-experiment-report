from pathlib import Path
import csv
import math
import zipfile
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "processed_results"
OUT_XLSX = OUT_DIR / "combined_temperature_data_checked.xlsx"


def read_selected_csv(path):
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            t_c = float(row["t / degC"])
            r_ohm = float(row["R / ohm (CH1)"])
            u_v = float(row["U / V (CH2)"])
            rows.append(
                {
                    "t_c": t_c,
                    "r_ohm": r_ohm,
                    "u_v": u_v,
                    "t_k": t_c + 273.15,
                    "inv_t": 1 / (t_c + 273.15),
                    "ln_r": math.log(r_ohm),
                }
            )
    return rows


def linear_fit(xs, ys):
    n = len(xs)
    mx = sum(xs) / n
    my = sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    slope = sxy / sxx
    intercept = my - slope * mx
    fitted = [intercept + slope * x for x in xs]
    residuals = [y - yh for y, yh in zip(ys, fitted)]
    ss_res = sum(e**2 for e in residuals)
    ss_tot = sum((y - my) ** 2 for y in ys)
    r2 = 1 - ss_res / ss_tot if ss_tot else 1
    pearson = math.copysign(math.sqrt(max(r2, 0)), slope)
    return {
        "slope": slope,
        "intercept": intercept,
        "r2": r2,
        "pearson": pearson,
        "ss_res": ss_res,
    }


def col_name(index):
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def cell_ref(row, col):
    return f"{col_name(col)}{row}"


def inline_cell(row, col, text, style=0):
    attrs = f' r="{cell_ref(row, col)}" t="inlineStr"'
    if style:
        attrs += f' s="{style}"'
    return f'<c{attrs}><is><t>{escape(str(text))}</t></is></c>'


def num_cell(row, col, value, style=0, formula=None):
    attrs = f' r="{cell_ref(row, col)}"'
    if style:
        attrs += f' s="{style}"'
    formula_xml = f"<f>{escape(formula)}</f>" if formula else ""
    return f"<c{attrs}>{formula_xml}<v>{value:.12g}</v></c>"


def worksheet_xml(title, rows):
    headers = [
        "序号",
        "t / ℃",
        "R / Ω (CH1)",
        "U / V (CH2)",
        "T / K",
        "1/T / K^-1",
        "ln(R/Ω)",
        "lnR拟合值",
        "lnR残差",
        "U拟合值",
        "U残差",
    ]
    therm = linear_fit([r["inv_t"] for r in rows], [r["ln_r"] for r in rows])
    pn = linear_fit([r["t_k"] for r in rows], [r["u_v"] for r in rows])

    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<cols>",
        '<col min="1" max="1" width="8" customWidth="1"/>',
        '<col min="2" max="11" width="16" customWidth="1"/>',
        '<col min="13" max="16" width="22" customWidth="1"/>',
        "</cols>",
        "<sheetData>",
    ]

    lines.append('<row r="1">')
    for col, header in enumerate(headers, start=1):
        lines.append(inline_cell(1, col, header, 1))
    lines.append("</row>")

    first_data_row = 2
    last_data_row = first_data_row + len(rows) - 1
    for offset, data in enumerate(rows):
        row = first_data_row + offset
        lines.append(f'<row r="{row}">')
        lines.append(num_cell(row, 1, offset + 1))
        lines.append(num_cell(row, 2, data["t_c"]))
        lines.append(num_cell(row, 3, data["r_ohm"]))
        lines.append(num_cell(row, 4, data["u_v"]))
        lines.append(num_cell(row, 5, data["t_k"], formula=f"B{row}+273.15"))
        lines.append(num_cell(row, 6, data["inv_t"], formula=f"1/E{row}"))
        lines.append(num_cell(row, 7, data["ln_r"], formula=f"LN(C{row})"))
        ln_fit = therm["intercept"] + therm["slope"] * data["inv_t"]
        u_fit = pn["intercept"] + pn["slope"] * data["t_k"]
        lines.append(num_cell(row, 8, ln_fit, formula=f"$N$5+$N$6*F{row}"))
        lines.append(num_cell(row, 9, data["ln_r"] - ln_fit, formula=f"G{row}-H{row}"))
        lines.append(num_cell(row, 10, u_fit, formula=f"$N$14+$N$15*E{row}"))
        lines.append(num_cell(row, 11, data["u_v"] - u_fit, formula=f"D{row}-J{row}"))
        lines.append("</row>")

    summary_start = last_data_row + 3
    summary = [
        (summary_start, "数据组", title, None, None),
        (summary_start + 2, "热敏电阻：lnR = a + b*(1/T)", "", None, None),
        (summary_start + 3, "截距 a", therm["intercept"], f"INTERCEPT(G{first_data_row}:G{last_data_row},F{first_data_row}:F{last_data_row})", None),
        (summary_start + 4, "斜率 b", therm["slope"], f"SLOPE(G{first_data_row}:G{last_data_row},F{first_data_row}:F{last_data_row})", None),
        (summary_start + 5, "Pearson r", therm["pearson"], f"CORREL(F{first_data_row}:F{last_data_row},G{first_data_row}:G{last_data_row})", None),
        (summary_start + 6, "R平方", therm["r2"], f"RSQ(G{first_data_row}:G{last_data_row},F{first_data_row}:F{last_data_row})", None),
        (summary_start + 7, "残差平方和", therm["ss_res"], f"SUMSQ(I{first_data_row}:I{last_data_row})", None),
        (summary_start + 11, "PN结：U = a + b*T", "", None, None),
        (summary_start + 12, "截距 a", pn["intercept"], f"INTERCEPT(D{first_data_row}:D{last_data_row},E{first_data_row}:E{last_data_row})", None),
        (summary_start + 13, "斜率 b", pn["slope"], f"SLOPE(D{first_data_row}:D{last_data_row},E{first_data_row}:E{last_data_row})", None),
        (summary_start + 14, "Pearson r", pn["pearson"], f"CORREL(E{first_data_row}:E{last_data_row},D{first_data_row}:D{last_data_row})", None),
        (summary_start + 15, "R平方", pn["r2"], f"RSQ(D{first_data_row}:D{last_data_row},E{first_data_row}:E{last_data_row})", None),
        (summary_start + 16, "残差平方和", pn["ss_res"], f"SUMSQ(K{first_data_row}:K{last_data_row})", None),
        (summary_start + 19, "说明", "E-G列、H-K列和右侧汇总区均含Excel公式，并写入了Python复算值。", None, None),
    ]
    for row_num, label, value, formula, _ in summary:
        lines.append(f'<row r="{row_num}">')
        lines.append(inline_cell(row_num, 1, label, 1 if formula is None and value == "" else 0))
        if isinstance(value, str):
            lines.append(inline_cell(row_num, 2, value))
        else:
            lines.append(num_cell(row_num, 2, value, formula=formula))
        lines.append("</row>")

    lines.extend(
        [
            "</sheetData>",
            f'<autoFilter ref="A1:K{last_data_row}"/>',
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
            "</worksheet>",
        ]
    )
    return "\n".join(lines)


def write_xlsx(sheet_data):
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="room_to_30C" sheetId="1" r:id="rId1"/>
<sheet name="30C_to_90C" sheetId="2" r:id="rId2"/>
<sheet name="combined_25C_to_90C" sheetId="3" r:id="rId3"/>
</sheets>
<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/>
</workbook>"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="1"><fill><patternFill patternType="none"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""

    with zipfile.ZipFile(OUT_XLSX, "w", compression=zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/styles.xml", styles)
        z.writestr("xl/worksheets/sheet1.xml", worksheet_xml("room_to_30C", sheet_data["room_to_30C"]))
        z.writestr("xl/worksheets/sheet2.xml", worksheet_xml("30C_to_90C", sheet_data["30C_to_90C"]))
        z.writestr("xl/worksheets/sheet3.xml", worksheet_xml("combined_25C_to_90C", sheet_data["combined_25C_to_90C"]))


def main():
    OUT_DIR.mkdir(exist_ok=True)
    sheet_data = {
        "room_to_30C": read_selected_csv(OUT_DIR / "room_to_30C_selected_temperature_data.csv"),
        "30C_to_90C": read_selected_csv(OUT_DIR / "30C_to_90C_selected_temperature_data.csv"),
        "combined_25C_to_90C": read_selected_csv(OUT_DIR / "combined_25C_to_90C_selected_temperature_data.csv"),
    }
    write_xlsx(sheet_data)
    print(f"wrote: {OUT_XLSX}")


if __name__ == "__main__":
    main()
