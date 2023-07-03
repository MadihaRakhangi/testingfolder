import pandas as pd
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt
import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
import numpy as np
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from tabulate import tabulate
from docx.shared import RGBColor
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

sf = pd.read_csv("eli-socket.csv")
lf = pd.read_csv("sugg-max-eli.csv")


sf1 = sf[
    [
        "SN",
        "Device Name",
        "Location",
        "Facility Area",
        "Earthing Configuration",
        "Type of Circuit Location",
        "Device Rating (A)",
        "Device Make",
        "Device Type",
        "Device Sensitivity (mA)",
        "No. of Phases",
        "Trip Curve",
    ]
]

df2 = sf[
    [
        "SN",
        "Device Name",
        "Device Rating (A)",
        "Device Type",
        "No. of Phases",
        "V_LN",
        "V_LE",
        "V_NE",
        "L1-ELI",
        "L2-ELI",
        "L3-ELI",
        "Psc (kA)",
    ]
]

sf_filled = sf.fillna("")
sf["Device Rating (A)"] = sf["Device Rating (A)"].astype(int)

Device_Rating = sf["Device Rating (A)"]
No_phase = sf["No. of Phases"]
T_Curve = sf["Trip Curve"]
new_column = []
result_column = []
P = 0
K = 0
TMS = 1
TDS = 1


def IEC(df1, df2):
    I = Is * ((((K * TMS) / Td) + 1) ** (1 / P))
    if row["Earthing Configuration"] == "TN":
        IEC_val_TN = row["V_LE"] / I
        new_column.append(round(IEC_val_TN, 4))
        if row["No. of Phases"] == 3:
            if (
                row["L1-ELI"] <= IEC_val_TN
                and row["L2-ELI"] <= IEC_val_TN
                and row["L3-ELI"] <= IEC_val_TN
            ):
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        elif row["No. of Phases"] == 1:
            if row["L1-ELI"] <= IEC_val_TN:
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        else:
            result_column.append("N/A")
    elif row["Earthing Configuration"] == "TT":
        IEC_val_TT = 50 / I
        new_column.append(round(IEC_val_TT, 4))
        if row["No. of Phases"] == 3:
            if (
                row["L1-ELI"] <= IEC_val_TT
                and row["L2-ELI"] <= IEC_val_TT
                and row["L3-ELI"] <= IEC_val_TT
            ):
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        elif row["No. of Phases"] == 1:
            if row["L1-ELI"] <= IEC_val_TT:
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        else:
            result_column.append("N/A")


def IEEE(df1, df2):
    I = Is * (((((A / ((Td / TDS) - B)) + 1)) ** (1 / p)))
    if row["Earthing Configuration"] == "TN":
        IEEE_val_TN = row["V_LE"] / I
        new_column.append(round(IEEE_val_TN, 4))
        if row["No. of Phases"] == 3:
            if (
                row["L1-ELI"] <= IEEE_val_TN
                and row["L2-ELI"] <= IEEE_val_TN
                and row["L3-ELI"] <= IEEE_val_TN
            ):
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        elif row["No. of Phases"] == 1:
            if row["L1-ELI"] <= IEEE_val_TN:
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        else:
            result_column.append("N/A")
    elif row["Earthing Configuration"] == "TT":
        IEEE_val_TT = 50 / I
        new_column.append(round(IEEE_val_TT, 4))
        if row["No. of Phases"] == 3:
            if (
                row["L1-ELI"] <= IEEE_val_TT
                and row["L2-ELI"] <= IEEE_val_TT
                and row["L3-ELI"] <= IEEE_val_TT
            ):
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        elif row["No. of Phases"] == 1:
            if row["L1-ELI"] <= IEEE_val_TT:
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        else:
            result_column.append("N/A")


for index, row in sf.iterrows():
    if row["Device Type"] == "MCB":
        rating = row[6]
        trip = row[11]  # Assuming the "Trip Curve" column is at index 10
        result_row = lf[lf["Device Rating (A)"] == rating]
        if trip in result_row.columns:
            val_MCB = result_row[trip].values[0]
        else:
            val_MCB = (
                0  # Set a default value when 'Trip Curve' value is not found in sugg-max-eli.csv
            )
        new_column.append(round(val_MCB, 2))

        if row["No. of Phases"] == 3:
            if row["L1-ELI"] <= val_MCB and row["L2-ELI"] <= val_MCB and row["L3-ELI"] <= val_MCB:
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        elif row["No. of Phases"] == 1:
            if row["L1-ELI"] <= val_MCB:
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        else:
            result_column.append("N/A")

    elif row["Device Type"] in ["RCD", "RCBO", "RCCB"] and row["Earthing Configuration"] == "TN":
        rccb_val_TN = (row["V_LE"] / row["Device Sensitivity (mA)"]) * 1000
        new_column.append(round(rccb_val_TN, 4))
        if row["No. of Phases"] == 3:
            if (
                row["L1-ELI"] <= rccb_val_TN
                and row["L2-ELI"] <= rccb_val_TN
                and row["L3-ELI"] <= rccb_val_TN
            ):
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        elif row["No. of Phases"] == 1:
            if row["L1-ELI"] <= rccb_val_TN:
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        else:
            result_column.append("N/A")

    elif row["Device Type"] in ["RCD", "RCBO", "RCCB"] and row["Earthing Configuration"] == "TT":
        rccb_val_TT = (50 / row["Device Sensitivity (mA)"]) * 1000
        new_column.append(round(rccb_val_TT, 4))
        if row["No. of Phases"] == 3:
            if (
                row["L1-ELI"] <= rccb_val_TT
                and row["L2-ELI"] <= rccb_val_TT
                and row["L3-ELI"] <= rccb_val_TT
            ):
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        elif row["No. of Phases"] == 1:
            if row["L1-ELI"] <= rccb_val_TT:
                result_column.append("Pass")
            else:
                result_column.append("Fail")
        else:
            result_column.append("N/A")

    elif row["Device Type"] == "MCCB" or row["Device Type"] == "ACB":
        if row["Type of Circuit Location"] == "Final":
            Td = 0.4
        elif row["Type of Circuit Location"] == "Distribution":
            Td = 5
        Is = row["Device Rating (A)"]
        if row[11] == "IEC Standard Inverse":
            P = 0.02
            K = 0.14
            IEC(sf1, df2)
        elif row[11] == "IEC Very Inverse":
            P = 1
            K = 13.5
            IEC(sf1, df2)
        elif row[11] == "IEC Long-Time Inverse":
            P = 1
            K = 120
            IEC(sf1, df2)
        elif row[11] == "IEC Extremely Inverse":
            P = 2
            K = 80
            IEC(sf1, df2)
        elif row[11] == "IEC Ultra Inverse":
            P = 2.5
            K = 315.2
            IEC(sf1, df2)
        elif row[11] == "IEEE Moderately Inverse":
            A = 0.0515
            B = 0.114
            p = 0.02
            IEEE(sf1, df2)
        elif row[11] == "IEEE Very Inverse":
            A = 19.61
            B = 0.491
            p = 2
            IEEE(sf1, df2)
        elif row[11] == "IEEE Extremely Inverse":
            A = 28.2
            B = 0.1217
            p = 2
            IEEE(sf1, df2)

new_column = pd.Series(new_column[: len(df2)], name="Suggested Max ELI (Ω)")
df2["Suggested Max ELI (Ω)"] = new_column
df2["Suggested Max ELI (Ω)"] = df2["Suggested Max ELI (Ω)"].apply(lambda x: "{:.2f}".format(x))
result_column = pd.Series(result_column[: len(df2)], name="Result")
df2["Result"] = result_column


def create_eli_table1(df1, doc):
    df1 = df1.fillna("")
    doc.add_heading("Earth Loop Impedance Test - Circuit Breaker", level=1)
    table_data = df1.iloc[:, 0:]
    table_str = tabulate(table_data, headers="keys", tablefmt="pipe")
    num_rows, num_cols = table_data.shape[0], table_data.shape[1]
    table = doc.add_table(rows=num_rows + 1, cols=num_cols)
    table.style = "Table Grid"
    table.autofit = False

    column_widths = {
        0: 0.2,
        1: 0.5,
        2: 0.6,
        3: 0.6,
        4: 0.8,
        5: 0.7,
        6: 0.6,
        7: 0.6,
        8: 0.55,
        9: 0.7,
        10: 0.41,
        11: 1.2,
    }

    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths.get(j, 1))
        table.cell(0, j).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Align cell text to the center

    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths[j])
        first_row_cells = table.rows[0].cells
        for cell in first_row_cells:
            cell_elem = cell._element
            tc_pr = cell_elem.get_or_add_tcPr()
            shading_elem = parse_xml(f'<w:shd {nsdecls("w")} w:fill="d9ead3"/>')
            tc_pr.append(shading_elem)

    for i, row in enumerate(table_data.itertuples(), start=1):
        for j, value in enumerate(row[1:], start=0):
            cell = table.cell(i, j)
            cell.text = str(value)
            cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Align cell text to the center
            if j == num_cols - 1:  # Apply background color only to the Result column
                result_cell = cell
                if value == "Pass":
                    shading_elm = parse_xml(
                        r'<w:shd {} w:fill="00FF00"/>'.format(nsdecls("w"))
                    )  # Green color
                    result_cell._tc.get_or_add_tcPr().append(shading_elm)
                elif value == "Fail":
                    shading_elm = parse_xml(
                        r'<w:shd {} w:fill="FF0000"/>'.format(nsdecls("w"))
                    )  # Red color
                    result_cell._tc.get_or_add_tcPr().append(shading_elm)

    for section in doc.sections:
        section.left_margin = Inches(0.2)

    font_size = 7
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
                    run.font.name = "Calibri"

    return doc


def create_eli_table2(df2, doc):
    table_data = df2.iloc[:, 0:]
    table_str = tabulate(table_data, headers="keys", tablefmt="pipe")
    num_rows, num_cols = table_data.shape[0], table_data.shape[1]
    table = doc.add_table(rows=num_rows + 1, cols=num_cols)
    table.style = "Table Grid"
    table.autofit = False

    column_widths = {
        0: 0.2,
        1: 0.44,
        2: 0.55,
        3: 0.59,
        4: 0.6,
        5: 0.58,
        6: 0.45,
        7: 0.47,
        8: 0.45,
        9: 0.41,
        10: 0.41,
        11: 0.41,
        12: 0.6,
        13: 0.9,
    }

    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths.get(j, 1))
        table.cell(0, j).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 

    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths[j])
        first_row_cells = table.rows[0].cells
        for cell in first_row_cells:
            cell_elem = cell._element
            tc_pr = cell_elem.get_or_add_tcPr()
            shading_elem = parse_xml(f'<w:shd {nsdecls("w")} w:fill="d9ead3"/>')
            tc_pr.append(shading_elem)

    for i, row in enumerate(table_data.itertuples(), start=1):
        for j, value in enumerate(row[1:], start=0):
            cell = table.cell(i, j)
            cell.text = str(value)
            cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER 
            if j == num_cols - 1:  # Apply background color only to the Result column
                result_cell = cell
                if value == "Pass":
                    shading_elm = parse_xml(
                        r'<w:shd {} w:fill="00FF00"/>'.format(nsdecls("w"))
                    )  # Green color
                    result_cell._tc.get_or_add_tcPr().append(shading_elm)
                elif value == "Fail":
                    shading_elm = parse_xml(
                        r'<w:shd {} w:fill="FF0000"/>'.format(nsdecls("w"))
                    )  # Red color
                    result_cell._tc.get_or_add_tcPr().append(shading_elm)

    for section in doc.sections:
        section.left_margin = Inches(0.2)

    font_size = 7
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
                    run.font.name = "Calibri"

    return doc


doc = Document()
doc = create_eli_table1(sf1, doc)
doc.add_paragraph("\n")
doc = create_eli_table2(df2, doc)
doc.save("ELI_Socket.docx")


