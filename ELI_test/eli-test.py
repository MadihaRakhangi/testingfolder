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
from docx.shared import RGBColor
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

df = pd.read_csv("eli-test.csv")
ef = pd.read_csv("sugg-max-eli.csv")

df["Device Rating (A)"] = df["Device Rating (A)"].astype(int)

Device_Rating = df["Device Rating (A)"]
No_phase = df["No. of Phases"]
T_Curve = df["Trip Curve"]
new_column = []
result_column = []

for index, row in df.iterrows():
    rating = row[7]
    trip = row[9]
    result_row = ef[ef["Device Rating (A)"] == rating]
    val = result_row[trip].values[0]
    new_column.append(round(val, 2))

    # Apply condition to determine result
    if row["No. of Phases"] == 3:
        if row["L1-ELI"] < val and row["L2-ELI"] < val and row["L3-ELI"] < val:
            result_column.append("Pass")
        else:
            result_column.append("Fail")
    elif row["No. of Phases"] == 1:
        if row["L1-ELI"] < val:
            result_column.append("Pass")
        else:
            result_column.append("Fail")
    else:
        result_column.append("N/A")

df["Suggested Max ELI (Ω)"] = new_column
df["Suggested Max ELI (Ω)"] = df["Suggested Max ELI (Ω)"].apply(lambda x: "{:.2f}".format(x))
df["Result"] = result_column


def create_eli_table(df):
    doc = Document()
    doc.add_heading("Earth Loop Impedance Test - Circuit Breaker", level=1)

    table_data = df.iloc[:, 0:]
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
        12: 0.4,
        13: 0.4,
        14: 0.4,
        15: 0.45,
        16: 0.5,
        17: 0.45,
    }

    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths.get(j, 1))
    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths[j])
        first_row_cells = table.rows[0].cells
    for cell in first_row_cells:
        cell_elem = cell._element
        tc_pr = cell_elem.get_or_add_tcPr()
        shading_elem = parse_xml(
            f'<w:shd {nsdecls("w")} w:fill="d9ead3"/>'
        )
        tc_pr.append(shading_elem)

    for i, row in enumerate(table_data.itertuples(), start=1):
        for j, value in enumerate(row[1:], start=0):
            cell = table.cell(i, j)
            cell.text = str(value)
            if j == num_cols - 1:  # Apply background color only to the Result column
                result_cell = cell
                if value == "Pass":
                    shading_elm = parse_xml(
                        r'<w:shd {} w:fill="00FF00"/>'.format(nsdecls('w'))
                    )  # Green color
                    result_cell._tc.get_or_add_tcPr().append(shading_elm)
                elif value == "Fail":
                    shading_elm = parse_xml(
                        r'<w:shd {} w:fill="FF0000"/>'.format(nsdecls('w'))
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

    return doc


def main():
    df = pd.read_csv("eli-test.csv")
    doc = Document()
    print(df)
    for section in doc.sections:
        section.left_margin = Inches(0.1)
    doc = create_eli_table(df)
    doc.save("eli-test.docx")


main()
