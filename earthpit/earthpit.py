import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO
import io

G="earthpit.csv"
ef = pd.read_csv(G)


def Earth_result(Elec_DistRatio, Mea_EarthResist):
    if Mea_EarthResist <= 2 and Elec_DistRatio >= 1:
        return "PASS - Test Electrodes are properly placed"
    elif Mea_EarthResist <= 2 and Elec_DistRatio < 1:
        return "PASS - Test Electrodes are not properly placed"
    elif Mea_EarthResist > 2 and Elec_DistRatio >= 1:
        return "FAIL - Test Electrodes are properly placed"
    elif Mea_EarthResist > 2 and Elec_DistRatio < 1:
        return "FAIL - Test Electrodes are not properly placed"
    else:
        return "Invalid"


def Earth_create_table(ef, doc):
    ef["Electrode Distance Ratio"] = round(
        ef["Nearest Electrode Distance"] / ef["Earth Electrode Depth"], 2
    )
    ef["Calculated Earth Resistance - Individual (Ω)"] = (
        ef["Measured Earth Resistance - Individual"] * ef["No. of Parallel Electrodes"]
    )

    ef["Result"] = ef.apply(
        lambda row: Earth_result(
            row["Electrode Distance Ratio"], row["Measured Earth Resistance - Individual"]
        ),
        axis=1,
    )

    table_data = ef.iloc[:, 0:]
    num_rows, num_cols = table_data.shape[0], table_data.shape[1]
    table = doc.add_table(rows=num_rows + 1, cols=num_cols)
    table.style = "Table Grid"
    table.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    table.autofit = False

    column_widths = {
        0: 0.25,
        1: 0.54,
        2: 0.55,
        3: 0.59,
        4: 0.48,
        5: 0.56,
        6: 0.54,
        7: 0.59,
        8: 0.37,
        9: 0.55,
        10: 0.8,
    }

    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths.get(j, 1))

    for i, row in enumerate(table_data.itertuples(), start=1):
        for j, value in enumerate(row[1:], start=0):
            table.cell(i, j).text = str(value)

    font_size = 7
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)

    return doc


def Earth_graph(ef):
    plt.figure(figsize=(15, 10))
    result_counts = ef["Result"].value_counts()
    plt.bar(result_counts.index, result_counts.values)
    plt.xlabel("Result")
    plt.ylabel("Count")
    plt.title("Earth Pit Electrode Test Results")
    graph = io.BytesIO()
    plt.savefig(graph)
    plt.close()
    return graph

def Earth_Pie(ef):
    plt.figure(figsize=(6, 6))
    result_counts = ef["Result"].value_counts()
    labels = result_counts.index
    values = result_counts.values
    colors = ["blue", "orange"]

    plt.pie(values, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
    plt.title("Earth Pit Electrode Test Results")
    plt.axis('equal')  # Equal aspect ratio ensures that the pie is drawn as a circle
    graph = io.BytesIO()
    plt.savefig(graph)
    plt.close()
    return graph


def main():
    ef = pd.read_csv("earthpit.csv")
    doc = Document()
    doc.add_heading("Earth Pit Electrode Test", 0)
    for section in doc.sections:
        section.left_margin = Inches(0.2)
    doc = Earth_create_table(ef, doc)
    graph_image = Earth_graph(ef)
    doc.add_picture(graph_image )
    graph_image = Earth_Pie(ef)
    doc.add_picture(graph_image)
    doc.save("Earth_Pit.docx")


main()