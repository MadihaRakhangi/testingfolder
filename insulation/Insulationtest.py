import pandas as pd
import matplotlib.pyplot as plt
import docx
import csv
import io
from docx.shared import Inches
from docx.shared import Pt

mf = pd.read_csv("Insulate.csv")


def insualtion_result(Nom_CVolt, T_Volt, Insu_R):
    if Nom_CVolt <= 50:
        if Insu_R >= 0.5 and T_Volt == 250:
            return "Satisfactory"
        else:
            return "Unsatisfactory"
    elif 50 < Nom_CVolt <= 500:
        if Insu_R >= 1 and T_Volt == 500:
            return "Satisfactory"
        else:
            return "Unsatisfactory"
    elif Nom_CVolt > 500:
        if Insu_R >= 1 and T_Volt == 1000:
            return "Satisfactory"
        else:
            return "Unsatisfactory"
    else:
        return "Invalid input"

def insulationrang(length):
    res2 = []
    for row in range(length):  # Adjusted the range to start from 0
        Nom_CVolt = mf.iloc[row, 7]
        T_Volt = mf.iloc[row, 9]
        Insu_R = mf.iloc[row, 13]
        res2.append(insualtion_result(Nom_CVolt, T_Volt, Insu_R))
        print(Nom_CVolt)
    return res2


def insulation_table(df, doc):
    table_data = df.iloc[:, 0:]
    num_rows, num_cols = table_data.shape
    table = doc.add_table(rows=num_rows + 1, cols=num_cols + 1)  # Add +1 for the "Result" column
    table.style = "Table Grid"
    table.autofit = False
    column_widths = {
        0: 0.2,
        1: 0.51,
        2: 0.55,
        3: 0.54,
        4: 0.38,
        5: 0.4,
        6: 0.5,
        7: 0.48,
        8: 0.5,
        9: 0.43,
        10: 0.4,
        11: 0.4,
        12: 0.4,
        13: 0.4,
        num_cols: 0.8  # Width for the "Result" column
    }
    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths[j])
    for i, row in enumerate(table_data.itertuples(), start=1):
        for j, value in enumerate(row[1:], start=0):
            table.cell(i, j).text = str(value)
    Results = insulationrang(num_rows)
    table.cell(0, num_cols).text = "Result"
    table.cell(0, num_cols).width = Inches(column_widths[num_cols])  # Set width for the "Result" column
    for i in range(0, num_rows):
        res_index = i - 1
        table.cell(i + 1, num_cols).text = Results[res_index]
    font_size = 6.5

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)

    return doc

def insulation_graph(df):
    df = pd.read_csv("Insulate.csv")
    x = df["Location"]
    y = df["Nominal Circuit Voltage"]
    plt.bar(x, y)
    plt.xlabel("Location")
    plt.ylabel("Nominal Circuit Voltage")
    plt.title("Nominal Circuit Voltage by Location")
    plt.savefig("chart.png")
    graph2 = io.BytesIO()
    plt.savefig(graph2)
    plt.close()
    return graph2


def insulation_pie(df):
    df = pd.read_csv("Insulate.csv")
    earthing_system_counts = df["Earthing System"].value_counts()
    plt.pie(earthing_system_counts, labels=earthing_system_counts.index, autopct="%1.1f%%")
    plt.axis("equal")
    plt.title("Earthing System Distribution")
    plt.savefig("pie_chart.png")
    graph4 = io.BytesIO()
    plt.savefig(graph4)
    plt.close()
    return graph4




def main():
    df = pd.read_csv("Insulate.csv")
    doc = docx.Document()
    doc = insulation_table(df, doc)

    graph_insulation = insulation_graph(df)
    doc.add_picture(graph_insulation)

    pie_insualtion= insulation_pie(df)
    doc.add_picture(pie_insualtion)
 
    doc.save("outputTEST.docx")


main()