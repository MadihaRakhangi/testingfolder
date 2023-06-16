import pandas as pd
import matplotlib.pyplot as plt
import docx
from docx import Document
import csv
import io
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor

M = "floor.csv"
df = pd.read_csv(M)

df['Applied Test Voltage (V)'] = pd.to_numeric(df['Applied Test Voltage (V)'], errors='coerce')
df['Measured Output Current (mA)'] = pd.to_numeric(df['Measured Output Current (mA)'], errors='coerce')

df['EffectiveResistance'] = df['Applied Test Voltage (V)'] / df['Measured Output Current (mA)']
df.to_csv('floorfinal.csv', index=False)



def result(Nom_EVolt, ATV, Eff_Floor, Dist_loc):
    if Nom_EVolt <= 500 and Dist_loc >= 1:
        if ATV == Nom_EVolt and Eff_Floor >= 50:
            return "pass"
        else:
            return "fail"
    elif Nom_EVolt > 500 and Dist_loc >= 1:
        if ATV == Nom_EVolt and Eff_Floor >= 100:
            return "pass"
        else:
            return "fail"
    elif Dist_loc <= 1:
        return "fail"
    else:
        return "Invalid input"


def rang(length):
    res = []
    for row in range(length):
        Nom_EVolt = df.iloc[row, 5]
        Dist_loc = df.iloc[row, 4]
        ATV = df.iloc[row, 6]
        Eff_Floor = df.iloc[row, 8]
        res.append(result(Nom_EVolt, ATV, Eff_Floor, Dist_loc))
    return res


def create_table(df, doc):
    table_data = df.iloc[:, :]
    num_rows, num_cols = table_data.shape
    table = doc.add_table(rows=num_rows + 1, cols=num_cols + 1)
    table.style = "Table Grid"
    table.autofit = False

    column_widths = {
        0: 0.4,
        1: 0.6,
        2: 0.7,
        3: 0.65,
        4: 0.5,
        5: 0.5,
        6: 0.5,
        7: 0.5,
        8: 0.5,
        9: 0.5,
        10:0.7,
        11:0.6,
    }

    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths[j])

    for i, row in enumerate(table_data.itertuples(), start=1):
        for j, value in enumerate(row[1:], start=0):
            if isinstance(value, float):
                value = "{:.2f}".format(value)
            table.cell(i, j).text = str(value)

    Results = rang(num_rows)
    table.cell(0, num_cols).text = "Result"
    table.cell(0, num_cols).width = Inches(0.6)
    for i in range(num_rows):
        table.cell(i + 1, num_cols).text = Results[i]
    table.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    font_size = 8

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)

    return doc



def scatter_graph(df):
    x = df["Distance from previous test location (m)"]
    y = df["EffectiveResistance"]
    plt.scatter(x, y)
    plt.xlabel("Distance from previous test location (m)")
    plt.ylabel("Effectivefloor")
    plt.title("Distance from previous test location (m) VS Effectivefloor")
    plt.savefig("scatter_graph.png")
    graph = io.BytesIO()
    plt.savefig(graph)
    plt.close()
    return graph


def pie_chart(df):
    df['Result'] = rang(df.shape[0])
    df_counts = df['Result'].value_counts()
    labels = df_counts.index.tolist()
    values = df_counts.values.tolist()

    plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
    plt.axis('equal')
    plt.title('Test Results')
    plt.savefig('pie_chart.png')
    graph = io.BytesIO()
    plt.savefig(graph)
    plt.close()
    return graph


def main():
    M = "floor.csv"
    df = pd.read_csv("floorfinal.csv")
    doc = docx.Document()
    doc.add_heading('FLOOR TEST', 0)
    doc = create_table(df, doc)
    scatter_graph_img = scatter_graph(df)
    doc.add_picture(scatter_graph_img, width=Inches(5), height=Inches(3))
    pie_chart_img = pie_chart(df)
    doc.add_picture(pie_chart_img, width=Inches(5), height=Inches(3))
    doc.save("floor.docx")


main()
