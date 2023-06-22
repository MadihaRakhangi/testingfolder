import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
import io

D="pol.csv"
af = pd.read_csv("pol.csv")





def polarity_result(line_neutral):
    if line_neutral == 230:
        return "OK"
    else :
        return "REVERSE"
    

def polarityrang(length):
    res4 = []
    for row in range(0, length):
        line_neutral = af.iloc[row, 5]
        res4.append(polarity_result(line_neutral))
    return res4


def polarity_table(af, doc):
    table_data = af.iloc[:, 0:]
    num_rows, num_cols = table_data.shape
    table = doc.add_table(rows=num_rows + 1, cols=num_cols + 1)
    table.style = "Table Grid"
    table.autofit = False
    column_widths = {
        0: 0.2,
        1: 0.51,
        2: 0.55,
        3: 0.54,
        4: 0.38,
        5: 0.56,
        6: 0.5,
        7: 0.48,
        8: 0.71,
       
    }
    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths[j])
    for i, row in enumerate(table_data.itertuples(), start=1):
        for j, value in enumerate(row[1:], start=0):
            table.cell(i, j).text = str(value)
    Results = polarityrang(num_rows)
    table.cell(0, num_cols).text = "Result"
    table.cell(0, num_cols).width = Inches(0.8)
    for i in range(num_rows):
        res_index = i
        table.cell(i + 1, num_cols).text = Results[res_index]
    font_size = 7

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
    return doc


def polarity_graph(af):
    x = af["Type of Supply"]
    y = af["Line to Neutral Voltage (V)"]
    plt.bar(x, y)
    plt.xlabel("Type of Supply")
    plt.ylabel("Line to Neutral Voltage (V)")
    plt.title("Type of Supply Type of Supply VS  Line to Neutral Voltage (V)")
    graph7 = io.BytesIO()
    plt.savefig(graph7)
    plt.close()
    return graph7

def polarity_pie(af):
    af['Result'] = polarityrang(af.shape[0])
    af_counts = af['Result'].value_counts()
    labels = af_counts.index.tolist()
    values = af_counts.values.tolist()

    plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
    plt.axis('equal')
    plt.title('Polarity Results')
    graph = io.BytesIO()
    plt.savefig(graph)
    plt.close()
    return graph

def main():
    doc = Document()
    doc.add_heading("POLARITY TEST", 0)
    doc = polarity_table(af, doc)

    graph_polarity = polarity_graph(af)
    doc.add_picture(graph_polarity, width=Inches(6))
    
    pie_polarity= polarity_pie(af)
    doc.add_picture(pie_polarity, width=Inches(5), height=Inches(3))

    doc.save("polarity.docx")


main()