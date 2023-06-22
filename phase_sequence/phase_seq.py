import pandas as pd
import matplotlib.pyplot as plt
import docx
from docx import Document
import csv
import io
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor

P="phasesequence.csv"
pf = pd.read_csv("phasesequence.csv")


def phase_result(phase_seq):
    if phase_seq == 'RYB':
        return "CLOCKWISE"
    else:
        return "ANTICLOCKWISE"

def phaserang(pf):
    res3 = []
    phase_seqs = pf["Phase Sequence"]
    for seq in phase_seqs:
        if seq == "RYB":
            res3.append("CLOCKWISE")
        elif seq == "RBY":
            res3.append("ANTICLOCKWISE")
        else:
            res3.append("UNKNOWN")
    return res3

# def rang(df):
#     res = []
#     phase_seqs = df["Phase Sequence"]
#     for phase_seq in phase_seqs:
#         res.append(result(phase_seq))
#     return res

def phase_table(df, doc):
    table_data = df.iloc[:, :]
    num_rows, num_cols = table_data.shape
    table = doc.add_table(rows=num_rows + 1, cols=num_cols + 1)
    table.style = "Table Grid"
    table.autofit = False

    column_widths = {
        0: 0.2,
        1: 0.4,
        2: 0.4,
        3: 0.4,
        4: 0.6,
        5: 0.3,
        6: 0.3,
        7: 0.4,
        8: 0.4,
        9: 0.4,
        10: 0.4,
        11: 0.4,
        12:0.6,
    }
    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths[j])
    for i, row in enumerate(table_data.itertuples(), start=0):
        for j, value in enumerate(row[1:], start=0):
            table.cell(i + 1, j).text = str(value)
    results = phaserang(pf)

    table.cell(0, num_cols).text = "Result"
    for i, result in enumerate(results, start=1):
        table.cell(i, num_cols).text = result
        table.cell(i, num_cols).width=Inches(0.8)

    font_size = 8
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
    return doc

def phase_graph(df):
    x = df["Phase Sequence"]
    y = df["V-L3-N"]
    plt.bar(x, y)
    plt.xlabel("Phase Sequence")
    plt.ylabel("V-L3-N")
    plt.title("Phase Sequence by V-L3-N")
    graph4 = io.BytesIO()
    plt.savefig(graph4)
    plt.close()
    return graph4

def phase_pie(pf):
    pf['Result'] = phaserang(pf)
    pf_counts = pf['Result'].value_counts()
    labels = pf_counts.index.tolist()
    values = pf_counts.values.tolist()
    plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
    plt.axis('equal')
    plt.title('Test Results')
    graph5 = io.BytesIO()
    plt.savefig(graph5)
    plt.close()
    return graph5

def main():
    pf = pd.read_csv("phasesequence.csv")
    doc = Document()
    doc.add_heading('Phase Sequence test', 0)
    doc = phase_table(pf, doc)
    bar_chart = phase_graph(pf)
    doc.add_picture(bar_chart, width=Inches(5), height=Inches(3))
    pie_diag = phase_pie(pf)
    doc.add_picture(pie_diag, width=Inches(5), height=Inches(3))
    doc.save("phasesequence.docx")

main()
