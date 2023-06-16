import pandas as pd
import matplotlib.pyplot as plt
import docx
from docx import Document
import csv
import io
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, RGBColor

df = pd.read_csv("tuesday.csv")

C_Loc = df.iloc[:, [0]]  # Conductor Location
Fac_A = df.iloc[:, [1]]  # Facility Area
N_Run = df.iloc[:, [2]]  # No of runs of Conductor
Cond_T = df.iloc[:, [3]]  # Conductor Type
Cond_S = df.iloc[:, [4]]  # Conductor Size (sq. mm)
Cond_L = df.iloc[:, [5]]  # Conductor Length (m)
Cond_Temp = df.iloc[:, [6]]  # Conductor Temperature (¡C)
Conti = df.iloc[:, [7]]  # Is Continuity found?
Load_IR = df.iloc[:, [8]]  # Lead Internal Resistance (?)
Cont_Resis = df.iloc[:, [9]]  # Continuity Resistance (?)


def alpha(Cond_T):
    if Cond_T == "Al":
        return 0.0038
    elif Cond_T == "Cu":
        return 0.00429
    elif Cond_T == "GI":
        return 0.00641
    elif Cond_T == "SS":
        return 0.003
    else:
        return None


df = pd.read_csv("tuesday.csv")
df['Corrected Continuity Resistance (Ω)'] = df['Continuity Resistance (?)'] - df['Lead Internal Resistance (?)']
alpha_values = df['Conductor Type'].apply(alpha)
df['Specific Conductor Resistance (MΩ/m) at 30°C'] = df['Corrected Continuity Resistance (Ω)'] / (
        1 + alpha_values * (df['Conductor Temperature (°C)'] - 30))

df['Specific Conductor Resistance (MΩ/m) at 30°C']=df['Specific Conductor Resistance (MΩ/m) at 30°C']/1000000

df['Specific Conductor Resistance (MΩ/m) at 30°C'] = df['Specific Conductor Resistance (MΩ/m) at 30°C'].apply(
    lambda x: format(x, 'E'))
df.to_csv('tuesday_updated.csv', index=False)

C_ContR = df.iloc[:, [10]]  # Corrected Continuity Resistance (?)
S_CondR = df.iloc[:, [11]]  # Specific Conductor Resistance (M?/m) at 30¡C


def result(Conti, C_ContR):
    if C_ContR <= 1:
        if Conti == "Yes":
            return "Pass"
        elif Conti == "No":
            return "Check Again"
        else:
            return "Invalid"
    elif C_ContR > 1:
        if Conti == "Yes":
            return "Fail"
        else:
            return "Fail"


def rang(length):
    res = []
    for row in range(length):
        Conti_val = Conti.iloc[row, 0]
        C_ContR_val = C_ContR.iloc[row, 0]
        res.append(result(Conti_val, C_ContR_val))
    return res


def create_table(df, doc):
    table_data = df.iloc[:, 0:]
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
        9: 0.43,
        10: 0.56,
        11: 0.56,  # Add this line
    }
    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths[j])
    for i, row in enumerate(table_data.itertuples(), start=1):
        for j, value in enumerate(row[1:], start=0):
            table.cell(i, j).text = str(value)
    Results = rang(num_rows)
    table.cell(0, num_cols).text = "Result"
    table.cell(0, num_cols).width = Inches(0.8)
    for i in range(num_rows):
        res_index = i
        table.cell(i + 1, num_cols).text = Results[res_index]
    font_size = 6.5

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
    return doc


def bar_graph(df):
    x = df["Conductor Type"]
    y = df["Conductor Temperature (°C)"]
    plt.bar(x, y)
    plt.xlabel("Conductor Type")
    plt.ylabel(" Corrected Continuity Resistance")
    plt.title("Conductor Type VS Corrected Continuity Resistance")
    plt.savefig("bar_graph.png")
    graph = io.BytesIO()
    plt.savefig(graph)
    plt.close()
    return graph

# def bar_graph(df):
#     fig, ax = plt.subplots()
#     ax.bar(df['Conductor Type'], df['Corrected Continuity Resistance (Ω)'])
#     ax.set_xlabel('Conductor Type')
#     ax.set_ylabel('Corrected Continuity Resistance')
#     ax.set_title('Conductor Type VS Corrected Continuity Resistance')
#     return fig



def pie_diagram(df):
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
    df = pd.read_csv("tuesday_updated.csv")
    doc = Document()
    doc.add_heading('RESISTANCE CONDUCTOR TEST', 0)
    doc = create_table(df, doc)
    bar_chart = bar_graph(df)
    doc.add_picture(bar_chart, width=Inches(5), height=Inches(3))
    pie_diag = pie_diagram(df)
    doc.add_picture(pie_diag, width=Inches(5), height=Inches(3))
    doc.save("resfinal.docx")


main()
