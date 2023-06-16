import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
import io

df = pd.read_csv("voltage.csv")

SN = df.iloc[:, [0]]  # Conductor Location
Cir_From = df.iloc[:, [1]]  # Facility Area
Cir_To = df.iloc[:, [2]]  # No of runs of Conductor
MV_From = df.iloc[:, [3]]  # Conductor Type
MV_To = df.iloc[:, [4]]  # Conductor Size (sq. mm)
NCV = df.iloc[:, [5]]  # Conductor Length (m)
Type_ISS = df.iloc[:, [6]]  # Conductor Temperature (¡C)
PoS = df.iloc[:, [7]]  # Is Continuity found?
Cable_Length = df.iloc[:, [8]]  # Lead Internal Resistance (?)
Cond_Type = df.iloc[:, [9]]
Ins_Type = df.iloc[:, [10]]  # Continuity Resistance (?)

df["Calculated Voltage Drop (V)"] = (
    df["Measured Voltage (V, L-N)[FROM]"] - df["Measured Voltage (V, L-N)[TO]"]
)
df["Voltage Drop %"] = (
    df["Calculated Voltage Drop (V)"] / df["Measured Voltage (V, L-N)[FROM]"]
) * 100
df["Voltage Drop %"] = df["Voltage Drop %"].round(decimals=2)
df.to_csv("voltage_upd.csv", index=False)

lim1 = 3
lim2 = 5
lim3 = 6
lim4 = 8



def result(VD, Type_ISS, PoS, Dist):
    if Dist <= 0:
        if VD <= lim1:
            if Type_ISS == "Public" and PoS == "Lighting":
                return "Pass"
            else:
                return "Fail"
        elif VD <= lim2:
            if Type_ISS == "Public" and PoS == "Other":
                return "Pass"
            else:
                return "Fail"
        elif VD <= lim3:
            if Type_ISS == "Private" and PoS == "Lighting":
                return "Pass"
            else:
                return "Fail"
        elif VD <= lim4:
            if Type_ISS == "Private" and PoS == "Other":
                return "Pass"
            else:
                return "Fail"
        else:
            return "Fail"

    elif 0 < Dist <= 100:
        if VD <= (lim1 + (Dist * 0.005)):
            if Type_ISS == "Public" and PoS == "Lighting":
                return "Pass"
            else:
                return "Fail"
        elif VD <= (lim2 + (Dist * 0.005)):
            if Type_ISS == "Public" and PoS == "Other":
                return "Pass"
            else:
                return "Fail"
        elif VD <= (lim3 + (Dist * 0.005)):
            if Type_ISS == "Private" and PoS == "Lighting":
                return "Pass"
            else:
                return "Fail"
        elif VD <= (lim4 + (Dist * 0.005)):
            if Type_ISS == "Private" and PoS == "Other":
                return "Pass"
            else:
                return "Fail"
        else:
            return "Fail"

    elif Dist > 100:
        if VD <= (lim1 + 0.5):
            if Type_ISS == "Public" and PoS == "Lighting":
                return "Pass"
            else:
                return "Fail"
        elif VD <= (lim2 + 0.5):
            if Type_ISS == "Public" and PoS == "Other":
                return "Pass"
            else:
                return "Fail"
        elif VD <= (lim3 + 0.5):
            if Type_ISS == "Private" and PoS == "Lighting":
                return "Pass"
            else:
                return "Fail"
        elif VD <= (lim4 + 0.5):
            if Type_ISS == "Private" and PoS == "Other":
                return "Pass"
            else:
                return "Fail"
        else:
            return "Fail"



def rang(length):
    res = []
    for row in range(0, length):
        VD_val = df.iloc[row, 12]
        Type_ISS_val = df.iloc[row, 6]
        PoS_val = df.iloc[row, 7]
        Dist = df.iloc[row, 8] - 100  # Calculate Dist for each row
        res.append(result(VD_val, Type_ISS_val, PoS_val, Dist))
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
        11: 0.56,
        12: 0.5,
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
    font_size = 7

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(font_size)
    return doc


def graph(df):
    x = df["Voltage Drop %"]
    y = df["Calculated Voltage Drop (V)"]
    plt.bar(x, y)
    plt.xlabel("Voltage Drop %")
    plt.ylabel("Calculated Voltage Drop (V)")
    plt.title("Calculated Voltage Drop (V) by Voltage Drop %")
    graph = io.BytesIO()
    plt.savefig(graph)
    plt.close()
    return graph


def main():
    doc = Document()
    doc.add_heading("RESISTANCE CONDUCTOR TEST", 0)
    doc = create_table(df, doc)
    histo_graph = graph(df)
    doc.add_picture(histo_graph, width=Inches(6))
    doc.save("voltage.docx")


main()