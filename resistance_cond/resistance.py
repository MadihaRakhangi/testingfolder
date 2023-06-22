import pandas as pd
import matplotlib.pyplot as plt
import docx
from docx import Document
import io
from docx.shared import Inches
from docx.shared import Pt

E="resistance.csv"
rf=pd.read_csv(E)

def alpha(Cond_T):                               # Function to calculate alpha value based on conductor typ      
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

# Calculate Corrected Continuity Resistance
rf['Corrected Continuity Resistance (Ω)'] = rf['Continuity Resistance (?)'] - rf['Lead Internal Resistance (?)']

# Calculate Specific Conductor Resistance at 30°C
alpha_values = rf['Conductor Type'].apply(alpha)
rf['Specific Conductor Resistance (MΩ/m) at 30°C'] = rf['Corrected Continuity Resistance (Ω)'] / (
    1 + alpha_values * (rf['Conductor Temperature (°C)'] - 30))
rf['Specific Conductor Resistance (MΩ/m) at 30°C'] /= 1000000
rf['Specific Conductor Resistance (MΩ/m) at 30°C'] = rf['Specific Conductor Resistance (MΩ/m) at 30°C'].apply(
    lambda x: format(x, 'E'))
rf.to_csv('resistance_updated.csv', index=False)


def resc_result(Conti, C_ContR):
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


def resc_rang(rf):
    res5 = []
    for row in range(0, len(rf)):
        Conti = rf.iloc[row]['Is Continuity found?']
        C_ContR = rf.iloc[row]['Corrected Continuity Resistance (Ω)']
        res5.append(resc_result(Conti, C_ContR))
    return res5


def resc_table(rf, doc):
    table_data = rf.iloc[:, 0:]
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
        12:0.5   # Add this line
    }
    for j, col in enumerate(table_data.columns):
        table.cell(0, j).text = col
        table.cell(0, j).width = Inches(column_widths[j])
    for i, row in enumerate(table_data.itertuples(), start=1):
        for j, value in enumerate(row[1:], start=0):
            table.cell(i, j).text = str(value)
    Results = resc_rang(rf)
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


def resc_graph(rf):
    x = rf["Conductor Type"]
    y = rf["Corrected Continuity Resistance (Ω)"]
    plt.bar(x, y)
    plt.xlabel("Conductor Type")
    plt.ylabel("Corrected Continuity Resistance")
    plt.title("Conductor Type VS Corrected Continuity Resistance")
    graph8= io.BytesIO()
    plt.savefig(graph8)
    plt.close()
    return graph8


def resc_pie(rf):
    rf['Result'] = resc_rang(rf)
    rf_counts = rf['Result'].value_counts()
    labels = rf_counts.index.tolist()
    values = rf_counts.values.tolist()

    plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=90)
    plt.axis('equal')
    plt.title('Test Results')
    graph9 = io.BytesIO()
    plt.savefig(graph9)
    plt.close()
    return graph9


def main():
    doc = Document()
    doc.add_heading('RESISTANCE CONDUCTOR TEST', 0)
    doc = resc_table(rf, doc)
    bar_resc = resc_graph(rf)
    doc.add_picture(bar_resc, width=Inches(5), height=Inches(3))
    pie_resc= resc_pie(rf)
    doc.add_picture(pie_resc, width=Inches(5), height=Inches(3))
    doc.save("resfinal.docx")


main()
