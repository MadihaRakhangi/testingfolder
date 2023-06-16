import pytest
from voltage import result, rang , create_table
import pandas as pd
from docx import Document
import os



test_df=pd.read_csv("voltage_upd.csv")


def test_result():
    assert result(2, "Public", "Lighting", 50) == "Pass"
    assert result(4, "Public", "Other", 50) == "Fail"
    # Add more test cases for different scenarios


def test_create_table():
    df = pd.DataFrame({                               #create  a test dataframe
        "Serial number": [1, 2],
        "Cable Length (m)":[60,80],
        "Insulation Type":["PVC","XLPE"]
    })

    doc = Document()                                  # Create a Document object
    doc = create_table(df, doc)                      
    temp_file = "temp_doc.docx"                        # Save the document as a temporary file
    temp_file = "temp_doc.docx"
    doc.save(temp_file)
    doc_read = Document(temp_file)                     # Read the temporary file and check if the table exists
    tables = doc_read.tables
    assert len(tables) == 1                             # Check if the first table in the document matches the DataFrame shape
    table = tables[0]
    assert len(table.rows) == df.shape[0] + 1            # Include header row
    assert len(table.columns) == df.shape[1]+1           #include the result coloumn
    os.remove(temp_file)    

def test_rang():
    # Create a sample DataFrame for testing
    df = pd.DataFrame({"VD_val": [2, 4], "Type_ISS_val": ["Public", "Public"], "PoS_val": ["Lighting", "Other"], "Dist": [50, 50]})

    # Call the rang function and validate the output
    assert rang(df.shape[0]) == ["Pass", "Fail"]
    # Add more test cases with different input data

