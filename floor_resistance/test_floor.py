import pytest
from floortest import result ,rang, create_table
import pandas as pd
import io
import os
import matplotlib as plt
from docx import Document


test_df=pd.read_csv("floorfinal.csv")

def test_result():
    assert result(230, 230, 55,1) == "pass"
    assert result(650, 650, 87,2) == "fail"

def test_rang():
    df = pd.DataFrame({                                  #create  a test dataframe
         "Nominal Voltage to Earth of System (V)": [230,230, 230],
         "Applied Test Voltage (V)": [230, 230,270],
         "EffectiveResistance": [50, 100, 50],
         "Distance from previous test location (m)":[1,1,1,]
    })
    assert rang(df.shape[0]) == ["pass","fail", "pass"]

def test_create_table():
    df = pd.DataFrame({                               #create  a test dataframe
        "Serial number": [1, 2, 3],
        "Applied Test Voltage (V)": [230, 230,270],
        "EffectiveResistance": [50, 100, 50],
        "Distance from previous test location (m)":[1,1,1,]
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


