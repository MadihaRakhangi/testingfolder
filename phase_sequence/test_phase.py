import pytest
from phase_seq import result ,rang, create_table
import pandas as pd
import io
import os
import matplotlib as plt
from docx import Document


test_df=pd.read_csv("phasesequence.csv")

def test_result():
    assert result("RYB") == "CLOCKWISE"
    assert result("RBY") == "ANTICLOCKWISE"

# def test_rang():
#     df = pd.DataFrame({                                  #create  a test dataframe
#          "Phase Sequence":["RYB","RBY"]
#     })
#     assert rang(df.shape[0]) == ["CLOCKWISE", "ANTICLOCKWISE"]

def test_rang():
    df = pd.DataFrame({
        "Phase Sequence": ["RYB", "RBY"]
    })
    assert rang(df) == ["CLOCKWISE", "ANTICLOCKWISE"]

def test_create_table():
    df = pd.DataFrame({                               #create  a test dataframe
        "Serial number": [1, 2],
        "Phase Sequence":["RYB","RBY"]
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