import pytest
from Instest import result , rang , graph , graph_pie , create_table
import pandas as pd
import io
import os
import matplotlib as plt
from matplotlib.image import imread
from matplotlib.pyplot import imshow

from matplotlib.testing.decorators import image_comparison
import tkinter
import filecmp
from docx import Document


test_df=pd.read_csv("Insulate.csv")

def test_result():
    assert result(10, 250, 0.5) == "Satisfactory"
    assert result(100, 500, 0.8) == "Unsatisfactory"


def test_rang():
    df = pd.DataFrame({                                  #create  a test dataframe
         "Nominal Circuit Voltage": [10, 100, 400],
         "Test Voltage": [250, 500, 1900],
         "Insulator Resistance": [0.5, 1.8, 0.9]
    })
    assert rang(df.shape[0]) == ["Satisfactory", "Satisfactory", "Unsatisfactory"]
    

def test_create_table():
    df = pd.DataFrame({                               #create  a test dataframe
        "Serial number": [1, 2, 3],
        "Location": ["A", "B", "C"],
        "Nominal Circuit Voltage": [100, 200, 300]
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
    os.remove(temp_file)                                 # Clean up the temporary file


# def test_graph():
   
#     df = pd.DataFrame({
#         "Location": ["A", "B", "C"],
#         "Nominal Circuit Voltage": [100, 200, 300]
#     })
#     graph_obj = graph(df)                         #create graph func and graph object
#     assert isinstance(graph_obj, io.BytesIO)      #verify the graph object
#     graph_obj.seek(0)  
#     loaded_graph = imread(graph_obj)                           # Reset the stream position to the beginning
#     # loaded_graph = plt.imread(graph_obj)           # Load the graph from the stream
#     assert loaded_graph is not None                 # Perform assertions on the loaded graph
#     plt.imshow(loaded_graph)
#     plt.show()