

import pytest
import pandas as pd
from docx import Document
import os
import io
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image
import matplotlib.pyplot as plt
from resistance import alpha, result, rang, create_table, bar_graph, pie_diagram
from docx import Document

@pytest.fixture
def sample_dataframe():
                                                           
    df = pd.DataFrame({                                           # Define a sample dataframe for testing
        'Conductor Type': ['Al', 'Cu'],
        'Conductor Temperature (°C)': [20, 40],
        'Continuity Resistance (Ω)': [1.1, 2.5],
        'Corrected Continuity Resistance (Ω)':[1 ,2.3],
        'Lead Internal Resistance (Ω)': [0.05, 0.1],
    })
    return df



def test_alpha():
    assert alpha('Al') == 0.0038
    assert alpha('Cu') == 0.00429
    assert alpha('GI') == 0.00641
    assert alpha('SS') == 0.003
    assert alpha('Invalid') is None



def test_result():
    assert result('Yes', 0.9) == 'Pass'
    assert result('No', 0.9) == 'Check Again'
    assert result('Yes', 1.2) == 'Fail'
    assert result('Invalid', 0.9) == 'Invalid'




def test_rang():
    df = pd.DataFrame({                                                          #create  a test dataframe
         "Continuity Resistance (Ω)": ["Yes","No",],
         "Corrected Continuity Resistance (Ω)": [0.9,1.2],
         
    })
    assert rang(df.shape[0]) == ["Pass","Fail"]



def test_create_table(sample_dataframe):
    doc = Document()                                                                   # Create a Document object
    doc = create_table(sample_dataframe, doc)                       
    temp_file = "temp_doc.docx"                                                 # Save the document as a temporary file
    temp_file = "temp_doc.docx"
    doc.save(temp_file)
    doc_read = Document(temp_file)                                              # Read the temporary file and check if the table exists
    tables = doc_read.tables
    assert len(tables) == 1                                                     # Check if the first table in the document matches the DataFrame shape
    table = tables[0]
    assert len(table.rows) == sample_dataframe.shape[0] + 1                     # Include header row
    assert len(table.columns) == sample_dataframe.shape[1]+1                    #include the result coloumn
    os.remove(temp_file)    

                                                                          
def test_bar_graph(sample_dataframe):
    graph = bar_graph(sample_dataframe)
    image = Image.open(graph)                                                                           # Convert BytesIO object to image using PIL
    image_array = np.array(image)                                                                       # Convert the image to a NumPy array
    assert image_array.shape[0] > 0                                                                     # Check if the image has a non-zero height
    assert image_array.shape[1] > 0                                                                     # Check if the image has a non-zero width
    graph.close()  

def test_pie_diagram(sample_dataframe):
    graph = pie_diagram(sample_dataframe)
   
