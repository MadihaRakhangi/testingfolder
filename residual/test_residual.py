# import pytest
# import pandas as pd
# import matplotlib.pyplot as plt
# from docx import Document
# from residual import Resi_create_table, Resi_graph, Resi_pie_chart

# @pytest.fixture
# def sample_data():
#     data = {
#         "Type": ["AC", "A"],
#         "Test Current (mA)": [10, 20],
#         "Rated Residual Operating Current,I?n (mA)": [20, 40],
#         "Device Tripped": ["Yes", "Yes"],
#         "Trip Time (ms)": ["100", "200"],
#     }
#     return pd.DataFrame(data)

# def test_Resi_create_table(sample_data):
#     doc = Document()
#     table_data = Resi_create_table(sample_data, doc)
#     # Write your assertions to check the table_data
#     assert table_data is not None
#     assert isinstance(table_data, pd.DataFrame)

# def test_Resi_graph(sample_data):
#     graph_filename = Resi_graph(sample_data)
#     # Write your assertions to check the graph_filename
#     assert graph_filename is not None
#     assert isinstance(graph_filename, str)
#     assert graph_filename.endswith(".png")

# def test_Resi_pie_chart(sample_data):
#     pie_filename = Resi_pie_chart(sample_data)
#     # Write your assertions to check the pie_filename
#     assert pie_filename is not None
#     assert isinstance(pie_filename, str)
#     assert pie_filename.endswith(".png")


