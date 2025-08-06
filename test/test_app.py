# tests/test_app.py
import pytest
import pandas as pd
import plotly.express as px
from app import load_data, create_bar_chart

def test_load_data():
    # Create a sample Excel file for testing
    sample_data = pd.DataFrame({"Product": ["A", "B"], "Sales": [100, 200]})
    sample_data.to_excel("test_data.xlsx", index=False)

    # Test the load_data function
    df = load_data("test_data.xlsx")
    assert not df.empty, "DataFrame should not be empty"
    assert list(df.columns) == ["Product", "Sales"], "Columns should match input data"
    assert len(df) == 2, "DataFrame should have 2 rows"

def test_create_bar_chart():
    # Test chart creation
    df = pd.DataFrame({"Product": ["A", "B"], "Sales": [100, 200]})
    fig = create_bar_chart(df, "Product", "Sales")
    assert fig is not None, "Chart should be created"
    assert fig.layout.xaxis.title.text == "Product", "X-axis title should be 'Product'"
    assert fig.layout.yaxis.title.text == "Sales", "Y-axis title should be 'Sales'"
