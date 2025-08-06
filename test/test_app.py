# tests/test_app.py
import pytest
import pandas as pd
from app import load_data  # Adjust import based on your app.py structure

def test_load_data():
    # Create a sample Excel file for testing
    sample_data = pd.DataFrame({"Product": ["A", "B"], "Sales": [100, 200]})
    sample_data.to_excel("test_data.xlsx", index=False)

    # Test the load_data function
    df = load_data("test_data.xlsx")
    assert not df.empty, "DataFrame should not be empty"
    assert list(df.columns) == ["Product", "Sales"], "Columns should match input data"
    assert len(df) == 2, "DataFrame should have 2 rows"
