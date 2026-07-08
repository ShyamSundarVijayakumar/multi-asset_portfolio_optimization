import pytest
import pandas as pd
import json
from pathlib import Path
from src.utils import check_is_cached, update_cache_timestamp, CACHE_LOG_PATH, read_clean_dataframe

# Fixture to create a dummy data file for testing
@pytest.fixture
def dummy_data(tmp_path):
    d = tmp_path / "data"
    d.mkdir()
    f = d / "test.csv"
    df = pd.DataFrame({
        'ISIN': ['A1', 'A1', 'A1', 'B2'], # Note the duplicate A1
        'Security Name': ['Stock1', 'Stock1', 'Stock1', 'Stock2'],
        'Is_Simulation': ['False', 'True', 'False', 'False']
    })
    df.to_csv(f, index=False)
    return f

def test_read_clean_dataframe_logic(dummy_data):
    # Test if it removes duplicates and simulations
    df = read_clean_dataframe(dummy_data)
    
    # Assert A1 duplicate is gone, and Simulation row is gone
    assert len(df) == 2
    
    # Verify the simulation row is actually gone
    assert 'True' not in df['Is_Simulation'].astype(str).values
    
    # Verify the duplicate A1 is gone (only one A1 remains)
    assert df['ISIN'].value_counts()['A1'] == 1

def test_cache_integration(tmp_path, monkeypatch):
    # 1. Create a temporary path for the test cache
    test_cache = tmp_path / ".engine_cache.json"
    
    # 2. Monkeypatch the CACHE_LOG_PATH in utils to point to our test file
    # This prevents the test from touching your real data!
    monkeypatch.setattr("src.utils.CACHE_LOG_PATH", test_cache)
    
    # 3. Test the "Update" function
    update_cache_timestamp("test_engine")
    
    # 4. Test the "Check" function (which now sees our test_cache)
    assert check_is_cached("test_engine") is True
    assert check_is_cached("wrong_engine") is False