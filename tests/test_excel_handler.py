import pytest
import pandas as pd
from core.excel_handler import ExcelHandler
from unittest.mock import patch, MagicMock

def test_read_excel_file_not_found():
    with pytest.raises(FileNotFoundError):
        ExcelHandler.read_excel('not_a_file.xlsx')

@patch('pandas.read_excel')
def test_read_excel_success(mock_read):
    mock_read.return_value = pd.DataFrame({'a': [1]})
    with patch('os.path.exists', return_value=True):
        df = ExcelHandler.read_excel('file.xlsx')
        assert 'a' in df.columns

def test_create_pivot_table():
    df = pd.DataFrame({'A': ['x', 'y'], 'B': [1, 2]})
    pivot = ExcelHandler.create_pivot_table(df, values='B', index=['A'])
    assert 'B' in pivot.columns

@patch('core.excel_handler.pd.ExcelWriter')
def test_write_excel(mock_writer):
    df = pd.DataFrame({'A': [1]})
    with patch('os.makedirs'), patch('os.path.exists', return_value=False):
        ExcelHandler.write_excel(df, 'out.xlsx')
    assert mock_writer.called

@patch('core.excel_handler.pd.ExcelWriter')
def test_write_multiple_sheets(mock_writer):
    dfs = {'Sheet1': pd.DataFrame({'A': [1]})}
    with patch('os.makedirs'), patch('os.path.exists', return_value=False):
        ExcelHandler.write_multiple_sheets(dfs, 'out.xlsx')
    assert mock_writer.called 