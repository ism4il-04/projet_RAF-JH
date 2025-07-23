import pytest
import pandas as pd
from core.raf_processor import RAFProcessor
from unittest.mock import MagicMock

def test_calculate_raf():
    df = pd.DataFrame({'Niveau de connexion': ['A'], 'Phase du projet': ['B']})
    out = RAFProcessor.calculate_raf(df)
    assert 'RAF' in out.columns

def test_add_raf_to_workbook():
    wb = MagicMock()
    df = pd.DataFrame({'RAF': [1, 2]})
    result = RAFProcessor.add_raf_to_workbook(wb, df)
    assert result is wb

def test_create_raf_summary_sheet():
    wb = MagicMock()
    df = pd.DataFrame({'Date de MEP': ['2023-01-01'], 'RAF': [1]})
    result = RAFProcessor.create_raf_summary_sheet(wb, df)
    assert result is wb 