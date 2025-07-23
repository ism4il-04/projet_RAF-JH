import pytest
import pandas as pd
from core.deployment_processor import DeploymentProcessor

def test_validate_dataframe():
    df = pd.DataFrame({'a': [1], 'b': [2]})
    assert DeploymentProcessor.validate_dataframe(df, ['a', 'b'])[0] is True
    assert DeploymentProcessor.validate_dataframe(df, ['a', 'c'])[0] is False

def test_calculate_raf():
    df = pd.DataFrame({'Niveau de connexion': ['A'], 'Phase du projet': ['B']})
    out = DeploymentProcessor.calculate_raf(df)
    assert 'RAF' in out.columns

def test_calculate_monthly_raf():
    df = pd.DataFrame({'Date de MEP': ['2023-01-01'], 'RAF': [1]})
    out = DeploymentProcessor.calculate_monthly_raf(df)
    assert 'Total RAF' in out.columns or out.empty 