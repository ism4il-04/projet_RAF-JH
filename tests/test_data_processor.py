import pytest
import pandas as pd
from core.data_processor import DataProcessor

def test_validate_dataframe():
    df = pd.DataFrame({'a': [1], 'b': [2]})
    assert DataProcessor.validate_dataframe(df, ['a', 'b'])[0] is True
    assert DataProcessor.validate_dataframe(df, ['a', 'c'])[0] is False

def test_create_connection_dict():
    df = pd.DataFrame({'Nom': ['P1', 'P2'], 'Val': [10, 20]})
    result = DataProcessor.create_connection_dict(df, 'Val')
    assert result == {'P1': 10, 'P2': 20}

def test_calculate_charge_jh():
    df = pd.DataFrame({'Soumise (h)': [8, 16]})
    result = DataProcessor.calculate_charge_jh(df)
    assert all(result['Charge JH'] == [1, 2])

def test_format_resource_summary():
    # Minimal test: just checks output is a DataFrame
    df = pd.DataFrame({'Ressource': ['R1'], 'Projet': ['P1'], 'Charge JH': [1], 'Montant total (Contrat) (Commande)': [100], 'Dernière Note': ['A'], 'Durée': [5]})
    conn = {'P1': 'C1'}
    phase = {'P1': 'PH1'}
    montant = {'P1': 100}
    out = DataProcessor.format_resource_summary(df, conn, phase, montant)
    assert isinstance(out, pd.DataFrame) 