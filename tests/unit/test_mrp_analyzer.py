import pytest
from src.core.mrp_analyzer import MRPAnalyzer

def test_mrp_analyzer_initialization():
    analyzer = MRPAnalyzer()
    assert analyzer is not None

def test_mrp_config():
    from src.core.mrp_analyzer import MRPConfig
    config = MRPConfig()
    assert isinstance(config.REQUIRED_COLUMNS, list)
    assert len(config.REQUIRED_COLUMNS) > 0
