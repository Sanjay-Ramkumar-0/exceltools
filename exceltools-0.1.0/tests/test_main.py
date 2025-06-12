import pytest
from exceltools.main import ExcelFile

@pytest.fixture
def excel_file():
    excel = ExcelFile().create_excel("test.xlsx")
    excel.set_value(1, 1, "A1")
    excel.set_value(1, 2, "B1")
    excel.set_value(2, 1, "A2")
    excel.set_value(2, 2, "B2")
    return excel

def test_remove_column(excel_file):
    excel_file.remove_column(1)
    assert excel_file.get_value(1, 1) == "B1"
    assert excel_file.get_dimensions()[1] == 1

def test_remove_row(excel_file):
    excel_file.remove_row(1)
    assert excel_file.get_value(1, 1) == "A2"
    assert excel_file.get_dimensions()[0] == 1

def test_invalid_column(excel_file):
    with pytest.raises(ValueError):
        excel_file.remove_column(3)  # Out of range

def test_invalid_row(excel_file):
    with pytest.raises(ValueError):
        excel_file.remove_row(3)  # Out of range