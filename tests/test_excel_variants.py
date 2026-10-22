"""Choose instrument variants using saved input, never tab order or QR markers."""
from io import BytesIO
import xml.etree.ElementTree as ET
import zipfile

from fastapi.testclient import TestClient
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
import pytest

from app.excel import NS, _prepare_workbook, inspect_excel
from app.excel_sheets import certificate_layouts
from app.main import create_app
from app.pdf import CertificateError


def variants(count=10, named=True):
    book = Workbook()
    register = book.active
    register.title = "1-Registro-Data"
    register["F49"] = count
    for number in [12, 10]:
        sheet = book.create_sheet(f"{7 if number == 12 else 8}-Certificado_{number}S")
        sheet["A1"] = "QCP-2026-0001"
        sheet["A3"] = f"RESULTADOS DE {number} SENSORES"
        sheet["D20"] = "Escanee este QR"
        sheet.print_area = "A1:L32"
    book.active = 1  # The active tab is deliberately the wrong one for 10.
    if named:
        book.defined_names.add(DefinedName("CantSensores", attr_text="'1-Registro-Data'!$F$49"))
    buffer = BytesIO()
    book.save(buffer)
    return buffer.getvalue()


@pytest.mark.parametrize("count", [10, 12])
def test_selects_only_the_variant_matching_the_saved_sensor_count(count):
    data = variants(count)
    layout = inspect_excel(data)
    assert layout.sheet_name.endswith(f"_{count}S")
    assert f"F49 indica {count} sensores" in layout.selection_reason
    assert len(certificate_layouts(data)) == 1
    prepared, *_, warnings = _prepare_workbook(data, False, layout=layout)
    with zipfile.ZipFile(BytesIO(prepared)) as archive:
        sheets = ET.fromstring(archive.read("xl/workbook.xml")).findall("s:sheets/s:sheet", NS)
        assert len(sheets) == 1 and sheets[0].get("name") == layout.sheet_name
    assert any(layout.selection_reason == warning for warning in warnings)


@pytest.mark.parametrize("options,match", [({"count": 11}, "cantidad de sensores")])
def test_invalid_sensor_selector_is_actionable_and_never_guessed(options, match):
    with pytest.raises(CertificateError, match=match):
        inspect_excel(variants(**options))


@pytest.mark.parametrize('options',[{'count':''},{'named':False}])
def test_missing_sensor_decision_uses_one_saved_active_format(options):
    data=variants(**options)
    layouts=certificate_layouts(data)
    assert len(layouts)==1
    assert layouts[0].sheet_name=='7-Certificado_12S'
    assert 'un solo formato' in layouts[0].selection_reason


def test_inspection_reports_the_selection_reason_before_emission(tmp_path):
    with TestClient(create_app(tmp_path, "https://example.com")) as client:
        response = client.post("/api/excel/inspect", files={"excel": ("horno.xlsx", variants())})
        assert response.status_code == 200
        assert response.json()["sheetName"] == "8-Certificado_10S"
        assert "10 sensores" in response.json()["selectionReason"]
        assert not list(tmp_path.iterdir())
