"""The Excel workbook must agree with the independent Python model (formula audit)."""
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
import market_model as MM  # noqa: E402

WB = load_workbook(ROOT / "model" / "ev2w_service_market.xlsx", data_only=True)
BLOCK = {"Bear": 4, "Base": 18, "Bull": 32}


def cached(sheet, ref):
    v = WB[sheet][ref].value
    if v is None:
        pytest.skip("workbook saved without calculated values")
    return v


@pytest.mark.parametrize("name", ["Bear", "Base", "Bull"])
def test_totals_match_python(name):
    py = MM.run(name)
    r0 = BLOCK[name]
    for k, y in enumerate(MM.YEARS):
        col = chr(ord("C") + k)
        assert abs(cached("Model", f"{col}{r0 + 8}") - py[y]["total"]) < 0.01, (name, y)
        assert abs(cached("Model", f"{col}{r0 + 2}") - py[y]["fleet"]) < 0.5, (name, y)


def test_no_battery_market_before_first_vintage_ages():
    py = MM.run("Base")
    assert py[2027]["battery"] == 0 and py[2028]["battery"] > 0      # FY22 vintage reaches age 6 in FY28


def test_scenarios_ordered():
    t = {n: MM.run(n)[2031]["total"] for n in MM.SCENARIOS}
    assert t["Bear"] < t["Base"] < t["Bull"]


def test_sensitivity_centre_equals_base():
    assert abs(cached("Sensitivity", "D22") - MM.run("Base")[2031]["total"]) < 0.01
