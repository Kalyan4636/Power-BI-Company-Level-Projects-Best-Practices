"""Structural + numerical tests.

Numerical tests recalculate the workbook in headless LibreOffice and compare every
chart data source against an independent pandas calculation.
"""
import shutil
import openpyxl
import pytest
from recalc import recalc

needs_lo = pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Dashboard filter cells (see docs/architecture.md)
CELL = {"Region": "B6", "Country": "D6", "Channel": "F6", "Customer": "H6", "Category": "J6", "Product": "L6", "From": "N6", "To": "P6"}


def test_structure(built):
    wb = openpyxl.load_workbook(built)
    assert wb.sheetnames == ["Dashboard", "Analysis", "Data", "Lists", "Guide"]
    d = wb["Dashboard"]
    assert len(d._charts) == 6
    assert len(d.data_validations.dataValidation) == 8
    for nm in ["sel_Region", "sel_Country", "sel_Channel", "sel_Customer", "sel_Category", "sel_Product", "sel_FromMonth", "sel_ToMonth", "m_From", "m_To"]:
        assert nm in wb.defined_names
    assert "SalesData" in wb["Data"].tables
    assert d.protection.sheet is True


def test_no_hardcoded_kpis(built):
    """KPI cells on the dashboard must be formulas, never pasted values."""
    d = openpyxl.load_workbook(built)["Dashboard"]
    for c in ["B9", "E9", "H9", "K9", "N9", "Q9"]:
        assert str(d[c].value).startswith("="), c


def _filtered(built, tmp_path, **flt):
    wb = openpyxl.load_workbook(built)
    for k, v in flt.items():
        wb["Dashboard"][CELL[k]] = v
    p = tmp_path / "t.xlsx"; wb.save(p)
    return openpyxl.load_workbook(recalc(str(p)), data_only=True)


def _errors(wb):
    bad = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("#"):
                    bad.append((ws.title, c.coordinate, c.value))
    return bad


SCENARIOS = {
    "all": ({}, lambda d: d.Revenue > -1),
    "camping_germany": ({"Category": "Camping", "Country": "Germany"},
                        lambda d: (d["Product Category"] == "Camping") & (d.Country == "Germany")),
    "west_online_returning_marAug": ({"Region": "Western Europe", "Channel": "Online Store", "Customer": "Returning", "From": "Mar", "To": "Aug"},
                                     lambda d: (d.Region == "Western Europe") & (d["Sales Channel"] == "Online Store") & (d["Customer Type"] == "Returning") & d.Month.between(3, 8)),
}


@needs_lo
@pytest.mark.parametrize("name", SCENARIOS)
def test_numbers_match_pandas(name, built, df, tmp_path):
    flt, mask_fn = SCENARIOS[name]
    wb = _filtered(built, tmp_path, **flt)
    assert not _errors(wb), "formula errors found"
    a = wb["Analysis"]; s = df[mask_fn(df)]
    # KPI block
    assert a["C9"].value == pytest.approx(s.Revenue.sum(), abs=0.01)
    assert a["C10"].value == pytest.approx(s.GP.sum(), abs=0.01)
    assert a["C12"].value == s["Units Sold"].sum()
    assert a["C13"].value == len(s)
    # every chart that ignores its own filter must equal pandas with that one filter removed
    def other(**drop):
        m = df.Revenue > -1
        full = {"Region": df.Region, "Country": df.Country, "Channel": df["Sales Channel"], "Customer": df["Customer Type"],
                "Category": df["Product Category"], "Product": df.Product}
        for k, v in flt.items():
            if k in full and k not in drop and v != "All": m &= full[k] == v
        if "Month" not in drop and ("From" in flt or "To" in flt):
            m &= df.Month.between(MONTHS.index(flt.get("From", "Jan")) + 1, MONTHS.index(flt.get("To", "Dec")) + 1)
        return df[m]
    # category block rows 39-43 (sorted: item G, selected H, excluded I) -> selected+excluded == revenue ignoring Category
    cat_total = sum((a.cell(r, 8).value or 0) + (a.cell(r, 9).value or 0) for r in range(39, 44))
    assert cat_total == pytest.approx(other(Category=1).Revenue.sum(), abs=0.01)
    ctry_total = sum((a.cell(r, 8).value or 0) + (a.cell(r, 9).value or 0) for r in range(63, 80))
    assert ctry_total == pytest.approx(other(Country=1).Revenue.sum(), abs=0.01)
    # selected-only totals equal fully filtered revenue
    assert sum(a.cell(r, 8).value for r in range(39, 44)) == pytest.approx(s.Revenue.sum(), abs=0.01)
    # monthly: selected months sum == filtered revenue
    assert sum(a.cell(r, 5).value for r in range(23, 35)) == pytest.approx(s.Revenue.sum(), abs=0.01)
    # top-5 products (units), selected series
    exp = s.groupby("Product")["Units Sold"].sum().nlargest(5)
    got = sorted([a.cell(r, 8).value for r in range(84, 89) if a.cell(r, 8).value], reverse=True)
    assert got == sorted(exp.tolist(), reverse=True)[:len(got)]
