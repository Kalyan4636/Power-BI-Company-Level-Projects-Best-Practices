import os, subprocess, sys
import pandas as pd
import pytest

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(ROOT, "scripts"))


@pytest.fixture(scope="session")
def built(tmp_path_factory):
    """Build the workbook once per test session."""
    out = tmp_path_factory.mktemp("build") / "dash.xlsx"
    subprocess.run([sys.executable, os.path.join(ROOT, "src", "build_dashboard.py"), "--out", str(out)], check=True)
    return str(out)


@pytest.fixture(scope="session")
def df():
    d = pd.read_csv(os.path.join(ROOT, "data", "sales_data.csv"), dtype={"ID": str}, parse_dates=["Order Date"])
    d["Revenue"] = d["Units Sold"] * d["Unit Price"] * (1 - d["Discount %"])
    d["COGS"] = d["Units Sold"] * d["Unit Cost"]
    d["GP"] = d["Revenue"] - d["COGS"]
    d["Month"] = d["Order Date"].dt.month
    return d
