PY ?= python3

.PHONY: install build test png clean

install:
	$(PY) -m pip install -r requirements.txt

build:            ## regenerate the workbook from data/sales_data.csv
	$(PY) src/build_dashboard.py --out dashboard/Sales_Performance_Dashboard.xlsx

test: build       ## structure + numeric tests (needs LibreOffice for the numeric ones)
	$(PY) -m pytest -q tests

png:              ## 4K PNG of the dashboard (needs LibreOffice + poppler)
	$(PY) scripts/render_png.py

clean:
	rm -rf build .pytest_cache __pycache__ tests/__pycache__ scripts/__pycache__
