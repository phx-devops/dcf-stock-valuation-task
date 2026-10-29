.DEFAULT_GOAL := help

PYTHON ?= python3
APP_DIR ?= $(abspath ../app)
OUT_DIR ?= $(APP_DIR)/output
INPUTS ?= $(APP_DIR)/inputs.json
ASSUMPTIONS ?= $(APP_DIR)/assumptions.json
MARKET_DATA ?= $(APP_DIR)/market_data.json
SOFFICE ?= soffice
PDFTOTEXT ?= pdftotext
SOFFICE_APP := /Applications/LibreOffice.app/Contents/MacOS/soffice
SOFFICE_BIN = $(shell command -v "$(SOFFICE)" 2>/dev/null || { test -x "$(SOFFICE_APP)" && printf '%s' "$(SOFFICE_APP)"; })

.PHONY: help install-dependencies extract-inputs build-workbook test-outputs

help:
	@printf '%s\n' \
		'install-dependencies  Install Python, LibreOffice, and PDF extraction dependencies' \
		'extract-inputs        Extract financial inputs from all supplied filings' \
		'build-workbook        Build the DCF workbook and summary outputs' \
		'test-outputs          Run the workbook output verifier'

install-dependencies:
	$(PYTHON) -m pip install openpyxl pytest
	@if [ -z "$(SOFFICE_BIN)" ]; then \
		if command -v brew >/dev/null 2>&1; then \
			brew install --cask libreoffice; \
		else \
			printf 'LibreOffice is missing. Install LibreOffice or make soffice available on PATH.\n' >&2; \
			exit 1; \
		fi; \
	fi
	@test -x "$(SOFFICE_BIN)" || { printf 'LibreOffice installation did not provide an executable soffice.\n' >&2; exit 1; }
	@if ! command -v "$(PDFTOTEXT)" >/dev/null 2>&1; then \
		if command -v brew >/dev/null 2>&1; then \
			brew install poppler; \
		else \
			printf 'pdftotext is missing. Install Poppler and ensure pdftotext is on PATH.\n' >&2; \
			exit 1; \
		fi; \
	fi

extract-inputs: install-dependencies
	@for filing in \
		"$(APP_DIR)/filings/NKE_10K_2024.pdf" \
		"$(APP_DIR)/filings/NKE_10K_2025.pdf" \
		"$(APP_DIR)/filings/Nike-Inc-2026-10K.pdf" \
		"$(APP_DIR)/filings/NKE_10Q_2027.pdf"; do \
		test -f "$$filing" || { printf 'Missing filing: %s\n' "$$filing" >&2; exit 1; }; \
	done
	@command -v "$(PDFTOTEXT)" >/dev/null 2>&1 || { printf 'pdftotext is missing. Run `make install-dependencies` first.\n' >&2; exit 1; }
	@set -eu; \
		tmp_dir=$$(mktemp -d); \
		trap 'rm -f "$$tmp_dir/NKE_10K_2024.txt" "$$tmp_dir/NKE_10K_2025.txt" "$$tmp_dir/Nike-Inc-2026-10K.txt" "$$tmp_dir/NKE_10Q_2027.pdf"; rmdir "$$tmp_dir"' EXIT HUP INT TERM; \
		"$(PDFTOTEXT)" -layout "$(APP_DIR)/filings/NKE_10K_2024.pdf" "$$tmp_dir/NKE_10K_2024.txt"; \
		"$(PDFTOTEXT)" -layout "$(APP_DIR)/filings/NKE_10K_2025.pdf" "$$tmp_dir/NKE_10K_2025.txt"; \
		"$(PDFTOTEXT)" -layout "$(APP_DIR)/filings/Nike-Inc-2026-10K.pdf" "$$tmp_dir/Nike-Inc-2026-10K.txt"; \
		cp "$(APP_DIR)/filings/NKE_10Q_2027.pdf" "$$tmp_dir/NKE_10Q_2027.pdf"; \
		$(PYTHON) extract_inputs.py --filings-dir "$$tmp_dir" \
			--source-filings-dir "$(APP_DIR)/filings" --out "$(INPUTS)"

build-workbook:
	@test -f "$(INPUTS)" || { printf 'Missing extracted inputs: %s\n' "$(INPUTS)" >&2; exit 1; }
	@test -f "$(ASSUMPTIONS)" || { printf 'Missing assumptions: %s\n' "$(ASSUMPTIONS)" >&2; exit 1; }
	@test -f "$(MARKET_DATA)" || { printf 'Missing market data: %s\n' "$(MARKET_DATA)" >&2; exit 1; }
	@test -x "$(SOFFICE_BIN)" || { printf 'LibreOffice is missing. Run `make install-dependencies` first.\n' >&2; exit 1; }
	SOFFICE="$(SOFFICE_BIN)" $(PYTHON) build_workbook.py \
		--inputs "$(INPUTS)" \
		--assumptions "$(ASSUMPTIONS)" \
		--market "$(MARKET_DATA)" \
		--out-dir "$(OUT_DIR)"

test-outputs:
	@test -x "$(SOFFICE_BIN)" || { printf 'LibreOffice is missing. Run `make install-dependencies` first.\n' >&2; exit 1; }
	APP_DIR="$(APP_DIR)" OUT_DIR="$(OUT_DIR)" SOFFICE="$(SOFFICE_BIN)" \
		$(PYTHON) -m pytest -q -p no:cacheprovider test_outputs.py