# ChannelWise Excel Reporting Automation

A Python-based Excel reporting automation system designed to process sales and returns data and update 13 business reporting workbooks while preserving important Excel structures.

## Features

- Automates 13 Excel reports across multiple business units.
- Processes sales and returns data from Excel exports.
- Supports channel-wise, business-unit revenue, and fashion/non-fashion reporting.
- Preserves Excel formatting, conditional formatting, charts, and external links using specialized preservation modules.
- Supports configuration-based report mappings.
- Creates backups before updating report workbooks.
- Provides audit and dry-run options for validation.
- Generates processing logs for troubleshooting.

## Supported Reports

**Channel-wise reports (5)**

- Giggles
- Protecto
- Safety Sure
- Be Dapper
- Be Desi

**Additional reports (8)**

- Business Unit Daily Revenue
- Fashion Tracker
- Non-Fashion Tracker
- BU Giggles Revenue
- BU Protecto Revenue
- BU Safety Sure Revenue
- BU Be Dapper Revenue
- BU Be Desi Revenue

## Main Python Files

| File | Purpose |
|---|---|
| `run_reports.py` | Main report automation runner |
| `rename_dumps.py` | Renames incoming data exports |
| `cfpatch.py` | Preserves conditional formatting |
| `chartpatch.py` | Preserves charts and drawings |
| `external_link_patch.py` | Handles Excel external-link preservation |
| `validate_report_integrity.py` | Validates report integrity |

## Installation

1. Clone the repository.
2. Create and activate a Python virtual environment.
3. Install the project dependencies.
4. Create a local `config.json` using `config.example.json` as a starting point.
5. Supply your own authorized Excel templates and input data.
6. Run an audit and dry-run before updating workbooks.

## Example Commands

```powershell
python run_reports.py --config config.json --audit
python run_reports.py --config config.json --dry-run
python run_reports.py --config config.json
```

The example configuration is intentionally incomplete. Users must provide their own report mappings and authorized data before running the system.

## Validation

The production runner was tested against 13 report workbooks in an isolated environment.

- 13 of 13 reports saved successfully.
- 22 Excel workbooks passed ZIP integrity checks.
- 630 internal Excel relationships were checked without broken references.
- Original workbook backups were verified.

These results describe the tested workbook set and do not guarantee compatibility with every Excel template.

## Data Privacy

This repository is intended to contain source code and documentation only. Confidential company sales data, customer information, report templates, credentials, and generated outputs must remain outside version control.