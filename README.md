# Lesson Plan Generator

A small **CLI** tool (Python) that reads a spreadsheet with the schedule of a technical course and generates a **Lesson Plan PDF** for a given week using a Word (`.docx`) template you provide.

## Features
- Reads **CSV** or **Excel** (`.xlsx`) files with `pandas`.
- Filters rows by the `Semana` (week) column.
- Replaces placeholders in the template (`{{ColumnName}}`) with data from the sheet.
- Produces a PDF via `docx2pdf` (macOS/Windows).
- Simple command‑line interface.

## Installation
```bash
# Inside the project folder
python -m venv .venv
source .venv/bin/activate  # macOS/Linux
# Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage
```bash
python -m lesson_plan_generator.main \
    --spreadsheet /caminho/para/planilha.xlsx \
    --week 3 \
    --output lesson_plan_semana_3.pdf
```
- `--spreadsheet` – path to the CSV/Excel file containing the schedule.
- `--week` – week identifier (must match exactly the value in the `Semana` column).
- `--output` – optional; defaults to `lesson_plan_week_<week>.pdf`.

The generated PDF will be saved at the location you specify.

## Template format
The template (`template/template.docx`) should contain placeholders in the form `{{ColumnName}}`. Example:
```
Título da aula: {{Título da aula}}
Objetivo da aula: {{Objetivo da aula}}
```
When the script runs, each placeholder is replaced with the corresponding value from the spreadsheet for the selected week.

## Limitations
- `docx2pdf` works only on macOS or Windows; on Linux you would need an alternative conversion method.
- The script expects a column that contains the week identifier (e.g., `Semana`). If your sheet uses a different name, adjust the code in `plan_generator.py` accordingly.

## License
MIT – feel free to adapt and extend.
