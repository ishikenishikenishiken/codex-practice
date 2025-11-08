# Excel Analyzer

This command-line tool summarizes numeric data in Excel workbooks. It reads one
or more sheets and prints statistics for every numeric column, along with an
optional correlation matrix. Existing HTML pages in the repository are left
untouched as requested.

## Requirements

- Python 3.9+
- [pandas](https://pandas.pydata.org/)
- A compatible Excel reader engine (``openpyxl`` is installed automatically by
  pandas when using ``pip install pandas``).

You can install the dependency with ``pip``:

```bash
pip install pandas
```

## Usage

```bash
python -m excel_analyzer.analyze_excel <path-to-workbook.xlsx>
```

Optional arguments:

- ``--sheet`` – analyze only the specified sheet name instead of every sheet.

The script prints summaries similar to the following:

```
Sheet: Sales Q1
Numeric column summary:
  revenue
    Count   : 12
    Missing : 0
    Mean    : 230.7500
    Std dev : 15.2300
    Min     : 205.0000
    Max     : 260.0000
Correlation matrix:
                revenue      cost
        revenue   1.0000   0.8123
           cost   0.8123   1.0000
```

This output helps you quickly understand the distribution and relationships in
the numeric columns of the workbook.
