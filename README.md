# Custom Excel Builder

Upload up to 5 Excel/CSV files (`.xlsx`, `.xlsb`, `.xlsm`, `.xls`, `.csv`, `.tsv`), tick the sheets and headers you need, and download one clean `.xlsx`.

Everything runs inside your browser (Python via [Pyodide](https://pyodide.org)). Files are read in place and **never uploaded anywhere**.

- First load takes about 20 seconds while the engine starts.
- Big files work but are slow: roughly 20,000 rows per 45 seconds. Keep the tab open.
- Excel's limit of 1,048,575 rows per sheet applies.

Files: `index.html` (UI), `worker.js` (runs Python off the main thread), `engine.py` (reads/writes the workbooks).
