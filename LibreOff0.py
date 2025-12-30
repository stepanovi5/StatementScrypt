import tabula
import pandas as pd
from pathlib import Path

pdf_folder = Path("pdf_statements")
out_folder = Path("xlsx_statements")
out_folder.mkdir(exist_ok=True)

for pdf in pdf_folder.glob("*.pdf"):
    dfs = tabula.read_pdf(pdf, pages="all", lattice=True)

    if not dfs:
        continue

    df = pd.concat(dfs, ignore_index=True)
    out_file = out_folder / (pdf.stem + ".xlsx")
    df.to_excel(out_file, index=False)
