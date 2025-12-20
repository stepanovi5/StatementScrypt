def process_all_statements():
    doc = XSCRIPTCONTEXT.getDocument()
    sheets = doc.Sheets

    COL_DEPOSIT = 3   # D
    COL_WITHDRAW = 4  # E
    COL_TID = 1       # B
    COL_RESULT = 7    # H

    START_ROW = 1     # со 2 строки
    MAX_ROWS = 10000  # запас, чтобы пройти весь лист

    for sheet in sheets:
        for row in range(START_ROW, MAX_ROWS):
            dep_cell = sheet.getCellByPosition(COL_DEPOSIT, row)
            wd_cell = sheet.getCellByPosition(COL_WITHDRAW, row)
            tid_cell = sheet.getCellByPosition(COL_TID, row)

            tid = tid_cell.String.strip()
            amount = None

            # строгое условие: ТОЛЬКО если есть TID
            if tid:
                if dep_cell.Value != 0:
                    amount = dep_cell.Value
                elif wd_cell.Value != 0:
                    amount = wd_cell.Value

            if amount is not None:
                sheet.getCellByPosition(COL_RESULT, row).String = f"{amount} {tid}"
            else:
                sheet.getCellByPosition(COL_RESULT, row).String = ""
