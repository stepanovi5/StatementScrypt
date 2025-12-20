import re

def process_two_columns():
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.CurrentController.ActiveSheet

    COL_AMOUNT = 0   # A
    COL_TEXT = 1     # B
    COL_RESULT = 2   # C

    START_ROW = 1
    MAX_ROWS = 10000

    # шаблон TID: берём первое "похожее" значение
    tid_pattern = re.compile(r'[A-Za-z0-9][A-Za-z0-9\-/_]{5,}')

    for row in range(START_ROW, MAX_ROWS):
        amount_cell = sheet.getCellByPosition(COL_AMOUNT, row)
        text_cell = sheet.getCellByPosition(COL_TEXT, row)

        amount = amount_cell.Value
        text = text_cell.String.strip()

        if amount == 0 or not text:
            sheet.getCellByPosition(COL_RESULT, row).String = ""
            continue

        match = tid_pattern.search(text)

        if match:
            tid = match.group(0)
            sheet.getCellByPosition(COL_RESULT, row).String = f"{amount} {tid}"
        else:
            sheet.getCellByPosition(COL_RESULT, row).String = ""
