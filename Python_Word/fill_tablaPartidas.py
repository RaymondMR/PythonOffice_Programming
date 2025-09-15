# python fill_tablaPartidas.py
# Rellena la "tabla 2" del documento Word activo con datos de la hoja "FINIQUITO" del primer libro abierto en Excel.
# La tabla 2 se asume que tiene 4 columnas.

import win32com.client

# === Instancias de Excel y Word (según tu contexto) ===
excelApp = win32com.client.GetObject(Class="Excel.Application")
workbook = excelApp.Workbooks[1]
sheet = workbook.Worksheets("FINIQUITO")

wordApp = win32com.client.GetObject(Class="Word.Application")
tables = wordApp.ActiveDocument.Tables
table2 = tables[1]  # "tabla 2" en tu texto, pero índice [1] según tu variable

# === Parámetros de trabajo ===
excel_start_row = 304
excel_end_row   = 314  # inclusive
excel_cols = {
    1: 4,   # Columna 1 de Word <- Col D (4) en Excel
    2: 6,   # Columna 2 de Word <- Col F (6)
    3: 8,   # Columna 3 de Word <- Col H (8)
    4: 10,  # Columna 4 de Word <- Col J (10)
}

word_header_rows = 1               # La fila 1 de Word se asume encabezado
rows_to_fill = excel_end_row - excel_start_row + 1
required_rows_in_table = word_header_rows + rows_to_fill

# === Asegurar que la tabla tenga suficientes filas ===
while table2.Rows.Count < required_rows_in_table:
    table2.Rows.Add()

def to_text(v):
    """Convierte valores de Excel a texto limpio para Word."""
    if v is None:
        return ""
    # Si es un float con punto decimal innecesario, lo limpiamos suavemente
    try:
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
    except Exception:
        pass
    return str(v)

def to_accounting(v):
    """Formatea el valor en formato contable: $ como prefijo y miles con comas."""
    if v is None:
        return ""
    try:
        num = float(v)
        return f"${num:,.2f}"
    except Exception:
        # Si no es numérico, retorna tal cual como texto
        return to_text(v)

# === Volcado de datos Excel -> Word ===
word_row = word_header_rows + 1  # Empezar a rellenar desde fila 2 en Word
for excel_row in range(excel_start_row, excel_end_row + 1):
    # Leer valores desde Excel
    d_val = sheet.Cells(excel_row, excel_cols[1]).Value  # D
    f_val = sheet.Cells(excel_row, excel_cols[2]).Value  # F
    h_val = sheet.Cells(excel_row, excel_cols[3]).Value  # H
    j_val = sheet.Cells(excel_row, excel_cols[4]).Value  # J

    # Escribir en Word (col 1..4)
    table2.Cell(word_row, 1).Range.Text = to_accounting(d_val)
    table2.Cell(word_row, 2).Range.Text = to_accounting(f_val)
    table2.Cell(word_row, 3).Range.Text = to_accounting(h_val)
    table2.Cell(word_row, 4).Range.Text = to_accounting(j_val)

    word_row += 1

# (Opcional) Ajustes visuales mínimos
# table2.AllowAutoFit = True
# table2.AutoFitBehavior(1)  # wdAutoFitContent = 1
print("Tabla rellenada correctamente de la fila 2 a la fila", required_rows_in_table)
