# python EnumerateTables.py
import win32com.client

# Conectarse a una instancia existente de Word
wordApp = win32com.client.GetObject(Class="Word.Application")

try:
    doc = wordApp.ActiveDocument
except Exception:
    raise RuntimeError("No hay un documento activo en Word.")

tables = doc.Tables
if tables.Count == 0:
    print("No se encontraron tablas en el documento activo.")
else:
    # Recorre todas las tablas
    for idx in range(1, tables.Count + 1):
        tbl = tables.Item(idx)
        try:
            # Obtiene el texto de la celda (1,1)
            raw_text = tbl.Cell(1, 1).Range.Text
            # En Word, el texto de la celda termina con marcadores de fin de celda: '\r\x07'
            clean_text = raw_text.rstrip("\r\x07")
        except Exception as e:
            clean_text = f"[Error al leer la celda (1,1): {e}]"

        # Imprime en el formato: tableNumber, celda(1,1).Value
        print(f"{idx}, {clean_text}")
