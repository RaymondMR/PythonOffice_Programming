import win32com.client
wordApp = win32com.client.GetObject(Class="Word.Application")
tables=wordApp.ActiveDocument.Tables
table2=tables[1]
print(table2.Cell(1,1).Range.Text)

