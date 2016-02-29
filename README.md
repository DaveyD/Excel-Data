# Excel-Data
Assorted data on Excel and VBA

```vbnet
  Function HelloWorld()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    For i = 1 to 100
      print(i)
    Next i
    For Each cell in myRange
      print(cell.Range.Address)
    Next cell
  End Function
  
```
