Attribute VB_Name = "removingDuplicatesOpt_2"
Option Explicit

Sub removingDuplicatesOpt_2()

Dim selection1 As Range
Dim selection2 As Range
Dim selection3 As Range
Dim selection4 As Range
Dim lastrow As Long
Dim lastrowC As Long

lastrow = Cells(Rows.Count, 1).End(xlUp).row

'created 2 dummy column
Range("B:C").Insert
Range("B1") = "ord"
Range("C1") = "xty"
Range("B2") = 1
Range("B3") = 2


'column ORD
Set selection1 = ActiveWorkbook.Worksheets(1).Range("B2:B3")
Set selection2 = ActiveWorkbook.Worksheets(1).Range("B2:B" & lastrow)

selection1.autofill Destination:=selection2

Range("A1:AN1").Select


'Sort ticket number (smallest to largest) "Ascending"
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("A1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(1).Sort
    .SetRange Range("A1:AX" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With




'column XTY
Range("C2").FormulaR1C1 = "=IF(RC[-2]=R[1]C[-2],1,""x"")"

Set selection3 = ActiveWorkbook.Worksheets(1).Range("C2")
Set selection4 = ActiveWorkbook.Worksheets(1).Range("C2:C" & lastrow)

selection3.autofill Destination:=selection4

selection4.Copy
selection4.PasteSpecial xlPasteValues

'Ascending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("C1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(1).Sort
    .SetRange Range("A1:AX" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'filtering
ActiveWorkbook.Worksheets(1).Range("C2").AutoFilter field:=3, Criteria1:=1

'Delete duplicate Rows
lastrowC = Cells(Rows.Count, 3).End(xlUp).row

Range(Cells(2, "C"), Cells(lastrowC, "C")).SpecialCells(xlCellTypeVisible).EntireRow.Delete
  
'Clear Filter
ActiveSheet.ShowAllData
    
    
'Ascending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("B1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(1).Sort
    .SetRange Range("A1:AX" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'delete dummy columns
Range("B:C").Delete

End Sub



