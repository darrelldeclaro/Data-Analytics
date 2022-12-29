Attribute VB_Name = "NAM_UD"
Public Sub NAM_Upload_Documentation()

Dim wsSrc As Worksheet
Dim wsDst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Upload Documentation\Raw Data\Resolved Tickets.xlsb"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Upload Documentation\FOR_CLEANING.xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - Upload Documentation\NAM Doc Upload Raw 2 (PBI).xlsx"


'*****PATH RESULT TO TEMPLATE*****'
Set wsSrc = Workbooks("Resolved Tickets.xlsb").Worksheets(6)
Set wsDst = Workbooks("FOR_CLEANING.xlsx").Worksheets(1)

'copy until last row of the source'
srclastrow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).row
dstlastrow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Offset(1).row

wsDst.Range("A2:AL" & dstlastrow).ClearContents

'copy and paste to destination'
wsSrc.Range("A2:AL" & srclastrow).Copy
    wsDst.Range("A2").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
'MsgBox ("Result to PBI Complete")'
Workbooks("Resolved Tickets.xlsb").Close

'MsgBox "Removing duplicates and highlighted fonts"

removingDuplicatesOpt_2

dataPreparation2

'Workbooks("SAP MANUAL_FILE_COMBINED RAW (PBI).xlsx").Save '

MsgBox ("column AJ for non-numeric entries and ERRORS.")


End Sub

Sub dataPreparation2()
Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim location As Range


'locating the cell value from source to destination, then paste
Set src = Workbooks("FOR_CLEANING.xlsx").Worksheets(1)
Set dst = Workbooks("NAM Doc Upload Raw 2 (PBI).xlsx").Worksheets(1)
Set tckId = src.Range("A2")
Set location = dst.Range("A:A").Find(tckId)

srclastrow = src.Cells(src.Rows.Count, "A").End(xlUp).row

src.Range("A2:AL" & srclastrow).Copy
dst.Range(location.Address).PasteSpecial xlPasteValues
Application.CutCopyMode = False

remove_selected_column

Workbooks("FOR_CLEANING.xlsx").Close savechanges:=True
Workbooks("NAM Doc Upload Raw 2 (PBI).xlsx").Save


Application.ScreenUpdating = False
Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
Application.ScreenUpdating = True


'Workbooks("BAL_INV Volume (PBI).xlsx").Close savechanges:=True
End Sub

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

Range("A1:AL1").Select


'Sort ticket number (smallest to largest) "Ascending"
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("A1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(1).Sort
    .SetRange Range("A1:AN" & lastrow)
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
    .SetRange Range("A1:AN" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'filtering
ActiveWorkbook.Worksheets(1).Range("C2").AutoFilter field:=3, Criteria1:=1



'Delete duplicate Rows
If Range("c2") = 1 Then
    lastrowC = Cells(Rows.Count, 3).End(xlUp).row
    Range(Cells(2, "C"), Cells(lastrowC, "C")).SpecialCells(xlCellTypeVisible).EntireRow.Delete
'Clear Filter
    ActiveSheet.ShowAllData
ElseIf c2 = x Then
    ActiveSheet.ShowAllData
End If

    
    
'Ascending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("B1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(1).Sort
    .SetRange Range("A1:AN" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'delete dummy columns
Range("B:C").Delete

End Sub


Sub remove_selected_column()

Dim lastrow As Long

lastrow = Cells(Rows.Count, 1).End(xlUp).row

Range("G2:G" & lastrow).ClearContents
Range("H2:H" & lastrow).ClearContents
Range("M2:M" & lastrow).ClearContents

End Sub






