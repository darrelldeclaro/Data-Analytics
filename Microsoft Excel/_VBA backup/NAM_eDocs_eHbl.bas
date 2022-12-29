Attribute VB_Name = "NAM_eDocs_eHbl"
Sub NAM_eDocs_eHbl()

eDocs

eDocs_toCompiler

eHbl

eHbl_toCompiler


End Sub

Sub eDocs()

Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long

Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - eDOCS\Extracted Data\08 Jul- Dec 2022.xlsb")
Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - eDOCS\edocs cleaned.xlsx")

Set src = Workbooks("08 Jul- Dec 2022.xlsb").Worksheets(5)
Set dst = Workbooks("edocs cleaned.xlsx").Worksheets(1)

Workbooks("edocs cleaned.xlsx").Worksheets(1).Activate
Workbooks("edocs cleaned.xlsx").Worksheets(1).Cells.Clear


srclastrow = src.Cells(src.Rows.Count, "A").End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, "A").End(xlUp).row

src.Range("A1:AJ" & srclastrow).Copy
dst.Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = False



removingDuplicatesOpt_2

Workbooks("08 Jul- Dec 2022.xlsb").Close savechanges:=False
Workbooks("edocs cleaned.xlsx").Save



End Sub

Sub eDocs_toCompiler()

Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim lastrow As Long
Dim location As Range

Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - eDOCS\BUNK2\eDocs2 (PBI 006).xlsx")

Set src = Workbooks("edocs cleaned.xlsx").Worksheets(1)
Set dst = Workbooks("eDocs2 (PBI 006).xlsx").Worksheets(1)

tckId = src.Range("A2")
Set location = dst.Range("A:A").Find(tckId)

srclastrow = src.Cells(src.Rows.Count, "A").End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, "A").End(xlUp).row

src.Range("A2:AJ" & srclastrow).Copy
dst.Range(location.Address).PasteSpecial xlPasteValues
Application.CutCopyMode = False

lastrow = dst.Cells(dst.Rows.Count, "A").End(xlUp).row
dst.Range("G2:G" & lastrow).ClearContents
dst.Range("H2:H" & lastrow).ClearContents
dst.Range("M2:M" & lastrow).ClearContents
dst.Range("N2:N" & lastrow).ClearContents

Workbooks("edocs cleaned.xlsx").Close savechanges:=True
Workbooks("eDocs2 (PBI 006).xlsx").Close savechanges:=True

End Sub


Sub eHbl()

Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long

Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - eDOCS\Extracted Data\08 Jul - Dec 2022_ehbl.xlsb")
Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - eDOCS\edocs H cleaned.xlsx")

Set src = Workbooks("08 Jul - Dec 2022_ehbl.xlsb").Worksheets(5)
Set dst = Workbooks("edocs H cleaned.xlsx").Worksheets(1)

srclastrow = src.Cells(src.Rows.Count, "A").End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, "A").End(xlUp).row

dst.Range("A2:AL" & dstlastrow).ClearContents
src.Range("A2:AJ" & srclastrow).Copy
dst.Range("A2").PasteSpecial xlPasteValues
Application.CutCopyMode = False


removingDuplicatesOpt_2


Workbooks("08 Jul - Dec 2022_ehbl.xlsb").Close savechanges:=False
Workbooks("edocs H cleaned.xlsx").Save

End Sub


Sub eHbl_toCompiler()

Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim tckId As Range
Dim location As Range

Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\NAM - eDOCS\BUNK2\eDocs2 (PBI H03).xlsx")

Set src = Workbooks("edocs H cleaned.xlsx").Worksheets(1)
Set dst = Workbooks("eDocs2 (PBI H03).xlsx").Worksheets(1)

Set tckId = src.Range("A2")
Set location = dst.Range("A:A").Find(tckId)

srclastrow = src.Cells(src.Rows.Count, "A").End(xlUp).row
dstlastrow = dst.Cells(dst.Rows.Count, "A").End(xlUp).row

src.Range("A2:AJ" & srclastrow).Copy
dst.Range(location.Address).PasteSpecial xlPasteValues
Application.CutCopyMode = False


dst.Range("G2:G" & dstlastrow).ClearContents
dst.Range("H2:H" & dstlastrow).ClearContents
dst.Range("M2:M" & dstlastrow).ClearContents
dst.Range("N2:N" & dstlastrow).ClearContents


Workbooks("edocs H cleaned.xlsx").Close savechanges:=True
Workbooks("eDocs2 (PBI H03).xlsx").Close savechanges:=True


End Sub



Sub removingDuplicatesOpt_2()

Dim selection1 As Range
Dim selection2 As Range
Dim selection3 As Range
Dim selection4 As Range
Dim lastrow As Long
Dim lastrowC As Long

lastrow = Cells(Rows.Count, "A").End(xlUp).row


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
    .SetRange Range("A1:AL" & lastrow)
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
    .SetRange Range("A1:AL" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'filtering
ActiveWorkbook.Worksheets(1).Range("C2").AutoFilter field:=3, Criteria1:=1



'Delete duplicate Rows
If Range("C2") = 1 Then
    
    lastrowC = Cells(Rows.Count, 3).End(xlUp).row
    
    Range(Cells(2, "C"), Cells(lastrowC, "C")).SpecialCells(xlCellTypeVisible).EntireRow.Delete
      
    'Clear Filter
    ActiveSheet.ShowAllData
ElseIf Range("C2") = "x" Then
    ActiveSheet.ShowAllData
End If

    
    
'Ascending
ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
ActiveWorkbook.Worksheets(1).Sort.SortFields.add Key:=Range("B1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.Worksheets(1).Sort
    .SetRange Range("A1:AL" & lastrow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'delete dummy columns
Range("B:C").Delete

End Sub





