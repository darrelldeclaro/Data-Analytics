Attribute VB_Name = "removingDuplicatesOpt_1"

Sub removingDuplicatesOpt1()

'NOTE:
'remove of duplicates function will start from bottom to top
'any new data appended from the bottom will be remove as duplicates
'thus, sorting data to descending before removing of duplicates
'then sorting to ascending to reutrn data to original form


Dim selection1 As Range
Dim selection2 As Range
Dim lastrow As Long

lastrow = Cells(Cells.Rows.Count, "A").End(xlUp).row

Range("G1").End(xlDown).Select
Cells(, 23).Range(Selection, Selection.End(xlDown)) = "retain"

'adding index column
Range("B:B").Insert
Range("B2") = 1
Range("B3") = 2

Set selection1 = ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Range("B2:B3")
Set selection2 = ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Range("B2:B" & lastrow)

'autofill
selection1.autofill Destination:=selection2


'sort to descending
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort
        .SetRange Range("A2:AD" & lastrow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
'remove duplicates
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Range("A1:AD" & lastrow).removeDuplicates Columns:=1, Header:=xlYes


'sort to ascending
ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort.SortFields.add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("MANUAL_FILE_COMBINED").Sort
        .SetRange Range("A2:AD" & lastrow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'delete dummy column
Columns("B:B").Delete
Columns("ac:ac").Delete

End Sub

