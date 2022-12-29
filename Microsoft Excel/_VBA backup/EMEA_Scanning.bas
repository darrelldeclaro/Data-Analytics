Attribute VB_Name = "EMEA_Scanning"
Option Explicit
Public Sub EMEA_Scanning()

'***** Initial Data Cleaning and Aggregation Process ****'
consolidation

'***** Second Phase Data Cleaning from OutLook and Aggregation Process ****'
cleaning_DPI

'***** Third Phase - Clean data from scan template to scan raw clean_RTQ prep ****'
complier_1

'***** forth Phase - Clean data from dpi autofill to scan raw clean_dpi autofill  ****'
complier_2

'***** final Phase - Clean data from complier to PBI ****'
complier_PBI

End Sub

Sub consolidation()

Dim wsSrc1 As Worksheet
Dim wsDst1 As Worksheet
Dim selection1 As Range
Dim selection2 As Range
Dim selection3 As Range
Dim selection4 As Range
Dim srclastrow1 As Long
Dim dstlastrow1 As Long
Dim lastrow As Long
Dim lastrowV As Long

Workbooks.Open "C:\Users\darrell.declaro\Downloads\merged_scanning.xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Scanning\Scan Team Template.xlsx"

'*****PATH RESULT TO TEMPLATE*****'
Set wsSrc1 = Workbooks("merged_scanning.xlsx").Worksheets(1)
Set wsDst1 = Workbooks("Scan Team Template.xlsx").Worksheets(1)

'copy until last row of the source'
srclastrow1 = wsSrc1.Cells(wsSrc1.Rows.Count, "A").End(xlUp).row

'find the first blank row of the destination'
dstlastrow1 = wsDst1.Cells(wsDst1.Rows.Count, "A").End(xlUp).row

'Removing existing data to destination
wsDst1.Range("A2:T" & dstlastrow1).ClearContents
wsDst1.Range("V5:AO" & dstlastrow1).ClearContents

'copy and paste'
wsSrc1.Range("A2:T" & srclastrow1).Copy
    wsDst1.Range("A2").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    
'applying calculated fields in yellow colored headers
lastrow = Cells(Rows.Count, "A").End(xlUp).row

Set selection1 = wsDst1.Range("V2")
Set selection2 = wsDst1.Range("V2:V" & lastrow)

selection1.autofill Destination:=selection2

Set selection3 = wsDst1.Range("X2:AO2")
Set selection4 = wsDst1.Range("X2:AO" & lastrow)

selection3.autofill Destination:=selection4


'filtering
wsDst1.Range("V2").AutoFilter field:=22, Criteria1:="OTHERS"

lastrowV = Cells(Rows.Count, "V").End(xlUp).row

        'Delete duplicate Rows
        lastrowV = Cells(Rows.Count, 3).End(xlUp).row
        
        wsDst1.Range(Cells(2, "V"), Cells(lastrowV, "V")).SpecialCells(xlCellTypeVisible).EntireRow.Delete
          
        'Clear Filter
        ActiveSheet.ShowAllData


    
    

Workbooks("merged_scanning.xlsx").Close
Workbooks("Scan Team Template.xlsx").Save


End Sub



Sub cleaningRTQ()

Dim wsSrc1 As Worksheet
Dim wsDst1 As Worksheet
Dim selection1 As Range
Dim selection2 As Range
Dim srclastrow1 As Long
Dim dstlastrow1 As Long
Dim lastrow As Long
Dim lastrowV As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Scanning\Scan Team Raw - Cleanup.xlsx"


'*****PATH RESULT TO TEMPLATE*****'
Set wsSrc1 = Workbooks("Scan Team Template.xlsx").Worksheets("Consolidated")
Set wsDst1 = Workbooks("Scan Team Raw - Cleanup.xlsx").Worksheets(1)

'copy until last row of the source'
srclastrow1 = wsSrc1.Cells(wsSrc1.Rows.Count, "A").End(xlUp).row

'find the first blank row of the destination'
dstlastrow1 = wsDst1.Cells(wsDst1.Rows.Count, "A").End(xlUp).Offset(1).row


'copy and paste'
wsSrc1.Range("X2:AO" & srclastrow1).Copy
    wsDst1.Range("A" & dstlastrow1).PasteSpecial xlPasteValues
    Application.CutCopyMode = False



removingDuplicatesOpt_2

    
'MsgBox ("Result to PBI Complete")'
Workbooks("Scan Team Template.xlsx").Close savechanges:=True
Workbooks("Scan Team Raw - Cleanup.xlsx").Save

'MsgBox ("Cleaning RTQ Done!")

End Sub


Sub cleaning_DPI_raw()

Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim lastrow As Long


'Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Scanning\Extracted Raw Data\DPI_AutoFiling_SEA_cleaned.xlsx"
'Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Scanning\Extracted Raw Data\DPI_AutoFiling_SEA_CLICK HERE TO OVERWRITE.csv"

Set src = Workbooks("DPI_AutoFiling_SEA_CLICK HERE TO OVERWRITE.csv").Worksheets(1)
'Set dst = Workbooks("DPI_AutoFiling_SEA_cleaned.xlsx").Worksheets(1)

srclastrow = src.Cells(src.Rows.Count, 1).End(xlUp).row
'dstlastrow = dst.Cells(dst.Rows.Count, 1).End(xlUp).row

'text to column
src.Range("A1:A" & srclastrow).TextToColumns _
    Destination:=Columns("A:R"), _
    DataType:=xlDelimited, _
   TextQualifier:=xlDoubleQuote, _
   ConsecutiveDelimiter:=True, _
   Tab:=True, _
   Semicolon:=True, _
   Comma:=False, _
   Space:=False, _
   Other:=False, _
   FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1))

'filtering
ActiveWorkbook.Worksheets(1).Range("C2").AutoFilter field:=9, Criteria1:=""

'Delete duplicate Rows
lastrow = Cells(Rows.Count, 3).End(xlUp).row

src.Range(Cells(2, "C"), Cells(lastrow, "C")).SpecialCells(xlCellTypeVisible).EntireRow.Delete
  
'Clear Filter
ActiveSheet.ShowAllData


End Sub






Sub cleaning_DPI()

Dim wsSrc1 As Worksheet
Dim wsDst1 As Worksheet
Dim selection1 As Range
Dim selection2 As Range
Dim srclastrow1 As Long
Dim dstlastrow1 As Long
Dim lastrow As Long
Dim lastrowV As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Scanning\Extracted Raw Data\DPI_AutoFiling_SEA_cleaned.xlsx"
Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Scanning\DPI_AutoFiling_SEA_Template.xlsx"

'*****PATH RESULT TO TEMPLATE*****'
Set wsSrc1 = Workbooks("DPI_AutoFiling_SEA_cleaned.xlsx").Worksheets(1)
Set wsDst1 = Workbooks("DPI_AutoFiling_SEA_Template.xlsx").Worksheets(1)

'copy until last row of the source'
srclastrow1 = wsSrc1.Cells(wsSrc1.Rows.Count, "A").End(xlUp).row

'find the first blank row of the destination'
dstlastrow1 = wsDst1.Cells(wsDst1.Rows.Count, "A").End(xlUp).row

'Removing existing data to destination
wsDst1.Range("A2:R" & dstlastrow1).ClearContents

'copy and paste'
wsSrc1.Range("A2:R" & srclastrow1).Copy
    wsDst1.Range("A2").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    
'applying calculated fields in yellow colored headers
lastrow = Cells(Rows.Count, "A").End(xlUp).row

wsDst1.Range("T5:AX" & lastrow).ClearContents

Set selection1 = wsDst1.Range("T2:AX2")
Set selection2 = wsDst1.Range("T2:AX" & lastrow)

selection1.autofill Destination:=selection2

Workbooks("DPI_AutoFiling_SEA_cleaned.xlsx").Close

Workbooks("DPI_AutoFiling_SEA_Template.xlsx").Save


End Sub

Sub complier_1()
Dim wsSrc As Worksheet
Dim wsDst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Scanning\Scan Team Raw - Cleanup.xlsx"

'*****PATH RESULT TO TEMPLATE*****'
Set wsSrc = Workbooks("Scan Team Template.xlsx").Worksheets("Consolidated")
Set wsDst = Workbooks("Scan Team Raw - Cleanup.xlsx").Worksheets(1)


'wsSrc2 TO wsDst4'
srclastrow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row

'find the first blank row of the destination'
dstlastrow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).Offset(1).row

'copy and paste'
wsSrc.Range("X2:AO" & srclastrow).Copy
    wsDst.Range("A" & dstlastrow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    

'MsgBox "Checking for custom field {EMEA Country} and  custom field {no. of transaction}"


Workbooks("Scan Team Template.xlsx").Close savechanges:=True

Workbooks("Scan Team Raw - Cleanup.xlsx").Save

End Sub

Sub complier_2()
Dim wsSrc As Worksheet
Dim wsDst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long



'*****PATH RESULT TO TEMPLATE*****'
Set wsSrc = Workbooks("DPI_AutoFiling_SEA_Template.xlsx").Worksheets(1)
Set wsDst = Workbooks("Scan Team Raw - Cleanup.xlsx").Worksheets(2)


'wsSrc2 TO wsDst4'
srclastrow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).row

'find the first blank row of the destination'
dstlastrow = wsDst.Cells(wsDst.Rows.Count, 1).End(xlUp).Offset(1).row

'copy and paste'
wsSrc.Range("AG2:AX" & srclastrow).Copy
    wsDst.Range("A" & dstlastrow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    

'MsgBox "Checking for custom field {EMEA Country} and  custom field {no. of transaction}"
'checkingForErrors


Workbooks("DPI_AutoFiling_SEA_Template.xlsx").Close savechanges:=True

'MsgBox ("Consolidation Done!")
End Sub

Sub complier_PBI()
Dim wsSrc1 As Worksheet
Dim wsSrc2 As Worksheet
Dim wsDst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim lastrow As Long

Workbooks.Open "C:\Users\darrell.declaro\Documents\Automation ver1.0\EMEA - Scanning\PBI Data Source\Scan Team Raw (PBI) - 2022.xlsx"

'*****PATH RESULT TO TEMPLATE*****'
Set wsSrc1 = Workbooks("Scan Team Raw - Cleanup.xlsx").Worksheets("RTQ Prep")
Set wsSrc2 = Workbooks("Scan Team Raw - Cleanup.xlsx").Worksheets("DPI Autofiling Prep")
Set wsDst = Workbooks("Scan Team Raw (PBI) - 2022.xlsx").Worksheets(1)



'wsSrc1(RTW prep) TO wsDst(PBI)'
srclastrow = wsSrc1.Cells(wsSrc1.Rows.Count, "A").End(xlUp).row

dstlastrow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).row

wsDst.Range("A2:R" & dstlastrow).ClearContents

wsSrc1.Range("A2:R" & srclastrow).Copy
    wsDst.Range("A2").PasteSpecial xlPasteValues
    Application.CutCopyMode = False


'wsSrc2(DPI autofilling) TO wsDst(PBI)'
srclastrow = wsSrc2.Cells(wsSrc2.Rows.Count, "A").End(xlUp).row

lastrow = wsDst.Cells(wsDst.Rows.Count, "A").End(xlUp).Offset(1).row

wsSrc2.Range("A2:R" & srclastrow).Copy
    wsDst.Range("A" & lastrow).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    

remove_selected_column

cleaning_PBI

'closing workbooks
Workbooks("Scan Team Raw - Cleanup.xlsx").Close savechanges:=True
Workbooks("Scan Team Raw (PBI) - 2022.xlsx").Save

Application.DisplayAlerts = False
ThisWorkbook.SaveAs "C:\Users\darrell.declaro\Documents\Automation ver1.0\ASP - Scanning\PBI Data Source", 51
ThisWorkbook.SaveAs "C:\Users\darrell.declaro\Documents\Automation ver1.0\SCA - Scanning\PBI Data Source", 51
Application.DisplayAlerts = True


MsgBox ("Check Columns information, Column E[OwnerName] for error, Column P and Q for incorrect input")
End Sub


Sub cleaning_PBI()

Dim ws As Worksheet
Dim lastrow_O As Long
Dim lastrow_P As Long
Dim lastrow_Q As Long

Set ws = Workbooks("Scan Team Raw (PBI) - 2022.xlsx").Worksheets(1)

ws.Range("A:A").AutoFilter field:=5, Criteria1:="stephan.hassenpflug", Operator:=xlFilterValues

ws.Range("A:A").AutoFilter field:=15, Criteria1:="*1900*", Operator:=xlFilterValues

ws.Range("O2").Select
ActiveCell.Offset(1).Select
Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1).Select
Loop

lastrow_O = Cells(Rows.Count, 1).End(xlUp).row
Range(Cells(2, "O"), Cells(lastrow_O, "O")).SpecialCells(xlCellTypeVisible).EntireRow.Delete

ws.ShowAllData

End Sub


Sub checkingForErrors()

Dim wsDst2 As Worksheet
Dim lastrow As Long
Dim colm As Range
Dim cell As Range

Set wsDst2 = Workbooks("Scan Team Raw - Cleanup.xlsx").Worksheets("RTQ Prep")
lastrow = wsDst2.Cells(wsDst2.Rows.Count, "P").End(xlUp).row
Set colm = wsDst2.Range("P2:P" & lastrow)


For Each cell In colm
    If Len(cell) > 2 Then
        cell = Left(cell, 2)
    ElseIf cell <> "" Then
        cell = ""
    End If
Next cell


Set colm = wsDst2.Range("Q2:Q" & lastrow)

For Each cell In colm
    If cell > 500 Then
        cell = 1
    ElseIf cell <> "" Then
        cell = ""
    End If
Next cell

End Sub



Sub removingDuplicatesOpt_2()

Dim src As Worksheet
Dim selection1 As Range
Dim selection2 As Range
Dim selection3 As Range
Dim selection4 As Range
Dim lastrow As Long
Dim lastrowC As Long

Set src = Workbooks("Scan Team Raw - Cleanup.xlsx").Worksheets(1)


lastrow = Cells(Rows.Count, 1).End(xlUp).row

'created 2 dummy column
src.Range("B:C").Insert
src.Range("B1") = "ord"
src.Range("C1") = "xty"
src.Range("B2") = 1
src.Range("B3") = 2


'column ORD
Set selection1 = src.Range("B2:B3")
Set selection2 = src.Range("B2:B" & lastrow)

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

If Range("C2") = 1 Then
        'Delete duplicate Rows
        lastrowC = Cells(Rows.Count, 3).End(xlUp).row
        
        Range(Cells(2, "C"), Cells(lastrowC, "C")).SpecialCells(xlCellTypeVisible).EntireRow.Delete
          
        'Clear Filter
        ActiveSheet.ShowAllData
    Else
        ActiveSheet.ShowAllData
End If
    
    
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



Sub remove_selected_column()

Dim lastrow As Long

lastrow = Cells(Rows.Count, 1).End(xlUp).row

Range("F2:F" & lastrow).ClearContents
Range("G2:G" & lastrow).ClearContents
Range("H2:H" & lastrow).ClearContents
Range("J2:J" & lastrow).ClearContents
Range("K2:K" & lastrow).ClearContents
Range("M2:M" & lastrow).ClearContents

End Sub






