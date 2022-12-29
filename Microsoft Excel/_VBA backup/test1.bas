Attribute VB_Name = "test1"
Sub test()


Dim src As Worksheet
Dim dst As Worksheet
Dim srclastrow As Long
Dim dstlastrow As Long
Dim selection1 As Range
Dim selection2 As Range

'Workbooks.Open "C:\Users\darrell.declaro\Documents\Analytics & Reporting\SCA - KNOPS\SCA KNOPS (PBI).xlsx"

Set src = Workbooks("SCA KNOPS - TEMPLATE.xlsx").Worksheets(1)
'Set dst = Workbooks("SCA KNOPS (PBI).xlsx").Worksheets(1)

srclastrow = src.Cells(src.Rows.Count, 1).End(xlUp).row
'dstlastrow = dst.Cells(dst.Rows.Count, 1).End(xlUp).row



'Change/update Column R "SSC SJO" based on the countries
src.Range("A2:AD" & srclastrow).AutoFilter field:=28, Criteria1:=Array("for clarification", "no data"), Operator:=xlFilterValues

src.Range("A2:AD" & srclastrow).AutoFilter field:=2, Criteria1:="*|*", Operator:=xlFilterValues

src.Activate
src.Range("R2").Select
ActiveCell.Offset(1, 0).Select
Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1, 0).Select
Loop

src.Range(Selection, ActiveCell.SpecialCells(xlLastCell).Offset(, -12)).FormulaR1C1 = "=INDEX(Reference!R2C9:R47C9,MATCH(LEFT(RC[-16],FIND(""|"",RC[-16])-2),Reference!R2C11:R47C11,0))"

ActiveSheet.ShowAllData


'filter for region SCA
src.Range("A2:AD" & srclastrow).AutoFilter field:=28, Criteria1:="SCA", Operator:=xlFilterValues





End Sub

