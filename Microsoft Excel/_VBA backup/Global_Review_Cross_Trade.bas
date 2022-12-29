Attribute VB_Name = "Global_Review_Cross_Trade"
Public Sub Global_Review_Cross_Trade()

cleaning_source1
cleaning_source2

End Sub


Sub cleaning_source1()

Dim src1 As Worksheet
Dim src2 As Worksheet
Dim dst1 As Worksheet
Dim dst2 As Worksheet
Dim src1r As Long
Dim dst1r As Long
Dim src2r As Long
Dim dst2r As Long


Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\GLOBAL - Review Cross Trade\Extracted Raw Data\Review_Cross_Trade_Report 2022.xlsx")
Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\GLOBAL - Review Cross Trade\GLOBAL Review Cross Trade - TEMPLATE.xlsx")
Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\GLOBAL - Review Cross Trade\BUNK\Review Cross Trade (PBI) - 01.xlsx")

Set src1 = Workbooks("Review_Cross_Trade_Report 2022.xlsx").Worksheets(1)
Set dst1 = Workbooks("GLOBAL Review Cross Trade - TEMPLATE.xlsx").Worksheets(1)
Set src2 = Workbooks("GLOBAL Review Cross Trade - TEMPLATE.xlsx").Worksheets(1)
Set dst2 = Workbooks("Review Cross Trade (PBI) - 01.xlsx").Worksheets(1)


'src1.Range("B2").Select
src1r = src1.Cells(src1.Rows.Count, 2).End(xlUp).row
dst1r = dst1.Cells(dst1.Rows.Count, 1).End(xlUp).row

dst1.Range("A2:O" & dst1r).ClearContents
dst1.Range("P3:BF" & dst1r).ClearContents

'raw to template
src1.Range("B3:P" & src1r).Copy
dst1.Range("A2").PasteSpecial xlPasteValues
Application.CutCopyMode = False

dst1.Range("P2:BF2").Copy
dst1.Range("P2:BF" & dst1r).PasteSpecial xlPasteFormulas
Application.CutCopyMode = False

Workbooks("Review_Cross_Trade_Report 2022.xlsx").Close savechanges:=False

   

'template to pbi
src2r = src2.Cells(src2.Rows.Count, 1).End(xlUp).row
dst2r = dst2.Cells(dst2.Rows.Count, 1).End(xlUp).row

dst2.Range("A2:BF" & dst2r).ClearContents

src2.Range("A2:BF" & src2r).Copy
dst2.Range("A2").PasteSpecial xlPasteValues
Application.CutCopyMode = False

With dst2.Range("C2:C" & dst2r)
    .NumberFormat = "general"
    .Value = .Value
End With

Workbooks("GLOBAL Review Cross Trade - TEMPLATE.xlsx").Close savechanges:=True
Workbooks("Review Cross Trade (PBI) - 01.xlsx").Save

End Sub



Sub cleaning_source2()

Dim src1 As Worksheet
Dim src2 As Worksheet
Dim dst1 As Worksheet
Dim dst2 As Worksheet
Dim src1r As Long
Dim dst1r As Long
Dim src2r As Long
Dim dst2r As Long
Dim lastrow As Long

Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\GLOBAL - Review Cross Trade\Extracted Raw Data\Review_Cross_Trade_Report_(SSC_HER_Productivity).xlsx")
Workbooks.Open ("C:\Users\darrell.declaro\Documents\Automation ver1.0\GLOBAL - Review Cross Trade\GLOBAL Review Cross Trade - TEMPLATE (RT data).xlsx")


Set src1 = Workbooks("Review_Cross_Trade_Report_(SSC_HER_Productivity).xlsx").Worksheets(1)
Set dst1 = Workbooks("GLOBAL Review Cross Trade - TEMPLATE (RT data).xlsx").Worksheets(1)
Set src2 = Workbooks("GLOBAL Review Cross Trade - TEMPLATE (RT data).xlsx").Worksheets(1)
Set dst2 = Workbooks("Review Cross Trade (PBI) - 01.xlsx").Worksheets(1)

'src1.Range("B4").Select
src1r = src1.Cells(src1.Rows.Count, 2).End(xlUp).row
dst1r = dst1.Cells(dst1.Rows.Count, 1).End(xlUp).row

dst1.Range("A2:J" & dst1r).ClearContents
dst1.Range("L3:BQ" & dst1r).ClearContents

'raw to template
src1.Range("B5:K" & src1r).Copy
dst1.Range("A2").PasteSpecial xlPasteValues
Application.CutCopyMode = False

lastrow = dst1.Cells(dst1.Rows.Count, 1).End(xlUp).row
dst1.Range("L2:BQ2").Copy
dst1.Range("L2:BQ" & lastrow).PasteSpecial xlPasteFormulas
Application.CutCopyMode = False

Workbooks("Review_Cross_Trade_Report_(SSC_HER_Productivity).xlsx").Close savechanges:=False


'removing entire row with "test" or "tst"
dst1.Range("A2:BQ" & lastrow).AutoFilter field:=2, Criteria1:="*test*", Operator:=xlFilterValues

lastrow1 = Cells(Rows.Count, 2).End(xlUp).row
Range(Cells(2, "B"), Cells(lastrow1, "B")).SpecialCells(xlCellTypeVisible).EntireRow.Delete

dst1.ShowAllData



'renaming "Review Cross Trade - 20 - 20" to "Review Cross Trade - Number TN's Checked 20 - Email Follow Up Set 20"
dst1.Range("B:B").Replace What:="Review Cross Trade - 20 - 20", Replacement:="Review Cross Trade - Number TN's Checked 20 - Email Follow Up Set 20"

'checking column BQ for errorneous inputs
For Each cell In dst1.Range("BQ2:BQ" & lastrow)
    If cell.Value > 100 Then
        cell.Value = 1
    End If
Next



'template to pbi
src2r = src2.Cells(src2.Rows.Count, 1).End(xlUp).row
dst2r = dst2.Cells(dst2.Rows.Count, 1).End(xlUp).Offset(1).row

src2.Range("L2:BQ" & src2r).Copy
dst2.Range("A" & dst2r).PasteSpecial xlPasteValues
Application.CutCopyMode = False

Workbooks("GLOBAL Review Cross Trade - TEMPLATE (RT data).xlsx").Close savechanges:=True
Workbooks("Review Cross Trade (PBI) - 01.xlsx").Close savechanges:=True

End Sub
