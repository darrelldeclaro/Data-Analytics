Attribute VB_Name = "remove_selected_column"
Sub remove_selected_column()

Dim lastrow As Long

lastrow = Cells(Rows.Count, 1).End(xlUp).row

'Range("F2:F" & lastrow).ClearContents
Range("G2:G" & lastrow).ClearContents
Range("H2:H" & lastrow).ClearContents
'Range("J2:J" & lastrow).ClearContents
'Range("K2:K" & lastrow).ClearContents
Range("M2:M" & lastrow).ClearContents
Range("N2:N" & lastrow).ClearContents

End Sub


