Attribute VB_Name = "SBO"
Sub SBO(wb, wsh_Path)
'Add Rack 1 SBO Data
Dim intn_Rack As Integer
Set wshRack = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wshRack.Name = "Rack"
'frmRack.Show
If Len(wsh_Path.Cells(7, 2)) > 5 Then
    Set wb2 = Workbooks.Open(wsh_Path.Cells(7, 2).Value2)
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Rack").Range("A1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("Rack").Range("B1")
    wb2.Sheets(1).Range("G:G").Copy Destination:=wb.Sheets("Rack").Range("C1")
    wb2.Close
    With Sheets("Rack")
        intn_Rack = .Cells(Rows.count, 1).End(xlUp).Row
        intn_Report = Sheets("Report").Cells(Rows.count, 6).End(xlUp).Row
        Dim intCountforSBO As Integer: intCountforSBO = 0
        For i = 2 To intn_Rack Step 1
        Dim LArrayRangeCL() As String
        LArrayRangeCL = Split(.Cells(i, 2).Value2, "_")
            If LArrayRangeCL(1) = "CL" Then
                intCountforSBO = intCountforSBO + 1
                .Cells(i, 4).Value2 = Right(.Cells(i, 1).Value2, Len(.Cells(i, 1)) - 4)
                .Cells(i, 5).Value2 = Right(.Cells(i, 2).Value2, Len(.Cells(i, 2)) - 1)
                .Cells(i, 5).Value2 = Left(.Cells(i, 5).Value2, Len(.Cells(i, 5)) - 3)
                    'Add data to report
                    Sheets("Report").Cells(intn_Report + intCountforSBO, 1).Value2 = .Cells(i, 3).Value2
                    Sheets("Report").Cells(intn_Report + intCountforSBO, 4).Value = 1
                    Sheets("Report").Cells(intn_Report + intCountforSBO, 5).Value2 = .Cells(i, 4).Value2
                    Sheets("Report").Cells(intn_Report + intCountforSBO, 6).Value2 = .Cells(i, 5).Value2
                    Sheets("Report").Cells(intn_Report + intCountforSBO, 13).Value2 = "WR_X_SBO"
            End If
        Next i
    End With
End If

End Sub
