Attribute VB_Name = "RDX"
Sub RDX(wb, wsh_Path)
'Add Rack 1 AI1 Data
Dim intn_AI As Integer
Set wshAI = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wshAI.Name = "AI"
'frmAI.Show
If Len(wsh_Path.Cells(8, 2)) > 5 Then
    Set wb2 = Workbooks.Open(wsh_Path.Cells(8, 2).Value2)
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("AI").Range("A1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("AI").Range("B1")
    wb2.Sheets(1).Range("G:G").Copy Destination:=wb.Sheets("AI").Range("C1")
    wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("AI").Range("D1")
    Dim BlockType As Integer: BlockType = wb2.Sheets(1).UsedRange.Columns.count
    Dim blockTypeRows As Integer: blockTypeRows = wb2.Sheets(1).UsedRange.Rows.count
    For i = 2 To BlockType Step 1
        If wb2.Sheets(1).Cells(1, i).Value = "Block type" Then
                wb2.Sheets(1).Activate
                wb2.Sheets(1).Select
                Range(Cells(1, i), Cells(blockTypeRows, i)).Copy
                wb.Sheets("AI").Activate
                wb.Sheets("AI").Select
                Range(Cells(1, 10), Cells(blockTypeRows, 10)).Select
                ActiveSheet.Paste

        End If
    Next

'    wb2.Sheets(1).Range("AB:AB").Copy Destination:=wb.Sheets("AI").Range("J1")
    wb2.Close

    
   
    
    With Sheets("AI")
        intn_AI = .Cells(Rows.count, 1).End(xlUp).Row
        For i = 2 To intn_AI Step 1
            .Cells(i, 5).Value2 = Right(.Cells(i, 1).Value2, Len(.Cells(i, 1)) - 4)
            .Cells(i, 6).Value2 = Left(.Cells(i, 2).Value2, 1)
            .Cells(i, 7).Value2 = Right(.Cells(i, 2).Value2, Len(.Cells(i, 2)) - 1)
            If .Cells(i, 6).Value2 = "V" Then
                .Cells(i, 7).Value2 = Left(.Cells(i, 7).Value2, Len(.Cells(i, 7)) - 5)
                .Cells(i, 8).Value2 = Right(.Cells(i, 2).Value2, 5)
            End If
        Next i
        .Cells(1, 5).Value2 = "Slot #"
        .Cells(1, 6).Value2 = "V/Q"
        .Cells(1, 7).Value2 = "Channel #"
        .Cells(1, 8).Value2 = "Low/High"
    End With
    With Sheets("AI").Range("A:H")
        .Cells.Sort Key1:=.Columns(Application.Match("Slot #", .Rows(1), 0)), Order1:=xlAscending, _
                    Key2:=.Columns(Application.Match("Channel #", .Rows(1), 0)), Order2:=xlAscending, _
                    Key3:=.Columns(Application.Match("Low/High", .Rows(1), 0)), Order2:=xlAscending, _
                    Orientation:=xlTopToBottom, Header:=xlYes
    End With
    
    
    'Add data to report
    With Sheets("AI")
        intn_Report = Sheets("Report").Cells(Rows.count, 6).End(xlUp).Row
        For i = 4 To intn_AI Step 3
            Sheets("Report").Cells(intn_Report + (i - 1) / 3, 1).Value2 = .Cells(i, 3).Value2
            Sheets("Report").Cells(intn_Report + (i - 1) / 3, 4).Value = 1
            Sheets("Report").Cells(intn_Report + (i - 1) / 3, 5).Value2 = .Cells(i, 5).Value2
            Sheets("Report").Cells(intn_Report + (i - 1) / 3, 6).Value2 = .Cells(i, 7).Value2
            Sheets("Report").Cells(intn_Report + (i - 1) / 3, 7).Value2 = .Cells(i - 2, 4).Value2
            Sheets("Report").Cells(intn_Report + (i - 1) / 3, 8).Value2 = .Cells(i - 1, 4).Value2
            Sheets("Report").Cells(intn_Report + i / 2, 13).Value2 = "RD_X_AI1"
        Next i
    End With
End If

'Add relevant data to NOC
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws2.Name = "Normal OC"
'frmNewDigit.Show
Set wb2 = Workbooks.Open(wsh_Path.Cells(14, 2).Value2)
wb2.Sheets(1).Range("L:L").Copy Destination:=wb.Sheets("Normal OC").Range("A1")
wb2.Sheets(1).Range("K:K").Copy Destination:=wb.Sheets("Normal OC").Range("B1")
wb2.Close
End Sub

