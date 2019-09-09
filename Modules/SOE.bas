Attribute VB_Name = "SOE"
Sub SOE(wb, wsh_Path)
'Add SOE Data
Dim intn_SOE As Integer
Set wshSOE = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wshSOE.Name = "SOE"
'frmSOE.Show
If Len(wsh_Path.Cells(9, 2)) > 5 Then
    Set wb2 = Workbooks.Open(wsh_Path.Cells(9, 2).Value2)
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("SOE").Range("A1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("SOE").Range("B1")
    wb2.Sheets(1).Range("G:G").Copy Destination:=wb.Sheets("SOE").Range("C1")
    wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("SOE").Range("D1")
    wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("SOE").Range("E1")
    wb2.Close
    With Sheets("SOE")
        intn_SOE = .Cells(Rows.count, 1).End(xlUp).Row
        intn_Report = Sheets("Report").Cells(Rows.count, 6).End(xlUp).Row
        For i = 2 To intn_SOE Step 1
            .Cells(i, 6).Value2 = Right(.Cells(i, 1).Value2, Len(.Cells(i, 1).Value2) - 4)
            .Cells(i, 7).Value2 = Left(.Cells(i, 2).Value2, 1)
            If .Cells(i, 7).Value2 = "Q" Then
                .Cells(i, 8).Value2 = Right(.Cells(i, 2).Value2, Len(.Cells(i, 2).Value2) - 1)
            Else: .Cells(i, 8).Value2 = Right(.Cells(i, 2).Value2, Len(.Cells(i, 2).Value2) - 3)
            End If
        Next i
        .Cells(1, 6).Value2 = "Slot #"
        .Cells(1, 7).Value2 = "I/Q"
        .Cells(1, 8).Value2 = "Channel #"
    End With
    With Sheets("SOE").Range("A:H")
            .Cells.Sort Key1:=.Columns(Application.Match("Slot #", .Rows(1), 0)), Order1:=xlAscending, _
                        Key2:=.Columns(Application.Match("Channel #", .Rows(1), 0)), Order2:=xlAscending, _
                        Key3:=.Columns(Application.Match("I/Q", .Rows(1), 0)), Order2:=xlAscending, _
                        Orientation:=xlTopToBottom, Header:=xlYes
    End With
    With Sheets("SOE")
        For i = 3 To intn_SOE Step 2
            Sheets("Report").Cells(intn_Report + (i - 1) / 2, 1).Value2 = .Cells(i, 3).Value2
            Sheets("Report").Cells(intn_Report + (i - 1) / 2, 4).Value = 1
            Sheets("Report").Cells(intn_Report + (i - 1) / 2, 5).Value2 = .Cells(i, 6).Value2
            Sheets("Report").Cells(intn_Report + (i - 1) / 2, 6).Value2 = .Cells(i, 8).Value2
            If .Cells(i - 1, 4).Value = 0 Then
                Sheets("Report").Cells(intn_Report + (i - 1) / 2, 14).Value2 = "NO"
            Else: Sheets("Report").Cells(intn_Report + (i - 1) / 2, 14).Value2 = "NC"
            End If
            Sheets("Report").Cells(intn_Report + (i - 1) / 2, 22).Value2 = .Cells(i, 5).Value2
            Sheets("Report").Cells(intn_Report + (i - 1) / 2, 13).Value2 = "RD_X_SOE"
        Next i
    End With
End If

'Add SOE_Message Data
Dim intn_SOE_Message As Integer
Set wshSOE_Message = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wshSOE_Message.Name = "SOE Message"
'frmSOE_Message.Show
If Len(wsh_Path.Cells(10, 2)) > 5 Then
    Set wb2 = Workbooks.Open(wsh_Path.Cells(10, 2).Value2)
    wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("SOE Message").Range("A1")
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("SOE Message").Range("B1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("SOE Message").Range("C1")
    wb2.Sheets(1).Range("H:H").Copy Destination:=wb.Sheets("SOE Message").Range("D1")
    wb2.Sheets(1).Range("N:N").Copy Destination:=wb.Sheets("SOE Message").Range("E1")
    wb2.Close
    With Sheets("SOE Message")
        intn_SOE_Message = .Cells(Rows.count, 1).End(xlUp).Row
        .Cells(1, 6).Value2 = "Slot #"
        .Cells(1, 7).Value2 = "Msg #"
        .Cells(1, 8).Value2 = "Sig #"
        .Cells(1, 9).Value2 = "Channel #"
        For i = 2 To intn_SOE_Message Step 1
            .Cells(i, 6).Value2 = Right(.Cells(i, 2).Value2, Len(.Cells(i, 2).Value2) - 4)
            .Cells(i, 7).Value2 = Right(.Cells(i, 3).Value2, 1)
            .Cells(i, 8).Value2 = Right(.Cells(i, 4).Value2, 1)
            If .Cells(i, 7).Value2 = 1 Then
                .Cells(i, 9).Value = .Cells(i, 8).Value - 1
            ElseIf .Cells(i, 7).Value2 = 2 Then
                .Cells(i, 9).Value = .Cells(i, 8).Value2 + 7
            ElseIf .Cells(i, 7).Value2 = 3 Then
                .Cells(i, 9).Value = .Cells(i, 8).Value + 15
            ElseIf .Cells(i, 7).Value2 = 4 Then
                .Cells(i, 9).Value = .Cells(i, 8).Value + 23
            End If
        Next i
    End With
    
    
    
    
    
    
    'Add Alarm Text to Report
    With Sheets("Report")
        For i = intn_Report To intn_Report + intn_SOE Step 1
            For j = 2 To intn_SOE_Message Step 1
                If .Cells(i, 5).Value2 = Sheets("SOE Message").Cells(j, 6).Value2 And _
                    .Cells(i, 6).Value2 = Sheets("SOE Message").Cells(j, 9).Value2 And _
                    .Cells(i, 22).Value2 = Sheets("SOE Message").Cells(j, 1).Value2 Then
                        .Cells(i, 15).Value2 = Sheets("SOE Message").Cells(j, 5).Value2
                End If
            Next j
        Next i
    End With
End If


        'Sort by Rack then Slot then Channel #
        With Sheets("Report").Range("A:Z")
                .Cells.Sort Key1:=.Columns(Application.Match("Type", .Rows(1), 0)), Order1:=xlAscending, _
                            Key2:=.Columns(Application.Match("Rack", .Rows(1), 0)), Order2:=xlAscending, _
                            Key3:=.Columns(Application.Match("Slot", .Rows(1), 0)), Order2:=xlAscending, _
                            Key3:=.Columns(Application.Match("Channel", .Rows(1), 0)), Order2:=xlAscending, _
                            Orientation:=xlTopToBottom, Header:=xlYes
        End With
                
                

'Move the SOE data
'Add the new soe tab
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws2.Name = "SOE_Seperator"
  
       
        Dim iRowsForSOE As Integer
        iRowsForSOE = Sheets("Report").UsedRange.Rows.count
            
        Dim iIndex As Integer
        iIndex = 1
        
        Dim bRunnung As Boolean
        bRunnung = True
        
        Dim iStartCount As Integer
        iStartCount = 0
        
        For q = 2 To iRowsForSOE Step 1
          
           '   If q > 659 Then
                    Dim strCurrentSym As String
                    ' get the current symbol
                    strCurrentSym = Sheets("Report").Cells(q, 2)
                      
                    ' check if the current symbol is the type we want
                    Dim strCheckType As String
                    strCheckType = Sheets("Report").Cells(q, 13)
                
                    
                    If strCheckType = "RD_X_SOE" Then
                             If bRunnung Then
                                iStartCount = q
                                bRunnung = False
                             End If
                             
                        'Sheets("SOE_Seperator").Cells(2, 2).Value = "test"
                        Sheets("Report").Rows(q).EntireRow.Copy
                        Sheets("SOE_Seperator").Range("A" & iIndex).PasteSpecial Paste:=xlValues
                        iIndex = iIndex + 1
                                           
                    End If
              
             ' End If
          
              

        Next
'
'        ' Delete SEO fom report
'        Do While iIndex >= 1
'
'            Sheets("Report").Rows(iStartCount).EntireRow.Delete
'
'         iIndex = iIndex - 1
'        Loop

End Sub
