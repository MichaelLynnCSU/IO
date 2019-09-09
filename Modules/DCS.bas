Attribute VB_Name = "DCS"
Sub HDCC()
Set wb3 = ThisWorkbook
Set wb = ThisWorkbook
blnPlaceHolder = False

    If blnPlaceHolder = True Then
    
        'Add block to report
        Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws3.Name = "Check_blocks"
            
        Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws3.Name = "Check_blocks2"
            
          
        DCSUI.Show
        
        Dim localData As String
        Dim dcsData As String
        Dim localRTUData As String
        Dim dcsRTUData As String
        
        Dim dcs2njh As String
        Dim dcs2chh As String
        Dim dcs2tfh As String
        
            localData = ThisWorkbook.Sheets("Check_blocks").Cells(1, 1).Value2
            dcsData = ThisWorkbook.Sheets("Check_blocks").Cells(1, 2).Value2
            localRTUData = ThisWorkbook.Sheets("Check_blocks").Cells(1, 3).Value2
            dcsRTUData = ThisWorkbook.Sheets("Check_blocks").Cells(1, 4).Value2
        
        'Add block to report
        Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws3.Name = localData
        Set oConn = CreateObject("ADODB.Connection")
        Set mrs = CreateObject("ADODB.Recordset")
        Dim myPath As String
        myPath = "X:\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\WIP Michael L\IO_List_WIP\WIP\105\export data\"
        oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & myPath & ";Extended Properties='text;HDR=YES;FMT=Delimited';"
        oConn.Open
        sSQLSting = "SELECT * From NJH_Info.csv"
        mrs.Open sSQLSting, oConn
        '=>Paste the data into a sheet
        ActiveSheet.Range("A:S").CopyFromRecordset mrs
        'Add block to report
        
        'Set wb3 = Workbooks.Open("\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\DCS data export\export data\" & localData & ".csv")
        'wb3.Sheets(1).Range("A:S").Copy Destination:=wb.Sheets(localData).Range("A:S")
        'wb3.Close
        
        
        'Add block to report
        Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws3.Name = dcsData
        Set wb3 = Workbooks.Open("\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\DCS data export\export data\" & dcsData & ".csv")
        wb3.Sheets(1).Range("A:S").Copy Destination:=wb.Sheets(dcsData).Range("A:S")
        wb3.Close
        
        'Add block to report
        Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws3.Name = localRTUData
        Set wb3 = Workbooks.Open("\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\DCS data export\export data\" & localRTUData & ".csv")
        wb3.Sheets(1).Range("A:S").Copy Destination:=wb.Sheets(localRTUData).Range("A:S")
        wb3.Close
        
        'Add block to report
        Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws3.Name = dcsRTUData
        Set wb3 = Workbooks.Open("\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\DCS data export\export data\" & dcsRTUData & ".csv")
        wb3.Sheets(1).Range("A:S").Copy Destination:=wb.Sheets(dcsRTUData).Range("A:S")
        wb3.Close
        
        
        
        ' part A
                
                Dim intRow As Integer
                Dim intCol As Integer
                ' compare DCS against local for messages
                intRow = Sheets(dcsData).UsedRange.Rows.count
                intCol = Sheets(localData).UsedRange.Rows.count
                
                      For k = 2 To intRow Step 1
                        For j = 2 To intCol Step 1
                            If Sheets(dcsData).Cells(k, 2).Value2 = Sheets(localData).Cells(j, 2).Value2 Then
                                If Sheets(dcsData).Cells(k, 4).Value2 = Sheets(localData).Cells(j, 4).Value2 Then
                                    If Sheets(dcsData).Cells(k, 9).Value2 = Sheets(localData).Cells(j, 9).Value2 Then
                                           If Replace(LCase(Sheets(dcsData).Cells(k, 12).Value2), " ", "") <> Replace(LCase(Sheets(localData).Cells(j, 12).Value2), " ", "") Then
                                                Sheets(dcsData).Cells(k, 16).Value2 = Sheets(localData).Cells(j, 12).Value2
                                                Sheets(dcsData).Cells(k, 11).Value2 = Sheets(localData).Cells(j, 10).Value2
                                            End If
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                     Next
        
        ' Part B
                Sheets(dcsData).Cells(1, 11).Value2 = "SP Local"
                Sheets(dcsData).Cells(1, 10).Value2 = "SP DCS"
                Dim search As Boolean: search = True
                'compare locals against DCS singals
                intRow = Sheets(localData).UsedRange.Rows.count
                intCol = Sheets(dcsData).UsedRange.Rows.count
                Dim count As Integer: count = 1
                      For k = 2 To intRow Step 1
                      search = True
                        For j = 2 To intCol Step 1
                            If Sheets(localData).Cells(k, 2).Value2 = Sheets(dcsData).Cells(j, 2).Value2 Then
                                If Replace(Sheets(localData).Cells(k, 4).Value2, " ", "") = Replace(Sheets(dcsData).Cells(j, 4).Value2, " ", "") Then
                                    If Sheets(localData).Cells(k, 9).Value2 = Sheets(dcsData).Cells(j, 9).Value2 Then
                                        Sheets(dcsData).Cells(k, 11).Value2 = Sheets(localData).Cells(j, 10).Value2
                                        search = False
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                        If search = True Then
                            Dim SP As String
                            SP = "A" & k & ":M" & k
                            Sheets(localData).Activate
                            Range(SP).Select
                            Selection.Copy
                            Sheets("Check_blocks").Select
                            Range("A" & count).Select
                            ActiveSheet.Paste
                            count = count + 1
                            search = False
                        End If
                     Next
                     
        'Part C
        
                ' compare DCS against local RTU messages
                row_range = Sheets(dcsRTUData).UsedRange.Rows.count
                row_range2 = Sheets(localRTUData).UsedRange.Rows.count
                
                      For k = 2 To intRow Step 1
                        For j = 2 To intCol Step 1
                            If Sheets(dcsRTUData).Cells(k, 2).Value2 = Sheets(localRTUData).Cells(j, 2).Value2 Then
                                If Sheets(dcsRTUData).Cells(k, 4).Value2 = Sheets(localRTUData).Cells(j, 4).Value2 Then
                                    If Sheets(dcsRTUData).Cells(k, 9).Value2 = Sheets(localRTUData).Cells(j, 9).Value2 Then
                                           If Replace(LCase(Sheets(dcsRTUData).Cells(k, 7).Value2), " ", "") <> Replace(LCase(Sheets(localRTUData).Cells(j, 7).Value2), " ", "") Then
                                                Sheets(dcsRTUData).Cells(k, 13).Value2 = Sheets(localRTUData).Cells(j, 7).Value2
                                            End If
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                     Next
                     
                     
        
        ' part D
                search = True
                'compare locals against DCS RTU
                intRow = Sheets(localRTUData).UsedRange.Rows.count
                intCol = Sheets(dcsRTUData).UsedRange.Rows.count
                count = 1
                      For k = 2 To intRow Step 1
                      search = True
                        For j = 2 To intCol Step 1
                            If Sheets(localRTUData).Cells(k, 2).Value2 = Sheets(dcsRTUData).Cells(j, 2).Value2 Then
                                If Replace(Sheets(localRTUData).Cells(k, 4).Value2, " ", "") = Replace(Sheets(dcsRTUData).Cells(j, 4).Value2, " ", "") Then
                                    If Sheets(localRTUData).Cells(k, 9).Value2 = Sheets(dcsRTUData).Cells(j, 9).Value2 Then
                                        search = False
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                        If search = True Then
                            SP = "A" & k & ":M" & k
                            Sheets(localRTUData).Activate
                            Range(SP).Select
                            Selection.Copy
                            Sheets("Check_blocks2").Select
                            Range("A" & count).Select
                            ActiveSheet.Paste
                            count = count + 1
                            search = False
                        End If
                     Next
                     
                                     
        '            Application.DisplayAlerts = False
        '                     Worksheets(l1).Delete
        '             Application.DisplayAlerts = True
        '
        '              Debug.Print "test"
        
    Else
    
            ' check for duplicates, NO/NC
            'Add block to report
            Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
                ws3.Name = "normalOpenCheck"
            Set wb3 = Workbooks.Open("\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\WIP Michael L\IO_List_WIP\WIP\116\DCS2\TFH IO List rev B.csv")
            wb3.Sheets(1).Range("A:AB").Copy Destination:=wb.Sheets("normalOpenCheck").Range("A:AB")
            wb3.Close
            
            Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
                ws3.Name = "duplicates"
                
            'algorithm part a
                Sheets("normalOpenCheck").Rows(1).EntireRow.Copy
                Sheets("duplicates").Rows(1).Select
                Sheets("duplicates").Paste
                
                Dim row_range3 As Integer
                Dim count3 As Integer
                count3 = 2
                
                row_range3 = Sheets("normalOpenCheck").UsedRange.Rows.count
                Do While row_range3 > 0
                    For k = (row_range3 - 1) To 1 Step -1
                        If Sheets("normalOpenCheck").Cells(k, 2).Value2 = Sheets("normalOpenCheck").Cells(row_range3, 2).Value2 Then
                            If Sheets("normalOpenCheck").Cells(row_range3, 2).Value2 <> "SPARE" Then
                            
                                Sheets("normalOpenCheck").Rows(k).EntireRow.Copy
                                Sheets("duplicates").Cells(count3, 1).Select
                                Sheets("duplicates").Paste
                                
                                Sheets("normalOpenCheck").Rows(row_range3).EntireRow.Copy
                                Sheets("duplicates").Cells(count3 + 1, 1).Select
                                Sheets("duplicates").Paste
                                
                                count3 = count3 + 2
                            End If
                        End If
                    Next
                row_range3 = row_range3 - 1
                Loop
            
    End If
    
End Sub
