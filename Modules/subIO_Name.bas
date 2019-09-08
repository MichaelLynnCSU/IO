Attribute VB_Name = "subIO_Name"
Sub subIO_DeleteSheets()

'Clean up workbook
On Error Resume Next


Application.DisplayAlerts = False
    Sheets("Signal Connections").Delete
    Sheets("HWConfig").Delete
    Sheets("Range").Delete
    Sheets("Alarm").Delete
    Sheets("Symbol Table").Delete
    Sheets("Rack").Delete
    Sheets("AI").Delete
    Sheets("SOE").Delete
    Sheets("SOE Message").Delete
    Sheets("DI Signal").Delete
    Sheets("DI").Delete
    Sheets("DI Alarm").Delete
    Sheets("File Paths").Delete
    Sheets("Report").Delete
Application.DisplayAlerts = True

End Sub


Sub subIO_Name()




'
' IO_Name Macro
'
'read in data simple save rerun
Application.ScreenUpdating = False
ActiveSheet.Name = "HWConfig"
Set wb = ThisWorkbook
Dim intc, intr As Integer
intr = 1
intc = 1
Set wsh_Path = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsh_Path.Name = "File Paths"
    wsh_Path.Cells(1, 1).Value2 = "File Name"
    wsh_Path.Cells(1, 2).Value2 = "File Path"
    wsh_Path.Columns("A:A").ColumnWidth = 20
    wsh_Path.Columns("B:B").ColumnWidth = 100
frmHWConfig.Show
Sheets("HWConfig").Select
Open wsh_Path.Cells(2, 2).Value2 For Input As #1
'keep data seperated in different columns (each group in original text file is seperated by an empty row)
Do Until EOF(1)
    Line Input #1, readLine
    If readLine = "" Then
        intc = intc + 1
        intr = 1
    Else
        ActiveSheet.Cells(intr, intc).Value = readLine
        intr = intr + 1
    End If
Loop
Close #1

'Delete any irrelevate sections (relevant sections are those with data for slots > 4)
For j = intc To 1 Step -1
    If Not InStr(Cells(1, j).Value, "DPADDRESS") > 0 Then
        Cells(1, j).EntireColumn.Delete
    End If
Next j
For j = intc To 1 Step -1
    If InStr(Cells(1, j).Value, " SLOT 1") > 0 Then
        Cells(1, j).EntireColumn.Delete
    End If
Next j
For j = intc To 1 Step -1
    If InStr(Cells(1, j).Value, " SLOT 2") > 0 Then
        Cells(1, j).EntireColumn.Delete
    End If
Next j
For j = intc To 1 Step -1
    If InStr(Cells(1, j).Value, " SLOT 3") > 0 Then
        Cells(1, j).EntireColumn.Delete
    End If
Next j

'Create template for report/summary
Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "Report"
With Sheets("Report")
    .Cells(1, 1).Value = "Symbol"
    .Cells(1, 2).Value = "Address"
    .Cells(1, 3).Value = "Comment"
    .Cells(1, 4).Value = "Rack"
    .Cells(1, 5).Value = "Slot"
    .Cells(1, 6).Value = "Channel"
    .Cells(1, 7).Value = "Range - LR"
    .Cells(1, 8).Value = "Range - HR"
    .Cells(1, 9).Value = "Alarm - WL"
    .Cells(1, 10).Value = "Alarm - WH"
    .Cells(1, 11).Value = "Alarm - AL"
    .Cells(1, 12).Value = "Alarm - AH"
    .Cells(1, 13).Value = "Type"
    .Cells(1, 14).Value = "NO/NC"
    .Cells(1, 15).Value = "Alarm Text"
    .Cells(1, 16).Value = "Block"
    .Cells(1, 17).Value = "Alarm Block"
    .Cells(1, 18).Value = "Chart"
    .Cells(1, 25).Value = "ET200M/RTU"
End With

Dim intn_Data, intn_Report, intk, intai As Integer
Dim strchannel, strsymbol, strcomment, strtype, straitype, strrack, strslot, strDig, strRTU As String

'Extract data and move to report
With Sheets("HWConfig")
    For j = intc To 1 Step -1
        intn_Data = .Cells(Rows.Count, j).End(xlUp).Row
        intn_Report = Sheets("Report").Cells(Rows.Count, 6).End(xlUp).Row
        intk = 0
        intai = 0
        'Digital?
        strDig = InStr(1, .Cells(1, j), "DI")
        'Extract rack #
        strrack = Mid(.Cells(1, j), InStr(1, .Cells(1, j), ",") + Len("DPADDRESS") + 3, 2)
        'Extract slot #
        strslot = Mid(.Cells(1, j), InStr(1, .Cells(1, j), ",") + Len("DPADDRESS") + 3 + Len(strrack) + 3 + Len("SLOT"), 1)
        'ET200M
        If InStr(1, .Cells(1, j), "IM 153-2") > 0 Then
            strRTU = "ET200M"
        End If
        'Extract symbol, comment and channel #
        For i = 1 To intn_Data Step 1
            If Left(.Cells(i, j).Value, 6) = "SYMBOL" Then
                'Keep track of how many
                intk = intk + 1
                'Extract channel # (between 1st and 2nd comma)
                strchannel = Trim(Mid(.Cells(i, j), InStr(1, .Cells(i, j), ",") + 1, 3))
                If Right(strchannel, 1) = "," Then
                    strchannel = Left(strchannel, Len(strchannel) - 1)
                End If
                'Extract symbol (between 2nd and 3rd comma)
                strsymbol = Mid(.Cells(i, j), InStr(1, .Cells(i, j), strchannel))
                strsymbol = Right(strsymbol, Len(strsymbol) - Len(strchannel) - 3)
                
                'Extract comment (to the right of last comma)
                strcomment = Right(strsymbol, Len(strsymbol) - InStr(strsymbol, ",") - 2)
                strcomment = Left(strcomment, Len(strcomment) - 1)
                strsymbol = Left(strsymbol, Len(strsymbol) - Len(strcomment) - 5)
                
                'Add channel #, symbol and comment to report
                Sheets("Report").Cells(intk + intn_Report, 6).Value2 = strchannel
                Sheets("Report").Cells(intk + intn_Report, 1).Value2 = strsymbol
                Sheets("Report").Cells(intk + intn_Report, 3).Value2 = strcomment
                Sheets("Report").Cells(intk + intn_Report, 4).Value2 = strrack
                Sheets("Report").Cells(intk + intn_Report, 5).Value2 = strslot
                Sheets("Report").Cells(intk + intn_Report, 24).Value2 = strDig
                If strDig > 0 Then
                    Sheets("Report").Cells(intk + intn_Report, 13).Value2 = "Digital"
                End If
                'Add RTU/ET200M to report
                Sheets("Report").Cells(intk + intn_Report, 25).Value2 = strRTU
            End If
            'Extract signal type
            If Left(.Cells(i, j).Value, 5) = "  AI_" Then
                intai = intai + 1
                strAI = Mid(.Cells(i, j), InStrRev(.Cells(i, j), ",") - 1, 1)
                straitype = Mid(.Cells(i, j), InStrRev(.Cells(i, j), ",") + 3)
                straitype = Left(straitype, Len(straitype) - 1)
                'Add signal type to report
                If strAI = 0 Then
                    For k = intn_Report + 1 To intk + intn_Report Step 1
                        If Sheets("Report").Cells(k, 6) = 0 Or Sheets("Report").Cells(k, 6) = 1 Then
                            If Not Len(Sheets("Report").Cells(k, 13)) > 0 Then
                                Sheets("Report").Cells(k, 13).Value2 = straitype
                            Else: Sheets("Report").Cells(k, 13).Value2 = Sheets("Report").Cells(k, 13).Value2 & ", " & straitype
                            End If
                        End If
                    Next k
                ElseIf strAI = 1 Then
                    For k = intn_Report + 1 To intk + intn_Report Step 1
                        If Sheets("Report").Cells(k, 6) = 2 Or Sheets("Report").Cells(k, 6) = 3 Then
                            If Not Len(Sheets("Report").Cells(k, 13)) > 0 Then
                                Sheets("Report").Cells(k, 13).Value2 = straitype
                            Else: Sheets("Report").Cells(k, 13).Value2 = Sheets("Report").Cells(k, 13).Value2 & ", " & straitype
                            End If
                        End If
                    Next k
                ElseIf strAI = 2 Then
                    For k = intn_Report + 1 To intk + intn_Report Step 1
                        If Sheets("Report").Cells(k, 6) = 4 Or Sheets("Report").Cells(k, 6) = 5 Then
                            If Not Len(Sheets("Report").Cells(k, 13)) > 0 Then
                                Sheets("Report").Cells(k, 13).Value2 = straitype
                            Else: Sheets("Report").Cells(k, 13).Value2 = Sheets("Report").Cells(k, 13).Value2 & ", " & straitype
                            End If
                        End If
                    Next k
                ElseIf strAI = 3 Then
                    For k = intn_Report + 1 To intk + intn_Report Step 1
                        If Sheets("Report").Cells(k, 6) = 6 Or Sheets("Report").Cells(k, 6) = 7 Then
                            If Not Len(Sheets("Report").Cells(k, 13)) > 0 Then
                                Sheets("Report").Cells(k, 13).Value2 = straitype
                            Else: Sheets("Report").Cells(k, 13).Value2 = Sheets("Report").Cells(k, 13).Value2 & ", " & straitype
                            End If
                        End If
                    Next k
                End If
            End If
        Next i
    Next j
End With

'Pull over block
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws2.Name = "Signal Connections"
frmCH_AI_Signals.Show
Set wb2 = Workbooks.Open(wsh_Path.Cells(3, 2).Value2)
wb2.Sheets(1).Range("K:K").Copy Destination:=wb.Sheets("Signal Connections").Range("A1")
wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Signal Connections").Range("B1")
wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Signal Connections").Range("C1")
wb2.Close
'Add block to report
Dim strLoc, strBlock As String
Dim intn_Signal, intn_Alarm, intmatch As Integer
intn_Report = Sheets("Report").Cells(Rows.Count, 6).End(xlUp).Row
intn_Signal = Sheets("Signal Connections").Cells(Rows.Count, 1).End(xlUp).Row
strBlock = ""
With Sheets("Report")
For i = 2 To intn_Report Step 1
    For j = 2 To intn_Signal Step 1
        If .Cells(i, 1).Value = Sheets("Signal Connections").Cells(j, 1).Value Then
            .Cells(i, 16).Value = Sheets("Signal Connections").Cells(j, 2).Value
            .Cells(i, 23).Value2 = Sheets("Signal Connections").Cells(j, 3).Value
        End If
    Next j
Next i
End With

'Pull Over Range Values
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws2.Name = "Range"
frmRanges.Show
If Not Len(wsh_Path.Cells(4, 2)) > 0 Then
    frmRanges.Show
End If
Set wb2 = Workbooks.Open(wsh_Path.Cells(4, 2).Value2)
wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Range").Range("A1")
wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("Range").Range("B1")
wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("Range").Range("C1")
wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Range").Range("D1")
wb2.Sheets(1).Range("L:L").Copy Destination:=wb.Sheets("Range").Range("E1")
wb2.Close
'sort by block then I/O Name
With Sheets("Range").Range("A:G")
        .Cells.Sort Key1:=.Columns(Application.Match("Block", .Rows(1), 0)), Order1:=xlAscending, _
                    Key2:=.Columns(Application.Match("I/O name", .Rows(1), 0)), Order2:=xlDescending, _
                    Orientation:=xlTopToBottom, Header:=xlYes
End With
intn_Range = Sheets("Range").Cells(Rows.Count, 1).End(xlUp).Row
'Add new block to report
Range("F2").Select
ActiveCell.FormulaR1C1 = "=MID(RC[-1],FIND(RC[-2],RC[-1])+LEN(RC[-2])+1,LEN(RC[-1])-FIND(RC[-2],RC[-1])+1)"
Range("G2").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],FIND("".U"",RC[-1])-1)"
Range("F2:G2").Select
Selection.AutoFill Destination:=Range("F2:G" & intn_Range)
Range("F2:G" & intn_Range).Select
'Add range to report
Sheets("Report").Select
Range("G2").Select
ActiveCell.FormulaR1C1 = "=MATCH(RC16,Range!C1,0)"
Range("G2").Select
Selection.AutoFill Destination:=Range("G2:H2"), Type:=xlFillDefault
Range("G2:H2").Select
Selection.AutoFill Destination:=Range("G2:H" & intn_Report), Type:=xlFillDefault
With Sheets("Report")
For i = 2 To intn_Report Step 1
    If Not Len(.Cells(i, 16).Value) > 0 Then
        .Cells(i, 7).Value = ""
        .Cells(i, 8).Value = ""
    Else:
        If IsError(Sheets("Range").Cells(.Cells(i, 7).Value + 2, 7)) = True Then
            .Cells(i, 17).Value = ""
        Else: .Cells(i, 17).Value = Sheets("Range").Cells(.Cells(i, 7).Value + 2, 7).Value
        End If
        .Cells(i, 7).Value = Sheets("Range").Cells(.Cells(i, 7).Value, 3).Value
        .Cells(i, 8).Value = Sheets("Range").Cells(.Cells(i, 8).Value + 1, 3).Value
    End If
Next i
End With

'Pull Over Alarm Values
Dim intz As Integer
Dim intn_Working As Integer
Dim strAlarmBlock As String
Dim intAH As Integer
Dim intWH As Integer
Dim intWL As Integer
Dim intAL As Integer
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws2.Name = "Alarm"
frmAlarm.Show


'Pull columns from Nickajack_Plant_NJH_Meas_Mon_Alarming
Set wb2 = Workbooks.Open(wsh_Path.Cells(5, 2).Value2)
wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Alarm").Range("A1")
wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("Alarm").Range("B1")
wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("Alarm").Range("C1")
wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Alarm").Range("D1")
wb2.Close

''sort by block then I/O Name
'With Sheets("Alarm").Range("A:C")
'        .Cells.Sort Key1:=.Columns(Application.Match("Block", .Rows(1), 0)), Order1:=xlAscending, MatchCase:=True, _
'                    Key2:=.Columns(Application.Match("I/O name", .Rows(1), 0)), Order2:=xlAscending, MatchCase:=True, _
'                    Orientation:=xlTopToBottom, Header:=xlYes
'End With
intn_Alarm = Sheets("Alarm").Cells(Rows.Count, 1).End(xlUp).Row


'Testing ouput for relevant sheets that exist
Debug.Print "Report: "
Debug.Print "Symbol: " & Sheets("Signal Connections").Cells(2, 1)

Debug.Print "Signal Connections: "
Debug.Print "Signal: " & Sheets("Signal Connections").Cells(1, 1)
Debug.Print "Block: " & Sheets("Signal Connections").Cells(1, 2)
Debug.Print "Chart: " & Sheets("Signal Connections").Cells(1, 3)

Debug.Print "Range: "
Debug.Print "Block: " & Sheets("Range").Cells(1, 1)
Debug.Print "Chart: " & Sheets("Range").Cells(1, 4)
Debug.Print "Interconnect: " & Sheets("Range").Cells(1, 5)

Debug.Print "Alarm: "
Debug.Print "Block: " & Sheets("Alarm").Cells(1, 1)
Debug.Print "Chart: " & Sheets("Alarm").Cells(1, 4)
Debug.Print "I/O Tag: " & Sheets("Alarm").Cells(1, 10)
Debug.Print "Value: " & Sheets("Alarm").Cells(1, 11)

'Start Part A

  Dim current_symbol As String
  
  Dim current_signal As String
  Dim current_signal_block As String
  Dim current_signal_chart As String
  
  Dim rows_Symbol As Integer
  Dim rows_Signal As Integer
  
  rows_Symbol = Sheets("Report").UsedRange.Rows.Count
  rows_Signal = Sheets("Signal Connections").UsedRange.Rows.Count
  
  Debug.Print "rows_Symbol: " & rows_Symbol
  Debug.Print "rows_Signal: " & rows_Signal
  
    For i = 2 To rows_Symbol Step 1
    
      For j = 2 To rows_Signal Step 1
      
        'get a signal(i) and symbol(j) value
        current_symbol = Sheets("Report").Cells(i, 1).Value2
        current_signal = Sheets("Signal Connections").Cells(j, 1).Value2
         ' Debug.Print "current_symbol: " & current_symbol
         ' Debug.Print "current_signal: " & current_signal
        
        'search for a signal and symbol match
        If current_symbol = current_signal Then
          current_signal_block = Sheets("Signal Connections").Cells(j, 2).Value2
          current_signal_chart = Sheets("Signal Connections").Cells(j, 3).Value2
           
'          Debug.Print "current_signal: " & current_signal
'          Debug.Print "current_signal_block: " & current_signal_block
'          Debug.Print "current_signal_chart: " & current_signal_chart
          
            'Start Part B
            
                                                  
                  Dim current_range_block As String
                  Dim current_range_chart As String
                  Dim current_range_interconnetion_block As String
                
                  Dim row_range As Integer
                  row_range = Sheets("Range").UsedRange.Rows.Count
                  
                      For k = 2 To row_range Step 1
                          'get a range_block(i) and range_chart value
                          current_range_block = Sheets("Range").Cells(k, 1).Value2
                          current_range_chart = Sheets("Range").Cells(k, 4).Value2
'                            Debug.Print "current_signal_block: " & current_signal_block
'                            Debug.Print "current_signal_chart: " & current_signal_chart
'                            Debug.Print "current_range_block: " & current_range_block
'                            Debug.Print "current_range_chart: " & current_range_chart


                 
                           'search for a symbol and range match
                        If current_signal_block = current_range_block Then
                          If current_signal_chart = current_range_chart Then
                          current_range_interconnetion_block = Sheets("Range").Cells(k, 5).Value2
                          
                          
                ' start bug check
'                 If current_symbol = "UNIT 1 STATOR TEMP #10" Then
'                   Debug.Print "current_symbol: " & current_symbol
'
'                   Debug.Print "current_signal: " & current_signal
'                   Debug.Print "current_signal_chart: " & current_signal_chart
'                   Debug.Print "current_signal_block: " & current_signal_block
'
'                   Debug.Print "current_range_block: " & current_range_block
'                   Debug.Print "current_range_chart: " & current_range_chart
'                   Debug.Print "current_range_interconnetion_block: " & current_range_interconnetion_block
'                 End If
                ' End bug check
                          
                          
                          
                          'Start Part C
                                                 
                          
                            Dim intEndPos As Integer
                            Dim intStartPos As Integer
                            Dim current_range_interconnetion_block_U As String
                            
                            
                           ' Start new String parsing alogrithm
                           
                            'Separate the block from the string
                            'the .U is a marker in the string to help find where the block name is located. Start from the end and search for .U and check if .U found
                            '  at end or in middle of interconnect string
                            intEndPos = InStrRev(current_range_interconnetion_block, ".U")
                            
                            'if .U is found and it is either at the very end or in the middle but has a " right after it then find the block name
                            If intEndPos > 0 Then
                              If (intEndPos = Len(current_range_interconnetion_block) - 1) Then  'check if .U at the end of the string
                                Debug.Print "It's at the end"
                              ElseIf Asc(Mid(current_range_interconnetion_block, intEndPos + 2, 1)) = 34 Then  'if .U in the middle check if there is double quote after the U (ascii of " is 34)
                                Debug.Print "It's in the middle and has double quote after the U"
                              Else
                                Debug.Print "string not found"
                                intEndPos = 0  'set to 0 so will not try to find block name
                              End If
                              If intEndPos > 0 Then 'the .U was found so now get the block name from the string
                                intStartPos = InStrRev(current_range_interconnetion_block, "\", intEndPos)
                                current_range_interconnetion_block_U = Mid(current_range_interconnetion_block, intStartPos + 1, intEndPos - intStartPos - 1)
                                Debug.Print "Found string:"; current_range_interconnetion_block_U
                              End If
                                
                           

                              ' End String parsing alogrithm


                          'Seperate the block from the string
'                            If InStr(current_range_interconnetion_block, ".U""") > 0 Then
'                                intEndPos = InStr(current_range_interconnetion_block, ".U")
'                                intStartPos = InStrRev(current_range_interconnetion_block, "\", intEndPos)
'                                current_range_interconnetion_block_U = Mid(current_range_interconnetion_block, intStartPos + 1, intEndPos - intStartPos - 1)
'                                Debug.Print "current_range_chart: " & current_range_chart
'                                Debug.Print "current_range_interconnetion_block_U: " & current_range_interconnetion_block_U
                     
                                               
                                    Dim intRowsAlarm As Integer
                                    Dim strCurrentAlarmBlock As String
                                    Dim strCurrentAlarmChart As String
                                    
                                    Dim IOTag As String
                                    Dim intAlarmWL As String
                                    Dim intAlarmWH As String
                                    Dim intAlarmAL As String
                                    Dim intAlarmAH As String
                                    
                                    intRowsAlarm = Sheets("Alarm").UsedRange.Rows.Count
                                    
                                        For m = 2 To intRowsAlarm Step 1
                                        
                                          'get a alarm_block(i) and  alarm_chart value
                                          strCurrentAlarmBlock = Sheets("Alarm").Cells(m, 1).Value2
                                          strCurrentAlarmChart = Sheets("Alarm").Cells(m, 4).Value2
                                          
                                          IOTag = Sheets("Alarm").Cells(m, 2).Value2
                                          
                                          
                                                  
                                          intAlarmAH = Sheets("Alarm").Cells(m, 11).Value2
                                          intAlarmWH = Sheets("Alarm").Cells(m, 11).Value2
                                          
                                          intAlarmWL = Sheets("Alarm").Cells(m, 11).Value2
                                          intAlarmAL = Sheets("Alarm").Cells(m, 11).Value2
                                          
                                                                                   
                                          
'                                          Debug.Print "current_range_chart: " & current_range_chart
'                                          Debug.Print "current_range_interconnetion_block_U: " & current_range_interconnetion_block_U
'
'                                          Debug.Print "strCurrentAlarmBlock: " & strCurrentAlarmBlock
'                                          Debug.Print "strCurrentAlarmChart: " & strCurrentAlarmChart
                                          
                                          
                                          'search for a range and alarm match
                                           If strCurrentAlarmBlock = current_range_interconnetion_block_U Then
                                             If strCurrentAlarmChart = current_range_chart Then
                                             
                                                'Debug.Print "IOTag: " & IOTag
                                                
                                                If IOTag = "U_AH" Then
                                                  Debug.Print "intAlarmAH: " & intAlarmAH
                                                  Sheets("Alarm").Cells(m, 3).Copy Sheets("Report").Cells(i, 12)
                                                  
                                                End If
                                                
                                                If IOTag = "U_WH" Then
                                                  Debug.Print "intAlarmWH: " & intAlarmWH
                                                    Sheets("Alarm").Cells(m, 3).Copy Sheets("Report").Cells(i, 10)
                                                End If
                                                
                                                If IOTag = "U_WL" Then
                                                  Debug.Print "intAlarmWL: " & intAlarmWL
                                                    Sheets("Alarm").Cells(m, 3).Copy Sheets("Report").Cells(i, 9)
                                                End If
                                                
                                                If IOTag = "U_AL" Then
                                                  Debug.Print "intAlarmAL: " & intAlarmAL
                                                    Sheets("Alarm").Cells(m, 3).Copy Sheets("Report").Cells(i, 11)
                                                End If
                                                
                                           
                                                Debug.Print "current_symbol: " & current_symbol
                                                
                                                Debug.Print "current_signal_block: " & current_signal_block
                                                Debug.Print "current_signal_chart: " & current_signal_chart
                                                
                                                Debug.Print "current_range_block: " & current_range_block
                                                Debug.Print "current_range_chart: " & current_range_chart
                                                  
                                                Debug.Print "current_range_interconnetion_block: " & current_range_interconnetion_block
                                                
                                                Debug.Print "strCurrentAlarmBlock: " & strCurrentAlarmBlock
                                                Debug.Print "current_range_interconnetion_block_U: " & current_range_interconnetion_block_U

                                                Debug.Print "strCurrentAlarmChart: " & strCurrentAlarmChart
                                                Debug.Print "current_range_chart: " & current_range_chart
                                                
                                            End If
                                          End If
                                                                                         
                                Next
                          End If
                          'End Part C
                          
                          End If
                        End If
                     Next
  
            
            'End Part B
          
        End If
        
     Next
     
    Next

'End Part A
  
  
'Add alarm to report
Sheets("Report").Range("Q:Q").Copy Destination:=wb.Sheets("Alarm").Range("H1")
With Sheets("Alarm")
    For i = 1 To intn_Report Step 1
        intAH = 0
        intWH = 0
        intWL = 0
        intAL = 0
        If Len(.Cells(i, 8).Value2) > 0 Then
            For j = 1 To intn_Alarm Step 1
                If .Cells(i, 17).Value2 = .Cells(j, 1).Value2 And .Cells(i, 23).Value2 = .Cells(j, 4).Value2 Then
                    If .Cells(j, 2).Value2 = "U_AH" Then
                        intAH = intAH + 1
                        If intAH > 1 Then
                            If Sheets("Report").Cells(i, 12).Value = .Cells(j, 3).Value Then
                                intAH = intAH - 1
                            End If
                        End If
                        Sheets("Report").Cells(i, 12).Value = .Cells(j, 3).Value
                    ElseIf .Cells(j, 2).Value2 = "U_WH" Then
                        intWH = intWH + 1
                        If intAH > 1 Then
                            If Sheets("Report").Cells(i, 10).Value = .Cells(j, 3).Value Then
                                intAH = intAH - 1
                            End If
                        End If
                        Sheets("Report").Cells(i, 10).Value = .Cells(j, 3).Value
                    ElseIf .Cells(j, 2).Value2 = "U_WL" Then
                        intWL = intWL + 1
                        If intWL > 1 Then
                            If Sheets("Report").Cells(i, 9).Value = .Cells(j, 3).Value Then
                                intWL = intWL - 1
                            End If
                        End If
                        Sheets("Report").Cells(i, 9).Value = .Cells(j, 3).Value
                    ElseIf .Cells(j, 2).Value2 = "U_AL" Then
                        intAL = intAL + 1
                        If intAL > 1 Then
                            If Sheets("Report").Cells(i, 11).Value = .Cells(j, 3).Value Then
                                intAL = intAL - 1
                            End If
                        End If
                        Sheets("Report").Cells(i, 11).Value = .Cells(j, 3).Value
                    End If
                End If
            Next j
        End If
    Next i
End With

'Pull Address
Dim FileNum As Long
Dim TotalFile As String
Dim Lines() As String
Dim strSymbolTable As String
Set wshSymbolText = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wshSymbolText.Name = "Symbol Table"
frmSymboTable.Show
Open wsh_Path.Cells(6, 2).Value2 For Input As #1
intc = 1
intr = 1
Do Until EOF(1)
    Line Input #1, readLine
    ActiveSheet.Cells(intr, intc).Value = readLine
    intr = intr + 1
Loop
Close #1
With Sheets("Symbol Table")
    intn_SymbolTable = .Cells(Rows.Count, 1).End(xlUp).Row
    For i = intn_SymbolTable To 1 Step -1
        .Cells(i, 2).Value2 = Right(Left(.Cells(i, 1).Value2, 28), Len(Left(.Cells(i, 1).Value2, 28)) - 4)
        .Cells(i, 3).Value2 = Trim(Mid(.Cells(i, 1).Value2, 29, 2))
        .Cells(i, 4).Value2 = Trim(Mid(.Cells(i, 1).Value2, 34, 7))
        If .Cells(i, 3).Value2 = "I" Or .Cells(i, 3).Value2 = "IW" Or .Cells(i, 3).Value2 = "Q" Or .Cells(i, 3).Value2 = "QW" Then
            .Cells(i, 6).Value = 1
        End If
        If .Cells(i, 3).Value2 = "I" Or .Cells(i, 3).Value2 = "Q" Then
            .Cells(i, 5).Value2 = .Cells(i, 3).Value2 & " " & Format(.Cells(i, 4).Value2, "0.0")
        Else: .Cells(i, 5).Value2 = .Cells(i, 3).Value2 & " " & .Cells(i, 4).Value2
        End If
        If .Cells(i, 3).Value2 = "I" Then
            .Cells(i, 7).Value2 = "DI 24V"
        End If
        If .Cells(i, 3).Value2 = "Q" Then
            .Cells(i, 7).Value2 = "DO 24V"
        End If
        If .Cells(i, 6).Value = "" Then
            .Cells(i, 6).EntireRow.Delete
        End If
    Next i
    intn_SymbolTable = .Cells(Rows.Count, 1).End(xlUp).Row
    Sheets("Report").Range("A:A").Copy Destination:=wb.Sheets("Symbol Table").Range("H1")
    For i = 1 To intn_Report Step 1
        For j = 1 To intn_SymbolTable Step 1
            If .Cells(i, 8).Value2 = Trim(.Cells(j, 2).Value2) Then
                Sheets("Report").Cells(i, 2).Value2 = .Cells(j, 5).Value2
                Sheets("Report").Cells(i, 13).Value2 = .Cells(j, 7).Value2
            End If
        Next j
    Next i
End With

'Add DI Values
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws2.Name = "DI Signal"
frmCH_DI_Signals.Show
Set wb2 = Workbooks.Open(wsh_Path.Cells(11, 2).Value2)
wb2.Sheets(1).Range("K:K").Copy Destination:=wb.Sheets("DI Signal").Range("A1")
wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("DI Signal").Range("B1")
wb2.Close
'Add block to report
Dim intn_DISignal As Integer
intn_Report = Sheets("Report").Cells(Rows.Count, 1).End(xlUp).Row
intn_DISignal = Sheets("DI Signal").Cells(Rows.Count, 1).End(xlUp).Row
With Sheets("Report")
.Cells(1, 23).Value = "Digital Block"
For i = 2 To intn_Report Step 1
    If .Cells(i, 24).Value > 0 Then
        For j = 2 To intn_Signal Step 1
            If .Cells(i, 1).Value = Sheets("DI Signal").Cells(j, 1).Value Then
                .Cells(i, 19).Value = Sheets("DI Signal").Cells(j, 2).Value
                .Cells(i, 13).Value2 = "Digital"
            End If
        Next j
    End If
Next i
End With

'Add DI interconnections
Dim intn_DI As Integer
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws2.Name = "DI"
frmDI.Show
Set wb2 = Workbooks.Open(wsh_Path.Cells(12, 2).Value2)
wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("DI").Range("A1")
wb2.Sheets(1).Range("L:L").Copy Destination:=wb.Sheets("DI").Range("B1")
wb2.Close
intn_DI = Sheets("DI").Cells(Rows.Count, 1).End(xlUp).Row
With Sheets("DI")
    Sheets("Report").Cells(1, 20).Value2 = "Sig #"
    Sheets("Report").Cells(1, 21).Value2 = "Block #"
    Sheets("Report").Cells(1, 22).Value2 = "Digital Chart"
    For i = 2 To intn_Report Step 1
        For j = 1 To intn_DI Step 1
            If Sheets("Report").Cells(i, 19).Value2 = .Cells(j, 1).Value2 Then
                'Sig #
                Sheets("Report").Cells(i, 20).Value2 = Right(.Cells(j, 2).Value2, Len(.Cells(j, 2).Value2) - InStrRev(.Cells(j, 2).Value2, ".I"))
                'Block#
                Sheets("Report").Cells(i, 21).Value2 = Left(.Cells(j, 2).Value2, Len(.Cells(j, 2)) - Len(Sheets("Report").Cells(i, 20)))
                Sheets("Report").Cells(i, 21).Value2 = Right(Sheets("Report").Cells(i, 21).Value2, Len(Sheets("Report").Cells(i, 21).Value2) - InStrRev(Sheets("Report").Cells(i, 21).Value2, "\"))
                'Chart
                Sheets("Report").Cells(i, 22).Value2 = Left(.Cells(j, 2).Value2, Len(.Cells(j, 2)) - (Len(.Cells(j, 2)) - InStrRev(.Cells(j, 2).Value2, "\")))
                If Len(Sheets("Report").Cells(i, 22).Value2) > 0 Then
                    Sheets("Report").Cells(i, 22).Value2 = Left(Sheets("Report").Cells(i, 22).Value2, Len(Sheets("Report").Cells(i, 22).Value2) - 1)
                End If
                Sheets("Report").Cells(i, 22).Value2 = Right(Sheets("Report").Cells(i, 22).Value2, Len(Sheets("Report").Cells(i, 22).Value2) - InStrRev(Sheets("Report").Cells(i, 22).Value2, ".IN"))
                If Len(Sheets("Report").Cells(i, 22).Value2) > 3 Then
                    Sheets("Report").Cells(i, 22).Value2 = Right(Sheets("Report").Cells(i, 22).Value2, Len(Sheets("Report").Cells(i, 22).Value2) - 4)
                End If
            End If
        Next j
    Next i
End With

'Add DI Alarm Text
Dim intn_DIAlarm As Integer
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws2.Name = "DI Alarm"
frmDIAlarm.Show
Set wb2 = Workbooks.Open(wsh_Path.Cells(13, 2).Value2)
wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("DI Alarm").Range("A1")
wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("DI Alarm").Range("B1")
wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("DI Alarm").Range("C1")
wb2.Sheets(1).Range("H:H").Copy Destination:=wb.Sheets("DI Alarm").Range("D1")
wb2.Sheets(1).Range("N:N").Copy Destination:=wb.Sheets("DI Alarm").Range("E1")
wb2.Close
intn_DIAlarm = Sheets("DI Alarm").Cells(Rows.Count, 1).End(xlUp).Row
With Sheets("DI Alarm")
    For i = 2 To intn_Report Step 1
        For j = 1 To intn_DIAlarm Step 1
            If Sheets("Report").Cells(i, 22).Value2 = .Cells(j, 1) And Sheets("Report").Cells(i, 21).Value2 = .Cells(j, 2).Value2 Then
                If Sheets("Report").Cells(i, 20).Value2 = "I_1" Then
                    If .Cells(j, 3).Value2 = "EV_ID1" And .Cells(j, 4) = "SIG_1" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_2" Then
                    If .Cells(j, 3).Value2 = "EV_ID1" And .Cells(j, 4) = "SIG_2" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_3" Then
                    If .Cells(j, 3).Value2 = "EV_ID1" And .Cells(j, 4) = "SIG_3" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_4" Then
                    If .Cells(j, 3).Value2 = "EV_ID1" And .Cells(j, 4) = "SIG_4" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_5" Then
                    If .Cells(j, 3).Value2 = "EV_ID1" And .Cells(j, 4) = "SIG_5" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_6" Then
                    If .Cells(j, 3).Value2 = "EV_ID1" And .Cells(j, 4) = "SIG_6" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_7" Then
                    If .Cells(j, 3).Value2 = "EV_ID1" And .Cells(j, 4) = "SIG_7" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_8" Then
                    If .Cells(j, 3).Value2 = "EV_ID1" And .Cells(j, 4) = "SIG_8" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_9" Then
                    If .Cells(j, 3).Value2 = "EV_ID2" And .Cells(j, 4) = "SIG_1" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_10" Then
                    If .Cells(j, 3).Value2 = "EV_ID2" And .Cells(j, 4) = "SIG_2" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_11" Then
                    If .Cells(j, 3).Value2 = "EV_ID2" And .Cells(j, 4) = "SIG_3" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_12" Then
                    If .Cells(j, 3).Value2 = "EV_ID2" And .Cells(j, 4) = "SIG_4" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_13" Then
                    If .Cells(j, 3).Value2 = "EV_ID2" And .Cells(j, 4) = "SIG_5" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_14" Then
                    If .Cells(j, 3).Value2 = "EV_ID2" And .Cells(j, 4) = "SIG_6" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_15" Then
                    If .Cells(j, 3).Value2 = "EV_ID2" And .Cells(j, 4) = "SIG_7" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                ElseIf Sheets("Report").Cells(i, 20).Value2 = "I_16" Then
                    If .Cells(j, 3).Value2 = "EV_ID2" And .Cells(j, 4) = "SIG_8" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                End If
            End If
        Next j
    Next i
End With

'Add Rack 1 SBO Data
Dim intn_Rack As Integer
Set wshRack = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wshRack.Name = "Rack"
frmRack.Show
If Len(wsh_Path.Cells(7, 2)) > 5 Then
    Set wb2 = Workbooks.Open(wsh_Path.Cells(7, 2).Value2)
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Rack").Range("A1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("Rack").Range("B1")
    wb2.Sheets(1).Range("G:G").Copy Destination:=wb.Sheets("Rack").Range("C1")
    wb2.Close
    With Sheets("Rack")
        intn_Rack = .Cells(Rows.Count, 1).End(xlUp).Row
        intn_Report = Sheets("Report").Cells(Rows.Count, 6).End(xlUp).Row
        For i = 2 To intn_Rack Step 2
            .Cells(i, 4).Value2 = Right(.Cells(i, 1).Value2, Len(.Cells(i, 1)) - 4)
            .Cells(i, 5).Value2 = Right(.Cells(i, 2).Value2, Len(.Cells(i, 2)) - 1)
            .Cells(i, 5).Value2 = Left(.Cells(i, 5).Value2, Len(.Cells(i, 5)) - 3)
            
            'Add data to report
            Sheets("Report").Cells(intn_Report + i / 2, 1).Value2 = .Cells(i, 3).Value2
            Sheets("Report").Cells(intn_Report + i / 2, 4).Value = 1
            Sheets("Report").Cells(intn_Report + i / 2, 5).Value2 = .Cells(i, 4).Value2
            Sheets("Report").Cells(intn_Report + i / 2, 6).Value2 = .Cells(i, 5).Value2
            Sheets("Report").Cells(intn_Report + i / 2, 13).Value2 = "WR_X_SBO"
        Next i
    End With
End If

'Add Rack 1 AI1 Data
Dim intn_AI As Integer
Set wshAI = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wshAI.Name = "AI"
frmAI.Show
If Len(wsh_Path.Cells(8, 2)) > 5 Then
    Set wb2 = Workbooks.Open(wsh_Path.Cells(8, 2).Value2)
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("AI").Range("A1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("AI").Range("B1")
    wb2.Sheets(1).Range("G:G").Copy Destination:=wb.Sheets("AI").Range("C1")
    wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("AI").Range("D1")
    wb2.Close
    With Sheets("AI")
        intn_AI = .Cells(Rows.Count, 1).End(xlUp).Row
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
        intn_Report = Sheets("Report").Cells(Rows.Count, 6).End(xlUp).Row
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

'Add SOE Data
Dim intn_SOE As Integer
Set wshSOE = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wshSOE.Name = "SOE"
frmSOE.Show
If Len(wsh_Path.Cells(9, 2)) > 5 Then
    Set wb2 = Workbooks.Open(wsh_Path.Cells(9, 2).Value2)
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("SOE").Range("A1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("SOE").Range("B1")
    wb2.Sheets(1).Range("G:G").Copy Destination:=wb.Sheets("SOE").Range("C1")
    wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("SOE").Range("D1")
    wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("SOE").Range("E1")
    wb2.Close
    With Sheets("SOE")
        intn_SOE = .Cells(Rows.Count, 1).End(xlUp).Row
        intn_Report = Sheets("Report").Cells(Rows.Count, 6).End(xlUp).Row
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
Set wshSOE_Message = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wshSOE_Message.Name = "SOE Message"
frmSOE_Message.Show
If Len(wsh_Path.Cells(10, 2)) > 5 Then
    Set wb2 = Workbooks.Open(wsh_Path.Cells(10, 2).Value2)
    wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("SOE Message").Range("A1")
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("SOE Message").Range("B1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("SOE Message").Range("C1")
    wb2.Sheets(1).Range("H:H").Copy Destination:=wb.Sheets("SOE Message").Range("D1")
    wb2.Sheets(1).Range("N:N").Copy Destination:=wb.Sheets("SOE Message").Range("E1")
    wb2.Close
    With Sheets("SOE Message")
        intn_SOE_Message = .Cells(Rows.Count, 1).End(xlUp).Row
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
            ElseIf .Cells(i, 7).Value2 = 3 And .Cells(i, 8).Value2 = 1 Then
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

'Clean up workbook
'Application.DisplayAlerts = False
'    Sheets("Signal Connections").Delete
'    Sheets("HWConfig").Delete
'    Sheets("Range").Delete
'    Sheets("Alarm").Delete
'    Sheets("Symbol Table").Delete
'    Sheets("Rack").Delete
'    Sheets("AI").Delete
'    Sheets("SOE").Delete
'    Sheets("SOE Message").Delete
'    Sheets("DI Signal").Delete
'    Sheets("DI").Delete
'    Sheets("DI Alarm").Delete
'Application.DisplayAlerts = True

'Sort by Rack then Slot then Channel #
With Sheets("Report").Range("A:Z")
        .Cells.Sort Key1:=.Columns(Application.Match("Type", .Rows(1), 0)), Order1:=xlAscending, _
                    Key2:=.Columns(Application.Match("Rack", .Rows(1), 0)), Order2:=xlAscending, _
                    Key3:=.Columns(Application.Match("Slot", .Rows(1), 0)), Order2:=xlAscending, _
                    Key3:=.Columns(Application.Match("Channel", .Rows(1), 0)), Order2:=xlAscending, _
                    Orientation:=xlTopToBottom, Header:=xlYes
End With

'Clean up report
'Sheets("Report").Range("V1:Z1").EntireColumn.Delete
'Sheets("Report").Range("P1:Q1").EntireColumn.Delete

'Format report
'Sheets("Report").Range("A:S").WrapText = True
'Sheets("Report").Columns("A:A").ColumnWidth = 30
'Sheets("Report").Columns("B:B").ColumnWidth = 7.43
'Sheets("Report").Columns("C:C").ColumnWidth = 26.5
'Sheets("Report").Columns("D:E").ColumnWidth = 4.29
'Sheets("Report").Columns("F:F").ColumnWidth = 7.29
'Sheets("Report").Columns("G:I").ColumnWidth = 9.71
'Sheets("Report").Columns("J:J").ColumnWidth = 10
'Sheets("Report").Columns("K:L").ColumnWidth = 9.71
'Sheets("Report").Columns("M:M").ColumnWidth = 30.14
'Sheets("Report").Columns("N:N").ColumnWidth = 6.29
'Sheets("Report").Columns("O:O").ColumnWidth = 31
'Sheets("Report").Range("A:Z").RowHeight = 30
'Sheets("Report").Rows("1:1").RowHeight = 15
'Sheets("Report").Range("A:S").AutoFilter

With Sheets("Report")
    intn_Report = .Cells(Rows.Count, 1).End(xlUp).Row
    .Cells(1, 1).EntireColumn.Insert
    .Cells(1, 1).Value2 = "Row #"
    For i = 2 To intn_Report Step 1
        .Cells(i, 1).Value = i - 1
    Next i
End With

'Top align cells
Dim wks As Worksheet
For Each wks In Worksheets
    wks.Cells.VerticalAlignment = xlTop
    wks.Cells.HorizontalAlignment = xlLeft
Next wks

'Add data to template
Set wbTemplate = Workbooks.Open("X:\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TEMPLATE IO List Report For Extraction Tool.xlsx")
wb.Sheets("Report").Range("B:B").Copy Destination:=wbTemplate.Sheets("File Paths").Range("B1")
wb.Sheets("Report").Range("A2:Z" & intn_Report).Copy
wbTemplate.Sheets("Report").Range("A2").PasteSpecial xlPasteValues

'SaveAs
frmSaveAs.Show

Application.ScreenUpdating = True
'Freeze top row
'ThisWorkbook.Sheets("Report").Select
'With Sheets("Report")
'    .Cells(2, 1).Select
'End With
'ActiveWindow.FreezePanes = True

'With ActiveSheet.PageSetup
'.PrintArea = "$A:P"
'.PrintTitleRows = "$1:$1"
'.LeftHeader = ""
'.CenterHeader = "&8&F"
'.LeftFooter = ""
'.CenterFooter = "&8&P of &N"
'.RightFooter = ""
'.LeftMargin = Application.InchesToPoints(0.25)
'.RightMargin = Application.InchesToPoints(0.25)
'.TopMargin = Application.InchesToPoints(0.5)
'.BottomMargin = Application.InchesToPoints(0.5)
'.HeaderMargin = Application.InchesToPoints(0.3)
'.FooterMargin = Application.InchesToPoints(0.05)
'.CenterHorizontally = True
'.Orientation = xlLandscape
'.FirstPageNumber = xlAutomatic
'.FitToPagesWide = 1
'.FitToPagesTall = False
'.Zoom = False
'End With

'
End Sub


