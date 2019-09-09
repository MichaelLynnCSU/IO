Attribute VB_Name = "subIO_Name"
Function IsArrayEmpty(arr As Variant) As Boolean

Dim index As Integer

index = -1
    On Error Resume Next
        index = UBound(arr)
    On Error GoTo 0

If (index = -1) Then IsArrayEmpty = True Else IsArrayEmpty = False

End Function

'Sub SelectAllCellsInSheet(SheetName As String)
'    lastCol = Sheets(SheetName).Range("a1").End(xlToRight).Column
'    Lastrow = Sheets(SheetName).Cells(1, 1).End(xlDown).Row
'    Sheets(SheetName).Range("A1", Sheets(SheetName).Cells(Lastrow, lastCol)).Select
'    Selection.NumberFormat = "”@”"
'    Selection.Replace What:="””", Replacement:="""""", LookAt:=xlPart, _
'    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
'    ReplaceFormat:=False
'    Selection.NumberFormat = "”@”"
'End Sub

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

Sub subRunIOName()
Call subIO_Name
End Sub
Sub subIO_Name()
' IO_Name Macro
'
'read in data simple save rerun
Application.ScreenUpdating = False
ActiveSheet.Name = "HWConfig"
Set wb = ThisWorkbook
Dim intc, intr As Integer
intr = 1
intc = 1
Application.DisplayAlerts = False

Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws2.Name = "CPU"

Set wsh_Path = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsh_Path.Name = "File Paths"
    wsh_Path.Cells(1, 1).Value2 = "File Name"
    wsh_Path.Cells(1, 2).Value2 = "File Path"
    wsh_Path.Columns("A:A").ColumnWidth = 20
    wsh_Path.Columns("B:B").ColumnWidth = 100
    frmPCS7PLCWithOneOrMoreRTU.Show

Sheets("HWConfig").Select

Open wsh_Path.Cells(2, 2).Value2 For Input As #1
'keep data seperated in different columns (each group in original text file is seperated by an empty row)
Do Until EOF(1)
    Line Input #1, readline
    If readline = "" Then
        intc = intc + 1
        intr = 1
    Else
        ' Add CPU
        Dim tArray() As String
        tArray = Split(readline, " ")
        Dim strInString As String
        strInString = tArray(0)
        If Trim(strInString) = "STATION" Then
           Sheets("CPU").Cells(1, 1).Value = readline
        End If
        ActiveSheet.Cells(intr, intc).Value = readline
        intr = intr + 1
    End If
Loop
Close #1

' anadarko
Dim AVP As Boolean
AVP = False
Dim APL As Boolean
APL = blnPlaceHolder


If Sheets("CPU").Cells(1, 1).Value = "STATION S7400H , ""AS1_H""" Or Sheets("CPU").Cells(1, 1).Value = "STATION S7400H , ""AS2_H""" Then
    AVP = True
End If

Dim ColumnLetter As String
Dim ResultHWCongif() As String
Dim hwSlot As String
Dim lengthx As Integer

If AVP = True Then

'Delete any irrelevate sections (relevant sections are those with data for slots > 4)
    For j = intc To 1 Step -1
        ColumnLetter = Split(Cells(1, j).Address, "$")(1)
        If Not InStr(Cells(1, j).Value, "DPADDRESS") > 0 Then
            Columns(j).Select
            Cells(1, j).EntireColumn.Delete
        End If
    Next j

    For j = intc To 1 Step -1
        ColumnLetter = Split(Cells(1, j).Address, "$")(1)
        ResultHWCongif = Split(Cells(1, j).Value, ",")
        lengthx = UBound(ResultHWCongif, 1) - LBound(ResultHWCongif, 1) + 1
        If lengthx > 0 Then
            hwSlot = Replace(ResultHWCongif(2), """", "")
            hwSlot = Replace(hwSlot, ",", "")
            hwSlot = Trim(hwSlot)
            Columns(j).Select
            If hwSlot = ("SLOT 1") Then
                Cells(1, j).EntireColumn.Delete
            End If
        End If
    Next j

    For j = intc To 1 Step -1
        ColumnLetter = Split(Cells(1, j).Address, "$")(1)
        ResultHWCongif = Split(Cells(1, j).Value, ",")
        lengthx = UBound(ResultHWCongif, 1) - LBound(ResultHWCongif, 1) + 1
        If lengthx > 0 Then
            hwSlot = Replace(ResultHWCongif(2), """", "")
            hwSlot = Replace(hwSlot, ",", "")
            hwSlot = Trim(hwSlot)
            Columns(j).Select
            If hwSlot = ("SLOT 2") Then
                Cells(1, j).EntireColumn.Delete
            End If
        End If
    Next j

    For j = intc To 1 Step -1
        ColumnLetter = Split(Cells(1, j).Address, "$")(1)
        ResultHWCongif = Split(Cells(1, j).Value, ",")
        lengthx = UBound(ResultHWCongif, 1) - LBound(ResultHWCongif, 1) + 1
        If lengthx > 0 Then
            hwSlot = Replace(ResultHWCongif(2), """", "")
            hwSlot = Replace(hwSlot, ",", "")
            hwSlot = Trim(hwSlot)
            Columns(j).Select
            If hwSlot = ("SLOT 3") Then
                Cells(1, j).EntireColumn.Delete
            End If
        End If
    Next j

Else:

    'Delete any irrelevate sections (relevant sections are those with data for slots > 4)
    For j = intc To 1 Step -1
        ColumnLetter = Split(Cells(1, j).Address, "$")(1)
        If Not InStr(Cells(1, j).Value, "DPADDRESS") > 0 Then
            Columns(j).Select
            Cells(1, j).EntireColumn.Delete
        End If
    Next j
    
    For j = intc To 1 Step -1
        ColumnLetter = Split(Cells(1, j).Address, "$")(1)
        If InStr(Cells(1, j).Value, " SLOT 1") > 0 Then
             Columns(j).Select
            Cells(1, j).EntireColumn.Delete
        End If
    Next j
    
    For j = intc To 1 Step -1
        ColumnLetter = Split(Cells(1, j).Address, "$")(1)
        If InStr(Cells(1, j).Value, " SLOT 2") > 0 Then
            Columns(j).Select
            Cells(1, j).EntireColumn.Delete
        End If
    Next j
    
    For j = intc To 1 Step -1
        ColumnLetter = Split(Cells(1, j).Address, "$")(1)
        If InStr(Cells(1, j).Value, " SLOT 3") > 0 Then
            Columns(j).Select
            Cells(1, j).EntireColumn.Delete
        End If
    Next j

End If
     
'Create template for report/summary
Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
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
    .Cells(1, 26).Value = "AI-Range"
End With


Dim mySignals() As Variant
mySignals = Array(Null)
ReDim Preserve mySignals(UBound(mySignals) + 1) 'Add next array element
mySignals(UBound(mySignals)) = Null         'Assign the array element
Dim intn_Data, intn_Report, intk, intai As Integer
Dim strchannel, strsymbol, strcomment, strtype, straitype, strrack, strslot, strDig, strRTU, strIOCard As String
Dim chnlSLotBreak() As String
Dim IOcardType() As String
Dim EmptyArrayCheck As Boolean
EmptyArrayCheck = False
'Extract data and move to report
With Sheets("HWConfig")
    For j = intc To 1 Step -1
        intn_Data = .Cells(Rows.count, j).End(xlUp).Row
        intn_Report = Sheets("Report").Cells(Rows.count, 6).End(xlUp).Row
        intk = 0
        intai = 0
        'Digital?
        strDig = InStr(1, .Cells(1, j), "DI")
        'Extract rack #
        strrack = Mid(.Cells(1, j), InStr(1, .Cells(1, j), ",") + Len("DPADDRESS") + 3, 2)
        'Extract slot #
        strslot = Mid(.Cells(1, j), InStr(1, .Cells(1, j), ",") + Len("DPADDRESS") + 3 + Len(strrack) + 3 + Len("SLOT"), 1)
        
        If AVP Then
            strrack = Replace(strrack, ",", "")
            Dim Result() As String
            Dim TextStrng As String
            TextStrng = ThisWorkbook.Sheets("HWConfig").Cells(1, j).Value
            Result() = Split(TextStrng, ",")
            Dim lengthy As Integer
            lengthy = UBound(Result, 1) - LBound(Result, 1) + 1
            
            If lengthy > 4 Then
                strIOCard = Result(4)
                IOcardType = Split(strIOCard, "_")
                If Trim(Replace(IOcardType(0), """", "")) = "VIM" Or Trim(Replace(IOcardType(0), """", "")) = "SDM" Or Trim(Replace(IOcardType(0), """", "")) = "SAM" Or Trim(Replace(IOcardType(0), """", "")) = "EAM" Then
                    chnlSLotBreak = Split(strIOCard, "_")
                    strrack = Mid(chnlSLotBreak(1), 2, 2)
                End If
            End If
            
             If lengthy > 0 Then
                Dim new_slot As String
                 new_slot = Result(2)
                 Result() = Split(new_slot, " ")
                 lengthy = UBound(Result, 1) - LBound(Result, 1) + 1
                 If lengthy > 2 Then
                    new_slot = Result(2)
                    strslot = new_slot
                    If Trim(Replace(IOcardType(0), """", "")) = "VIM" Or Trim(Replace(IOcardType(0), """", "")) = "SDM" Or Trim(Replace(IOcardType(0), """", "")) = "SAM" Or Trim(Replace(IOcardType(0), """", "")) = "EAM" Then
                        strslot = Mid(chnlSLotBreak(1), 5, 2)
                    End If
                 End If
             End If
             
        End If
        
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

                Dim chanAnadarkoOffset As Integer
             
                'Add channel #, symbol and comment to report
                If (Not IsArrayEmpty(IOcardType)) Then
                    If Trim(Replace(IOcardType(0), """", "")) = "VIM" Or Trim(Replace(IOcardType(0), """", "")) = "SDM" Or Trim(Replace(IOcardType(0), """", "")) = "SAM" Or Trim(Replace(IOcardType(0), """", "")) = "EAM" Then
                       Sheets("Report").Cells(intk + intn_Report, 6).Value2 = (strchannel + 1)
                    Else:
                       Sheets("Report").Cells(intk + intn_Report, 6).Value2 = strchannel
                    End If
                Else:
                    Sheets("Report").Cells(intk + intn_Report, 6).Value2 = strchannel
                End If
                
                Sheets("Report").Cells(intk + intn_Report, 1).Value2 = strsymbol
                ReDim Preserve mySignals(UBound(mySignals) + 1)   'Add next array element
                mySignals(UBound(mySignals)) = strsymbol         'Assign the array element
                Sheets("Report").Cells(intk + intn_Report, 3).Value2 = strcomment
                Sheets("Report").Cells(intk + intn_Report, 4).Value2 = strrack
                Sheets("Report").Cells(intk + intn_Report, 5).Value2 = strslot
                Sheets("Report").Cells(intk + intn_Report, 13).Value2 = strIOCard
                Sheets("Report").Cells(intk + intn_Report, 24).Value2 = strDig
                If AVP = False Then
                    If strDig > 0 Then
                        Sheets("Report").Cells(intk + intn_Report, 13).Value2 = "Digital"
                    End If
                End If
                'Add RTU/ET200M to report
                Sheets("Report").Cells(intk + intn_Report, 25).Value2 = strRTU
            End If
        Next i
    Next j
End With


'Pull over block
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws2.Name = "Signal Connections"
'frmCH_AI_Signals.Show
Set wb2 = Workbooks.Open(wsh_Path.Cells(3, 2).Value2)
wb2.Sheets(1).Range("K:K").Copy Destination:=wb.Sheets("Signal Connections").Range("A1")
wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Signal Connections").Range("B1")
wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Signal Connections").Range("C1")
wb2.Close
'Add block to report

If AVP Then
 Sheets("Signal Connections").Columns("B:B").Cut Destination:=Columns("G:G")
    Sheets("Signal Connections").Columns("C:C").Cut Destination:=Columns("B:B")
    Sheets("Signal Connections").Columns("G:G").Cut Destination:=Columns("C:C")
    Sheets("Signal Connections").Cells(1, 2).Value2 = "Block"
    Sheets("Signal Connections").Cells(1, 3).Value2 = "Chart"
End If
    
Dim strLoc, strBlock As String
Dim intn_Signal, intn_Alarm, intmatch As Integer
intn_Report = Sheets("Report").Cells(Rows.count, 6).End(xlUp).Row
intn_Signal = Sheets("Signal Connections").Cells(Rows.count, 1).End(xlUp).Row
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

       
Dim mySignalsLen As Integer
mySignalsLen = UBound(mySignals) - LBound(mySignals) + 1

'Pull Over Range Values
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws2.Name = "Range"
'frmRanges.Show
If Not Len(wsh_Path.Cells(4, 2)) > 0 Then
'    frmRanges.Show
End If

Set wb2 = Workbooks.Open(wsh_Path.Cells(4, 2).Value2)
wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Range").Range("A1")
wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("Range").Range("B1")
wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("Range").Range("C1")
wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Range").Range("D1")
wb2.Sheets(1).Range("L:L").Copy Destination:=wb.Sheets("Range").Range("E1")
wb2.Close
'sort by block then I/O Name

If AVP Then
 Sheets("Range").Columns("A:A").Cut Destination:=Columns("G:G")
    Sheets("Range").Columns("D:D").Cut Destination:=Columns("A:A")
    Sheets("Range").Columns("G:G").Cut Destination:=Columns("D:D")
    Sheets("Range").Cells(1, 1).Value2 = "Block"
    Sheets("Range").Cells(1, 4).Value2 = "Chart"
Else

    Dim rows_Range2 As Integer: rows_Range2 = Sheets("Range").UsedRange.Rows.count
    Do While rows_Range2 > 0
        Dim curr_IO2 As String: curr_IO2 = Sheets("Range").Range("B" & rows_Range2).Value
        Dim curr_IO3 As String: curr_IO3 = Sheets("Range").Range("A" & rows_Range2).Value
        If curr_IO2 <> "Scale.High" And curr_IO2 <> "Scale.Low" And curr_IO2 <> "PV_Out" And curr_IO2 <> "VHRANGE" And curr_IO2 <> "VLRANGE" And curr_IO2 <> "V" Or curr_IO3 = "2" Then
            Rows(rows_Range2).Delete
        End If
        rows_Range2 = rows_Range2 - 1
    Loop
    
End If
   
            
intn_Range = Sheets("Range").Cells(Rows.count, 1).End(xlUp).Row
'Add new block to report
Range("F2").Select
ActiveCell.FormulaR1C1 = "=MID(RC[-1],FIND(RC[-2],RC[-1])+LEN(RC[-2])+1,LEN(RC[-1])-FIND(RC[-2],RC[-1])+1)"
Range("G2").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-1],FIND("".PV"",RC[-1])-1)"
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
On Error Resume Next
    For i = 2 To intn_Report Step 1
        If Not Len(.Cells(i, 16).Value) > 0 Then
            .Cells(i, 7).Value = ""
            .Cells(i, 8).Value = ""
        Else:
            If IsError(Sheets("Range").Cells(.Cells(i, 7).Value + 2, 7)) = True Then
                .Cells(i, 17).Value = ""
            Else:
                .Cells(i, 17).Value = Sheets("Range").Cells(.Cells(i, 7).Value + 2, 7).Value
            End If
                .Cells(i, 7).Value = Sheets("Range").Cells(.Cells(i, 7).Value, 3).Value
                .Cells(i, 8).Value = Sheets("Range").Cells(.Cells(i, 8).Value + 1, 3).Value
        End If
    Next i
Resume
End With
    

'Add DI Values
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws2.Name = "DI Signal"
'frmCH_DI_Signals.Show
Set wb2 = Workbooks.Open(wsh_Path.Cells(11, 2).Value2)

If AVP Then
    wb2.Sheets(1).Range("K:K").Copy Destination:=wb.Sheets("DI Signal").Range("A1")
    wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("DI Signal").Range("B1")
Else:
    wb2.Sheets(1).Range("K:K").Copy Destination:=wb.Sheets("DI Signal").Range("A1")
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("DI Signal").Range("B1")
End If
wb2.Close

If AVP Then
    Sheets("DI Signal").Cells(1, 2).Value2 = "Block"
End If

If AVP = True Then
Dim diCHeck As Integer
diCHeck = Sheets("DI Signal").UsedRange.Rows.count
    For i = 1 To intn_Report Step 1
        For j = 1 To diCHeck Step 1
            If Trim(Sheets("Report").Cells(i, 1).Value2) = Trim(Sheets("DI Signal").Cells(j, 1).Value2) Then
                Sheets("Report").Cells(i, 24).Value2 = "1"
                Sheets("Report").Cells(i, 16).Value2 = Sheets("DI Signal").Cells(j, 2).Value2
            End If
        Next j
    Next i
Else:
'Add block to report
Dim intn_DISignal As Integer
intn_Report = Sheets("Report").Cells(Rows.count, 1).End(xlUp).Row
intn_DISignal = Sheets("DI Signal").Cells(Rows.count, 1).End(xlUp).Row
    With Sheets("Report")
    .Cells(1, 23).Value = "Digital Block"
    For i = 2 To intn_Report Step 1
        If .Cells(i, 24).Value > 0 Then
            For j = 2 To intn_Signal Step 1
                If .Cells(i, 1).Value = Sheets("DI Signal").Cells(j, 1).Value Then
                        .Cells(i, 19).Value = Sheets("DI Signal").Cells(j, 2).Value
                    If AVP = False Then
                        .Cells(i, 13).Value2 = "Digital"
                    End If
                End If
            Next j
        End If
    Next i
    End With
End If

'Pull Address
Dim FileNum As Long
Dim TotalFile As String
Dim Lines() As String
Dim strSymbolTable As String
Set wshSymbolText = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wshSymbolText.Name = "Symbol Table"
'frmSymboTable.Show
Open wsh_Path.Cells(6, 2).Value2 For Input As #1
intc = 1
intr = 1
Do Until EOF(1)
    Line Input #1, readline
    ActiveSheet.Cells(intr, intc).Value = readline
    intr = intr + 1
Loop
Close #1
With Sheets("Symbol Table")
    intn_SymbolTable = .Cells(Rows.count, 1).End(xlUp).Row
    For i = intn_SymbolTable To 1 Step -1
        .Cells(i, 2).Value2 = Right(Left(.Cells(i, 1).Value2, 28), Len(Left(.Cells(i, 1).Value2, 28)) - 4)
        .Cells(i, 3).Value2 = Trim(Mid(.Cells(i, 1).Value2, 29, 2))
        .Cells(i, 4).Value2 = Trim(Mid(.Cells(i, 1).Value2, 34, 7))
        If .Cells(i, 3).Value2 = "I" Or .Cells(i, 3).Value2 = "IW" Or .Cells(i, 3).Value2 = "Q" Or .Cells(i, 3).Value2 = "QW" Then
            .Cells(i, 6).Value = 1
        End If
        If .Cells(i, 3).Value2 = "I" Or .Cells(i, 3).Value2 = "Q" Then
            .Cells(i, 5).Value2 = .Cells(i, 3).Value2 & " " & Format(.Cells(i, 4).Value2, "0.0")
        Else:
            .Cells(i, 5).Value2 = .Cells(i, 3).Value2 & " " & .Cells(i, 4).Value2
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
 End With
 
If AVP = True Then
intn_SymbolTable = Sheets("Symbol Table").UsedRange.Rows.count
    For i = 1 To intn_Report Step 1
         For j = 1 To intn_SymbolTable Step 1
                If Trim(Sheets("Report").Cells(i, 1).Value2) = Trim(Sheets("Symbol Table").Cells(j, 2).Value2) Then
                    Sheets("Report").Cells(i, 2).Value2 = Trim(Sheets("Symbol Table").Cells(j, 5).Value2)
                    If Trim(Sheets("Symbol Table").Cells(j, 3).Value2) = "I" Then
                       Sheets("Report").Cells(i, 13).Value2 = "DI 24V" & Sheets("Report").Cells(i, 13).Value2
                    ElseIf Trim(Sheets("Symbol Table").Cells(j, 3).Value2) = "Q" Then
                       Sheets("Report").Cells(i, 13).Value2 = "DO 24V" & Sheets("Report").Cells(i, 13).Value2
                    ElseIf Trim(Sheets("Symbol Table").Cells(j, 3).Value2) = "IW" Then
                       Sheets("Report").Cells(i, 13).Value2 = "AI" & Sheets("Report").Cells(i, 13).Value2
                    ElseIf Trim(Sheets("Symbol Table").Cells(j, 3).Value2) = "QW" Then
                        Sheets("Report").Cells(i, 13).Value2 = "AO" & Sheets("Report").Cells(i, 13).Value2
                    Else:
                       Sheets("Report").Cells(i, 13).Value2 = "None"
                    End If
                End If
            Next j
      Next i
Else:
    intn_SymbolTable = Sheets("Symbol Table").Cells(Rows.count, 1).End(xlUp).Row
    Sheets("Report").Range("A:A").Copy Destination:=wb.Sheets("Symbol Table").Range("H1")
    For i = 1 To intn_Report Step 1
        For j = 1 To intn_SymbolTable Step 1
            If Sheets("Symbol Table").Cells(i, 8).Value2 = Trim(Sheets("Symbol Table").Cells(j, 2).Value2) Then
                Sheets("Report").Cells(i, 2).Value2 = Sheets("Symbol Table").Cells(j, 5).Value2
                Sheets("Report").Cells(i, 13).Value2 = Sheets("Symbol Table").Cells(j, 7).Value2
            End If
        Next j
    Next i
    End If
    


'Pull Over Alarm Values
Dim intz As Integer
Dim intn_Working As Integer
Dim strAlarmBlock As String
Dim intAH As Integer
Dim intWH As Integer
Dim intWL As Integer
Dim intAL As Integer
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws2.Name = "Alarm"
'frmAlarm.Show


'Pull columns from Nickajack_Plant_NJH_Meas_Mon_Alarming
Set wb2 = Workbooks.Open(wsh_Path.Cells(5, 2).Value2)
If AVP Then
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Alarm").Range("D1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("Alarm").Range("B1")
    wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("Alarm").Range("C1")
    wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Alarm").Range("A1")
    Else:
    wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Alarm").Range("A1")
    wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("Alarm").Range("B1")
    wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("Alarm").Range("C1")
    wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Alarm").Range("D1")
End If
wb2.Close

    Dim rows_Range4 As Integer: rows_Range4 = Sheets("Alarm").UsedRange.Rows.count
    Do While rows_Range4 > 0
        Dim curr_IO4 As String: curr_IO4 = Sheets("Alarm").Range("B" & rows_Range4).Value
        If curr_IO4 <> "PV_AH_Lim" And curr_IO4 <> "PV_WH_Lim" And curr_IO4 <> "PV_WL_Lim" And curr_IO4 <> "PV_AL_Lim" And curr_IO4 <> "U_AH" And curr_IO4 <> "U_WH" And curr_IO4 <> "U_WL" And curr_IO4 <> "U_AL" Then
            Sheets("Alarm").Rows(rows_Range4).Delete
        End If
        rows_Range4 = rows_Range4 - 1
    Loop
    
    
                            
                                  
If AVP Then
    Sheets("Alarm").Cells(1, 1).Value2 = "Block"
    Sheets("Alarm").Cells(1, 4).Value2 = "Chart"
    
End If
           

'Start Part A

Dim current_symbol As String
Dim current_signal As String
Dim current_signal_block As String
Dim current_signal_chart As String
Dim rows_Signal As Integer: rows_Signal = Sheets("Signal Connections").UsedRange.Rows.count
    
For i = 2 To (mySignalsLen - 1) Step 1
    For j = 2 To rows_Signal Step 1
        current_symbol = mySignals(i)
        current_signal = Sheets("Signal Connections").Cells(j, 1).Value2
        
        'search for a signal and symbol match
        If current_symbol = current_signal Then
            current_signal_block = Sheets("Signal Connections").Cells(j, 2).Value2
            current_signal_chart = Sheets("Signal Connections").Cells(j, 3).Value2
                      
            'Start Part B
                                            
            Dim current_range_block As String
            Dim current_range_chart As String
            Dim current_range_interconnetion_block As String
            Dim current_range_interconnetion_block_U As String
            Dim row_range As Integer
            row_range = Sheets("Range").UsedRange.Rows.count
                              
            For k = 2 To row_range Step 1
'                'get a range_block(i) and range_chart value
                current_range_block = Sheets("Range").Cells(k, 1).Value2
                current_range_chart = Sheets("Range").Cells(k, 4).Value2
'
'
'                'search for a symbol and range match
                If current_signal_block = current_range_block Then
                If current_signal_chart = current_range_chart Then
                If APL Then
                    current_range_interconnetion_block = Sheets("Range").Cells(k, 7).Value2
                    current_range_interconnetion_block_U = current_range_interconnetion_block
                Else
                        current_range_interconnetion_block = Sheets("Range").Cells(k, 5).Value2
    
                        'Start Part C
    
                        Dim intEndPos As Integer
                        Dim intStartPos As Integer
    
    '                     Start new String parsing alogrithm
    '
    '                    Separate the block from the string
    '                    the .U is a marker in the string to help find where the block name is located. Start from the end and search for .U and check if .U found
    '                      at end or in middle of interconnect string
    '                    compistate for the new plant
                        If AVP Then
                            intEndPos = InStrRev(current_range_interconnetion_block, ".PV")
                            If intEndPos > 0 Then
                                Dim Result2() As String
                                Result2() = Split(current_range_interconnetion_block, "\")
                                current_range_interconnetion_block_U = Result2(0)
                                current_range_interconnetion_block_U = Replace(current_range_interconnetion_block_U, """", "")
                            End If
                        Else:
                            intEndPos = InStrRev(current_range_interconnetion_block, ".U")
                        End If
    
                        'if .U is found and it is either at the very end or in the middle but has a " right after it then find the block name
                        If intEndPos > 0 Then
                            If (intEndPos = Len(current_range_interconnetion_block) - 1) Then  'check if .U at the end of the string
                            '                                Debug.Print "It's at the end"
                            ElseIf Asc(Mid(current_range_interconnetion_block, intEndPos + 2, 1)) = 34 Then  'if .U in the middle check if there is double quote after the U (ascii of " is 34)
                            '                                Debug.Print "It's in the middle and has double quote after the U"
                            Else
                            '                                Debug.Print "string not found"
                            intEndPos = 0  'set to 0 so will not try to find block name
                            End If
                            If intEndPos > 0 Then 'the .U was found so now get the block name from the string
                                intStartPos = InStrRev(current_range_interconnetion_block, "\", intEndPos)
                                current_range_interconnetion_block_U = Mid(current_range_interconnetion_block, intStartPos + 1, intEndPos - intStartPos - 1)
                                '                                Debug.Print "Found string:"; current_range_interconnetion_block_U
                            End If
                            ' End String parsing alogrithm
                        End If
                    End If
    

                    Dim intRowsAlarm As Integer
                    Dim strCurrentAlarmBlock As String
                    Dim strCurrentAlarmChart As String

                    Dim IOTag As String
                    intRowsAlarm = Sheets("Alarm").UsedRange.Rows.count

                For m = 2 To intRowsAlarm Step 1

                    'get a alarm_block(i) and  alarm_chart value
                    strCurrentAlarmBlock = Sheets("Alarm").Cells(m, 1).Value2
                    strCurrentAlarmChart = Sheets("Alarm").Cells(m, 4).Value2
    
                    IOTag = Sheets("Alarm").Cells(m, 2).Value2
    
                    'search for a range and alarm match
                    If strCurrentAlarmBlock = current_range_interconnetion_block_U Then
        
                        If IOTag = "U_AH" Or IOTag = "PV_AH_Lim" Then
                            '                                                  Debug.Print "intAlarmAH: " & intAlarmAH
                            Sheets("Alarm").Cells(m, 3).Copy Sheets("Report").Cells(i, 12)
                            End If
                        If IOTag = "U_WH" Or IOTag = "PV_WH_Lim" Then
                            '                                                  Debug.Print "intAlarmWH: " & intAlarmWH
                            Sheets("Alarm").Cells(m, 3).Copy Sheets("Report").Cells(i, 10)
                        End If
                        If IOTag = "U_WL" Or IOTag = "PV_WL_Lim" Then
                            '                                                  Debug.Print "intAlarmWL: " & intAlarmWL
                            Sheets("Alarm").Cells(m, 3).Copy Sheets("Report").Cells(i, 9)
                        End If
                        If IOTag = "U_AL" Or IOTag = "PV_AL_Lim" Then
                            '                                                  Debug.Print "intAlarmAL: " & intAlarmAL
                            Sheets("Alarm").Cells(m, 3).Copy Sheets("Report").Cells(i, 11)
                        End If
                    End If
                Next m
                
                End If
'                'End Part C
                End If
            Next k
            'End Part B
        
        End If
    Next j
Next i
'End Part A
  
If AVP = False Then '-----------------------------------------------skip

Dim rows_symbol_Report As Integer: rows_symbol_Report = Sheets("Report").UsedRange.Rows.count
Dim cols_HWConfig_T As Integer: cols_HWConfig_T = Sheets("HWConfig").UsedRange.Columns.count
Dim rows_HWConfig_T As Integer: rows_HWConfig_T = Sheets("HWConfig").UsedRange.Rows.count

  For q = 2 To (mySignalsLen - 1) Step 1
    'Debug.Print "CURRENT COUNT ", q
    ' after symbol is matched with its messafwes set it nack to -1
    Dim digCheck As Boolean:  digCheck = False
    
    Dim target_channel As String
    target_channel = "-1"
    
    Dim target_message As String
    target_message = ""
    
    Dim target_address_AI As String
    target_address = ""

    
    Dim symbol_from_report As String
    symbol_from_report = mySignals(q)
                            
    For i = 2 To cols_HWConfig_T Step 1
        For j = 1 To rows_HWConfig_T Step 1
        'start part A of algorithm
        
        Dim HWConfig_line As String
        HWConfig_line = Sheets("HWConfig").Cells(j, i).Value2
          
          'get signal from HWConfig and match it to symbol from report
          If InStr(HWConfig_line, ",") > 0 Then
          Dim LArray() As String
          LArray = Split(HWConfig_line, ",")
            If IsEmpty(LArray) Then
            ' do nothing
            Else
          ' UBound(LArray, 1) gives the upper limit of the first dimension, which is 5.
            x = UBound(LArray, 1) - LBound(LArray, 1) + 1
                If x > 4 Then
                ' remove the quotes for the comparison
                 Dim cleanSTRAI As String
                 cleanSTRAI = Replace(LArray(0), """", "")
                 Dim findSpaceAdd As Integer: findSpaceAdd = InStr(cleanSTRAI, " ")
                
                  If Mid(cleanSTRAI, 1, findSpaceAdd - 1) = "DPSUBSYSTEM" Then
                    target_address_AI = LArray(4)
                  End If
                End If
            End If
                                  
                  
            intEndPos = InStr(HWConfig_line, ",")
            intStartPos = 1
            Dim signal_from_HWCONFIG As String
            signal_from_HWCONFIG = ""
            signal_from_HWCONFIG = Mid(HWConfig_line, intStartPos, intEndPos - 1)
            Dim remander_current_symbol_T As String
            remander_current_symbol = Mid(HWConfig_line, intEndPos + 2, Len(HWConfig_line))
                  
                                         
            If Trim(signal_from_HWCONFIG) = Trim("SYMBOL  I") Or Trim(signal_from_HWCONFIG) = Trim("SYMBOL  O") Then
                 Dim strET200 As String
                 strET200 = Sheets("Report").Cells(q, 25).Value2
                 
                 If Len(strET200) < 1 Then
                      ' repair the rackk 33 et200 errors
                      Sheets("report").Cells(q, 25).Value = "ET200M"
                 End If
                                                                                              
                intEndPos = InStr(remander_current_symbol, ",")
                intStartPos = 1
                Dim current_channel_T As String
                current_channel_T = Mid(remander_current_symbol, intStartPos, intEndPos - 1)
                remander_current_symbol = Mid(remander_current_symbol, intEndPos + 2, Len(HWConfig_line))
                intEndPos = InStr(remander_current_symbol, ",")
                intStartPos = 1
                Dim current_signal_T As String
                current_signal_T = Mid(remander_current_symbol, intStartPos + 1, intEndPos - 3)
                 
                 If current_channel_T = "0" Or current_channel_T = "1" Then
                    current_channel_T = "0"
                 ElseIf current_channel_T = "2" Or current_channel_T = "3" Then
                    current_channel_T = "1"
                 ElseIf current_channel_T = "4" Or current_channel_T = "5" Then
                    current_channel_T = "2"
                 ElseIf current_channel_T = "6" Or current_channel_T = "7" Then
                    current_channel_T = "3"
                 Else
                    'Debug.Print "WE GOT HERE, OUT OF SIGNAL RANGE ", current_channel_T
                 End If
                                                                                                                
                If Trim(current_signal_T) = Trim(symbol_from_report) Then
                Trim (Replace(target_address_AI, """", ""))
                    If Trim(Replace(target_address_AI, """", "")) = "DO16xDC24V/0.5A" Or Trim(Replace(target_address_AI, """", "")) = "DI16xDC24V" Then
                         digCheck = True
                    Exit For
                End If
                    target_channel = current_channel_T
            End If
           End If
    End If
                  
                              
                
        ' end part A of algorithm
        
        ' part B of parse String algorithm
        
 
            'get symbol
            If InStr(HWConfig_line, ",") > 0 Then

                  intEndPos = InStr(HWConfig_line, ",")
                  intStartPos = 1
                  Dim current_symbol_T2
                  current_symbol_T2 = Mid(HWConfig_line, intStartPos, intEndPos - 1)
                  'Debug.Print "current_symbol_T2 for AI_type", Trim(current_symbol_T2)
                          
         
                            
                If Trim(current_symbol_T2) = Trim("AI_TYPE") Or Trim(current_symbol_T2) = Trim("AO_TYPE") Then
                        
                       If target_channel <> "-1" Then
                          'Debug.Print "TARGET CHANNEL ", target_channel
                    
                          'get AI_type
                          intEndPos = InStr(HWConfig_line, ",")
                          intStartPos = 1
                          Dim current_AI_4_type As String
                          current_AI_4_type = Mid(HWConfig_line, intStartPos, intEndPos - 1)
                          'Debug.Print current_AI_4_type
                         

                            'store the rest of the string
                            Dim remander_AI_4_type As String
                            remander_AI_4_type = Mid(HWConfig_line, intEndPos + 2, Len(HWConfig_line))
                           'Debug.Print remander_AI_4_type


                            'get AI_ID_type
                            intEndPos = InStr(remander_AI_4_type, ",")
                            intStartPos = 1
                            Dim current_ID_AI_4_type As String
                            current_ID_AI_4_type = Mid(remander_AI_4_type, intStartPos, intEndPos - 1)
                            'Debug.Print "check current ID ", current_ID_AI_4_type

                            'store the rest of the string
                            Dim remander_ID_Range_4_type
                            remander_ID_Range_4_type = Mid(remander_AI_4_type, intEndPos + 2, Len(remander_AI_4_type))
                           'Debug.Print "remander_ID_Range_4_type, "; remander_ID_Range_4_type

                            'get AI_channel_type
                            intEndPos = InStr(remander_ID_Range_4_type, ",")
                            intStartPos = 1
                            Dim current_channel_AI_4_type
                            current_channel_AI_4_type = Mid(remander_ID_Range_4_type, intStartPos, intEndPos - 1)
                          'Debug.Print "get AI_channel_type ", Trim(current_channel_AI_4_type)

                            'store the rest of the string, thing left is the messages
                            Dim remander_messages_AI_4_type
                            remander_messages_AI_4_type = Mid(remander_ID_Range_4_type, intEndPos + 2, Len(remander_ID_Range_4_type))
                           'Debug.Print "current sybmol AI messages", remander_messages_AI_4_type
                              
                                  If Trim(target_channel) = Trim(current_channel_AI_4_type) Then
                                       target_message = remander_messages_AI_4_type
                                   End If
                 
                      End If

               End If
            
        End If


    ' end part B of algorithm
    
    ' part C of String alogrithm

                'get symbol
                If InStr(HWConfig_line, ",") > 0 Then

                    intEndPos = InStr(HWConfig_line, ",")
                    intStartPos = 1
                    Dim current_symbol_T
                    current_symbol_T = Mid(HWConfig_line, intStartPos, intEndPos - 1)
                    
                    
                    
                     If Trim(current_symbol_T) = Trim("AI_RANGE") Or Trim(current_symbol_T) = Trim("AO_RANGE") Then
                        'Debug.Print "FOUND RANGE LINE ", Trim(HWConfig_line)
                        If target_channel <> "-1" Then
                      'Debug.Print "CURRENT RANGE CHANNEL ", target_channel


                        intEndPos = InStr(HWConfig_line, ",")
                        intStartPos = 1
                        Dim current_Range_4_type2
                        current_Range_4_type2 = Mid(HWConfig_line, intStartPos, intEndPos - 1)
                       'Debug.Print current_Range_4_type

                        'store the rest of the string
                        Dim remander_current_symbol2
                        remander_current_symbol2 = Mid(HWConfig_line, intEndPos + 2, Len(HWConfig_line))
                        'Debug.Print remander_current_symbol


                        'get AI_ID_type
                        intEndPos = InStr(remander_current_symbol2, ",")
                        intStartPos = 1
                        Dim AI_ID_type As String
                        AI_ID_type = Mid(remander_current_symbol2, intStartPos, intEndPos - 1)

                        'store the rest of the string
                        remander_current_symbol2 = Mid(remander_current_symbol2, intEndPos + 2, Len(remander_current_symbol2))

                        'get AI_channel_type
                        intEndPos = InStr(remander_current_symbol2, ",")
                        intStartPos = 1
                        Dim current_channel_Range_4_type2
                        current_channel_Range_4_type2 = Mid(remander_current_symbol2, intStartPos, intEndPos - 1)

                        'store the rest of the string, thing left is the messages
                        Dim remander_messages_Range_4_type2
                        remander_messages_Range_4_type2 = Mid(remander_current_symbol2, intEndPos + 2, Len(remander_current_symbol2))

                              If Trim(target_channel) = Trim(current_channel_Range_4_type2) Then
                                  target_message = remander_messages_Range_4_type2 & target_message & target_address_AI
                                  target_channel = ""
                              End If

                     End If
                 End If
            End If

     ' end C of String alogrithm
                 
              ' reset current_symbol_T
              current_symbol_T = ""
        Next j
    If digCheck Then Exit For
      Next i
      If Len(target_message) > 1 Then
                Dim TxtRng  As Range
                Set TxtRng = Sheets("Report").Cells(q, 13)
                TxtRng.Value = target_message
                target_message = ""
      End If
  Next q
End If '----------------------------------------------- end skip

'Add DI interconnections
Dim intn_DI As Integer
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws2.Name = "DI"
'frmDI.Show
Set wb2 = Workbooks.Open(wsh_Path.Cells(12, 2).Value2)

If AVP Then
        wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("DI").Range("A1")
Else
        wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("DI").Range("A1")
End If

wb2.Sheets(1).Range("L:L").Copy Destination:=wb.Sheets("DI").Range("B1")
wb2.Close
intn_DI = Sheets("DI").Cells(Rows.count, 1).End(xlUp).Row
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
Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws2.Name = "DI Alarm"
'frmDIAlarm.Show
Set wb2 = Workbooks.Open(wsh_Path.Cells(13, 2).Value2)
wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("DI Alarm").Range("A1")
wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("DI Alarm").Range("B1")
wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("DI Alarm").Range("C1")
wb2.Sheets(1).Range("H:H").Copy Destination:=wb.Sheets("DI Alarm").Range("D1")

If AVP Then
        wb2.Sheets(1).Range("E:E").Copy Destination:=wb.Sheets("DI Alarm").Range("E1")
Else
        wb2.Sheets(1).Range("N:N").Copy Destination:=wb.Sheets("DI Alarm").Range("E1")
End If

wb2.Close
intn_DIAlarm = Sheets("DI Alarm").Cells(Rows.count, 1).End(xlUp).Row

If AVP Then
With Sheets("DI Alarm")
    For i = 2 To intn_Report Step 1
        For j = 1 To intn_DIAlarm Step 1
            If Sheets("Report").Cells(i, 16).Value2 = .Cells(j, 1) Then
                    If .Cells(j, 3).Value2 = "MsgEvId1" And .Cells(j, 4) = "SIG_1" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                    If .Cells(j, 3).Value2 = "MsgEvId1" And .Cells(j, 4) = "SIG_2" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                    If .Cells(j, 3).Value2 = "MsgEvId1" And .Cells(j, 4) = "SIG_3" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                    If .Cells(j, 3).Value2 = "MsgEvId1" And .Cells(j, 4) = "SIG_4" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                    If .Cells(j, 3).Value2 = "MsgEvId1" And .Cells(j, 4) = "SIG_5" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                    If .Cells(j, 3).Value2 = "MsgEvId1" And .Cells(j, 4) = "SIG_6" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                    If .Cells(j, 3).Value2 = "MsgEvId1" And .Cells(j, 4) = "SIG_7" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
                    If .Cells(j, 3).Value2 = "MsgEvId1" And .Cells(j, 4) = "SIG_8" Then
                        Sheets("Report").Cells(i, 15).Value2 = .Cells(j, 5).Value2
                    End If
            End If
        Next j
    Next i
End With
Else:
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
End If

Call SBO.SBO(wb, wsh_Path)
Call RDX.RDX(wb, wsh_Path)

'-----------------------------------new code for AI & NOC

    Dim rows_symbol_Report_AI
    rows_symbol_Report_AI = Sheets("Report").UsedRange.Rows.count
    Dim rows_HWConfig_AI
    rows_HWConfig_AI = Sheets("AI").UsedRange.Rows.count
    Dim rows_HWConfig_NOC
    rows_HWConfig_NOC = Sheets("Normal OC").UsedRange.Rows.count
    For q = 2 To rows_symbol_Report_AI Step 1
        Dim symbol_from_report_AI As String
        symbol_from_report_AI = Sheets("Report").Cells(q, 1).Value2
       ' Debug.Print symbol_from_report_AI
            'Add missing (AI) RDX block types
            For j = 2 To rows_HWConfig_AI Step 1
                Dim IOComment As String
                current_IOComment = Sheets("AI").Cells(j, 3)
               ' Debug.Print current_IOComment
                If Trim(symbol_from_report_AI) = Trim(current_IOComment) Then
                Dim checkForEmpty As String
                checkForEmpty = Sheets("Report").Cells(q, 13)
                 If Len(checkForEmpty) < 1 Then
                     Dim current_IOBlock As String
                     current_IOBlock = Sheets("AI").Cells(j, 10)
                    ' Debug.Print current_IOBlock
                     Sheets("AI").Cells(j, 10).Copy Sheets("Report").Cells(q, 13)
                 End If
                End If
            Next
               
           ' add NOCs from new digital lists to the report
           For k = 2 To rows_HWConfig_NOC Step 1
                Dim noc_signal As String
                noc_signal = Sheets("Normal OC").Cells(k, 1)
                'Debug.Print noc_signal
                If Trim(symbol_from_report_AI) = Trim(noc_signal) Then
                      If noc_signal <> "Spare" Then
                         Dim current_NOC As String
                         current_NOC = Sheets("Normal OC").Cells(k, 2)
                         'Debug.Print current_NOC
                         'Debug.Print noc_signal
                         Sheets("Normal OC").Cells(k, 2).Copy Sheets("Report").Cells(q, 14)
                      End If
                End If
        Next

          
        Dim seperateString As String
        seperateString = Sheets("Report").Cells(q, 13)
        Dim range2
        Dim type2
        Dim POS
        POS = InStr(seperateString, """")
        range2 = Mid(seperateString, POS + 1, Len(seperateString))
        POS = InStr(range2, """")
        range2 = Mid(range2, POS + 1, Len(range2))
        'Debug.Print range2
        range2 = Replace(range2, """", "")
        type2 = Mid(seperateString, 1, POS)
        'Debug.Print type2
        type2 = Replace(type2, """", "")
        Dim LArrayRange() As String
        LArrayRange = Split(range2, " ")
        Dim icheckarraysize
                
         ' UBound(LArray, 1) gives the upper limit of the first dimension, which is 5.
        icheckarraysize = UBound(LArrayRange, 1) - LBound(LArrayRange, 1) + 1
        If icheckarraysize > 1 Then
          Dim strNewString
          strNewString = LArrayRange(1)
          strNewString = strNewString & ", " & LArrayRange(0)
          Sheets("Report").Cells(q, 13).Value = Trim(strNewString)
          Sheets("Report").Cells(q, 26).Value = Trim(type2)
        End If
            
    Next
    
'-----------------------------------end code for AI & NOC from the report to the new tab

Call SOE.SOE(wb, wsh_Path)

With Sheets("Report")
    intn_Report = .Cells(Rows.count, 1).End(xlUp).Row
    .Cells(1, 1).EntireColumn.Insert
    .Cells(1, 1).Value2 = "Row #"
    For i = 2 To intn_Report Step 1
        .Cells(i, 1).Value = i - 1
    Next i
End With

'shift tab
    Sheets("Report").Activate
    Sheets("Report").Columns("AA:AA").Select
    Selection.Cut
    Sheets("Report").Columns("O:O").Select
    Selection.Insert Shift:=xlToRight
    
    
'Top align cells
Dim wks As Worksheet
For Each wks In Worksheets
    wks.Cells.VerticalAlignment = xlTop
    wks.Cells.HorizontalAlignment = xlLeft
Next wks

'Add data to template
Set wbTemplate = Workbooks.Open("X:\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TEMPLATE IO List Report For Extraction Tool.xlsx")
wb.Sheets("Report").Range("A2:AA" & intn_Report).Copy
wbTemplate.Sheets("Report").Range("A2").PasteSpecial xlPasteValues


' change the name of the report
Dim strCPUtemplateName As String
strCPUtemplateName = wb.Sheets("CPU").Cells(1, 1).Value2
Dim TestArray() As String
TestArray = Split(strCPUtemplateName, ",")
strCPUtemplateName = TestArray(1)

Dim strTestForSpec As Integer
strTestForSpec = InStr(1, strCPUtemplateName, "/")

If strTestForSpec > 0 Then
   strCPUtemplateName = Replace(strCPUtemplateName, "/", "-")
End If

wbTemplate.Sheets("Report").Name = Replace(strCPUtemplateName, """", "")
wb.Sheets("Report").Range("A2:AA" & intn_Report).Copy
wbTemplate.Sheets("SOE").Range("A2").PasteSpecial xlPasteValues
wb.Sheets("Report").Range("A2:AA" & intn_Report).Copy
wbTemplate.Sheets("SBO").Range("A2").PasteSpecial xlPasteValues
wb.Sheets("Report").Range("A2:AA" & intn_Report).Copy
wbTemplate.Sheets("RDX").Range("A2").PasteSpecial xlPasteValues



' Delete everything but DI/DO and AI8x12Bit from template tab report
Dim iReportCount As Integer: iReportCount = wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).UsedRange.Rows.count
     Do While iReportCount > 1
         Dim strCurrReportType
         strCurrReportType = wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Cells(iReportCount, 14).Value
         If strCurrReportType = "RD_X_SOE" Or strCurrReportType = "WR_X_SBO" Or Trim(strCurrReportType) = "RD_X_AI1" Or Trim(strCurrReportType) = "RD_X_AI16" Then
           wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Rows(iReportCount).EntireRow.Delete
         End If

      iReportCount = iReportCount - 1
     Loop
        
' Delete everything but SOE from template tab report
Dim iSOECount As Integer: iSOECount = wbTemplate.Sheets("SOE").UsedRange.Rows.count
Dim iCurrCountRows As Integer
iCurrCountRows = 2
    Do While iSOECount > 1
        Dim strCurrSOEType
        strCurrSOEType = wbTemplate.Sheets("SOE").Cells(iSOECount, 14).Value
        If strCurrSOEType <> "RD_X_SOE" Then
          wbTemplate.Sheets("SOE").Rows(iSOECount).EntireRow.Delete
        Else
          wbTemplate.Sheets("SOE").Cells(iCurrCountRows, 1).Value = iCurrCountRows
          iCurrCountRows = iCurrCountRows + 1
          Dim tempStr As String: tempStr = wbTemplate.Sheets("SOE").Cells(iSOECount, 17).Value
          Dim tempStrLen As Integer: tempStrLen = Len(wbTemplate.Sheets("SOE").Cells(iSOECount, 17).Value)
            If Len(wbTemplate.Sheets("SOE").Cells(iSOECount, 17).Value) = 0 Then
                If Trim(wbTemplate.Sheets("SOE").Cells(iSOECount, 2).Value) <> "SPARE" Then
                    Range("Q" & iSOECount).Select
                    ActiveCell.FormulaR1C1 = wbTemplate.Sheets("SOE").Cells(iSOECount, 2).Value
                End If
            End If
        End If
    iSOECount = iSOECount - 1
    Loop
    iCurrCountRows = 1
Dim iSOECount2 As Integer: iSOECount2 = wbTemplate.Sheets("SOE").UsedRange.Rows.count
Do While iSOECount2 > 1
    If Len(wbTemplate.Sheets("SOE").Cells(iSOECount2, 17).Value) = 0 Then
        If Trim(wbTemplate.Sheets("SOE").Cells(iSOECount2, 2).Value) = "SPARE" Then
            Range("Q" & iSOECount2).Select
            ActiveCell.FormulaR1C1 = wbTemplate.Sheets("SOE").Cells(iSOECount2, 2).Value & " " & (iSOECount2 - 1)
        End If
    End If
iSOECount2 = iSOECount2 - 1
Loop
                    
 ' Delete everything but SBO from template tab report
Dim iSBOCount As Integer: iSBOCount = wbTemplate.Sheets("SBO").UsedRange.Rows.count
     Do While iSBOCount > 1
         Dim strCurrSBOType
         strCurrSBOType = wbTemplate.Sheets("SBO").Cells(iSBOCount, 14).Value
         If strCurrSBOType <> "WR_X_SBO" Then
           wbTemplate.Sheets("SBO").Rows(iSBOCount).EntireRow.Delete
         End If
      iSBOCount = iSBOCount - 1
     Loop
     
' Delete everything but RDX from template tab report
Dim irdxCOUNT As Integer: irdxCOUNT = wbTemplate.Sheets("RDX").UsedRange.Rows.count

     Do While irdxCOUNT > 1
         Dim strCurrRDXType
         strCurrRDXType = wbTemplate.Sheets("RDX").Cells(irdxCOUNT, 14).Value
         If Trim(strCurrRDXType) <> "RD_X_AI1" And Trim(strCurrRDXType) <> "RD_X_AI16" Then
           wbTemplate.Sheets("RDX").Rows(irdxCOUNT).EntireRow.Delete
         End If

      irdxCOUNT = irdxCOUNT - 1
     Loop
   

wb.Sheets("File Paths").Range("A2:AA" & intn_Report).Copy
wbTemplate.Sheets("File Paths").Range("A2").PasteSpecial xlPasteValues

'Sort by Rack then Slot then Channel #

' sort tabs ASC rack/slot/channel
ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort. _
    SortFields.Clear
ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort. _
    SortFields.Add2 Key:=Range("Table1[Rack]"), SortOn:=xlSortOnValues, Order _
    :=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort. _
    SortFields.Add2 Key:=Range("Table1[Slot]"), SortOn:=xlSortOnValues, Order _
    :=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort. _
    SortFields.Add2 Key:=Range("Table1[Chnl]"), SortOn:=xlSortOnValues, Order _
    :=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort. _
    SortFields.Add2 Key:=Range("Table1[Alarm-AH]"), SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort. _
    SortFields.Add2 Key:=Range("Table1[Alarm-WH]"), SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort. _
    SortFields.Add2 Key:=Range("Table1[Alarm-WL]"), SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort. _
    SortFields.Add2 Key:=Range("Table1[Alarm-AL]"), SortOn:=xlSortOnValues, _
    Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1"). _
    Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

ActiveWorkbook.Worksheets("SOE").ListObjects("Table15").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("SOE").ListObjects("Table15").Sort.SortFields.Add2 _
    Key:=Range("Table15[Rack]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("SOE").ListObjects("Table15").Sort.SortFields.Add2 _
    Key:=Range("Table15[Slot]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("SOE").ListObjects("Table15").Sort.SortFields.Add2 _
    Key:=Range("Table15[Chnl]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("SOE").ListObjects("Table15").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


ActiveWorkbook.Worksheets("RDX").ListObjects("Table134").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("RDX").ListObjects("Table134").Sort.SortFields.Add2 _
    Key:=Range("Table134[Rack]"), SortOn:=xlSortOnValues, Order:=xlAscending _
    , DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("RDX").ListObjects("Table134").Sort.SortFields.Add2 _
    Key:=Range("Table134[Slot]"), SortOn:=xlSortOnValues, Order:=xlAscending _
    , DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("RDX").ListObjects("Table134").Sort.SortFields.Add2 _
    Key:=Range("Table134[Chnl]"), SortOn:=xlSortOnValues, Order:=xlAscending _
    , DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("RDX").ListObjects("Table134").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

    
ActiveWorkbook.Worksheets("SBO").ListObjects("Table13").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("SBO").ListObjects("Table13").Sort.SortFields.Add2 _
    Key:=Range("Table13[Rack]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("SBO").ListObjects("Table13").Sort.SortFields.Add2 _
    Key:=Range("Table13[Slot]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("SBO").ListObjects("Table13").Sort.SortFields.Add2 _
    Key:=Range("Table13[Chnl]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
    DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("SBO").ListObjects("Table13").Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
    
    
    
wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Activate
Range("Table1[Ditial?]").Select
Selection.ListObject.ListColumns(26).Delete

wbTemplate.Sheets("SOE").Activate
Range("Table15[Ditial?]").Select
Selection.ListObject.ListColumns(26).Delete

wbTemplate.Sheets("SBO").Activate
Range("Table13[Ditial?]").Select
Selection.ListObject.ListColumns(26).Delete

wbTemplate.Sheets("RDX").Activate
Range("Table134[Ditial?]").Select
Selection.ListObject.ListColumns(26).Delete
        
' RDX range swap and type inconsistency fix
Dim iRdxSwap As Integer: iRdxSwap = wbTemplate.Sheets("RDX").UsedRange.Rows.count
For i = 2 To iRdxSwap Step 1
' fix type inconsistency
Range("N" & i).Select
Dim strCurrType As String
strCurrType = ActiveCell.FormulaR1C1
If strCurrType <> "RD_X_AI16" Then
    ActiveCell.FormulaR1C1 = "RD_X_AI16"
End If
    
' swap ranges
wbTemplate.Sheets("RDX").Activate
Range("H" & i).Select
Dim tempLH As String
tempLH = ActiveCell.FormulaR1C1

Range("I" & i).Select
Dim tempHI As String
tempHI = ActiveCell.FormulaR1C1

Range("H" & i).Select
ActiveCell.FormulaR1C1 = tempHI

Range("I" & i).Select
ActiveCell.FormulaR1C1 = tempLH
Next
   
   
Dim iReportReCount As Integer: iReportReCount = wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).UsedRange.Rows.count
Dim iGeRcount As Integer
iGeRcount = 1
For i = 2 To iReportReCount Step 1
    wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Activate
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = iGeRcount
    iGeRcount = iGeRcount + 1
Next
iGeRcount = 1
    
' insert unit col
Columns("AC:AC").Select
Selection.Cut
Columns("J:J").Select
Selection.Insert Shift:=xlToRight

' move chart column inbetween alarm text and block
Columns("U:U").Select
Application.CutCopyMode = False
Selection.Cut
Columns("S:S").Select
Selection.Insert Shift:=xlToRight


' hide everyhing after the block column
Columns("U:AA").Select
Selection.EntireColumn.Hidden = True


iReportReCount = wbTemplate.Sheets("SOE").UsedRange.Rows.count
For i = 2 To iReportReCount Step 1
    wbTemplate.Sheets("SOE").Activate
    Range("A" & i).Select
    ActiveCell.FormulaR1C1 = iGeRcount
    iGeRcount = iGeRcount + 1
Next
iGeRcount = 1

' insert unit col
Columns("AC:AC").Select
Selection.Cut
Columns("J:J").Select
Selection.Insert Shift:=xlToRight

' move chart column inbetween alarm text and block
Columns("U:U").Select
Application.CutCopyMode = False
Selection.Cut
Columns("S:S").Select
Selection.Insert Shift:=xlToRight

' hide SOE alarms and ranges
Range("H:I,K:N").Select
Range("Table15[[#Headers],[Alarm-WL]]").Activate
Selection.EntireColumn.Hidden = True

' hide everyhing after the block column
Columns("U:AA").Select
Selection.EntireColumn.Hidden = True

'increaes column size for the alarm text
Columns("R:R").Select
Selection.ColumnWidth = 60
    
'change size of the font for alarm text
Columns("R:R").Select
With Selection.Font
    .Name = "Calibri"
    .Size = 9
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontMinor
End With
    
' negate change of size to header
Range("Table15[[#Headers],[Alarm Text]]").Select
With Selection.Font
    .Name = "Calibri"
    .Size = 7
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ThemeColor = xlThemeColorLight1
    .TintAndShade = 0
    .ThemeFont = xlThemeFontMinor
End With
    
iReportReCount = wbTemplate.Sheets("SBO").UsedRange.Rows.count
For i = 2 To iReportReCount Step 1
  wbTemplate.Sheets("SBO").Activate
  Range("A" & i).Select
  ActiveCell.FormulaR1C1 = iGeRcount
  iGeRcount = iGeRcount + 1
Next
iGeRcount = 1

' insert unit col
  Columns("AC:AC").Select
  Selection.Cut
  Columns("J:J").Select
  Selection.Insert Shift:=xlToRight
  
  ' move chart column inbetween alarm text and block
  Columns("U:U").Select
  Application.CutCopyMode = False
  Selection.Cut
  Columns("S:S").Select
  Selection.Insert Shift:=xlToRight
    
' hide SBO alarms and ranges
  Range("H:I,K:K,L:L,M:M,N:N").Select
  Range("Table13[[#Headers],[Alarm-AH]]").Activate
  Selection.EntireColumn.Hidden = True
  
  ' hide everyhing after the block column
  Columns("U:AA").Select
  Selection.EntireColumn.Hidden = True
    
    
iReportReCount = wbTemplate.Sheets("RDX").UsedRange.Rows.count
For i = 2 To iReportReCount Step 1
  wbTemplate.Sheets("RDX").Activate
  Range("A" & i).Select
  ActiveCell.FormulaR1C1 = iGeRcount
  iGeRcount = iGeRcount + 1
Next
      
'insert unit col
Columns("AC:AC").Select
Selection.Cut
Columns("J:J").Select
Selection.Insert Shift:=xlToRight
    
' move chart column inbetween alarm text and block
Columns("U:U").Select
Application.CutCopyMode = False
Selection.Cut
Columns("S:S").Select
Selection.Insert Shift:=xlToRight
    
' hide everyhing after the block column
Columns("U:AA").Select
Selection.EntireColumn.Hidden = True
    
If AVP = True Then

     ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[Rack]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[Slot]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[Chnl]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(Replace(strCPUtemplateName, """", "")).ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[Address]"), SortOn:=xlSortOnValues, Order:=xlAscending _
        , DataOption:=xlSortNormal
        
            
    Dim checkForApacsLateral() As String
    ' Delete APACS alternative channel
    Dim iReportCountAPACS As Integer: iReportCountAPACS = wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).UsedRange.Rows.count
    iReportCountAPACS = iReportCountAPACS - 1
    wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Activate
        Do While iReportCountAPACS > 1
            Range("O" & iReportCountAPACS - 1).Select
            checkForApacsLateral = Split(ActiveCell.FormulaR1C1, "_")
            If checkForApacsLateral(0) = "DI 24V ""SDM" Or checkForApacsLateral(0) = "DI 24V ""SAM" Or checkForApacsLateral(0) = "DI 24V ""VIM" Or checkForApacsLateral(0) = "DI 24V ""EAM" Then
                Range("O" & iReportCountAPACS).Select
                checkForApacsLateral = Split(ActiveCell.FormulaR1C1, "_")
                If checkForApacsLateral(0) = "DO 24V ""SDM" Or checkForApacsLateral(0) = "DO 24V ""SAM" Or checkForApacsLateral(0) = "DO 24V ""VIM" Or checkForApacsLateral(0) = "DO 24V ""EAM" Then
                     Range("T" & iReportCountAPACS - 1).Select
                     If ActiveCell.FormulaR1C1 <> "" Then
                        wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Rows(iReportCountAPACS).EntireRow.Delete
                     End If
                End If
            End If
            iReportCountAPACS = iReportCountAPACS - 1
        Loop

    
    iReportCountAPACS = wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).UsedRange.Rows.count
    iReportCountAPACS = iReportCountAPACS - 1
    wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Activate
    Do While iReportCountAPACS > 1
        Range("O" & iReportCountAPACS).Select
        checkForApacsLateral = Split(ActiveCell.FormulaR1C1, "_")
        If checkForApacsLateral(0) = "DO 24V ""SDM" Or checkForApacsLateral(0) = "DO 24V ""SAM" Or checkForApacsLateral(0) = "DO 24V ""VIM" Or checkForApacsLateral(0) = "DO 24V ""EAM" Then
            Range("O" & iReportCountAPACS - 1).Select
            checkForApacsLateral = Split(ActiveCell.FormulaR1C1, "_")
                If checkForApacsLateral(0) = "DI 24V ""SDM" Or checkForApacsLateral(0) = "DI 24V ""SAM" Or checkForApacsLateral(0) = "DI 24V ""VIM" Or checkForApacsLateral(0) = "DI 24V ""EAM" Then
                     Range("T" & iReportCountAPACS).Select
                     If ActiveCell.FormulaR1C1 <> "" Then
                        Range("T" & iReportCountAPACS - 1).Select
                         If ActiveCell.FormulaR1C1 = "" Then
                            wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Rows(iReportCountAPACS - 1).EntireRow.Delete
                         End If
                     End If
                End If
        End If
        iReportCountAPACS = iReportCountAPACS - 1
    Loop
        
       
    iReportCountAPACS = wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).UsedRange.Rows.count
    iReportCountAPACS = iReportCountAPACS - 1
    wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Activate
    Do While iReportCountAPACS > 1
        Range("O" & iReportCountAPACS).Select
        checkForApacsLateral = Split(ActiveCell.FormulaR1C1, "_")
        If checkForApacsLateral(0) = "DI 24V ""SDM" Or checkForApacsLateral(0) = "DI 24V ""SAM" Or checkForApacsLateral(0) = "DI 24V ""VIM" Or checkForApacsLateral(0) = "DI 24V ""EAM" Then
             Range("T" & iReportCountAPACS).Select
             If ActiveCell.FormulaR1C1 = "" Then
                Range("D" & iReportCountAPACS).Select
                   ActiveCell.FormulaR1C1 = ""
                   Range("C" & iReportCountAPACS).Select
                   Dim APACSADD As String
                   APACSADD = ActiveCell.FormulaR1C1
                   Range("B" & iReportCountAPACS).Select
                   ActiveCell.FormulaR1C1 = APACSADD
             End If
        End If
        If checkForApacsLateral(0) = "DO 24V ""SDM" Or checkForApacsLateral(0) = "DO 24V ""SAM" Or checkForApacsLateral(0) = "DO 24V ""VIM" Or checkForApacsLateral(0) = "DO 24V ""EAM" Then
           Range("T" & iReportCountAPACS).Select
             If ActiveCell.FormulaR1C1 = "" Then
                Range("D" & iReportCountAPACS).Select
                    ActiveCell.FormulaR1C1 = ""
                   Range("C" & iReportCountAPACS).Select
                   Dim APACSADD2 As String
                   APACSADD2 = ActiveCell.FormulaR1C1
                   Range("B" & iReportCountAPACS).Select
                   ActiveCell.FormulaR1C1 = APACSADD2
             End If
        End If
        iReportCountAPACS = iReportCountAPACS - 1
    Loop
                            
Application.DisplayAlerts = False
Worksheets("SOE").Delete
Worksheets("SBO").Delete
Worksheets("RDX").Delete
Application.DisplayAlerts = True
    
' end AVP
End If


Application.CutCopyMode = False
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







