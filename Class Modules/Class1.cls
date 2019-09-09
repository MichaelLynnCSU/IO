VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' CLASS MODULE CODE
' Class name: Class1



Private Function Cols(SheetName As String) As Integer
    Dim cols_ As Integer: cols_ = Sheets(SheetName).UsedRange.Columns.count
    For k = 2 To cols_ Step 1
    If Sheets(SheetName).Cells(1, k).Value2 = "Block type" Then
    'function return in the same name as the function
        Cols = k
    End If
    Next

End Function

Public Function mySaveAsCSV()
   
Application.DisplayAlerts = False
For Each mySheet In ActiveWorkbook.Worksheets

    myPath = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\WIP Michael L\IO_List_WIP\WIP\sheets\"
    If mySheet.Name = "Sheet1" Then
    mySheet.Delete
    Else
    ActiveWorkbook.Sheets(mySheet.index).Copy
    ActiveWorkbook.saveAs Filename:=myPath & mySheet.Name, FileFormat:=xlCSV, CreateBackup:=True
    ActiveWorkbook.Close
    End If
Next mySheet
Application.DisplayAlerts = True

End Function

Public Function DiOu(SheetName As String, e As BlockTypeStruct)
    ' new APL CH_AI -> PCS7AnIn
    Dim targetCOL As Integer
    targetCOL = Cols(SheetName)
    Dim intRow As Integer
    intRow = Sheets(SheetName).UsedRange.Rows.count
    For k = 2 To intRow Step 1
        If Sheets(SheetName).Cells(k, targetCOL).Value2 = "CH_DO" Then
            Sheets(SheetName).Cells(k, targetCOL).Value2 = e.CH_DO
        End If
        If Sheets(SheetName).Cells(k, 6).Value2 = "I" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_In"
        ElseIf LCase(Sheets(SheetName).Cells(k, 6).Value2) = "value" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_Out"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "QBAD" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "Bad"
        End If
    Next
    

End Function

Public Function DI(SheetName As String, e As BlockTypeStruct)
    Dim targetCOL As Integer
    targetCOL = Cols(SheetName)
    Dim intRow As Integer
    intRow = Sheets(SheetName).UsedRange.Rows.count
    For k = 2 To intRow Step 1
        If Sheets(SheetName).Cells(k, targetCOL).Value2 = "CH_DI" Then
            Sheets(SheetName).Cells(k, targetCOL).Value2 = e.CH_DI
        End If
        If LCase(Sheets(SheetName).Cells(k, 6).Value2) = "value" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_In"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "Q" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_Out"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "QBAD" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "Bad"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "LAST_ON" And Sheets(SheetName).Cells(k, 10).Value2 = "0" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "Feature.Bit29=0"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "LAST_ON" And Sheets(SheetName).Cells(k, 10).Value2 = "1" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "Feature.Bit29=1"
        End If
    Next


End Function

Public Function AO(SheetName As String, e As BlockTypeStruct)
    Dim targetCOL As Integer
    targetCOL = Cols(SheetName)
    
    Dim intRow As Integer
    intRow = Sheets(SheetName).UsedRange.Rows.count
    For k = 2 To intRow Step 1
        If Sheets(SheetName).Cells(k, targetCOL).Value2 = "CH_AO" Then
            Sheets(SheetName).Cells(k, targetCOL).Value2 = e.CH_AO
        End If
        
        If Sheets(SheetName).Cells(k, 6).Value2 = "U" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_In"
        ElseIf LCase(Sheets(SheetName).Cells(k, 6).Value2) = "value" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_Out"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "UHRANGE" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "Scale.High"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "ULRANGE" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "Scale.Low"
        End If
    Next k
   
    Dim datablockf1, datablockf2 As Boolean
    Dim trigger As String

    If SheetName = "CH_AI_Ranges" Then
           'interconnection algorithm start
        intRow = Sheets(SheetName).UsedRange.Rows.count
        Dim strDug As Integer
        For k = 2 To intRow Step 1
        trigger = ""
        datablockf1 = False
        datablockf2 = False
            If Sheets(SheetName).Cells(k, 6).Value2 = "PV_Out" Then
                Dim Result() As String
                Dim TextStrng As String
                TextStrng = Sheets(SheetName).Cells(k, 12).Value2
                If TextStrng <> "" Then
                    Result() = Split(TextStrng, """")
                    Dim length As Integer
                    length = UBound(Result, 1) - LBound(Result, 1) + 1
                    For j = 0 To (length - 1) Step 1
                    strDig = InStr(1, Result(j), ".U")
                        If strDig > 0 Then
                        Dim test As String
                            test = Mid(Result(j), Len(Result(j)) - 1)
                            If Mid(Result(j), Len(Result(j)) - 1) = ".U" Then
                                datablockf1 = True
                                Debug.Print Result(j)
                            End If
                        End If
                    Next
                    For i = 0 To (length - 1) Step 1
                    strDig = InStr(1, Result(i), "\")
                        If strDig < 1 Then
                            If Result(i) <> "" Then
                                datablockf2 = True
                                trigger = Result(i)
                                Debug.Print Result(i)
                            End If
                        End If
                    Next
                End If
            End If
            If datablockf1 = True And datablockf2 = True Then
                Debug.Print trigger
                Sheets(SheetName).Cells(k, 11).Value2 = trigger
                Sheets(SheetName).Cells(k, 12).Value2 = Replace(Sheets(SheetName).Cells(k, 12).Value2, trigger, "")
            End If
        Next k
    End If

 
End Function

Public Function AI(SheetName As String, e As BlockTypeStruct)
    ' new APL CH_AI -> PCS7AnIn
    
    Dim targetCOL As Integer
    targetCOL = Cols(SheetName)
    Dim intRow As Integer
    intRow = Sheets(SheetName).UsedRange.Rows.count
    For k = 2 To intRow Step 1
            If Sheets(SheetName).Cells(k, targetCOL).Value2 = "CH_AI" And SheetName = "CH_AI_Signals" Then
                Sheets(SheetName).Cells(k, targetCOL).Value2 = e.CH_AI
            End If
            If Sheets(SheetName).Cells(k, targetCOL).Value2 = "CH_AI" And SheetName = "CH_AI_Ranges" Then
                Sheets(SheetName).Cells(k, targetCOL).Value2 = e.CH_AI
                
                    If Sheets(SheetName).Cells(k, 6).Value2 = "V" Then
                        Dim Result2() As String
                        Dim TextStrng2 As String
                        Dim Result3() As String
                        TextStrng2 = Sheets(SheetName).Cells(k, 12).Value2
                        If TextStrng2 <> "" Then
                            Dim intEndPos2 As Integer
                            intEndPos2 = InStrRev(TextStrng2, ".U")
                            If intEndPos2 > 0 Then
                                Result2() = Split(TextStrng2, ".U""")
                                Result3() = Split(Result2(0), "\")
                                Dim len1 As Integer
                                len1 = UBound(Result3, 1) - LBound(Result3, 1) + 1
                                If len1 > 2 Then
                                    Sheets(SheetName).Cells(k, 4).Value2 = Result3(2)
                                    Sheets(SheetName).Cells(k, targetCOL).Value2 = "MonAnL"
                                Else
                                    Sheets(SheetName).Cells(k, 4).Value2 = Result3(1)
                                    Sheets(SheetName).Cells(k, targetCOL).Value2 = "MonAnL"
                                End If
                            End If
                        End If
                    End If
            End If

            If Sheets(SheetName).Cells(k, 6).Value2 = "V" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "PV_Out"
            ElseIf LCase(Sheets(SheetName).Cells(k, 6).Value2) = "value" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "PV_In"
            ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "VHRANGE" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "Scale.High"
            ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "VLRANGE" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "Scale.Low"
            ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "Last_ON" And Sheets(SheetName).Cells(k, 10).Value2 = "0" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "Feature.Bit29=0"
            ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "Last_ON" And Sheets(SheetName).Cells(k, 10).Value2 = "1" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "Feature.Bit29=1"
            ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "SUBS_V" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "SubsPV_In"
            ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "LAST_ON" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "Feature.Bit30"
            ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "CH_F_HL" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "HighLimit"
            ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "CH_F_LL" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "LowLimit"
            ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "QBAD" Then
                Sheets(SheetName).Cells(k, 6).Value2 = "Bad"
            End If
    Next k
    
    Dim datablockf1, datablockf2 As Boolean
    Dim trigger As String

    If SheetName = "CH_AI_Ranges" Then
           'interconnection algorithm start
        intRow = Sheets(SheetName).UsedRange.Rows.count
        Dim strDug As Integer
        For k = 2 To intRow Step 1
        trigger = ""
        datablockf1 = False
        datablockf2 = False
            If Sheets(SheetName).Cells(k, 6).Value2 = "PV_Out" Then
                Dim Result() As String
                Dim TextStrng As String
                TextStrng = Sheets(SheetName).Cells(k, 12).Value2
                If TextStrng <> "" Then
                    Result() = Split(TextStrng, """")
                    Dim length As Integer
                    length = UBound(Result, 1) - LBound(Result, 1) + 1
                    For j = 0 To (length - 1) Step 1
                    strDig = InStr(1, Result(j), ".U")
                        If strDig > 0 Then
                        Dim test As String
                            test = Mid(Result(j), Len(Result(j)) - 1)
                            If Mid(Result(j), Len(Result(j)) - 1) = ".U" Then
                                datablockf1 = True
                                Debug.Print Result(j)
                            End If
                        End If
                    Next
                    For i = 0 To (length - 1) Step 1
                    strDig = InStr(1, Result(i), "\")
                        If strDig < 1 Then
                            If Result(i) <> "" Then
                                datablockf2 = True
                                trigger = Result(i)
                                Debug.Print Result(i)
                            End If
                        End If
                    Next
                End If
            End If
            If datablockf1 = True And datablockf2 = True Then
                Debug.Print trigger
                Sheets(SheetName).Cells(k, 11).Value2 = trigger
                Sheets(SheetName).Cells(k, 12).Value2 = Replace(Sheets(SheetName).Cells(k, 12).Value2, trigger, "")
            End If
        Next k
    End If



  
End Function

Public Function messages(SheetName As String, e As BlockTypeStruct)
    Dim targetCOL As Integer
    targetCOL = Cols(SheetName)
    Dim intRow As Integer
    intRow = Sheets(SheetName).UsedRange.Rows.count
    Dim myNewstring As String
    For k = 2 To intRow Step 1
        If Sheets(SheetName).Cells(k, targetCOL).Value2 = "MEAS_MON" Then
            Sheets(SheetName).Cells(k, targetCOL).Value2 = e.MEAS_MON
        End If
        
        myNewstring = ""
        'part A “Longest Common Subsequence” interative
        Dim Result() As String
        Dim TextStrng As String
        Dim TextStrng2 As String
        TextStrng = Sheets(SheetName).Cells(k, 14).Value2
        TextStrng2 = Sheets(SheetName).Cells(k + 1, 14).Value2
        If TextStrng <> "" And TextStrng2 <> "" Then
        Result() = Split(TextStrng, " ")
            If Result(0) <> "$$BlockComment$$" And Sheets(SheetName).Cells(k, 5).Value2 = "Meas.value monitoring block" And Sheets(SheetName).Cells(k, 8).Value2 = "SIG_1" Then
                For counter = 1 To Len(TextStrng) Step 1
                    If Mid(TextStrng, counter, 1) = Mid(TextStrng2, counter, 1) Then
                        If Len(TextStrng2) > counter Then
                            myNewstring = myNewstring + Mid(TextStrng, counter, 1)
                        End If
                    End If
                Next counter
                Sheets(SheetName).Cells(k, 5).Value2 = myNewstring
                Sheets(SheetName).Cells(k + 1, 5).Value2 = myNewstring
                Sheets(SheetName).Cells(k + 2, 5).Value2 = myNewstring
            End If
        End If
    Next k

  
End Function

Public Function MeasMon(SheetName As String, e As BlockTypeStruct)
    Dim targetCOL As Integer
    targetCOL = Cols(SheetName)
    Dim intRow As Integer
    intRow = Sheets(SheetName).UsedRange.Rows.count
    For k = 2 To intRow Step 1
        If Sheets(SheetName).Cells(k, targetCOL).Value2 = "MEAS_MON" Then
            Sheets(SheetName).Cells(k, targetCOL).Value2 = e.MEAS_MON
        End If
        If Sheets(SheetName).Cells(k, 6).Value2 = "U_AH" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_AH_Lim"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "U_WH" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_WH_Lim"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "U_WL" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_WL_Lim"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "U_AL" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_AL_Lim"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "M_SUP_AH" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_AH_En"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "M_SUP_WH" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_WH_En"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "M_SUP_WL" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_WL_En"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "M_SUP_AL" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_AL_En"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "HYS" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_Hyst"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "HYS" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_Hyst"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "QH_ALM" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_AH_AC"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "QL_ALM" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_AL_AC"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "QH_WRN" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_WH_WRN"
        ElseIf Sheets(SheetName).Cells(k, 6).Value2 = "QL_WRN" Then
            Sheets(SheetName).Cells(k, 6).Value2 = "PV_WL_WRN"
        End If
        
        
        ' supression key
        If Sheets(SheetName).Cells(k, 10).Value2 = "Suppress HH=Yes" Or Sheets(SheetName).Cells(k, 10).Value2 = "Suppress LL=Yes" Or Sheets(SheetName).Cells(k, 10).Value2 = "Suppress H=Yes" Or Sheets(SheetName).Cells(k, 10).Value2 = "Suppress L=Yes" Then
            Sheets(SheetName).Cells(k, 10).Value2 = "1"
        ElseIf Sheets(SheetName).Cells(k, 10).Value2 = "Suppress HH=No" Or Sheets(SheetName).Cells(k, 10).Value2 = "Suppress LL=No" Or Sheets(SheetName).Cells(k, 10).Value2 = "Suppress H=No" Or Sheets(SheetName).Cells(k, 10).Value2 = "Suppress L=No" Then
            Sheets(SheetName).Cells(k, 10).Value2 = "0"
        End If
    Next k

'    Dim C As Integer
'    C = Sheets(SheetName).Cells.SpecialCells(xlLastCell).Column
'    Do Until C = 0
'    If WorksheetFunction.CountA(Columns(C)) = 0 Then
'        Columns(C).Delete
'    End If
'    C = C - 1
'    Loop
'
'    Dim lastrow As Integer
'    Dim lastCol As Integer
'    lastrow = Sheets(SheetName).Cells(1, 1).End(xlDown).Row
'    lastCol = Sheets(SheetName).Range("a1").End(xlToRight).Column
'    'wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Range("A1", wbTemplate.Sheets(Replace(strCPUtemplateName, """", "")).Cells(Lastrow, lastCol)).Select
'    For i = 1 To lastrow Step 1
'        For j = 1 To lastCol Step 1
'            If Not IsNumeric(Sheets(SheetName).Cells(i, j).Value) Then
'                Sheets(SheetName).Cells(i, j) = """" + Sheets(SheetName).Cells(i, j).Value + """"
'            ElseIf IsEmpty(Sheets(SheetName).Cells(i, j).Value) Then
'                Sheets(SheetName).Cells(i, j) = """"""
'            End If
'        Next j
'    Next i
End Function
