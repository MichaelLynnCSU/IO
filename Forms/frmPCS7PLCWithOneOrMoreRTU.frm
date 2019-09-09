VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPCS7PLCWithOneOrMoreRTU 
   Caption         =   "File Input"
   ClientHeight    =   14505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12225
   OleObjectBlob   =   "frmPCS7PLCWithOneOrMoreRTU.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPCS7PLCWithOneOrMoreRTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCH_AI_Ranges_Click()
  Dim strRanges As Variant
  strRanges = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_AI_Ranges File To Be Opened")
  TextBox11.Value = strRanges
End Sub

Private Sub btnCH_AI_Signal_Click()
  Dim strCH_AI_Singals As Variant
  strCH_AI_Singals = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_AI_Signals File To Be Opened")
  TextBox3.Value = strCH_AI_Singals
End Sub

Private Sub btnCH_DI_Click()
  Dim strDI As Variant
  strDI = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_DI File To Be Opened")
  TextBox6.Value = strDI
End Sub

Private Sub btnCH_DI_Signal_Click()
  Dim strCH_DI_Singals As Variant
  strCH_DI_Singals = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_DI_Signals File To Be Opened")
  TextBox5.Value = strCH_DI_Singals
End Sub

Private Sub btnCH_DI_Signal_NO_NC_Click()
  Dim strDigit As Variant
  strDigit = Application.GetOpenFilename(FileFilter:="Text Files(*.csv),*.csv", Title:="Select New Digit File To Be Opened")
  TextBox9.Value = strDigit
End Sub

Private Sub btnExit_Click()
  Unload frmPCS7PLCWithOneOrMoreRTU
End Sub

Private Sub btnHW_Config_Click()
  Dim strHWConfig As Variant
  strHWConfig = Application.GetOpenFilename(FileFilter:="Text Files(*.cfg),*.cfg", Title:="Select HW Config File To Be Opened")
  TextBox8.Value = strHWConfig
End Sub

Private Sub btnMeas_Mon_Alarming_Click()
  Dim strAlarm As Variant
  strAlarm = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select Meas_Mon_Alarming File To Be Opened")
  TextBox2.Value = strAlarm
End Sub

Private Sub btnMessage_Block_Click()
  Dim strDIAlarm As Variant
  strDIAlarm = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select Message_Block File To Be Opened")
  TextBox7.Value = strDIAlarm
End Sub

Private Sub btnRD_X_AI1_Click()
  Dim strAI As Variant
  strAI = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select RD_X_AI1 File To Be Opened")
  TextBox1.Value = strAI
End Sub

Private Sub btnRD_X_SOE_Click()
  Dim strSOE As Variant
  strSOE = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select RD_X_SOE File To Be Opened")
  TextBox12.Value = strSOE
End Sub

Private Sub btnRD_X_SOE_Message_Click()
  Dim strSOE_Message As Variant
  strSOE_Message = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select RD_X_SOE_Message File To Be Opened")
  TextBox13.Value = strSOE_Message
End Sub

Private Sub btnRun_Click()

    If Me.Controls("CheckBox1").Value = False Then
      blnPlaceHolder = False
    Else
        blnPlaceHolder = True
    End If
    
    ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "HW Config File"
    If Len(TextBox8.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = TextBox8.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\test_hwconfig.cfg"
    End If
    
    ThisWorkbook.Sheets("File Paths").Cells(3, 1).Value2 = "CH_AI_Singals"
    If Len(TextBox3.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = TextBox3.Value
    Else
       ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\test_analog_DI\CHH_Ch_AI_Signals.csv"
    End If
    
    ThisWorkbook.Sheets("File Paths").Cells(4, 1).Value2 = "CH_AI_Ranges"
    If Len(TextBox11.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = TextBox11.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\test_analog_DI\CHH_Ch_AI_Ranges.csv"
    End If
    
    ThisWorkbook.Sheets("File Paths").Cells(5, 1).Value2 = "Meas_Mon_Alarming"
    If Len(TextBox2.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = TextBox2.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\test_analog_DI\CHH_MEAS_MON_ALARMING.csv"
    End If
    
    ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "Symbol Table File"
    If Len(TextBox14.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = TextBox14.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\test_SYMBOLTABLE.asc"
    End If
    
    ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
    If Len(TextBox10.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = TextBox10.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\rtu_test\test__csv.csv"
    End If
  
    ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
    If Len(TextBox1.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = TextBox1.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\rtu_test\test__csv.csv"
    End If
  
    ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
    If Len(TextBox12.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = TextBox12.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\rtu_test\test__csv.csv"
    End If
    
    ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
    If Len(TextBox13.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = TextBox13.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\rtu_test\test__csv.csv"
    End If
   
    ThisWorkbook.Sheets("File Paths").Cells(11, 1).Value2 = "CH_DI_Singals"
    If Len(TextBox9.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = TextBox9.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\test_analog_DI\CHH_CH_DI_Signals.csv"
    End If
        
    
    ThisWorkbook.Sheets("File Paths").Cells(12, 1).Value2 = "CH_DI"
    If Len(TextBox5.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = TextBox5.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\test_analog_DI\CHH_CH_DI.csv"
    End If
   
  
    ThisWorkbook.Sheets("File Paths").Cells(13, 1).Value2 = "Message_Block"
    If Len(TextBox7.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = TextBox7.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\test_analog_DI\CHH_MESSAGE_BLOCK.csv"
    End If
   
    ThisWorkbook.Sheets("File Paths").Cells(14, 1).Value2 = "CH_DI_Signals_NO-NC mod"
    If Len(TextBox9.Value) > 0 Then
        ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = TextBox9.Value
    Else
        ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\testfiles\test_analog_DI\CHH_CH_DI_Signals_NO-NC mod.csv"
    End If
      

      
   Unload Me


End Sub

Private Sub btnSymbolTable_Click()
  Dim strSymbolText As Variant
  strSymbolText = Application.GetOpenFilename(FileFilter:="Text Files(*.asc),*.asc", Title:="Select Symbol Table File To Be Opened")
  TextBox14.Value = strSymbolText
End Sub

Private Sub btnTestRunAutoFill_Click()
  UserForm1.Show
  Unload Me
End Sub

Private Sub btnWR_X_SBO_Click()
  Dim strrack As Variant
  strrack = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select WR_X_SBO File To Be Opened")
  TextBox10.Value = strrack
End Sub



Private Sub CheckBox1_Click()

End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub UserForm_Click()

End Sub
