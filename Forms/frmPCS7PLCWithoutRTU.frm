VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPCS7PLCWithoutRTU 
   Caption         =   "File Input"
   ClientHeight    =   9615.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12225
   OleObjectBlob   =   "frmPCS7PLCWithoutRTU.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPCS7PLCWithoutRTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnChangePLC_Click()
    
  If cmbBoxPLCSelection.Value = "PCS7 PLC With One or More RTU" Then
    Unload Me
    frmPCS7PLCWithOneOrMoreRTU.Show
    
  ElseIf cmbBoxPLCSelection.Value = "PCS7 SOE PLC" Then
    Unload Me
    frmPCS7SOEPLC.Show
    
  Else
    MsgBox "Incorrect type Selected"
  End If
End Sub

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnHWConfig_Click()
  Dim strHWConfig As Variant
  strHWConfig = Application.GetOpenFilename(FileFilter:="Text Files(*.cfg),*.cfg", Title:="Select HW Config File To Be Opened")
  TextBox1.Value = strHWConfig
End Sub

Private Sub btnMessage_Click()
  Dim strDIAlarm As Variant
  strDIAlarm = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select Message_Block File To Be Opened")
  TextBox3.Value = strDIAlarm
End Sub

Private Sub btnParamExportAIRanges_Click()
  Dim strRanges As Variant
  strRanges = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_AI_Ranges File To Be Opened")
  TextBox4.Value = strRanges
End Sub

Private Sub btnParamExportAORanges_Click()
  Dim strParamAORanges As Variant
  strParamAORanges = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select Meas_Mon_Alarming File To Be Opened")
  TextBox5.Value = strParamAORanges
End Sub

Private Sub btnParamlarm_Click()
  Dim strAlarm As Variant
  strAlarm = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select Meas_Mon_Alarming File To Be Opened")
  TextBox6.Value = strAlarm
End Sub

Private Sub btnRun_Click()
Dim blnPlaceHolder As Boolean
  Dim i As Long
  blnPlaceHolder = True
  i = 1
  For i = 1 To 7
    If Me.Controls("CheckBox" & i).Value = False Then
      blnPlaceHolder = False
      Exit For
    End If
  Next i
  
  If blnPlaceHolder = True Then
  
   ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "HW Config File"
   ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = strHWConfig
    
   ThisWorkbook.Sheets("File Paths").Cells(4, 1).Value2 = "CH_AI_Ranges"
   ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = strRanges
  
   ThisWorkbook.Sheets("File Paths").Cells(5, 1).Value2 = "Meas_Mon_Alarming"
   ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = strAlarm
  
   ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "Symbol Table File"
   ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = strSymbolText
  
   ThisWorkbook.Sheets("File Paths").Cells(13, 1).Value2 = "Message_Block"
   ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = strDIAlarm
   
   ThisWorkbook.Sheets("File Paths").Cells(16, 1).Value2 = "Parameter Export"
   ThisWorkbook.Sheets("File Paths").Cells(16, 2).Value2 = strParamAORanges
   
   ThisWorkbook.Sheets("File Paths").Cells(16, 1).Value2 = "Signal Export"
   ThisWorkbook.Sheets("File Paths").Cells(16, 2).Value2 = strSignalExportASPLC
   
   Unload Me
  
  Else
    MsgBox "Missing required fields"
  End If
End Sub

Private Sub btnSignalExportAS_Click()
  Dim strSignalExportASPLC As Variant
  strSignalExportASPLC = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select Meas_Mon_Alarming File To Be Opened")
  TextBox6.Value = strSignalExportASPLC
End Sub

Private Sub btnSymbolTable_Click()
  Dim strSymbolText As Variant
  strSymbolText = Application.GetOpenFilename(FileFilter:="Text Files(*.asc),*.asc", Title:="Select Symbol Table File To Be Opened")
  TextBox2.Value = strSymbolText
End Sub

Private Sub btnTestRunAutoFill_Click()
   ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "HW Config File"
   ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\nas-longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\Exported Data Files\NJH_HWConfig.cfg"
    
   ThisWorkbook.Sheets("File Paths").Cells(4, 1).Value2 = "CH_AI_Ranges"
   ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\nas-longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\Exported Data Files\Nickajack_Plant_NJH_CH_AI_Ranges.csv"
  
   ThisWorkbook.Sheets("File Paths").Cells(5, 1).Value2 = "Meas_Mon_Alarming"
   ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\nas-longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\Exported Data Files\Nickajack_Plant_NJH_Meas_Mon_Alarming.csv"
  
   ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "Symbol Table File"
   ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\nas-longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\Exported Data Files\SymbolTable.asc"
  
   ThisWorkbook.Sheets("File Paths").Cells(13, 1).Value2 = "Message_Block"
   ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "\\nas-longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\Exported Data Files\Nickajack_Plant_NJH_Message_Block.csv"
   
   ThisWorkbook.Sheets("File Paths").Cells(15, 1).Value2 = "Parameter Export"
   ThisWorkbook.Sheets("File Paths").Cells(15, 2).Value2 = "File missing"
   
   ThisWorkbook.Sheets("File Paths").Cells(16, 1).Value2 = "Signal Export"
   ThisWorkbook.Sheets("File Paths").Cells(16, 2).Value2 = "File missing"
   
   Unload Me
   
End Sub

Private Sub UserForm_Initialize()
    cmbBoxPLCSelection.List = Array("PCS7 PLC Without RTU", "PCS7 PLC With One or More RTU", "PCS7 SOE PLC")
    cmbBoxPLCSelection.Value = "PCS7 PLC With One or More RTU"
End Sub
