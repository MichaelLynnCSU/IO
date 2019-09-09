VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPCS7SOEPLC 
   Caption         =   "File Input"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12225
   OleObjectBlob   =   "frmPCS7SOEPLC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPCS7SOEPLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCH_AI_Signal_Click()
  Dim strCH_AI_Singals As Variant
  strCH_AI_Singals = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_AI_Signals File To Be Opened")
  TextBox3.Value = strCH_AI_Singals
End Sub

Private Sub btnCH_DI_Signal_Click()
  Dim strDigit As Variant
  strDigit = Application.GetOpenFilename(FileFilter:="Text Files(*.csv),*.csv", Title:="Select New Digit File To Be Opened")
  TextBox4.Value = strDigit
End Sub

Private Sub btnChangePLC_Click()
If cmbBoxPLCSelection.Value = "PCS7 PLC Without RTU" Then
    Unload Me
    frmPCS7PLCWithoutRTU.Show
    
  ElseIf cmbBoxPLCSelection.Value = "PCS7 PLC With One or More RTU" Then
    Unload Me
    frmPCS7PLCWithOneOrMoreRTU.Show
    
  Else
    MsgBox "Incorrect type Selected"
  End If
End Sub

Private Sub btnExit_Click()
  Unload Me
End Sub

Private Sub btnMeas_Mon_Alarming_Click()
  Dim strSymbolText As Variant
  strSymbolText = Application.GetOpenFilename(FileFilter:="Text Files(*.asc),*.asc", Title:="Select Symbol Table File To Be Opened")
  TextBox2.Value = strSymbolText
End Sub

Private Sub btnRD_X_AI1_Click()
  Dim strHWConfig As Variant
  strHWConfig = Application.GetOpenFilename(FileFilter:="Text Files(*.cfg),*.cfg", Title:="Select HW Config File To Be Opened")
  TextBox1.Value = strHWConfig
End Sub

Private Sub btnRun_Click()
  Dim blnPlaceHolder As Boolean
  Dim i As Long
  blnPlaceHolder = True
  i = 1
  For i = 1 To 4
    If Me.Controls("CheckBox" & i).Value = False Then
      blnPlaceHolder = False
      Exit For
    End If
  Next i
  
  If blnPlaceHolder = True Then
  
   ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "HW Config File"
   ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = strHWConfig
  
   ThisWorkbook.Sheets("File Paths").Cells(3, 1).Value2 = "CH_AI_Singals"
   ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = strCH_AI_Singals
    
   ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "Symbol Table File"
   ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = strSymbolText
  
   ThisWorkbook.Sheets("File Paths").Cells(14, 1).Value2 = "CH_DI_Signals_NO-NC mod"
   ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = strDigit
      
   Unload Me
  
  Else
    MsgBox "Missing required fields"
  End If

End Sub

Private Sub btnTestRunAutoFill_Click()
  ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "HW Config File"
  ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\nas-longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\Exported Data Files\NJH_HWConfig.cfg"
  
  ThisWorkbook.Sheets("File Paths").Cells(3, 1).Value2 = "CH_AI_Singals"
  ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\nas-longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\Exported Data Files\Nickajack_Plant_NJH_CH_AI_Signals.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "Symbol Table File"
  ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\nas-longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\Exported Data Files\SymbolTable.asc"
  
  ThisWorkbook.Sheets("File Paths").Cells(14, 1).Value2 = "CH_DI_Signals_NO-NC mod"
  ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "\\nas-longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\NJH\Exported Data Files\Nickajack_Plant_NJH_CH_DI_Signals_NO-NC mod.csv"
    
  Unload Me
End Sub

Private Sub UserForm_Initialize()
    cmbBoxPLCSelection.List = Array("PCS7 PLC Without RTU", "PCS7 PLC With One or More RTU", "PCS7 SOE PLC")
    cmbBoxPLCSelection.Value = "PCS7 PLC With One or More RTU"
End Sub
