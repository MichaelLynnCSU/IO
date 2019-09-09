VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectPLCType 
   Caption         =   "Select one of the three PLC types"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3375
   OleObjectBlob   =   "frmSelectPLCType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectPLCType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOpenPLCType_Click()

  If ComboBox1.Value = "PCS7 PLC Without RTU" Then
    Unload Me
    frmPCS7PLCWithoutRTU.Show
    
  ElseIf ComboBox1.Value = "PCS7 PLC With One or More RTU" Then
    Unload Me
    frmPCS7PLCWithOneOrMoreRTU.Show
    
  ElseIf ComboBox1.Value = "PCS7 SOE PLC" Then
    Unload Me
    frmPCS7SOEPLC.Show
    
  Else
    MsgBox "Incorrect type Selected"
  End If
  
End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub UserForm_Initialize()
    ComboBox1.List = Array("PCS7 PLC Without RTU", "PCS7 PLC With One or More RTU", "PCS7 SOE PLC")
    ComboBox1.Value = "PCS7 PLC With One or More RTU"
End Sub

