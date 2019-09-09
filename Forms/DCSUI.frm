VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DCSUI 
   Caption         =   "UserForm2"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "DCSUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DCSUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton1_Click()
    
If ComboBox1.Value = "DCS_NJH" Then
  Unload Me
End If

If ComboBox1.Value = "DCS_CHH" Then
    Unload Me
End If

If ComboBox1.Value = "DCS_TFH" Then
    Unload Me
End If

If ComboBox1.Value = "DCS2_NJH" Then
  Unload Me
End If

End Sub

Private Sub UserForm_Click()

End Sub


Private Sub UserForm_Initialize()
    ComboBox1.List = Array("DCS_NJH", "DCS_CHH", "DCS_TFH", "DCS2_NJH")
    ComboBox1.Value = "DCS_NJH"
End Sub

