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

    If Me.Controls("CheckBox1").Value = False Then
      blnPlaceHolder = False
    Else
        blnPlaceHolder = True
    End If
    
If ComboBox1.Value = "DCS_NJH" Then
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 1).Value2 = "NJH-Info"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 2).Value2 = "HDCC_NJH_Info"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 3).Value2 = "NJH-RTU-Info"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 4).Value2 = "HDCC_NJH_RTU_Info"
  Unload Me
End If

If ComboBox1.Value = "DCS_CHH" Then
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 1).Value2 = "CHH_Info"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 2).Value2 = "HDCC_CHH_Info"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 3).Value2 = "CHH-RTU-Info"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 4).Value2 = "HDCC_CHH_RTU_Info"
    Unload Me
End If

If ComboBox1.Value = "DCS_TFH" Then
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 1).Value2 = "TFH_Info"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 2).Value2 = "HDCC_TFH-Info"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 3).Value2 = "TFH_RTU_Info"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 4).Value2 = "HDCC_TFH_RTU_Info"
    Unload Me
End If

If ComboBox1.Value = "DCS2_NJH" Then
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 1).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\WIP Michael L\IO_List_WIP\WIP\116\DCS2\NJH IO List Rev B"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\WIP Michael L\IO_List_WIP\WIP\116\DCS2\CHH IO List rev B"
    ThisWorkbook.Sheets("Check_blocks").Cells(1, 3).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\WIP Michael L\IO_List_WIP\WIP\116\DCS2\TFH IO List rev B"
  Unload Me
End If

End Sub

Private Sub UserForm_Click()

End Sub


Private Sub UserForm_Initialize()
    ComboBox1.List = Array("DCS_NJH", "DCS_CHH", "DCS_TFH", "DCS2_NJH")
    ComboBox1.Value = "DCS_NJH"
End Sub

