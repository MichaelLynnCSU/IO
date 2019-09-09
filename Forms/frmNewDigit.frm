VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNewDigit 
   Caption         =   "UserForm1"
   ClientHeight    =   2355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmNewDigit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNewDigit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub frmDigit_Click()
    Dim strDigit As Variant
    strDigit = Application.GetOpenFilename(FileFilter:="Text Files(*.csv),*.csv", Title:="Select New Digit File To Be Opened")
    If strDigit = False Then
        Exit Sub
    End If
    ThisWorkbook.Sheets("File Paths").Cells(14, 1).Value2 = "New Digital File"
    ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = strDigit
    Unload frmNewDigit
End Sub

