VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDI 
   Caption         =   "UserForm1"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2670
   OleObjectBlob   =   "frmDI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDI_Click()
    Dim strDI As Variant
    strDI = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_DI File To Be Opened")
    If strDI = False Then
        Exit Sub
    End If
    ThisWorkbook.Sheets("File Paths").Cells(12, 1).Value2 = "CH_DI"
    ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = strDI
    Unload frmDI
End Sub
