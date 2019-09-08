VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSOE 
   Caption         =   "UserForm1"
   ClientHeight    =   1560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2880
   OleObjectBlob   =   "frmSOE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSOE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSOE_Click()
    Dim strSOE As Variant
    strSOE = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select RD_X_SOE File To Be Opened")
    If strSOE = False Then
        Unload frmSOE
    End If
    ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
    ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = strSOE
    Unload frmSOE
End Sub
