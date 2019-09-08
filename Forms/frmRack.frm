VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRack 
   Caption         =   "UserForm1"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3255
   OleObjectBlob   =   "frmRack.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRack_Click()
    Dim strrack As Variant
    strrack = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select WR_X_SBO File To Be Opened")
    If strrack = False Then
        Unload frmRack
    End If
    ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
    ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = strrack
    Unload frmRack
End Sub
