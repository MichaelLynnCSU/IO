VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAI 
   Caption         =   "UserForm1"
   ClientHeight    =   1800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3240
   OleObjectBlob   =   "frmAI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAI_Click()
    Dim strAI As Variant
    strAI = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select RD_X_AI1 File To Be Opened")
    If strAI = False Then
        Unload frmRack
    End If
    ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
    ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = strAI
    Unload frmAI
End Sub
