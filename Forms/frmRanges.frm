VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRanges 
   Caption         =   "Select CH_AI_Ranges File"
   ClientHeight    =   1584
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3030
   OleObjectBlob   =   "frmRanges.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRanges_Click()
    Dim strRanges As Variant
    strRanges = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_AI_Ranges File To Be Opened")
    If strRanges = False Then
        Exit Sub
    End If
    ThisWorkbook.Sheets("File Paths").Cells(4, 1).Value2 = "CH_AI_Ranges"
    ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = strRanges
    Unload frmRanges
End Sub
