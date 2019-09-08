VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCH_AI_Signals 
   Caption         =   "Browse CH_AI_Signals File"
   ClientHeight    =   1470
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   2910
   OleObjectBlob   =   "frmCH_AI_Signals.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCH_AI_Signals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCH_AI_Signals_Click()
    Dim strCH_AI_Singals As Variant
    strCH_AI_Singals = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_AI_Signals File To Be Opened")
    If strCH_AI_Singals = False Then
        Exit Sub
    End If
    ThisWorkbook.Sheets("File Paths").Cells(3, 1).Value2 = "CH_AI_Singals"
    ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = strCH_AI_Singals
    Unload frmCH_AI_Signals
End Sub
