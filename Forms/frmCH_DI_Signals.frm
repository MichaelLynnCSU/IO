VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCH_DI_Signals 
   Caption         =   "UserForm1"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2910
   OleObjectBlob   =   "frmCH_DI_Signals.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCH_DI_Signals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCH_DI_Signals_Click()
    Dim strCH_DI_Singals As Variant
    strCH_DI_Singals = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select CH_DI_Signals File To Be Opened")
    If strCH_DI_Singals = False Then
        Exit Sub
    End If
    ThisWorkbook.Sheets("File Paths").Cells(11, 1).Value2 = "CH_DI_Singals"
    ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = strCH_DI_Singals
    Unload frmCH_DI_Signals
End Sub
