VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHWConfig 
   Caption         =   "Browse HW Config File"
   ClientHeight    =   1704
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3015
   OleObjectBlob   =   "frmHWConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHWConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHWConfigFile_Click()
    Dim strHWConfig As Variant
    strHWConfig = Application.GetOpenFilename(FileFilter:="Text Files(*.cfg),*.cfg", Title:="Select HW Config File To Be Opened")
    If strHWConfig = False Then
        Exit Sub
    End If
    ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "HW Config File"
    ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = strHWConfig
    Unload frmHWConfig
End Sub
