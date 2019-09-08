VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDIAlarm 
   Caption         =   "UserForm1"
   ClientHeight    =   1320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2640
   OleObjectBlob   =   "frmDIAlarm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDIAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDIAlarm_Click()
    Dim strDIAlarm As Variant
    strDIAlarm = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select Message_Block File To Be Opened")
    If strDIAlarm = False Then
        Exit Sub
    End If
    ThisWorkbook.Sheets("File Paths").Cells(13, 1).Value2 = "Message_Block"
    ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = strDIAlarm
    Unload frmDIAlarm
End Sub
