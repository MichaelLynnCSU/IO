VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAlarm 
   Caption         =   "UserForm1"
   ClientHeight    =   1560
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3000
   OleObjectBlob   =   "frmAlarm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlarm_Click()
    Dim strAlarm As Variant
    strAlarm = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select Meas_Mon_Alarming File To Be Opened")
    If strAlarm = False Then
        Exit Sub
    End If
    ThisWorkbook.Sheets("File Paths").Cells(5, 1).Value2 = "Meas_Mon_Alarming"
    ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = strAlarm
    Unload frmAlarm
End Sub
