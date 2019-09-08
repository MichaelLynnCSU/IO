VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSOE_Message 
   Caption         =   "UserForm1"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3240
   OleObjectBlob   =   "frmSOE_Message.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSOE_Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSOE_Message_Click()
    Dim strSOE_Message As Variant
    strSOE_Message = Application.GetOpenFilename(FileFilter:="Excel Files (*.csv), *.csv", Title:="Select RD_X_SOE_Message File To Be Opened")
    If strSOE_Message = False Then
        Unload frmRack
    End If
    ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
    ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = strSOE_Message
    Unload frmSOE_Message
End Sub
