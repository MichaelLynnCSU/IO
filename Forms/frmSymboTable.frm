VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSymboTable 
   Caption         =   "UserForm1"
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2670
   OleObjectBlob   =   "frmSymboTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSymboTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSymbolTable_Click()
    Dim strSymbolText As Variant
    strSymbolText = Application.GetOpenFilename(FileFilter:="Text Files(*.asc),*.asc", Title:="Select Symbol Table File To Be Opened")
    If strSymbolText = False Then
        Exit Sub
    End If
    ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "Symbol Table File"
    ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = strSymbolText
    Unload frmSymboTable
End Sub
