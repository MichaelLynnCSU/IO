VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSaveAs 
   Caption         =   "UserForm1"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3165
   OleObjectBlob   =   "frmSaveAs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSaveAs_Click()
    If Len(TextBox1) > 0 Then
        ActiveWorkbook.SaveAs Filename:=Left(Sheets("File Paths").Cells(2, 2), InStrRev(Sheets("File Paths").Cells(2, 2), "\")) & TextBox1.Value & "_" & TextBox2.Value & "_IO List Report_" & Format(Date, "MM-DD-YYYY") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = False
    Else:
        ActiveWorkbook.SaveAs Filename:=Left(Sheets("File Paths").Cells(2, 2), InStrRev(Sheets("File Paths").Cells(2, 2), "\")) & ComboBox1.Value & "_" & TextBox2.Value & "_IO List Report_" & Format(Date, "MM-DD-YYYY") & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = False
    End If
    ActiveWorkbook.Sheets("Report").Name = TextBox2.Value
    Unload frmSaveAs
End Sub

Private Sub UserForm_Initialize()
    ComboBox1.List = Array("Anadarko", "Colorado Mills", "FMC", "Genesis", "HEXCEL", "Hidd", "Innovative Contorls", "InteGrow", "James Engineering", _
        "LADWP", "Linde", "LSI", "Matheson", "Messer", "Norican", "ParkPlus", "PGW", "Pigler Automation", "Prime Control", "Siemens Canada", _
        "Siemens DEMAG", "SIMPLOT", "Sinclair", "Sunshine Paper Co", "Tesla", "TRONOX", "TVA", "UR Energy", "Usace", "VSI Parylene", "Other - Specify Below")
    ComboBox1.Value = "LSI"
End Sub
