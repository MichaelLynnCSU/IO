VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PAAUI 
   Caption         =   "PAAUI"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "PAAUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PAAUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

If ComboBox1.Value = "NJH" Then

 'Hardcoded inputs for testing purposes only. Delete when finished.


  ThisWorkbook.Sheets("File Paths").Cells(1, 1).Value2 = "CH_AI_Singals"
  ThisWorkbook.Sheets("File Paths").Cells(1, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\PAA\NJH_CH_AI_Signals.csv"

  ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "CH_AI_Ranges"
  ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\PAA\NJH_CH_AI_Ranges.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(3, 1).Value2 = "CH_AO_Ranges"
  ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\PAA\NJH_CH_AO_Ranges.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(4, 1).Value2 = "Meas_Mon_Alarming"
  ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\PAA\NJH_Meas_Mon_Alarming.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(5, 1).Value2 = "CH_DI_Singals"
  ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\PAA\NJH_CH_DI_Signals.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "CH_DI"
  ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\PAA\NJH_CH_DI.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "CH_DO"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\PAA\NJH_CH_DO.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "Message_Block"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\PAA\NJH_Message_Block.csv"

  Unload Me
End If

If ComboBox1.Value = "CHH" Then

 'Hardcoded inputs for testing purposes only. Delete when finished.

 ' converts for all CH, the hwconfig and symbol dont matter
  ThisWorkbook.Sheets("File Paths").Cells(1, 1).Value2 = "CH_AI_Singals"
  ThisWorkbook.Sheets("File Paths").Cells(1, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_AI_Signals.csv"

  ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "CH_AI_Ranges"
  ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_AI_Ranges.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(3, 1).Value2 = "CH_AO_Ranges"
  ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_AO_Ranges.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(4, 1).Value2 = "Meas_Mon_Alarming"
  ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_AI_MEAS_MON_ALARMING.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(5, 1).Value2 = "CH_DI_Singals"
  ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI_Signals.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "CH_DI"
  ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "CH_DO"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"
  
  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "Message_Block"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_Message_Block.csv"

  Unload Me
End If
End Sub

Private Sub UserForm_Click()

End Sub


Private Sub UserForm_Initialize()
    ComboBox1.List = Array("NJH", "CHH")
    ComboBox1.Value = "NJH"
End Sub


