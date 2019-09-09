VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1
   Caption         =   "UserForm1"
   ClientHeight    =   2145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4995
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()

If ComboBox1.Value = "NJH" Then

 blnPlaceHolder = False

 'Hardcoded inputs for testing purposes only. Delete when finished.
  Dim myPath As String
  path = ThisWorkbook.path
  ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "HW Config File"
  ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = path & "\NJH_HWConfig.cfg"

  ThisWorkbook.Sheets("File Paths").Cells(3, 1).Value2 = "CH_AI_Singals"
  ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = path & "\Nickajack_Plant_NJH_CH_AI_Signals.csv"

  ThisWorkbook.Sheets("File Paths").Cells(4, 1).Value2 = "CH_AI_Ranges"
  ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = path & "\Nickajack_Plant_NJH_CH_AI_Ranges.csv"

  ThisWorkbook.Sheets("File Paths").Cells(5, 1).Value2 = "Meas_Mon_Alarming"
  ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = path & "\Nickajack_Plant_NJH_CH_AI_Meas_Mon_Alarming.csv"

  ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "Symbol Table File"
  ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = path & "\SymbolTable.asc"

  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = path & "\Nickajack_Plant_NJH_WR_X_SBO.csv"

  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = path & "\Nickajack_Plant_NJH_RD_X_AI1.csv"

  ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = path & "\Nickajack_Plant_NJH_RD_X_SOE.csv"

  ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
  ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = path & "\Nickajack_Plant_NJH_RD_X_SOE_Messages.csv"

  ThisWorkbook.Sheets("File Paths").Cells(11, 1).Value2 = "CH_DI_Singals"
  ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = path & "\Nickajack_Plant_NJH_CH_DI_Signals.csv"

  ThisWorkbook.Sheets("File Paths").Cells(12, 1).Value2 = "CH_DI"
  ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = path & "\Nickajack_Plant_NJH_CH_DI.csv"

  ThisWorkbook.Sheets("File Paths").Cells(13, 1).Value2 = "Message_Block"
  ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = path & "\Nickajack_Plant_NJH_CH_DI_Message_Block.csv"

  ThisWorkbook.Sheets("File Paths").Cells(14, 1).Value2 = "CH_DI_Signals_NO-NC mod"
  ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = path & "\Nickajack_Plant_NJH_CH_DI_Signals_NO-NC mod.csv"

  Unload Me

 End If

     If ComboBox1.Value = "CHH_Master_RED" Then

      blnPlaceHolder = False

 ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\MASTER_RED\chh_master_red_HWCONFIG.cfg"
 ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_Ch_AI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_Ch_AI_Ranges.csv"
 ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_AI_MEAS_MON_ALARMING.csv"
 ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\MASTER_RED\CHH_MASTER_RED_SYMBOLTABLE.asc"
 ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI.csv"
 ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_DI_MESSAGE_BLOCK.csv"
 ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI_Signals_NO-NC mod.csv"


  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
  ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"
    Unload Me
 End If


 If ComboBox1.Value = "CHH_SOE_Master" Then

  blnPlaceHolder = False

 ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\SOE_MASTER\chh_soe_master_hwconfig.cfg"
 ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_Ch_AI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_Ch_AI_Ranges.csv"
 ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_AI_MEAS_MON_ALARMING.csv"
 ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\SOE_MASTER\CHH_SOE_MASTER_SYMBOLTABLE.asc"
 ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI.csv"
 ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_DI_MESSAGE_BLOCK.csv"
 ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI_Signals_NO-NC mod.csv"


  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\SOE_MASTER\CHH_CHH_SOE_MASTER_C_Parameters.csv"

  ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
  ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\SOE_MASTER\CHH_CHH_SOE_MASTER_Messages.csv"
    Unload Me
 End If

 If ComboBox1.Value = "CHH_Unit_1-2_RED" Then

  blnPlaceHolder = False

 ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\UNIT_1-2_RED\chh_unit_1-2_red_hwconfig.cfg"
 ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_Ch_AI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_Ch_AI_Ranges.csv"
 ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_AI_MEAS_MON_ALARMING.csv"
 ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\UNIT_1-2_RED\CHH_UNIT_1-2_RED_SYMBOLTABLE.asc"
 ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI.csv"
 ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_DI_MESSAGE_BLOCK.csv"
 ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI_Signals_NO-NC mod.csv"


  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
  ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"
    Unload Me
 End If

 If ComboBox1.Value = "CHH_Unit_3-4_RED" Then

  blnPlaceHolder = False

 ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\UNIT_3-4_RED\chh_unit_3_4_red_hwconfig.cfg"
 ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_Ch_AI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_Ch_AI_Ranges.csv"
 ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_AI_MEAS_MON_ALARMING.csv"
 ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\UNIT_3-4_RED\CHH_UNIT_3-4_RED_SYMBOLTABLE.asc"
 ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI.csv"
 ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_DI_MESSAGE_BLOCK.csv"
 ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\CHH_CH_DI_Signals_NO-NC mod.csv"


  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
  ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"
    Unload Me
 End If

 If ComboBox1.Value = "TFH" Then

  blnPlaceHolder = False

 ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\testfiles\test_hwconfig.cfg"
 ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\test_analog_DI\CHH_Ch_AI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\test_analog_DI\CHH_Ch_AI_Ranges.csv"
 ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\test_analog_DI\CHH_CH_AI_MEAS_MON_ALARMING.csv"
 ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\testfiles\test_SYMBOLTABLE.asc"
 ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\test_analog_DI\CHH_CH_DI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\test_analog_DI\CHH_CH_DI.csv"
 ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\test_analog_DI\CHH_DI_MESSAGE_BLOCK.csv"
 ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\test_analog_DI\CHH_CH_DI_Signals_NO-NC mod.csv"


  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\Exported Data Files\CHH_TFH_ComRTU_WR_X_SBO16.csv"

  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\Exported Data Files\CHH_TFH_ComRTU_RD_X_AI16.csv"

  ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\Exported Data Files\CHH_TFH_ComRTU_RD_X_SOE32T.csv"

  ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
  ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\TFH\Exported Data Files\CHH_TFH_ComRTU_RD_X_SOE32T_Messages.csv"

 Unload Me

 End If

  If ComboBox1.Value = "AS1" Then

   blnPlaceHolder = False

 ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS1\as1_h.cfg"
 ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS1\IOList_Mods\PV_AS1_AI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS1\IOList_Mods\PV_AS1_AI_Range.csv"
 ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS1\PV_AS1__AI_MonAnL_Alarms.csv"
 ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS1\AS1_Symbols.asc"
 ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS1\IOList_Mods\PV_AS1_DI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS1\IOList_Mods\PV_AS1_DI.csv"
 ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS1\PV_AS1_DI_MonDiL_Alarms.csv"
 ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"


  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
  ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

 End If


   If ComboBox1.Value = "AS2" Then

   blnPlaceHolder = False

 ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS2\as2_h.cfg"
 ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS2\IOList_Mods\PV_AS2_AI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS2\IOList_Mods\PV_AS2_AI_Range.csv"
 ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS2\PV_AS2__MonAnL_Alarms.csv"
 ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS2\AS2_Symbols.asc"
 ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS2\IOList_Mods\PV_AS2_DI_Signals.csv"
 ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS2\IOList_Mods\PV_AS2_DI.csv"
 ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\APV\Export\AS2\PV_AS2_MonDiL_Alarms.csv"
 ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"


  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"

  ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
  ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = "\\NAS-Longmont\Project\Customer\LSI\LSI001 - TVA IROCS\07 - IO List Tool\CHH\Exported Data Files\testfiles\test__csv.csv"
    Unload Me
 End If
 Unload Me

 If ComboBox1.Value = "PAA" Then

blnPlaceHolder = True

  ThisWorkbook.Sheets("File Paths").Cells(2, 1).Value2 = "HW Config File"
  ThisWorkbook.Sheets("File Paths").Cells(2, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_CH_HWConfig.cfg"

  ThisWorkbook.Sheets("File Paths").Cells(3, 1).Value2 = "CH_AI_Singals"
  ThisWorkbook.Sheets("File Paths").Cells(3, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_CH_AI_Signals.csv"

  ThisWorkbook.Sheets("File Paths").Cells(4, 1).Value2 = "CH_AI_Ranges"
  ThisWorkbook.Sheets("File Paths").Cells(4, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_CH_AI_Ranges.csv"

  ThisWorkbook.Sheets("File Paths").Cells(5, 1).Value2 = "Meas_Mon_Alarming"
  ThisWorkbook.Sheets("File Paths").Cells(5, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_CH_AI_Meas_Mon_Alarming.csv"

  ThisWorkbook.Sheets("File Paths").Cells(6, 1).Value2 = "Symbol Table File"
  ThisWorkbook.Sheets("File Paths").Cells(6, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\SymbolTable.asc"

  ThisWorkbook.Sheets("File Paths").Cells(7, 1).Value2 = "WR_X_SBO - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(7, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_WR_X_SBO.csv"

  ThisWorkbook.Sheets("File Paths").Cells(8, 1).Value2 = "RD_X_AI1 - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(8, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_RD_X_AI1.csv"

  ThisWorkbook.Sheets("File Paths").Cells(9, 1).Value2 = "RD_X_SOE - Rack 1"
  ThisWorkbook.Sheets("File Paths").Cells(9, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_RD_X_SOE.csv"

  ThisWorkbook.Sheets("File Paths").Cells(10, 1).Value2 = "RD_X_SOE_Message"
  ThisWorkbook.Sheets("File Paths").Cells(10, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_RD_X_SOE_Messages.csv"

  ThisWorkbook.Sheets("File Paths").Cells(11, 1).Value2 = "CH_DI_Singals"
  ThisWorkbook.Sheets("File Paths").Cells(11, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_CH_DI_Signals.csv"

  ThisWorkbook.Sheets("File Paths").Cells(12, 1).Value2 = "CH_DI"
  ThisWorkbook.Sheets("File Paths").Cells(12, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_CH_DI.csv"

  ThisWorkbook.Sheets("File Paths").Cells(13, 1).Value2 = "Message_Block"
  ThisWorkbook.Sheets("File Paths").Cells(13, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_Message_Block.csv"

  ThisWorkbook.Sheets("File Paths").Cells(14, 1).Value2 = "CH_DI_Signals_NO-NC mod"
  ThisWorkbook.Sheets("File Paths").Cells(14, 2).Value2 = "C:\Users\MIchael.Lynn\Desktop\latest_Import\NJH_CH_DI_Signals_NO-NC mod.csv"

  Unload Me

 End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()
    ComboBox1.List = Array("NJH", "CHH_Master_RED", "CHH_SOE_Master", "CHH_Unit_1-2_RED", "CHH_Unit_3-4_RED", "TFH", "AS1", "AS2", "PAA")
    ComboBox1.Value = "NJH"
End Sub
