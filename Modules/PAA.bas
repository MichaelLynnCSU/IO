Attribute VB_Name = "PAA"
Sub PAA()
    Set wb3 = ThisWorkbook
    Set wb = ThisWorkbook
    Dim myWorkBooks As New Class1
    
    Set wsh_Path = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    wsh_Path.Name = "File Paths"
    wsh_Path.Cells(1, 1).Value2 = "File Name"
    wsh_Path.Cells(1, 2).Value2 = "File Path"
    wsh_Path.Columns("A:A").ColumnWidth = 20
    wsh_Path.Columns("B:B").ColumnWidth = 100
    
    PAAUI.Show
    
    Dim e As BlockTypeStruct
    e.CH_DO = "PCS7DiOu"
    e.CH_DI = "PCS7DiIn"
    e.CH_AO = "PCS7AnOu"
    e.CH_AI = "PCS7AnIn"
    e.MEAS_MON = "MonAnL"
    
    'Add block to report
    'IF set methd throw error make the the file is clsed
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws3.Name = "CH_AI_Signals"
    Set wb3 = Workbooks.Open(wsh_Path.Cells(1, 2).Value2)
    wb3.Sheets(1).Range("A:AD").Copy Destination:=wb.Sheets("CH_AI_Signals").Range("A:AD")
    wb3.Close
    
    
   myWorkBooks.AI "CH_AI_Signals", e
   
    'IF set methd throw error make the the file is clsed
    'Add block to report
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws3.Name = "CH_AI_Ranges"
    Set wb3 = Workbooks.Open(wsh_Path.Cells(2, 2).Value2)
    wb3.Sheets(1).Range("A:AD").Copy Destination:=wb.Sheets("CH_AI_Ranges").Range("A:AD")
    wb3.Close
    
   myWorkBooks.AI "CH_AI_Ranges", e
        
        'IF set methd throw error make the the file is clsed
        'Add block to report
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws3.Name = "CH_AO_Ranges"
    Set wb3 = Workbooks.Open(wsh_Path.Cells(3, 2).Value2)
    wb3.Sheets(1).Range("A:AD").Copy Destination:=wb.Sheets("CH_AO_Ranges").Range("A:AD")
    wb3.Close

  myWorkBooks.AO "CH_AO_Ranges", e
   
   
    'IF set methd throw error make the the file is clsed
    'Add block to report
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws3.Name = "Meas_Mon_Alarming"
    Set wb3 = Workbooks.Open(wsh_Path.Cells(4, 2).Value2)
    wb3.Sheets(1).Range("A:AD").Copy Destination:=wb.Sheets("Meas_Mon_Alarming").Range("A:AD")
    wb3.Close

   myWorkBooks.MeasMon "Meas_Mon_Alarming", e
    
    'IF set methd throw error make the the file is clsed
    'Add block to report
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws3.Name = "CH_DI_Signals"
    Set wb3 = Workbooks.Open(wsh_Path.Cells(5, 2).Value2)
    wb3.Sheets(1).Range("A:AD").Copy Destination:=wb.Sheets("CH_DI_Signals").Range("A:AD")
    wb3.Close
    
  myWorkBooks.DI "CH_DI_Signals", e
        
    'IF set methd throw error make the the file is clsed
    'Add block to report
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws3.Name = "CH_DI"
    Set wb3 = Workbooks.Open(wsh_Path.Cells(6, 2).Value2)
    wb3.Sheets(1).Range("A:AD").Copy Destination:=wb.Sheets("CH_DI").Range("A:AD")
    wb3.Close
    
   myWorkBooks.DI "CH_DI", e
  
    'IF set methd throw error make the the file is clsed
    'Add block to report
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws3.Name = "CH_DO"
    Set wb3 = Workbooks.Open(wsh_Path.Cells(7, 2).Value2)
    wb3.Sheets(1).Range("A:AD").Copy Destination:=wb.Sheets("CH_DO").Range("A:AD")
    wb3.Close
    
  myWorkBooks.DiOu "CH_DO", e
    'IF set methd throw error make the the file is clsed
    'Add block to report
    Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws3.Name = "Message_Block"
    Set wb3 = Workbooks.Open(wsh_Path.Cells(8, 2).Value2)
    wb3.Sheets(1).Range("A:AK").Copy Destination:=wb.Sheets("Message_Block").Range("A:AK")
    wb3.Close
    
    'need to add this
   myWorkBooks.messages "Message_Block", e


End Sub
