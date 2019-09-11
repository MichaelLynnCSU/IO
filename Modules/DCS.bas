Attribute VB_Name = "DCS"
Sub DCS()

        Set wb = ThisWorkbook
        Dim myPath As String
        myPath = ThisWorkbook.Path
        
        'Pull over block
        Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws2.Name = "Signal Connections"
        'frmCH_AI_Signals.Show
        Set wb2 = Workbooks.Open(myPath & "\Exported Data Files\Nickajack_Plant_NJH_CH_AI_Signals.csv")
        wb2.Sheets(1).Range("K:K").Copy Destination:=wb.Sheets("Signal Connections").Range("A1")
        wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Signal Connections").Range("B1")
        wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Signal Connections").Range("C1")
        wb2.Close
        
                'Pull over block
        Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws2.Name = "Range Connections"
        'frmCH_AI_Signals.Show
        Set wb2 = Workbooks.Open(myPath & "\Exported Data Files\Nickajack_Plant_NJH_CH_AI_Ranges.csv")
        wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Range Connections").Range("A1")
        wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Range Connections").Range("B1")
        wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("Range Connections").Range("C1")
        wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("Range Connections").Range("D1")
        wb2.Sheets(1).Range("M:M").Copy Destination:=wb.Sheets("Range Connections").Range("E1")
        wb2.Close
        
        'Pull over block
        Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
            ws2.Name = "Limits Connections"
        'frmCH_AI_Signals.Show
        Set wb2 = Workbooks.Open(myPath & "\Exported Data Files\Nickajack_Plant_NJH_CH_AI_Meas_Mon_Alarming.csv")
        wb2.Sheets(1).Range("D:D").Copy Destination:=wb.Sheets("Limits Connections").Range("A1")
        wb2.Sheets(1).Range("B:B").Copy Destination:=wb.Sheets("Limits Connections").Range("B1")
        wb2.Sheets(1).Range("F:F").Copy Destination:=wb.Sheets("Limits Connections").Range("C1")
        wb2.Sheets(1).Range("J:J").Copy Destination:=wb.Sheets("Limits Connections").Range("D1")
        wb2.Close
        
        Set oConn = CreateObject("ADODB.Connection")
        oConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName _
        & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
        oConn.Open
        
        
        Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws2.Name = "Output"
        Sheets("Output").Cells(1, 1).Value = "Symbol"
        Sheets("Output").Cells(1, 2).Value = "Chart"
        Sheets("Output").Cells(1, 3).Value = "Block"
        Sheets("Output").Cells(1, 4).Value = "IO_Name"
        Sheets("Output").Cells(1, 5).Value = "Value"
        
        
        Set mrs = CreateObject("ADODB.Recordset")
        Set mrs2 = CreateObject("ADODB.Recordset")
        
        sSQLSting = "" _
        & "SELECT Signal.[Signal], Signal.[Chart], Signal.[Block], Range.[I/O name], Range.[Value], Range.[nBlock] " _
        & "From ([Signal Connections$] AS Signal " _
        & "INNER JOIN [Range Connections$] AS Range ON Signal.[Chart] = Range.[Chart] AND Signal.[Block] = Range.[Block]) " _
        & "WHERE Signal.[Signal] = ""U1 GEN BEARING X-AXIS"""
        mrs.Open sSQLSting, oConn
        ActiveSheet.Range("A2").CopyFromRecordset mrs
        mrs.Close
        
        
        sSQLSting = "" _
        & "SELECT Signal.[Signal], Range.[Chart], Range.[nBlock], Limit.[I/O name], Limit.[Value] " _
        & "From (([Signal Connections$] AS Signal " _
        & "INNER JOIN [Range Connections$] AS Range ON Signal.[Chart] = Range.[Chart] AND Signal.[Block] = Range.[Block]) " _
        & "INNER JOIN [Limits Connections$] AS Limit ON Range.[Chart] = Limit.[Chart] AND Range.[nBlock] = Limit.[Block]) " _
        & "WHERE Signal.[Signal] = ""U1 GEN BEARING X-AXIS"""
        mrs2.Open sSQLSting, oConn
        ActiveSheet.Range("A6").CopyFromRecordset mrs2
        mrs2.Close
        
        
        
        
        
        oConn.Close

End Sub
