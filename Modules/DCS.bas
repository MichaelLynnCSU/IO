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
        wb2.Close
        
        
        Set oConn = CreateObject("ADODB.Connection")
        oConn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName _
& ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
        oConn.Open
           
        sSQLSting = "" _
        & "SELECT Signal.[Signal], Signal.[Chart], Signal.[Block], Range.[I/O name], Range.[Value] " _
        & "From [Signal Connections$] AS Signal INNER JOIN [Range Connections$] AS Range " _
        & "ON Signal.[Chart] = Range.[Chart] AND Signal.[Block] = Range.[Block] WHERE Signal.[Signal] = ""U3 GATE LIMIT"""
        Set mrs = CreateObject("ADODB.Recordset")
        mrs.Open sSQLSting, oConn
        
        Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws2.Name = "Output"
        ActiveSheet.Range("A2").CopyFromRecordset mrs
        mrs.Close
        
        oConn.Close
        '        sSQLSting = "" _
'        & "SELECT Signal.[Signal], Signal.[Chart], Signal.[Block], Range.[I/O name], Range.[Value] " _
'        & "From Nickajack_Plant_NJH_CH_AI_Signals.csv AS Signal, Nickajack_Plant_NJH_CH_AI_Ranges.csv AS Range " _
'        & "WHERE Signal.[Signal] = ""U3 GATE LIMIT"""
'        mrs.Open sSQLSting, oConn
'        ActiveSheet.Range("A1").CopyFromRecordset mrs
'        mrs.Close
End Sub
