Attribute VB_Name = "Module1"
Sub subTest()
Dim current_range_interconnetion_block As String '"BK2_TEMP\PH_A161kV.U"GEN3_LOAD_RED\3.T1_XFMR"GEN4_LOAD_RED\4.T1_XFMR"AI2_PT272_287\SLOT18.V2


Dim intEndPos As Integer
Dim intStartPos As Integer
Dim strBlock As String

  current_range_interconnetion_block = "BK2_TEMP\PH_A161kV.UGEN3_LOAD_RED\3.T1_XFMRGEN4_LOAD_RED\4.T1_XFMRAI2_PT272_287\SLOT18.V2"
  'If InStr(current_range_interconnetion_block, ".U") > 0 Then
    intEndPos = InStr(current_range_interconnetion_block, ".U")
    intStartPos = InStrRev(current_range_interconnetion_block, "\", intEndPos)
    strBlock = Mid(current_range_interconnetion_block, intStartPos + 1, intEndPos - intStartPos - 1)
    Debug.Print intEndPos, intStartPos, strBlock
 ' End If
  
End Sub




Sub SubTest2()



Dim intEndPos As Integer
Dim intStartPos As Integer
Dim strBlock As String

Dim symbol_4_type As String

Dim current_symbol As String
Dim current_channel As String
Dim current_message As String
Dim current_signal As String
Dim remander_current_symbol As String


Dim AI_4_type As String

Dim current_AI_4_type As String
Dim remander_AI_4_type As String

Dim current_ID_AI_4_type As String
Dim remander_ID_AI_4_type As String

Dim current_channel_AI_4_type As Integer
Dim remander_channel_AI_4_type As String

Dim remander_messages_AI_4_type As String


Dim Range_4_type As String

Dim current_Range_4_type As String
Dim remander_Range_4_type As String

Dim current_ID_Range_4_type As String
Dim remander_ID_Range_4_type As String

Dim current_channel_Range_4_type As Integer
Dim remander_channel_Range_4_type As String

Dim remander_messages_Range_4_type As String

' part A of String alogrithm

'symbol_4_type = "Symbol I, 0, ""GEN4 COOL WTR INLET TEMP"", ""GENERATOR COOLING WATER INLET TEMP"""
symbol_4_type = "Symbol I, 1, ""U4 BRG COOL DSCH TEMP"", ""BEARING COOLING WATER DISCHARGE TEMP"""


'get symbol
intEndPos = InStr(symbol_4_type, ",")
intStartPos = 1
current_symbol = Mid(symbol_4_type, intStartPos, intEndPos - 1)
Debug.Print current_symbol

'store the rest of the string
remander_current_symbol = Mid(symbol_4_type, intEndPos + 2, Len(symbol_4_type))
Debug.Print remander_current_symbol

'get channel
intEndPos = InStr(remander_current_symbol, ",")
intStartPos = 1
current_channel = Mid(remander_current_symbol, intStartPos, intEndPos - 1)
Debug.Print current_channel

'store the rest of the string
remander_current_symbol = Mid(remander_current_symbol, intEndPos + 2, Len(symbol_4_type))
Debug.Print remander_current_symbol

'get signal
intEndPos = InStr(remander_current_symbol, ",")
intStartPos = 1
current_signal = Mid(remander_current_symbol, intStartPos + 1, intEndPos - 3)
Debug.Print current_signal

' end part A of algorithm


' part B of parse String algorithm

AI_4_type = "AI_TYPE, AI, 0, ""CURRENT_(4-WIRE_TRANSDUCER)"""
Debug.Print AI_4_type, "testing string in new module"

'get AI_type
intEndPos = InStr(AI_4_type, ",")
intStartPos = 1
current_AI_4_type = Mid(AI_4_type, intStartPos, intEndPos - 1)
Debug.Print current_AI_4_type

'store the rest of the string
remander_AI_4_type = Mid(AI_4_type, intEndPos + 2, Len(AI_4_type))
Debug.Print remander_AI_4_type


'get AI_ID_type
intEndPos = InStr(remander_AI_4_type, ",")
intStartPos = 1
current_ID_AI_4_type = Mid(remander_AI_4_type, intStartPos, intEndPos - 1)
Debug.Print current_ID_AI_4_type

'store the rest of the string
remander_ID_Range_4_type = Mid(remander_AI_4_type, intEndPos + 2, Len(remander_AI_4_type))
Debug.Print remander_ID_Range_4_type

'get AI_channel_type
intEndPos = InStr(remander_ID_Range_4_type, ",")
intStartPos = 1
current_channel_AI_4_type = Mid(remander_ID_Range_4_type, intStartPos, intEndPos - 1)
Debug.Print Trim(current_channel_AI_4_type)

'store the rest of the string, thing left is the messages
remander_messages_AI_4_type = Mid(remander_ID_AI_4_type, intEndPos + 2, Len(remander_ID_AI_4_type))
Debug.Print remander_messages_AI_4_type

' end part B of algorithm

' part C of String alogrithm

Range_4_type = "AI_RANGE, AI , 0, ""4_TO_20_MA"""
Debug.Print Range_4_type

'get AI_type
intEndPos = InStr(Range_4_type, ",")
intStartPos = 1
current_Range_4_type = Mid(Range_4_type, intStartPos, intEndPos - 1)
Debug.Print current_AI_4_type

'store the rest of the string
remander_Range_4_type = Mid(Range_4_type, intEndPos + 2, Len(Range_4_type))
Debug.Print remander_Range_4_type


'get AI_ID_type
intEndPos = InStr(remander_Range_4_type, ",")
intStartPos = 1
current_ID_Range_4_type = Mid(remander_Range_4_type, intStartPos, intEndPos - 1)
Debug.Print current_ID_Range_4_type

'store the rest of the string
remander_ID_AI_4_type = Mid(remander_Range_4_type, intEndPos + 2, Len(remander_Range_4_type))
Debug.Print remander_ID_AI_4_type

'get AI_channel_type
intEndPos = InStr(remander_ID_AI_4_type, ",")
intStartPos = 1
current_channel_Range_4_type = Mid(remander_ID_AI_4_type, intStartPos, intEndPos - 1)
Debug.Print Trim(current_channel_Range_4_type)

'store the rest of the string, thing left is the messages
remander_messages_Range_4_type = Mid(remander_ID_AI_4_type, intEndPos + 2, Len(remander_ID_AI_4_type))
Debug.Print remander_messages_Range_4_type



' end C of String alogrithm

End Sub
