Attribute VB_Name = "Module1"
Declare Function OPENCOM Lib "Port" (ByVal A$) As Integer
Declare Sub CLOSECOM Lib "Port" ()
Declare Sub SENDBYTE Lib "Port" (ByVal b%)
Declare Function READBYTE Lib "Port" () As Integer
Declare Sub DTR Lib "Port" (ByVal b%)
Declare Sub RTS Lib "Port" (ByVal b%)
Declare Sub TXD Lib "Port" (ByVal b%)
Declare Function CTS Lib "Port" () As Integer
Declare Function DSR Lib "Port" () As Integer
Declare Function RI Lib "Port" () As Integer
Declare Function DCD Lib "Port" () As Integer
Declare Sub DELAY Lib "Port" (ByVal b%)
Declare Sub TIMEINIT Lib "Port" ()
Declare Sub TIMEINITUS Lib "Port" ()
Declare Function TIMEREAD Lib "Port" () As Long
Declare Function TIMEREADUS Lib "Port" () As Long
Declare Sub DELAYUS Lib "Port" (ByVal l As Long)
Declare Sub REALTIME Lib "Port" (ByVal i As Boolean)

Public Sub Pause(Dauer)
start = Timer
Do While Timer < start + Dauer
DoEvents
Loop
End Sub


