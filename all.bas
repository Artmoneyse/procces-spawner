Attribute VB_Name = "all"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Wait As Boolean
Public BitMaskConf As Long
Public StartFileOnExit As String
Public CloseOnFast As Boolean 'Если быстро оффнули то вырубить

Public Function readParam(param As String) As Boolean
readParam = False
''''''''''''''''''''''''''
If InStr(1, param, "-th_wait") > 0 Then
Wait = True
readParam = True
End If
''''''''''''''''''''''''''
If InStr(1, param, "-th_closeonfast") > 0 Then
CloseOnFast = True
readParam = True
End If
'''''''''''''''''''''''''''
If InStr(1, param, "-th_bitmask=") > 0 Then
BitMaskConf = Replace(param, "-th_bitmask=", "")
readParam = True
End If
'''''''''''''''''''''''''''
If InStr(1, param, "-th_startfileonexit=") > 0 Then
StartFileOnExit = Replace(param, "-th_startfileonexit=", "")
readParam = True
End If
'''''''''''''''''''''''''''
End Function
