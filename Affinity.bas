Attribute VB_Name = "Affinity"
Option Explicit
Private Declare Function WTSEnumerateProcesses Lib "wtsapi32.dll" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Function SetProcessAffinityMask Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwProcessAffinityMask As Long) As Long
Private Declare Function GetProcessAffinityMask Lib "kernel32" (ByVal hProcess As Long, ByRef ProcessMask As Long, ByRef SystemMask As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Sub WTSFreeMemory Lib "wtsapi32.dll" (ByVal pMemory As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WTS_CURRENT_SERVER_HANDLE = 0&
Private Type WTS_PROCESS_INFO
 SessionID As Long
 ProcessID As Long
 pProcessName As Long
 pUserSid As Long
 End Type

'**************************************
' Name: SetProcessAffinityMask API
' Description:This code allows you to set the process affinity on a running thread for multi processor computers.
'Get the process ID of a running thread by filename. Get the process handle by the PID. Set your custom Affinity Mask. Apply the affinity mask to the process thread.
' By: Justin Ploski (from psc cd)
'
' Inputs:Process Name is the only variable input.
'You ll need to customize the affinity mask - see comments by the MyMask assignment for information.
'
' Returns:The SetProcessAffinity API returns 0 if the affinity was set correctly. Anything <> 0 is an error.
'
' Assumes:This code is Windows API that allows you to specify a running process name and obtain the PID (Process ID). From the PID, it then obtains a handle on the process. You then set a custom affinity BitMask for the process, and pass the handle and affinity mask to the SetProcessAffinity function. Use GetCurrentProcess() API (always returns long -1) in place of the application handle to set the affinity on the current application.
'See the comments when setting the MyMask variable to customize which processors will be used.
'This is my first submission. I've been leeching off PlanetSourceCode for years, so I figured it's time to give something back. I've seen alot of questions but not many answers related to process affinity for multiprocessors. Please comment if you find this code useful.
'
' Side Effects:If using a single processor machine, the only valid process affinity is CPU0.
'**************************************

Public Function SetAffinity(lngPID As Long, AffinityBit As Long)
 
 Const PROCESS_QUERY_INFORMATION = 1024
 Const PROCESS_VM_READ = 16
 Const MAX_PATH = 260
 Const STANDARD_RIGHTS_REQUIRED = &HF0000
 Const SYNCHRONIZE = &H100000
 Const PROCESS_ALL_ACCESS = &H1F0FFF
 Const TH32CS_SNAPPROCESS = &H2&
 Const hNull = 0
 Const WIN95_System_Found = 1
 Const WINNT_System_Found = 2
 Const Default_Log_Size = 10000000
 Const Default_Log_Days = 0
 Const SPECIFIC_RIGHTS_ALL = &HFFFF
 Const STANDARD_RIGHTS_ALL = &H1F0000
 
 Dim BitMasks() As Long, NumMasks As Long, LoopMasks As Long
 Dim MyMask As Long
 Const AffinityMask As Long = &HF ' 00001111b
 
' Dim lngPID As Long
 Dim lngHwndProcess
' lngPID = GetProcessID(strImageName)
 'If lngPID = 0 Then
' MsgBox "Could not get process ID of " & strImageName, vbCritical, "Error"
' Exit Sub
' End If
 lngHwndProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, lngPID)
 If lngHwndProcess = 0 Then
 MsgBox "Could not obtain a handle for the Process ID: " & lngPID, vbCritical, "Error"
 Exit Function
 End If
 BitMasks() = GetBitMasks(AffinityMask)
'Dim i As Long
'For i = 0 To UBound(BitMasks) 'перебираем весь список
'MsgBox BitMasks(i)
'Next i
''''''''''''''''''''''''''''''''''''''''''''''
'0  CPU0
'1 CPU1

'3 CPU0 CPU1
'4 CPU2
'5 CPU0 CPU2
'6 CPU1 CPU2
'7 CPU0 CPU1 + CPU2
'8 CPU3
'9 CPU0 CPU3
'0 a CPU1 + CPU3
'0 b CPU0 + CPU1 + CPU3
'0 c CPU2 + CPU3
'0d CPU0+CPU2+CPU3
'0e CPU1+CPU2+CPU3
'0 f CPU0 + CPU1 + CPU2 + CPU3
'''''''''''''''''''''''''''''''''''''''''''
 'Use CPU0
' MyMask = BitMasks(AffinityBit) '«адаем соответствие
 
 'Use CPU1
 'MyMask = BitMasks(1)
 'Use CPU0 and CPU1
 'MyMask = BitMasks(0) Or BitMasks(1)
 
 'The CPUs to use are specified by the array index.
 'To use CPUs 0, 2, and 4, you would use:
 'MyMask = BitMasks(0) Or BitMasks(2) Or BitMasks(4)
 'To Set Affinity, pass the application handle and your custom affinity mask:
 'SetProcessAffinityMask(lngHwndProcess, MyMask)
 'Use GetCurrentProcess() API instead of lngHwndProcess to set affinity on the current app.
 If SetProcessAffinityMask(lngHwndProcess, AffinityBit) = 1 Then
' MsgBox "Affinity Set", vbInformation, "Success"
 Else
 MsgBox "Failed To Set Affinity", vbCritical, "Failure"
 End If
 
 
End Function

Private Function GetBitMasks(ByVal inValue As Long) As Long()
 Dim RetArr() As Long, NumRet As Long
 Dim LoopBits As Long, BitMask As Long
 Const HighBit As Long = &H80000000
 ReDim RetArr(0 To 31) As Long
 For LoopBits = 0 To 30
 BitMask = 2 ^ LoopBits
 If (inValue And BitMask) Then
 RetArr(NumRet) = BitMask
 NumRet = NumRet + 1
 End If
 Next LoopBits
 If (inValue And HighBit) Then
 RetArr(NumRet) = HighBit
 NumRet = NumRet + 1
 End If
 If (NumRet > 0) Then ' Trim unused array items and return array
 If (NumRet < 32) Then ReDim Preserve RetArr(0 To NumRet - 1) As Long
 GetBitMasks = RetArr
 End If
End Function
