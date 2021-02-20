VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "process_spawner"
   ClientHeight    =   3756
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "process_spawner"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3756
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Wait = False
BitMaskConf = 0
Dim sArgs() As String
Dim StartString As String
    
    Dim iLoop As Integer
    sArgs = Split(Command$, " ")
    For iLoop = 0 To UBound(sArgs)
        If readParam(sArgs(iLoop)) <> True Then
       ' MsgBox sArgs(iLoop)
        StartString = StartString & sArgs(iLoop) & " "
        End If
    Next
'Shell StartString
If Len(StartString) < 1 Then
MsgBox "Start string is empty!"
End
End If

Open App.Path & "\process_spawner.log" For Output As #1
Print #1, StartString
Close #1


If Wait = True Then
    Dim t As String
    t = GetTickCount
    ExecuteAndWait StartString
    If CloseOnFast = True And GetTickCount - t < 2000 Then End
Else
    Dim hPid As Long
    hPid = Shell(StartString)
    Call SetAffinity(hPid, BitMaskConf)
End If

If Len(StartFileOnExit) > 0 Then Call Shell(StartFileOnExit, vbNormalFocus)
End
End Sub



