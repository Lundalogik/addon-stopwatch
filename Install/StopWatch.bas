Attribute VB_Name = "StopWatch"
Option Explicit

Dim oDictionary As New Dictionary
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const PRINT_DEBUG As Boolean = True
Private Const PRINT_LOG As Boolean = True

Public Sub Start(sKey As String, Optional bReset As Boolean = True)
    On Error GoTo ErrorHandler
    
    If (oDictionary Is Nothing) Then
        Set oDictionary = New Dictionary
    End If
        
    If (oDictionary.Exists(sKey) = True And bReset = True) Then
        oDictionary(sKey) = GetTickCount
        Call DebugPrint("StopWatch reset """ + sKey + """")
    ElseIf (oDictionary.Exists(sKey) = False) Then
        Call oDictionary.Add(sKey, GetTickCount)
        Call DebugPrint("StopWatch started """ + sKey + """")
    End If
    
    Exit Sub
ErrorHandler:
    UI.ShowError ("StopWatch.Start")
End Sub

Public Sub Remove(sKey As String)
    On Error GoTo ErrorHandler
    
    If (oDictionary Is Nothing) Then
        Set oDictionary = New Dictionary
    End If
        
    If (oDictionary.Exists(sKey) = True) Then
        Call oDictionary.Remove(sKey)
        Call DebugPrint("StopWatch """ + sKey + """ removed...")
    Else
        DebugPrint ("No StopWatch with key """ + sKey)
    End If
    
    Exit Sub
ErrorHandler:
    UI.ShowError ("StopWatch.Remove")
End Sub

Public Function PrintTime(sKey As String) As Long
    On Error GoTo ErrorHandler
    
    If (oDictionary Is Nothing) Then
        Set oDictionary = New Dictionary
    End If
        
    If (oDictionary.Exists(sKey) = True) Then
        Dim lTime As Long
        lTime = (GetTickCount - oDictionary.Item(sKey))
        
        Call DebugPrint("StopWatch """ + sKey + """ TIME: " + VBA.CStr(lTime))
        Call oDictionary.Remove(sKey)
    Else
        Call DebugPrint("No StopWatch with key """ + sKey)
    End If
    
    Exit Function
ErrorHandler:
    UI.ShowError ("StopWatch.PrintTime")
End Function

Private Sub DebugPrint(sText As String)
On Error GoTo ErrorHandler
    If PRINT_DEBUG Then
        Debug.Print sText
    End If
    If PRINT_LOG Then
        Call Application.log.Add(lkEventLogTypeDebug, "StopWatch", "", sText)
    End If

Exit Sub
ErrorHandler:
    Call UI.ShowError("StopWatch.DebugPrint")
End Sub
