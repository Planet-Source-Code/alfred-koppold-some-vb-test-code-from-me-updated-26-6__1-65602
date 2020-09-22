Attribute VB_Name = "modSubClass"
Option Explicit

Private colScrollEvents As New Collection


'for subclassing ...
Private oldWinProc As Long
Private Const GWL_WNDPROC = (-4)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_VSCROLL = &H115

Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_LEFT = 6
Private Const SB_RIGHT = 7
Private Const SB_ENDSCROLL = 8

Public Function WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim objScrollEvents As CScrollEvents
    Dim lPos As Long
    Dim lScroll As Long
    
    'Set objScrollEvents = colScrollEvents.Item(CStr(hwnd))
    For Each objScrollEvents In colScrollEvents
        If objScrollEvents.hwnd = hwnd Then Exit For
    Next objScrollEvents
    
    If Not objScrollEvents Is Nothing Then
        Select Case wMsg
            Case WM_VSCROLL
                objScrollEvents.TriggerCount colScrollEvents.count
                lScroll = (wParam And &HFFFF&)
                lPos = (wParam And &H7FFFFFFF) / &HFFFF&
                Select Case lScroll
                    Case SB_LINEUP
                        objScrollEvents.TriggerScrollLine -1
                    Case SB_LINEDOWN
                        objScrollEvents.TriggerScrollLine 1
                    Case SB_PAGEUP
                        objScrollEvents.TriggerScrollPages -1
                    Case SB_PAGEDOWN
                        objScrollEvents.TriggerScrollPages 1
                    Case SB_THUMBPOSITION
                        objScrollEvents.TriggerScrollPos lPos
                    Case SB_THUMBTRACK
                        objScrollEvents.TriggerScrollTrack lPos
                    Case SB_ENDSCROLL
                        objScrollEvents.TriggerScroll
                    Case Else
                        Debug.Print wParam
                End Select
                WndProc = 0
                Exit Function
        End Select
    End If
    
    'Pass on messages
    WndProc = CallWindowProc(oldWinProc, hwnd, wMsg, wParam, lParam)
    
End Function

Public Function HookWindow(hwnd As Long) As CScrollEvents
    Dim objScrollEvents As CScrollEvents
    
    'store the old message handler.
    oldWinProc = GetWindowLong(hwnd, GWL_WNDPROC)
    'set the message handler to ours.
    SetWindowLong hwnd, GWL_WNDPROC, AddressOf WndProc
    
    Set objScrollEvents = New CScrollEvents
    objScrollEvents.hwnd = hwnd
    colScrollEvents.Add objScrollEvents, CStr(hwnd)
    Set HookWindow = objScrollEvents
End Function

Public Sub UnHookWindow(hwnd As Long)
    Dim objScrollEvents As CScrollEvents
    
    'Sets procedure for handling events back
    '     to the original.
    Set objScrollEvents = colScrollEvents.Item(CStr(hwnd))
    colScrollEvents.Remove CStr(hwnd)
    Set objScrollEvents = Nothing
    
    SetWindowLong hwnd, GWL_WNDPROC, oldWinProc
End Sub


