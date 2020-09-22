Attribute VB_Name = "modMouseScroll"
'Added:
'Use: For use with mouse wheel

Option Explicit

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Const WM_MOUSEWHEEL = &H20A
Public Const GWL_WNDPROC = (-4)

Public IsHooked As Boolean
Dim PrevProc As Long
Dim m_hWnd As Long

Public Sub Hook(hwnd As Long)
    On Error Resume Next
    PrevProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
    IsHooked = True
End Sub

Public Sub UnHook()
    If IsHooked Then
        SetWindowLong m_hWnd, GWL_WNDPROC, PrevProc
        IsHooked = False
    End If
End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    
    If uMsg = WM_MOUSEWHEEL Then
        Dim ctrl As Control
        Dim strVar As String
        If (wParam \ 65536) < 0 Then
            For Each ctrl In Screen.ActiveForm.Controls
                If TypeOf ctrl Is ISPanel Then
                    If ctrl.ScrollDown Then Exit For
                Else
                    'sometimes loses if the control is an IsPanel or not setting
                    'modifying this variable prevents that for some unknown reason
                    strVar = ctrl.Name
                    DoEvents
                End If
            Next
        Else
            For Each ctrl In Screen.ActiveForm.Controls
                If TypeOf ctrl Is ISPanel Then
                    If ctrl.ScrollUp Then Exit For
                Else
                    'sometimes loses if the control is an IsPanel or not setting
                    'modifying this variable prevents that for some unknown reason
                    strVar = ctrl.Name
                    DoEvents
                End If
            Next
        End If
    End If
    m_hWnd = hwnd
End Function

