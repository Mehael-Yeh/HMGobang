Attribute VB_Name = "Module2"
Public Const GWL_WNDPROC = (-4)
Public Const WM_MOUSEWHEEL = &H20A
Public OldProcAddr As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Function MyWinProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
If msg <> WM_MOUSEWHEEL Then
    MyWinProc = CallWindowProc(OldProcAddr, hwnd, msg, wp, lp)
Else
    Debug.Print msg
    Dim s As String
    s = Hex(wp)
    If Len(s) < 8 Then s = String(8 - Len(s), "0") & s
    Dim zDelta As Long
    zDelta = CInt("&h" & Left(s, 4)) * (-1)
    If Form2.VScroll1.Value + zDelta < Form2.VScroll1.Min Then
        Form2.VScroll1.Value = Form2.VScroll1.Min
    ElseIf Form2.VScroll1.Value + zDelta > Form2.VScroll1.Max Then
        Form2.VScroll1.Value = Form2.VScroll1.Max
    Else
        Form2.VScroll1.Value = Form2.VScroll1.Value + zDelta
    End If
End If
End Function

