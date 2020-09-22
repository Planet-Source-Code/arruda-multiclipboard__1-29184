Attribute VB_Name = "Module1"
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function Shell_NotifyIcon Lib "Shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallBackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200

Public IconTray As NOTIFYICONDATA
Dim ShiftV As Boolean, ShiftC As Boolean, NumStatus As Boolean
Dim VarText(0 To 3) As Variant

Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal dwTime As Long)
    
    On Error Resume Next
    If GetAsyncKeyState(vbKeyEscape) < 0 Then
        If Form1.Mn.Visible Then Form1.Mn.Visible = False
    End If
    
    V1 = GetAsyncKeyState(vbKeyControl)
    V2 = GetAsyncKeyState(vbKeyC)
    ShiftC = (V1 < 0) And (V2 < 0)
    
    T1 = IIf((GetAsyncKeyState(vbKey1) Or GetAsyncKeyState(vbKeyNumpad1)) < 0, 1, 0)
    T2 = IIf((GetAsyncKeyState(vbKey2) Or GetAsyncKeyState(vbKeyNumpad2)) < 0, 1, 0)
    T3 = IIf((GetAsyncKeyState(vbKey3) Or GetAsyncKeyState(vbKeyNumpad3)) < 0, 1, 0)
    T4 = IIf((GetAsyncKeyState(vbKey4) Or GetAsyncKeyState(vbKeyNumpad4)) < 0, 1, 0)
    
    NumStatus = False
    If T1 > 0 Then
        NumStatus = True
    Else
        If T2 > 0 Then
            NumStatus = True
        Else
            If T3 > 0 Then
                NumStatus = True
            Else
                If T4 > 0 Then NumStatus = True
            End If
        End If
    End If
    ShiftV = (V1 < 0) And (NumStatus)

    If ShiftC Then
        If T1 > 0 Then VarText(0) = Clipboard.GetText()
        If T2 > 0 Then VarText(1) = Clipboard.GetText()
        If T3 > 0 Then VarText(2) = Clipboard.GetText()
        If T4 > 0 Then VarText(3) = Clipboard.GetText()
    End If
    
    If ShiftV Then
        Clipboard.Clear
        If T1 > 0 Then Clipboard.SetText VarText(0)
        If T2 > 0 Then Clipboard.SetText VarText(1)
        If T3 > 0 Then Clipboard.SetText VarText(2)
        If T4 > 0 Then Clipboard.SetText VarText(3)
    End If
    
    
End Sub
