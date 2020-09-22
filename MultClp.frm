VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   450
   ClientLeft      =   1500
   ClientTop       =   1530
   ClientWidth     =   1560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "MultClp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   104
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Menu Mn 
      Caption         =   "Multi Clipboard"
      Visible         =   0   'False
      Begin VB.Menu Sm1 
         Caption         =   "Close Multi-Clipboard"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
        
    IconTray.cbSize = Len(IconTray)
    IconTray.hwnd = hwnd
    IconTray.hIcon = Me.Icon
    IconTray.uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
    IconTray.uCallBackMessage = WM_MOUSEMOVE
    IconTray.szTip = "Multi-Clipboard" + Chr(0)
    IconTray.uID = 1&
    
    Shell_NotifyIcon NIM_ADD, IconTray
    SetTimer hwnd, 0, 1, AddressOf TimerProc
    RegisterServiceProcess GetCurrentProcessId(), 1


End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    If X = WM_LBUTTONUP Or X = WM_RBUTTONUP Then PopupMenu Mn, 0
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    X = Shell_NotifyIcon(NIM_DELETE, IconTray)
    KillTimer Me.hwnd, 0
    End
    
End Sub

Private Sub Sm1_Click()

    Form_Unload False

End Sub


