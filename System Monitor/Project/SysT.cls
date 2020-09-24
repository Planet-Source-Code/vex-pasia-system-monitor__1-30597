VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysT"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
  
Private FormHandle As Long
Private mvarbRunningInTray As Boolean
Private SysIcon As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Property Let bRunningInTray(ByVal vData As Boolean)
    mvarbRunningInTray = vData
End Property


Property Get bRunningInTray() As Boolean
    bRunningInTray = mvarbRunningInTray
End Property

Public Sub ShowIcon(ByRef sysTrayForm As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = 512
    SysIcon.hIcon = sysTrayForm.Icon
    SysIcon.szTip = sysTrayForm.Caption + Chr(0)
        
    Shell_NotifyIcon 0, SysIcon
    mvarbRunningInTray = True
End Sub

Public Sub RemoveIcon(sysTrayForm As Form)
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    SysIcon.uId = vbNull
    SysIcon.uFlags = 7
    SysIcon.ucallbackMessage = vbNull
    SysIcon.hIcon = sysTrayForm.Icon
    SysIcon.szTip = Chr(0)
    Shell_NotifyIcon 2, SysIcon
    'If sysTrayForm.Visible = False Then sysTrayForm.Show    'Incase user can't see form
    mvarbRunningInTray = False
End Sub

Public Sub ChangeIcon(sysTrayForm As Form, picNewIcon As Image)

If mvarbRunningInTray = True Then   'If running in the tray
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    'SysIcon.uId = vbNull
    'SysIcon.uFlags = 7
    'SysIcon.ucallbackMessage = 512
    SysIcon.hIcon = picNewIcon.Picture
    'SysIcon.szTip = sysTrayForm.Caption + Chr(0)
    Shell_NotifyIcon 1, SysIcon
End If
End Sub

Public Sub ChangeToolTip(sysTrayForm As Form, strNewTip As String)

If mvarbRunningInTray = True Then   'If running in the tray
    SysIcon.cbSize = Len(SysIcon)
    SysIcon.hwnd = sysTrayForm.hwnd
    SysIcon.szTip = strNewTip & Chr(0)
    Shell_NotifyIcon 1, SysIcon
End If

End Sub
