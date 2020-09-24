Attribute VB_Name = "AlwOnTop"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd _
As Long, ByVal hWndInsertAfter_ As Long, ByVal X _
As Long, ByVal y_ As Long, ByVal cx As Long, ByVal cy _
As Long, ByVal wFlags_ As Long) As Long

Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40

Public Function Always_On_Top(ByVal H, FrmX As Long, _
FrmY As Long, Hght As Long, Wdth As Long, YesAOT As Boolean)
If YesAOT = True Then
SetWindowPos H, conHwndTopmost, FrmX, FrmY, Wdth, Hght, _
conSwpNoActivate
ElseIf YesAOT = False Then
SetWindowPos H, conHwndNoTopmost, FrmX, FrmY, Wdth, Hght, _
conSwpShowWindow
End If
End Function

