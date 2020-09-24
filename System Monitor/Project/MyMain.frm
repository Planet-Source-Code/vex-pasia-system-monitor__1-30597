VERSION 5.00
Begin VB.Form Mymain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Monitor"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MyMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Disk Space"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4455
      Begin SysMon.ProgressCntrl DS 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
      End
   End
   Begin SysMon.ProgressCntrl OptmB 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "Optimize"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame CPU 
      Caption         =   "CPU Usage"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4455
      Begin SysMon.ProgressCntrl CPUB 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   3000
   End
   Begin VB.Frame UM 
      Caption         =   "Physical Memory"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin SysMon.ProgressCntrl UMB 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
      End
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show Me"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Mymain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Dim memoryInfo As MEMORYSTATUS
Dim lastpcent As Single, lastTot As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

'Registry API
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function StrFormatByteSize Lib _
    "shlwapi" Alias "StrFormatByteSizeA" (ByVal _
    dw As Long, ByVal pszBuf As String, ByRef _
    cchBuf As Long) As String
    
'unused
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

'reg constants
Private Const REG_SZ = 1        'a null terminated string
Private Const REG_DWORD = 4     'a double word is 4 bytes (A.K.A. Long variable)
Private Const HKEY_DYN_DATA = &H80000006    'reg key root
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_SET_VALUE = &H2
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&

'for regcreatekey, but unused
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Dim regkey() As Long  'used to keep open a reg key throughout the program
Dim last As Long      'last processor usage %
Dim lastavg As Long   'last avg position
Dim sum As Double     'used to calc the avg
Dim cnt As Double     'used to calc the avg
Dim stime As Single   'start time
Dim WinDir As String  'windows directory
Dim MyIcon As New SysT

Function GetMemoryInfo()
  'get memory info
  DoEvents
  GlobalMemoryStatus memoryInfo
    
  Totp1 = Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10
  Availp1 = Int(memoryInfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10
  pcent = Int(Availp1 / Totp1 * 100)
  
  If pcent < 10 Then cmdOpt_Click
  
  lastpcent = pcent
  lastTot = memoryInfo.dwMemoryLoad
  
  UMB.Caption = "Physical Mem Free: " & Format(Availp1) & " MB of " & Format(Totp1) & " MB (" & Format(lastpcent) & "%)"
  MyIcon.ChangeToolTip Me, UMB.Caption
  
  UMB.Value = Format(lastpcent)
  DS.Max = DiscSpace("C:", 1) 'Get Total Disk Space in Drive C:
  DS.Value = DiscSpace("C:", 0) 'Get Available Disk Space in Drive C:
  DS.Caption = "Free Space in Drive C: " & FormatKB(DiscSpace("C:", 0)) & " out of " & FormatKB(DiscSpace("C:", 1))
  
  use = GetCPUUsage()  'returns a percentage
  sum = sum + use * (Timer1.Interval \ 50)
  cnt = cnt + Timer1.Interval \ 50

  CPUB.Value = Int(sum / cnt)
  CPUB.Caption = "Processor Usage (avg: " & Format(Int(sum / cnt)) & "%) - Time: " & Format(Time - stime, "hh:mm:ss")
End Function

Private Function DiscSpace(ByVal DrvLetter As String, ByVal intOption As Integer)
Dim Status As Long
Dim TotalBytes As Currency
Dim FreeBytes As Currency
Dim BytesAvailableToCaller As Currency
Status = GetDiskFreeSpaceEx(DrvLetter, BytesAvailableToCaller, _
TotalBytes, FreeBytes)
If Status <> 0 Then
    Select Case intOption
        Case 0: DiscSpace = FreeBytes * 10000
        Case 1: DiscSpace = TotalBytes * 10000
    End Select
End If
End Function

Private Function FormatKB(ByVal Amount As Long) _
    As String
    Dim Buffer As String
    Dim Result As String
    Buffer = Space$(255)
    Result = StrFormatByteSize(Amount, Buffer, _
    Len(Buffer))

    If InStr(Result, vbNullChar) > 1 Then
        FormatKB = Left$(Result, InStr(Result, _
            vbNullChar) - 1)
    End If
End Function

Private Sub cmdOpt_Click()
    On Error Resume Next
    ReDim a(20) As String
    
    Dim j As Integer
    OptmB.Max = 20

    For j = 0 To 20
            OptmB.Value = j
            a(j) = Space$(500000)
            DoEvents
            OptmB.Caption = "[" & j / 20 * 100 & "%] Optimizing..."
    Next j
    OptmB.Caption = "Done."
    OptmB.Value = 0
End Sub

Private Sub Form_Load()
Call Always_On_Top(Me.hwnd, Me.Left / Screen.TwipsPerPixelX, _
Me.Top / Screen.TwipsPerPixelY, Me.Height / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, True)

MyIcon.ShowIcon Me

  OptmB.Caption = ""
  CPUB.Max = 100
  UMB.Max = 100
  ReDim regkey(4)     'we use 5 different reg keys
  Call InitializeCPU  'tell windows we want it to monitor the processor
  use = GetCPUUsage() 'get the % usage of the processor
  sum = use * (Timer1.Interval \ 50)
  cnt = Timer1.Interval \ 50
  stime = Time
GetMemoryInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub

Private Sub mnuExit_Click()
DoEvents
CloseCPU
MyIcon.RemoveIcon Me
End
End Sub

Private Sub mnuShow_Click()
Me.Visible = True
Me.SetFocus
End Sub

Private Sub Timer1_Timer()
GetMemoryInfo
End Sub

Public Function GetCPUUsage() As Long
    Dim Data As Long
    Dim hret As Long
    DoEvents
    hret = RegQueryValueEx(regkey(0), "KERNEL\CPUUsage", 0&, REG_DWORD, Data, 4)
    GetCPUUsage = Data
End Function

Public Function CloseCPU() As Long
    Dim Data As Long
    Dim hret As Long
    DoEvents
    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StopStat", regkey(4))
    hret = RegQueryValueEx(regkey(4), "KERNEL\CPUUsage", 0&, REG_DWORD, Data, 4)
    hret = RegCloseKey(regkey(4))
    hret = RegCloseKey(regkey(0))
End Function

'Initialize the CPU meter stats
Public Function InitializeCPU() As Long
    Dim Data As Long
    Dim hret As Long
    DoEvents
    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", regkey(0))
    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartSrv", regkey(1))
    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StopSrv", regkey(2))
    hret = RegQueryValueEx(regkey(1), "KERNEL", 0&, REG_DWORD, Data, 4)
    hret = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", regkey(3))
    hret = RegQueryValueEx(regkey(3), "KERNEL\CPUUsage", 0&, REG_DWORD, Data, 4)
    hret = RegCloseKey(regkey(3))
    hret = RegQueryValueEx(regkey(2), "KERNEL", 0&, REG_DWORD, Data, 4)
    hret = RegCloseKey(regkey(1))
    hret = RegCloseKey(regkey(2))
    
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim Result As Long
Dim msg As Long
    
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    
    Select Case msg
        Case 517
            Me.PopupMenu mnu
        Case 515
            mnuShow_Click
    End Select
End Sub
