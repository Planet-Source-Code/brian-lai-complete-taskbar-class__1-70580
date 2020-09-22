VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetLayeredWindowAttributes Lib "user32" (ByVal Hwnd As Long, ByRef pcrKey As Long, ByRef pbAlpha As Byte, ByRef pdwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal Hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As TBarData) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal Hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
 
Private Type TBarData
    cbSize As Long
    Hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type

Dim K As TBarData

'Class Functions


Property Let Clock(What As Boolean)
    On Error Resume Next
    ShowWindow ClockHwnd, IIf(What, 1, 0)
End Property

Property Get Clock() As Boolean
    On Error Resume Next
    Clock = IsWindowVisible(ClockHwnd)
End Property

Property Get Hwnd() As Long
    Hwnd = TaskbarHwnd 'wrapper
End Property

Property Let Left(What As Long)
    'MOVEWINDOW does not work on the taskbar
    'MoveWindow TaskbarHwnd, What, Me.Top, Me.Width, Me.Height, 1
End Property

Property Get Left() As Long
    On Error Resume Next
    K = GetTBarData
    Left = K.rc.Left
End Property

Property Let Top(What As Long)
    'MOVEWINDOW does not work on the taskbar
    'MoveWindow TaskbarHwnd, Me.Left, What, Me.Width, Me.Height, 1
End Property

Property Get Top() As Long
    On Error Resume Next
    K = GetTBarData
    Top = K.rc.Top
End Property

Property Let Width(What As Long)
    'MOVEWINDOW does not work on the taskbar
    'MoveWindow TaskbarHwnd, Me.Left, Me.Top, What, Me.Height, 1
End Property

Property Get Width() As Long
    On Error Resume Next
    K = GetTBarData
    Width = K.rc.Right - K.rc.Left
End Property

Property Let Height(What As Long)
    'MOVEWINDOW does not work on the taskbar
    'MoveWindow TaskbarHwnd, Me.Left, Me.Top, Me.Width, What, 1
End Property

Property Get Height() As Long
    On Error Resume Next
    K = GetTBarData
    Height = K.rc.Bottom - K.rc.Top
End Property

Property Let Programs(What As Boolean)
    On Error Resume Next
    ShowWindow ProgramsHwnd, IIf(What, 1, 0)
End Property

Property Get Programs() As Boolean
    On Error Resume Next
    Programs = IsWindowVisible(ProgramsHwnd)
End Property

Property Let StartButton(What As Boolean)
    On Error Resume Next
    ShowWindow StartButtonHwnd, IIf(What, 1, 0)
End Property

Property Get StartButton() As Boolean
    On Error Resume Next
    StartButton = IsWindowVisible(StartButtonHwnd)
End Property

Property Let Systray(What As Boolean)
    On Error Resume Next
    ShowWindow SystrayHwnd, IIf(What, 1, 0)
End Property

Property Get Systray() As Boolean
    On Error Resume Next
    Systray = IsWindowVisible(SystrayHwnd)
End Property

Property Let Transparency(Perc As Integer)
    On Error Resume Next
    MakeTransparent TaskbarHwnd, Perc
End Property

Property Get Transparency() As Integer
    On Error Resume Next
    Transparency = HowTransparent(TaskbarHwnd)
End Property

Property Let Visible(What As Boolean)
    On Error Resume Next
    ShowWindow TaskbarHwnd, IIf(What, 1, 0)
End Property

Property Get Visible() As Boolean
    On Error Resume Next
    Visible = IsWindowVisible(TaskbarHwnd)
End Property
















'Modular Functions
 
Private Function GetTBarData() As TBarData
    On Error Resume Next
    Dim TBar As TBarData
    SHAppBarMessage &H5, TBar
    GetTBarData = TBar
    
End Function

Private Function TaskbarHwnd() As Long
    On Error Resume Next
    TaskbarHwnd = FindWindow("Shell_Traywnd", vbNullString)
End Function

Public Function ProgramsHwnd() As Long
    On Error Resume Next
    ProgramsHwnd = FindWindowEx(TaskbarHwnd, 0, "ReBarWindow32", vbNullString)
    ProgramsHwnd = FindWindowEx(ProgramsHwnd, 0, "MSTaskSwWClass", vbNullString)
    ProgramsHwnd = FindWindowEx(ProgramsHwnd, 0, "ToolbarWindow32", vbNullString)
End Function

Public Function StartButtonHwnd() As Long
    On Error Resume Next
    StartButtonHwnd = FindWindowEx(TaskbarHwnd, 0, "button", vbNullString)
End Function

Public Function SystrayHwnd() As Long
    On Error Resume Next
    SystrayHwnd = FindWindowEx(TaskbarHwnd, 0, "traynotifywnd", vbNullString)
End Function

Public Function ClockHwnd() As Long
    On Error Resume Next
    ClockHwnd = FindWindowEx(SystrayHwnd, 0, "TrayClockWClass", vbNullString)
End Function

Private Function isTransparent(ByVal Hwnd As Long) As Boolean
    On Error Resume Next
    Dim Msg As Long
    Msg = GetWindowLong(Hwnd, GWL_EXSTYLE)
    If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
      isTransparent = True
    Else
      isTransparent = False
    End If
    If Err Then
      isTransparent = False
    End If
End Function

Private Function HowTransparent(ByVal Hwnd As Long) As Long
    On Error Resume Next
    Dim L As Long
    Dim B As Byte
    
    If isTransparent(Hwnd) = True Then
        If GetLayeredWindowAttributes(Hwnd, L, B, 0) Then
            L = CLng(CLng(B) / 2.55)
            HowTransparent = L
        Else
            HowTransparent = 100
        End If
    Else
        HowTransparent = 100
    End If
End Function

Private Function MakeTransparent(ByVal Hwnd As Long, Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 100 Then
        MakeTransparent = 50 'just to make everyone feel happy when they input shit values
    ElseIf Perc = 100 Then
        MakeOpaque Hwnd
        MakeTransparent = 0
    Else
        'Perc = Perc  'percentage...
      Msg = GetWindowLong(Hwnd, GWL_EXSTYLE)
      Msg = Msg Or WS_EX_LAYERED
      SetWindowLong Hwnd, GWL_EXSTYLE, Msg
      SetLayeredWindowAttributes Hwnd, 0, Perc * 2.55, LWA_ALPHA
      MakeTransparent = 0
    End If
    If Err Then
      MakeTransparent = 2
    End If
End Function

Private Function MakeOpaque(ByVal Hwnd As Long) As Long
    On Error Resume Next
    Dim Msg As Long
    Msg = GetWindowLong(Hwnd, GWL_EXSTYLE)
    Msg = Msg And Not WS_EX_LAYERED
    SetWindowLong Hwnd, GWL_EXSTYLE, Msg
    SetLayeredWindowAttributes Hwnd, 0, 0, LWA_ALPHA
    MakeOpaque = 0
    If Err Then MakeOpaque = 2
End Function
