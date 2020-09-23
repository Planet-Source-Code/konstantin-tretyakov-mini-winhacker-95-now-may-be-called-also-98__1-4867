Attribute VB_Name = "modAPI"
'Project WinHack
'Copyright Tretyakov Konstantin (kt_ee@yahoo.com)
'You may use this code for free, if you give me some credit
'At least remember, thet it is not fair to put your name on what you didn't do

'And I would surely appreciate, if you mail me the program (or link to it)
'you created, using this code, (or if you somehow modified this one)

'API function declarations for WinHack

Option Explicit
Public m_hWnd As Long
Public MeOnTop As Boolean
Public Const ColChange = 5000

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
'Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
'Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long 'nCmdShow: 0 = hide, 1 = restore, 2 = minimize, 3 = maximize
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
'Declare Function ShowOwnedPopups Lib "user32" (ByVal hwnd As Long, ByVal fShow As Long) As Long
'works like SendMessage
'PostMessage hwnd, WM_CLOSE, 0, 0&
'This closes a program ,cool ?
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Const WM_CLOSE = &H10

'Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
'Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
'Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long

Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function AnyPopup Lib "user32" () As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Declare Function EnumWindows Lib "user32" (ByVal lpfn As Long, lParam As Any) As Boolean
'Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'This is used to set/deset "Always on top"
'Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

' ShowWindow() Commands
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_MAX = 10

Private Function CallBackFunc(ByVal hwnd As Long, lParam As Control) As Long
    Dim strhWnd$, strClass$, TempCurY&
    If (IsWindowVisible(hwnd) And (frmWinHack.chkVisible.Value = 1)) Or frmWinHack.chkVisible.Value = 0 Then
        strhWnd = "&H" & Hex(hwnd)
        strClass = GetClass(hwnd)
        With frmWinHack
        Load .Label1(.Label1.Count)
        .Label1(.Label1.Count - 1).Caption = GetWindowCaption(hwnd) & " [" & strClass & "]"
        .Label1(.Label1.Count - 1).Tag = hwnd
        .Label1(.Label1.Count - 1).Top = .Label1(.Label1.Count - 2).Top + .Label1(.Label1.Count - 2).Height
        .Label1(.Label1.Count - 1).Visible = True
        .Label1(.Label1.Count - 1).BackColor = IIf(IsWindowVisible(hwnd) <> 0, .Label1(0).BackColor, .Label1(0).BackColor + ColChange)
        End With
        'TempCurY = lParam.CurrentY
        'lParam.Print strhWnd, strClass ', GetWindowCaption(hWnd)
        'lParam.CurrentY = TempCurY
        'lParam.CurrentX = 300
        'lParam.Print GetWindowCaption(hwnd)
    End If
    CallBackFunc = True
End Function
Public Function GetClass$(hwnd As Long)
    Dim lRetChars&, sClassName$
    sClassName = Space(256)
    lRetChars = GetClassName(hwnd, sClassName, Len(sClassName))
    GetClass = Left(sClassName, lRetChars)
'    Debug.Print "GetClass hWnd = " & hWnd
'    Debug.Print "Class = " & Left(sClassName, lRetChars)
End Function
Public Function GetWindowCaption$(ByVal hwnd As Long)
    Dim sRetCapt$, RetLen As Long
    RetLen = GetWindowTextLength(hwnd) + 1
    sRetCapt = Space(RetLen)
    RetLen = GetWindowText(hwnd, sRetCapt, RetLen)
    GetWindowCaption = Left(sRetCapt, RetLen)
End Function
Public Sub CallBackDemo(frmName As Control)
'    frmName.Cls
'    frmName.Print "Handle", "Class Name", Space(45), "     Window Caption"
'    frmName.Print "-----------", "-------------------", Space(45), "     --------------------------"
    EnumWindows AddressOf CallBackFunc, frmName
End Sub


Public Function SetTopWindow(hwnd As Long, bState As Boolean) As Boolean
  If bState = True Then 'Put the window on top
    SetTopWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  ElseIf bState = False Then ' Turn off the TopMost flag
    SetTopWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
  Else
    Debug.Print "bState Unknown."
    SetTopWindow = False
  End If
End Function

Function GetCoords(ByVal hwnd As Long, Optional Left As Long, Optional Top As Long, Optional Right As Long, Optional Bottom As Long) As Boolean
    Dim TempRect As RECT
    GetCoords = GetWindowRect(hwnd, TempRect)
    With TempRect
        Left = .Left
        Top = .Top
        Right = .Right
        Bottom = .Bottom
    End With
End Function

