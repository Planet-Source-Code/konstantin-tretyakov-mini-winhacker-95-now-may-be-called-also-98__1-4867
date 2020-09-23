VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmWinHack 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mini WinHacker 95"
   ClientHeight    =   4320
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5865
   Icon            =   "frmWinHack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   300
      Left            =   4620
      TabIndex        =   26
      Top             =   3960
      Width           =   1160
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3360
      TabIndex        =   25
      Top             =   3960
      Width           =   1160
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2100
      TabIndex        =   24
      Top             =   3960
      Width           =   1160
   End
   Begin VB.Frame Frame 
      Caption         =   "About"
      Height          =   3252
      Index           =   3
      Left            =   180
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   5532
      Begin VB.Image imgBeauty 
         Height          =   480
         Index           =   3
         Left            =   120
         Picture         =   "frmWinHack.frx":030A
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblAbout 
         Caption         =   "About"
         Height          =   2832
         Left            =   660
         TabIndex        =   23
         Top             =   240
         Width           =   4752
      End
   End
   Begin VB.Frame Frame 
      BorderStyle     =   0  'None
      Height          =   3204
      Index           =   1
      Left            =   192
      TabIndex        =   33
      Top             =   480
      Visible         =   0   'False
      Width           =   5484
      Begin VB.Frame Frame3 
         Caption         =   "System"
         Height          =   1800
         Left            =   0
         TabIndex        =   35
         Top             =   1080
         Width           =   5484
         Begin VB.TextBox txtDupName 
            Height          =   288
            Left            =   2520
            TabIndex        =   22
            Top             =   1320
            Width           =   2772
         End
         Begin VB.TextBox txtPrinters 
            Height          =   288
            Left            =   2520
            TabIndex        =   20
            Top             =   960
            Width           =   2772
         End
         Begin VB.TextBox txtRecycle 
            Height          =   288
            Left            =   2520
            TabIndex        =   16
            Top             =   240
            Width           =   2772
         End
         Begin VB.TextBox txtControlPanel 
            Height          =   288
            Left            =   2520
            TabIndex        =   18
            Top             =   600
            Width           =   2772
         End
         Begin VB.Label Label3 
            Caption         =   "&Dial-Up Networking Folder name:"
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   2412
         End
         Begin VB.Label Label3 
            Caption         =   "Printers &Folder name:"
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   1932
         End
         Begin VB.Label Label3 
            Caption         =   "&Recycle Bin name:"
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1572
         End
         Begin VB.Label Label3 
            Caption         =   "Control &Panel name:"
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1932
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "User information"
         Height          =   1020
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   5484
         Begin VB.TextBox txtOrg 
            Height          =   288
            Left            =   2520
            TabIndex        =   14
            Top             =   600
            Width           =   2772
         End
         Begin VB.TextBox txtOwner 
            Height          =   288
            Left            =   2520
            TabIndex        =   12
            Top             =   240
            Width           =   2772
         End
         Begin VB.Label Label3 
            Caption         =   "Registered Or&ganization:"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   1932
         End
         Begin VB.Label Label3 
            Caption         =   "Registered O&wner:"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1572
         End
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Window List"
      Height          =   3252
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   480
      Visible         =   0   'False
      Width           =   5592
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   60
         Top             =   780
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   2052
         Left            =   60
         TabIndex        =   38
         Top             =   660
         Width           =   5232
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   612
            Left            =   0
            ScaleHeight     =   615
            ScaleWidth      =   5235
            TabIndex        =   39
            Top             =   0
            Width           =   5232
            Begin VB.Label Label1 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Window Caption and Class"
               Height          =   252
               Index           =   0
               Left            =   0
               TabIndex        =   40
               Top             =   -240
               UseMnemonic     =   0   'False
               Width           =   5232
            End
         End
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Show Visible &Windows only"
         Height          =   192
         Left            =   120
         TabIndex        =   27
         Top             =   220
         Value           =   1  'Checked
         Width           =   2832
      End
      Begin VB.VScrollBar v1 
         Height          =   2472
         LargeChange     =   10
         Left            =   5340
         TabIndex        =   28
         Top             =   360
         Width           =   192
      End
      Begin VB.Label lblWMngLabel 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Window Caption and Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   60
         TabIndex        =   42
         Top             =   420
         UseMnemonic     =   0   'False
         Width           =   5232
      End
      Begin VB.Label Label4 
         Caption         =   "To access menu click on a window name with your right mouse button,  doubleclick to see window information."
         Height          =   372
         Left            =   180
         TabIndex        =   41
         Top             =   2760
         Width           =   4992
      End
   End
   Begin ComctlLib.TabStrip Tab1 
      Height          =   3804
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5736
      _ExtentX        =   10107
      _ExtentY        =   6720
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Shell"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Set shell options"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Captions"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Change some captions and user information"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Wind&ow Manager"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "A tool to hack your windows"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Abo&ut"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "About Mini WinHack"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame 
      Caption         =   "General"
      Height          =   3276
      Index           =   0
      Left            =   192
      TabIndex        =   29
      Top             =   480
      Width           =   5484
      Begin VB.CheckBox chkDUpStart 
         Caption         =   "&Dial-Up Networking Folder contents on Start Menu"
         Height          =   192
         Left            =   672
         TabIndex        =   6
         Top             =   2640
         Width           =   4752
      End
      Begin VB.CheckBox chkPRNStart 
         Caption         =   "Printers &Folder contents on Start Menu"
         Height          =   192
         Left            =   672
         TabIndex        =   5
         Top             =   2400
         Width           =   4752
      End
      Begin VB.CheckBox chkCPanelStart 
         Caption         =   "Control &Panel Folder contents on Start Menu"
         Height          =   192
         Left            =   672
         TabIndex        =   4
         Top             =   2160
         Width           =   3552
      End
      Begin VB.CheckBox chkBMPIcon 
         Caption         =   "Show the icon for &BMP files as a mini picture of itself"
         Height          =   204
         Left            =   672
         TabIndex        =   3
         Top             =   1728
         Width           =   4620
      End
      Begin VB.CheckBox chkIconWrap 
         Caption         =   "Wrap &icon title"
         Height          =   204
         Left            =   672
         TabIndex        =   2
         Top             =   1056
         Width           =   1644
      End
      Begin VB.CheckBox chkWinAni 
         Caption         =   "&Window animation"
         Height          =   204
         Left            =   672
         TabIndex        =   1
         Top             =   384
         Width           =   1644
      End
      Begin VB.Frame Frame2 
         Caption         =   "Menu delay"
         Height          =   1356
         Left            =   2400
         TabIndex        =   30
         Top             =   192
         Width           =   2988
         Begin ComCtl2.UpDown udnMenuDelay 
            Height          =   288
            Left            =   1585
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   480
            Width           =   156
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   327681
            Value           =   400
            BuddyControl    =   "txtMenuDelay"
            BuddyDispid     =   196628
            OrigLeft        =   768
            OrigTop         =   384
            OrigRight       =   924
            OrigBottom      =   684
            Increment       =   10
            Max             =   1000
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtMenuDelay 
            Height          =   288
            Left            =   1056
            MaxLength       =   5
            TabIndex        =   9
            Text            =   "0"
            Top             =   480
            Width           =   528
         End
         Begin VB.Label Label2 
            Caption         =   "&Menu show delay:"
            Height          =   204
            Index           =   1
            Left            =   192
            TabIndex        =   8
            Top             =   288
            Width           =   1836
         End
         Begin VB.Label lblMenuNote 
            Caption         =   "0 = Fastest / 1000 = Slowest"
            Height          =   396
            Left            =   192
            TabIndex        =   32
            Top             =   864
            Width           =   2604
         End
         Begin VB.Label Label2 
            Caption         =   "milliseconds"
            Height          =   204
            Index           =   0
            Left            =   1824
            TabIndex        =   31
            Top             =   576
            Width           =   1068
         End
      End
      Begin VB.CheckBox chkRBinStart 
         Caption         =   "&Recycle Bin Folder contents on Start Menu"
         Height          =   192
         Left            =   672
         TabIndex        =   7
         Top             =   2880
         Width           =   4752
      End
      Begin VB.Image imgBeauty 
         Height          =   480
         Index           =   4
         Left            =   180
         Picture         =   "frmWinHack.frx":074C
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image imgBeauty 
         Height          =   480
         Index           =   2
         Left            =   195
         Picture         =   "frmWinHack.frx":0A56
         Top             =   1635
         Width           =   480
      End
      Begin VB.Image imgBeauty 
         Height          =   480
         Index           =   1
         Left            =   195
         Picture         =   "frmWinHack.frx":0D60
         Top             =   960
         Width           =   480
      End
      Begin VB.Image imgBeauty 
         Height          =   480
         Index           =   0
         Left            =   195
         Picture         =   "frmWinHack.frx":106A
         Top             =   285
         Width           =   480
      End
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   -96
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmWinHack.frx":1374
            Key             =   "sinewave"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuLabel 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuShowWindow 
         Caption         =   "Show Window"
      End
      Begin VB.Menu mnuShowNoActivate 
         Caption         =   "Show Window (do not activate)"
      End
      Begin VB.Menu mnuHideWindow 
         Caption         =   "Hide Window"
      End
      Begin VB.Menu mnuPostMessage 
         Caption         =   "Destroy Window"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaximizeWindow 
         Caption         =   "Maximize Window"
      End
      Begin VB.Menu mnuRestoreWindow 
         Caption         =   "Restore Window"
      End
      Begin VB.Menu mnuMinimizeWindow 
         Caption         =   "Minimize Window"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBringtoTop 
         Caption         =   "Bring Window to Top"
      End
      Begin VB.Menu mnuAlOnTopStart 
         Caption         =   "Always on Top"
         Begin VB.Menu mnuAlwaysOnTopTrue 
            Caption         =   "Yes"
         End
         Begin VB.Menu mnuAlwaysOnTopFalse 
            Caption         =   "No"
         End
      End
      Begin VB.Menu mnuFlWinStart 
         Caption         =   "Flash Window"
         Begin VB.Menu mnuFlashWindowTrue 
            Caption         =   "Yes"
         End
         Begin VB.Menu mnuFlashWindowFalse 
            Caption         =   "No"
         End
      End
      Begin VB.Menu mnuSetWindowText 
         Caption         =   "Set New Window Caption..."
      End
      Begin VB.Menu mnuMoveWindow 
         Caption         =   "Move Window..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeeHandle 
         Caption         =   "Window Info"
      End
   End
End
Attribute VB_Name = "frmWinHack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project WinHack
'Copyright Tretyakov Konstantin (kt_ee@yahoo.com)
'You may use this code for free, if you give me some credit
'At least remember, thet it is not fair to put your name on what you didn't do

'And I would surely appreciate, if you mail me the program (or link to it)
'you created, using this code, (or if you somehow modified this one)

Option Explicit


 'Two Menus Are Not Visible on this form
 'First is mnuReport (Caption = "&Report")
 'If it's visible, then the user may "check" or "uncheck" it, and if it's checked
 'a MsgBox appears after every window operation with the result, returned by the function
 'Second menu, that's not visible is mnuCloseWindow
 'it's function is the same as of mnuMinimizeWindow
 
Private Sub Label1_DblClick(Index As Integer)
'Report window status
If Index Then
Dim i%
For i = 1 To Label1.UBound
Label1(i).BackColor = IIf(IsWindowVisible(Label1(i).Tag) <> 0, Label1(0).BackColor, Label1(0).BackColor + ColChange)
Next
Label1(Index).BackColor = vbRed
m_hWnd = Label1(Index).Tag
mnuSeeHandle_Click
End If
End Sub

'Dim res As Long

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 2 And Index Then
Dim i%
For i = 1 To Label1.UBound
Label1(i).BackColor = IIf(IsWindowVisible(Label1(i).Tag) <> 0, Label1(0).BackColor, Label1(0).BackColor + ColChange)
Next
Label1(Index).BackColor = vbRed
m_hWnd = Label1(Index).Tag
PopupMenu mnuLabel
End If

End Sub

'Private Sub mnuAbout_Click()
'MsgBox "Window Manager V" & App.Major & "." & App.Minor & App.Revision & vbCrLf & "Copyright© 1998, Tretyakov Konstantin" & vbCrLf & "All Rights Reserved"
'End Sub

Private Sub mnuAlwaysOnTopFalse_Click()
If SetTopWindow(m_hWnd, False) = True And m_hWnd = Me.hwnd Then MeOnTop = False
End Sub

Private Sub mnuAlwaysOnTopTrue_Click()
   If SetTopWindow(m_hWnd, True) = True And m_hWnd = Me.hwnd Then MeOnTop = True
End Sub

Private Sub mnuBringToTop_Click()
BringWindowToTop m_hWnd
End Sub

'Private Sub mnuCloseWindow_Click()
'MB (CloseWindow(m_hWnd))
'End Sub

'Private Sub mnuDestroyWindow_Click()
'MB (DestroyWindow(m_hWnd))
'End Sub

'Only later I found out, that 'FlashWindow' is used by
'the programs to turn user's attention to them
'by flashing on the taskbar

Private Sub mnuFlashWindowFalse_Click()
FlashWindow m_hWnd, False
End Sub

Private Sub mnuFlashWindowTrue_Click()
FlashWindow m_hWnd, True
End Sub

'
Private Sub mnuHideWindow_Click()
ShowWindow m_hWnd, 0 'nCmdShow: 0 = hide, 1 = restore, 2 = minimize, 3 = maximize

End Sub

'Private Sub mnuIsIconic_Click()
'MsgBox IsIconic(m_hWnd)
'End Sub

'Private Sub mnuIsZoomed_Click()
'MsgBox IsZoomed(m_hWnd)
'End Sub
'
'Private Sub mnuOpenIcon_Click()
'MB (OpenIcon(m_hWnd))
'End Sub
'
Private Sub mnuMaximizeWindow_Click()
ShowWindow m_hWnd, 3 'nCmdShow: 0 = hide, 1 = restore, 2 = minimize, 3 = maximize

End Sub

Private Sub mnuMinimizeWindow_Click()
ShowWindow m_hWnd, 6 'nCmdShow: 0 = hide, 1 = restore, 2 = minimize, 3 = maximize

End Sub

Private Sub mnuMoveWindow_Click()
    Load frmMove
    With frmMove
    .lblInfo.Caption = "Caption: " & GetWindowCaption(m_hWnd) & vbCrLf & "Class: " & GetClass(m_hWnd) & vbCrLf & "Handle: " & m_hWnd & "  (" & Hex(m_hWnd) & "h)" & vbCrLf & "Visible: " & IIf(IsWindowVisible(m_hWnd) <> 0, "Yes", "No")
    Dim a&, b&, c&, d&
    GetCoords m_hWnd, a, b, c, d
    .txtCoord(0) = CStr(a)
    .txtCoord(1) = CStr(b)
    .txtCoord(2) = CStr(c - a)
    .txtCoord(3) = CStr(d - b)
    .Show 1, Me
    End With
End Sub

Private Sub mnuPostMessage_Click()
PostMessage m_hWnd, WM_CLOSE, 0, 0&
End Sub

'Private Sub mnuReport_Click()
'mnuReport.Checked = Not mnuReport.Checked
'End Sub

Private Sub mnuRestoreWindow_Click()
ShowWindow m_hWnd, 9 'nCmdShow: 0 = hide, 1 = restore, 2 = minimize, 3 = maximize

End Sub

Private Sub mnuSeeHandle_Click()
Dim a&, b&, c&, d&
GetCoords m_hWnd, a, b, c, d
'MsgBox "Caption: " & GetWindowCaption(m_hWnd) & vbCrLf & "Class: " & GetClass(m_hWnd) & vbCrLf & "Handle: " & m_hWnd & "  (" & "&H" & Hex(m_hWnd) & ")" & vbCrLf & "Visible: " & IIf(IsWindowVisible(m_hWnd) <> 0, "Yes", "No") & vbCrLf & vbCrLf & "Window Coordinates:" & vbCrLf & "Left:  " & Chr(9) & CStr(a) & vbCrLf & "Top:   " & Chr(9) & CStr(b) & vbCrLf & "Width: " & Chr(9) & CStr(c - a) & vbCrLf & "Height:" & Chr(9) & CStr(d - b), vbInformation, "Window Manager - Window Info"
MsgBox "Caption: " & GetWindowCaption(m_hWnd) & vbCrLf & "Class: " & GetClass(m_hWnd) & vbCrLf & "Handle: " & m_hWnd & "  (" & Hex(m_hWnd) & "h)" & vbCrLf & "Visible: " & IIf(IsWindowVisible(m_hWnd) <> 0, "Yes", "No") & vbCrLf & vbCrLf & "Window Coordinates:" & vbCrLf & "Left:  " & Chr(9) & CStr(a) & vbCrLf & "Top:   " & Chr(9) & CStr(b) & vbCrLf & "Width: " & Chr(9) & CStr(c - a) & vbCrLf & "Height:" & Chr(9) & CStr(d - b), vbInformation, "Window Manager - Window Info"
End Sub

'Private Sub mnuSetRefresh_Click()
'On Error GoTo erhELP
'Dim a As Long
'a = CLng(InputBox("Please, input new interval in milliseconds (1-30000):", "Set new refresh rate", "1000"))
'If a > 0 And a < 30001 Then Timer1.Interval = a Else Err.Raise 3000
'Exit Sub
'erhELP:
'MsgBox "Wrong Values"
'End Sub

Private Sub mnuSetWindowText_Click()
Dim a As String
Dim c$, d&
d = GetWindowTextLength(m_hWnd) + 1
c = Space(d)
GetWindowText m_hWnd, c, d
a = InputBox("Enter new window caption here:", "Set New Window Caption", Left(c, d))
If a <> "" Then SetWindowText m_hWnd, a
End Sub

Private Sub mnuShowNoActivate_Click()
ShowWindow m_hWnd, 4  'nCmdShow: 0 = hide, 1 = restore, 2 = minimize, 3 = maximize

End Sub

Private Sub mnuShowwindow_Click()
ShowWindow m_hWnd, 5  'nCmdShow: 0 = hide, 1 = restore, 2 = minimize, 3 = maximize

End Sub

Private Sub Timer1_Timer()
'Label1(0).BackColor = QBColor(3)
'Label1(0).FontBold = False
Dim i%
For i = 1 To Label1.Count - 1
Unload Label1(i)
Next
CallBackDemo Picture1
Picture1.Height = Label1(Label1.Count - 1).Top + Label1(Label1.Count - 1).Height + 2
v1.Max = (Picture1.Height - Frame4.Height) / Label1(0).Height + 1
''Label1(0).BackColor = vbmagenta
'Label1(0).FontBold = True
v1.Visible = Picture1.Height > Frame4.Height
End Sub

Private Sub v1_Change()
Picture1.Top = -v1.Value * Label1(0).Height
End Sub


Private Sub chkBMPIcon_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkBMPIcon_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub chkCPanelStart_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkDUpStart_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkIconWrap_Click()
    NeedRestart = True
    cmdApply.Enabled = True
End Sub

Private Sub chkIconWrap_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub chkPRNStart_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkRBinStart_Click()
    cmdApply.Enabled = True
End Sub

Private Sub chkWinAni_Click()
    NeedRestart = True
    cmdApply.Enabled = True
End Sub

Private Sub chkWinAni_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub


Private Sub cmdApply_Click()
    On Error GoTo ErrorHandler
    cmdApply.Enabled = False
    SetStringKey HKEY_CURRENT_USER, MinAnim, "MinAnimate", chkWinAni.Value
    SetStringKey HKEY_CURRENT_USER, MinAnim, "IconTitleWrap", chkIconWrap.Value
    SetStringKey HKEY_CURRENT_USER, MenuDelay, "MenuShowDelay", txtMenuDelay
    SetStringKey HKEY_LOCAL_MACHINE, MainRoot & RecycleBin, , txtRecycle
    SetStringKey HKEY_LOCAL_MACHINE, MainRoot & ControlPanel, , txtControlPanel
    SetStringKey HKEY_LOCAL_MACHINE, MainRoot & PrintersReg, , txtPrinters
    SetStringKey HKEY_LOCAL_MACHINE, MainRoot & DialUp, , txtDupName
    SetStringKey HKEY_LOCAL_MACHINE, WinInfo, "RegisteredOwner", txtOwner
    SetStringKey HKEY_LOCAL_MACHINE, WinInfo, "RegisteredOrganization", txtOrg
    
    Dim Directory As String
    Directory = SysRoot & "\Start Menu" & "\*." & ControlPanel
    If (chkCPanelStart.Value = 1) Then
        If Dir(Directory, vbDirectory) = "" Then
        MkDir SysRoot & "\Start Menu" & CPSN & ControlPanel
        End If
    Else
        If Dir(Directory, vbDirectory) <> "" Then
        RmDir SysRoot & "\Start Menu\" & Dir(Directory, vbDirectory)
        End If
    End If
    
    Directory = SysRoot & "\Start Menu" & "\*." & PrintersReg
    If (chkPRNStart.Value = 1) Then
        If Dir(Directory, vbDirectory) = "" Then
        MkDir SysRoot & "\Start Menu" & PRNSN & PrintersReg
        End If
    Else
        If Dir(Directory, vbDirectory) <> "" Then
        RmDir SysRoot & "\Start Menu\" & Dir(Directory, vbDirectory)
        End If
    End If
    
    Directory = SysRoot & "\Start Menu" & "\*." & DialUp
    If (chkDUpStart.Value = 1) Then
        If Dir(Directory, vbDirectory) = "" Then
        MkDir SysRoot & "\Start Menu" & DUPNSN & DialUp
        End If
    Else
        If Dir(Directory, vbDirectory) <> "" Then
        RmDir SysRoot & "\Start Menu\" & Dir(Directory, vbDirectory)
        End If
    End If
    
    Directory = SysRoot & "\Start Menu" & "\*." & RecycleBin
    If (chkRBinStart.Value = 1) Then
        If Dir(Directory, vbDirectory) = "" Then
        MkDir SysRoot & "\Start Menu" & RBSN & RecycleBin
        End If
    Else
        If Dir(Directory, vbDirectory) <> "" Then
        RmDir SysRoot & "\Start Menu\" & Dir(Directory, vbDirectory)
        End If
    End If
    
    If chkBMPIcon.Value = 1 Then
        SetStringKey HKEY_LOCAL_MACHINE, BmpView, , "%1"
    Else
        SetStringKey HKEY_LOCAL_MACHINE, BmpView, , GetStringKey(HKEY_LOCAL_MACHINE, WinInfo, "ProgramFilesDir") & "\Accessories\MSPAINT.EXE,1"
    End If
    MsgBox "Some changes may take effect only after you restart your computer.", vbInformation, "Important !"
    Exit Sub
ErrorHandler:
    MsgBox "Something Wrong !", vbCritical, "Error"
    Err.Clear
    Resume Next
End Sub

Private Sub cmdOK_Click()
    If cmdApply.Enabled Then cmdApply_Click
'    If NeedRestart Then
'        If MsgBox("Some changes will take effect only after you restart your computer." & vbCrLf & "Do you wish to restart your computer now ?", vbYesNo + vbQuestion + vbDefaultButton2, "Restart ?") = vbYes Then
'            If RestartComputer = False Then MsgBox "Unable to restart system", vbCritical, "Error"
'        End If
'    End If
    Unload Me
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdApply_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    
End Sub

Private Sub Form_Load()
'Load all the values
    On Error GoTo ErrorHandler
    SysRoot = GetStringKey(HKEY_LOCAL_MACHINE, WinInfo, "SystemRoot")
    Tab1.ZOrder 1
    lblAbout.Caption = "Mini WinHacker 95 (with Window Manager™), Version 1.03" & vbCrLf & "Copyright© 1998, Tretyakov Konstantin (kt_ee@yahoo.com)" & vbCrLf & vbCrLf & "This is a shareware." & vbCrLf & vbCrLf & "For full documentation on Mini WinHacker 95 call me (52-05-46)" & vbCrLf & vbCrLf & "Warning: This program is not protected by copyright law, and international treaties. Unauthorized reproduction or distribution of this program,or any portion of it, may not result in severe civil and criminal penalties, and will not be prosecuted to the maximum extent possible under law!!!"
    lblMenuNote.Caption = "0 = Fastest / 1000 = Slowest" & vbCrLf & "Default is around 400 ms"
    chkWinAni.Value = Val(GetStringKey(HKEY_CURRENT_USER, MinAnim, "MinAnimate"))
    chkIconWrap.Value = Val(GetStringKey(HKEY_CURRENT_USER, MinAnim, "IconTitleWrap"))
    chkBMPIcon.Value = IIf(GetStringKey(HKEY_LOCAL_MACHINE, BmpView) = "%1", 1, 0)
    txtMenuDelay = GetStringKey(HKEY_CURRENT_USER, MenuDelay, "MenuShowDelay")
    txtRecycle = GetStringKey(HKEY_LOCAL_MACHINE, MainRoot & RecycleBin)
    txtControlPanel = GetStringKey(HKEY_LOCAL_MACHINE, MainRoot & ControlPanel)
    txtPrinters = GetStringKey(HKEY_LOCAL_MACHINE, MainRoot & PrintersReg)
    txtDupName = GetStringKey(HKEY_LOCAL_MACHINE, MainRoot & DialUp)
    txtOwner = GetStringKey(HKEY_LOCAL_MACHINE, WinInfo, "RegisteredOwner")
    txtOrg = GetStringKey(HKEY_LOCAL_MACHINE, WinInfo, "RegisteredOrganization")
    chkCPanelStart.Value = IIf(Dir(SysRoot & "\Start Menu" & "\*." & ControlPanel, vbDirectory) = "", 0, 1)
    chkPRNStart.Value = IIf(Dir(SysRoot & "\Start Menu" & "\*." & PrintersReg, vbDirectory) = "", 0, 1)
    chkDUpStart.Value = IIf(Dir(SysRoot & "\Start Menu" & "\*." & DialUp, vbDirectory) = "", 0, 1)
    chkRBinStart.Value = IIf(Dir(SysRoot & "\Start Menu" & "\*." & RecycleBin, vbDirectory) = "", 0, 1)
    
    cmdApply.Enabled = False
    Exit Sub
ErrorHandler:
MsgBox "Something Wrong !", vbCritical, "Error !"
End Sub

Private Sub Tab1_Click()
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To Tab1.Tabs.Count - 1
        If i = Tab1.SelectedItem.Index - 1 Then
'            Frame(i).Left = 210
'            Frame(i).Enabled = True
            Frame(i).Visible = True
        Else
'            Frame(i).Left = -20000
'            Frame(i).Enabled = False
            Frame(i).Visible = False
        End If
    Next
    If (Tab1.SelectedItem.Index = 3) Then
        Timer1_Timer
        Timer1.Enabled = True
    End If
End Sub

Private Sub Tab1_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub txtControlPanel_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtDupName_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtMenuDelay_Change()
    NeedRestart = True
    cmdApply.Enabled = True
    If Val(txtMenuDelay) > 1000 Then txtMenuDelay = "1000": Exit Sub
    If Val(txtMenuDelay) < 0 Then txtMenuDelay = "0": Exit Sub
    If IsNumeric(txtMenuDelay) = False Or InStr(txtMenuDelay, ",") <> 0 Then txtMenuDelay = Val(txtMenuDelay)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = Tab1.SelectedItem.Index
        If i = Tab1.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set Tab1.SelectedItem = Tab1.Tabs(1)
        Else
            'increment the tab
            Set Tab1.SelectedItem = Tab1.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub txtMenuDelay_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub txtOrg_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtOwner_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtPrinters_Change()
    cmdApply.Enabled = True
End Sub

Private Sub txtRecycle_Change()
    cmdApply.Enabled = True
    NeedRestart = True
End Sub
