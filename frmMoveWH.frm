VERSION 5.00
Begin VB.Form frmMove 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Window Manager - Move Window"
   ClientHeight    =   2715
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   4890
   Icon            =   "frmMoveWH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCoord 
      Height          =   288
      Index           =   3
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   8
      Top             =   1920
      Width           =   552
   End
   Begin VB.TextBox txtCoord 
      Height          =   288
      Index           =   2
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "0"
      Top             =   1620
      Width           =   552
   End
   Begin VB.TextBox txtCoord 
      Height          =   288
      Index           =   1
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "0"
      Top             =   1320
      Width           =   552
   End
   Begin VB.TextBox txtCoord 
      Height          =   288
      Index           =   0
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "0"
      Top             =   1020
      Width           =   552
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   312
      Left            =   2520
      TabIndex        =   10
      Top             =   2340
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Move"
      Default         =   -1  'True
      Height          =   312
      Left            =   1320
      TabIndex        =   9
      Top             =   2340
      Width           =   1092
   End
   Begin VB.Label lblCoord 
      Caption         =   "Window &height:"
      Height          =   252
      Index           =   3
      Left            =   180
      TabIndex        =   7
      Top             =   1980
      Width           =   2232
   End
   Begin VB.Label lblCoord 
      Caption         =   "Window &width:"
      Height          =   252
      Index           =   2
      Left            =   180
      TabIndex        =   5
      Top             =   1680
      Width           =   2232
   End
   Begin VB.Label lblCoord 
      Caption         =   "Top left corner &Y coordinate:"
      Height          =   252
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   1380
      Width           =   2232
   End
   Begin VB.Label lblCoord 
      Caption         =   "Top left corner &X coordinate:"
      Height          =   252
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   1080
      Width           =   2232
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Window:"
      Height          =   912
      Left            =   120
      TabIndex        =   0
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   4632
   End
End
Attribute VB_Name = "frmMove"
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

'Winhack dialog to move a window

Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If MoveWindow(m_hWnd, Val(txtCoord(0)), Val(txtCoord(1)), Val(txtCoord(2)), Val(txtCoord(3)), 1) = 0 Then MsgBox "Unable to perform operation !", vbCritical, "Error !"
    Unload Me
End Sub

Private Sub Form_Load()
    If MeOnTop Then SetTopWindow Me.hwnd, True
End Sub

