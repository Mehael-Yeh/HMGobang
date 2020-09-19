VERSION 5.00
Begin VB.Form Dialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "慧木五子棋(HMGobang)"
   ClientHeight    =   1890
   ClientLeft      =   6450
   ClientTop       =   4770
   ClientWidth     =   6030
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Dialog.frx":1084A
   ScaleHeight     =   1890
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   3500
      Left            =   120
      Top             =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "HMGobang.lovelyh.xyz"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   180
      Left            =   120
      MouseIcon       =   "Dialog.frx":31DF4
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1560
      Width           =   1800
   End
   Begin VB.Label lblTitle 
      Caption         =   "慧木五子棋(HMGobang)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "慧木菌 版权所有(C) 保留最终解释权 "
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label LabelD 
      AutoSize        =   -1  'True
      Caption         =   "子"
      BeginProperty Font 
         Name            =   "田氏颜体大字库"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   510
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   450
   End
   Begin VB.Label LabelD 
      AutoSize        =   -1  'True
      Caption         =   "五  棋"
      BeginProperty Font 
         Name            =   "田氏颜体大字库"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label LabelD 
      AutoSize        =   -1  'True
      Caption         =   "慧木"
      BeginProperty Font 
         Name            =   "田氏颜体大字库"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   510
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Form1.Show
    Unload Dialog
End Sub

Private Sub Form_Load()
    BB1 = "v0.1.2a Ace"
    BB2 = "Build 28/8/2020"
    lblVersion.Caption = "版本：" + BB1 + " " + BB2
End Sub

Private Sub Label1_Click()
    Form1.Show
    Unload Dialog
End Sub

Private Sub Label2_Click()
    On Error GoTo Err
    Call ShellExecute(hwnd, "open", "www.HMGobang.lovelyh.xyz", vbNullString, vbNullString, conSwNormal)
Err:
End Sub


Private Sub LabelD_Click(Index As Integer)
    frmAbout.Show
End Sub

Private Sub lblTitle_Click()
    Form1.Show
    Unload Dialog
End Sub

Private Sub lblVersion_Click()
    Form1.Show
    Unload Dialog
End Sub

Private Sub Timer1_Timer()
    Form1.Show
    Unload Dialog
End Sub
