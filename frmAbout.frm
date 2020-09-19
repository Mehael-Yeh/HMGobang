VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于慧木五子棋"
   ClientHeight    =   4545
   ClientLeft      =   2340
   ClientTop       =   2235
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmAbout.frx":424A
   ScaleHeight     =   3137.04
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "华文宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1320
      Width           =   5415
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   465
      Left            =   4080
      TabIndex        =   0
      Top             =   3360
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "官网(&G)"
      Height          =   465
      Left            =   4080
      TabIndex        =   1
      Top             =   3960
      Width           =   1485
   End
   Begin VB.Label Label1 
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
      Left            =   720
      TabIndex        =   7
      Top             =   720
      Width           =   450
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1380
   End
   Begin VB.Label Label1 
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
      TabIndex        =   5
      Top             =   120
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   2236.306
      Y2              =   2236.306
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
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
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
      Left            =   1800
      TabIndex        =   4
      Top             =   840
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "敬告: "
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   3630
   End
   Begin VB.Menu V 
      Caption         =   "版本"
      Begin VB.Menu GXRZ 
         Caption         =   "更新日志"
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload frmAbout
End Sub

Private Sub cmdSysInfo_Click()
    On Error GoTo Err
    Call ShellExecute(hwnd, "open", "www.HMGobang.lovelyh.xyz", vbNullString, vbNullString, conSwNormal)
Err:
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "版本：" + BB1 + " " + BB2
    Text1.Text = "本软件创作时间：19/7/2020-" & vbCrLf & "v0 Ace，中文为艾斯，即火拳艾斯，意以纪念前女友（初恋）喜欢的《海贼王》中的一名角色。" & vbCrLf & "游戏设计明显比较难的部分应该是棋盘的生成和AI的设计，棋盘设计上的生成系统由于覆盖的关系，225个点位调试起来很麻烦，所以在代码中用了很多占内存的循环语句，当然一定程度上也加强了它的稳定性；AI的设计上借鉴了很多大佬的Python逻辑构造，不得不让我感叹Visual Basic是一门过时的语言了，但VB的易用性是毋庸置疑的。在最后一个小版本我应该会加上“禁手”这个功能，不过要很多时间去调试就是了。因为我本身也仅仅是想做个很基础的五子棋游戏，并没有想过把这个软件做得太完美，它可以很完美，但永远不是我心目中最完美的，保留那种感觉就是最美好的。这也是我开源的一部分原因。" & vbCrLf & "由于我的个人原因，有些人错过了，并且可能不再回来，而这个软件也就只停留在第一大版本，从此不再更新，这是我最后的浪漫，感谢你一路的相伴。" & vbCrLf & "特别鸣谢：邱旺、李义圣等对本软件功能和代码开发上提出的建议。"
    Text1.Locked = True
    lblDisclaimer.Caption = "敬告：" & vbCrLf & vbCrLf & "慧木菌 版权所有(C)" & vbCrLf & vbCrLf & "本软件仅供学习使用，有任何意见可联系作者QQ:1175592624"
End Sub

Private Sub GXRZ_Click()
    Load Form2
    Form2.Show
End Sub

