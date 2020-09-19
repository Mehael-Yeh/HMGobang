VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "更新日志"
   ClientHeight    =   4215
   ClientLeft      =   5940
   ClientTop       =   3975
   ClientWidth     =   6390
   ClipControls    =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6390
   StartUpPosition =   2  '屏幕中心
   Begin VB.VScrollBar VScroll1 
      Height          =   4215
      LargeChange     =   5000
      Left            =   4600
      Max             =   30000
      SmallChange     =   70
      TabIndex        =   3
      Top             =   0
      Width           =   450
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5160
      Top             =   2880
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4185
      ScaleWidth      =   4425
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "更新日志"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5400
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Form2
End Sub


Private Sub Form_Load()
    OldProcAddr = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf MyWinProc)
End Sub

Private Sub Timer1_Timer()
    Picture1.FontSize = 28
    Picture1.CurrentX = 100
    Picture1.CurrentY = 100
    Picture1.ForeColor = &H80&
    Picture1.Print "V0 Ace 艾斯"
        Picture1.FontSize = 15
        Picture1.ForeColor = vbBlack
        Picture1.Print
        Picture1.CurrentX = 120
        Picture1.Print "V0.1.0a Build 3/8/2020"
            Picture1.Print
            Picture1.FontSize = 12
            Picture1.CurrentX = 120
            Picture1.Print "1.主程序代码已达到1024行，初步实现"
            Picture1.CurrentX = 250
            Picture1.Print "了积分机制，能辨识棋局中的连珠个数"
            Picture1.Print
            Picture1.CurrentX = 120
            Picture1.Print "2.加入“更新日志”"
        Picture1.FontSize = 15
        Picture1.Print
        Picture1.CurrentX = 120
        Picture1.Print "V0.1.0b Build 8/8/2020"
            Picture1.Print
            Picture1.FontSize = 12
            Picture1.CurrentX = 120
            Picture1.Print "1.修正“横孤二”时无法判断的情况"
            Picture1.Print
            Picture1.CurrentX = 120
            Picture1.Print "2.第一阶段AI适配，初步实现人机对弈"
            Picture1.CurrentX = 120
            Picture1.Print "3.加入AI落子显示"
        Picture1.FontSize = 15
        Picture1.Print
        Picture1.CurrentX = 120
        Picture1.Print "V0.1.1 Build 26/8/2020"
            Picture1.Print
            Picture1.FontSize = 12
            Picture1.CurrentX = 120
            Picture1.Print "1.简单模式AI完全适配！"
            Picture1.Print
            Picture1.CurrentX = 120
            Picture1.Print "2.修正落子显示重开后无法清除的情况"
        Picture1.FontSize = 15
        Picture1.Print
        Picture1.CurrentX = 120
        Picture1.Print "V0.1.2a Build 28/8/2020"
            Picture1.Print
            Picture1.FontSize = 12
            Picture1.CurrentX = 120
            Picture1.Print "1.去除当前胜率显示，加入悔棋按钮"
            Picture1.Print
            Picture1.CurrentX = 120
            Picture1.Print "2.修正悔棋系统"
            Picture1.Print
            Picture1.CurrentX = 120
            Picture1.Print "3.取消胜利判定后的退出引导"
            Picture1.Print
            Picture1.CurrentX = 120
            Picture1.Print "4.调整面板属性，修复部分Bug；部分Bu"
            Picture1.CurrentX = 250
            Picture1.Print "g问题已确定，还未定位"
            Picture1.Print
            Picture1.CurrentX = 120
            Picture1.Print "5.主程序代码已达到2982行，五子棋初"
            Picture1.CurrentX = 250
            Picture1.Print "代模版完全架构完成"
            
End Sub

Private Sub VScroll1_Change()
    Picture1.Top = -VScroll1.Value
    Picture1.Height = Picture1.Height + VScroll1.Value
End Sub
