VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "慧木五子棋(HMGobang)"
   ClientHeight    =   7590
   ClientLeft      =   3900
   ClientTop       =   1635
   ClientWidth     =   9720
   Icon            =   "HMGobang.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9720
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9240
      Top             =   7200
   End
   Begin VB.CommandButton Command3 
      Caption         =   "悔棋"
      BeginProperty Font 
         Name            =   "魚石行書"
         Size            =   21.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7850
      TabIndex        =   239
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "白子"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7560
      TabIndex        =   236
      Top             =   4440
      Width           =   2055
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000C0&
         Caption         =   "坐下(&W)"
         BeginProperty Font 
            Name            =   "字酷堂黄楷体(个人非商业版)"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   237
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "仿宋"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   238
         Top             =   700
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "黑子"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7560
      TabIndex        =   233
      Top             =   2520
      Width           =   2055
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000C0&
         Caption         =   "坐下(&B)"
         BeginProperty Font 
            Name            =   "字酷堂黄楷体(个人非商业版)"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   235
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "仿宋"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   234
         Top             =   700
         Width           =   195
      End
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   7440
      Top             =   7200
   End
   Begin VB.PictureBox Picture1 
      FillColor       =   &H80000001&
      Height          =   7215
      Left            =   0
      ScaleHeight     =   7155
      ScaleWidth      =   7275
      TabIndex        =   5
      Top             =   0
      Width           =   7335
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   224
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   230
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   223
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   229
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   222
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   228
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   221
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   227
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   220
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   226
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   219
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   225
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   218
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   224
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   217
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   223
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   216
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   222
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   215
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   221
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   214
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   220
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   213
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   219
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   212
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   218
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   211
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   217
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   210
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   216
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   209
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   215
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   208
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   214
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   207
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   213
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   206
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   212
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   205
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   211
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   204
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   210
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   203
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   209
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   202
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   208
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   201
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   207
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   200
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   206
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   199
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   205
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   198
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   204
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   197
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   203
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   196
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   202
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   195
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   201
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   194
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   200
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   193
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   199
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   192
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   198
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   191
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   197
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   190
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   196
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   189
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   195
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   188
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   194
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   187
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   193
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   186
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   192
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   185
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   191
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   184
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   190
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   183
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   189
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   182
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   188
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   181
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   187
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   180
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   186
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   179
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   185
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   178
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   184
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   177
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   183
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   176
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   182
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   175
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   181
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   174
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   180
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   173
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   179
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   172
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   178
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   171
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   177
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   170
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   176
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   169
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   175
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   168
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   174
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   167
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   173
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   166
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   172
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   165
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   171
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   164
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   170
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   163
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   169
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   162
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   168
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   161
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   167
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   160
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   166
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   159
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   165
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   158
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   164
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   157
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   163
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   156
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   162
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   155
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   161
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   154
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   160
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   153
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   159
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   152
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   158
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   151
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   157
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   150
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   156
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   149
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   155
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   0
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   154
         Top             =   0
         Width           =   420
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   420
         Index           =   0
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   1
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   153
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   2
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   152
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   3
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   151
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   4
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   150
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   5
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   149
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   6
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   148
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   7
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   147
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   8
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   146
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   9
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   145
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   10
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   144
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   11
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   143
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   12
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   142
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   13
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   141
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   14
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   140
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   15
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   139
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   16
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   138
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   17
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   137
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   18
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   136
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   19
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   135
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   20
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   134
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   21
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   133
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   22
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   132
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   23
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   131
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   24
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   130
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   25
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   129
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   26
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   128
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   27
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   127
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   28
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   126
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   29
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   125
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   30
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   124
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   31
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   123
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   32
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   122
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   33
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   121
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   34
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   120
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   35
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   119
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   36
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   118
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   37
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   117
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   38
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   116
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   39
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   115
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   40
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   114
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   41
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   113
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   42
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   112
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   43
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   111
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   44
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   110
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   45
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   109
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   46
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   108
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   47
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   107
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   48
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   106
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   49
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   105
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   50
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   104
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   51
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   103
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   52
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   102
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   53
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   101
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   54
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   100
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   55
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   99
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   56
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   98
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   57
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   97
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   58
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   96
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   59
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   95
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   60
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   94
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   61
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   93
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   62
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   92
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   63
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   91
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   64
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   90
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   65
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   89
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   66
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   88
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   67
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   87
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   68
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   86
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   69
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   85
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   70
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   84
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   71
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   83
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   72
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   82
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   73
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   81
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   74
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   80
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   75
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   79
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   76
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   78
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   77
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   77
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   78
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   76
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   79
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   75
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   80
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   74
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   81
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   73
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   82
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   72
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   83
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   71
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   84
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   70
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   85
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   69
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   86
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   68
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   87
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   67
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   88
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   66
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   89
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   65
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   90
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   64
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   91
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   63
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   92
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   62
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   93
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   61
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   94
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   60
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   95
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   59
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   96
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   58
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   97
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   57
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   98
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   56
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   99
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   55
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   100
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   54
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   101
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   53
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   102
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   52
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   103
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   51
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   104
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   50
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   105
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   49
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   106
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   48
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   107
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   47
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   108
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   46
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   109
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   45
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   110
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   44
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   111
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   43
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   112
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   42
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   113
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   41
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   114
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   40
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   115
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   39
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   116
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   38
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   117
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   37
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   118
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   36
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   119
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   35
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   120
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   34
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   121
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   33
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   122
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   32
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   123
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   31
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   124
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   30
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   125
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   29
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   126
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   28
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   127
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   27
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   128
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   26
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   129
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   25
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   130
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   24
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   131
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   23
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   132
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   22
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   133
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   21
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   134
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   20
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   135
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   19
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   136
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   18
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   137
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   17
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   138
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   16
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   139
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   15
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   140
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   14
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   141
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   13
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   142
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   12
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   143
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   11
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   144
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   10
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   145
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   9
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   146
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   8
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   147
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   7
         Top             =   0
         Width           =   420
      End
      Begin VB.Label Lab1 
         BackStyle       =   0  'Transparent
         Height          =   420
         Index           =   148
         Left            =   0
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   0
         Width           =   420
      End
   End
   Begin VB.ComboBox combol1 
      BeginProperty Font 
         Name            =   "苏新诗柳楷简"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "HMGobang.frx":1084A
      Left            =   8040
      List            =   "HMGobang.frx":10857
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   8880
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Left            =   8400
      Top             =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "举棋不悔大丈夫："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7560
      TabIndex        =   240
      Top             =   6240
      Width           =   1680
   End
   Begin VB.Label TS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   232
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label JDT 
      BackColor       =   &H009BC679&
      Height          =   375
      Left            =   0
      TabIndex        =   231
      Top             =   7200
      Width           =   15
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "当前难度："
      BeginProperty Font 
         Name            =   "苏新诗柳楷简"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7440
      TabIndex        =   3
      Top             =   1320
      Width           =   1800
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
      Left            =   8400
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
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
      Left            =   7920
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
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
      Left            =   7440
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Menu KS 
      Caption         =   "开始(&S)"
      Begin VB.Menu ND 
         Caption         =   "难度"
         Begin VB.Menu Game 
            Caption         =   "简单"
            Index           =   0
         End
         Begin VB.Menu Game 
            Caption         =   "中等"
            Index           =   1
         End
         Begin VB.Menu Game 
            Caption         =   "困难"
            Index           =   2
         End
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu YXGZ 
         Caption         =   "游戏规则"
      End
      Begin VB.Menu GY 
         Caption         =   "关于HMGobang"
      End
   End
   Begin VB.Menu SZ 
      Caption         =   "设置"
      Visible         =   0   'False
      Begin VB.Menu NDF 
         Caption         =   "难度调整"
         Begin VB.Menu Game1 
            Caption         =   "简单"
            Index           =   0
         End
         Begin VB.Menu Game1 
            Caption         =   "中等"
            Index           =   1
         End
         Begin VB.Menu Game1 
            Caption         =   "困难"
            Index           =   2
         End
      End
      Begin VB.Menu fjx 
         Caption         =   "-"
      End
      Begin VB.Menu Quit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Remember() As Integer '记局悔棋系统
Private Sub combol1_Click()
    If Di <> -1 And N = 0 Then    '内部启动
        ans = MsgBox("您真的要开启新的对局吗？", 1 + 0 + 32, "慧木五子棋(HMGobobang)")      '重置弹窗
        If ans = 1 Then '重置确认
            Lab1(V).BorderStyle = 0 '落子位置显示清除
            Sh = -1     '动画起始
            Label1(0).Top = -400
            Label1(1).Left = 7560
            Label1(2).Left = 8780
            Timer1.Enabled = True
            Timer1.Interval = 30
            Di = combol1.ListIndex
            Difficility = combol1.ListIndex + 1  '难度设置
            gameover = False: way = 0
            Call Clsp
            Call Newz
            Call JD
            Call ZX
            Call FB
        ElseIf ans = 2 Then     '重置取消
            Cancel = True
        End If
    ElseIf Di <> -1 And N = 1 Then    '外部启动
        Lab1(V).BorderStyle = 0 '落子位置显示清除
        Sh = -1     '动画起始
        Label1(0).Top = -400
        Label1(1).Left = 7560
        Label1(2).Left = 8780
        Timer1.Enabled = True
        Timer1.Interval = 30
        Di = combol1.ListIndex
        Difficility = combol1.ListIndex + 1  '难度设置
        gameover = False: way = 0
        Call Clsp
        Call Newz
        Call JD
        Call ZX
        Call FB
    ElseIf Di = -1 Then     '原始启动
        Lab1(V).BorderStyle = 0 '落子位置显示清除
        Sh = -1     '动画起始
        Label1(0).Top = -400
        Label1(1).Left = 7560
        Label1(2).Left = 8780
        Timer1.Enabled = True
        Timer1.Interval = 30
        Di = combol1.ListIndex
        Difficility = combol1.ListIndex + 1  '难度设置
        gameover = False: way = 0
        Call Clsp
        Call Newz
        Call JD
        Call ZX
        Call FB
    End If
End Sub

Private Sub Command1_Click()
    Label3.Caption = "人脑"
    Label4.Caption = "电脑"
    Command1.Visible = False
    Command2.Visible = False
    Call JH
    Color_P = &H0&  '人的棋子为黑色
    Color_C = &HFFFFFF  '机器的棋子颜色为白色
End Sub

Private Sub Command2_Click()
    Label3.Caption = "电脑"
    Label4.Caption = "人脑"
    Command1.Visible = False
    Command2.Visible = False
    Call JH
    Color_C = &H0&  '机器的棋子颜色为黑色
    Color_P = &HFFFFFF  '人的棋子为白色
End Sub

Private Sub Command3_Click()
    Ren_Y = (Remember(chess - 1) Mod 15 + 1) * 494 - 234
    Ren_X = (Remember(chess - 1) \ 15 + 1) * 494 - 234
    Jiqi_Y = (Remember(chess) Mod 15 + 1) * 494 - 234
    Jiqi_X = (Remember(chess) \ 15 + 1) * 494 - 234
    Shape1(Remember(chess)).Visible = False    '机器悔棋开始
    Shape1(Remember(chess)).FillColor = &H8000&
    Lab1(Remember(chess)).BorderStyle = 0 '落子位置显示清除
    Lab1(Remember(chess)).Enabled = True    '机器悔棋结束
    Shape1(Remember(chess - 1)).Visible = False '人悔棋开始
    Shape1(Remember(chess - 1)).FillColor = &H8000&
    Lab1(Remember(chess - 1)).Enabled = True  '人悔棋结束
    Timer4.Enabled = True
    Timer4.Interval = 10
    chess = chess - 2
    If chess < 2 Then
        Command3.Enabled = False
    Else
        Command3.Enabled = True
    End If
    way = 0 '清除胜负判定
End Sub
Private Sub Timer4_Timer()
    Call Fixp
End Sub

Public Sub Fixp()
    If Ren_X <> 260 And Ren_X <> 7176 And Ren_Y <> 260 And Ren_Y <> 7176 Then '中间补棋盘(人)
        Picture1.Line (Ren_Y, Ren_X - 245)-(Ren_Y, Ren_X + 245) '棋盘纵列(人)
        Picture1.Line (Ren_Y - 245, Ren_X)-(Ren_Y + 245, Ren_X) '棋盘横列(人)
        Picture1.DrawWidth = 1
    End If
    If Jiqi_X <> 260 And Jiqi_X <> 7176 And Jiqi_Y <> 260 And Jiqi_Y <> 7176 Then  '中间补棋盘(机器)
        Picture1.Line (Jiqi_Y, Jiqi_X - 245)-(Jiqi_Y, Jiqi_X + 245) '棋盘纵列(机器)
        Picture1.Line (Jiqi_Y - 245, Jiqi_X)-(Jiqi_Y + 245, Jiqi_X) '棋盘横列(机器)
        Picture1.DrawWidth = 1
    End If
    If Ren_X = 260 And Ren_Y <> 260 And Ren_Y <> 7176 Then '上边框补棋盘(人)
        Picture1.DrawWidth = 3
        Picture1.Line (Ren_Y - 245, Ren_X)-(Ren_Y + 245, Ren_X) '棋盘横列(人)
        Picture1.DrawWidth = 1
        Picture1.Line (Ren_Y, Ren_X)-(Ren_Y, Ren_X + 245) '棋盘纵列(人)
    ElseIf Ren_X = 7176 And Ren_Y <> 7176 And Ren_Y <> 260 Then '下边框补棋盘(人)
        Picture1.DrawWidth = 3
        Picture1.Line (Ren_Y - 245, Ren_X)-(Ren_Y + 245, Ren_X) '棋盘横列(人)
        Picture1.DrawWidth = 1
        Picture1.Line (Ren_Y, Ren_X - 245)-(Ren_Y, Ren_X) '棋盘纵列(人)
    ElseIf Ren_Y = 260 And Ren_X <> 7176 And Ren_X <> 260 Then '左边框补棋盘(人)
        Picture1.Line (Ren_Y, Ren_X)-(Ren_Y + 245, Ren_X)  '棋盘横列(人)
        Picture1.DrawWidth = 3
        Picture1.Line (Ren_Y, Ren_X - 245)-(Ren_Y, Ren_X + 245) '棋盘纵列(人)
        Picture1.DrawWidth = 1
    ElseIf Ren_Y = 7176 And Ren_X <> 260 And Ren_X <> 7176 Then '右边框补棋盘(人)
        Picture1.Line (Ren_Y - 245, Ren_X)-(Ren_Y, Ren_X)  '棋盘横列(人)
        Picture1.DrawWidth = 3
        Picture1.Line (Ren_Y, Ren_X - 245)-(Ren_Y, Ren_X + 245) '棋盘纵列(人)
        Picture1.DrawWidth = 1
    End If
    If Ren_X = 260 And Ren_Y = 260 Then '上左上角补棋盘(人)
        Picture1.DrawWidth = 3
        Picture1.Line (Ren_Y, Ren_X)-(Ren_Y + 245, Ren_X)  '棋盘横列(人)
        Picture1.Line (Ren_Y, Ren_X)-(Ren_Y, Ren_X + 245) '棋盘纵列(人)
        Picture1.DrawWidth = 1
    ElseIf Ren_X = 7176 And Ren_Y = 7176 Then '右下角补棋盘(人)
        Picture1.DrawWidth = 3
        Picture1.Line (Ren_Y - 245, Ren_X)-(Ren_Y, Ren_X)  '棋盘横列(人)
        Picture1.Line (Ren_Y, Ren_X - 245)-(Ren_Y, Ren_X) '棋盘纵列(人)
        Picture1.DrawWidth = 1
    ElseIf Ren_Y = 260 And Ren_X = 7176 Then '左下角补棋盘(人)
        Picture1.DrawWidth = 3
        Picture1.Line (Ren_Y, Ren_X)-(Ren_Y + 245, Ren_X)  '棋盘横列(人)
        Picture1.Line (Ren_Y, Ren_X - 245)-(Ren_Y, Ren_X)  '棋盘纵列(人)
        Picture1.DrawWidth = 1
    ElseIf Ren_Y = 7176 And Ren_X = 260 Then '右上角补棋盘(人)
        Picture1.DrawWidth = 3
        Picture1.Line (Ren_Y - 245, Ren_X)-(Ren_Y, Ren_X)  '棋盘横列(人)
        Picture1.Line (Ren_Y, Ren_X)-(Ren_Y, Ren_X + 245)  '棋盘纵列(人)
        Picture1.DrawWidth = 1
    End If
    If Jiqi_X = 260 And Jiqi_Y <> 260 And Jiqi_Y <> 7176 Then '上边框补棋盘(机器)
        Picture1.DrawWidth = 3
        Picture1.Line (Jiqi_Y - 245, Jiqi_X)-(Jiqi_Y + 245, Jiqi_X) '棋盘横列(机器)
        Picture1.DrawWidth = 1
        Picture1.Line (Jiqi_Y, Jiqi_X)-(Jiqi_Y, Jiqi_X + 245) '棋盘纵列(机器)
    ElseIf Jiqi_X = 7176 And Jiqi_Y <> 7176 And Jiqi_Y <> 260 Then '下边框补棋盘(机器)
        Picture1.DrawWidth = 3
        Picture1.Line (Jiqi_Y - 245, Jiqi_X)-(Jiqi_Y + 245, Jiqi_X) '棋盘横列(机器)
        Picture1.DrawWidth = 1
        Picture1.Line (Jiqi_Y, Jiqi_X - 245)-(Jiqi_Y, Jiqi_X) '棋盘纵列(机器)
    ElseIf Jiqi_Y = 260 And Jiqi_X <> 7176 And Jiqi_X <> 260 Then '左边框补棋盘(机器)
        Picture1.Line (Jiqi_Y, Jiqi_X)-(Jiqi_Y + 245, Jiqi_X)  '棋盘横列(机器)
        Picture1.DrawWidth = 3
        Picture1.Line (Jiqi_Y, Jiqi_X - 245)-(Jiqi_Y, Jiqi_X + 245) '棋盘纵列(机器)
        Picture1.DrawWidth = 1
    ElseIf Jiqi_Y = 7176 And Jiqi_X <> 260 And Jiqi_X <> 7176 Then '右边框补棋盘(机器)
        Picture1.Line (Jiqi_Y - 245, Jiqi_X)-(Jiqi_Y, Jiqi_X)  '棋盘横列(机器)
        Picture1.DrawWidth = 3
        Picture1.Line (Jiqi_Y, Jiqi_X - 245)-(Jiqi_Y, Jiqi_X + 245) '棋盘纵列(机器)
        Picture1.DrawWidth = 1
    End If
    If Jiqi_X = 260 And Jiqi_Y = 260 Then '上左上角补棋盘(机器)
        Picture1.DrawWidth = 3
        Picture1.Line (Jiqi_Y, Jiqi_X)-(Jiqi_Y + 245, Jiqi_X)  '棋盘横列(机器)
        Picture1.Line (Jiqi_Y, Jiqi_X)-(Jiqi_Y, Jiqi_X + 245) '棋盘纵列(机器)
        Picture1.DrawWidth = 1
    ElseIf Jiqi_X = 7176 And Jiqi_Y = 7176 Then '右下角补棋盘(机器)
        Picture1.DrawWidth = 3
        Picture1.Line (Jiqi_Y - 245, Jiqi_X)-(Jiqi_Y, Jiqi_X)  '棋盘横列(机器)
        Picture1.Line (Jiqi_Y, Jiqi_X - 245)-(Jiqi_Y, Jiqi_X) '棋盘纵列(机器)
        Picture1.DrawWidth = 1
    ElseIf Jiqi_Y = 260 And Jiqi_X = 7176 Then '左下角补棋盘(机器)
        Picture1.DrawWidth = 3
        Picture1.Line (Jiqi_Y, Jiqi_X)-(Jiqi_Y + 245, Jiqi_X)  '棋盘横列(机器)
        Picture1.Line (Jiqi_Y, Jiqi_X - 245)-(Jiqi_Y, Jiqi_X)  '棋盘纵列(机器)
        Picture1.DrawWidth = 1
    ElseIf Jiqi_Y = 7176 And Jiqi_X = 260 Then '右上角补棋盘(机器)
        Picture1.DrawWidth = 3
        Picture1.Line (Jiqi_Y - 245, Jiqi_X)-(Jiqi_Y, Jiqi_X)  '棋盘横列(机器)
        Picture1.Line (Jiqi_Y, Jiqi_X)-(Jiqi_Y, Jiqi_X + 245)  '棋盘纵列(机器)
        Picture1.DrawWidth = 1
    End If
    If Ren_X = 1742 And Ren_Y = 1742 Or Jiqi_X = 1742 And Jiqi_Y = 1742 Then
        Picture1.Circle (1742, 1742), 45
    ElseIf Ren_X = 3718 And Ren_Y = 3718 Or Jiqi_X = 3718 And Jiqi_Y = 3718 Then
        Picture1.Circle (3718, 3718), 45
    ElseIf Ren_X = 5694 And Ren_Y = 5694 Or Jiqi_X = 5694 And Jiqi_Y = 5694 Then
        Picture1.Circle (5694, 5694), 45
    ElseIf Ren_X = 5694 And Ren_Y = 1742 Or Jiqi_X = 5694 And Jiqi_Y = 1742 Then
        Picture1.Circle (1742, 5694), 45
    ElseIf Ren_X = 1742 And Ren_Y = 5694 Or Jiqi_X = 1742 And Jiqi_Y = 5694 Then
        Picture1.Circle (5694, 1742), 45
    End If
    Timer4.Enabled = False
End Sub
Private Sub Form_Load()
    Form1.Caption = "慧木五子棋(HMGobang)" + " " + BB1 + " " + BB2
    Di = -1     '初始启动标志
    Difficility = 0   '难度预存
    Lab1(V).BorderStyle = 0 '落子位置显示清除
    Sh = -1     '动画起始
    Label1(0).Top = -400
    Label1(1).Left = 7560
    Label1(2).Left = 8780
    Timer1.Enabled = True
    Timer1.Interval = 30
    For i = 0 To 224
        Lab1(i).Move -1000, -1000   '避免误触
    Next i
    Score_P = 1
    Score_C = 1     '重置积分
    Command3.Enabled = False    '锁定悔棋按钮
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu SZ     '导入弹出菜单
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ans = MsgBox("您真的要退出吗？", 1 + 0 + 32, "慧木五子棋(HMGobobang)")
    If ans = 1 Then
        End     '退出
    ElseIf ans = 2 Then
        Cancel = True
    End If
End Sub

Private Sub Game_Click(Index As Integer)
    If Di <> -1 Then   '重置判断
        ans = MsgBox("您真的要开启新的对局吗？", 1 + 0 + 32, "慧木五子棋(HMGobobang)")
        If ans = 1 Then     '重置确认
            Lab1(V).BorderStyle = 0 '落子位置显示清除
            N = 1   '外部启动标志
            Call Chongzhi(Index)
            combol1.ListIndex = Index
            Difficility = Index + 1 '难度设置
            gameover = False: way = 0
        ElseIf ans = 2 Then    '重置取消
            Cancel = True
        End If
    Else
        combol1.ListIndex = Index
        Difficility = Index + 1 '难度设置
        gameover = False: way = 0
    End If
End Sub

Private Sub GY_Click()
    frmAbout.Show   '导入关于菜单
End Sub

Private Sub Lab1_Click(Index As Integer)
    chess = chess + 1   '人下棋并录入次数
    ReDim Preserve Remember(chess)
    Remember(chess) = Index
    Lab1(V).BorderStyle = 0 '落子位置显示清除
    Shape1(Index).Visible = True
    Shape1(Index).FillColor = Color_P   '人下棋的颜色
    Shape1(Index).BorderColor = Color_C '棋子边框颜色交换
    Lab1(Index).Enabled = False
    Score_P = 1: Score_C = 1            '置零积分
    Call Score(Score_P, Score_C)
    If chess < 2 Then
        Command3.Enabled = False
    Else
        Command3.Enabled = True
    End If
    If Difficility = 1 And gameover = False Then Call DuanNao1(Index)
    If way = 1 Then
        MsgBox "恭喜您！获胜了！", 0 + 48, "慧木五子棋(HMGobobang)"
        gameover = True
    End If
    If way = 2 Then
        Shape1(i).Visible = True
        Shape1(i).FillColor = Color_C   '机器下棋的颜色
        Shape1(i).BorderColor = Color_P '棋子边框颜色交换
        Lab1(i).Enabled = False
        MsgBox "对不起！AI获胜！", 0 + 48, "慧木五子棋(HMGobobang)"
        gameover = True
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu SZ     '导入弹出菜单
End Sub

Private Sub Quit_Click()
    ans = MsgBox("您真的要退出吗？", 1 + 0 + 32, "慧木五子棋(HMGobobang)")
    If ans = 1 Then
        End     '退出
    ElseIf ans = 2 Then
        Cancel = True
    End If
End Sub

Private Sub Timer1_Timer()  '启动“慧木”动画1+“五子棋”动画2引入
    Sh = Sh + 1
    If Sh = 1 Then
        Label1(0).Visible = True
    ElseIf Sh = 2 Then
        Timer2.Enabled = True
        Timer2.Interval = 10
    End If
    Label1(0).Top = Label1(0).Top + 20
    If Label1(0).Top = 0 Then Timer1.Interval = 0
    Call Newp
End Sub

Private Sub Timer2_Timer()  '启动“五子棋”动画2
    If Sh = 2 Then Label1(1).Visible = True
    If Sh = 2 Then Label1(2).Visible = True
    Label1(1).Left = Label1(1).Left + 10
    Label1(2).Left = Label1(2).Left - 10
    If Label1(2).Left = 8400 Then Timer2.Interval = 0
End Sub
Private Sub Game1_Click(Index As Integer)
    If Di <> -1 Then   '重置判断
        ans = MsgBox("您真的要开启新的对局吗？", 1 + 0 + 32, "慧木五子棋(HMGobobang)")
        If ans = 1 Then     '重置确认
            Lab1(V).BorderStyle = 0 '落子位置显示清除
            N = 1   '外部启动标志
            Call Chongzhi(Index)
            combol1.ListIndex = Index
            Difficility = Index + 1 '难度设置
            gameover = False: way = 0
        ElseIf ans = 2 Then    '重置取消
            Cancel = True
        End If
    Else
        combol1.ListIndex = Index
        Difficility = Index + 1 '难度设置
        gameover = False: way = 0
    End If
End Sub
Private Sub Newz()   '重绘落子
    Shape1(0).Move 40, 40   '绘子
    Lab1(0).Move 40, 40
    For i = 1 To 224
        If X = 0 Then
            Load Shape1(i)
        End If
        Select Case i   '落子排布
            Case 1 To 14
                Shape1(i).Move 40 + 494 * i, 40
                Lab1(i).Move 40 + 494 * i, 40
            Case 15 To 29
                Shape1(i).Move 40 + 494 * (i - 15), 534
                Lab1(i).Move 40 + 494 * (i - 15), 534
            Case 30 To 44
                Shape1(i).Move 40 + 494 * (i - 30), 1028
                Lab1(i).Move 40 + 494 * (i - 30), 1028
            Case 45 To 59
                Shape1(i).Move 40 + 494 * (i - 45), 1522
                Lab1(i).Move 40 + 494 * (i - 45), 1522
            Case 60 To 74
                Shape1(i).Move 40 + 494 * (i - 60), 2016
                Lab1(i).Move 40 + 494 * (i - 60), 2016
            Case 75 To 89
                Shape1(i).Move 40 + 494 * (i - 75), 2510
                Lab1(i).Move 40 + 494 * (i - 75), 2510
            Case 90 To 104
                Shape1(i).Move 40 + 494 * (i - 90), 3004
                Lab1(i).Move 40 + 494 * (i - 90), 3004
            Case 105 To 119
                Shape1(i).Move 40 + 494 * (i - 105), 3498
                Lab1(i).Move 40 + 494 * (i - 105), 3498
            Case 120 To 134
                Shape1(i).Move 40 + 494 * (i - 120), 3992
                Lab1(i).Move 40 + 494 * (i - 120), 3992
            Case 135 To 149
                Shape1(i).Move 40 + 494 * (i - 135), 4486
                Lab1(i).Move 40 + 494 * (i - 135), 4486
            Case 150 To 164
                Shape1(i).Move 40 + 494 * (i - 150), 4980
                Lab1(i).Move 40 + 494 * (i - 150), 4980
            Case 165 To 179
                Shape1(i).Move 40 + 494 * (i - 165), 5474
                Lab1(i).Move 40 + 494 * (i - 165), 5474
            Case 180 To 194
                Shape1(i).Move 40 + 494 * (i - 180), 5968
                Lab1(i).Move 40 + 494 * (i - 180), 5968
            Case 195 To 209
                Shape1(i).Move 40 + 494 * (i - 195), 6462
                Lab1(i).Move 40 + 494 * (i - 195), 6462
            Case 210 To 224
                Shape1(i).Move 40 + 494 * (i - 210), 6956
                Lab1(i).Move 40 + 494 * (i - 210), 6956
        End Select
        Shape1(i).BackStyle = 1
        Shape1(i).FillColor = &H8000&
        Lab1(i).MousePointer = vbCrosshair  '空格激发鼠标变化，提示允许落子
    Next i
    X = 1
    Shape1(0).FillColor = &H8000&
End Sub
Private Sub Clsp()   '清除落子解放空格
    Call Newz   '确保每个落子都得到生成加载
    For i = 0 To 224
        Shape1(i).Visible = False   '落子清除
        Lab1(i).Enabled = True      '空格解放
    Next i
    Score_P = 1
    Score_C = 1     '重置积分
End Sub
Private Sub Newp()   '重绘棋盘
    Picture1.Scale (120, 120)-(7320, 7320)     '绘棋盘（始）
    Picture1.BackColor = &H80000002
    For i = 1 To 13
        Picture1.Line (260 + 494 * i, 240)-(260 + 494 * i, 7176)    '棋盘纵列
        Picture1.Line (260, 260 + 494 * i)-(7176, 260 + 494 * i)    '棋盘横列
    Next i
    For i = 0 To 2
        Picture1.Circle (1742 + 1976 * i, 1742 + 1976 * i), 45  '左上到右下三个标记点
    Next i
    Picture1.Circle (5694, 1742), 45                 '右上标记点
    Picture1.Circle (1742, 5694), 45                 '左下标记点
    Picture1.DrawWidth = 3
    Picture1.Line (260, 260)-(7176, 260)    '上框
    Picture1.Line (260, 260)-(260, 7176)    '左框
    Picture1.Line (7176, 260)-(7176, 7176)  '右框
    Picture1.Line (260, 7176)-(7176, 7176)  '下框
    Picture1.DrawWidth = 1 '绘棋盘（末）
End Sub

Private Sub Timer3_Timer()  '进度条运行
    If JDT.Width < 7315 Then
        JDT.Width = JDT.Width + 170
        TS.Caption = "Loading..."
        Picture1.Enabled = False
    Else
        Timer3.Enabled = False
        TS.Caption = ""
        Picture1.Enabled = True
    End If
End Sub
Private Sub JD()     '进度条初始
    JDT.Width = 15
    Timer3.Enabled = True
    If chess < 2 Then
        Command3.Enabled = False
    Else
        Command3.Enabled = True
    End If
End Sub
Private Sub Chongzhi(Index As Integer)    '外重置，稳定特殊结构
    Sh = -1     '动画起始
    Label1(0).Top = -400
    Label1(1).Left = 7560
    Label1(2).Left = 8780
    Timer1.Enabled = True
    Timer1.Interval = 30
    Di = Index
    Call Clsp
    Call Newz
    Call JD
    Call ZX
    Call FB
End Sub

Private Sub ZX()     '坐下按钮的生成
    Command1.Visible = True
    Command2.Visible = True
End Sub
Private Sub FB()     '封闭所有棋子
    For i = 0 To 224
        Lab1(i).Enabled = False
    Next i
End Sub
Private Sub JH()     '激活所有棋子
    Sh = 0     '动画起始
    Label1(0).Top = -40
    Label1(1).Left = 7560
    Label1(2).Left = 8780
    Timer1.Enabled = True
    Timer1.Interval = 30
    Di = Index
    Call Clsp
    Call Newz
End Sub
Private Sub P_wuzi(Score_P As Long, Score_C As Long, i As Integer)   '“五子”判断与积分
    '人的判定开始
    For m = 0 To 14     '行数
        If i >= 0 + 15 * m And i <= 10 + 15 * m Then
            '横单五
            If Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Shape1(i + 4).FillColor And Shape1(i + 4).FillColor = Color_P Then Score_P = Score_P + 5000000: gameover = True: way = 1
        End If
    Next m
    For m = 0 To 10     '列数
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖单五
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 60).FillColor = Shape1(i + 45).FillColor And Shape1(i + 60).FillColor = Color_P Then Score_P = Score_P + 5000000: gameover = True: way = 1
        End If
        If i >= 60 + 15 * m And i <= 70 + 15 * m Then
            '左下到右上单五
            If Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Shape1(i - 56).FillColor And Shape1(i - 56).FillColor = Color_P Then Score_P = Score_P + 5000000: gameover = True: way = 1
        End If
        If i >= 0 + 15 * m And i <= 10 + 15 * m Then
            '左上到右下单五
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Shape1(i + 64).FillColor And Shape1(i + 64).FillColor = Color_P Then Score_P = Score_P + 5000000: gameover = True: way = 1
        End If
    Next m
    '人的判定结束，机器的判定开始
    For m = 0 To 14     '行数
        If i >= 0 + 15 * m And i <= 10 + 15 * m Then
            '横单五
            If Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Shape1(i + 4).FillColor And Shape1(i + 4).FillColor = Color_C Then Score_C = Score_C + 5000000: gameover = True: way = 2
        End If
    Next m
    For m = 0 To 10     '列数
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖单五
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 60).FillColor = Shape1(i + 45).FillColor And Shape1(i + 60).FillColor = Color_C Then Score_C = Score_C + 5000000: gameover = True: way = 2
        End If
        If i >= 60 + 15 * m And i <= 70 + 15 * m Then
            '左下到右上单五
            If Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Shape1(i - 56).FillColor And Shape1(i - 56).FillColor = Color_C Then Score_C = Score_C + 5000000: gameover = True: way = 2
        End If
        If i >= 0 + 15 * m And i <= 10 + 15 * m Then
            '左上到右下单五
           If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Shape1(i + 64).FillColor And Shape1(i + 64).FillColor = Color_C Then Score_C = Score_C + 5000000: gameover = True: way = 2
        End If
    Next m
    '机器的判定结束
End Sub
Private Sub P_dansi(Score_P As Long, Score_C As Long, i As Integer)     '“单四”判断与积分
    '人的判定开始
    For m = 0 To 14     '行数
        If i = 0 + 15 * m Then
            '横单四左封(盘封)
            If Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_P And Shape1(i + 4).FillColor = &H8000& Then Score_P = Score_P + 5000
        End If
        If i >= 1 + 15 * m And i <= 10 + 15 * m Then
            '横单四左封(棋封)
            If Shape1(i - 1).FillColor = Color_C And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_P And Shape1(i + 4).FillColor = &H8000& Then Score_P = Score_P + 5000
            '横单四右封(棋封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_P And Shape1(i + 4).FillColor = Color_C Then Score_P = Score_P + 5000
        End If
        If i = 11 + 15 * m Then
            '横单四右封(盘封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_P Then Score_P = Score_P + 5000
        End If
    Next m
    For m = 0 To 14     '列数
        If i = m Then
            '竖单四上封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_P And Shape1(i + 60).FillColor = &H8000& Then Score_P = Score_P + 5000
        End If
        If i = 165 + m Then
            '竖单四下封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_P And Shape1(i - 15).FillColor = &H8000& Then Score_P = Score_P + 5000
        End If
    Next m
    For m = 1 To 11
        If i = 209 + m Then
            '左下到右上单四左下封(下盘封)
            If Shape1(i - 56).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_P Then Score_P = Score_P + 5000
        End If
        If i = 41 + 15 * m Then
            '左下到右上单四右上封(右盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_P Then Score_P = Score_P + 5000
        End If
        If i = -1 + m Then
            '左上到右下单四左上封(上盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_P And Shape1(i + 64).FillColor = &H8000& Then Score_P = Score_P + 5000
        End If
        If i = 165 + m Then
            '左上到右下单四右下封(下盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_P And Shape1(i - 16).FillColor = &H8000& Then Score_P = Score_P + 5000
        End If
    Next m
    For m = 1 To 10     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖单四上封(棋封)
            If Shape1(i - 15).FillColor = Color_C And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_P And Shape1(i + 60).FillColor = &H8000& Then Score_P = Score_P + 5000
            '竖单四下封(棋封)
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_P And Shape1(i + 60).FillColor = Color_C Then Score_P = Score_P + 5000
        End If
        If i >= 46 + 15 * m And i <= 55 + 15 * m Then
            '左下到右上单四左下封(棋封)
            If Shape1(i + 14).FillColor = Color_C And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_P And Shape1(i - 56).FillColor = &H8000& Then Score_P = Score_P + 5000
            '左下到右上单四右上封(棋封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_P And Shape1(i - 56).FillColor = Color_C Then Score_P = Score_P + 5000
        End If
        If i = 45 + m Then
            '左下到右上单四右上封(上盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_P Then Score_P = Score_P + 5000
        End If
        If i = 45 + 15 * m Then
            '左下到右上单四左下封(左盘封)
            If Shape1(i - 56).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_P Then Score_P = Score_P + 5000
        End If
        If i >= 1 + 15 * m And i <= 10 + 15 * m Then
            '左上到右下单四左上封(棋封)
            If Shape1(i - 16).FillColor = Color_C And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_P And Shape1(i + 64).FillColor = &H8000& Then Score_P = Score_P + 5000
            '左上到右下单四右下封(棋封)
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_P And Shape1(i + 64).FillColor = Color_C Then Score_P = Score_P + 5000
        End If
        If i = 0 + 15 * m Then
            '左上到右下单四左上封(左盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_P And Shape1(i + 64).FillColor = &H8000& Then Score_P = Score_P + 5000
        End If
        If i = 11 + 15 * m Then
            '左上到右下单四右下封(右盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_P And Shape1(i - 16).FillColor = &H8000& Then Score_P = Score_P + 5000
        End If
    Next m
    '人的判定结束，机器的判定开始
    For m = 0 To 14     '行数
        If i = 0 + 15 * m Then
            '横单四左封(盘封)
            If Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_C And Shape1(i + 4).FillColor = &H8000& Then Score_C = Score_C + 5000
        End If
        If i >= 1 + 15 * m And i <= 10 + 15 * m Then
            '横单四左封(棋封)
            If Shape1(i - 1).FillColor = Color_P And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_C And Shape1(i + 4).FillColor = &H8000& Then Score_C = Score_C + 5000
            '横单四右封(棋封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_C And Shape1(i + 4).FillColor = Color_P Then Score_C = Score_C + 5000
        End If
        If i = 11 + 15 * m Then
            '横单四右封(盘封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_C Then Score_C = Score_C + 5000
        End If
    Next m
    For m = 0 To 14     '列数
        If i = m Then
            '竖单四上封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_C And Shape1(i + 60).FillColor = &H8000& Then Score_C = Score_C + 5000
        End If
        If i = 165 + m Then
            '竖单四下封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_C And Shape1(i - 15).FillColor = &H8000& Then Score_C = Score_C + 5000
        End If
    Next m
    For m = 1 To 11
        If i = 209 + m Then
            '左下到右上单四左下封(下盘封)
            If Shape1(i - 56).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_C Then Score_C = Score_C + 5000
        End If
        If i = 41 + 15 * m Then
            '左下到右上单四右上封(右盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_C Then Score_C = Score_C + 5000
        End If
        If i = -1 + m Then
            '左上到右下单四左上封(上盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_C And Shape1(i + 64).FillColor = &H8000& Then Score_C = Score_C + 5000
        End If
        If i = 165 + m Then
            '左上到右下单四右下封(下盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_C And Shape1(i - 16).FillColor = &H8000& Then Score_C = Score_C + 5000
        End If
    Next m
    For m = 1 To 10     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖单四上封(棋封)
            If Shape1(i - 15).FillColor = Color_P And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_C And Shape1(i + 60).FillColor = &H8000& Then Score_C = Score_C + 5000
            '竖单四下封(棋封)
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_C And Shape1(i + 60).FillColor = Color_P Then Score_C = Score_C + 5000
        End If
        If i >= 46 + 15 * m And i <= 55 + 15 * m Then
            '左下到右上单四左下封(棋封)
            If Shape1(i + 14).FillColor = Color_P And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_C And Shape1(i - 56).FillColor = &H8000& Then Score_C = Score_C + 5000
            '左下到右上单四右上封(棋封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_C And Shape1(i - 56).FillColor = Color_P Then Score_C = Score_C + 5000
        End If
        If i = 45 + m Then
            '左下到右上单四右上封(上盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_C Then Score_C = Score_C + 5000
        End If
        If i = 45 + 15 * m Then
            '左下到右上单四左下封(左盘封)
            If Shape1(i - 56).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_C Then Score_C = Score_C + 5000
        End If
        If i >= 1 + 15 * m And i <= 10 + 15 * m Then
            '左上到右下单四左上封(棋封)
            If Shape1(i - 16).FillColor = Color_P And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_C And Shape1(i + 64).FillColor = &H8000& Then Score_C = Score_C + 5000
            '左上到右下单四右下封(棋封)
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_C And Shape1(i + 64).FillColor = Color_P Then Score_C = Score_C + 5000
        End If
        If i = 0 + 15 * m Then
            '左上到右下单四左上封(左盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_C And Shape1(i + 64).FillColor = &H8000& Then Score_C = Score_C + 5000
        End If
        If i = 11 + 15 * m Then
            '左上到右下单四右下封(右盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_C And Shape1(i - 16).FillColor = &H8000& Then Score_C = Score_C + 5000
        End If
    Next m
    '机器的判定结束
End Sub
Private Sub P_gusi(Score_P As Long, Score_C As Long, i As Integer)     '“孤四”判断与积分
    '人的判定开始
    For m = 0 To 14     '行数
        If i >= 1 + 15 * m And i <= 10 + 15 * m Then
            '横孤四
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_P And Shape1(i + 4).FillColor = &H8000& Then Score_P = Score_P + 400000
        End If
    Next m
    For m = 1 To 10     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖孤四
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_P And Shape1(i + 60).FillColor = &H8000& Then Score_P = Score_P + 400000
        End If
        If i >= 46 + 15 * m And i <= 55 + 15 * m Then
            '左下到右上孤四
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_P And Shape1(i - 56).FillColor = &H8000& Then Score_P = Score_P + 400000
        End If
        If i >= 1 + 15 * m And i <= 10 + 15 * m Then
            '左上到右下孤四
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_P And Shape1(i + 64).FillColor = &H8000& Then Score_P = Score_P + 400000
        End If
    Next m
    '人的判定结束，机器的判定开始
    For m = 0 To 14     '行数
        If i >= 1 + 15 * m And i <= 10 + 15 * m Then
            '横孤四
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Shape1(i + 3).FillColor And Shape1(i + 3).FillColor = Color_C And Shape1(i + 4).FillColor = &H8000& Then Score_C = Score_C + 400000
        End If
    Next m
    For m = 1 To 10     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖孤四
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Shape1(i + 45).FillColor And Shape1(i + 45).FillColor = Color_C And Shape1(i + 60).FillColor = &H8000& Then Score_C = Score_C + 400000
        End If
        If i >= 46 + 15 * m And i <= 55 + 15 * m Then
            '左下到右上孤四
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Shape1(i - 42).FillColor And Shape1(i - 42).FillColor = Color_C And Shape1(i - 56).FillColor = &H8000& Then Score_C = Score_C + 400000
        End If
        If i >= 1 + 15 * m And i <= 10 + 15 * m Then
            '左上到右下孤四
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Shape1(i + 48).FillColor And Shape1(i + 48).FillColor = Color_C And Shape1(i + 64).FillColor = &H8000& Then Score_C = Score_C + 400000
        End If
    Next m
    '机器的判定结束
End Sub
Private Sub P_dansan(Score_P As Long, Score_C As Long, i As Integer)     '“单三”判断与积分
    '人的判定开始
    For m = 0 To 14     '行数
        If i = 0 + 15 * m Then
            '横单三左封(盘封)
            If Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_P And Shape1(i + 3).FillColor = &H8000& Then Score_P = Score_P + 1000
        End If
        If i >= 1 + 15 * m And i <= 11 + 15 * m Then
            '横单三左封(棋封)
            If Shape1(i - 1).FillColor = Color_C And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_P And Shape1(i + 3).FillColor = &H8000& Then Score_P = Score_P + 1000
            '横单三右封(棋封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_P And Shape1(i + 3).FillColor = Color_C Then Score_P = Score_P + 1000
        End If
        If i = 12 + 15 * m Then
            '横单三右封(盘封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_P Then Score_P = Score_P + 1000
        End If
    Next m
    For m = 0 To 14     '列数
        If i = m Then
            '竖单三上封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_P And Shape1(i + 45).FillColor = &H8000& Then Score_P = Score_P + 1000
        End If
        If i = 180 + m Then
            '竖单三下封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_P And Shape1(i - 15).FillColor = &H8000& Then Score_P = Score_P + 1000
        End If
    Next m
    For m = 1 To 12
        If i = 209 + m Then
            '左下到右上单三左下封(下盘封)
            If Shape1(i - 42).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_P Then Score_P = Score_P + 1000
        End If
        If i = 27 + 15 * m Then
            '左下到右上单三右上封(右盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_P Then Score_P = Score_P + 1000
        End If
        If i = -1 + m Then
            '左上到右下单三左上封(上盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_P And Shape1(i + 48).FillColor = &H8000& Then Score_P = Score_P + 1000
        End If
        If i = 180 + m Then
            '左上到右下单三右下封(下盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_P And Shape1(i - 16).FillColor = &H8000& Then Score_P = Score_P + 5000
        End If
    Next m
    For m = 1 To 11     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖单三上封(棋封)
            If Shape1(i - 15).FillColor = Color_C And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_P And Shape1(i + 45).FillColor = &H8000& Then Score_P = Score_P + 1000
            '竖单三下封(棋封)
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_P And Shape1(i + 45).FillColor = Color_C Then Score_P = Score_P + 1000
        End If
        If i >= 31 + 15 * m And i <= 41 + 15 * m Then
            '左下到右上单三左下封(棋封)
            If Shape1(i + 14).FillColor = Color_C And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_P And Shape1(i - 42).FillColor = &H8000& Then Score_P = Score_P + 1000
            '左下到右上单三右上封(棋封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_P And Shape1(i - 42).FillColor = Color_C Then Score_P = Score_P + 1000
        End If
        If i = 30 + m Then
            '左下到右上单三右上封(上盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_P Then Score_P = Score_P + 1000
        End If
        If i = 30 + 15 * m Then
            '左下到右上单三左下封(左盘封)
            If Shape1(i - 42).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_P Then Score_P = Score_P + 1000
        End If
        If i >= 1 + 15 * m And i <= 11 + 15 * m Then
            '左上到右下单三左上封(棋封)
            If Shape1(i - 16).FillColor = Color_C And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_P And Shape1(i + 48).FillColor = &H8000& Then Score_P = Score_P + 1000
            '左上到右下单三右下封(棋封)
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_P And Shape1(i + 48).FillColor = Color_C Then Score_P = Score_P + 1000
        End If
        If i = 0 + 15 * m Then
            '左上到右下单三左上封(左盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_P And Shape1(i + 48).FillColor = &H8000& Then Score_P = Score_P + 1000
        End If
        If i = 12 + 15 * m Then
            '左上到右下单三右下封(右盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_P And Shape1(i - 16).FillColor = &H8000& Then Score_P = Score_P + 1000
        End If
    Next m
    '人的判定结束，机器的判定开始
    For m = 0 To 14     '行数
        If i = 0 + 15 * m Then
            '横单三左封(盘封)
            If Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_C And Shape1(i + 3).FillColor = &H8000& Then Score_C = Score_C + 1000
        End If
        If i >= 1 + 15 * m And i <= 11 + 15 * m Then
            '横单三左封(棋封)
            If Shape1(i - 1).FillColor = Color_P And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_C And Shape1(i + 3).FillColor = &H8000& Then Score_C = Score_C + 1000
            '横单三右封(棋封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_C And Shape1(i + 3).FillColor = Color_P Then Score_C = Score_C + 1000
        End If
        If i = 12 + 15 * m Then
            '横单三右封(盘封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_C Then Score_C = Score_C + 1000
        End If
    Next m
    For m = 0 To 14     '列数
        If i = m Then
            '竖单三上封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_C And Shape1(i + 45).FillColor = &H8000& Then Score_C = Score_C + 1000
        End If
        If i = 180 + m Then
            '竖单三下封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_C And Shape1(i - 15).FillColor = &H8000& Then Score_C = Score_C + 1000
        End If
    Next m
    For m = 1 To 12
        If i = 209 + m Then
            '左下到右上单三左下封(下盘封)
            If Shape1(i - 42).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_C Then Score_C = Score_C + 1000
        End If
        If i = 27 + 15 * m Then
            '左下到右上单三右上封(右盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_C Then Score_C = Score_C + 1000
        End If
        If i = -1 + m Then
            '左上到右下单三左上封(上盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_C And Shape1(i + 48).FillColor = &H8000& Then Score_C = Score_C + 1000
        End If
        If i = 180 + m Then
            '左上到右下单三右下封(下盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_C And Shape1(i - 16).FillColor = &H8000& Then Score_C = Score_C + 5000
        End If
    Next m
    For m = 1 To 11     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖单三上封(棋封)
            If Shape1(i - 15).FillColor = Color_P And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_C And Shape1(i + 45).FillColor = &H8000& Then Score_C = Score_C + 1000
            '竖单三下封(棋封)
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_C And Shape1(i + 45).FillColor = Color_P Then Score_C = Score_C + 1000
        End If
        If i >= 31 + 15 * m And i <= 41 + 15 * m Then
            '左下到右上单三左下封(棋封)
            If Shape1(i + 14).FillColor = Color_P And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_C And Shape1(i - 42).FillColor = &H8000& Then Score_C = Score_C + 1000
            '左下到右上单三右上封(棋封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_C And Shape1(i - 42).FillColor = Color_P Then Score_C = Score_C + 1000
        End If
        If i = 30 + m Then
            '左下到右上单三右上封(上盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_C Then Score_C = Score_C + 1000
        End If
        If i = 30 + 15 * m Then
            '左下到右上单三左下封(左盘封)
            If Shape1(i - 42).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_C Then Score_C = Score_C + 1000
        End If
        If i >= 1 + 15 * m And i <= 11 + 15 * m Then
            '左上到右下单三左上封(棋封)
            If Shape1(i - 16).FillColor = Color_P And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_C And Shape1(i + 48).FillColor = &H8000& Then Score_C = Score_C + 1000
            '左上到右下单三右下封(棋封)
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_C And Shape1(i + 48).FillColor = Color_P Then Score_C = Score_C + 1000
        End If
        If i = 0 + 15 * m Then
            '左上到右下单三左上封(左盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_C And Shape1(i + 48).FillColor = &H8000& Then Score_C = Score_C + 1000
        End If
        If i = 12 + 15 * m Then
            '左上到右下单三右下封(右盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_C And Shape1(i - 16).FillColor = &H8000& Then Score_C = Score_C + 1000
        End If
    Next m
    '机器的判定结束
End Sub
Private Sub P_gusan(Score_P As Long, Score_C As Long, i As Integer)     '“孤三”判断与积分
    '人的判定开始
    For m = 0 To 14     '行数
        If i >= 1 + 15 * m And i <= 11 + 15 * m Then
            '横孤三
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_P And Shape1(i + 3).FillColor = &H8000& Then Score_P = Score_P + 70000
        End If
    Next m
    For m = 1 To 11     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖孤三
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_P And Shape1(i + 45).FillColor = &H8000& Then Score_P = Score_P + 70000
        End If
        If i >= 31 + 15 * m And i <= 41 + 15 * m Then
            '左下到右上孤三
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_P And Shape1(i - 42).FillColor = &H8000& Then Score_P = Score_P + 70000
        End If
        If i >= 1 + 15 * m And i <= 11 + 15 * m Then
            '左上到右下孤三
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_P And Shape1(i + 48).FillColor = &H8000& Then Score_P = Score_P + 70000
        End If
    Next m
    '人的判定结束，机器的判定开始
    For m = 0 To 14     '行数
        If i >= 1 + 15 * m And i <= 11 + 15 * m Then
            '横孤三
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i + 2).FillColor And Shape1(i + 2).FillColor = Color_C And Shape1(i + 3).FillColor = &H8000& Then Score_C = Score_C + 70000
        End If
    Next m
    For m = 1 To 11     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖孤三
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 30).FillColor And Shape1(i + 30).FillColor = Color_C And Shape1(i + 45).FillColor = &H8000& Then Score_C = Score_C + 70000
        End If
        If i >= 31 + 15 * m And i <= 41 + 15 * m Then
            '左下到右上孤三
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i - 28).FillColor And Shape1(i - 28).FillColor = Color_C And Shape1(i - 42).FillColor = &H8000& Then Score_C = Score_C + 70000
        End If
        If i >= 1 + 15 * m And i <= 11 + 15 * m Then
            '左上到右下孤三
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Shape1(i + 32).FillColor And Shape1(i + 32).FillColor = Color_C And Shape1(i + 48).FillColor = &H8000& Then Score_C = Score_C + 70000
        End If
    Next m
    '机器的判定结束
End Sub

Private Sub P_daner(Score_P As Long, Score_C As Long, i As Integer)     '“单二”判断与积分
    '人的判定开始
    For m = 0 To 14     '行数
        If i = 0 + 15 * m Then
            '横单二左封(盘封)
            If Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_P And Shape1(i + 2).FillColor = &H8000& Then Score_P = Score_P + 200
        End If
        If i >= 1 + 15 * m And i <= 12 + 15 * m Then
            '横单二左封(棋封)
            If Shape1(i - 1).FillColor = Color_C And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_P And Shape1(i + 2).FillColor = &H8000& Then Score_P = Score_P + 200
            '横单二右封(棋封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_P And Shape1(i + 2).FillColor = Color_C Then Score_P = Score_P + 200
        End If
        If i = 13 + 15 * m Then
            '横单二右封(盘封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_P Then Score_P = Score_P + 200
        End If
    Next m
    For m = 0 To 14     '列数
        If i = m Then
            '竖单二上封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_P And Shape1(i + 30).FillColor = &H8000& Then Score_P = Score_P + 200
        End If
        If i = 195 + m Then
            '竖单二下封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_P And Shape1(i - 15).FillColor = &H8000& Then Score_P = Score_P + 200
        End If
    Next m
    For m = 1 To 13
        If i = 209 + m Then
            '左下到右上单二左下封(下盘封)
            If Shape1(i - 28).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_P Then Score_P = Score_P + 200
        End If
        If i = 13 + 15 * m Then
            '左下到右上单二右上封(右盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_P Then Score_P = Score_P + 200
        End If
        If i = -1 + m Then
            '左上到右下单二左上封(上盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_P And Shape1(i + 32).FillColor = &H8000& Then Score_P = Score_P + 200
        End If
        If i = 195 + m Then
            '左上到右下单二右下封(下盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_P And Shape1(i - 16).FillColor = &H8000& Then Score_P = Score_P + 200
        End If
    Next m
    For m = 1 To 12     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖单二上封(棋封)
            If Shape1(i - 15).FillColor = Color_C And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_P And Shape1(i + 30).FillColor = &H8000& Then Score_P = Score_P + 200
            '竖单二下封(棋封)
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_P And Shape1(i + 30).FillColor = Color_C Then Score_P = Score_P + 200
        End If
        If i >= 16 + 15 * m And i <= 27 + 15 * m Then
            '左下到右上单二左下封(棋封)
            If Shape1(i + 14).FillColor = Color_C And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_P And Shape1(i - 28).FillColor = &H8000& Then Score_P = Score_P + 200
            '左下到右上单二右上封(棋封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_P And Shape1(i - 28).FillColor = Color_C Then Score_P = Score_P + 200
        End If
        If i = 15 + m Then
            '左下到右上单二右上封(上盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_P Then Score_P = Score_P + 200
        End If
        If i = 15 + 15 * m Then
            '左下到右上单二左下封(左盘封)
            If Shape1(i - 28).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_P Then Score_P = Score_P + 200
        End If
        If i >= 1 + 15 * m And i <= 12 + 15 * m Then
            '左上到右下单二左上封(棋封)
            If Shape1(i - 16).FillColor = Color_C And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_P And Shape1(i + 32).FillColor = &H8000& Then Score_P = Score_P + 200
            '左上到右下单二右下封(棋封)
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_P And Shape1(i + 32).FillColor = Color_C Then Score_P = Score_P + 200
        End If
        If i = 0 + 15 * m Then
            '左上到右下单二左上封(左盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_P And Shape1(i + 32).FillColor = &H8000& Then Score_P = Score_P + 200
        End If
        If i = 13 + 15 * m Then
            '左上到右下单二右下封(右盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_P And Shape1(i - 16).FillColor = &H8000& Then Score_P = Score_P + 200
        End If
    Next m
    '人的判定结束，机器的判定开始
    For m = 0 To 14     '行数
        If i = 0 + 15 * m Then
            '横单二左封(盘封)
            If Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_C And Shape1(i + 2).FillColor = &H8000& Then Score_C = Score_C + 200
        End If
        If i >= 1 + 15 * m And i <= 12 + 15 * m Then
            '横单二左封(棋封)
            If Shape1(i - 1).FillColor = Color_P And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_C And Shape1(i + 2).FillColor = &H8000& Then Score_C = Score_C + 200
            '横单二右封(棋封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_C And Shape1(i + 2).FillColor = Color_P Then Score_C = Score_C + 200
        End If
        If i = 13 + 15 * m Then
            '横单二右封(盘封)
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_C Then Score_C = Score_C + 200
        End If
    Next m
    For m = 0 To 14     '列数
        If i = m Then
            '竖单二上封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_C And Shape1(i + 30).FillColor = &H8000& Then Score_C = Score_C + 200
        End If
        If i = 195 + m Then
            '竖单二下封(盘封)
            If Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_C And Shape1(i - 15).FillColor = &H8000& Then Score_C = Score_C + 200
        End If
    Next m
    For m = 1 To 13
        If i = 209 + m Then
            '左下到右上单二左下封(下盘封)
            If Shape1(i - 28).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_C Then Score_C = Score_C + 200
        End If
        If i = 13 + 15 * m Then
            '左下到右上单二右上封(右盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_C Then Score_C = Score_C + 200
        End If
        If i = -1 + m Then
            '左上到右下单二左上封(上盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_C And Shape1(i + 32).FillColor = &H8000& Then Score_C = Score_C + 200
        End If
        If i = 195 + m Then
            '左上到右下单二右下封(下盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_C And Shape1(i - 16).FillColor = &H8000& Then Score_C = Score_C + 200
        End If
    Next m
    For m = 1 To 12     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖单二上封(棋封)
            If Shape1(i - 15).FillColor = Color_P And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_C And Shape1(i + 30).FillColor = &H8000& Then Score_C = Score_C + 200
            '竖单二下封(棋封)
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_C And Shape1(i + 30).FillColor = Color_P Then Score_C = Score_C + 200
        End If
        If i >= 16 + 15 * m And i <= 27 + 15 * m Then
            '左下到右上单二左下封(棋封)
            If Shape1(i + 14).FillColor = Color_P And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_C And Shape1(i - 28).FillColor = &H8000& Then Score_C = Score_C + 200
            '左下到右上单二右上封(棋封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_C And Shape1(i - 28).FillColor = Color_P Then Score_C = Score_C + 200
        End If
        If i = 15 + m Then
            '左下到右上单二右上封(上盘封)
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_C Then Score_C = Score_C + 200
        End If
        If i = 15 + 15 * m Then
            '左下到右上单二左下封(左盘封)
            If Shape1(i - 28).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_C Then Score_C = Score_C + 200
        End If
        If i >= 1 + 15 * m And i <= 12 + 15 * m Then
            '左上到右下单二左上封(棋封)
            If Shape1(i - 16).FillColor = Color_P And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_C And Shape1(i + 32).FillColor = &H8000& Then Score_C = Score_C + 200
            '左上到右下单二右下封(棋封)
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_C And Shape1(i + 32).FillColor = Color_P Then Score_C = Score_C + 200
        End If
        If i = 0 + 15 * m Then
            '左上到右下单二左上封(左盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_C And Shape1(i + 32).FillColor = &H8000& Then Score_C = Score_C + 200
        End If
        If i = 13 + 15 * m Then
            '左上到右下单二右下封(右盘封)
            If Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_C And Shape1(i - 16).FillColor = &H8000& Then Score_C = Score_C + 200
        End If
    Next m
    '机器的判定结束
End Sub

Private Sub P_guer(Score_P As Long, Score_C As Long, i As Integer)     '“孤二”判断与积分
 '人的判定开始
    For m = 0 To 14     '行数
        If i >= 1 + 15 * m And i <= 12 + 15 * m Then
            '横孤二
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_P And Shape1(i + 2).FillColor = &H8000& Then Score_P = Score_P + 250
        End If
    Next m
    For m = 1 To 12     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖孤二
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_P And Shape1(i + 30).FillColor = &H8000& Then Score_P = Score_P + 250
        End If
        If i >= 16 + 15 * m And i <= 27 + 15 * m Then
            '左下到右上孤二
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_P And Shape1(i - 28).FillColor = &H8000& Then Score_P = Score_P + 250
        End If
        If i >= 1 + 15 * m And i <= 12 + 15 * m Then
            '左上到右下孤二
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_P And Shape1(i + 32).FillColor = &H8000& Then Score_P = Score_P + 250
        End If
    Next m
    '人的判定结束，机器的判定开始
    For m = 0 To 14     '行数
        If i >= 1 + 15 * m And i <= 12 + 15 * m Then
            '横孤二
            If Shape1(i - 1).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Color_C And Shape1(i + 2).FillColor = &H8000& Then Score_C = Score_C + 250
        End If
    Next m
    For m = 1 To 12     '列数(去双头)
        If i >= 0 + 15 * m And i <= 14 + 15 * m Then
            '竖孤二
            If Shape1(i - 15).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Color_C And Shape1(i + 30).FillColor = &H8000& Then Score_C = Score_C + 250
        End If
        If i >= 16 + 15 * m And i <= 27 + 15 * m Then
            '左下到右上孤二
            If Shape1(i + 14).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Color_C And Shape1(i - 28).FillColor = &H8000& Then Score_C = Score_C + 250
        End If
        If i >= 1 + 15 * m And i <= 12 + 15 * m Then
            '左上到右下孤二
            If Shape1(i - 16).FillColor = &H8000& And Shape1(i).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = Color_C And Shape1(i + 32).FillColor = &H8000& Then Score_C = Score_C + 250
        End If
    Next m
    '机器的判定结束
End Sub

Private Sub P_yizi(Score_P As Long, Score_C As Long, i As Integer)     '“一子”判断与积分
    For m = 0 To 12
        If i >= 16 + 15 * m And i <= 28 + 15 * m Then
            '人的单子判定
            If Shape1(i - 1).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i - 16).FillColor And Shape1(i - 16).FillColor = Shape1(i - 15).FillColor And Shape1(i - 15).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i + 14).FillColor And Shape1(i + 14).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = &H8000& And Shape1(i).FillColor = Color_P Then Score_P = Score_P + 10
            '机器的单子判定
            If Shape1(i - 1).FillColor = Shape1(i + 1).FillColor And Shape1(i + 1).FillColor = Shape1(i - 16).FillColor And Shape1(i - 16).FillColor = Shape1(i - 15).FillColor And Shape1(i - 15).FillColor = Shape1(i - 14).FillColor And Shape1(i - 14).FillColor = Shape1(i + 14).FillColor And Shape1(i + 14).FillColor = Shape1(i + 15).FillColor And Shape1(i + 15).FillColor = Shape1(i + 16).FillColor And Shape1(i + 16).FillColor = &H8000& And Shape1(i).FillColor = Color_C Then Score_C = Score_C + 10
        End If
    Next m
End Sub

Private Sub Score(Score_P As Long, Score_C As Long)
    Dim a As Integer
    For a = 0 To 224
        Call P_wuzi(Score_P, Score_C, a)
        Call P_gusi(Score_P, Score_C, a)
        Call P_dansi(Score_P, Score_C, a)
        Call P_gusan(Score_P, Score_C, a)
        Call P_dansan(Score_P, Score_C, a)
        Call P_daner(Score_P, Score_C, a)
        Call P_guer(Score_P, Score_C, a)
        Call P_yizi(Score_P, Score_C, a)
    Next a
End Sub
Private Sub DuanNao1(i As Integer)
    Dim m As Integer
    Dim B As Integer
    Dim Ai As Integer
    Ai = 0
    If i = 0 Then   '0起特殊情况
        B = Int(3 * Rnd)
        If B = 0 And Lab1(0).Enabled = True Then
            Call Ai1(1): Ai = 1
        ElseIf B = 1 And Lab1(1).Enabled = True Then
            Call Ai1(15): Ai = 1
        ElseIf B = 2 And Lab1(2).Enabled = True Then
            Call Ai1(16): Ai = 1
        Else
            For m = 17 To 18
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 2 To 3
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 33 To 30 Step -1
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 48 To 45 Step -1
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
        End If
    ElseIf i = 210 Then '210起特殊情况
        B = Int(3 * Rnd)
        If B = 0 And Lab1(211).Enabled = True Then
            Call Ai1(211): Ai = 1
        ElseIf B = 1 And Lab1(195).Enabled = True Then
            Call Ai1(195): Ai = 1
        ElseIf B = 2 And Lab1(196).Enabled = True Then
            Call Ai1(196): Ai = 1
        Else
            For m = 197 To 198
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 212 To 213
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 183 To 180 Step -1
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 168 To 165 Step -1
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
        End If
    ElseIf i = 14 Then  '14起特殊情况
        B = Int(3 * Rnd)
        If B = 0 And Lab1(13).Enabled = True Then
            Call Ai1(13): Ai = 1
        ElseIf B = 1 And Lab1(29).Enabled = True Then
            Call Ai1(29): Ai = 1
        ElseIf B = 2 And Lab1(28).Enabled = True Then
            Call Ai1(28): Ai = 1
        Else
            For m = 27 To 26
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 12 To 11
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 44 To 41 Step -1
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 59 To 56 Step -1
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                    If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
        End If
    ElseIf i = 224 Then '224起特殊情况
        B = Int(3 * Rnd)
    If B = 0 And Lab1(209).Enabled = True Then
Call Ai1(209): Ai = 1
        ElseIf B = 1 And Lab1(223).Enabled = True Then
            Call Ai1(223): Ai = 1
        ElseIf B = 2 And Lab1(208).Enabled = True Then
            Call Ai1(208): Ai = 1
        Else
            For m = 207 To 206
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 222 To 221
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 194 To 191 Step -1
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
            For m = 179 To 176 Step -1
                If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                    Call Ai1(m): Ai = 1
                    Exit For
                ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                    Shape1(m).FillColor = &H8000&  '还原颜色
                End If
            Next m
        End If
    ElseIf i = 15 Or i = 30 Or i = 45 Or i = 165 Or i = 180 Or i = 195 Then '左侧边角特殊情况
        For m = i + 2 To i + 1 Step -1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 19 To i + 15 Step -1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 11 To i - 15 Step -1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 3 To i + 4
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
    ElseIf i = 29 Or i = 44 Or i = 59 Or i = 179 Or i = 194 Or i = 209 Then '右侧边角特殊情况
        For m = i - 2 To i - 1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 11 To i + 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 18 To i - 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 3 To i - 4 Step -1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
    ElseIf i = 1 Or i = 2 Or i = 3 Or i = 11 Or i = 12 Or i = 13 Then '上侧边角特殊情况
        For m = i + 30 To i + 15 Step -15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 61 To i + 1 Step -15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 59 To i - 1 Step -15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 45 To i + 60 Step 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
    ElseIf i = 211 Or i = 212 Or i = 213 Or i = 221 Or i = 222 Or i = 223 Then '下侧边角特殊情况
        For m = i - 30 To i - 15 Step 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 61 To i - 1 Step 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 59 To i + 1 Step 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 45 To i - 60 Step -15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
    ElseIf i = 60 Or i = 75 Or i = 90 Or i = 105 Or i = 120 Or i = 135 Or i = 150 Then '左侧边特殊情况
        For m = i + 2 To i + 1 Step -1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 13 To i - 15 Step -1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 17 To i + 15 Step -1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 32 To i + 30 Step -1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 30 To i - 28 Step -1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        If Lab1(i + 3).Enabled = True And Ai = 0 Then Shape1(i + 3).FillColor = Color_C '机器下棋的颜色
        Score_P = 1: Score_C = 1
        Call Score(Score_P, Score_C)
        If Ai = 0 And Lab1(i + 3).Enabled = True And Score_C >= Score_P Then
            Call Ai1(i + 3): Ai = 1
        ElseIf Lab1(i + 3).Enabled = True And Score_C < Score_P Then
            Shape1(i + 3).FillColor = &H8000& '还原颜色
        End If
    ElseIf i = 74 Or i = 89 Or i = 104 Or i = 119 Or i = 134 Or i = 149 Or i = 164 Then '右侧边特殊情况
        For m = i - 2 To i - 1 Step 1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 13 To i + 15 Step 1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 17 To i - 15 Step 1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 32 To i - 30 Step 1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 28 To i + 30 Step 1
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        If Lab1(i - 3).Enabled = True And Ai = 0 Then Shape1(i - 3).FillColor = Color_C '机器下棋的颜色
        Score_P = 1: Score_C = 1
        Call Score(Score_P, Score_C)
        If Ai = 0 And Lab1(i - 3).Enabled = True And Score_C >= Score_P Then
            Call Ai1(i - 3): Ai = 1
        ElseIf Lab1(i - 3).Enabled = True And Score_C < Score_P Then
            Shape1(i - 3).FillColor = &H8000& '还原颜色
        End If
    ElseIf i = 4 Or i = 5 Or i = 6 Or i = 7 Or i = 8 Or i = 9 Or i = 10 Then '上侧边特殊情况
        For m = i + 30 To i + 15 Step -15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 29 To i - 1 Step -15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 31 To i + 1 Step -15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 32 To i + 2 Step -15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i + 28 To i - 2 Step -15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        If Lab1(i + 45).Enabled = True And Ai = 0 Then Shape1(i + 45).FillColor = Color_C '机器下棋的颜色
        Score_P = 1: Score_C = 1
        Call Score(Score_P, Score_C)
        If Ai = 0 And Lab1(i + 45).Enabled = True And Score_C >= Score_P Then
            Call Ai1(i + 45): Ai = 1
        ElseIf Lab1(i + 45).Enabled = True And Score_C < Score_P Then
            Shape1(i + 45).FillColor = &H8000& '还原颜色
        End If
    ElseIf i = 214 Or i = 215 Or i = 216 Or i = 217 Or i = 218 Or i = 219 Or i = 220 Then '上侧边特殊情况
        For m = i - 30 To i - 15 Step 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 29 To i + 1 Step 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 31 To i - 1 Step 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 32 To i - 2 Step 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        For m = i - 28 To i + 2 Step 15
            If Lab1(m).Enabled = True And Ai = 0 Then Shape1(m).FillColor = Color_C '机器下棋的颜色
            Score_P = 1: Score_C = 1
            Call Score(Score_P, Score_C)
            If Ai = 0 And Lab1(m).Enabled = True And Score_C >= Score_P Then
                Call Ai1(m): Ai = 1
                Exit For
            ElseIf Lab1(m).Enabled = True And Score_C < Score_P Then
                Shape1(m).FillColor = &H8000&  '还原颜色
            End If
        Next m
        If Lab1(i - 45).Enabled = True And Ai = 0 Then Shape1(i - 45).FillColor = Color_C '机器下棋的颜色
        Score_P = 1: Score_C = 1
        Call Score(Score_P, Score_C)
        If Ai = 0 And Lab1(i - 45).Enabled = True And Score_C >= Score_P Then
            Call Ai1(i - 45): Ai = 1
        ElseIf Lab1(i - 45).Enabled = True And Score_C < Score_P Then
            Shape1(i - 45).FillColor = &H8000& '还原颜色
        End If
    ElseIf i = 16 Then  '16起特殊情况
        B = Int(3 * Rnd)
        If B = 0 And Lab1(i + 32).Enabled = True Then
            Call Ai1(i + 32): Ai = 1
        ElseIf B = 1 And Lab1(i + 17).Enabled = True Then
            Call Ai1(i + 17): Ai = 1
        ElseIf B = 2 And Lab1(i + 31).Enabled = True Then
            Call Ai1(i + 31): Ai = 1
        Else
            q = Int(8 * Rnd + 1)
            If Ai = 0 And q = 1 And Lab1(i - 16).Enabled = True Then
                Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 16): Ai = 1
                Else
                    Shape1(i - 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 2 And Lab1(i - 15).Enabled = True Then
                Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 15): Ai = 1
                Else
                    Shape1(i - 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 3 And Lab1(i - 14).Enabled = True Then
                Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 14): Ai = 1
                Else
                    Shape1(i - 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 4 And Lab1(i - 1).Enabled = True Then
                Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 1): Ai = 1
                Else
                    Shape1(i - 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 5 And Lab1(i + 1).Enabled = True Then
                Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 1): Ai = 1
                Else
                    Shape1(i + 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 6 And Lab1(i + 14).Enabled = True Then
                Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 14): Ai = 1
                Else
                    Shape1(i + 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 7 And Lab1(i + 15).Enabled = True Then
                Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 15): Ai = 1
                Else
                    Shape1(i + 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 8 And Lab1(i + 16).Enabled = True Then
                Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 16): Ai = 0
                Else
                    Shape1(i + 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 Then
                For q = 1 To 8
                    If q = 1 And Lab1(i - 16).Enabled = True Then
                        Call Ai1(i - 16): Ai = 1
                    ElseIf q = 2 And Lab1(i - 15).Enabled = True Then
                        Call Ai1(i - 15): Ai = 1
                    ElseIf q = 3 And Lab1(i - 14).Enabled = True Then
                        Call Ai1(i - 14): Ai = 1
                    ElseIf q = 4 And Lab1(i - 1).Enabled = True Then
                        Call Ai1(i - 1): Ai = 1
                    ElseIf q = 5 And Lab1(i + 1).Enabled = True Then
                        Call Ai1(i + 1): Ai = 1
                    ElseIf q = 6 And Lab1(i + 14).Enabled = True Then
                        Call Ai1(i + 14): Ai = 1
                    ElseIf q = 7 And Lab1(i + 15).Enabled = True Then
                        Call Ai1(i + 15): Ai = 1
                    ElseIf q = 8 And Lab1(i + 16).Enabled = True Then
                        Call Ai1(i + 16): Ai = 1
                    End If
                Next q
            End If
        End If
    ElseIf i = 28 Then  '28起特殊情况
        B = Int(3 * Rnd)
        If B = 0 And Lab1(i + 28).Enabled = True Then
            Call Ai1(i + 28): Ai = 1
        ElseIf B = 1 And Lab1(i + 28).Enabled = True Then
            Call Ai1(i + 29): Ai = 1
        ElseIf B = 2 And Lab1(i + 13).Enabled = True Then
            Call Ai1(i + 13): Ai = 1
        Else
            q = Int(8 * Rnd + 1)
            If Ai = 0 And q = 1 And Lab1(i - 16).Enabled = True Then
                Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 16): Ai = 1
                Else
                    Shape1(i - 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 2 And Lab1(i - 15).Enabled = True Then
                Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 15): Ai = 1
                Else
                    Shape1(i - 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 3 And Lab1(i - 14).Enabled = True Then
                Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 14): Ai = 1
                Else
                    Shape1(i - 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 4 And Lab1(i - 1).Enabled = True Then
                Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 1): Ai = 1
                Else
                    Shape1(i - 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 5 And Lab1(i + 1).Enabled = True Then
                Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 1): Ai = 1
                Else
                    Shape1(i + 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 6 And Lab1(i + 14).Enabled = True Then
                Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 14): Ai = 1
                Else
                    Shape1(i + 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 7 And Lab1(i + 15).Enabled = True Then
                Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 15): Ai = 1
                Else
                    Shape1(i + 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 8 And Lab1(i + 16).Enabled = True Then
                Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 16): Ai = 0
                Else
                    Shape1(i + 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 Then
                For q = 1 To 8
                    If q = 1 And Lab1(i - 16).Enabled = True Then
                        Call Ai1(i - 16): Ai = 1
                    ElseIf q = 2 And Lab1(i - 15).Enabled = True Then
                        Call Ai1(i - 15): Ai = 1
                    ElseIf q = 3 And Lab1(i - 14).Enabled = True Then
                        Call Ai1(i - 14): Ai = 1
                    ElseIf q = 4 And Lab1(i - 1).Enabled = True Then
                        Call Ai1(i - 1): Ai = 1
                    ElseIf q = 5 And Lab1(i + 1).Enabled = True Then
                        Call Ai1(i + 1): Ai = 1
                    ElseIf q = 6 And Lab1(i + 14).Enabled = True Then
                        Call Ai1(i + 14): Ai = 1
                    ElseIf q = 7 And Lab1(i + 15).Enabled = True Then
                        Call Ai1(i + 15): Ai = 1
                    ElseIf q = 8 And Lab1(i + 16).Enabled = True Then
                        Call Ai1(i + 16): Ai = 1
                    End If
                Next q
            End If
        End If
    ElseIf i = 196 Then '196起特殊情况
        B = Int(3 * Rnd)
        If B = 0 And Lab1(i - 28).Enabled = True Then
            Call Ai1(i - 28): Ai = 1
        ElseIf B = 1 And Lab1(i - 29).Enabled = True Then
            Call Ai1(i - 29): Ai = 1
        ElseIf B = 2 And Lab1(i - 13).Enabled = True Then
            Call Ai1(i - 13): Ai = 1
        Else
            q = Int(8 * Rnd + 1)
            If Ai = 0 And q = 1 And Lab1(i - 16).Enabled = True Then
                Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 16): Ai = 1
                Else
                    Shape1(i - 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 2 And Lab1(i - 15).Enabled = True Then
                Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 15): Ai = 1
                Else
                    Shape1(i - 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 3 And Lab1(i - 14).Enabled = True Then
                Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 14): Ai = 1
                Else
                    Shape1(i - 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 4 And Lab1(i - 1).Enabled = True Then
                Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 1): Ai = 1
                Else
                    Shape1(i - 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 5 And Lab1(i + 1).Enabled = True Then
                Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 1): Ai = 1
                Else
                    Shape1(i + 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 6 And Lab1(i + 14).Enabled = True Then
                Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 14): Ai = 1
                Else
                    Shape1(i + 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 7 And Lab1(i + 15).Enabled = True Then
                Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 15): Ai = 1
                Else
                    Shape1(i + 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 8 And Lab1(i + 16).Enabled = True Then
                Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 16): Ai = 0
                Else
                    Shape1(i + 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 Then
                For q = 1 To 8
                    If q = 1 And Lab1(i - 16).Enabled = True Then
                        Call Ai1(i - 16): Ai = 1
                    ElseIf q = 2 And Lab1(i - 15).Enabled = True Then
                        Call Ai1(i - 15): Ai = 1
                    ElseIf q = 3 And Lab1(i - 14).Enabled = True Then
                        Call Ai1(i - 14): Ai = 1
                    ElseIf q = 4 And Lab1(i - 1).Enabled = True Then
                        Call Ai1(i - 1): Ai = 1
                    ElseIf q = 5 And Lab1(i + 1).Enabled = True Then
                        Call Ai1(i + 1): Ai = 1
                    ElseIf q = 6 And Lab1(i + 14).Enabled = True Then
                        Call Ai1(i + 14): Ai = 1
                    ElseIf q = 7 And Lab1(i + 15).Enabled = True Then
                        Call Ai1(i + 15): Ai = 1
                    ElseIf q = 8 And Lab1(i + 16).Enabled = True Then
                        Call Ai1(i + 16): Ai = 1
                    End If
                Next q
            End If
        End If
    ElseIf i = 208 Then '208起特殊情况
        B = Int(3 * Rnd)
        If B = 0 And Lab1(i - 17).Enabled = True Then
            Call Ai1(i - 17): Ai = 1
        ElseIf B = 1 And Lab1(i - 32).Enabled = True Then
            Call Ai1(i - 32): Ai = 1
        ElseIf B = 2 And Lab1(i - 31).Enabled = True Then
            Call Ai1(i - 31): Ai = 1
        Else
            q = Int(8 * Rnd + 1)
            If Ai = 0 And q = 1 And Lab1(i - 16).Enabled = True Then
                Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 16): Ai = 1
                Else
                    Shape1(i - 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 2 And Lab1(i - 15).Enabled = True Then
                Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 15): Ai = 1
                Else
                    Shape1(i - 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 3 And Lab1(i - 14).Enabled = True Then
                Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 14): Ai = 1
                Else
                    Shape1(i - 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 4 And Lab1(i - 1).Enabled = True Then
                Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 1): Ai = 1
                Else
                    Shape1(i - 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 5 And Lab1(i + 1).Enabled = True Then
                Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 1): Ai = 1
                Else
                    Shape1(i + 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 6 And Lab1(i + 14).Enabled = True Then
                Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 14): Ai = 1
                Else
                    Shape1(i + 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 7 And Lab1(i + 15).Enabled = True Then
                Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 15): Ai = 1
                Else
                    Shape1(i + 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 8 And Lab1(i + 16).Enabled = True Then
                Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 16): Ai = 0
                Else
                    Shape1(i + 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 Then
                For q = 1 To 8
                    If q = 1 And Lab1(i - 16).Enabled = True Then
                        Call Ai1(i - 16): Ai = 1
                    ElseIf q = 2 And Lab1(i - 15).Enabled = True Then
                        Call Ai1(i - 15): Ai = 1
                    ElseIf q = 3 And Lab1(i - 14).Enabled = True Then
                        Call Ai1(i - 14): Ai = 1
                    ElseIf q = 4 And Lab1(i - 1).Enabled = True Then
                        Call Ai1(i - 1): Ai = 1
                    ElseIf q = 5 And Lab1(i + 1).Enabled = True Then
                        Call Ai1(i + 1): Ai = 1
                    ElseIf q = 6 And Lab1(i + 14).Enabled = True Then
                        Call Ai1(i + 14): Ai = 1
                    ElseIf q = 7 And Lab1(i + 15).Enabled = True Then
                        Call Ai1(i + 15): Ai = 1
                    ElseIf q = 8 And Lab1(i + 16).Enabled = True Then
                        Call Ai1(i + 16): Ai = 1
                    End If
                Next q
            End If
        End If
    ElseIf i = 17 Or i = 18 Or i = 19 Or i = 20 Or i = 21 Or i = 22 Or i = 23 Or i = 24 Or i = 25 Or i = 26 Or i = 27 Then
        B = Int(3 * Rnd)
        If B = 0 And Lab1(i + 30).Enabled = True Then
            Call Ai1(i + 30): Ai = 1
        ElseIf B = 1 And Lab1(i + 29).Enabled = True Then
            Call Ai1(i + 29): Ai = 1
        ElseIf B = 2 And Lab1(i + 31).Enabled = True Then
            Call Ai1(i + 31): Ai = 1
        Else
            q = Int(8 * Rnd + 1)
            If Ai = 0 And q = 1 And Lab1(i - 16).Enabled = True Then
                Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 16): Ai = 1
                Else
                    Shape1(i - 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 2 And Lab1(i - 15).Enabled = True Then
                Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 15): Ai = 1
                Else
                    Shape1(i - 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 3 And Lab1(i - 14).Enabled = True Then
                Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 14): Ai = 1
                Else
                    Shape1(i - 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 4 And Lab1(i - 1).Enabled = True Then
                Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 1): Ai = 1
                Else
                    Shape1(i - 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 5 And Lab1(i + 1).Enabled = True Then
                Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 1): Ai = 1
                Else
                    Shape1(i + 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 6 And Lab1(i + 14).Enabled = True Then
                Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 14): Ai = 1
                Else
                    Shape1(i + 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 7 And Lab1(i + 15).Enabled = True Then
                Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 15): Ai = 1
                Else
                    Shape1(i + 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 8 And Lab1(i + 16).Enabled = True Then
                Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 16): Ai = 0
                Else
                    Shape1(i + 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 Then
                For q = 1 To 8
                    If q = 1 And Lab1(i - 16).Enabled = True Then
                        Call Ai1(i - 16): Ai = 1
                    ElseIf q = 2 And Lab1(i - 15).Enabled = True Then
                        Call Ai1(i - 15): Ai = 1
                    ElseIf q = 3 And Lab1(i - 14).Enabled = True Then
                        Call Ai1(i - 14): Ai = 1
                    ElseIf q = 4 And Lab1(i - 1).Enabled = True Then
                        Call Ai1(i - 1): Ai = 1
                    ElseIf q = 5 And Lab1(i + 1).Enabled = True Then
                        Call Ai1(i + 1): Ai = 1
                    ElseIf q = 6 And Lab1(i + 14).Enabled = True Then
                        Call Ai1(i + 14): Ai = 1
                    ElseIf q = 7 And Lab1(i + 15).Enabled = True Then
                        Call Ai1(i + 15): Ai = 1
                    ElseIf q = 8 And Lab1(i + 16).Enabled = True Then
                        Call Ai1(i + 16): Ai = 1
                    End If
                Next q
            End If
        End If
    ElseIf i = 197 Or i = 198 Or i = 199 Or i = 200 Or i = 201 Or i = 202 Or i = 203 Or i = 204 Or i = 205 Or i = 206 Or i = 207 Then
        B = Int(3 * Rnd)
        If B = 0 And Lab1(i - 30).Enabled = True Then
            Call Ai1(i - 30): Ai = 1
        ElseIf B = 1 And Lab1(i - 29).Enabled = True Then
            Call Ai1(i - 29): Ai = 1
        ElseIf B = 2 And Lab1(i - 31).Enabled = True Then
            Call Ai1(i - 31): Ai = 1
        Else
            q = Int(8 * Rnd + 1)
            If Ai = 0 And q = 1 And Lab1(i - 16).Enabled = True Then
                Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 16): Ai = 1
                Else
                    Shape1(i - 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 2 And Lab1(i - 15).Enabled = True Then
                Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 15): Ai = 1
                Else
                    Shape1(i - 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 3 And Lab1(i - 14).Enabled = True Then
                Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 14): Ai = 1
                Else
                    Shape1(i - 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 4 And Lab1(i - 1).Enabled = True Then
                Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 1): Ai = 1
                Else
                    Shape1(i - 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 5 And Lab1(i + 1).Enabled = True Then
                Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 1): Ai = 1
                Else
                    Shape1(i + 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 6 And Lab1(i + 14).Enabled = True Then
                Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 14): Ai = 1
                Else
                    Shape1(i + 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 7 And Lab1(i + 15).Enabled = True Then
                Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 15): Ai = 1
                Else
                    Shape1(i + 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 8 And Lab1(i + 16).Enabled = True Then
                Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 16): Ai = 0
                Else
                    Shape1(i + 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 Then
                For q = 1 To 8
                    If q = 1 And Lab1(i - 16).Enabled = True Then
                        Call Ai1(i - 16): Ai = 1
                    ElseIf q = 2 And Lab1(i - 15).Enabled = True Then
                        Call Ai1(i - 15): Ai = 1
                    ElseIf q = 3 And Lab1(i - 14).Enabled = True Then
                        Call Ai1(i - 14): Ai = 1
                    ElseIf q = 4 And Lab1(i - 1).Enabled = True Then
                        Call Ai1(i - 1): Ai = 1
                    ElseIf q = 5 And Lab1(i + 1).Enabled = True Then
                        Call Ai1(i + 1): Ai = 1
                    ElseIf q = 6 And Lab1(i + 14).Enabled = True Then
                        Call Ai1(i + 14): Ai = 1
                    ElseIf q = 7 And Lab1(i + 15).Enabled = True Then
                        Call Ai1(i + 15): Ai = 1
                    ElseIf q = 8 And Lab1(i + 16).Enabled = True Then
                        Call Ai1(i + 16): Ai = 1
                    End If
                Next q
            End If
        End If
    ElseIf i = 31 Or i = 46 Or i = 61 Or i = 76 Or i = 91 Or i = 106 Or i = 121 Or i = 136 Or i = 151 Or i = 166 Or i = 181 Then
        B = Int(3 * Rnd)
        If B = 0 And Lab1(i + 2).Enabled = True Then
            Call Ai1(i + 2): Ai = 1
        ElseIf B = 1 And Lab1(i - 13).Enabled = True Then
            Call Ai1(i - 13): Ai = 1
        ElseIf B = 2 And Lab1(i + 17).Enabled = True Then
            Call Ai1(i + 17): Ai = 1
        Else
            q = Int(8 * Rnd + 1)
            If Ai = 0 And q = 1 And Lab1(i - 16).Enabled = True Then
                Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 16): Ai = 1
                Else
                    Shape1(i - 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 2 And Lab1(i - 15).Enabled = True Then
                Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 15): Ai = 1
                Else
                    Shape1(i - 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 3 And Lab1(i - 14).Enabled = True Then
                Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 14): Ai = 1
                Else
                    Shape1(i - 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 4 And Lab1(i - 1).Enabled = True Then
                Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 1): Ai = 1
                Else
                    Shape1(i - 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 5 And Lab1(i + 1).Enabled = True Then
                Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 1): Ai = 1
                Else
                    Shape1(i + 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 6 And Lab1(i + 14).Enabled = True Then
                Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 14): Ai = 1
                Else
                    Shape1(i + 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 7 And Lab1(i + 15).Enabled = True Then
                Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 15): Ai = 1
                Else
                    Shape1(i + 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 8 And Lab1(i + 16).Enabled = True Then
                Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 16): Ai = 0
                Else
                    Shape1(i + 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 Then
                For q = 1 To 8
                    If q = 1 And Lab1(i - 16).Enabled = True Then
                        Call Ai1(i - 16): Ai = 1
                    ElseIf q = 2 And Lab1(i - 15).Enabled = True Then
                        Call Ai1(i - 15): Ai = 1
                    ElseIf q = 3 And Lab1(i - 14).Enabled = True Then
                        Call Ai1(i - 14): Ai = 1
                    ElseIf q = 4 And Lab1(i - 1).Enabled = True Then
                        Call Ai1(i - 1): Ai = 1
                    ElseIf q = 5 And Lab1(i + 1).Enabled = True Then
                        Call Ai1(i + 1): Ai = 1
                    ElseIf q = 6 And Lab1(i + 14).Enabled = True Then
                        Call Ai1(i + 14): Ai = 1
                    ElseIf q = 7 And Lab1(i + 15).Enabled = True Then
                        Call Ai1(i + 15): Ai = 1
                    ElseIf q = 8 And Lab1(i + 16).Enabled = True Then
                        Call Ai1(i + 16): Ai = 1
                    End If
                Next q
            End If
        End If
    ElseIf i = 43 Or i = 58 Or i = 73 Or i = 88 Or i = 103 Or i = 118 Or i = 133 Or i = 148 Or i = 163 Or i = 178 Or i = 193 Then
        B = Int(3 * Rnd)
        If B = 0 And Lab1(i - 2).Enabled = True Then
            Call Ai1(i - 2): Ai = 1
        ElseIf B = 1 And Lab1(i + 13).Enabled = True Then
            Call Ai1(i + 13): Ai = 1
        ElseIf B = 2 And Lab1(i - 17).Enabled = True Then
            Call Ai1(i - 17): Ai = 1
        Else
            q = Int(8 * Rnd + 1)
            If Ai = 0 And q = 1 And Lab1(i - 16).Enabled = True Then
                Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 16): Ai = 1
                Else
                    Shape1(i - 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 2 And Lab1(i - 15).Enabled = True Then
                Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 15): Ai = 1
                Else
                    Shape1(i - 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 3 And Lab1(i - 14).Enabled = True Then
                Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 14): Ai = 1
                Else
                    Shape1(i - 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 4 And Lab1(i - 1).Enabled = True Then
                Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i - 1): Ai = 1
                Else
                    Shape1(i - 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 5 And Lab1(i + 1).Enabled = True Then
                Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 1): Ai = 1
                Else
                    Shape1(i + 1).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 6 And Lab1(i + 14).Enabled = True Then
                Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 14): Ai = 1
                Else
                    Shape1(i + 14).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 7 And Lab1(i + 15).Enabled = True Then
                Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 15): Ai = 1
                Else
                    Shape1(i + 15).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 And q = 8 And Lab1(i + 16).Enabled = True Then
                Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                Score_P = 1: Score_C = 1
                Call Score(Score_P, Score_C)
                If Score_C >= Score_P Then
                    Call Ai1(i + 16): Ai = 0
                Else
                    Shape1(i + 16).FillColor = &H8000& '还原颜色
                End If
            ElseIf Ai = 0 Then
                For q = 1 To 8
                    If q = 1 And Lab1(i - 16).Enabled = True Then
                        Call Ai1(i - 16): Ai = 1
                    ElseIf q = 2 And Lab1(i - 15).Enabled = True Then
                        Call Ai1(i - 15): Ai = 1
                    ElseIf q = 3 And Lab1(i - 14).Enabled = True Then
                        Call Ai1(i - 14): Ai = 1
                    ElseIf q = 4 And Lab1(i - 1).Enabled = True Then
                        Call Ai1(i - 1): Ai = 1
                    ElseIf q = 5 And Lab1(i + 1).Enabled = True Then
                        Call Ai1(i + 1): Ai = 1
                    ElseIf q = 6 And Lab1(i + 14).Enabled = True Then
                        Call Ai1(i + 14): Ai = 1
                    ElseIf q = 7 And Lab1(i + 15).Enabled = True Then
                        Call Ai1(i + 15): Ai = 1
                    ElseIf q = 8 And Lab1(i + 16).Enabled = True Then
                        Call Ai1(i + 16): Ai = 1
                    End If
                Next q
            End If
        End If
    Else
        For m = 0 To 10
            For r = 0 To 10
                If i = 32 + 15 * m + r Then
                    For q = 1 To 8
                        If q = 1 And Ai = 0 And Lab1(i - 16).Enabled = True Then
                            Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                            score_c1 = Score_C
                            score_p1 = Score_P
                            Score_P = 1: Score_C = 1
                            Call Score(Score_P, Score_C)
                            If Score_C >= 200 * score_c1 Or 1.005 * Score_P <= score_p1 Then
                                Call Ai1(i - 16): Ai = 1
                            Else
                                Shape1(i - 16).FillColor = &H8000& '还原颜色
                            End If
                        ElseIf q = 2 And Ai = 0 And Lab1(i - 15).Enabled = True Then
                            Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                            score_c1 = Score_C
                            score_p1 = Score_P
                            Score_P = 1: Score_C = 1
                            Call Score(Score_P, Score_C)
                            If Score_C >= 200 * score_c1 Or 1.005 * Score_P <= score_p1 Then
                                Call Ai1(i - 15): Ai = 1
                            Else
                                Shape1(i - 15).FillColor = &H8000& '还原颜色
                            End If
                        ElseIf q = 3 And Ai = 0 And Lab1(i - 14).Enabled = True Then
                            Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                            score_c1 = Score_C
                            score_p1 = Score_P
                            Score_P = 1: Score_C = 1
                            Call Score(Score_P, Score_C)
                            If Score_C >= 200 * score_c1 Or 1.005 * Score_P <= score_p1 Then
                                Call Ai1(i - 14): Ai = 1
                            Else
                                Shape1(i - 14).FillColor = &H8000& '还原颜色
                            End If
                        ElseIf q = 4 And Ai = 0 And Lab1(i - 1).Enabled = True Then
                            Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                            score_c1 = Score_C
                            score_p1 = Score_P
                            Score_P = 1: Score_C = 1
                            Call Score(Score_P, Score_C)
                            If Score_C >= 200 * score_c1 Or 1.005 * Score_P <= score_p1 Then
                                Call Ai1(i - 1): Ai = 1
                            Else
                                Shape1(i - 1).FillColor = &H8000& '还原颜色
                            End If
                        ElseIf q = 5 And Ai = 0 And Lab1(i + 1).Enabled = True Then
                              Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                            score_c1 = Score_C
                            score_p1 = Score_P
                            Score_P = 1: Score_C = 1
                            Call Score(Score_P, Score_C)
                            If Score_C >= 200 * score_c1 Or 1.005 * Score_P <= score_p1 Then
                                Call Ai1(i + 1): Ai = 1
                            Else
                                Shape1(i + 1).FillColor = &H8000& '还原颜色
                            End If
                        ElseIf q = 6 And Ai = 0 And Lab1(i + 14).Enabled = True Then
                            Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                            score_c1 = Score_C
                            score_p1 = Score_P
                            Score_P = 1: Score_C = 1
                            Call Score(Score_P, Score_C)
                            If Score_C >= 200 * score_c1 Or 1.005 * Score_P <= score_p1 Then
                                Call Ai1(i + 14): Ai = 1
                            Else
                                Shape1(i + 14).FillColor = &H8000& '还原颜色
                            End If
                        ElseIf q = 7 And Ai = 0 And Lab1(i + 15).Enabled = True Then
                            Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                            score_c1 = Score_C
                            score_p1 = Score_P
                            Score_P = 1: Score_C = 1
                            Call Score(Score_P, Score_C)
                            If Score_C >= 200 * score_c1 Or 1.005 * Score_P <= score_p1 Then
                                Call Ai1(i + 15): Ai = 1
                            Else
                                Shape1(i + 15).FillColor = &H8000& '还原颜色
                            End If
                        ElseIf q = 8 And Ai = 0 And Lab1(i + 16).Enabled = True Then
                            Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                            score_c1 = Score_C
                            score_p1 = Score_P
                            Score_P = 1: Score_C = 1
                            Call Score(Score_P, Score_C)
                            If Score_C >= 200 * score_c1 Or 1.005 * Score_P <= score_p1 Then
                                Call Ai1(i + 16): Ai = 1
                            Else
                                Shape1(i + 16).FillColor = &H8000& '还原颜色
                            End If
                        End If
                    Next q
                    q = Int(16 * Rnd) + 1
                    If Ai = 0 And q = 1 And Lab1(i - 16).Enabled = True Then
                        Shape1(i - 16).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i - 16): Ai = 1
                        Else
                            Shape1(i - 16).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 2 And Lab1(i - 15).Enabled = True Then
                        Shape1(i - 15).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i - 15): Ai = 1
                        Else
                            Shape1(i - 15).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 3 And Lab1(i - 14).Enabled = True Then
                        Shape1(i - 14).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i - 14): Ai = 1
                        Else
                            Shape1(i - 14).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 4 And Lab1(i - 1).Enabled = True Then
                        Shape1(i - 1).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i - 1): Ai = 1
                        Else
                            Shape1(i - 1).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 5 And Lab1(i + 1).Enabled = True Then
                        Shape1(i + 1).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i + 1): Ai = 1
                        Else
                            Shape1(i + 1).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 6 And Lab1(i + 14).Enabled = True Then
                        Shape1(i + 14).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i + 14): Ai = 1
                        Else
                            Shape1(i + 14).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 7 And Lab1(i + 15).Enabled = True Then
                        Shape1(i + 15).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i + 15): Ai = 1
                        Else
                            Shape1(i + 15).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 8 And Lab1(i + 16).Enabled = True Then
                        Shape1(i + 16).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i + 16): Ai = 0
                        Else
                            Shape1(i + 16).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 9 And Lab1(i - 32).Enabled = True Then
                        Shape1(i - 32).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i - 32): Ai = 1
                        Else
                            Shape1(i - 32).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 10 And Lab1(i - 30).Enabled = True Then
                        Shape1(i - 30).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i - 30): Ai = 1
                        Else
                            Shape1(i - 30).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 11 And Lab1(i - 2).Enabled = True Then
                        Shape1(i - 2).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i - 2): Ai = 1
                        Else
                            Shape1(i - 2).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 12 And Lab1(i + 2).Enabled = True Then
                        Shape1(i + 2).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i + 2): Ai = 1
                        Else
                            Shape1(i + 2).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 13 And Lab1(i + 28).Enabled = True Then
                        Shape1(i + 28).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i + 28): Ai = 1
                        Else
                            Shape1(i + 28).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 14 And Lab1(i + 30).Enabled = True Then
                        Shape1(i + 30).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i + 30): Ai = 1
                        Else
                            Shape1(i + 30).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 15 And Lab1(i + 32).Enabled = True Then
                        Shape1(i + 32).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i + 32): Ai = 1
                        Else
                            Shape1(i + 32).FillColor = &H8000& '还原颜色
                        End If
                    ElseIf Ai = 0 And q = 16 And Lab1(i - 28).Enabled = True Then
                        Shape1(i - 28).FillColor = Color_C '机器下棋的颜色
                        score_c1 = Score_C
                        score_p1 = Score_P
                        Score_P = 1: Score_C = 1
                        Call Score(Score_P, Score_C)
                        If Score_C >= Score_P Or Score_P <= score_p1 And score_c1 <= Score_C Then
                            Call Ai1(i - 28): Ai = 1
                        Else
                            Shape1(i - 28).FillColor = &H8000& '还原颜色
                        End If
                    End If
                End If
            Next r
        Next m
    End If
    If Ai = 0 Then Call DuanNao1(i)
End Sub

Private Sub Ai1(i As Integer)
    chess = chess + 1   '机器下棋并录入次数
    ReDim Preserve Remember(chess)
    Remember(chess) = i
    Shape1(i).Visible = True
    Shape1(i).FillColor = Color_C '机器下棋的颜色
    Shape1(i).BorderColor = Color_P '棋子边框颜色交换
    Lab1(i).BorderStyle = 1 '显示落子位置
    Lab1(i).Enabled = False
    V = i
    If chess < 2 Then
        Command3.Enabled = False
    Else
        Command3.Enabled = True
    End If
End Sub

