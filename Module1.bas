Attribute VB_Name = "Module1"
Public BB1 As String    '版本号
Public BB2 As String    '版本最后修订时间
Public Sh As Integer    '动画时间同步
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long '导入外部链接设定
Public Di As Integer    '判断是否已经开局
Public N As Integer     '重置来源标志
Public X As Integer     '判断是否落子已经加载过
Public Color_P  As String   '人的棋子颜色
Public Color_C  As String   '机器的棋子颜色
Public Score_P As Long      '人的分数
Public Score_C As Long      '机器的分数
Public gameover As Boolean  '结束游戏标志
Public way As Integer   '游戏结束方式(胜负)
Public V As Integer     '上一次的AI落子位置
Public Difficility As Integer '判断当前难度
Public chess As Integer     '当前落子总数
Public Ren_X As Long
Public Ren_Y As Long
Public Jiqi_X As Long
Public Jiqi_Y As Long
'补棋盘所需参数
