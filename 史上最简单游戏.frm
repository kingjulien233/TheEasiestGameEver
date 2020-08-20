VERSION 5.00
Begin VB.Form 史上最简单游戏 
   Caption         =   "史上最简单游戏"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6090
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CommandS 
      Caption         =   "开始!"
      Height          =   495
      Left            =   2400
      TabIndex        =   25
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CommandE 
      Caption         =   "退出!"
      Height          =   495
      Left            =   2400
      TabIndex        =   27
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton CommandM 
      Caption         =   "主菜单!"
      Height          =   495
      Left            =   2400
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CommandR 
      Caption         =   "重试!"
      Height          =   495
      Left            =   2400
      TabIndex        =   24
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   3600
   End
   Begin VB.Label LabelEND 
      Caption         =   "什么！？你竟然通关了这款游戏！？大神请受我一拜！"
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      TabIndex        =   30
      Top             =   1320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "TheEasiestGameEver   NaiveGames   Copyright2014"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5520
      Width           =   5775
   End
   Begin VB.Label LabelT 
      Alignment       =   2  'Center
      Caption         =   "史上最简单游戏"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   28
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label LabelAI9 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI14 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label0 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   4920
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI12 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   3360
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI13 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI10 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI11 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI15 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI16 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI17 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI18 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI19 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI7 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI8 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI4 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI3 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI2 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI1 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI6 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelAI5 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Height          =   5415
      Left            =   5760
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "史上最简单游戏"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer                '确定按下左或右按键
Dim b As Integer                '确定按下上或下按键
Dim c As Single                 '确定松开左或右按键
Dim d As Single                 '确定松开上或下按键
Dim e As Integer                '判定红色方块是否在运动边界内
Dim f As Integer
Dim g As Integer                '确定红色方块的位置
Dim h As Integer
Dim i As Integer                '判定红色方块是否在运动边界内
Dim j As Integer
Dim k As Integer                '确定红色方块的位置
Dim l As Integer

Private Sub CommandE_Click()
End                             '结束游戏
End Sub

Private Sub CommandM_Click()
LabelT.Visible = True           '返回菜单，隐藏红色方块与蓝色方块
LabelAI1.Visible = False
LabelAI2.Visible = False
LabelAI3.Visible = False
LabelAI4.Visible = False
LabelAI5.Visible = False
LabelAI6.Visible = False
LabelAI7.Visible = False
LabelAI8.Visible = False
LabelAI9.Visible = False
LabelAI10.Visible = False
LabelAI11.Visible = False
LabelAI12.Visible = False
LabelAI13.Visible = False
LabelAI14.Visible = False
LabelAI15.Visible = False
LabelAI16.Visible = False
LabelAI17.Visible = False
LabelAI18.Visible = False
LabelAI19.Visible = False
Label0.Visible = False
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
CommandS.Visible = True
CommandE.Visible = True
CommandR.Visible = False
CommandM.Visible = False
Timer1.Enabled = False          '关闭计时器
End Sub

Private Sub CommandR_Click()
a = 0                           '将控制运动的变量恢复起始值
b = 0
c = 0
d = 0
Timer1.Enabled = True           '开启计时器
CommandR.Visible = False
CommandM.Visible = False
LabelAI1.Left = 360             '恢复红色方块的位置
LabelAI2.Left = 360
LabelAI3.Left = 360
LabelAI4.Left = 360
LabelAI5.Left = 5400
LabelAI6.Left = 5400
LabelAI7.Left = 5400
LabelAI8.Left = 5400
LabelAI9.Top = 240
LabelAI10.Top = 240
LabelAI11.Top = 240
LabelAI12.Top = 240
LabelAI13.Top = 240
LabelAI14.Top = 240
LabelAI15.Top = 4560
LabelAI16.Top = 4560
LabelAI17.Top = 4560
LabelAI18.Top = 4560
LabelAI19.Top = 4560
Label1.Left = 2880              '恢复蓝色方块的位置
Label1.Top = 240
End Sub

Private Sub CommandS_Click()
LabelT.Visible = False          '隐藏菜单，显示红色方块与蓝色方块
LabelAI1.Visible = True
LabelAI2.Visible = True
LabelAI3.Visible = True
LabelAI4.Visible = True
LabelAI5.Visible = True
LabelAI6.Visible = True
LabelAI7.Visible = True
LabelAI8.Visible = True
LabelAI9.Visible = True
LabelAI10.Visible = True
LabelAI11.Visible = True
LabelAI12.Visible = True
LabelAI13.Visible = True
LabelAI14.Visible = True
LabelAI15.Visible = True
LabelAI16.Visible = True
LabelAI17.Visible = True
LabelAI18.Visible = True
LabelAI19.Visible = True
Label0.Visible = True
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
CommandS.Visible = False
CommandE.Visible = False
Timer1.Enabled = True           '启动红色方块的运动与蓝色方块的移动
a = 0                           '将控制运动的变量恢复起始值
b = 0
c = 0
d = 0
LabelAI1.Left = 360             '确定红色方块起始位置
LabelAI2.Left = 360
LabelAI3.Left = 360
LabelAI4.Left = 360
LabelAI5.Left = 5400
LabelAI6.Left = 5400
LabelAI7.Left = 5400
LabelAI8.Left = 5400
LabelAI9.Top = 240
LabelAI10.Top = 240
LabelAI11.Top = 240
LabelAI12.Top = 240
LabelAI13.Top = 240
LabelAI14.Top = 240
LabelAI15.Top = 4560
LabelAI16.Top = 4560
LabelAI17.Top = 4560
LabelAI18.Top = 4560
LabelAI19.Top = 4560
Label1.Left = 2880
Label1.Top = 240
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then     '按下左或右按键时为变量a赋值
  a = -1
Else
  If KeyCode = vbKeyRight Then
    a = 1
  Else
    If KeyCode = vbKeyUp Then   '按下上或下按键时为变量b赋值
      b = -1
    Else
      If KeyCode = vbKeyDown Then '
        b = 1
      End If
    End If
  End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then     '松开左或右按键时为变量a赋值
  a = 0
Else
  If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then      '松开上或下按键时为变量b赋值
    b = 0
  End If
End If
End Sub

Private Sub Form_Load()
z = 1
End Sub

Private Sub Timer1_Timer()
If LabelAI1.Left <= 360 Then    '判定红色方块是否在运动边界内
  e = -1
Else
  If LabelAI1.Left >= 5400 Then
    e = 1
  End If
End If
If LabelAI5.Left <= 360 Then
  f = -1
Else
  If LabelAI5.Left >= 5400 Then
    f = 1
  End If
End If
If LabelAI9.Top <= 240 Then
  i = -1
Else
  If LabelAI9.Top >= 4560 Then
    i = 1
  End If
End If
If LabelAI15.Top <= 240 Then
  j = -1
Else
  If LabelAI15.Top >= 4560 Then
    j = 1
  End If
End If
If e = 1 Then                   '若红色方块在运动边界外则改变其运动方向
  LabelAI1.Left = LabelAI1.Left - 50
  LabelAI2.Left = LabelAI2.Left - 50
  LabelAI3.Left = LabelAI3.Left - 50
  LabelAI4.Left = LabelAI4.Left - 50
Else
  If e = -1 Then
    LabelAI1.Left = LabelAI1.Left + 50
    LabelAI2.Left = LabelAI2.Left + 50
    LabelAI3.Left = LabelAI3.Left + 50
    LabelAI4.Left = LabelAI4.Left + 50
  End If
End If
If f = 1 Then
  LabelAI5.Left = LabelAI5.Left - 50
  LabelAI6.Left = LabelAI6.Left - 50
  LabelAI7.Left = LabelAI7.Left - 50
  LabelAI8.Left = LabelAI8.Left - 50
Else
  If f = -1 Then
    LabelAI5.Left = LabelAI5.Left + 50
    LabelAI6.Left = LabelAI6.Left + 50
    LabelAI7.Left = LabelAI7.Left + 50
    LabelAI8.Left = LabelAI8.Left + 50
  End If
End If
If i = 1 Then
  LabelAI9.Top = LabelAI9.Top - 50
  LabelAI10.Top = LabelAI10.Top - 50
  LabelAI11.Top = LabelAI11.Top - 50
  LabelAI12.Top = LabelAI12.Top - 50
  LabelAI13.Top = LabelAI13.Top - 50
  LabelAI14.Top = LabelAI14.Top - 50
Else
  If i = -1 Then
    LabelAI9.Top = LabelAI9.Top + 50
    LabelAI10.Top = LabelAI10.Top + 50
    LabelAI11.Top = LabelAI11.Top + 50
    LabelAI12.Top = LabelAI12.Top + 50
    LabelAI13.Top = LabelAI13.Top + 50
    LabelAI14.Top = LabelAI14.Top + 50
  End If
End If
If j = 1 Then
  LabelAI15.Top = LabelAI15.Top - 50
  LabelAI16.Top = LabelAI16.Top - 50
  LabelAI17.Top = LabelAI17.Top - 50
  LabelAI18.Top = LabelAI18.Top - 50
  LabelAI19.Top = LabelAI19.Top - 50
Else
  If j = -1 Then
    LabelAI15.Top = LabelAI15.Top + 50
    LabelAI16.Top = LabelAI16.Top + 50
    LabelAI17.Top = LabelAI17.Top + 50
    LabelAI18.Top = LabelAI18.Top + 50
    LabelAI19.Top = LabelAI19.Top + 50
  End If
End If
g = LabelAI1.Left               '确定红色方块的位置以判定是否在运动边界内
h = LabelAI5.Left
k = LabelAI9.Top
l = LabelAI15.Top
c = c + (a - c) / 20            '控制蓝色方块运动流畅度
Label1.Left = Label1.Left + c * 20      '控制蓝色方块运动速度
d = d + (b - d) / 20
Label1.Top = Label1.Top + d * 20
If Label1.Left >= 5520 Or Label1.Left <= 240 Or Label1.Top >= 4920 Or Label1.Top <= 120 Then        '蓝色方块出界的死亡判定
  Timer1.Enabled = False
  CommandR.Visible = True
  CommandM.Visible = True
End If
If Label1.Left <= g + 240 And Label1.Left >= g - 240 And ((Label1.Top <= 1440 And Label1.Top >= 960) Or (Label1.Top <= 2400 And Label1.Top >= 1920) Or (Label1.Top <= 3360 And Label1.Top >= 2880) Or (Label1.Top <= 4320 And Label1.Top >= 3840)) Then
  Timer1.Enabled = False
  CommandR.Visible = True
  CommandM.Visible = True
End If
If Label1.Left <= h + 240 And Label1.Left >= h - 240 And ((Label1.Top <= 960 And Label1.Top >= 480) Or (Label1.Top <= 1920 And Label1.Top >= 1440) Or (Label1.Top <= 2880 And Label1.Top >= 2400) Or (Label1.Top <= 3840 And Label1.Top >= 3360)) Then
  Timer1.Enabled = False
  CommandR.Visible = True
  CommandM.Visible = True
End If
If Label1.Top <= k + 240 And Label1.Top >= k - 240 And ((Label1.Left <= 720 And Label1.Left >= 240) Or (Label1.Left <= 1680 And Label1.Left >= 1200) Or (Label1.Left <= 2640 And Label1.Left >= 2160) Or (Label1.Left <= 3600 And Label1.Left >= 3120) Or (Label1.Left <= 4560 And Label1.Left >= 4080) Or (Label1.Left <= 5520 And Label1.Left >= 5040)) Then
  Timer1.Enabled = False
  CommandR.Visible = True
  CommandM.Visible = True
End If
If Label1.Top <= l + 240 And Label1.Top >= l - 240 And ((Label1.Left <= 1200 And Label1.Left >= 720) Or (Label1.Left <= 2160 And Label1.Left >= 1680) Or (Label1.Left <= 3120 And Label1.Left >= 2640) Or (Label1.Left <= 4080 And Label1.Left >= 3600) Or (Label1.Left <= 5040 And Label1.Left >= 4560)) Then
  Timer1.Enabled = False
  CommandR.Visible = True
  CommandM.Visible = True
End If
If Label1.Top >= 4680 Then      '蓝色方块到达终点的成功判定
  LabelEND.Visible = True
  LabelAI1.Visible = False
  LabelAI2.Visible = False
  LabelAI3.Visible = False
  LabelAI4.Visible = False
  LabelAI5.Visible = False
  LabelAI6.Visible = False
  LabelAI7.Visible = False
  LabelAI8.Visible = False
  LabelAI9.Visible = False
  LabelAI10.Visible = False
  LabelAI11.Visible = False
  LabelAI12.Visible = False
  LabelAI13.Visible = False
  LabelAI14.Visible = False
  LabelAI15.Visible = False
  LabelAI16.Visible = False
  LabelAI17.Visible = False
  LabelAI18.Visible = False
  LabelAI19.Visible = False
  Label0.Visible = False
  Label1.Visible = False
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  CommandE.Visible = True
  CommandR.Visible = False
  CommandM.Visible = False
  Timer1.Enabled = False
End If
End Sub
