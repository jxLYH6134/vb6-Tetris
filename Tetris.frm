VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "俄罗斯方块"
   ClientHeight    =   4590
   ClientLeft      =   12420
   ClientTop       =   5250
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4305
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   4080
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4595
      Left            =   0
      ScaleHeight     =   4560
      ScaleWidth      =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   2915
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Height          =   360
         Index           =   3
         Left            =   600
         TabIndex        =   113
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   112
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Height          =   360
         Index           =   1
         Left            =   600
         TabIndex        =   111
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   360
         Index           =   111
         Left            =   2520
         TabIndex        =   121
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   360
         Index           =   110
         Left            =   2160
         TabIndex        =   120
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   360
         Index           =   109
         Left            =   1800
         TabIndex        =   119
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   360
         Index           =   108
         Left            =   1440
         TabIndex        =   118
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   360
         Index           =   107
         Left            =   1080
         TabIndex        =   117
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   360
         Index           =   106
         Left            =   720
         TabIndex        =   116
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   360
         Index           =   105
         Left            =   360
         TabIndex        =   115
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   360
         Index           =   104
         Left            =   0
         TabIndex        =   114
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   103
         Left            =   2520
         TabIndex        =   110
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   102
         Left            =   2160
         TabIndex        =   109
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   101
         Left            =   1800
         TabIndex        =   108
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   100
         Left            =   1440
         TabIndex        =   107
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   99
         Left            =   1080
         TabIndex        =   106
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   98
         Left            =   720
         TabIndex        =   105
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   97
         Left            =   360
         TabIndex        =   104
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   96
         Left            =   0
         TabIndex        =   103
         Top             =   4200
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   95
         Left            =   2520
         TabIndex        =   102
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   94
         Left            =   2160
         TabIndex        =   101
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   93
         Left            =   1800
         TabIndex        =   100
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   92
         Left            =   1440
         TabIndex        =   99
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   91
         Left            =   1080
         TabIndex        =   98
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   90
         Left            =   720
         TabIndex        =   97
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   89
         Left            =   360
         TabIndex        =   96
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   88
         Left            =   0
         TabIndex        =   95
         Top             =   3840
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   87
         Left            =   2520
         TabIndex        =   94
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   86
         Left            =   2160
         TabIndex        =   93
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   85
         Left            =   1800
         TabIndex        =   92
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   84
         Left            =   1440
         TabIndex        =   91
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   83
         Left            =   1080
         TabIndex        =   90
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   82
         Left            =   720
         TabIndex        =   89
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   81
         Left            =   360
         TabIndex        =   88
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   80
         Left            =   0
         TabIndex        =   87
         Top             =   3480
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   79
         Left            =   2520
         TabIndex        =   86
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   78
         Left            =   2160
         TabIndex        =   85
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   77
         Left            =   1800
         TabIndex        =   84
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   76
         Left            =   1440
         TabIndex        =   83
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   75
         Left            =   1080
         TabIndex        =   82
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   74
         Left            =   720
         TabIndex        =   81
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   73
         Left            =   360
         TabIndex        =   80
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   72
         Left            =   0
         TabIndex        =   79
         Top             =   3120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   71
         Left            =   2520
         TabIndex        =   78
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   70
         Left            =   2160
         TabIndex        =   77
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   69
         Left            =   1800
         TabIndex        =   76
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   68
         Left            =   1440
         TabIndex        =   75
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   67
         Left            =   1080
         TabIndex        =   74
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   66
         Left            =   720
         TabIndex        =   73
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   65
         Left            =   360
         TabIndex        =   72
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   64
         Left            =   0
         TabIndex        =   71
         Top             =   2760
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   63
         Left            =   2520
         TabIndex        =   70
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   62
         Left            =   2160
         TabIndex        =   69
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   61
         Left            =   1800
         TabIndex        =   68
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   60
         Left            =   1440
         TabIndex        =   67
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   59
         Left            =   1080
         TabIndex        =   66
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   58
         Left            =   720
         TabIndex        =   65
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   57
         Left            =   360
         TabIndex        =   64
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   56
         Left            =   0
         TabIndex        =   63
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   55
         Left            =   2520
         TabIndex        =   62
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   54
         Left            =   2160
         TabIndex        =   61
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   53
         Left            =   1800
         TabIndex        =   60
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   52
         Left            =   1440
         TabIndex        =   59
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   51
         Left            =   1080
         TabIndex        =   58
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   50
         Left            =   720
         TabIndex        =   57
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   49
         Left            =   360
         TabIndex        =   56
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   48
         Left            =   0
         TabIndex        =   55
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   47
         Left            =   2520
         TabIndex        =   54
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   46
         Left            =   2160
         TabIndex        =   53
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   45
         Left            =   1800
         TabIndex        =   52
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   44
         Left            =   1440
         TabIndex        =   51
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   43
         Left            =   1080
         TabIndex        =   50
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   42
         Left            =   720
         TabIndex        =   49
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   41
         Left            =   360
         TabIndex        =   48
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   40
         Left            =   0
         TabIndex        =   47
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   39
         Left            =   2520
         TabIndex        =   46
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   38
         Left            =   2160
         TabIndex        =   45
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   37
         Left            =   1800
         TabIndex        =   44
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   36
         Left            =   1440
         TabIndex        =   43
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   35
         Left            =   1080
         TabIndex        =   42
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   34
         Left            =   720
         TabIndex        =   41
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   33
         Left            =   360
         TabIndex        =   40
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   32
         Left            =   0
         TabIndex        =   39
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   31
         Left            =   2520
         TabIndex        =   38
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   30
         Left            =   2160
         TabIndex        =   37
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   29
         Left            =   1800
         TabIndex        =   36
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   28
         Left            =   1440
         TabIndex        =   35
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   27
         Left            =   1080
         TabIndex        =   34
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   26
         Left            =   720
         TabIndex        =   33
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   25
         Left            =   360
         TabIndex        =   32
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   24
         Left            =   0
         TabIndex        =   31
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   23
         Left            =   2520
         TabIndex        =   30
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   22
         Left            =   2160
         TabIndex        =   29
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   21
         Left            =   1800
         TabIndex        =   28
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   20
         Left            =   1440
         TabIndex        =   27
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   19
         Left            =   1080
         TabIndex        =   26
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   18
         Left            =   720
         TabIndex        =   25
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   17
         Left            =   360
         TabIndex        =   24
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   16
         Left            =   0
         TabIndex        =   23
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   15
         Left            =   2520
         TabIndex        =   22
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   14
         Left            =   2160
         TabIndex        =   21
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   13
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   12
         Left            =   1440
         TabIndex        =   19
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   11
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   10
         Left            =   720
         TabIndex        =   17
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   9
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   8
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   7
         Left            =   2520
         TabIndex        =   14
         Top             =   -120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   6
         Left            =   2160
         TabIndex        =   13
         Top             =   -120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   5
         Left            =   1800
         TabIndex        =   12
         Top             =   -120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   4
         Left            =   1440
         TabIndex        =   11
         Top             =   -120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         Top             =   -120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   2
         Left            =   720
         TabIndex        =   9
         Top             =   -120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   -120
         Width           =   360
      End
      Begin VB.Label Label2 
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   -120
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "下个"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   975
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Height          =   120
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   120
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Height          =   120
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   120
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3360
      Top             =   4080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3000
      Top             =   4080
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "已暂停"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3000
      TabIndex        =   122
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nxt%, Atv%, Drt%, Hgt%, Clr%, Mve As Boolean, Srt As Boolean
   '下个，当前，方向，最高，清除，移动锁定，      开始锁定

Private Sub Form_Activate()
Form1.Left = (Screen.Width - Form1.Width) / 2 '窗口居中
Form1.Top = (Screen.Height - Form1.Height) / 2

For a = 0 To 3 '初始化界面
    Label1(a).Left = 0
    Label1(a).Top = -360
Next a

For b = 0 To 103
    Label2(b).BackStyle = 0
Next b

For c = 0 To 1
    Label3(c).Left = 0
    Label3(c).Top = -120
Next c

Label4.Left = 960
Label4.Top = 1920

Mve = False '设定初始值
Srt = True

MsgBox "使用A/D左右移动 W变换方块 S加速下落", , "Tetris" '提示信息
End Sub

Private Sub Command1_Click()
Randomize

If Srt Then
    For i = 0 To 103 '刷新界面
        Label2(i).BackStyle = 0
    Next i
    
    Nxt = Int(Rnd * 5 + 1) '随机初始图形
    
    Drt = 0 '变量初始值
    Mve = True
    Srt = False
    Text1.Text = "0"
    Command1.Caption = "暂停"
    
    rePlace '跳转
Else
    If Timer1.Enabled Then '暂停游戏
        Timer1.Enabled = False
        Picture1.Visible = False
        Command1.Caption = "继续"
    ElseIf Not Timer3.Enabled Then '重新开始前倒计时
        Label4.Caption = "3"
        Timer3.Enabled = True
        Command1.Caption = "暂停"
    End If
End If
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim F As Boolean

Select Case KeyCode '键盘操控
    Case 65 'A
        F = Mve
        For i = 0 To 3
            If Label1(i).Left = 0 Then F = False: Exit For
            For j = 8 To 103
                If Label2(j).BackStyle = 1 And Label1(i).Left = Label2(j).Left + 360 And Label1(i).Top = Label2(j).Top Then F = False: Exit For
        Next j, i
        If F Then
            For k = 0 To 3
                Label1(k).Left = Label1(k).Left - 360
            Next k
        End If
    Case 68 'D
        F = Mve
        For i = 0 To 3
            If Label1(i).Left = 2520 Then F = False: Exit For
            For j = 8 To 103
                If Label2(j).BackStyle = 1 And Label1(i).Left = Label2(j).Left - 360 And Label1(i).Top = Label2(j).Top Then F = False: Exit For
        Next j, i
        If F Then
            For k = 0 To 3
                Label1(k).Left = Label1(k).Left + 360
            Next k
        End If
    Case 83 'S
        If Timer1.Enabled Then '加速下落
            Mve = False
            Timer1.Interval = 20
        End If
    Case 87 'W
        If Timer1.Enabled Then Rotate '跳转
End Select
End Sub

Private Sub Timer1_Timer()
Dim F As Boolean
F = True

For i = 0 To 3 '碰撞检测
    For j = 8 To 111
        If Label1(i).Top = Label2(j).Top - 360 And Label1(i).Left = Label2(j).Left And Label2(j).BackStyle = 1 Then F = False: Exit For
Next j, i

If F Then
    For a = 0 To 3 '下落
        Label1(a).Top = Label1(a).Top + 360
    Next a
Else
    Mve = False '刷新变量
    Timer1.Enabled = False
    Timer1.Interval = 300
    
    For i = 0 To 3 '图形替换
        For j = 103 To 0 Step -1
            If Label1(i).Top = Label2(j).Top And Label1(i).Left = Label2(j).Left Then
                Label2(j).BackStyle = 1
                Label2(j).BackColor = Label1(i).BackColor
            End If
    Next j, i
    
    For b = 0 To 3 '移位
        Label1(b).Left = 0
        Label1(b).Top = -360
    Next b
    
    For c = 0 To 7 '结束判定
        If Label2(c).BackStyle = 1 Then F = True: Exit For
    Next c
    
    If F Then
        Srt = True '结束游戏
        Command1.Caption = "开始"
        MsgBox "游戏结束", , "Tetris"
    Else
        For i = 96 To 8 Step -8 '消除判定
            F = True
            For j = i To i + 7
                If Label2(j).BackStyle = 0 Then F = False
            Next j
            If F Then Timer2.Enabled = True: Exit For
        Next i
        
        If Not F Then
            If Srt = False Then Text1.Text = Val(Text1.Text) + 1 '普通计分
            rePlace '跳转
        End If
    End If
End If
End Sub

Private Sub Timer2_Timer()
Dim F As Boolean, N%, S%

If Clr = 0 Then
    For i = 8 To 96 Step 8 '删除整排方块
        F = True
        For j = i To i + 7
            If Label2(j).BackStyle = 0 Then F = False: Exit For
        Next j
        If F Then
            For k = i To i + 7
                Label2(k).BackStyle = 0
                Label2(k).BackColor = &HFFFFFF
            Next k
            N = N + 1
        End If
    Next i
    
    S = N * 10 '整排递增计分
    For l = 1 To N - 1
        S = S + l * 5
    Next l
    Text1.Text = Val(Text1.Text) + S
    Clr = 1
ElseIf Clr = 1 Then
    For i = 8 To 96 Step 8 '下落悬空方块
        F = True
        For j = i To i + 7
            If Label2(j).BackStyle = 1 Then F = False: Exit For
        Next j
        If F Then
            For k = i + 7 To 8 Step -1
                Label2(k).BackStyle = Label2(k - 8).BackStyle
                Label2(k).BackColor = Label2(k - 8).BackColor
            Next k
        End If
    Next i
    Clr = 2
ElseIf Clr = 2 Then
    Timer2.Enabled = False '继续循环
    Clr = 0
    rePlace
End If
End Sub

Private Sub Timer3_Timer()
Label4.Caption = Val(Label4.Caption) - 1 '倒计时

If Label4.Caption = "0" Then '继续游戏
    Label4.Caption = "已暂停"
    Timer3.Enabled = False
    Timer1.Enabled = True
    Picture1.Visible = True
End If
End Sub

Private Sub Command1_LostFocus()
Command1.SetFocus '指定焦点
End Sub

Private Sub Frame1_DblClick() '作弊码(？)
If Text1.Text <> "" Then
    Srt = True
    Text1.Text = 999999
End If
End Sub


Private Sub rePlace()
Randomize

Atv = Nxt '获取下个图形

Select Case Atv '方块归位
    Case 1
        Label1(0).Left = 1440
        Label1(1).Left = 1080
        Label1(2).Left = 720
        Label1(3).Left = 1800
        Label1(0).Top = -480
        Label1(1).Top = -480
        Label1(2).Top = -480
        Label1(3).Top = -480
        Label1(0).BackColor = &H8080FF
        Label1(1).BackColor = &H8080FF
        Label1(2).BackColor = &H8080FF
        Label1(3).BackColor = &H8080FF
    Case 2
        Label1(0).Left = 1440
        Label1(1).Left = 1440
        Label1(2).Left = 1080
        Label1(3).Left = 1800
        Label1(0).Top = -480
        Label1(1).Top = -840
        Label1(2).Top = -480
        Label1(3).Top = -480
        Label1(0).BackColor = &H80C0FF
        Label1(1).BackColor = &H80C0FF
        Label1(2).BackColor = &H80C0FF
        Label1(3).BackColor = &H80C0FF
    Case 3
        Label1(0).Left = 1440
        Label1(1).Left = 1080
        Label1(2).Left = 1080
        Label1(3).Left = 1800
        Label1(0).Top = -480
        Label1(1).Top = -840
        Label1(2).Top = -480
        Label1(3).Top = -480
        Label1(0).BackColor = &H80FFFF
        Label1(1).BackColor = &H80FFFF
        Label1(2).BackColor = &H80FFFF
        Label1(3).BackColor = &H80FFFF
    Case 4
        Label1(0).Left = 1440
        Label1(1).Left = 1800
        Label1(2).Left = 1800
        Label1(3).Left = 1080
        Label1(0).Top = -480
        Label1(1).Top = -840
        Label1(2).Top = -480
        Label1(3).Top = -480
        Label1(0).BackColor = &H80FF80
        Label1(1).BackColor = &H80FF80
        Label1(2).BackColor = &H80FF80
        Label1(3).BackColor = &H80FF80
    Case 5
        Label1(0).Left = 1080
        Label1(1).Left = 1440
        Label1(2).Left = 1080
        Label1(3).Left = 1440
        Label1(0).Top = -840
        Label1(1).Top = -840
        Label1(2).Top = -480
        Label1(3).Top = -480
        Label1(0).BackColor = &HFFFF80
        Label1(1).BackColor = &HFFFF80
        Label1(2).BackColor = &HFFFF80
        Label1(3).BackColor = &HFFFF80
    Case 6
        Label1(0).Left = 1440
        Label1(1).Left = 1440
        Label1(2).Left = 1080
        Label1(3).Left = 1800
        Label1(0).Top = -840
        Label1(1).Top = -480
        Label1(2).Top = -840
        Label1(3).Top = -480
        Label1(0).BackColor = &HFF8080
        Label1(1).BackColor = &HFF8080
        Label1(2).BackColor = &HFF8080
        Label1(3).BackColor = &HFF8080
    Case 7
        Label1(0).Left = 1440
        Label1(1).Left = 1440
        Label1(2).Left = 1800
        Label1(3).Left = 1080
        Label1(0).Top = -840
        Label1(1).Top = -480
        Label1(2).Top = -840
        Label1(3).Top = -480
        Label1(0).BackColor = &HFF80FF
        Label1(1).BackColor = &HFF80FF
        Label1(2).BackColor = &HFF80FF
        Label1(3).BackColor = &HFF80FF
End Select

Nxt = Int(Rnd * 7 + 1) '随机下个图形

Select Case Nxt '图形预览
    Case 1
        Label3(0).Left = 240
        Label3(0).Top = 480
        Label3(0).Width = 240
        Label3(0).Height = 120
        Label3(0).BackColor = &H8080FF
        Label3(1).Left = 480
        Label3(1).Top = 480
        Label3(1).Width = 240
        Label3(1).Height = 120
        Label3(1).BackColor = &H8080FF
    Case 2
        Label3(0).Left = 420
        Label3(0).Top = 420
        Label3(0).Width = 120
        Label3(0).Height = 120
        Label3(0).BackColor = &H80C0FF
        Label3(1).Left = 300
        Label3(1).Top = 540
        Label3(1).Width = 360
        Label3(1).Height = 120
        Label3(1).BackColor = &H80C0FF
    Case 3
        Label3(0).Left = 300
        Label3(0).Top = 420
        Label3(0).Width = 120
        Label3(0).Height = 120
        Label3(0).BackColor = &H80FFFF
        Label3(1).Left = 300
        Label3(1).Top = 540
        Label3(1).Width = 360
        Label3(1).Height = 120
        Label3(1).BackColor = &H80FFFF
    Case 4
        Label3(0).Left = 540
        Label3(0).Top = 420
        Label3(0).Width = 120
        Label3(0).Height = 120
        Label3(0).BackColor = &H80FF80
        Label3(1).Left = 300
        Label3(1).Top = 540
        Label3(1).Width = 360
        Label3(1).Height = 120
        Label3(1).BackColor = &H80FF80
    Case 5
        Label3(0).Left = 360
        Label3(0).Top = 420
        Label3(0).Width = 240
        Label3(0).Height = 120
        Label3(0).BackColor = &HFFFF80
        Label3(1).Left = 360
        Label3(1).Top = 540
        Label3(1).Width = 240
        Label3(1).Height = 120
        Label3(1).BackColor = &HFFFF80
    Case 6
        Label3(0).Left = 300
        Label3(0).Top = 420
        Label3(0).Width = 240
        Label3(0).Height = 120
        Label3(0).BackColor = &HFF8080
        Label3(1).Left = 420
        Label3(1).Top = 540
        Label3(1).Width = 240
        Label3(1).Height = 120
        Label3(1).BackColor = &HFF8080
    Case 7
        Label3(0).Left = 420
        Label3(0).Top = 420
        Label3(0).Width = 240
        Label3(0).Height = 120
        Label3(0).BackColor = &HFF80FF
        Label3(1).Left = 300
        Label3(1).Top = 540
        Label3(1).Width = 240
        Label3(1).Height = 120
        Label3(1).BackColor = &HFF80FF
End Select

Drt = 0 '刷新变量
Mve = True

Timer1.Enabled = True '开始下落
End Sub

Private Sub Rotate()
Dim F As Boolean, S%

F = Mve '初始可用性

Select Case Atv '分流形状
    Case 1
        Select Case Drt Mod 2 '分流方向
            Case 0
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top - 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top - 720 And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top + 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(1).Left = Label1(0).Left
                    Label1(1).Top = Label1(0).Top - 360
                    Label1(2).Left = Label1(0).Left
                    Label1(2).Top = Label1(0).Top - 720
                    Label1(3).Left = Label1(0).Left
                    Label1(3).Top = Label1(0).Top + 360
                End If
            Case 1
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left - 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(0).Left - 720 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(0).Left + 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(1).Left = Label1(0).Left - 360
                    Label1(1).Top = Label1(0).Top
                    Label1(2).Left = Label1(0).Left - 720
                    Label1(2).Top = Label1(0).Top
                    Label1(3).Left = Label1(0).Left + 360
                    Label1(3).Top = Label1(0).Top
                End If
        End Select
    Case 2
        Select Case Drt Mod 4
            Case 0
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top + 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(2).Left = Label1(0).Left + 360
                    Label1(3).Left = Label1(0).Left
                    Label1(3).Top = Label1(0).Top + 360
                End If
            Case 1
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left - 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(1).Top = Label1(0).Top + 360
                    Label1(3).Left = Label1(0).Left - 360
                    Label1(3).Top = Label1(0).Top
                End If
            Case 2
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top - 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(2).Left = Label1(0).Left - 360
                    Label1(3).Left = Label1(0).Left
                    Label1(3).Top = Label1(0).Top - 360
                End If
            Case 3
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top - 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(1).Top = Label1(0).Top - 360
                    Label1(3).Left = Label1(0).Left + 360
                    Label1(3).Top = Label1(0).Top
                End If
        End Select
    Case 3
        Select Case Drt Mod 4
            Case 0
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top - 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top - 720 And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(3).Left And Label2(i).Top = Label1(3).Top - 720 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(1).Left = Label1(0).Left
                    Label1(2).Top = Label1(0).Top - 720
                    Label1(2).Left = Label1(0).Left
                    Label1(3).Top = Label1(0).Top - 720
                End If
            Case 1
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left + 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(1).Left - 360 And Label2(i).Top = Label1(1).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(1).Left + 360 And Label2(i).Top = Label1(1).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(0).Left = Label1(1).Left + 360
                    Label1(2).Left = Label1(1).Left - 360
                    Label1(2).Top = Label1(1).Top
                    Label1(3).Top = Label1(1).Top
                End If
            Case 2
                For i = 0 To 111
                    If Label2(i).Left = Label1(1).Left And Label2(i).Top = Label1(1).Top + 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(1).Left And Label2(i).Top = Label1(1).Top - 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(2).Left And Label2(i).Top = Label1(2).Top + 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(0).Left = Label1(1).Left
                    Label1(2).Top = Label1(1).Top + 360
                    Label1(3).Left = Label1(1).Left
                    Label1(3).Top = Label1(1).Top - 360
                End If
            Case 3
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left + 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(1).Left - 360 And Label2(i).Top = Label1(1).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    
                    Label1(1).Left = Label1(2).Left
                    Label1(3).Left = Label1(0).Left + 360
                    Label1(3).Top = Label1(0).Top
                End If
        End Select
    Case 4
        Select Case Drt Mod 4
            Case 0
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top - 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(0).Left And Label2(i).Top = Label1(0).Top - 720 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(1).Left = Label1(0).Left
                    Label1(3).Left = Label1(0).Left
                    Label1(3).Top = Label1(0).Top - 720
                End If
            Case 1
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left - 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(1).Left + 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(1).Left - 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(0).Left = Label1(1).Left - 360
                    Label1(2).Top = Label1(1).Top
                    Label1(3).Left = Label1(1).Left - 360
                    Label1(3).Top = Label1(1).Top
                End If
            Case 2
                For i = 0 To 111
                    If Label2(i).Left = Label1(1).Left And Label2(i).Top = Label1(1).Top + 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(1).Left And Label2(i).Top = Label1(1).Top - 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(3).Left And Label2(i).Top = Label1(3).Top - 360 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(0).Left = Label1(1).Left
                    Label1(2).Left = Label1(1).Left
                    Label1(2).Top = Label1(1).Top - 360
                    Label1(3).Top = Label1(1).Top - 360
                End If
            Case 3
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left + 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(0).Left - 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(1).Left + 360 And Label2(i).Top = Label1(1).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(1).Left = Label1(0).Left + 360
                    Label1(2).Left = Label1(0).Left + 360
                    Label1(2).Top = Label1(0).Top
                    Label1(3).Top = Label1(0).Top
                End If
        End Select
    Case 6
        Select Case Drt Mod 2
            Case 0
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left + 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(3).Left And Label2(i).Top = Label1(3).Top - 720 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(2).Left = Label1(0).Left + 360
                    Label1(3).Top = Label1(0).Top - 360
                End If
            Case 1
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left - 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(3).Left And Label2(i).Top = Label1(3).Top + 720 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(2).Left = Label1(0).Left - 360
                    Label1(3).Top = Label1(0).Top + 360
                End If
        End Select
    Case 7
        Select Case Drt Mod 2
            Case 0
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left - 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(3).Left And Label2(i).Top = Label1(3).Top - 720 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(2).Left = Label1(0).Left - 360
                    Label1(3).Top = Label1(0).Top - 360
                End If
            Case 1
                For i = 0 To 111
                    If Label2(i).Left = Label1(0).Left + 360 And Label2(i).Top = Label1(0).Top And Label2(i).BackStyle = 1 Then F = False: Exit For
                    If Label2(i).Left = Label1(3).Left And Label2(i).Top = Label1(3).Top + 720 And Label2(i).BackStyle = 1 Then F = False: Exit For
                Next i
                If F Then
                    Label1(2).Left = Label1(0).Left + 360
                    Label1(3).Top = Label1(0).Top + 360
                End If
        End Select
End Select

If F Then
    Drt = Drt + 1 '方向累加
    Do '防出界
        S = 0
        For i = 0 To 3
            If Label1(i).Left < 0 Then S = 1
            If Label1(i).Left > 2520 Then S = -1
        Next i
        For j = 0 To 3
            Label1(j).Left = Label1(j).Left + S * 360
        Next j
    Loop Until S = 0
End If
End Sub
