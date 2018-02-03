VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10110
   ClientLeft      =   2985
   ClientTop       =   2460
   ClientWidth     =   8250
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   20.25
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00808000&
   Icon            =   "workspace.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   8250
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3720
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   4560
      Width           =   1212
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3960
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   3840
      Top             =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   3960
      TabIndex        =   0
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   150
      Left            =   0
      Top             =   0
      Width           =   492
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   0
      Left            =   3480
      Top             =   6960
      Width           =   252
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   1
      Left            =   360
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   2
      Left            =   720
      Shape           =   1  'Square
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   3
      Left            =   1080
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   4
      Left            =   1440
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   5
      Left            =   1800
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   6
      Left            =   2160
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   7
      Left            =   2520
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   8
      Left            =   0
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   9
      Left            =   360
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   10
      Left            =   720
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   11
      Left            =   1080
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   12
      Left            =   1440
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   13
      Left            =   1800
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   14
      Left            =   2160
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   15
      Left            =   2520
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   16
      Left            =   2880
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   17
      Left            =   3240
      Top             =   0
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   18
      Left            =   2880
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   19
      Left            =   3240
      Top             =   360
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   20
      Left            =   0
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   21
      Left            =   360
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   22
      Left            =   720
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   23
      Left            =   1080
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   24
      Left            =   1440
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   25
      Left            =   1800
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   26
      Left            =   2160
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   27
      Left            =   2520
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   28
      Left            =   0
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   29
      Left            =   360
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   30
      Left            =   720
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   31
      Left            =   1080
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   32
      Left            =   1440
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   33
      Left            =   1800
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   34
      Left            =   2160
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   35
      Left            =   2520
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   36
      Left            =   2880
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   37
      Left            =   3240
      Top             =   720
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   38
      Left            =   2880
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   39
      Left            =   3240
      Top             =   1080
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   40
      Left            =   0
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   41
      Left            =   360
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   42
      Left            =   720
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   43
      Left            =   1080
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   44
      Left            =   1440
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   45
      Left            =   1800
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   46
      Left            =   2160
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   47
      Left            =   2520
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   48
      Left            =   0
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   49
      Left            =   360
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   50
      Left            =   720
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   51
      Left            =   1080
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   52
      Left            =   1440
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   53
      Left            =   1800
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   54
      Left            =   2160
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   55
      Left            =   2520
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   56
      Left            =   2880
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   57
      Left            =   3240
      Top             =   1440
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   58
      Left            =   2880
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   59
      Left            =   3240
      Top             =   1800
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   60
      Left            =   0
      Top             =   2160
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   61
      Left            =   360
      Top             =   2160
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   62
      Left            =   720
      Top             =   2160
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   63
      Left            =   1080
      Top             =   2160
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   64
      Left            =   1440
      Top             =   2160
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   65
      Left            =   1800
      Top             =   2760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   66
      Left            =   2160
      Top             =   2160
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   67
      Left            =   2520
      Top             =   2160
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   68
      Left            =   0
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   69
      Left            =   360
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   70
      Left            =   720
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   71
      Left            =   1080
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   72
      Left            =   1440
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   73
      Left            =   1800
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   74
      Left            =   2160
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   75
      Left            =   2520
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   76
      Left            =   2880
      Top             =   2160
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   77
      Left            =   3240
      Top             =   2160
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   78
      Left            =   2880
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   79
      Left            =   3240
      Top             =   2520
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   80
      Left            =   0
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   81
      Left            =   360
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   82
      Left            =   720
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   83
      Left            =   1080
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   84
      Left            =   1440
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   85
      Left            =   1800
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   86
      Left            =   2160
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   87
      Left            =   2520
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   88
      Left            =   0
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   89
      Left            =   360
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   90
      Left            =   720
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   91
      Left            =   1080
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   92
      Left            =   1440
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   93
      Left            =   1800
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   94
      Left            =   2160
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   95
      Left            =   2520
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   96
      Left            =   2880
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   97
      Left            =   3240
      Top             =   2880
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   98
      Left            =   2880
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   99
      Left            =   3240
      Top             =   3240
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   100
      Left            =   3240
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   101
      Left            =   2880
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   102
      Left            =   3240
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   103
      Left            =   2880
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   104
      Left            =   2520
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   105
      Left            =   2160
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   106
      Left            =   1800
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   107
      Left            =   1440
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   108
      Left            =   1080
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   109
      Left            =   720
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   110
      Left            =   360
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   111
      Left            =   0
      Top             =   5040
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   112
      Left            =   2520
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   113
      Left            =   2160
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   114
      Left            =   1800
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   115
      Left            =   1440
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   116
      Left            =   1080
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   117
      Left            =   720
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   118
      Left            =   360
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   119
      Left            =   0
      Top             =   4680
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   120
      Left            =   3240
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   121
      Left            =   2880
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   122
      Left            =   3240
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   123
      Left            =   2880
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   124
      Left            =   2520
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   125
      Left            =   2160
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   126
      Left            =   1800
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   127
      Left            =   1440
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   128
      Left            =   1080
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   129
      Left            =   720
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   130
      Left            =   360
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   131
      Left            =   0
      Top             =   4320
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   132
      Left            =   2520
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   133
      Left            =   2160
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   134
      Left            =   1800
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   135
      Left            =   1440
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   136
      Left            =   1080
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   137
      Left            =   720
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   138
      Left            =   360
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   139
      Left            =   0
      Top             =   3960
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   140
      Left            =   3240
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   141
      Left            =   2880
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   142
      Left            =   2520
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   143
      Left            =   2160
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   144
      Left            =   1800
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   145
      Left            =   1440
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   146
      Left            =   1080
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   147
      Left            =   720
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   148
      Left            =   360
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   149
      Left            =   0
      Top             =   3600
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   151
      Left            =   2880
      Top             =   6840
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   152
      Left            =   3240
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   153
      Left            =   2880
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   154
      Left            =   2520
      Top             =   6840
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   155
      Left            =   2160
      Top             =   6840
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   156
      Left            =   1800
      Top             =   6840
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   157
      Left            =   1440
      Top             =   6840
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   158
      Left            =   1080
      Top             =   6840
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   159
      Left            =   720
      Top             =   6840
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   160
      Left            =   360
      Top             =   6840
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   161
      Left            =   0
      Top             =   6840
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   162
      Left            =   2520
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   163
      Left            =   2160
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   164
      Left            =   1800
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   165
      Left            =   1440
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   166
      Left            =   1080
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   167
      Left            =   720
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   168
      Left            =   360
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   169
      Left            =   0
      Top             =   6480
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   170
      Left            =   3240
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   171
      Left            =   2880
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   172
      Left            =   3240
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   173
      Left            =   2880
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   174
      Left            =   2520
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   175
      Left            =   2160
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   176
      Left            =   1800
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   177
      Left            =   1440
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   178
      Left            =   1080
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   179
      Left            =   720
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   180
      Left            =   360
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   181
      Left            =   0
      Top             =   6120
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   182
      Left            =   2520
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   183
      Left            =   2160
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   184
      Left            =   1800
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   185
      Left            =   1440
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   186
      Left            =   1080
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   187
      Left            =   720
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   188
      Left            =   360
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   189
      Left            =   0
      Top             =   5760
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   190
      Left            =   3240
      Shape           =   1  'Square
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   191
      Left            =   2880
      Shape           =   1  'Square
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   192
      Left            =   2520
      Shape           =   1  'Square
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   193
      Left            =   2160
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   194
      Left            =   1800
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   195
      Left            =   1440
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   196
      Left            =   1080
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   197
      Left            =   720
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   198
      Left            =   360
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape1 
      Height          =   372
      Index           =   199
      Left            =   0
      Top             =   5400
      Width           =   372
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   492
      Index           =   0
      Left            =   3960
      Top             =   2280
      Width           =   492
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   492
      Index           =   1
      Left            =   3960
      Top             =   2760
      Width           =   492
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   492
      Index           =   2
      Left            =   3960
      Top             =   3240
      Width           =   492
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   492
      Index           =   3
      Left            =   3960
      Top             =   3720
      Width           =   492
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB语言版俄罗斯方块
'Totoo、Aoo34智造（一个人的两个名字），一些方块，很多计算

Const WN As Integer = 10, HN As Integer = 20, StrBig As Integer = 300
Const Boxl As Integer = 372, BoxNum As Integer = 200




Private Sub Combo1_DropDown()
Turn
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = TimeLen
CheckTop
Fail
Cleaner
XFull
End Sub

Private Sub Form_Load()
    For a = 0 To 3
With Shape2(a)
.Width = Boxl
.Height = Boxl
End With
    Next a
With Label2
.Height = StrBig
.Width = StrBig * 8
End With
Label2.Move StrBig * 8, Boxl * 20
With Combo1
.Width = StrBig * 8
End With
Combo1.Move 0, Boxl * 20
Form1.Caption = "w,a,s,d分别为变形、左、右及降落"

TimeLen = 200
Timer1.Interval = 1000

Call ClearUpEr

ShapeAdd
    
End Sub
 
Private Sub ClearUpEr()
'Totoo作品
With Form1
.Width = WN * Boxl + 200
.Height = HN * Boxl + Combo1.Height + 400
End With
    Dim Ia As Integer, ib As Integer
    Dim x(BoxNum) As Integer, y(BoxNum) As Integer
    x(1) = 0
    y(1) = 0
        For a = 0 To 199
With Shape1(a)
.Width = Boxl * (Iret + 1)
.Height = Boxl * (Iret + 1)
End With
    Ia = Ia + 1
        If (Ia <> 0) And (a Mod WN = 0) Then Ia = 0: ib = ib + 1
    x(a) = Boxl * Ia
    y(a) = Boxl * (ib - 1)
    Shape1(a).Move x(a), y(a)
        Next a
'Totoo作品
End Sub

Sub ShapeAdd()
'Totoo作品
Dim Sret As Integer
x(1) = 0: y(1) = 0
        For j = 2 To 4
        If j = 4 Then
            If x(3) = 1 And y(3) = 1 Then
                        Rndget Sret, 2
            If Sret = 0 Then GoTo Four:
            End If
        End If
    Rndget Sret, 2
    If Sret = 1 Then
        Sret = j
        NextBox Sret, Sret - 1, 1, 1
    Else
        Sret = j
        NextBox Sret, Sret - 1, 1, 0
    End If

        Next j
        
If 1 = 2 Then
Four:
Rndget Sret, 2
Select Case x(2)
    Case 1:
            If Sret = 1 Then
            NextBox 4, 2, 1, 1
            Else
            NextBox 4, 3, -1, 1
            End If
    Case 0:
            If Sret = 1 Then
            NextBox 4, 2, 1, 0
            Else
            NextBox 4, 3, -1, 0
            End If
End Select
End If
initialize:
        For a = 1 To 4
With Shape2(a - 1)
.Move x(a) * Boxl, y(a) * Boxl
.Width = Boxl
.Height = Boxl
End With
        Next a
corect:
    Dim reta3, reta4 As Integer
        For a = 1 To 4
    reta3 = x(a)
        If reta3 > reta4 Then: reta4 = reta3
        Next a
    Randomize
    reta3 = Fix(Rnd * (9 - reta4)) + 1
        For a = 1 To 4
    x(a) = x(a) + reta3
        Next a
'Totoo作品

End Sub

Sub Cleaner()
'Totoo作品，中国智造
    For a = 1 To 10
        For b = 1 To 20
            If BF(a, b) = 1 Then
Shape1(a + (b - 1) * 10 - 1).FillStyle = 0
            Else
Shape1(a + (b - 1) * 10 - 1).FillStyle = 1
            End If
        Next b
    Next a

End Sub


Sub CheckTop()
    'Totoo作品，中国智造
On Error GoTo done:
        For a = 1 To 4
    If x(a) + 1 < 19 Then On Error Resume Next
    If y(a) > 18 Then GoTo done:
    If BF(x(a) + 1, y(a) + 2) = 1 Then GoTo done:

On Error GoTo Over:
    If x(a) + 1 > 20 Or x(a) + 1 < 1 Then GoTo Over:
        Next a
    If 1 = 2 Then
Over:
    Call ClsBox
        'Timelen = 500
        Call ShapeAdd
        'MsgBox "GameOver!": End
    End If
    If 1 = 2 Then
done:
        For a = 1 To 4
            If BF(x(a) + 1, y(a) + 1) = 1 Then GoTo Over:
        Next a
        For a = 1 To 4
    BF(x(a) + 1, y(a) + 1) = 1
        Next a
    Call ShapeAdd: If BottomAsk = True Then TimeLen = 500: BottomAsk = False
    End If
Pass:
End Sub

Private Sub Turn()
    Dim ret As Integer
    
        For a = 1 To 4
        
            ret = x(a) - x(3): mY(a) = ret + y(3)
            ret = y(a) - y(3): mX(a) = -ret + x(3)
        
        Next a
    
ComeTure
End Sub

Sub XFull() 'Totoo作品，中国智造
    Dim Ia As Integer, I As Integer
    Dim mY As Integer, BfRet(1 To 10, 1 To 20) As Integer
    Dim Cleanit As Boolean
        For b = 1 To 20
            For a = 1 To 10
                If BF(a, b) = 1 Then Ia = Ia + 1
            Next a
                If Ia = 10 Then I = I + 1: Toper(I) = b:  '记录满格
    Ia = 0
        Next b
    If I <> 0 Then
        For b = 1 To I
            For a = 1 To 10
        BF(a, Toper(b)) = 0
            Next a
socre = socre + 1
            Next b
Label2.Caption = "完成：" & Str(socre)
    End If
    If (Clean = True) Then
        For a = 1 To 10
    Cleanit = False
            For b = 1 To 20
        mY = 0
        mY = BF(a, b)
        If BF(a, b) = 1 Then
                For c = 1 To I
            If Toper(c) <> 0 Then
                If b < Toper(c) Then
                mY = mY + 1
                Cleanit = True
                End If
            End If
            If c = I Then
                If b + mY > 20 Then GoTo Pass:
            BfRet(a, b + mY - 1) = 1
                If 1 = 2 Then
Pass:
                For d = 1 To 10
                BfRet(a, 20) = 1
                Next d
                End If
        End If
    Next c
    End If
    mY = 0
    Next b
    If Cleanit = True Then
    For b = 1 To 20
    BF(a, b) = BfRet(a, b)
    BfRet(a, b) = 0
    Next b
    End If
Next a
End If
    For L = 1 To I
    Toper(L) = 0
    Next L
End Sub

Private Sub Save()
    Dim SFN As String
    CommonDialog1.ShowOpen
    SFN = CommonDialog1.FileName
    If SFN <> "" Then
    Open SFN & ".totooDat" For Output As #1
    For a = 1 To 10
    For b = 1 To 20
    Print #1, BF(a, b)
    Next b, a
    Print socre
    Close #1
    End If
End Sub


Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
        Case 65, 37: MoveLeft
        Case 68, 39: MoveRight
        Case 87, 38: Turn
        Case 83, 40: TimeLen = 20: BottomAsk = True
        End Select
    If KeyCode = 13 Then
        EntI = EntI + 1
            If EntI Mod 2 = 1 Then
            TimeLen = 10
            Else: TimeLen = 1000: End If
    End If
End Sub

Private Sub Fail()
    Clean = True
        For a = 1 To 4
    y(a) = y(a) + 1
Shape2(a - 1).Move x(a) * Boxl, y(a) * Boxl
        Next a
End Sub


