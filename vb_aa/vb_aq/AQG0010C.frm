VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQG0010C 
   Caption         =   "板坯称重申请单查询及确认"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text12 
      Height          =   345
      Left            =   5880
      TabIndex        =   19
      Top             =   1590
      Width           =   2235
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   1560
      TabIndex        =   15
      Top             =   1560
      Width           =   3375
      Begin Threed.SSOption SSOption1 
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "待称重"
      End
      Begin Threed.SSOption SSOption2 
         Height          =   315
         Left            =   1065
         TabIndex        =   17
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "待修改"
      End
      Begin Threed.SSOption SSOption3 
         Height          =   315
         Left            =   2070
         TabIndex        =   18
         Top             =   0
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   196609
         Caption         =   "已确认"
      End
   End
   Begin VB.TextBox Text11 
      Height          =   315
      Left            =   7320
      TabIndex        =   14
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox Text10 
      Height          =   315
      Left            =   3000
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   315
      Left            =   1560
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   315
      Left            =   8880
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   8880
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   8880
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   7320
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   7320
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   7320
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   3375
   End
   Begin InDate.UDate UDate2 
      Height          =   315
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.74
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.UDate UDate1 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.74
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   6885
      Left            =   9630
      TabIndex        =   1
      Top             =   2130
      Width           =   5505
      _Version        =   393216
      _ExtentX        =   9710
      _ExtentY        =   12144
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   6
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQG0010C.frx":0000
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   6885
      Left            =   120
      TabIndex        =   0
      Top             =   2130
      Width           =   9495
      _Version        =   393216
      _ExtentX        =   16748
      _ExtentY        =   12144
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQG0010C.frx":0349
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "申请单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   120
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "申请单日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   120
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "车号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   5880
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "厚度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   5880
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "宽度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   5880
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "长度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   120
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "钢种"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   5880
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "板坯号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   120
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "状态"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   18800
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   18800
      Y1              =   1980
      Y2              =   1980
   End
End
Attribute VB_Name = "AQG0010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
