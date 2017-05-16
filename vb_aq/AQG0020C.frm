VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQG0020C 
   Caption         =   "板坯称重申请单编辑"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text21 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10830
      TabIndex        =   22
      Top             =   1710
      Width           =   1305
   End
   Begin VB.TextBox Text20 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10830
      TabIndex        =   21
      Top             =   1380
      Width           =   1305
   End
   Begin VB.TextBox Text19 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10830
      TabIndex        =   20
      Top             =   1050
      Width           =   1305
   End
   Begin VB.TextBox Text18 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10830
      TabIndex        =   19
      Top             =   720
      Width           =   1305
   End
   Begin VB.TextBox Text17 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10830
      TabIndex        =   18
      Top             =   390
      Width           =   1305
   End
   Begin VB.TextBox Text16 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10830
      TabIndex        =   17
      Top             =   60
      Width           =   1305
   End
   Begin VB.TextBox Text15 
      Height          =   315
      Left            =   7290
      TabIndex        =   16
      Top             =   1050
      Width           =   1305
   End
   Begin VB.TextBox Text14 
      Height          =   315
      Left            =   6150
      TabIndex        =   15
      Top             =   1710
      Width           =   795
   End
   Begin VB.TextBox Text13 
      Height          =   315
      Left            =   6150
      TabIndex        =   14
      Top             =   1380
      Width           =   2445
   End
   Begin VB.TextBox Text12 
      Height          =   315
      Left            =   6150
      TabIndex        =   13
      Top             =   1050
      Width           =   1125
   End
   Begin VB.TextBox Text11 
      Height          =   315
      Left            =   6150
      TabIndex        =   12
      Top             =   720
      Width           =   1755
   End
   Begin VB.TextBox Text10 
      Height          =   315
      Left            =   6150
      TabIndex        =   11
      Top             =   390
      Width           =   1755
   End
   Begin VB.TextBox Text9 
      Height          =   315
      Left            =   6150
      TabIndex        =   10
      Top             =   60
      Width           =   1755
   End
   Begin VB.TextBox Text8 
      Height          =   315
      Left            =   1530
      TabIndex        =   9
      Top             =   1710
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   1530
      TabIndex        =   8
      Top             =   1380
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   2190
      TabIndex        =   7
      Top             =   1050
      Width           =   1785
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1530
      TabIndex        =   6
      Top             =   1050
      Width           =   645
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   2190
      TabIndex        =   5
      Top             =   720
      Width           =   1785
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1530
      TabIndex        =   4
      Top             =   720
      Width           =   645
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1530
      TabIndex        =   3
      Top             =   390
      Width           =   2445
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1530
      TabIndex        =   2
      Top             =   60
      Width           =   2445
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   2925
      Left            =   90
      TabIndex        =   1
      Top             =   5880
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   5159
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
      MaxCols         =   11
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQG0020C.frx":0000
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2775
      Left            =   90
      TabIndex        =   0
      Top             =   2520
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   4895
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
      MaxCols         =   15
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQG0020C.frx":047E
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   90
      Top             =   60
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   90
      Top             =   390
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "车辆号"
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
      Left            =   90
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "目的仓库"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   90
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "申请单明细"
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
      Left            =   90
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "可选板坯"
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
      Left            =   90
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "启运仓库"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   90
      Top             =   1380
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "板坯数量"
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
      Left            =   4710
      Top             =   60
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   4710
      Top             =   390
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   4710
      Top             =   1050
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   4710
      Top             =   720
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   4710
      Top             =   1380
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "订单号"
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   90
      Top             =   1710
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "总重量"
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   4710
      Top             =   1710
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "订单序列号"
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
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   9360
      Top             =   60
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "录入日期"
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   9360
      Top             =   390
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "录入人"
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   9360
      Top             =   1050
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "修改人"
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   9360
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "修改日期"
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
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   9360
      Top             =   1380
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "称重日期"
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
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   9360
      Top             =   1710
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "称重人"
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
   Begin InDate.ULabel ULabel21 
      Height          =   315
      Left            =   3480
      Top             =   1380
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      Caption         =   "块"
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
   Begin InDate.ULabel ULabel22 
      Height          =   315
      Left            =   3480
      Top             =   1710
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      Caption         =   "吨"
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
   Begin InDate.ULabel ULabel23 
      Height          =   315
      Left            =   7920
      Top             =   720
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   "mm"
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
   Begin InDate.ULabel ULabel24 
      Height          =   315
      Left            =   7920
      Top             =   390
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   "mm"
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
   Begin InDate.ULabel ULabel25 
      Height          =   315
      Left            =   7920
      Top             =   60
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Caption         =   "mm"
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
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   18800
      Y1              =   5460
      Y2              =   5460
   End
   Begin VB.Line Line3 
      X1              =   -30
      X2              =   18770
      Y1              =   5430
      Y2              =   5430
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   -60
      X2              =   18740
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Line Line1 
      X1              =   -1230
      X2              =   17570
      Y1              =   2070
      Y2              =   2070
   End
End
Attribute VB_Name = "AQG0020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
