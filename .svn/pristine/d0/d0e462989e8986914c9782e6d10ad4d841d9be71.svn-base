VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQB0150C 
   Caption         =   "交付条件设计结果查询 - AQB0150C"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   1665
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.OptionButton opt_KND 
      BackColor       =   &H00E0E0E0&
      Caption         =   "客户特殊要求"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   11145
      TabIndex        =   41
      Top             =   1570
      Width           =   1500
   End
   Begin VB.TextBox txt_Design_STS 
      Height          =   315
      Left            =   7770
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt_RECT_NAME 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7470
      Width           =   975
   End
   Begin VB.TextBox txt_FLT_NAME 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6225
      Width           =   975
   End
   Begin VB.TextBox txt_STRT_NAME 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5820
      Width           =   975
   End
   Begin VB.TextBox txt_RECT_TYP 
      Height          =   300
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   22
      Top             =   7470
      Width           =   345
   End
   Begin VB.TextBox txt_FLT_TYP 
      Height          =   300
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   21
      Top             =   6225
      Width           =   345
   End
   Begin VB.TextBox txt_STRT_KND 
      Height          =   300
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   20
      Top             =   5820
      Width           =   345
   End
   Begin VB.TextBox txt_LEN_TOL_UNIT 
      Height          =   300
      Left            =   14880
      TabIndex        =   13
      Top             =   2640
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt_WID_TOL_UNIT 
      Height          =   300
      Left            =   14880
      TabIndex        =   12
      Top             =   2160
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt_THK_TOL_UNIT 
      Height          =   300
      Left            =   14880
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.OptionButton opt_KND 
      BackColor       =   &H00E0E0E0&
      Caption         =   "综合交付条件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   8055
      TabIndex        =   10
      Top             =   1570
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.OptionButton opt_KND 
      BackColor       =   &H00E0E0E0&
      Caption         =   "标准交付条件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   9600
      TabIndex        =   9
      Top             =   1570
      Width           =   1500
   End
   Begin VB.TextBox txt_SHAPE_GRD_NAME 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   2730
      MaxLength       =   80
      TabIndex        =   8
      Top             =   1550
      Width           =   2295
   End
   Begin VB.TextBox txt_SURF_GRD_NAME 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   2745
      MaxLength       =   80
      TabIndex        =   7
      Top             =   1970
      Width           =   2295
   End
   Begin VB.TextBox txt_DEV_STD_CD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   310
      Left            =   2160
      MaxLength       =   7
      TabIndex        =   6
      Top             =   2400
      Width           =   2865
   End
   Begin VB.TextBox txt_SHAPE_GRD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1550
      Width           =   555
   End
   Begin VB.TextBox txt_SURF_GRD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   2160
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1970
      Width           =   555
   End
   Begin VB.TextBox txt_KND 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   6600
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "4"
      Top             =   150
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txt_ins_emp 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   6360
      MaxLength       =   7
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt_ORD_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1650
      MaxLength       =   11
      TabIndex        =   1
      Tag             =   "订单号"
      Top             =   135
      Width           =   2025
   End
   Begin VB.TextBox txt_ORD_ITEM 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   5190
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "序列号"
      Top             =   135
      Width           =   1005
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   210
      Top             =   135
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "订单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   3750
      Top             =   135
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "序列号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   210
      Top             =   690
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      Caption         =   "标准代码"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   2
      Left            =   2310
      Top             =   690
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      Caption         =   "发布年度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   3
      Left            =   3210
      Top             =   690
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "品种"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   4
      Left            =   4440
      Top             =   690
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "厚度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   5
      Left            =   5400
      Top             =   690
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "宽度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   6
      Left            =   6360
      Top             =   690
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "长度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   7
      Left            =   7590
      Top             =   690
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "交货日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   8
      Left            =   8790
      Top             =   690
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "客户"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   9
      Left            =   10020
      Top             =   690
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "订单量"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   10
      Left            =   11250
      Top             =   690
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "特殊要求"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   11
      Left            =   12480
      Top             =   690
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   556
      Caption         =   "订单用途"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   12
      Left            =   360
      Top             =   2400
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "交付条件代表标准"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   14
      Left            =   360
      Top             =   1965
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "表面等级"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   15
      Left            =   360
      Top             =   1545
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "形状等级"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel txt_STDSPEC 
      Height          =   345
      Left            =   210
      Top             =   990
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_STDSPEC_YY 
      Height          =   345
      Left            =   2310
      Top             =   990
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_PROD_CD 
      Height          =   345
      Left            =   3210
      Top             =   990
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_ORD_LEN 
      Height          =   345
      Left            =   6360
      Top             =   990
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_DEL_TO_DATE 
      Height          =   345
      Left            =   7590
      Top             =   990
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_CUST_CD 
      Height          =   345
      Left            =   8790
      Top             =   990
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_UNIT_WGT 
      Height          =   345
      Left            =   10020
      Top             =   990
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_CUST_SPEC_NO 
      Height          =   345
      Left            =   11250
      Top             =   990
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_ENDUSE_CD 
      Height          =   345
      Left            =   12480
      Top             =   990
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   330
      Index           =   13
      Left            =   6240
      Top             =   1545
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   582
      Caption         =   "标准分类"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ul_INS_DATE 
      Height          =   300
      Left            =   1650
      Top             =   8220
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      Caption         =   ""
      Alignment       =   0
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   44
      Left            =   360
      Top             =   8220
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "编制日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   45
      Left            =   7800
      Top             =   8220
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "编制人"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   46
      Left            =   3960
      Top             =   8220
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "修改日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   47
      Left            =   11460
      Top             =   8220
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "修改人"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ul_INS_EMP 
      Height          =   300
      Left            =   9075
      Top             =   8220
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      Caption         =   ""
      Alignment       =   0
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ul_UPD_DATE 
      Height          =   300
      Left            =   5235
      Top             =   8220
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      Caption         =   ""
      Alignment       =   0
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ul_UPD_EMP 
      Height          =   300
      Left            =   12750
      Top             =   8220
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      Caption         =   ""
      Alignment       =   0
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   16
      Left            =   360
      Top             =   3090
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "项目"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   20
      Left            =   1680
      Top             =   3090
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "类型"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   21
      Left            =   3120
      Top             =   3090
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "下限值"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   22
      Left            =   4440
      Top             =   3090
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "上限值"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   23
      Left            =   5775
      Top             =   3090
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   529
      Caption         =   "单位"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   17
      Left            =   360
      Top             =   3570
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "公差"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   18
      Left            =   1710
      Top             =   4020
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      Caption         =   "宽度"
      Alignment       =   1
      BackColor       =   12632256
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   19
      Left            =   1710
      Top             =   4470
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      Caption         =   "长度"
      Alignment       =   1
      BackColor       =   12632256
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   48
      Left            =   1710
      Top             =   3570
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      Caption         =   "厚度"
      Alignment       =   1
      BackColor       =   12632256
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_WID_TOL_MIN 
      Height          =   300
      Left            =   3120
      TabIndex        =   14
      Top             =   4030
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_LEN_TOL_MIN 
      Height          =   300
      Left            =   3120
      TabIndex        =   15
      Top             =   4475
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_THK_TOL_MAX 
      Height          =   300
      Left            =   4440
      TabIndex        =   16
      Top             =   3570
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_WID_TOL_MAX 
      Height          =   300
      Left            =   4440
      TabIndex        =   17
      Top             =   4020
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_LEN_TOL_MAX 
      Height          =   300
      Left            =   4440
      TabIndex        =   18
      Top             =   4470
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_THK_TOL_MIN 
      Height          =   315
      Left            =   3120
      TabIndex        =   19
      Top             =   3570
      Width           =   1200
      _Version        =   262145
      _ExtentX        =   2117
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   60
      Left            =   5760
      Top             =   3570
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   61
      Left            =   5760
      Top             =   4020
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   62
      Left            =   5760
      Top             =   4470
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   24
      Left            =   360
      Top             =   6225
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "不平度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   25
      Left            =   360
      Top             =   7050
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "塔形度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   26
      Left            =   360
      Top             =   7470
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "直角度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   39
      Left            =   360
      Top             =   5820
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "镰刀弯"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   41
      Left            =   4320
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      Caption         =   "总上限值"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   42
      Left            =   5520
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      Caption         =   "单位上限值"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   43
      Left            =   6735
      Top             =   5400
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   529
      Caption         =   "单位"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   49
      Left            =   6720
      Top             =   7470
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   55
      Left            =   6720
      Top             =   7050
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   58
      Left            =   6720
      Top             =   6225
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   59
      Left            =   6720
      Top             =   5820
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_RECT_MAX 
      Height          =   300
      Left            =   4320
      TabIndex        =   26
      Top             =   7470
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_STRT_MAX 
      Height          =   300
      Left            =   4320
      TabIndex        =   27
      Top             =   5820
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_FLT_MAX 
      Height          =   300
      Left            =   4320
      TabIndex        =   28
      Top             =   6240
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_TELSCOPE_MAX 
      Height          =   300
      Left            =   4320
      TabIndex        =   29
      Top             =   7050
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_STRT_UNIT_MAX 
      Height          =   300
      Left            =   5520
      TabIndex        =   30
      Top             =   5820
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   1
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_FLT_UNIT_MAX 
      Height          =   300
      Left            =   5520
      TabIndex        =   31
      Top             =   6225
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_TELSCOPE_UNIT_MAX 
      Height          =   300
      Left            =   5520
      TabIndex        =   32
      Top             =   7050
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   63
      Left            =   360
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "项目"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   64
      Left            =   1680
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "类型"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   27
      Left            =   7850
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "同板差"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   28
      Left            =   7850
      Top             =   4530
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "楔形度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   29
      Left            =   7850
      Top             =   5040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "凸度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   30
      Left            =   7850
      Top             =   5550
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "波浪度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   31
      Left            =   9360
      Top             =   6690
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "宽度方向"
      Alignment       =   1
      BackColor       =   12632256
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   32
      Left            =   9360
      Top             =   7170
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "长度方向"
      Alignment       =   1
      BackColor       =   12632256
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   38
      Left            =   9360
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "厚度方向"
      Alignment       =   1
      BackColor       =   12632256
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   40
      Left            =   7850
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "切斜"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   50
      Left            =   13785
      Top             =   4560
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   51
      Left            =   13785
      Top             =   5055
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   52
      Left            =   13785
      Top             =   5565
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   53
      Left            =   13785
      Top             =   6060
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   54
      Left            =   13785
      Top             =   3585
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   56
      Left            =   13785
      Top             =   6675
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   57
      Left            =   13785
      Top             =   7170
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_SAPE_WAVE 
      Height          =   300
      Left            =   12285
      TabIndex        =   33
      Top             =   5550
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_THK_SHR_DRAG_MAX 
      Height          =   300
      Left            =   12285
      TabIndex        =   34
      Top             =   6090
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_WID_SHR_DRAG_MAX 
      Height          =   300
      Left            =   12285
      TabIndex        =   35
      Top             =   6690
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_THK_DVT_MAX 
      Height          =   300
      Left            =   12280
      TabIndex        =   36
      Top             =   3600
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   1
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_VRT_MAX 
      Height          =   300
      Left            =   12280
      TabIndex        =   37
      Top             =   4530
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   1
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_CROWN_MAX 
      Height          =   300
      Left            =   12285
      TabIndex        =   38
      Top             =   5040
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   1
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_LEN_SHR_DRAG_MAX 
      Height          =   300
      Left            =   12285
      TabIndex        =   39
      Top             =   7170
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   33
      Left            =   7850
      Top             =   3090
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "项目"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   34
      Left            =   9360
      Top             =   3090
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "类型"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   35
      Left            =   10800
      Top             =   3090
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "下限值"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   36
      Left            =   12280
      Top             =   3090
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "上限值"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   37
      Left            =   13815
      Top             =   3090
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "单位"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_ORD_THK 
      Height          =   345
      Left            =   4440
      Top             =   990
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel txt_ORD_WID 
      Height          =   345
      Left            =   5400
      Top             =   990
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   46
      Left            =   6240
      Top             =   1965
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "LP板宽度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   47
      Left            =   9720
      Top             =   1965
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      Caption         =   "LP板厚度1/2/3"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   48
      Left            =   6240
      Top             =   2400
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   556
      Caption         =   "LP板长度1/2/3/4/5"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_LP_WID 
      Height          =   315
      Left            =   8040
      TabIndex        =   42
      Tag             =   "订单宽度"
      Top             =   1965
      Width           =   1155
      _Version        =   262145
      _ExtentX        =   2028
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   6
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_LP_THK1 
      Height          =   315
      Left            =   11580
      TabIndex        =   43
      Tag             =   "订单宽度"
      Top             =   1965
      Width           =   900
      _Version        =   262145
      _ExtentX        =   1587
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   6
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_LP_THK3 
      Height          =   315
      Left            =   13380
      TabIndex        =   44
      Tag             =   "订单宽度"
      Top             =   1965
      Width           =   1005
      _Version        =   262145
      _ExtentX        =   1764
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   6
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_LP_THK2 
      Height          =   315
      Left            =   12480
      TabIndex        =   45
      Tag             =   "订单宽度"
      Top             =   1965
      Width           =   900
      _Version        =   262145
      _ExtentX        =   1587
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   6
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN1 
      Height          =   315
      Left            =   8115
      TabIndex        =   46
      Tag             =   "Tot_Wgt"
      Top             =   2400
      Width           =   1245
      _Version        =   262145
      _ExtentX        =   2205
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   8
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN2 
      Height          =   315
      Left            =   9360
      TabIndex        =   47
      Tag             =   "Tot_Wgt"
      Top             =   2400
      Width           =   1245
      _Version        =   262145
      _ExtentX        =   2205
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   8
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN3 
      Height          =   315
      Left            =   10620
      TabIndex        =   48
      Tag             =   "Tot_Wgt"
      Top             =   2400
      Width           =   1245
      _Version        =   262145
      _ExtentX        =   2205
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   8
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN4 
      Height          =   315
      Left            =   11880
      TabIndex        =   49
      Tag             =   "Tot_Wgt"
      Top             =   2400
      Width           =   1245
      _Version        =   262145
      _ExtentX        =   2196
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   8
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ORD_LP_LEN5 
      Height          =   315
      Left            =   13125
      TabIndex        =   50
      Tag             =   "Tot_Wgt"
      Top             =   2400
      Width           =   1245
      _Version        =   262145
      _ExtentX        =   2196
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   8
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   65
      Left            =   7850
      Top             =   4050
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "异板差"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   66
      Left            =   13740
      Top             =   4095
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_DIFF_DVT_MAX 
      Height          =   300
      Left            =   12280
      TabIndex        =   51
      Top             =   4050
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   1
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   67
      Left            =   3120
      Top             =   5400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      Caption         =   "测量长度"
      Alignment       =   1
      BackColor       =   16761024
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   68
      Left            =   360
      Top             =   6645
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      Caption         =   "不平度2"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
      BorderStyle     =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   69
      Left            =   6720
      Top             =   6645
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_FLT_UNIT_MAX2 
      Height          =   300
      Left            =   5520
      TabIndex        =   52
      Top             =   6645
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_FLT_LEN 
      Height          =   300
      Left            =   3120
      TabIndex        =   53
      Top             =   6225
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_FLT_LEN2 
      Height          =   300
      Left            =   3120
      TabIndex        =   54
      Top             =   6645
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   70
      Left            =   1680
      Top             =   4920
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      Caption         =   "对角线"
      Alignment       =   1
      BackColor       =   12632256
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit sdb_DIA_TOL_MIN 
      Height          =   300
      Left            =   3120
      TabIndex        =   55
      Top             =   4920
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_DIA_TOL_MAX 
      Height          =   300
      Left            =   4440
      TabIndex        =   56
      Top             =   4920
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   71
      Left            =   5760
      Top             =   4920
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      Caption         =   "mm"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit SDB_STRT_LENGTH 
      Height          =   300
      Left            =   3120
      TabIndex        =   57
      Top             =   5820
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   2
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   1425
      Index           =   4
      Left            =   240
      Top             =   1440
      Width           =   14565
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   12
      X1              =   12120
      X2              =   12120
      Y1              =   3000
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   11
      X1              =   10680
      X2              =   10680
      Y1              =   3000
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   9
      X1              =   9240
      X2              =   9240
      Y1              =   3000
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   8
      X1              =   1620
      X2              =   1620
      Y1              =   5400
      Y2              =   7920
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   585
      Index           =   3
      Left            =   240
      Top             =   8040
      Width           =   14565
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   2655
      Index           =   2
      Left            =   240
      Top             =   5280
      Width           =   7455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   2265
      Index           =   1
      Left            =   240
      Top             =   3000
      Width           =   7455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   4935
      Index           =   0
      Left            =   7680
      Top             =   3000
      Width           =   7125
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   7
      X1              =   13680
      X2              =   13680
      Y1              =   3000
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   5
      X1              =   4200
      X2              =   4200
      Y1              =   5400
      Y2              =   7920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   3060
      X2              =   3060
      Y1              =   3000
      Y2              =   5160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   1620
      X2              =   1620
      Y1              =   3000
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   14760
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   330
      Index           =   1
      Left            =   8010
      Top             =   1545
      Width           =   4695
   End
End
Attribute VB_Name = "AQB0150C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量设计
'-- Program Name      交付条件设计结果修改及查询
'-- Program ID        AQB0150C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.08.21
'-- Description       交付条件设计结果修改及查询
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"
    
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'TOP
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
                Call Gp_Ms_Collection(txt_ORD_NO, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_ORD_ITEM, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(txt_KND, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
               Call Gp_Ms_Collection(txt_Design_STS, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_STDSPEC, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_STDSPEC_YY, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_PROD_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ORD_THK, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_ORD_LEN, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_DEL_TO_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_CUST_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_UNIT_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_CUST_SPEC_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_ENDUSE_CD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_DEV_STD_CD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_SURF_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_SURF_GRD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
             Call Gp_Ms_Collection(txt_SHAPE_GRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_SHAPE_GRD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Body
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
           Call Gp_Ms_Collection(sdb_LEN_TOL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_LEN_TOL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_WID_TOL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_WID_TOL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_THK_TOL_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_THK_TOL_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_STRT_KND, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_STRT_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              
              Call Gp_Ms_Collection(sdb_STRT_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(sdb_STRT_UNIT_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_FLT_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_FLT_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
               Call Gp_Ms_Collection(sdb_FLT_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_FLT_UNIT_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(sdb_TELSCOPE_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_TELSCOPE_UNIT_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_RECT_TYP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_RECT_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              
              Call Gp_Ms_Collection(sdb_RECT_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_THK_DVT_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(sdb_VRT_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_CROWN_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_SAPE_WAVE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_THK_SHR_DRAG_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_WID_SHR_DRAG_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(sdb_LEN_SHR_DRAG_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
'Insert Emp
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
               Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(ul_INS_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(ul_INS_EMP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(ul_UPD_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(ul_UPD_EMP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
                
'LP板尺寸 刘翔 2012.11.16
            Call Gp_Ms_Collection(txt_ORD_LP_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ORD_LP_THK1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ORD_LP_THK2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ORD_LP_THK3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ORD_LP_LEN1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ORD_LP_LEN2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ORD_LP_LEN3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ORD_LP_LEN4, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ORD_LP_LEN5, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 '顾汉锋 20160623
            Call Gp_Ms_Collection(sdb_DIA_TOL_MIN, " ", " ", " ", "", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(sdb_DIA_TOL_MAX, " ", " ", " ", "", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_STRT_LENGTH, " ", " ", " ", "", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_FLT_LEN, " ", " ", " ", "", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_FLT_LEN2, " ", " ", " ", "", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(sdb_FLT_UNIT_MAX2, " ", " ", " ", "", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(sdb_DIFF_DVT_MAX, " ", " ", " ", "", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         

    'MASTER Collection
     Mc1.Add Item:="AQB0150C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQB0150C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
     'Picture2.BackColor = &HE0E0E0

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        
        Case "txt_DEV_STD_CD"       '交付条件代表标准
            sCode = "DEV_STD_CD"
        
        Case "txt_SURF_GRD"         '表面等级
            sCode = "Q0050"
            Set oCodeName = txt_SURF_GRD_NAME
            
        Case "txt_SHAPE_GRD"        '形状等级
            sCode = "Q0051"
            Set oCodeName = txt_SHAPE_GRD_NAME
        
        Case "txt_STRT_KND"         '镰刀弯
            sCode = "Q0030"
            Set oCodeName = txt_STRT_NAME
        
        Case "txt_FLT_TYP"          '不平度
            sCode = "Q0029"
            Set oCodeName = txt_FLT_NAME
        
        Case "txt_RECT_TYP"         '直角度
            sCode = "Q0041"
            Set oCodeName = txt_RECT_NAME
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name, True)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))

    txt_ORD_NO.Text = sOrderNo
    txt_ORD_ITEM.Text = sOrderItem

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    txt_ORD_NO.Text = ""
    txt_ORD_ITEM.Text = ""
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    pControl(1).SetFocus
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()
        
    Dim sMesg As String
    
    sMesg = Gf_Ms_NeceCheck(pControl)
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(pControl)
        If sMesg = "OK" Then
        
            If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Gf_subMasterLock(Mc1, Trim(txt_Design_STS.Text))
                Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            Else
                Call Gp_Ms_Cls(Mc1("rControl"))
            
            End If
            
        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
            
        End If
    Else
    
        sMesg = sMesg + " 必须输入！"
        Call Gp_MsgBoxDisplay(sMesg)
            
    End If
    
    If Len(txt_DEL_TO_DATE.Caption) = 8 Then txt_DEL_TO_DATE.Caption = Format(txt_DEL_TO_DATE.Caption, "####-##-##")
    
End Sub

Public Sub Form_Pro()
    If txt_DEV_STD_CD.Enabled = False Then Exit Sub
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        txt_ins_emp.Text = sUserID
        If subMinMaxValueCheck = False Then Exit Sub
        
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub


Private Sub opt_KND_Click(Index As Integer)
      txt_KND.Text = Index
      Call Form_Ref
End Sub


'下限值 , 上限值 Check
Private Function subMinMaxValueCheck() As Boolean
    
    If subValueCheck(sdb_LEN_TOL_MIN, sdb_LEN_TOL_MAX) = False Then Exit Function
    If subValueCheck(sdb_WID_TOL_MIN, sdb_WID_TOL_MAX) = False Then Exit Function
    If subValueCheck(sdb_THK_TOL_MIN, sdb_THK_TOL_MAX) = False Then Exit Function
    
    subMinMaxValueCheck = True

End Function


'下限值 , 上限值 Check
Private Function subValueCheck(ByVal dMin As Object, ByVal dMax As Object) As Boolean

    If Trim(dMin.Text) = "" And Trim(dMax.Text) = "" Then
        subValueCheck = True
        Exit Function
    End If

    If Trim(dMin.Text) <> "" And Trim(dMax.Text) <> "" Then
        If Val(dMin) <= Val(dMax) Then
            subValueCheck = True
        Else
            subValueCheck = False
            Call Gp_MsgBoxDisplay(dMax.Name & " is more then " & dMin.Name, "I")
            dMin.SetFocus
        End If
    Else
        subValueCheck = True
    End If
        
End Function


