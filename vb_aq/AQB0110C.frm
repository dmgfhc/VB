VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQB0110C 
   Caption         =   "成分设计结果查询 - AQB0110C"
   ClientHeight    =   9090
   ClientLeft      =   -1155
   ClientTop       =   1635
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.OptionButton opt_KND 
      BackColor       =   &H00E0E0E0&
      Caption         =   "成品成分"
      Height          =   315
      Index           =   4
      Left            =   7080
      TabIndex        =   13
      Top             =   1335
      Width           =   1095
   End
   Begin VB.TextBox txt_Design_STS 
      Height          =   315
      Left            =   6900
      TabIndex        =   12
      Top             =   150
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.OptionButton opt_KND 
      BackColor       =   &H00E0E0E0&
      Caption         =   "标准成分"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   11
      Top             =   1350
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton opt_KND 
      BackColor       =   &H00E0E0E0&
      Caption         =   "企标成分"
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   10
      Top             =   1335
      Width           =   1065
   End
   Begin VB.OptionButton opt_KND 
      BackColor       =   &H00E0E0E0&
      Caption         =   "客户特殊要求成分"
      Height          =   315
      Index           =   2
      Left            =   4830
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txt_STLGRD_GRP_NAME 
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
      Left            =   10230
      TabIndex        =   8
      Top             =   1730
      Width           =   4845
   End
   Begin VB.TextBox txt_STLGRD_GRP 
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
      Left            =   8940
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1730
      Width           =   1275
   End
   Begin VB.TextBox txt_ins_emp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13560
      MaxLength       =   7
      TabIndex        =   6
      Tag             =   "INS_EMP"
      Top             =   1395
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.TextBox txt_STLGRD_NAME 
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
      Left            =   3060
      TabIndex        =   5
      Top             =   1730
      Width           =   3975
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
      Left            =   1620
      MaxLength       =   11
      TabIndex        =   3
      Tag             =   "订单号"
      Top             =   120
      Width           =   2235
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
      Left            =   5340
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "序列号"
      Top             =   120
      Width           =   645
   End
   Begin VB.TextBox txt_STLGRD 
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
      Top             =   1730
      Width           =   1395
   End
   Begin VB.TextBox txt_KND 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   13020
      TabIndex        =   0
      Text            =   "1"
      Top             =   1425
      Visible         =   0   'False
      Width           =   495
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   180
      Top             =   540
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   556
      Caption         =   "标准代码"
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
      Index           =   2
      Left            =   2250
      Top             =   540
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   556
      Caption         =   "发布年度"
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
      Index           =   3
      Left            =   3120
      Top             =   540
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "品种"
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
      Index           =   4
      Left            =   3870
      Top             =   540
      Width           =   1230
      _ExtentX        =   2170
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   5
      Left            =   5100
      Top             =   540
      Width           =   1230
      _ExtentX        =   2170
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   6
      Left            =   6330
      Top             =   540
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
      Index           =   7
      Left            =   7530
      Top             =   540
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "交货日期"
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
      Index           =   8
      Left            =   8760
      Top             =   540
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
      Index           =   9
      Left            =   9990
      Top             =   540
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "订单产品重量"
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
      Index           =   10
      Left            =   11220
      Top             =   540
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
      Index           =   11
      Left            =   12450
      Top             =   540
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   556
      Caption         =   "订单用途"
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
      Index           =   0
      Left            =   180
      Top             =   120
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
      Left            =   3900
      Top             =   120
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
      Index           =   12
      Left            =   180
      Top             =   1725
      Width           =   1365
      _ExtentX        =   2408
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   14
      Left            =   7530
      Top             =   1725
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "钢种Group"
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
      ForeColor       =   16711680
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   6975
      Left            =   135
      TabIndex        =   4
      Top             =   2175
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   12303
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
      MaxCols         =   22
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQB0110C.frx":0000
   End
   Begin InDate.ULabel txt_STDSPEC 
      Height          =   345
      Left            =   180
      Top             =   840
      Width           =   2085
      _ExtentX        =   3678
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
      Left            =   2250
      Top             =   840
      Width           =   885
      _ExtentX        =   1561
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
      Left            =   3120
      Top             =   840
      Width           =   765
      _ExtentX        =   1349
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
   Begin InDate.ULabel txt_ORD_THK 
      Height          =   345
      Left            =   3870
      Top             =   840
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
   Begin InDate.ULabel txt_ORD_WID 
      Height          =   345
      Left            =   5100
      Top             =   840
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
      Left            =   6330
      Top             =   840
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
      Left            =   7560
      Top             =   840
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
      Left            =   8760
      Top             =   840
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
      Left            =   9990
      Top             =   840
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
      Left            =   11220
      Top             =   840
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
      Left            =   12450
      Top             =   840
      Width           =   2595
      _ExtentX        =   4577
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
      Height          =   315
      Index           =   13
      Left            =   180
      Top             =   1305
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "标准分类"
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   1680
      Top             =   1305
      Width           =   6645
   End
End
Attribute VB_Name = "AQB0110C"
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
'-- Program Name      成分设计结果修改及查询
'-- Program ID        AQB0110C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.08.21
'-- Description       成分设计结果修改及查询
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

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim ArrayRecords As Variant

'----------------------------------------------------
'Form_Define
'----------------------------------------------------
Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Hsheet"

   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   
         Call Gp_Ms_Collection(txt_ORD_NO, "p", "n", " ", "i", " ", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ORD_ITEM, "p", "n", " ", "i", " ", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
         Call Gp_Ms_Collection(txt_STLGRD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_STLGRD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STLGRD_GRP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_STLGRD_GRP_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_KND, "p", " ", " ", " ", " ", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
    'MASTER Collection
    Mc1.Add Item:="AQB0110C.P_REFER_MASTER", Key:="P-R"
    Mc1.Add Item:="AQB0110C.P_MODIFY_MASTER", Key:="P-M"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQB0110C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQB0110C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQB0110C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=2, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
            
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
        Case "txt_STLGRD"               '钢种
            sCode = "STLGRD"
            Set oCodeName = txt_STLGRD_NAME
        
        Case "txt_STLGRD_GRP"           '钢种Group
            sCode = "Q0048"
            Set oCodeName = txt_STLGRD_GRP_NAME
                    
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub


'----------------------------------------------------
'Form_Activate
'----------------------------------------------------
Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    Call GP_MENU_SHOW_HIDE("11F12F")
    
End Sub

'----------------------------------------------------
'Form_KeyPress
'----------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub



'----------------------------------------------------
'Form_Load
'----------------------------------------------------
Private Sub Form_Load()

    Dim x As Boolean

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
        
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 4)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 10)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 16)
    
    Screen.MousePointer = vbDefault
    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
      
    ArrayRecords = GF_GetChemicalCode
   
    txt_ORD_NO.Text = sOrderNo
    txt_ORD_ITEM.Text = sOrderItem
    
    ss1.TextTip = SS_TEXTTIP_FLOATINGFOCUSONLY
    ss1.TextTipDelay = 250
    x = ss1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
    
    Call GP_MENU_SHOW_HIDE("11F12F")

End Sub

'----------------------------------------------------
'Form_QueryUnload
'----------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

'----------------------------------------------------
'Spread_Can
'----------------------------------------------------
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
    Call GS_SetChemicalSpreadLineColor(ss1, "0915")
      
End Sub


'----------------------------------------------------
'Form_Cls
'----------------------------------------------------
Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        pControl(1).SetFocus
    End If
    
    Call GP_MENU_SHOW_HIDE("11F12F")
    
    
End Sub

'----------------------------------------------------
'Form_Ref
'----------------------------------------------------
Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    sMesg = Gf_Ms_NeceCheck(pControl)
    If sMesg = "OK" Then

        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then

            If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
                Call Gf_Sp_Display(M_CN1, Proc_Sc("Sc").Item("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), Proc_Sc("Sc").Item("pColumn"))
                Call subSpreadLockCol
                Call Gf_subMasterLock(Mc1, Trim(txt_Design_STS.Text))
                Call Gp_Ms_ControlLock(Mc1("pControl"), True)
                Call GS_SetChemicalLength(ss1, ArrayRecords, "041016", txt_KND.Text)
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'                Call GP_ChemCode_RowHeader_Clear("AQB0110C", ss1, 1, 1)
            End If

        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
        End If

    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)

    End If
    
    Dim i As Integer
    
    Call GS_SetChemicalSpreadLineColor(ss1, "0915")
    
    Call GP_MENU_SHOW_HIDE("11F12F")
    
    If Len(txt_DEL_TO_DATE.Caption) = 8 Then txt_DEL_TO_DATE.Caption = Format(txt_DEL_TO_DATE.Caption, "####-##-##")
    
    Exit Sub
    
    

Refer_Err:

End Sub

'----------------------------------------------------
'Form_Pro
'----------------------------------------------------
Public Sub Form_Pro()
         
    If txt_STLGRD.Enabled = False Then Exit Sub
    
    If Gf_Mc_Authority(sAuthority, Mc1, Proc_Sc("Sc")) Then
        txt_ins_emp.Text = sUserID
        
         If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1)
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call GS_SetChemicalLength(ss1, ArrayRecords, "041016", txt_KND.Text)
        End If
                
    End If
    
     Call subSpreadLockCol
     Call GS_SetChemicalSpreadLineColor(ss1, "0915")
     Call GP_MENU_SHOW_HIDE("11F12F")
End Sub

'----------------------------------------------------
'Form_Exc
'----------------------------------------------------
Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

'----------------------------------------------------
'Form_Ins
'----------------------------------------------------
Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)

End Sub

'----------------------------------------------------
'Form_Del
'----------------------------------------------------
Public Sub Form_Del()

    If Not Gf_Ms_AllDel(M_CN1, Proc_Sc("Sc"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub

'----------------------------------------------------
'Form_Cpy
'----------------------------------------------------
Public Sub form_Cpy()
    
    Call Gf_Ms_Copy(Mc1)
    
End Sub

'----------------------------------------------------
'Form_Pst
'----------------------------------------------------
Public Sub form_Pst()
    
    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21, "P")
    End If
    
End Sub

'----------------------------------------------------
'Form_Exit
'----------------------------------------------------
Public Sub Form_Exit()
    Unload Me
End Sub

'----------------------------------------------------
'Master_Cpy
'----------------------------------------------------
Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

'----------------------------------------------------
'Master_Pst
'----------------------------------------------------
Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
End Sub

'----------------------------------------------------
'Spread_Cpy
'----------------------------------------------------
Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

'----------------------------------------------------
'Spread_Pst
'----------------------------------------------------
Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)
    
End Sub

'----------------------------------------------------
'Spread_Del
'----------------------------------------------------
Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("Sc"))

End Sub


'----------------------------------------------------
'opt_KND_Click
'----------------------------------------------------
Private Sub opt_KND_Click(Index As Integer)
    txt_KND.Text = Index
    Call Form_Ref
End Sub

'----------------------------------------------------
'ss1_BlockSelected
'----------------------------------------------------
Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

'----------------------------------------------------
'ss1_Click
'----------------------------------------------------
Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
        
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

'----------------------------------------------------
'ss1_ComboSelChange
'----------------------------------------------------
Private Sub ss1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    
    ss1.Col = Col
    ss1.Row = Row
    
    Call GF_GetCeqValue(ss1, Me.Name, txt_KND.Text)
    
    Call subSpreadLockCol
    
    
End Sub

'----------------------------------------------------
'ss1_EditMode
'----------------------------------------------------
Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)
    End If
    
End Sub

'----------------------------------------------------
'ss1_KeyDown
'----------------------------------------------------
Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

'----------------------------------------------------
'ss1_KeyUp
'----------------------------------------------------
Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim sTemp_Code As String
    Dim iCol As Integer

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol

        Case 4, 10, 16

            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1

                DD.sWitch = "SP"
              '  DD.sKey = "Q0001"
                DD.rControl.Add Item:=ss1.ActiveCol
                DD.nameType = "2"

                Call GF_CHEM_SEQ(M_CN1, KeyCode)
                
                Call GS_SetChemicalLength(ss1, ArrayRecords, "041016", txt_KND.Text)
        
            End If

    End Select

    ' ss1.SetFocus

End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Call GP_ChemCode_RowHeader_Clear("AQB0110C", ss1, NewRow, NewCol)
End Sub

'----------------------------------------------------
'ss1_LostFocus
'----------------------------------------------------
Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

'----------------------------------------------------
'ss1_TextTipFetch
'----------------------------------------------------
Private Sub ss1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    'TipText = ss1.CellTag
    'ShowTip = True
    Dim sTip As String
    ShowTip = True
    'With ss1
        
    sTip = GF_GetCellMaxLength(ss1, ss1.ActiveRow, ss1.ActiveCol)
    TipText = sTip
    'End With
End Sub


'----------------------------------------------------
'subSpreadLockCol
'----------------------------------------------------
Private Sub subSpreadLockCol()

    Dim i As Integer

    With ss1
            
        If txt_Design_STS.Text = "A" Then
            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 5: .Lock = True
                .Col = 6: .Lock = True
                .Col = 7: .Lock = True
                .Col = 8: .Lock = True
                .Col = 11: .Lock = True
                .Col = 12: .Lock = True
                .Col = 13: .Lock = True
                .Col = 14: .Lock = True
                .Col = 17: .Lock = True
                .Col = 18: .Lock = True
                .Col = 19: .Lock = True
                .Col = 20: .Lock = True
            Next i
        Else
            For i = 1 To .MaxRows
                .Row = i
                
                .Col = 5: .Lock = False
                .Col = 6: .Lock = False
                .Col = 8: .Lock = False
                .Col = 11: .Lock = False
                .Col = 12: .Lock = False
                .Col = 14: .Lock = False
                .Col = 17: .Lock = False
                .Col = 18: .Lock = False
                .Col = 20: .Lock = False
                If txt_KND.Text = "3" Then
                    .Col = 7: .Lock = False
                    .Col = 13: .Lock = False
                    .Col = 19: .Lock = False
                Else
                    .Col = 7:  .Lock = True
                    .Col = 13: .Lock = True
                    .Col = 19: .Lock = True
                End If
            Next i
        End If
        
        
    End With

End Sub

Private Sub subMasterLock()
    If txt_Design_STS.Text = "A" Then
        Call Gp_Ms_ControlLock(Mc1("iControl"), True)
    Else
        Call Gp_Ms_ControlLock(Mc1("iControl"), False)
    End If
End Sub


