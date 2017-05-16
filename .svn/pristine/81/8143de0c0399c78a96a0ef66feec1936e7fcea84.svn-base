VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0080C 
   Caption         =   "成分/材质设计标准修改及查询 - AQC0080C"
   ClientHeight    =   8925
   ClientLeft      =   -1155
   ClientTop       =   1635
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_Smp_No 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AQC0080C.frx":0000
      Left            =   11535
      List            =   "AQC0080C.frx":000D
      TabIndex        =   10
      Top             =   135
      Width           =   2130
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
      Left            =   10020
      TabIndex        =   5
      Top             =   1215
      Width           =   5025
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
      Left            =   8790
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1215
      Width           =   1230
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
      Left            =   3270
      TabIndex        =   3
      Top             =   1215
      Width           =   4245
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
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   2
      Tag             =   "订单号"
      Top             =   120
      Width           =   1560
   End
   Begin VB.TextBox txt_ORD_ITEM 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4725
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "序列号"
      Top             =   120
      Width           =   420
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
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   0
      Top             =   1215
      Width           =   1710
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
      Left            =   3885
      Top             =   540
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   6330
      Top             =   540
      Width           =   1215
      _ExtentX        =   2143
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
      Left            =   7560
      Top             =   540
      Width           =   1215
      _ExtentX        =   2143
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
      Left            =   8775
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
      Left            =   10005
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
      Left            =   11235
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
      Left            =   12465
      Top             =   540
      Width           =   2565
      _ExtentX        =   4524
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
      Left            =   3330
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
      Left            =   180
      Top             =   1215
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
      Index           =   14
      Left            =   7530
      Top             =   1215
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "钢种Group"
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7365
      Left            =   195
      TabIndex        =   6
      Top             =   1620
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   12991
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "AQC0080C.frx":001A
      Begin FPSpread.vaSpread ss3 
         Height          =   7365
         Left            =   8520
         TabIndex        =   12
         Top             =   0
         Width           =   6315
         _Version        =   393216
         _ExtentX        =   11139
         _ExtentY        =   12991
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
         SpreadDesigner  =   "AQC0080C.frx":008C
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   7365
         Left            =   3930
         TabIndex        =   13
         Top             =   0
         Width           =   4500
         _Version        =   393216
         _ExtentX        =   7937
         _ExtentY        =   12991
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
         SpreadDesigner  =   "AQC0080C.frx":042A
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   7365
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3840
         _Version        =   393216
         _ExtentX        =   6773
         _ExtentY        =   12991
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
         MaxCols         =   12
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0080C.frx":09B3
      End
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   735
      Left            =   105
      TabIndex        =   7
      Top             =   9855
      Visible         =   0   'False
      Width           =   14925
      _Version        =   393216
      _ExtentX        =   26326
      _ExtentY        =   1296
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
      MaxCols         =   278
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0080C.frx":0FC1
   End
   Begin Threed.SSCommand cmd_Ord_Search 
      Height          =   345
      Left            =   5415
      TabIndex        =   9
      Top             =   105
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "订单查询"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmd_Run_Test 
      Height          =   345
      Left            =   13755
      TabIndex        =   11
      Top             =   105
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "再综合判定"
      BevelWidth      =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "综合判定不合格产品"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   9705
      TabIndex        =   8
      Top             =   195
      Width           =   1815
   End
End
Attribute VB_Name = "AQC0080C"
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
'-- Program Name      成分/材质设计标准修改及查询
'-- Program ID        AQC0080C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2005.11.04
'-- Description       成分/材质设计标准修改及查询
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
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
    Dim iRow   As Long
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Hsheet"

   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   
         Call Gp_Ms_Collection(txt_ORD_NO, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ORD_ITEM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_STDSPEC, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STDSPEC_YY, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_PROD_CD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ORD_THK, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ORD_WID, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ORD_LEN, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_DEL_TO_DATE, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_cust_cd, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_UNIT_WGT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_CUST_SPEC_NO, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_ENDUSE_CD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_STLGRD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_STLGRD_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STLGRD_GRP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_STLGRD_GRP_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
    'MASTER Collection
    Mc1.Add Item:="AQC0080C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
         
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQC0080C.P_MODIFY1", Key:="P-M"
    Sc1.Add Item:="AQC0080C.P_SREFER1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=2, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
     
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
      Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
      Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
      
      Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
      Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
      Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
      Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
      Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
      
      Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
      Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     'louyannan 20101126 zra
     Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     
         
     Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 33, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 34, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 35, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 36, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 37, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 38, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 39, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 40, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 41, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 42, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 43, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 44, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 45, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 46, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 47, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 48, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 49, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 50, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 51, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 52, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 53, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 54, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 55, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 56, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 57, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 58, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 59, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 60, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 61, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 62, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 63, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 64, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 65, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 66, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 67, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 68, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 69, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 70, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 71, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 72, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 73, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 74, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 75, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 76, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 77, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 78, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 79, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 80, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 81, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 82, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 83, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 84, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 85, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 86, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 87, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 88, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 89, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 90, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 91, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 92, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     Call Gp_Sp_Collection(ss2, 93, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 94, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 95, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 96, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 97, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     
     'louyannan 201011226 hgt_zra
      Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
     
     Call Gp_Sp_Collection(ss2, 98, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 99, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 100, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 101, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 102, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 103, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 104, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 105, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 106, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 107, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 108, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 109, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 110, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 111, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 112, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 113, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 114, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 115, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 116, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 117, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 118, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 119, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 120, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 121, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 122, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 123, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 124, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 125, " ", " ", " ", "  ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 126, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 127, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    Call Gp_Sp_Collection(ss2, 128, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 129, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 130, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 131, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 132, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 133, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 134, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 135, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 136, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 137, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 138, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 139, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 140, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 141, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 142, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 143, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 144, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 145, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 146, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 147, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 148, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 149, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 150, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 151, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 152, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 153, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 154, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 155, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 156, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 157, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
'20090515 sun bin
    Call Gp_Sp_Collection(ss2, 158, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 159, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 160, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 161, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 162, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
'20090515 sun bin
    Call Gp_Sp_Collection(ss2, 163, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 164, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 165, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 166, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 167, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 168, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 169, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 170, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 171, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 172, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 173, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 174, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 175, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 176, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 177, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 178, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 179, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 180, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 181, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 182, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     
    Call Gp_Sp_Collection(ss2, 183, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 184, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 185, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 186, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 187, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    Call Gp_Sp_Collection(ss2, 188, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQC0080C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:="AQC0080C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=2, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
     
       'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
         
    'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=3, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 5, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 6, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 7, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 8, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 9, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 10, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 11, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 12, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
         
    'Spread_Collection
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="AQC0080C.P_SREFER4", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=2, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=sc3, Key:="Sc3"
    Proc_Sc.Add Item:=sc4, Key:="Sc4"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    sc4.Item("Spread").Col = 0
    sc4.Item("Spread").Row = 0
    sc4.Item("Spread").Text = "◎"
    Call MenuToolSet
            
End Sub

Private Sub MenuToolSet()

    With MDIMain.MenuTool
        .Buttons(4).Enabled = False                 'Save
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub


'Private Sub cbo_loc_Click(Index As Integer)
'If cbo_loc(Index).Value = "1" Then
'cbo_loc(Abs(Index - 1)).Value = "0"
'Else
'cbo_loc(Abs(Index - 1)).Value = 1
'
'
'End Sub

Private Sub cmd_Ord_Search_Click()
    If Len(Trim(txt_ORD_NO)) < 9 Then
        Call Gp_MsgBoxDisplay("请输入订单号(> 8)")
        Exit Sub
    End If
    Call Gf_Sp_Display(M_CN1, Proc_Sc("Sc4").Item("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc4").Item("P-R"), "R", Mc1("pControl")), Proc_Sc("Sc4").Item("pColumn"))
    Call Form_Ref
    
End Sub

'----------------------------------------------------
'Form_Activate
'----------------------------------------------------
Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuToolSet
    
'    Call GP_MENU_SHOW_HIDE("11F12F")
    
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
        
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    With ss3
        .Row = 0: .Row2 = 0
        .Col = 6: .Col2 = 6
        
        .BlockMode = True
        
        .CellType = SS_CELL_TYPE_STATIC_TEXT
        .TypeHAlign = SS_CELL_H_ALIGN_CENTER
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .TypeTextWordWrap = True
        
        .BackColor = &HE1E4CD
        .ForeColor = BLUE
        
        .BlockMode = False

    End With
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc4")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    Call Gf_Sp_Cls(Proc_Sc("Sc4"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc4")("Spread"), "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
        
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
          
    ArrayRecords = GF_GetChemicalCode
'    Call GP_MENU_SHOW_HIDE("11F12F")

End Sub

'----------------------------------------------------
'Form_QueryUnload
'----------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc3")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc4")("Spread"), "Q-System.INI", Me.Name)
    
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
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    Set sc4 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

'----------------------------------------------------
'Spread_Can
'----------------------------------------------------
Public Sub Spread_Can()

    ss1.Col = 0
    ss1.Row = 0
    If ss1.Text = "◎" Then
        Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc1"))
    End If
    
    ss3.Col = 0
    ss3.Row = 0
    If ss3.Text = "◎" Then
        Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc3"))
    End If
End Sub


'----------------------------------------------------
'Form_Cls
'----------------------------------------------------
Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc1")) Then
        If Gf_Sp_Cls(Proc_Sc("Sc3")) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("pControl"), False)
            pControl(1).SetFocus
        End If
    End If
    Call MenuToolSet
    
End Sub

'----------------------------------------------------
'Form_Ref
'----------------------------------------------------
Public Sub Form_Ref()
    Dim iRow        As Long
    Dim iCol        As Long

On Error GoTo Refer_Err

    Dim sMesg As String

    If Gf_Sp_ProceExist(Proc_Sc("Sc1").Item("Spread")) Then Exit Sub
    If Gf_Sp_ProceExist(Proc_Sc("Sc3").Item("Spread")) Then Exit Sub
    
    If Len(txt_ORD_NO.Text) = 11 And Len(txt_ORD_ITEM.Text) = 2 Then
        If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            Call Gf_Sp_Display(M_CN1, Proc_Sc("Sc1").Item("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc1").Item("P-R"), "R", Mc1("pControl")), Proc_Sc("Sc1").Item("pColumn"))
            Call Gf_Sp_Display(M_CN1, Proc_Sc("Sc2").Item("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")), Proc_Sc("Sc2").Item("pColumn"))
            Call SS3_DataEdit
 
            Call SetChemicalLength(ss1, ArrayRecords, "03", "1")
            Call Gp_Sp_BlockColor(ss3, 3, 6, 1, ss3.MaxRows, , &HC0FFFF)
        End If
       
    End If
    
    ss1.BlockMode = True
    ss1.Row = 1:  ss1.Row2 = ss1.MaxRows
    ss1.Col = 6:  ss1.Col2 = 6
    ss1.Lock = False
    ss1.BlockMode = False
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 3
        If Trim(ss1.Text) <> "Ceq" Then
            ss1.BlockMode = True
            ss1.Row = iRow: ss1.Row2 = iRow
            ss1.Col = 6:    ss1.Col2 = 6
            ss1.Lock = True
            ss1.BlockMode = False
        End If
    Next iRow
    
'    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MenuToolSet

    Exit Sub

Refer_Err:

End Sub

'----------------------------------------------------
'Form_Pro
'----------------------------------------------------
Public Sub Form_Pro()
    Dim iRow        As Long
    Dim iCol        As Long
    Dim iCnt1       As Long
    Dim iCnt2       As Long
    
    iCnt1 = 0:      iCnt2 = 0
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 0
        If Trim(ss1.Text) <> "" Then
            iCnt1 = 1
            Exit For
        End If
    Next iRow
    
    For iRow = 1 To ss3.MaxRows
        ss3.Row = iRow
        ss3.Col = 0
        If Trim(ss3.Text) <> "" Then
            iCnt2 = 1
            Exit For
        End If
    Next iRow
        
    If iCnt1 > 0 Then
        If Not Gf_Sp_Process(M_CN1, Proc_Sc("Sc1"), Mc1) Then Exit Sub
    End If
    
    If iCnt2 > 0 Then
        If Not Gf_Sp_Process(M_CN1, Proc_Sc("Sc2"), Mc1) Then Exit Sub
    End If
  
    ss1.MaxRows = 0
    ss2.MaxRows = 0
    ss3.MaxRows = 0
    
    Call Form_Ref
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MenuToolSet
    
End Sub

'----------------------------------------------------
'Form_Pro
'----------------------------------------------------
Public Sub SS3_DataEdit()
    Dim iRow        As Long
    Dim iCol        As Long
    Dim CNT         As Long
    Dim sQtyName    As String
    Dim sQtyFl      As String
    Dim sQtyMin     As String
    Dim sQtyMinAve  As String
    Dim sQtyMax     As String
    Dim sQuery      As String
    Dim AdoRs           As adodb.Recordset
    If ss2.MaxRows = 0 Then
        ss3.MaxRows = 0
        Exit Sub
    End If
    CNT = 0
    For iRow = 1 To (ss2.MaxCols - 3) / 5
        iCol = (iRow * 4) + (CNT - 1)
        
        ss2.Row = 1
        ss2.Col = iCol
        sQtyName = Trim(ss2.Text & "")
        ss2.Col = iCol + 1
        sQtyMin = Trim(ss2.Text & "")
        ss2.Col = iCol + 2
        sQtyMinAve = Trim(ss2.Text & "")
        ss2.Col = iCol + 3
        sQtyMax = Trim(ss2.Text & "")
        ss2.Col = iCol + 4
        sQtyFl = Trim(ss2.Text & "")
        
        
        ss3.MaxRows = iRow
        ss3.Row = iRow
        ss3.Col = 2
        ss3.Text = sQtyName
        ss3.Col = 3
        ss3.Text = sQtyMin
        ss3.Col = 4
        ss3.Text = sQtyMinAve
        ss3.Col = 5
        ss3.Text = sQtyMax
        ss3.Col = 6
        ss3.Text = sQtyFl

        CNT = CNT + 1
    Next iRow
    
    ss3.Col = 1
    ss3.Row = 1:     ss3.Text = "拉伸试验"
    ss3.Row = 11:     ss3.Text = "追加拉伸试验"
    ss3.Row = 25:    ss3.Text = "高温拉伸试验"
    ss3.Row = 32:    ss3.Text = "追加高温拉伸试验"
    ss3.Row = 38:    ss3.Text = "冲击、时效"
    ss3.Row = 42:    ss3.Text = "其它"
    ss3.Row = 45:    ss3.Text = "金相检验"
    
    '--------------------配置化项目显示 王成  2012.12.13-----------------------------------------------------
    
     Set AdoRs = New adodb.Recordset
            sQuery = "{call AQC0080C.P_SREFER2_CONFIG('" + Trim(txt_ORD_NO.Text) + "','" + Trim(txt_ORD_ITEM.Text) + "')}"
            AdoRs.Open sQuery, M_CN1, adOpenKeyset

            Set AdoRs = M_CN1.Execute(sQuery)

            If Not AdoRs.EOF And Not AdoRs.BOF Then
                  ArrayRecords = AdoRs.GetRows
                  Call subSpreadView_Config(ArrayRecords)
            End If

            AdoRs.Close
            Erase ArrayRecords
    
 '-----------------------------------------------------------------------------------------------------------
    
End Sub

'----------------------------------------------------
'Form_Exc
'----------------------------------------------------
Public Sub Form_Exc()
    ss1.Col = 0
    ss1.Row = 0
    If ss1.Text = "◎" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If
    
    ss3.Col = 0
    ss3.Row = 0
    If ss3.Text = "◎" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If
    
    ss4.Col = 0
    ss4.Row = 0
    If ss4.Text = "◎" Then
        Call Gp_Sp_Excel(Me, Proc_Sc("Sc4")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If
End Sub

'----------------------------------------------------
'Form_Exit
'----------------------------------------------------
Public Sub Form_Exit()
    Unload Me
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
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = "◎"

    sc3.Item("Spread").Col = 0
    sc3.Item("Spread").Row = 0
    sc3.Item("Spread").Text = ""
    
    sc4.Item("Spread").Col = 0
    sc4.Item("Spread").Row = 0
    sc4.Item("Spread").Text = ""
End Sub

Private Sub ss1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    
    ss1.Col = Col
    ss1.Row = Row
    
    Call GF_GetCeqValue(ss1, "AQC0080C", "1")
    
End Sub

Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc1")("Spread"), 2)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), 7)
        
        ss1.Row = Row
        ss1.Col = 1:      ss1.Text = Trim(txt_ORD_NO.Text)
        ss1.Col = 2:      ss1.Text = Trim(txt_ORD_ITEM.Text)
    End If
    
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

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = ""
    
    sc3.Item("Spread").Col = 0
    sc3.Item("Spread").Row = 0
    sc3.Item("Spread").Text = "◎"
    
    sc4.Item("Spread").Col = 0
    sc4.Item("Spread").Row = 0
    sc4.Item("Spread").Text = ""
End Sub

Private Sub ss3_EditChange(ByVal Col As Long, ByVal Row As Long)

    Dim sQtyFl      As String
    Dim sQtyMin     As String
    Dim sQtyMinAve  As String
    Dim sQtyMax     As String

            
    If Row < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc3")("Spread"), 2)
        If Col = 6 Then
            ss3.Row = Row
            ss3.Col = 6
            sQtyFl = Trim(ss3.Text & "")
            
            If sQtyFl <> "A" And sQtyFl <> "B" And sQtyFl <> "C" And sQtyFl <> "" Then
                Call Gp_MsgBoxDisplay("请再输入判定")
                ss3.Col = 0: ss3.Text = ""
                ss3.Col = 6: ss3.Text = ""
                Exit Sub
            End If
            
            ss2.Row = 1
            ss2.Col = (Row * 5) + 2
            ss2.Text = sQtyFl
        End If
        
        If Col = 3 Then
           ss3.Row = Row
           ss3.Col = 3
           sQtyMin = Trim(ss3.Text & "")
           ss2.Row = 1
           ss2.Col = (Row * 5) - 1
           ss2.Text = sQtyMin
        End If
        
        If Col = 4 Then
           ss3.Row = Row
           ss3.Col = 4
           sQtyMinAve = Trim(ss3.Text & "")
           ss2.Row = 1
           ss2.Col = (Row * 5)
           ss2.Text = sQtyMinAve
        End If

        If Col = 5 Then
           ss3.Row = Row
           ss3.Col = 5
           sQtyMax = Trim(ss3.Text & "")
           ss2.Row = 1
           ss2.Col = (Row * 5) + 1
           ss2.Text = sQtyMax
        End If

        ss2.Col = ss2.MaxCols
        ss2.Text = sUserID
        ss2.Col = 0
        ss2.Text = "Update"
    End If
    
End Sub

Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim sQtyFl      As String
            
    If Row < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc3")("Spread"), 2)
        
        Call ss3_EditChange(ss3.Col, ss3.Row)
    End If
    
End Sub

Private Sub ss3_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim str_orgin As String
    If KeyCode = vbKeyF4 Then
    
        With ss3
            .Col = .ActiveCol
            .Row = .ActiveRow
            If .ActiveCol = 3 Then
                DD.sWitch = "MS"
                DD.sKey = "Q0002"
                DD.rControl.Add Item:=ss3
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                Call ss3_EditChange(ss3.Col, ss3.Row)
            End If
        End With
        
    End If
    
End Sub

Private Sub ss3_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub


Private Sub ss4_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = ""
    
    sc3.Item("Spread").Col = 0
    sc3.Item("Spread").Row = 0
    sc3.Item("Spread").Text = ""
    
    sc4.Item("Spread").Col = 0
    sc4.Item("Spread").Row = 0
    sc4.Item("Spread").Text = "◎"
End Sub

Private Sub ss4_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
    
    ss4.Row = Row
    ss4.Col = 1
    txt_ORD_NO.Text = Trim(ss4.Text & "")
    ss4.Col = 2
    txt_ORD_ITEM.Text = Trim(ss4.Text & "")

    Call Form_Ref
End Sub

Private Sub ss4_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Public Sub SetChemicalLength(ByVal vSP As vaSpread, ByVal ArrayRecords As Variant, ByVal sCol As String, ByVal sKnd As String)

    Dim i As Integer
    Dim j As Integer
    Dim dblChem As Double
        
    With vSP
         
        For i = 1 To .MaxRows
           
           .Col = sCol: .Row = i
            
            For j = 0 To UBound(ArrayRecords, 2)
                If ArrayRecords(1, j) = Trim(.Text) Then
                    dblChem = ArrayRecords(2, j)
                    Exit For
                End If
            Next j
        
            Call subSetChemLength(vSP, dblChem, .Col, .Row, sKnd)
        
        Next i
                
    End With

End Sub


Private Sub ComboBoxEdit()
    Dim SQL   As String
    Dim AdoRs As adodb.Recordset

    Set AdoRs = New adodb.Recordset
    
    SQL = ""
    SQL = " SELECT        PROD_CD,  SMP_NO " & vbCrLf
    SQL = SQL & "   FROM  GP_PLATE          " & vbCrLf
    SQL = SQL & "  WHERE  ORD_NO    = '" & Trim(txt_ORD_NO.Text) & "'" & vbCrLf
    SQL = SQL & "    AND  ORD_ITEM  = '" & Trim(txt_ORD_ITEM.Text) & "'" & vbCrLf
    SQL = SQL & "    AND  REC_STS   = '2'   " & vbCrLf
    SQL = SQL & "    AND  PROC_CD   = 'QAE' " & vbCrLf
    SQL = SQL & "  GROUP  BY  SMP_NO , PROD_CD   " & vbCrLf
    SQL = SQL & "  UNION ALL " & vbCrLf
    SQL = SQL & " SELECT  'HC' PROD_CD,  SMP_NO " & vbCrLf
    SQL = SQL & "   FROM  GP_COIL           " & vbCrLf
    SQL = SQL & "  WHERE  ORD_NO    = '" & Trim(txt_ORD_NO.Text) & "'" & vbCrLf
    SQL = SQL & "    AND  ORD_ITEM  = '" & Trim(txt_ORD_ITEM.Text) & "'" & vbCrLf
    SQL = SQL & "    AND  REC_STS   = '2'     " & vbCrLf
    SQL = SQL & "    AND  PROC_CD   = 'QAE'   " & vbCrLf
    SQL = SQL & "  GROUP  BY  SMP_NO  ,PROD_CD  " & vbCrLf
'    SQL = SQL & "  UNION ALL                " & vbCrLf
'    SQL = SQL & " SELECT  'LP' PROD_CD,  SMP_NO " & vbCrLf
'    SQL = SQL & "   FROM  GP_PLATE          " & vbCrLf
'    SQL = SQL & "  WHERE  ORD_NO    = '" & Trim(txt_ORD_NO.Text) & "'" & vbCrLf
'    SQL = SQL & "    AND  ORD_ITEM  = '" & Trim(txt_ORD_ITEM.Text) & "'" & vbCrLf
'    SQL = SQL & "    AND  REC_STS   = '2'   " & vbCrLf
'    SQL = SQL & "    AND  PROC_CD   = 'QAE' " & vbCrLf
'    SQL = SQL & "  GROUP  BY  SMP_NO        " & vbCrLf
    SQL = SQL & "  ORDER  BY  SMP_NO        " & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    cbo_Smp_No.Clear
    
    Do Until AdoRs.EOF
       cbo_Smp_No.AddItem AdoRs.Fields("PROD_CD").Value & " " & AdoRs.Fields("SMP_NO").Value
       AdoRs.MoveNext
    Loop
    
    AdoRs.Close
      
End Sub

Private Sub cmd_Run_Test_Click()

    Dim OutParam(2, 4)      As Variant
    Dim ret_Result_ErrMsg   As String
    Dim sSmpNo              As String
    Dim sProdCd             As String
    Dim sQuery              As String
    
    Dim adoCmd As adodb.Command
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    On Error GoTo Process_Exec_ERROR

    sProdCd = Left(cbo_Smp_No, 2)
    sSmpNo = Mid(cbo_Smp_No, 4, Len(cbo_Smp_No))
        
    If (sProdCd <> "PP" And sProdCd <> "HC" And sProdCd <> "LP") Or Trim(sSmpNo) = "" Then
        Call Gp_MsgBoxDisplay("取样号错误!")
        Exit Sub
    End If
        
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 2

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    Set adoCmd.ActiveConnection = M_CN1
            
    '---------squery(CALL AQT1320P)----------------------
    sQuery = "{CALL AQT1320P('" & sSmpNo & "','" & sProdCd & "',?,?)}"
    '-------------------------------------------------------
    
    adoCmd.CommandType = adCmdText
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_code") <> "YY" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        GoTo Process_Exec_ERROR
    End If
    Set adoCmd = Nothing
    
    Call Gp_MsgBoxDisplay("处理完了..!!", "I")
    Call ComboBoxEdit
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error & "   " & ret_Result_ErrMsg)
End Sub

Private Sub txt_ORD_NO_Change()
    If Len(txt_ORD_NO.Text) = 11 And Len(txt_ORD_ITEM.Text) = 2 Then Call ComboBoxEdit
End Sub

Private Sub txt_ORD_ITEM_Change()
    If Len(txt_ORD_NO.Text) = 11 And Len(txt_ORD_ITEM.Text) = 2 Then Call ComboBoxEdit
End Sub

'--------------------配置化项目显示 王成  2012.12.13-----------------------------------------------------

Private Sub subSpreadView_Config(ByVal strArr As Variant)

    Dim i As Integer
    Dim OLD_MAXROWS As Integer
    
    If UBound(strArr, 2) < 0 Then Exit Sub
    
    With ss3
        OLD_MAXROWS = .MaxRows
        .MaxRows = .MaxRows + UBound(strArr, 2) + 1
        .Row = OLD_MAXROWS + 1
        .Col = 1
        .Text = "配置化项目"
        For i = 1 To UBound(strArr, 2) + 1
            .Row = OLD_MAXROWS + i
            .Col = 2: .Text = GF_NullChange(strArr(0, i - 1))
            .Col = 3: .Text = GF_NullChange(strArr(1, i - 1)) & ""
            .Col = 4: .Text = GF_NullChange(strArr(2, i - 1)) & ""
            .Col = 5: .Text = GF_NullChange(strArr(3, i - 1)) & ""
            .Col = 6: .Text = GF_NullChange(strArr(4, i - 1)) & ""
        Next i
            
    End With

End Sub
'----------------------------------------------------------------------------------------------------------
