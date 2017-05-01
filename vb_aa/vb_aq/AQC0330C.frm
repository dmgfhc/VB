VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0330C 
   Caption         =   "综合判定结果详细查询 - AQC0330C"
   ClientHeight    =   10500
   ClientLeft      =   180
   ClientTop       =   1560
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10500
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_SMP_CUT_LOC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   10620
      TabIndex        =   5
      Text            =   "B"
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   525
      Left            =   3690
      TabIndex        =   1
      Top             =   -90
      Width           =   6765
      Begin VB.OptionButton opt_SMP_CUT_LOC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头/尾"
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
         Index           =   5
         Left            =   5670
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton opt_SMP_CUT_LOC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头/中/尾"
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
         Index           =   4
         Left            =   4635
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton opt_SMP_CUT_LOC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "头部"
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
         Index           =   2
         Left            =   2565
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton opt_SMP_CUT_LOC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中部"
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
         Index           =   3
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   975
      End
      Begin VB.OptionButton opt_SMP_CUT_LOC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "尾部"
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
         Index           =   1
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Value           =   -1  'True
         Width           =   975
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   13
         Left            =   60
         Top             =   150
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "取样位置 "
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
   End
   Begin VB.TextBox txt_PROD_NO 
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
      MaxLength       =   14
      TabIndex        =   0
      Tag             =   "产品编号"
      Top             =   60
      Width           =   1905
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   180
      Top             =   600
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      Caption         =   "订单号/序列号"
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
      Left            =   2220
      Top             =   600
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   556
      Caption         =   "计划产品标准"
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
      Left            =   7110
      Top             =   600
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   556
      Caption         =   "企标钢种"
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
      Left            =   2220
      Top             =   1230
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   556
      Caption         =   "实际产品标准"
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
      Left            =   8820
      Top             =   600
      Width           =   2370
      _ExtentX        =   4180
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
      Index           =   8
      Left            =   4830
      Top             =   600
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   556
      Caption         =   "订单尺寸"
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
      Left            =   4830
      Top             =   1230
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   556
      Caption         =   "实际尺寸"
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
      Left            =   7080
      Top             =   1230
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "试样编号"
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
      Left            =   11205
      Top             =   600
      Width           =   3990
      _ExtentX        =   7038
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
      Index           =   0
      Left            =   180
      Top             =   60
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "产品编号"
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
      Left            =   9570
      Top             =   1230
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Caption         =   "综合等级"
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
   Begin InDate.ULabel txt_ORD_NO 
      Height          =   315
      Left            =   180
      Top             =   900
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_PROD_GRD 
      Height          =   315
      Left            =   9570
      Top             =   1530
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_PLAN_STD 
      Height          =   315
      Left            =   2220
      Top             =   900
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_STLGRD 
      Height          =   315
      Left            =   7110
      Top             =   900
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_STDSPEC 
      Height          =   315
      Left            =   2220
      Top             =   1530
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_ENDUSE 
      Height          =   315
      Left            =   8820
      Top             =   900
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_ORD_SIZE 
      Height          =   315
      Left            =   4830
      Top             =   900
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      Caption         =   ""
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_PROD_SIZE 
      Height          =   315
      Left            =   4830
      Top             =   1560
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   556
      Caption         =   ""
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_SMP_NO 
      Height          =   315
      Left            =   7080
      Top             =   1530
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_CUST 
      Height          =   315
      Left            =   11205
      Top             =   900
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
      Height          =   315
      Index           =   3
      Left            =   10470
      Top             =   1230
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Caption         =   "表面等级"
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
   Begin InDate.ULabel txt_SUF_GRD 
      Height          =   315
      Left            =   10470
      Top             =   1530
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
      Height          =   315
      Index           =   12
      Left            =   11370
      Top             =   1230
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   556
      Caption         =   "UST标准/结果"
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
   Begin InDate.ULabel txt_UST_STD_GRD 
      Height          =   315
      Left            =   11370
      Top             =   1530
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
      Height          =   315
      Index           =   16
      Left            =   180
      Top             =   1230
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Caption         =   "生产日期"
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
   Begin InDate.ULabel txt_PROD_DATE 
      Height          =   315
      Left            =   180
      Top             =   1530
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
      Height          =   315
      Index           =   17
      Left            =   1200
      Top             =   1230
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Caption         =   "判定日期"
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
   Begin InDate.ULabel txt_DSC_DATE 
      Height          =   315
      Left            =   1200
      Top             =   1530
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
      Height          =   315
      Index           =   18
      Left            =   8610
      Top             =   1230
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "冶炼炉号"
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
   Begin InDate.ULabel txt_HEAT_NO 
      Height          =   315
      Left            =   8610
      Top             =   1530
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
      Height          =   315
      Index           =   14
      Left            =   2220
      Top             =   1860
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      Caption         =   "发放日期"
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
      Index           =   15
      Left            =   3450
      Top             =   1860
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      Caption         =   "发放人员"
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
      Index           =   19
      Left            =   4860
      Top             =   1860
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      Caption         =   "提货单号"
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
   Begin InDate.ULabel txt_PRINT_DATE 
      Height          =   315
      Left            =   2220
      Top             =   2160
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_PRINT_EMP 
      Height          =   315
      Left            =   3450
      Top             =   2160
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      Caption         =   ""
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin InDate.ULabel txt_SHP_ISP_NO 
      Height          =   315
      Left            =   4860
      Top             =   2160
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
      Height          =   315
      Index           =   23
      Left            =   180
      Top             =   1860
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      Caption         =   "质量证明书编号"
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
   Begin InDate.ULabel txt_CERT_NO 
      Height          =   315
      Left            =   180
      Top             =   2160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   1
      BackColor       =   15529975
      BackgroundStyle =   1
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   6405
      Left            =   120
      TabIndex        =   8
      Top             =   2535
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   11298
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "AQC0330C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   6405
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   3855
         _Version        =   393216
         _ExtentX        =   6800
         _ExtentY        =   11298
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0330C.frx":0072
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   6405
         Left            =   3945
         TabIndex        =   10
         Top             =   0
         Width           =   4485
         _Version        =   393216
         _ExtentX        =   7911
         _ExtentY        =   11298
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
         MaxCols         =   9
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0330C.frx":05F0
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   6405
         Left            =   8520
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   6435
         _Version        =   393216
         _ExtentX        =   11351
         _ExtentY        =   11298
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   13
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0330C.frx":0CB7
      End
   End
End
Attribute VB_Name = "AQC0330C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      综合判定结果详细查询成分
'-- Program ID        AQC0330C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.8. 26
'-- Description       综合判定结果详细查询成分
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType   As String           'Form Type
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
Dim sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Const cMin = 2
Private Const cMax = 3
Private Const cAveMin = 4
Private Const cRst1 = 5
Private Const cRst2 = 6
Private Const cRst3 = 7
Private Const cRst4 = 8
Private Const cRst5 = 9
Private Const cRst6 = 10
Private Const cRstAve = 11
Private Const cUnit = 12
Private Const cDsc = 13

Private Const OLD_MAXROWS = 69

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_PROD_NO, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ORD_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_PLAN_STD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ORD_SIZE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_stlgrd, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_ENDUSE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_CUST, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_PROD_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_DSC_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_STDSPEC, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_PROD_SIZE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_SMP_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HEAT_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_PROD_GRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_SUF_GRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_UST_STD_GRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_CERT_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_PRINT_DATE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_PRINT_EMP, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_SHP_ISP_NO, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              
     
    'MASTER Collection
    Mc1.Add Item:="AQC0330C.P_REFER_HEADER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
  
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQC0330C.P_REFER_SS1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        
    
    'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AQC0330C.P_REFER_SS3", Key:="P-R"
    sc3.Add Item:=pColumn1, Key:="pColumn"
    sc3.Add Item:=nColumn1, Key:="nColumn"
    sc3.Add Item:=aColumn1, Key:="aColumn"
    sc3.Add Item:=mColumn1, Key:="mColumn"
    sc3.Add Item:=iColumn1, Key:="iColumn"
    sc3.Add Item:=lColumn1, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
    

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    Proc_Sc.Add Item:=sc3, Key:="Sc3"
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = "冶炼"
    sc3.Item("Spread").Col = 0
    sc3.Item("Spread").Row = 0
    sc3.Item("Spread").Text = "成品"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
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

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(ss2, False)
    Call Gp_Sp_ReadOnlySet(ss2)

    'Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    ss1.RowHeight(-1) = 12.54
    ss1.RowHeight(0) = 24
    ss1.BackColorStyle = BackColorStyleUnderGrid
    ss1.GrayAreaBackColor = &HE0E0E0
    ss1.GridColor = &H808040
    ss1.ShadowColor = &HE1E4CD
    ss1.ShadowDark = &H808040
    ss1.SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
    
    ss2.RowHeight(-1) = 12.54
    ss2.RowHeight(0) = 24
    ss2.BackColorStyle = BackColorStyleUnderGrid
    ss2.GrayAreaBackColor = &HE0E0E0
    ss2.GridColor = &H808040
    ss2.ShadowColor = &HE1E4CD
    ss2.ShadowDark = &H808040
    ss2.SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
    
    ss3.RowHeight(-1) = 12.54
    ss3.RowHeight(0) = 24
    ss3.BackColorStyle = BackColorStyleUnderGrid
    ss3.GrayAreaBackColor = &HE0E0E0
    ss3.GridColor = &H808040
    ss3.ShadowColor = &HE1E4CD
    ss3.ShadowDark = &H808040
    ss3.SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040

    Call MATR_ITEM

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss3, "Q-System.INI", Me.Name)

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
    
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
    Set sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call sub_ss2_data_clear
        txt_PROD_NO.Enabled = True
        pControl(1).SetFocus
        Call Gf_Sp_Cls(Proc_Sc("SC3"))
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err
    
    Dim sQuery As String
    Dim sMesg As String
    Dim AdoRs As adodb.Recordset
    Dim i As Integer
    
    Dim v_chem_rslt, v_chem_rslt_fp, v_chem_diff, v_chem_diff_min, v_chem_diff_max, v_chem_min, v_chem_max As Double
                        
    ss2.ReDraw = True
    
    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
    
        Call Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False)
        Call Gf_Sp_Refer(M_CN1, sc3, Mc1, Mc1("nControl"), Mc1("mControl"), False)

        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        Call subSpreadCheck1
     End If
     
     
    With ss3
        For i = 1 To .MaxRows
            .Row = i
            .Col = 3: v_chem_rslt_fp = Val(.Text)
            .Col = 4: v_chem_rslt = Val(.Text)
            .Col = 5: v_chem_diff = Val(.Text)
            .Col = 6: v_chem_min = Val(.Text)
            .Col = 7: v_chem_max = Val(.Text)
            .Col = 8: v_chem_diff_min = Val(.Text)
            .Col = 9: v_chem_diff_max = Val(.Text)
            
            
'          If v_chem_rslt_fp < v_chem_rslt And v_chem_diff < v_chem_diff_min Then
'            Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, i, i, , &HFF)
'          End If
'
'          If v_chem_rslt_fp > v_chem_rslt And v_chem_diff > v_chem_diff_max Then
'            Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, i, i, , &HFF)
'          End If
           
           If v_chem_rslt_fp < v_chem_min Or v_chem_rslt_fp > v_chem_max Then
             Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, i, i, , &HFF)
           End If
           
            If v_chem_min > 0 And v_chem_rslt_fp = 0 Then
             Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, i, i, , &HFF)
           End If

        Next i

    End With
     
                       
    Set AdoRs = New adodb.Recordset
    
    sQuery = "{call AQC0330C.P_REFER_SS2('" + Trim(txt_PROD_NO.Text) + "','" + txt_SMP_CUT_LOC.Text + "')}"
                    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If AdoRs.BOF Or AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        GoTo Refer_Err
    End If
    
    Call subSpreadDataView(AdoRs)

    AdoRs.Close
'    Set AdoRs = Nothing
'--------------------配置化项目显示 王成  2012.12.14-------------------------------------------------------------------------
    
    Dim ArrayRecords As Variant
    sQuery = "{call AQC0330C.P_SREFER_SS2_CONFIG('" + Trim(txt_PROD_NO.Text) + "','" + Trim(txt_SMP_CUT_LOC.Text) + "')}"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Set AdoRs = M_CN1.Execute(sQuery)
    
    If Not AdoRs.EOF And Not AdoRs.BOF Then
      ArrayRecords = AdoRs.GetRows
      Call subSpreadView_Config(ArrayRecords)
    End If
    
    AdoRs.Close
    Erase ArrayRecords
    
    '----------------------------------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    
    Exit Sub

Refer_Err:
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub opt_SMP_CUT_LOC_Click(Index As Integer)
    If opt_SMP_CUT_LOC(1).Value = True Then
        txt_SMP_CUT_LOC.Text = "B"
    ElseIf opt_SMP_CUT_LOC(2).Value = True Then
        txt_SMP_CUT_LOC.Text = "T"
    ElseIf opt_SMP_CUT_LOC(3).Value = True Then
        txt_SMP_CUT_LOC.Text = "M"
    ElseIf opt_SMP_CUT_LOC(4).Value = True Then
        txt_SMP_CUT_LOC.Text = "A"
    ElseIf opt_SMP_CUT_LOC(5).Value = True Then
        txt_SMP_CUT_LOC.Text = "Y"
    Else
        txt_SMP_CUT_LOC.Text = "B"
    End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub subSpreadCheck1()
    
 Dim i As Long
 Dim j As Long
 

    With ss1
        
        For i = 1 To .MaxRows
            .Row = i
            If GF_GET_CELL_VALUE2(ss1, i, 2) = "" And GF_GET_CELL_VALUE2(ss1, i, 3) = "" And Gf_Get_Cell_Value(ss1, i, 4) = "" Then
                '.Row = i
                .RowHidden = True
            Else
                Call Gp_Sp_CellColor(ss1, 5, i)
                If GF_GET_CELL_VALUE2(ss1, i, 2) <> 0 And GF_GET_CELL_VALUE2(ss1, i, 2) <> "" Then
                   If GF_GET_CELL_VALUE2(ss1, i, 2) > GF_GET_CELL_VALUE2(ss1, i, 5) Or _
                      GF_GET_CELL_VALUE2(ss1, i, 5) = 0 Then
                      Call Gp_Sp_CellColor(ss1, 5, i, RED)
                   End If
                End If
                If GF_GET_CELL_VALUE2(ss1, i, 3) <> 0 And GF_GET_CELL_VALUE2(ss1, i, 3) <> "" Then
                   If GF_GET_CELL_VALUE2(ss1, i, 3) < GF_GET_CELL_VALUE2(ss1, i, 5) Or _
                      GF_GET_CELL_VALUE2(ss1, i, 5) = 0 Then
                      Call Gp_Sp_CellColor(ss1, 5, i, RED)
                   End If
                End If
                
                '成品成分超标检查
                If GF_GET_CELL_VALUE2(ss1, i, 6) <> 0 And GF_GET_CELL_VALUE2(ss1, i, 6) <> "" Then
                   If GF_GET_CELL_VALUE2(ss1, i, 6) > GF_GET_CELL_VALUE2(ss1, i, 8) Or _
                      GF_GET_CELL_VALUE2(ss1, i, 8) = 0 Then
                      Call Gp_Sp_CellColor(ss1, 8, i, RED)
                   End If
                End If
                If GF_GET_CELL_VALUE2(ss1, i, 7) <> 0 And GF_GET_CELL_VALUE2(ss1, i, 7) <> "" Then
                   If GF_GET_CELL_VALUE2(ss1, i, 7) < GF_GET_CELL_VALUE2(ss1, i, 8) Or _
                      GF_GET_CELL_VALUE2(ss1, i, 8) = 0 Then
                      Call Gp_Sp_CellColor(ss1, 8, i, RED)
                   End If
                End If
                
                .RowHidden = False
                
                'J = J + 1
                '.Col = 0: .Text = J
                
            End If
        Next i
                
    End With
   
    
End Sub

Private Sub sub_ss2_data_clear()

    Dim i As Integer

    With ss2
    
        For i = 1 To ss2.MaxRows
        
            .Row = i
            
            .Col = 2: .Text = ""
            .Col = 3: .Text = ""
            .Col = 4: .Text = ""
            .Col = 5: .Text = ""
            .Col = 6: .Text = ""
            .Col = 7: .Text = ""
            .Col = 8: .Text = ""
            .Col = 9: .Text = ""
            .Col = 10: .Text = ""
            .Col = 11: .Text = ""
            .Col = 13: .Text = ""
        
        Next i

    
    End With

End Sub



Private Sub MATR_ITEM()

    Dim i As Integer
    Dim iRow As Integer
    Dim sMatr(69) As String
    Dim sDCSC(69) As String
    
'----------------------------------------------------------------------------------------------- 1
    sMatr(1) = "拉伸试验 - 屈服强度"
    sMatr(2) = "拉伸试验 - 抗拉强度"
    sMatr(3) = "拉伸试验 - 断面收缩率"
    sMatr(4) = "拉伸试验 - 厚度方向断面收缩率"
    sMatr(5) = "拉伸试验 - 断后伸长率"
    sMatr(6) = "拉伸试验 - 屈强比"
    sMatr(7) = "拉伸试验 - 规定非比例伸长应力"
    sMatr(8) = "拉伸试验 - 规定总伸长应力"
    sMatr(9) = "拉伸试验 - 规定残余伸长应力"

'20090805 SUN BIN START
'----------------------------------------------------------------------------------------------- 2
    sMatr(10) = "追加拉伸试验 - 屈服强度"
    sMatr(11) = "追加拉伸试验 - 抗拉强度"
    sMatr(12) = "追加拉伸试验 - 断面收缩率"
    sMatr(13) = "追加拉伸试验 - 断后伸长率"
    sMatr(14) = "追加拉伸试验 - 屈强比"
    sMatr(15) = "追加拉伸试验 - 规定非比例伸长应力"
    sMatr(16) = "追加拉伸试验 - 规定总伸长应力"
    sMatr(17) = "追加拉伸试验 - 规定残余伸长应力"
    
'20090805 SUN BIN END
'----------------------------------------------------------------------------------------------- 3
    sMatr(18) = "高溫拉伸试验 - 屈服强度"
    sMatr(19) = "高溫拉伸试验 - 抗拉强度"
    sMatr(20) = "高溫拉伸试验 - 断后伸长率"
    sMatr(21) = "高溫拉伸试验 - 断面收缩率"
    sMatr(22) = "高溫拉伸试验 - 厚度方向断面收缩率"
    sMatr(23) = "高溫拉伸试验 - 规定非比例伸长应力"
    sMatr(24) = "高溫拉伸试验 - 规定残余伸长应力"

'20090805 SUN BIN START
'----------------------------------------------------------------------------------------------- 4
    sMatr(25) = "追加高溫拉伸试验 - 屈服强度"
    sMatr(26) = "追加高溫拉伸试验 - 抗拉强度"
    sMatr(27) = "追加高溫拉伸试验 - 断后伸长率"
    sMatr(28) = "追加高溫拉伸试验 - 断面收缩率"
    sMatr(29) = "追加高溫拉伸试验 - 规定非比例伸长应力"
    sMatr(30) = "追加高溫拉伸试验 - 规定残余伸长应力"
'20090805 SUN BIN END

'----------------------------------------------------------------------------------------------- 5
    sMatr(31) = "冲击试验"
    sMatr(32) = "冲击试验断面纤维率"
    sMatr(33) = "冲击试验侧向膨胀值"
    sMatr(34) = "追加冲击试验"
    sMatr(35) = "追加冲击试验断面纤维率"
    sMatr(36) = "追加冲击试验侧向膨胀值"
    sMatr(37) = "时效冲击试验"
    sMatr(38) = "时效冲击试验断面纤维率"
    sMatr(39) = "追加时效冲击试验"
    sMatr(40) = "追加时效冲击试验断面纤维率"
    
'----------------------------------------------------------------------------------------------- 6
    sMatr(41) = "硬度试验"
    sMatr(42) = "追加硬度试验"
    sMatr(43) = "弯曲试验"
    sMatr(44) = "追加弯曲试验"
    sMatr(45) = "反复弯曲试验"
    sMatr(46) = "焊接硬度试验"
    sMatr(47) = "焊缝弯曲试验"
    sMatr(48) = "超声波探伤 UST"
    sMatr(49) = "锻平试验"
    sMatr(50) = "淬透性试验"
    sMatr(51) = "抗氢裂能力试验"
    sMatr(52) = "硫化物腐蚀裂纹试验"
    sMatr(53) = "重力撕裂试验"
'----------------------------------------------------------------------------------------------- 7
    sMatr(54) = "脱碳层试验"
    sMatr(55) = "晶粒度试验"
    sMatr(56) = "奥氏体晶粒度试验"
    sMatr(57) = "硫印试验"
    sMatr(58) = "酸浸试验"
    sMatr(59) = "断口检验"
    sMatr(60) = "非金属夹杂试验"
    sMatr(61) = "带状组织"
    sMatr(62) = "NDT重力撕裂试验"
       
'------------------edit by 耿学玉 20110215 uel 、追加UEL、应力比1-5------------------------------------------
    sMatr(63) = "均匀变形伸长率UEL"
    sMatr(64) = "追加均匀变形伸长率UEL"
    sMatr(65) = "追加应力比项目1"
    sMatr(66) = "追加应力比项目2"
    sMatr(67) = "追加应力比项目3"
    sMatr(68) = "追加应力比项目4"
    sMatr(69) = "追加应力比项目5"
       
'----------------------------------------------------------------------------------------------- 1
    sDCSC(1) = "MPa"
    sDCSC(2) = "MPa"
    sDCSC(3) = "%"
    sDCSC(4) = "%"
    sDCSC(5) = "%"
    sDCSC(6) = "%"
    sDCSC(7) = "MPa"
    sDCSC(8) = "MPa"
    sDCSC(9) = "MPa"

'----------------------------------------------------------------------------------------------- 2
    sDCSC(10) = "MPa"
    sDCSC(11) = "MPa"
    sDCSC(12) = "%"
    sDCSC(13) = "%"
    sDCSC(14) = "%"
    sDCSC(15) = "MPa"
    sDCSC(16) = "MPa"
    sDCSC(17) = "MPa"
'----------------------------------------------------------------------------------------------- 3
    sDCSC(18) = "MPa"
    sDCSC(19) = "MPa"
    sDCSC(20) = "%"
    sDCSC(21) = "%"
    sDCSC(22) = "%"
    sDCSC(23) = "MPa"
    sDCSC(24) = "MPa"

'----------------------------------------------------------------------------------------------- 4
    sDCSC(25) = "MPa"
    sDCSC(26) = "MPa"
    sDCSC(27) = "%"
    sDCSC(28) = "%"
    sDCSC(29) = "MPa"
    sDCSC(30) = "MPa"
'----------------------------------------------------------------------------------------------- 5
    sDCSC(31) = "J"
    sDCSC(32) = "%"
    sDCSC(33) = "%"
    sDCSC(34) = "J"
    sDCSC(35) = "%"
    sDCSC(36) = "%"
    sDCSC(37) = "J/cm2"
    sDCSC(38) = "%"
    sDCSC(39) = "J/cm2"
    sDCSC(40) = "%"
'----------------------------------------------------------------------------------------------- 6
    sDCSC(41) = ""
    sDCSC(42) = ""
    sDCSC(43) = ""
    sDCSC(44) = ""
    sDCSC(45) = "次"
    sDCSC(46) = ""
    sDCSC(47) = ""
    sDCSC(48) = "级"
    sDCSC(49) = ""
    sDCSC(50) = ""
    sDCSC(51) = ""
    sDCSC(52) = "%"
    sDCSC(53) = "%"
'----------------------------------------------------------------------------------------------- 7
    sDCSC(54) = "mm"
    sDCSC(55) = "级"
    sDCSC(56) = "级"
    sDCSC(57) = "级"
    sDCSC(58) = "级"
    sDCSC(59) = "级"
    sDCSC(60) = "级"
    sDCSC(61) = "级"
    sDCSC(62) = ""
'------------------edit by 耿学玉 20110215 uel 、追加UEL、应力比1-5------------------------------------------
    sDCSC(63) = "%"
    sDCSC(64) = "%"
    sDCSC(65) = ""
    sDCSC(66) = ""
    sDCSC(67) = ""
    sDCSC(68) = ""
    sDCSC(69) = ""
   
    
    With ss2
    
        .MaxRows = 1
        .MaxRows = 69
    
        For i = 1 To 69
            .Row = i
            .Col = 1: .Text = sMatr(i)
            .Col = 12: .Text = sDCSC(i)
        Next i

    End With
    
    
    Call subSpreadCheck2

End Sub


Private Sub subSpreadDataView(ByRef AdoRs As adodb.Recordset)

    'On Error GoTo Refer_Err
'-------------------------------------------------------------------------------------------------- 1
    
'屈服强度 - 01
    Call GP_SET_CELL_VALUE(ss2, 1, cMin, AdoRs("YP_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 1, cMax, AdoRs("YP_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 1, cRst1, AdoRs("YP_RST"))
    Call GP_SET_CELL_VALUE(ss2, 1, cDsc, AdoRs("YP_DSC_RST"))

'抗拉强度 - 02
    Call GP_SET_CELL_VALUE(ss2, 2, cMin, AdoRs("TS_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 2, cMax, AdoRs("TS_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 2, cRst1, AdoRs("TS_RST"))
    Call GP_SET_CELL_VALUE(ss2, 2, cDsc, AdoRs("TS_DSC_RST"))
                                                                                                                                                                                                                                                             
'断面收缩率 - 03
    Call GP_SET_CELL_VALUE(ss2, 3, cMin, AdoRs("RA_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 3, cMax, AdoRs("RA_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 3, cRst1, AdoRs("RA_RST"))
    Call GP_SET_CELL_VALUE(ss2, 3, cRst2, AdoRs("RA_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 3, cRst3, AdoRs("RA_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 3, cRstAve, AdoRs("RA_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 3, cDsc, AdoRs("RA_DSC_RST"))
    
    '厚度方向断面收缩率 louyannan 20101124
    Call GP_SET_CELL_VALUE(ss2, 4, cMin, AdoRs("ZRA_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 4, cAveMin, AdoRs("ZRA_AVE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 4, cRst1, AdoRs("ZRA_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 4, cRst2, AdoRs("ZRA_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 4, cRst3, AdoRs("ZRA_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 4, cRstAve, AdoRs("ZRA_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 4, cDsc, AdoRs("ZRA_DSC_RST"))
    
                                                                                                                                                                                                                                                             
'断后伸长率 - 04
    Call GP_SET_CELL_VALUE(ss2, 5, cMin, AdoRs("EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 5, cMax, AdoRs("EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 5, cRst1, AdoRs("EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 5, cDsc, AdoRs("EL_DSC_RST"))
                                                                                                                                                                                                                                                             
                                                                                                                                                                                                                                                             
                                                                                                                                                                                                                                                             
'屈强比 - 05
     Call GP_SET_CELL_VALUE(ss2, 6, cMin, AdoRs("YR_MIN"))
     Call GP_SET_CELL_VALUE(ss2, 6, cMax, AdoRs("YR_MAX"))
     Call GP_SET_CELL_VALUE(ss2, 6, cRst1, AdoRs("YR_RST"))
     Call GP_SET_CELL_VALUE(ss2, 6, cDsc, AdoRs("YR_DSC_RST"))
                                                                                                                                                                                                                                                             
                                                                                                                                                                                                                                                             
'规定非比例伸长应力 - 06
    Call GP_SET_CELL_VALUE(ss2, 7, cMin, AdoRs("SNPP_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 7, cMax, AdoRs("SNPP_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 7, cRst1, AdoRs("SNPP_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 7, cDsc, AdoRs("SNPP_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'规定总伸长应力 - 07
    Call GP_SET_CELL_VALUE(ss2, 8, cMin, AdoRs("SG_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 8, cMax, AdoRs("SG_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 8, cRst1, AdoRs("SG_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 8, cDsc, AdoRs("SG_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'规定残余伸长应力 - 08
    Call GP_SET_CELL_VALUE(ss2, 9, cMin, AdoRs("SP_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 9, cMax, AdoRs("SP_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 9, cRst1, AdoRs("SP_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 9, cDsc, AdoRs("SP_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'-------------------------------------------------------------------------------------------------- 2
'20090805 SUN BIN START
    '屈服强度 - 09
    Call GP_SET_CELL_VALUE(ss2, 10, cMin, AdoRs("A_YP_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 10, cMax, AdoRs("A_YP_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 10, cRst1, AdoRs("A_YP_RST"))
    Call GP_SET_CELL_VALUE(ss2, 10, cDsc, AdoRs("A_YP_DSC_RST"))

'抗拉强度 - 10
    Call GP_SET_CELL_VALUE(ss2, 11, cMin, AdoRs("A_TS_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 11, cMax, AdoRs("A_TS_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 11, cRst1, AdoRs("A_TS_RST"))
    Call GP_SET_CELL_VALUE(ss2, 11, cDsc, AdoRs("A_TS_DSC_RST"))
                                                                                                                                                                                                                                                           
'断面收缩率 - 11
    Call GP_SET_CELL_VALUE(ss2, 12, cMin, AdoRs("A_RA_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 12, cMax, AdoRs("A_RA_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 12, cRst1, AdoRs("A_RA_RST"))
    Call GP_SET_CELL_VALUE(ss2, 12, cRst2, AdoRs("A_RA_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 12, cRst3, AdoRs("A_RA_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 12, cRstAve, AdoRs("A_RA_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 12, cDsc, AdoRs("A_RA_DSC_RST"))
                                                                                                                                                                                                                                                           
'断后伸长率 - 12
    Call GP_SET_CELL_VALUE(ss2, 13, cMin, AdoRs("A_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 13, cMax, AdoRs("A_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 13, cRst1, AdoRs("A_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 13, cDsc, AdoRs("A_EL_DSC_RST"))
                                                                                                                                                                                                                                              
'屈强比 - 13
     Call GP_SET_CELL_VALUE(ss2, 14, cMin, AdoRs("A_YR_MIN"))
     Call GP_SET_CELL_VALUE(ss2, 14, cMax, AdoRs("A_YR_MAX"))
     Call GP_SET_CELL_VALUE(ss2, 14, cRst1, AdoRs("A_YR_RST"))
     Call GP_SET_CELL_VALUE(ss2, 14, cDsc, AdoRs("A_YR_DSC_RST"))
                                                                                                                                                                                                                                                           
                                                                                                                                                                                                                                                             
'规定非比例伸长应力 - 14
    Call GP_SET_CELL_VALUE(ss2, 15, cMin, AdoRs("A_SNPP_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 15, cMax, AdoRs("A_SNPP_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 15, cRst1, AdoRs("A_SNPP_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 15, cDsc, AdoRs("A_SNPP_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'规定总伸长应力 - 15
    Call GP_SET_CELL_VALUE(ss2, 16, cMin, AdoRs("A_SG_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 16, cMax, AdoRs("A_SG_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 16, cRst1, AdoRs("A_SG_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 16, cDsc, AdoRs("A_SG_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'规定残余伸长应力 - 16
    Call GP_SET_CELL_VALUE(ss2, 17, cMin, AdoRs("A_SP_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 17, cMax, AdoRs("A_SP_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 17, cRst1, AdoRs("A_SP_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 17, cDsc, AdoRs("A_SP_EL_DSC_RST"))
'20090805 SUN BIN END
                                                                                                                                                                                                                                                             
'-------------------------------------------------------------------------------------------------- 3

'屈服强度 - 17
    Call GP_SET_CELL_VALUE(ss2, 18, cMin, AdoRs("HGT_YP_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 18, cMax, AdoRs("HGT_YP_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 18, cRst1, AdoRs("HGT_YP_RST"))
    Call GP_SET_CELL_VALUE(ss2, 18, cDsc, AdoRs("HGT_YP_DSC_RST"))
                                                                                                                                                                                                                                                             
'抗拉强度 - 18
    Call GP_SET_CELL_VALUE(ss2, 19, cMin, AdoRs("HGT_TS_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 19, cMax, AdoRs("HGT_TS_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 19, cRst1, AdoRs("HGT_TS_RST"))
    Call GP_SET_CELL_VALUE(ss2, 19, cDsc, AdoRs("HGT_TS_DSC_RST"))
    
    
' 断后伸长率 -20
    Call GP_SET_CELL_VALUE(ss2, 20, cMin, AdoRs("HGT_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 20, cMax, AdoRs("HGT_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 20, cRst1, AdoRs("HGT_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 20, cDsc, AdoRs("HGT_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'断面收缩率 - 19
    Call GP_SET_CELL_VALUE(ss2, 21, cMin, AdoRs("HGT_RA_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 21, cMax, AdoRs("HGT_RA_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 21, cRst1, AdoRs("HGT_RA_RST"))
    Call GP_SET_CELL_VALUE(ss2, 21, cRst1, AdoRs("HGT_RA_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 21, cRst1, AdoRs("HGT_RA_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 21, cRst1, AdoRs("HGT_RA_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 21, cDsc, AdoRs("HGT_RA_DSC_RST"))
    
  
    
 '厚度方向断面收缩率 louyannan 20101124
    Call GP_SET_CELL_VALUE(ss2, 22, cMin, AdoRs("HGT_ZRA_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 22, cAveMin, AdoRs("HGT_ZRA_AVE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 22, cRst1, AdoRs("HGT_ZRA_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 22, cRst2, AdoRs("HGT_ZRA_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 22, cRst3, AdoRs("HGT_ZRA_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 22, cRstAve, AdoRs("HGT_ZRA_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 22, cDsc, AdoRs("HGT_ZRA_DSC_RST"))
                                                                                                                                                                                                                                                             
                    

                                                                                                                                                                                                                                                             
'规定非比例伸长应力 - 21
    Call GP_SET_CELL_VALUE(ss2, 23, cMin, AdoRs("HGT_SNPP_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 23, cMax, AdoRs("HGT_SNPP_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 23, cRst1, AdoRs("HGT_SNPP_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 23, cDsc, AdoRs("HGT_SNPP_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'规定残余伸长应力 - 22
    Call GP_SET_CELL_VALUE(ss2, 24, cMin, AdoRs("HGT_SP_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 24, cMax, AdoRs("HGT_SP_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 24, cRst1, AdoRs("HGT_SP_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 24, cDsc, AdoRs("HGT_SP_EL_DSC_RST"))
    
                                                                                                                                                                                                                                                             
'-------------------------------------------------------------------------------------------------- 4

'屈服强度 - 23
    Call GP_SET_CELL_VALUE(ss2, 25, cMin, AdoRs("A_HGT_YP_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 25, cMax, AdoRs("A_HGT_YP_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 25, cRst1, AdoRs("A_HGT_YP_RST"))
    Call GP_SET_CELL_VALUE(ss2, 25, cDsc, AdoRs("A_HGT_YP_DSC_RST"))
                                                                                                                                                                                                                                                             
'抗拉强度 - 24
    Call GP_SET_CELL_VALUE(ss2, 26, cMin, AdoRs("A_HGT_TS_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 26, cMax, AdoRs("A_HGT_TS_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 26, cRst1, AdoRs("A_HGT_TS_RST"))
    Call GP_SET_CELL_VALUE(ss2, 26, cDsc, AdoRs("A_HGT_TS_DSC_RST"))
                                                                                                                                                                                                                                                             
'断面收缩率 - 25
    Call GP_SET_CELL_VALUE(ss2, 27, cMin, AdoRs("A_HGT_RA_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 27, cMax, AdoRs("A_HGT_RA_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 27, cRst1, AdoRs("A_HGT_RA_RST"))
    Call GP_SET_CELL_VALUE(ss2, 27, cRst1, AdoRs("A_HGT_RA_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 27, cRst1, AdoRs("A_HGT_RA_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 27, cRst1, AdoRs("A_HGT_RA_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 27, cDsc, AdoRs("A_HGT_RA_DSC_RST"))
                                                                                                                                                                                                                                                             
'断后伸长率 - 26
    Call GP_SET_CELL_VALUE(ss2, 28, cMin, AdoRs("A_HGT_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 28, cMax, AdoRs("A_HGT_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 28, cRst1, AdoRs("A_HGT_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 28, cDsc, AdoRs("A_HGT_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'规定非比例伸长应力 - 27
    Call GP_SET_CELL_VALUE(ss2, 29, cMin, AdoRs("A_HGT_SNPP_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 29, cMax, AdoRs("A_HGT_SNPP_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 29, cRst1, AdoRs("A_HGT_SNPP_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 29, cDsc, AdoRs("A_HGT_SNPP_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'规定残余伸长应力 - 28
    Call GP_SET_CELL_VALUE(ss2, 30, cMin, AdoRs("A_HGT_SP_EL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 30, cMax, AdoRs("A_HGT_SP_EL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 30, cRst1, AdoRs("A_HGT_SP_EL_RST"))
    Call GP_SET_CELL_VALUE(ss2, 30, cDsc, AdoRs("A_HGT_SP_EL_DSC_RST"))
                                                                                                                                                                                                                                                             
'-------------------------------------------------------------------------------------------------- 5

                                                                                                                                                                                                                                                             
'冲击试验 - 29
    Call GP_SET_CELL_VALUE(ss2, 31, cMin, AdoRs("IMPACT_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 31, cAveMin, AdoRs("IMPACT_AVE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 31, cRst1, AdoRs("IMPACT_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 31, cRst2, AdoRs("IMPACT_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 31, cRst3, AdoRs("IMPACT_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 31, cRst4, AdoRs("IMPACT_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 31, cRst5, AdoRs("IMPACT_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 31, cRst6, AdoRs("IMPACT_RST6"))
    Call GP_SET_CELL_VALUE(ss2, 31, cRstAve, AdoRs("IMPACT_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 31, cDsc, AdoRs("IMPACT_DSC_RST"))
                                                                                                                                                                                                                                                             
'冲击试验 - 断面纤维率 - 30
    Call GP_SET_CELL_VALUE(ss2, 32, cMin, AdoRs("IMPACT_RATE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 32, cMax, AdoRs("IMPACT_RATE_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 32, cRst1, AdoRs("IMPACT_RATE_RST"))
    Call GP_SET_CELL_VALUE(ss2, 32, cRst2, AdoRs("IMPACT_RATE_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 32, cRst3, AdoRs("IMPACT_RATE_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 32, cRst4, AdoRs("IMPACT_RATE_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 32, cRst5, AdoRs("IMPACT_RATE_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 32, cRst6, AdoRs("IMPACT_RATE_RST6"))
    Call GP_SET_CELL_VALUE(ss2, 32, cRstAve, AdoRs("IMPACT_RATE_RST_AVE"))
   '冲击-侧膨胀 louyannan 20101124
    Call GP_SET_CELL_VALUE(ss2, 33, cMin, AdoRs("SID_EXPAIN_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 33, cAveMin, AdoRs("SID_EXPAIN_AVE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 33, cRst1, AdoRs("SID_EXPAIN_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 33, cRst2, AdoRs("SID_EXPAIN_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 33, cRst3, AdoRs("SID_EXPAIN_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 33, cRst4, AdoRs("SID_EXPAIN_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 33, cRst5, AdoRs("SID_EXPAIN_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 33, cRst6, AdoRs("SID_EXPAIN_RST6"))
    Call GP_SET_CELL_VALUE(ss2, 33, cRstAve, AdoRs("SID_EXPAIN_RST_AVE"))
                                                                                                                                                                                                                                                             
'追加冲击试验 - 31
    Call GP_SET_CELL_VALUE(ss2, 34, cMin, AdoRs("A_IMPACT_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 34, cAveMin, AdoRs("A_IMPACT_AVE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 34, cRst1, AdoRs("A_IMPACT_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 34, cRst2, AdoRs("A_IMPACT_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 34, cRst3, AdoRs("A_IMPACT_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 34, cRst4, AdoRs("A_IMPACT_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 34, cRst5, AdoRs("A_IMPACT_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 34, cRst6, AdoRs("A_IMPACT_RST6"))
    Call GP_SET_CELL_VALUE(ss2, 34, cRstAve, AdoRs("A_IMPACT_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 34, cDsc, AdoRs("A_IMPACT_DSC_RST"))
                                                                                                                                                                                                                                                             
'追加冲击试验 - 断面纤维率 - 32
    Call GP_SET_CELL_VALUE(ss2, 35, cMin, AdoRs("A_IMPACT_RATE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 35, cMax, AdoRs("A_IMPACT_RATE_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 35, cRst1, AdoRs("A_IMPACT_RATE_RST"))
    Call GP_SET_CELL_VALUE(ss2, 35, cRst2, AdoRs("A_IMPACT_RATE_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 35, cRst3, AdoRs("A_IMPACT_RATE_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 35, cRst4, AdoRs("A_IMPACT_RATE_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 35, cRst5, AdoRs("A_IMPACT_RATE_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 35, cRst6, AdoRs("A_IMPACT_RATE_RST6"))
    Call GP_SET_CELL_VALUE(ss2, 35, cRstAve, AdoRs("A_IMPACT_RATE_RST_AVE"))
    
     '冲击-侧膨胀 louyannan 20101124
    Call GP_SET_CELL_VALUE(ss2, 36, cMin, AdoRs("A_SID_EXPAIN_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 36, cAveMin, AdoRs("A_SID_EXPAIN_AVE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 36, cRst1, AdoRs("A_SID_EXPAIN_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 36, cRst2, AdoRs("A_SID_EXPAIN_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 36, cRst3, AdoRs("A_SID_EXPAIN_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 36, cRst4, AdoRs("A_SID_EXPAIN_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 36, cRst5, AdoRs("A_SID_EXPAIN_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 36, cRst6, AdoRs("A_SID_EXPAIN_RST6"))
    Call GP_SET_CELL_VALUE(ss2, 36, cRstAve, AdoRs("A_SID_EXPAIN_RST_AVE"))
                                                                                                                                                                                                                                                             
'时效冲击试验 - 33
    Call GP_SET_CELL_VALUE(ss2, 37, cMin, AdoRs("TIM_IMPACT_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 37, cAveMin, AdoRs("TIM_IMPACT_AVE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 37, cRst1, AdoRs("TIM_IMPACT_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 37, cRst2, AdoRs("TIM_IMPACT_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 37, cRst3, AdoRs("TIM_IMPACT_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 37, cRst4, AdoRs("TIM_IMPACT_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 37, cRst5, AdoRs("TIM_IMPACT_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 37, cRst6, AdoRs("TIM_IMPACT_RST6"))
    Call GP_SET_CELL_VALUE(ss2, 37, cRstAve, AdoRs("TIM_IMPACT_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 37, cDsc, AdoRs("TIM_IMPACT_DSC_RST"))
                                                                                                                                                                                                                                                             
'时效冲击试验 - 断面纤维率 - 34
    Call GP_SET_CELL_VALUE(ss2, 38, cMin, AdoRs("TIM_IMPACT_RATE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 38, cMax, AdoRs("TIM_IMPACT_RATE_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 38, cRst1, AdoRs("TIM_IMPACT_RATE_RST"))
                                                                                                                                                                                                                                                             
'追加时效冲击试验- 35
    Call GP_SET_CELL_VALUE(ss2, 39, cMin, AdoRs("A_TIM_IMPACT_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 39, cAveMin, AdoRs("A_TIM_IMPACT_AVE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 39, cRst1, AdoRs("A_TIM_IMPACT_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 39, cRst2, AdoRs("A_TIM_IMPACT_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 39, cRst3, AdoRs("A_TIM_IMPACT_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 39, cRst4, AdoRs("A_TIM_IMPACT_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 39, cRst5, AdoRs("A_TIM_IMPACT_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 39, cRst6, AdoRs("A_TIM_IMPACT_RST6"))
    Call GP_SET_CELL_VALUE(ss2, 39, cRstAve, AdoRs("A_TIM_IMPACT_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 39, cDsc, AdoRs("A_TIM_IMPACT_DSC_RST"))
                                                                                                                                                                                                                                                             
'追加时效冲击试验 - 断面纤维率 - 36
    Call GP_SET_CELL_VALUE(ss2, 40, cMin, AdoRs("A_TIM_IMPACT_RATE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 40, cMax, AdoRs("A_TIM_IMPACT_RATE_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 40, cRst1, AdoRs("A_TIM_IMPACT_RATE_RST"))
                                                                                                                                                                                                                                                             
'-------------------------------------------------------------------------------------------------- 6
                                                                                                                                                                                                                                                             
'硬度- 37
    Call GP_SET_CELL_VALUE(ss2, 41, cMin, AdoRs("HARD_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 41, cMax, AdoRs("HARD_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 41, cRst1, AdoRs("HARD_RST"))
    Call GP_SET_CELL_VALUE(ss2, 41, cDsc, AdoRs("HARD_DSC_RST"))
    
'追加硬度- 38
    Call GP_SET_CELL_VALUE(ss2, 42, cMin, AdoRs("A_HARD_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 42, cMax, AdoRs("A_HARD_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 42, cRst1, AdoRs("A_HARD_RST"))
    Call GP_SET_CELL_VALUE(ss2, 42, cDsc, AdoRs("A_HARD_DSC_RST"))
                                                                                                                                                                                                                                                             
'弯曲试验 - 39
    Call GP_SET_CELL_VALUE(ss2, 43, cRst1, AdoRs("BEND_RST"))
    Call GP_SET_CELL_VALUE(ss2, 43, cDsc, AdoRs("BEND_DSC_RST"))
    
'追加弯曲试验 - 40
    Call GP_SET_CELL_VALUE(ss2, 44, cRst1, AdoRs("A_BEND_RST"))
    Call GP_SET_CELL_VALUE(ss2, 44, cDsc, AdoRs("A_BEND_DSC_RST"))
                                                                                                                                                                                                                                                             
'反复弯曲 - 41
    Call GP_SET_CELL_VALUE(ss2, 45, cMin, AdoRs("RPT_BEND_TMS"))
    Call GP_SET_CELL_VALUE(ss2, 45, cRst1, AdoRs("RPT_BEND_RST"))
    Call GP_SET_CELL_VALUE(ss2, 45, cDsc, AdoRs("RPT_BEND_DSC_RST"))
                                                                                                                                                                                                                                                             
'焊缝硬度 - 42
    Call GP_SET_CELL_VALUE(ss2, 46, cMin, AdoRs("WLD_HARD_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 46, cMax, AdoRs("WLD_HARD_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 46, cRst1, AdoRs("WLD_HARD_RST"))
    Call GP_SET_CELL_VALUE(ss2, 46, cUnit, AdoRs("WLD_HARD_UNIT"))
    Call GP_SET_CELL_VALUE(ss2, 46, cDsc, AdoRs("WLD_HARD_DSC_RST"))
                                                                                                                                                                                                                                                             
'焊缝弯曲 - 43
    Call GP_SET_CELL_VALUE(ss2, 47, cRst1, AdoRs("WLD_BEND_RST"))
    Call GP_SET_CELL_VALUE(ss2, 47, cDsc, AdoRs("WLD_BEND_DSC_RST"))
                                                                                                                                                                                                                                                             
'超声波探伤（UST）- 44
    Call GP_SET_CELL_VALUE(ss2, 48, cMin, AdoRs("UST_GRD"))
    Call GP_SET_CELL_VALUE(ss2, 48, cRst1, AdoRs("UST_GRD_RST"))
    Call GP_SET_CELL_VALUE(ss2, 48, cDsc, AdoRs("UST_DSC_RST"))
                                                                                                                                                                                                                                                             
'锻平 - 45
    Call GP_SET_CELL_VALUE(ss2, 49, cRst1, AdoRs("FOAT_RST"))
    Call GP_SET_CELL_VALUE(ss2, 49, cDsc, AdoRs("FOAT_DSC_RST"))
                                                                                                                                                                                                                                                             
'淬透性 - 46
    Call GP_SET_CELL_VALUE(ss2, 50, cMin, AdoRs("JOMINY_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 50, cMax, AdoRs("JOMINY_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 50, cRst1, AdoRs("JOMINY_RST_TOP"))
    Call GP_SET_CELL_VALUE(ss2, 50, cRst2, AdoRs("JOMINY_RST_TOP2"))
    Call GP_SET_CELL_VALUE(ss2, 50, cRst3, AdoRs("JOMINY_RST_TOP3"))
    Call GP_SET_CELL_VALUE(ss2, 50, cDsc, AdoRs("JOMINY_DSC_RST"))
                                                                                                                                                                                                                                                             
'抗氢裂能力 - 47
    Call GP_SET_CELL_VALUE2(ss2, 51, cMax, 3, AdoRs("HIC_CSR_MAX"), AdoRs("HIC_CLR_MAX"), AdoRs("HIC_CTR_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 51, cRst1, AdoRs("HIC_CSR_RST"))
    Call GP_SET_CELL_VALUE(ss2, 51, cRst2, AdoRs("HIC_CLR_RST"))
    Call GP_SET_CELL_VALUE(ss2, 51, cRst3, AdoRs("HIC_CTR_RST"))
    Call GP_SET_CELL_VALUE(ss2, 51, cDsc, AdoRs("HIC_CWR_DSC_RST"))
                                                                                                                                                                                                                                                             
'硫化物腐蚀裂纹 - 48
    Call GP_SET_CELL_VALUE(ss2, 52, cMax, AdoRs("SSCC_YP_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 52, cRst1, AdoRs("SSCC_YP_RST"))
    Call GP_SET_CELL_VALUE(ss2, 52, cDsc, AdoRs("SSCC_YP_DSC_RST"))
                                                                                                                                                                                                                                                             
'重力撕裂试验 - 49
    Call GP_SET_CELL_VALUE(ss2, 53, cMin, AdoRs("DWTT_YP_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 53, cAveMin, AdoRs("DWTT_YP_AVE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 53, cRst1, AdoRs("DWTT_YP_RST"))
    Call GP_SET_CELL_VALUE(ss2, 53, cRst2, AdoRs("DWTT_YP_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 53, cRstAve, AdoRs("DWTT_YP_RST_AVE"))
    Call GP_SET_CELL_VALUE(ss2, 53, cDsc, AdoRs("DWTT_YP_DSC_RST"))
                                                                                                                                                                                                                                                             
'-------------------------------------------------------------------------------------------------- 5

                                                                                                                                                                                                                                                             
'脱碳层 - 50
    Call GP_SET_CELL_VALUE(ss2, 54, cMax, AdoRs("RMV_CAR_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 54, cRst1, AdoRs("RMV_CAR_RST"))
    Call GP_SET_CELL_VALUE(ss2, 54, cDsc, AdoRs("RMV_CAR_DSC_RST"))
                                                                                                                                                                                                                                                             
'晶粒度 - 51
    Call GP_SET_CELL_VALUE(ss2, 55, cMin, AdoRs("GRAIN_SIZE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 55, cMax, AdoRs("GRAIN_SIZE_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 55, cRst1, AdoRs("GRAIN_SIZE_RST"))
    Call GP_SET_CELL_VALUE(ss2, 55, cDsc, AdoRs("GRAIN_SIZE_DSC_RST"))
    
'奥氏体晶粒度 - 52
    Call GP_SET_CELL_VALUE(ss2, 56, cMin, AdoRs("OST_GRAIN_SIZE_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 56, cMax, AdoRs("OST_GRAIN_SIZE_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 56, cRst1, AdoRs("OST_GRAIN_SIZE_RST"))
    Call GP_SET_CELL_VALUE(ss2, 56, cDsc, AdoRs("OST_GRAIN_SIZE_DSC_RST"))
                                                                                                                                                                                                                                                             
'硫印 - 53
    Call GP_SET_CELL_VALUE(ss2, 57, cMax, AdoRs("S_PRINT_DRG"))
    Call GP_SET_CELL_VALUE(ss2, 57, cRst1, AdoRs("S_PRINT_RST"))
    Call GP_SET_CELL_VALUE(ss2, 57, cDsc, AdoRs("S_PRINT_DSC_RST"))
                                                                                                                                                                                                                                                             
'酸浸检验 - 54
    Call GP_SET_CELL_VALUE2(ss2, 58, cMax, 5, AdoRs("ACD_DFT_GRD1"), AdoRs("ACD_DFT_GRD2"), AdoRs("ACD_DFT_GRD3"), AdoRs("ACD_DFT_GRD4"), AdoRs("ACD_DFT_GRD5"))
    Call GP_SET_CELL_VALUE(ss2, 58, cRst1, AdoRs("ACD_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 58, cRst2, AdoRs("ACD_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 58, cRst3, AdoRs("ACD_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 58, cRst4, AdoRs("ACD_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 58, cRst5, AdoRs("ACD_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 58, cDsc, AdoRs("ACD_DSC_RST"))
                                                                                                                                                                                                                                                             
'断口检验 - 55
    Call GP_SET_CELL_VALUE2(ss2, 59, cMax, 5, AdoRs("FRACT_GRD1"), AdoRs("FRACT_GRD2"), AdoRs("FRACT_GRD3"), AdoRs("FRACT_GRD4"), AdoRs("FRACT_GRD5"))
    Call GP_SET_CELL_VALUE(ss2, 59, cRst1, AdoRs("FRACT_GRD_RST1"))
    Call GP_SET_CELL_VALUE(ss2, 59, cRst2, AdoRs("FRACT_GRD_RST2"))
    Call GP_SET_CELL_VALUE(ss2, 59, cRst3, AdoRs("FRACT_GRD_RST3"))
    Call GP_SET_CELL_VALUE(ss2, 59, cRst4, AdoRs("FRACT_GRD_RST4"))
    Call GP_SET_CELL_VALUE(ss2, 59, cRst5, AdoRs("FRACT_GRD_RST5"))
    Call GP_SET_CELL_VALUE(ss2, 59, cDsc, AdoRs("FRACT_GRD_RST"))
                                                                                                                                                                                                                                                             
'非金属夹杂 - 56
    Call GP_SET_CELL_VALUE2(ss2, 60, cMax, 10, AdoRs("NON_METAL_AGRD1"), AdoRs("NON_METAL_AGRD2"), AdoRs("NON_METAL_AGRD3"), AdoRs("NON_METAL_AGRD4"), AdoRs("NON_METAL_BGRD1"), AdoRs("NON_METAL_BGRD2"), AdoRs("NON_METAL_BGRD3"), AdoRs("NON_METAL_BGRD4"), AdoRs("NON_METAL_DS_GRD"), AdoRs("NON_METAL_TIN_GRD"))
    Call GP_SET_CELL_VALUE2(ss2, 60, cRst1, 2, AdoRs("NON_METAL_ARST1"), AdoRs("NON_METAL_ARST2"))
    Call GP_SET_CELL_VALUE2(ss2, 60, cRst2, 2, AdoRs("NON_METAL_ARST3"), AdoRs("NON_METAL_ARST4"))
    Call GP_SET_CELL_VALUE2(ss2, 60, cRst3, 2, AdoRs("NON_METAL_BRST1"), AdoRs("NON_METAL_BRST2"))
    Call GP_SET_CELL_VALUE2(ss2, 60, cRst4, 2, AdoRs("NON_METAL_BRST3"), AdoRs("NON_METAL_BRST4"))
    Call GP_SET_CELL_VALUE2(ss2, 60, cRst5, 2, AdoRs("NON_METAL_DS_RST"), AdoRs("NON_METAL_TIN_RST"))
    Call GP_SET_CELL_VALUE(ss2, 60, cDsc, AdoRs("NON_MATAL_DSC_RST"))
                                                                                                                                                                                                                                                             
'带状组织 - 57
    Call GP_SET_CELL_VALUE(ss2, 61, cMax, AdoRs("BELT_STR_GRD"))
    Call GP_SET_CELL_VALUE(ss2, 61, cRst1, AdoRs("BELT_STR_GRD_RST"))
    Call GP_SET_CELL_VALUE(ss2, 61, cDsc, AdoRs("BELT_STR_DSC_RST"))
    '---------------------------------------------------LOUYANNAN 20101202-----------------------------------------------
     Call GP_SET_CELL_VALUE(ss2, 62, cRst1, AdoRs("NDT_RST"))
     Call GP_SET_CELL_VALUE(ss2, 62, cDsc, AdoRs("NDT_DSC_RST"))
     
     Call GP_SET_CELL_VALUE(ss2, 62, cDsc, AdoRs("NDT_DSC_RST"))
     Call GP_SET_CELL_VALUE(ss2, 62, cDsc, AdoRs("NDT_DSC_RST"))
     Call GP_SET_CELL_VALUE(ss2, 62, cDsc, AdoRs("NDT_DSC_RST"))
     
'------------------------------------------------------------
'UEL 均匀变形伸长率UEL - 63
    Call GP_SET_CELL_VALUE(ss2, 63, cMin, AdoRs("UEL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 63, cMax, AdoRs("UEL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 63, cRst1, AdoRs("UEL_VAL"))
    Call GP_SET_CELL_VALUE(ss2, 63, cDsc, AdoRs("UEL_DSC_RST"))
    
'UEL 追加均匀变形伸长率UEL - 64
    Call GP_SET_CELL_VALUE(ss2, 64, cMin, AdoRs("A_UEL_MIN"))
    Call GP_SET_CELL_VALUE(ss2, 64, cMax, AdoRs("A_UEL_MAX"))
    Call GP_SET_CELL_VALUE(ss2, 64, cRst1, AdoRs("A_UEL_VAL"))
    Call GP_SET_CELL_VALUE(ss2, 64, cDsc, AdoRs("A_UEL_DSC_RST"))
    
'应力比项目1 - 65
    Call GP_SET_CELL_VALUE(ss2, 65, cMin, AdoRs("A_STRESS_RT_MIN1"))
    Call GP_SET_CELL_VALUE(ss2, 65, cMax, AdoRs("A_STRESS_RT_MAX1"))
    Call GP_SET_CELL_VALUE(ss2, 65, cRst1, AdoRs("A_STRESS_RT_VAL1"))
    Call GP_SET_CELL_VALUE(ss2, 65, cDsc, AdoRs("A_STRESS_RT_DSC_RST1"))
    
'应力比项目1 - 66
    Call GP_SET_CELL_VALUE(ss2, 66, cMin, AdoRs("A_STRESS_RT_MIN2"))
    Call GP_SET_CELL_VALUE(ss2, 66, cMax, AdoRs("A_STRESS_RT_MAX2"))
    Call GP_SET_CELL_VALUE(ss2, 66, cRst1, AdoRs("A_STRESS_RT_VAL2"))
    Call GP_SET_CELL_VALUE(ss2, 66, cDsc, AdoRs("A_STRESS_RT_DSC_RST2"))
    
    
'应力比项目1 - 67
    Call GP_SET_CELL_VALUE(ss2, 67, cMin, AdoRs("A_STRESS_RT_MIN3"))
    Call GP_SET_CELL_VALUE(ss2, 67, cMax, AdoRs("A_STRESS_RT_MAX3"))
    Call GP_SET_CELL_VALUE(ss2, 67, cRst1, AdoRs("A_STRESS_RT_VAL3"))
    Call GP_SET_CELL_VALUE(ss2, 67, cDsc, AdoRs("A_STRESS_RT_DSC_RST3"))
    
'应力比项目1 - 68
    Call GP_SET_CELL_VALUE(ss2, 68, cMin, AdoRs("A_STRESS_RT_MIN4"))
    Call GP_SET_CELL_VALUE(ss2, 68, cMax, AdoRs("A_STRESS_RT_MAX4"))
    Call GP_SET_CELL_VALUE(ss2, 68, cRst1, AdoRs("A_STRESS_RT_VAL4"))
    Call GP_SET_CELL_VALUE(ss2, 68, cDsc, AdoRs("A_STRESS_RT_DSC_RST4"))
    
'应力比项目1 - 69
    Call GP_SET_CELL_VALUE(ss2, 69, cMin, AdoRs("A_STRESS_RT_MIN5"))
    Call GP_SET_CELL_VALUE(ss2, 69, cMax, AdoRs("A_STRESS_RT_MAX5"))
    Call GP_SET_CELL_VALUE(ss2, 69, cRst1, AdoRs("A_STRESS_RT_VAL5"))
    Call GP_SET_CELL_VALUE(ss2, 69, cDsc, AdoRs("A_STRESS_RT_DSC_RST5"))
    
'冲击试验增加温度显示 2012。4.9 刘翔
    ss2.Col = 1
    '冲击
    ss2.Row = 31
    Call GP_SET_CELL_VALUE(ss2, 31, 1, ss2.Text & "( " & AdoRs("IMPACT_TMP") & " ℃ )")
    '追加冲击
    ss2.Row = 34
    Call GP_SET_CELL_VALUE(ss2, 34, 1, ss2.Text & "( " & AdoRs("A_IMPACT_TMP") & " ℃ )")
    '时效冲击
    ss2.Row = 37
    Call GP_SET_CELL_VALUE(ss2, 37, 1, ss2.Text & "( " & AdoRs("TIM_IMPACT_TMP") & " ℃ )")
    '追加时效冲击
    ss2.Row = 39
    Call GP_SET_CELL_VALUE(ss2, 39, 1, ss2.Text & "( " & AdoRs("A_TIM_IMPACT_TMP") & " ℃ )")

    
    
    
        Select Case AdoRs("SMP_CUT_LOC")
    
        Case "B"
            opt_SMP_CUT_LOC(1).Value = True
        Case "T"
            opt_SMP_CUT_LOC(2).Value = True
        Case "M"
            opt_SMP_CUT_LOC(3).Value = True
        Case "A"
            opt_SMP_CUT_LOC(4).Value = True
        Case "Y"
            opt_SMP_CUT_LOC(5).Value = True
        Case Else
            opt_SMP_CUT_LOC(1).Value = True
    End Select
    
    Call subSpreadCheck2
        
    Exit Sub
    
Refer_Err:
    
    Screen.MousePointer = vbDefault

End Sub


Private Sub subSpreadCheck2()
    
On Error GoTo Refer_Err
    
 Dim i As Long
 Dim j As Long
    
    With ss2
        
        j = 0
        
        For i = 1 To .MaxRows
        
            .Row = i
     
            If ((Gf_Get_Cell_Value(ss2, i, 2) = 0 Or Gf_Get_Cell_Value(ss2, i, 2) = "") And _
                (Gf_Get_Cell_Value(ss2, i, 3) = 0 Or Gf_Get_Cell_Value(ss2, i, 3) = "") And _
                (Gf_Get_Cell_Value(ss2, i, 4) = 0 Or Gf_Get_Cell_Value(ss2, i, 4) = "") And _
                (Gf_Get_Cell_Value(ss2, i, 5) = 0 Or Gf_Get_Cell_Value(ss2, i, 5) = "") And _
                (Gf_Get_Cell_Value(ss2, i, 6) = 0 Or Gf_Get_Cell_Value(ss2, i, 6) = "") And _
                (Gf_Get_Cell_Value(ss2, i, 7) = 0 Or Gf_Get_Cell_Value(ss2, i, 7) = "") And _
                (Gf_Get_Cell_Value(ss2, i, 8) = 0 Or Gf_Get_Cell_Value(ss2, i, 8) = "") And _
                (Gf_Get_Cell_Value(ss2, i, 9) = 0 Or Gf_Get_Cell_Value(ss2, i, 9) = "") And _
                (Gf_Get_Cell_Value(ss2, i, 10) = 0 Or Gf_Get_Cell_Value(ss2, i, 10) = "") And _
                (Gf_Get_Cell_Value(ss2, i, 11) = 0 Or Gf_Get_Cell_Value(ss2, i, 11) = "")) Then

                .Row = i
                .RowHidden = True
            Else
                .RowHidden = False
                j = j + 1
                .Col = 0: .Text = j
            End If
        Next i
                
    End With
    
    Exit Sub
    
Refer_Err:
    
    Screen.MousePointer = vbDefault
       
End Sub

'--------------------配置化项目显示 王成  2012.12.14-----------------------------------------------------

Private Sub subSpreadView_Config(ByVal strArr As Variant)

    Dim i As Integer
    
    If UBound(strArr, 2) < 0 Then Exit Sub
    
    With ss2
        .MaxRows = .MaxRows + UBound(strArr, 2) + 1
        For i = 1 To UBound(strArr, 2) + 1
            .Row = OLD_MAXROWS + i
            .Col = 1: .Text = GF_NullChange(strArr(0, i - 1))
            .Col = 2: .Text = GF_NullChange(strArr(1, i - 1)) & ""
            .Col = 3: .Text = GF_NullChange(strArr(2, i - 1)) & ""
            .Col = 4: .Text = GF_NullChange(strArr(3, i - 1)) & ""
            .Col = 5: .Text = GF_NullChange(strArr(4, i - 1)) & ""
            .Col = 6: .Text = GF_NullChange(strArr(5, i - 1)) & ""
            .Col = 7: .Text = GF_NullChange(strArr(6, i - 1)) & ""
            .Col = 8: .Text = GF_NullChange(strArr(7, i - 1)) & ""
            .Col = 9: .Text = GF_NullChange(strArr(8, i - 1)) & ""
            .Col = 10: .Text = GF_NullChange(strArr(9, i - 1)) & ""
            .Col = 11: .Text = GF_NullChange(strArr(10, i - 1)) & ""
            .Col = 12: .Text = GF_NullChange(strArr(11, i - 1)) & ""
            .Col = 13: .Text = GF_NullChange(strArr(12, i - 1)) & ""
        Next i
            
    End With
    
    Call subSpreadCheck2

End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub SSSplitter1_SplitterStartDrag(ByVal SplitterBarType As Long, ByVal BorderPanes As SSSplitter.Panes, Cancel As Boolean)

End Sub
