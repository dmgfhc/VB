VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQC0360C 
   Caption         =   "综合判定不合格品处理_AQC0360C"
   ClientHeight    =   9090
   ClientLeft      =   -1710
   ClientTop       =   3330
   ClientWidth     =   14985
   FillColor       =   &H80000012&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14985
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_PROD3_KND 
      Height          =   310
      Left            =   9825
      TabIndex        =   27
      Text            =   "M"
      Top             =   1920
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt_grd_chg 
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
      Left            =   7770
      TabIndex        =   23
      Top             =   1920
      Width           =   645
   End
   Begin VB.TextBox txt_PLT 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   5715
      MaxLength       =   2
      TabIndex        =   22
      Tag             =   "PLT"
      Top             =   180
      Width           =   465
   End
   Begin VB.TextBox txt_enuse_chg_name 
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
      Left            =   8205
      TabIndex        =   21
      Top             =   1350
      Width           =   1635
   End
   Begin VB.TextBox txt_enuse_org_name 
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
      Left            =   5010
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1350
      Width           =   1590
   End
   Begin VB.TextBox txt_grd_org 
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
      Left            =   4515
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   1920
      Width           =   2070
   End
   Begin VB.TextBox txt_enuse_org 
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
      Left            =   4530
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1350
      Width           =   480
   End
   Begin VB.TextBox txt_enuse_chg 
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
      Left            =   7785
      TabIndex        =   16
      Top             =   1350
      Width           =   435
   End
   Begin VB.TextBox txt_stdspec_chg 
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
      Left            =   7770
      TabIndex        =   14
      Top             =   780
      Width           =   2070
   End
   Begin VB.TextBox txt_stdspec_org 
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
      Left            =   4530
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   780
      Width           =   2070
   End
   Begin VB.TextBox txt_smp_no 
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
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   14
      TabIndex        =   12
      Top             =   780
      Width           =   2070
   End
   Begin VB.TextBox txt_STDSPEC_NAME 
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
      Left            =   11100
      MaxLength       =   18
      TabIndex        =   9
      Top             =   165
      Width           =   750
   End
   Begin VB.TextBox txt_STDSPEC 
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
      Left            =   9060
      TabIndex        =   8
      Top             =   165
      Width           =   2070
   End
   Begin VB.TextBox txt_PROD_CD 
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
      Left            =   7380
      MaxLength       =   11
      TabIndex        =   1
      Tag             =   "产品代码"
      Top             =   180
      Width           =   435
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   6645
      Left            =   135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2340
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   11721
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   24
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0360C.frx":0000
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   195
      Top             =   180
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "综合判定日期"
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
      Left            =   6315
      Top             =   180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin InDate.UDate dtp_DSC_DATE_1 
      Height          =   300
      Left            =   1515
      TabIndex        =   2
      Top             =   180
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.UDate dtp_DSC_DATE_2 
      Height          =   300
      Left            =   3015
      TabIndex        =   3
      Top             =   180
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin Threed.SSCommand cmd_AQC0330C 
      Height          =   375
      Left            =   13635
      TabIndex        =   4
      Top             =   765
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
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
      Caption         =   "成分／材质详细"
      BevelWidth      =   1
   End
   Begin Threed.SSCommand cmd_Run_Test 
      Height          =   375
      Left            =   12015
      TabIndex        =   7
      Top             =   765
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   3
      Left            =   7995
      Top             =   165
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "标准编号"
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
   Begin CSTextLibCtl.sidbEdit sdb_ORD_THK 
      Height          =   300
      Left            =   13485
      TabIndex        =   10
      Top             =   165
      Width           =   600
      _Version        =   262145
      _ExtentX        =   1058
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
      Text            =   " 0.00"
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
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   2
      Left            =   12015
      Top             =   165
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "产品厚度/宽度"
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
   Begin CSTextLibCtl.sidbEdit sdb_ORD_WID 
      Height          =   300
      Left            =   14100
      TabIndex        =   11
      Top             =   165
      Width           =   1020
      _Version        =   262145
      _ExtentX        =   1799
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
      Text            =   " 0.00"
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
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   5
      Left            =   195
      Top             =   780
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "试样号"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   6
      Left            =   3420
      Top             =   780
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "当前标准"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   7
      Left            =   6675
      Top             =   780
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "改判标准"
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
   Begin Threed.SSCommand cmd_confirm 
      Height          =   375
      Left            =   10395
      TabIndex        =   15
      Top             =   765
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
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
      Caption         =   "确认改判"
      BevelWidth      =   1
   End
   Begin VB.TextBox Txt_Stand_No 
      Height          =   270
      Left            =   9870
      TabIndex        =   5
      Top             =   4245
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Txt_EndUse_CD 
      Height          =   270
      Left            =   10500
      TabIndex        =   6
      Top             =   3390
      Visible         =   0   'False
      Width           =   975
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   4
      Left            =   3435
      Top             =   1350
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "当前用途"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   8
      Left            =   6690
      Top             =   1350
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "改判用途"
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
      Index           =   9
      Left            =   3420
      Top             =   1920
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "当前等级"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   10
      Left            =   6675
      Top             =   1920
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "改判等级"
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
      ForeColor       =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_PROD_THK 
      Height          =   300
      Left            =   1290
      TabIndex        =   19
      Top             =   1350
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
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
      ReadOnly        =   -1  'True
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
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   11
      Left            =   180
      Top             =   1350
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "产品厚度"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel7 
      Height          =   300
      Index           =   0
      Left            =   4635
      Top             =   180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "工厂"
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
   Begin InDate.ULabel ULabel7 
      Height          =   300
      Index           =   1
      Left            =   10440
      Top             =   1200
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   529
      Caption         =   "强制放行规则说明:"
      Alignment       =   0
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
   Begin InDate.ULabel ULabel7 
      Height          =   300
      Index           =   2
      Left            =   10440
      Top             =   1560
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   529
      Caption         =   "1.产品强制放行时,修改等级录入:""A"""
      Alignment       =   0
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
      ForeColor       =   255
   End
   Begin InDate.ULabel ULabel7 
      Height          =   300
      Index           =   3
      Left            =   10440
      Top             =   1920
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   529
      Caption         =   "2.产品强制放行时,仅可以逐张/卷进行放行操作"
      Alignment       =   0
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
      ForeColor       =   255
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   315
      Left            =   8430
      TabIndex        =   24
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   196609
      BackStyle       =   1
      Begin VB.OptionButton opt_Confirm_MATR 
         BackColor       =   &H00E0E0E0&
         Caption         =   "材质"
         Height          =   270
         Left            =   630
         TabIndex        =   26
         Top             =   30
         Value           =   -1  'True
         Width           =   660
      End
      Begin VB.OptionButton opt_Confirm_CHEM 
         BackColor       =   &H00E0E0E0&
         Caption         =   "成份"
         Height          =   270
         Left            =   30
         TabIndex        =   25
         Top             =   30
         Width           =   690
      End
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   15
      X2              =   15240
      Y1              =   645
      Y2              =   645
   End
End
Attribute VB_Name = "AQC0360C"
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
'-- Program Name      综合判定不合格产品处理
'-- Program ID        AQC0360C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.08.21
'-- Description       综合判定不合格产品处理
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
Public sPLT_Authority As String     'Active User Plant Authority Setting

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

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_PROD_CD, "p", "n", " ", " ", " ", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dtp_DSC_DATE_1, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dtp_DSC_DATE_2, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_STDSPEC, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ORD_THK, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(sdb_ORD_WID, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_PROD3_KND, "p", " ", " ", " ", " ", "a", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQC0360C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQC0360C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQC0360C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cmd_AQC0320C_Click()
    If Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 1) <> "" Then
        AQC0320C.txt_PROD_NO.Text = Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 1)
        Call AQC0320C.Form_Ref

    End If
End Sub

Private Sub cmd_AQC0330C_Click()

    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If
    
    If Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 1) <> "" Then
       AQC0330C.txt_PROD_NO.Text = Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 1)
       AQC0330C.Form_Ref
    End If
End Sub


Private Sub cmd_Confirm_Click()
    Dim iRow As Long
    
    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If
   
'   If txt_grd_chg.Text = "3" And txt_PROD3_KND.Text = "" Then
'      MsgBox ("请选择是改判为成份协议还是材质协议！")
'   End If
    
If txt_stdspec_chg = "" Or txt_enuse_chg = "" Then
   Call Gp_MsgBoxDisplay("请输入修改标准，修改用途", "I")
   Exit Sub
Else
    If (Trim(txt_stdspec_chg) = Trim(txt_stdspec_org)) _
                    And (Trim(txt_enuse_chg) = Trim(txt_enuse_org)) _
                         And (Trim(txt_grd_chg) = Trim(txt_grd_org)) Then
    
       Call Gp_MsgBoxDisplay("请检查输入修改标准、修改用途、修改等级", "I")
       Exit Sub
    Else
        If Trim(txt_grd_chg) = "A" Then
            Call Gp_MsgBoxDisplay("强制放行仅可逐张进行", "I")
            Exit Sub
        End If
            With ss1
                 For iRow = 1 To .MaxRows
                     .Row = iRow
                     .Col = 2
                     If .Text = txt_SMP_NO Then
                        .Col = 6
                        .Text = txt_stdspec_chg
                        .Col = 10
                        .Text = txt_grd_chg
                        .Col = 18
                        .Text = txt_enuse_chg
                        .Col = 21
                        .Text = sUserID
                        .Col = 22
                        .Text = sUserName
''20100208 SUN BIN START
'
                        If txt_grd_chg.Text = "3" And txt_PROD3_KND.Text = "C" Then
                            .Col = 7
                            .Text = "3"
                            .Col = 9
                            .Text = "Z"
                        End If

                        If txt_grd_chg.Text = "3" And txt_PROD3_KND.Text = "M" Then
                            .Col = 7
                            .Text = "1"
                            .Col = 9
                            .Text = "3"
                        End If
''20100208 SUN BIN END
                        If Not (Change_Grade_Check(iRow, 10)) Then
                            Call Gp_MsgBoxDisplay("产品等级修改错误", "I")
                           .Col = 10
                           .Text = ""
                           .Col = 6
                           .Text = txt_stdspec_org
                           .Col = 18
                           .Text = txt_enuse_org
                           Exit Sub
                        End If
                        .Col = 0
                        .Text = "Update"

                     End If
                 Next iRow
            End With
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
        Case "txt_PROD_CD"             '产品
            sCode = "B0005"
            
        Case "txt_STDSPEC"              '标准
            sCode = "STDSPEC"
            Set oCodeName = txt_STDSPEC_NAME
        
        Case "txt_stdspec_chg"              '标准
            sCode = "STDSPEC2"
            
        Case "txt_enuse_chg"            '订单用途
             AQX0010C.Show 1
             Exit Sub
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
    
Err_Track:

End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuToolSet
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
    
    sPLT_Authority = Gf_PLT_Authority(Me.Name)
    If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
       txt_plt.Text = sPLT_Authority
    Else
       txt_plt.Text = ""
    End If

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuToolSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    cmd_Confirm.Enabled = False
    cmd_Run_Test.Enabled = False
    cmd_AQC0330C.Enabled = False
    dtp_DSC_DATE_1.Text = Date
    dtp_DSC_DATE_2.Text = Date
    txt_PROD_CD.Text = "PP"
    Screen.MousePointer = vbDefault

End Sub

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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuToolSet
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        dtp_DSC_DATE_1.Text = ""
        dtp_DSC_DATE_2.Text = ""
        txt_PROD_CD.Text = "PP"
        txt_STDSPEC.Text = ""
        txt_STDSPEC_NAME.Text = ""
        sdb_ORD_THK.Value = 0
        sdb_ORD_WID.Value = 0
        txt_SMP_NO = ""
        txt_stdspec_org = ""
        txt_stdspec_chg = ""
        txt_enuse_org = ""
        txt_enuse_org_name = ""
        txt_enuse_chg = ""
        txt_enuse_chg_name = ""
        txt_grd_org = ""
        txt_grd_chg = ""
        sdb_PROD_THK.Value = 0
        cmd_Run_Test.Enabled = False
        cmd_Confirm.Enabled = False
        cmd_AQC0330C.Enabled = False
        
        If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
           txt_plt.Text = sPLT_Authority
        Else
           txt_plt.Text = ""
        End If
        
        pControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuToolSet
        cmd_Confirm.Enabled = True
        cmd_Run_Test.Enabled = True
        cmd_AQC0330C.Enabled = True
    Else
        cmd_Confirm.Enabled = False
        cmd_Run_Test.Enabled = False
        cmd_AQC0330C.Enabled = False
    End If
    
    txt_SMP_NO = ""
    txt_stdspec_org = ""
    txt_stdspec_chg = ""
    txt_enuse_org = ""
    txt_enuse_org_name = ""
    txt_enuse_chg = ""
    txt_enuse_chg_name = ""
    txt_grd_org = ""
    txt_grd_chg = ""
    sdb_PROD_THK.Value = 0
    txt_PROD3_KND.Text = "M"
    opt_Confirm_MATR.Value = True
    Txt_Stand_No.Text = Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 6)
    txt_ENDUSE_CD.Text = Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 18)


End Sub

Public Sub Form_Pro()
Dim i As Long

    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If
    

    For i = 1 To ss1.MaxRows
         ss1.Col = 0
         ss1.Row = i
         If ss1.Text = "Update" Then
'            ss1.Col = 10
'            If ss1.Text = txt_grd_org Then
'               Call Gp_MsgBoxDisplay("修改标准与原标准相同", "I")
'               Exit Sub
'            End If
            ss1.Col = 6
            If ss1.Text = "" Then
               Call Gp_MsgBoxDisplay("修改标准等级为空", "I")
               Exit Sub
            End If
            ss1.Col = 10
            If ss1.Text = "" Then
               Call Gp_MsgBoxDisplay("修改等级为空", "I")
               Exit Sub
            End If
            
         End If
         
    Next i
    If Change_Grade_Check(ss1.ActiveRow, 10) Then
        If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuToolSet
            Call OS_ACE3020P_CALL
        End If
    End If
End Sub

Public Sub Spread_Cpy()

'    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)
    
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

Private Sub opt_Confirm_CHEM_Click()
    
    If opt_Confirm_CHEM.Value = True Then
        txt_PROD3_KND.Text = "C"
    End If

End Sub

Private Sub opt_Confirm_MATR_Click()

    If opt_Confirm_MATR.Value = True Then
       txt_PROD3_KND.Text = "M"
    End If
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
    Dim iOldRow As Long
    
        iOldRow = Row
        
    If Col = 10 Then
        If Not (Change_Grade_Check(Row, Col)) Then
            Call Gp_MsgBoxDisplay("产品等级修改错误", "I")
            ss1.Col = 10
            ss1.Text = ""
            Call ss1.SetActiveCell(10, iOldRow)
        ElseIf (Change_Grade_Check(Row, Col)) Then
'20100208 SUN BIN START
           With ss1
            .Row = .ActiveRow
            .Col = 10
              If .Text = "3" Then
                 .Col = 7
                 .Lock = False
                 .Col = 9
                 .Lock = False
              End If
            End With
        End If
'20100208 SUN BIN END
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
     If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub

     With ss1
          .Row = .ActiveRow
          .Col = 2
          txt_SMP_NO = .Text
          .Col = 6
          txt_stdspec_org = .Text
'          .Col = 8
'          If .Text = "1" Then
'             txt_grd_chg = "2"
'          Else
'             txt_grd_chg = .Text
'          End If
          
          .Col = 11
          txt_grd_org = .Text
          .Col = 14
          sdb_PROD_THK.Value = .Text
          .Col = 18
          txt_enuse_org = .Text
          .Col = 21
          .Text = sUserID
          .Col = 22
          .Text = sUserName
     End With

End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)
    End If
    
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If
    
    If KeyCode = vbKeyF4 Then
        With ss1
            Select Case .ActiveCol
            Case 6
         
                    Set DD.sPname = Me.ss1
    
                    DD.sWitch = "SP"
                    DD.rControl.Add Item:=6
    
                    .Row = .ActiveRow
    
                    DD.nameType = "2"
    
                    Call Gf_StdSPEC_DD(M_CN1, KeyCode)
            Case 10
                Set DD.sPname = Me.ss1
                    .Row = .ActiveRow
                    DD.sWitch = "SP"
                    DD.sKey = "Q0034"
                    DD.rControl.Add Item:=10
                    DD.nameType = "2"
                    Call Gf_Common_DD(M_CN1, KeyCode)
            Case 18
                Set DD.sPname = Me.ss1
                .Row = .ActiveRow
                DD.sWitch = "SP"
                DD.sKey = Mid(txt_PROD_CD, 1, 1)
                DD.rControl.Add Item:=18
                Call Gf_Usage_DD(M_CN1, KeyCode)
                
                    
            End Select
            
        End With
    End If
End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
End Sub

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    Txt_Stand_No.Text = Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 6)
    txt_ENDUSE_CD.Text = Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 18)
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

Public Function OS_ACE3020P_CALL() As Boolean

On Error GoTo SpreadPro_Error
  
    Dim sQuery As String
    
    Dim AdoRs As adodb.Recordset
    
    Set AdoRs = New adodb.Recordset
    
    Dim adoCmd3 As adodb.Command

    OS_ACE3020P_CALL = True
    
    Screen.MousePointer = vbHourglass
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd3 = New adodb.Command
    
    adoCmd3.CommandType = adCmdText
    Set adoCmd3.ActiveConnection = M_CN1
    sQuery = "{call ACB3020P(?)}"
    adoCmd3.CommandText = sQuery
    
    adoCmd3.Parameters.Append adoCmd3.CreateParameter("Messg", adVariant, adParamOutput)
    
    adoCmd3.Execute , , adExecuteNoRecords
    
    Screen.MousePointer = vbDefault
    
    Exit Function
    
SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd3 = Nothing
    
    OS_ACE3020P_CALL = False
    Err.Raise Err.Number, Err.Description

End Function

Private Function Change_Grade_Check(ByVal iRow As Long, ByVal iCol As Long) As Boolean

    Dim iFice_Grade  As Integer
    Dim sStand_No    As String
    Dim sEndUse_CD   As String
    Dim iChang_Grade As Integer
    Dim sChang_Grade As String
    Dim iOld_Grade   As Integer
    Dim iChem_Grade  As Integer
    Dim iMate_Grade  As Integer
    Dim iMax_Grade   As Integer
    
        If iCol <> 10 Then
         Exit Function
        Else
            With ss1
                .Row = iRow
                .Col = 6
                    sStand_No = .Text
                .Col = 7
                    iChem_Grade = Val(.Text)
                .Col = 8
                    iFice_Grade = Val(.Text)
                .Col = 9
                    iMate_Grade = Val(.Text)
                .Col = 18
                    sEndUse_CD = .Text
                .Col = 10
                    If .Text <> "A" Then
                        iChang_Grade = Val(.Text)
                        If iChang_Grade = 0 Then
                            Change_Grade_Check = False
                            Exit Function
                        End If
                        sChang_Grade = "S"
                    Else
                        iChang_Grade = 1
                        sChang_Grade = .Text
                    End If
                    
                .Col = 11
                    iOld_Grade = Val(.Text)
                
                iMax_Grade = MAX_Grade_Check(iChem_Grade, iFice_Grade, iMate_Grade)
        

                If iChang_Grade < iFice_Grade Then
                    Change_Grade_Check = False
                    Exit Function
                End If
                                
                If iChang_Grade >= iMax_Grade Then
                    Change_Grade_Check = True
                ElseIf sStand_No = Trim(Txt_Stand_No.Text) And sEndUse_CD = txt_ENDUSE_CD.Text Then
                    If sChang_Grade = "A" Then
                        Change_Grade_Check = True
                    Else
                        Change_Grade_Check = False
                    End If
                Else
                    Change_Grade_Check = True
                End If
            End With
        End If
End Function

Function MAX_Grade_Check(ByVal iChem_Grade As Integer, ByVal iFice_Grade As Integer, ByVal iMate_Grade As Integer) As Integer
    Dim iMax_Grade As Integer
    
'        If iChem_Grade > iFice_Grade Then
'            iMax_Grade = iChem_Grade
'        Else
'            iMax_Grade = iFice_Grade
'        End If
'
'        If iMax_Grade >= iMate_Grade Then
'            iMax_Grade = iMax_Grade
'        Else
'            iMax_Grade = iMate_Grade
'        End If
        iMax_Grade = iFice_Grade
        MAX_Grade_Check = iMax_Grade

End Function


Private Sub cmd_Run_Test_Click()
    Dim IORD As String
    Dim IORD_NO, IORD_ITEM As String
    
    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If
    
    If Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 1) <> "" Then
       Unload AQC0080C
       IORD = Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 4)
       IORD_NO = Mid(IORD, 1, 11)
       IORD_ITEM = Mid(IORD, 13, 2)
       AQC0080C.txt_ORD_NO.Text = IORD_NO
       AQC0080C.txt_ORD_ITEM.Text = IORD_ITEM
       AQC0080C.Form_Ref
    End If
 
End Sub
Private Sub Usage_SELECT()
    
    Dim sQuery          As String
    Dim sMesg           As String
    Dim AdoRs           As adodb.Recordset
    Dim ArrayRecords    As Variant
    
    On Error GoTo Error_Rtn

    Set AdoRs = New adodb.Recordset
    
    If txt_stdspec_chg = "" Then
        Exit Sub
    End If

    sQuery = " SELECT A.enduse_cd ,B.ENDUSE_NAME FROM qp_std_usage A, qp_ord_usage B "
    sQuery = sQuery + " WHERE A.ENDUSE_CD=B.ENDUSE_CD and a.prod_knd=b.prod_knd "
    sQuery = sQuery + "   AND A.PROD_KND = SUBSTR( '" & Trim(txt_PROD_CD.Text) & "', 1, 1)"
    sQuery = sQuery + "   AND  A.STDSPEC  = '" & Trim(txt_stdspec_chg.Text) & "' AND  A.THK_MIN < = " & sdb_PROD_THK.Value
    sQuery = sQuery + "   AND  A.THK_MAX  > = " & sdb_PROD_THK.Value & " AND NVL(A.ENDUSE_FL,'N') IN ('N','T','C')"
'    sQuery = sQuery + "   AND ROWNUM=1"
              
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not (AdoRs.BOF And AdoRs.EOF) Then
       ArrayRecords = AdoRs.GetRows
       txt_enuse_chg = ArrayRecords(0, 0)
       txt_enuse_chg_name = ArrayRecords(1, 0)
    Else
       MsgBox ("当前改判标准没有订单用途,改判后该产品使用原有订单用途!")
       txt_enuse_chg = txt_enuse_org
    End If
    
    Set AdoRs = Nothing
    Set ArrayRecords = Nothing
    
Error_Rtn:
    Set AdoRs = Nothing
    Set ArrayRecords = Nothing

End Sub

Private Sub txt_enuse_chg_Change()
   If Trim(txt_enuse_chg.Text) <> "" Then
      txt_enuse_chg_name = Gf_EnduseNameFind(M_CN1, Mid(Trim(txt_PROD_CD), 1, 1), txt_enuse_chg.Text)
   End If

End Sub

Private Sub txt_enuse_org_Change()
   If Trim(txt_enuse_org.Text) <> "" Then
      txt_enuse_org_name = Gf_EnduseNameFind(M_CN1, Mid(Trim(txt_PROD_CD), 1, 1), txt_enuse_org.Text)
   End If

End Sub

Private Sub txt_grd_chg_Change()
   If txt_grd_chg.Text = "3" And txt_PROD3_KND.Text = "" Then
      MsgBox ("请选择是改判为成份协议还是材质协议！")
   End If
End Sub

Private Sub txt_stdspec_chg_Change()
  If txt_stdspec_chg.Text = "" Then
     txt_enuse_chg.Text = ""
     txt_enuse_chg_name = ""
  End If
      Call Usage_SELECT
End Sub

Private Sub txt_stdspec_chg_LostFocus()
      Call Usage_SELECT
End Sub
Public Function Gf_EnduseNameFind(Conn As adodb.Connection, Code1 As String, Code2 As String) As Variant

On Error GoTo CodeFind_Error

    Dim sQuery As String
    Dim AdoRs As adodb.Recordset
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_EnduseNameFind = "FAIL": Exit Function
    End If
    
    Set AdoRs = New adodb.Recordset

    sQuery = "SELECT Enduse_name  FROM qp_ord_usage WHERE prod_knd = '" & Code1 & "' and enduse_cd='" & Code2 & "'"
      
   ' Debug.Print sQuery
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        If Not AdoRs.EOF Then
            Gf_EnduseNameFind = IIf(VarType(AdoRs.Fields(0)) = vbNull, "", AdoRs.Fields(0))
        End If
        
    Else
        Gf_EnduseNameFind = ""
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

CodeFind_Error:

    Set AdoRs = Nothing
    Gf_EnduseNameFind = "FAIL"

End Function
Private Sub MenuToolSet()
     
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
    
End Sub

Private Sub TXT_PLT_Change()

    If txt_plt.Text = "C3" Then
       txt_PROD_CD.Text = "PP"
    End If

End Sub

Private Sub TXT_PLT_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub


