VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACB5031C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "库内不同作业区转区作业实绩查询_ACB5031C"
   ClientHeight    =   10410
   ClientLeft      =   150
   ClientTop       =   1170
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10410
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_SIZE_KND 
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
      Height          =   315
      ItemData        =   "ACB5031C.frx":0000
      Left            =   13215
      List            =   "ACB5031C.frx":0010
      TabIndex        =   38
      Top             =   1200
      Width           =   720
   End
   Begin VB.TextBox text_cur_inv_1 
      Alignment       =   2  'Center
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
      Left            =   3865
      MaxLength       =   3
      TabIndex        =   37
      Tag             =   "仓库"
      Top             =   480
      Width           =   720
   End
   Begin VB.TextBox Text_PROC_CD 
      Alignment       =   2  'Center
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
      Left            =   3865
      MaxLength       =   3
      TabIndex        =   36
      Tag             =   "CD_MANA_NO"
      Top             =   1200
      Width           =   720
   End
   Begin VB.TextBox TXT_ORD_NO 
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
      Left            =   9060
      MaxLength       =   11
      TabIndex        =   35
      Tag             =   "CD_MANA_NO"
      Top             =   1200
      Width           =   1380
   End
   Begin VB.ComboBox CBO_ORD_ITEM 
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
      Left            =   10440
      TabIndex        =   34
      Top             =   1200
      Width           =   750
   End
   Begin VB.TextBox TXT_CUST_CD 
      Alignment       =   2  'Center
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
      Left            =   6030
      MaxLength       =   6
      TabIndex        =   33
      Top             =   1200
      Width           =   1590
   End
   Begin VB.TextBox TXT_MAT_NO 
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
      Left            =   3870
      MaxLength       =   14
      TabIndex        =   31
      Tag             =   "物料号"
      Top             =   1620
      Width           =   1965
   End
   Begin VB.TextBox TXT_AD_CD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15540
      MaxLength       =   1
      TabIndex        =   30
      Text            =   "A"
      Top             =   660
      Visible         =   0   'False
      Width           =   345
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   915
      Left            =   60
      TabIndex        =   24
      Top             =   0
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1614
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.OptionButton OPT_Y 
         BackColor       =   &H00E0E0E0&
         Caption         =   "营销转区"
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
         Left            =   180
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   270
         Width           =   1110
      End
      Begin VB.OptionButton OPT_N 
         BackColor       =   &H00E0E0E0&
         Caption         =   "分厂转区"
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
         Left            =   1350
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   270
         Value           =   -1  'True
         Width           =   1110
      End
   End
   Begin VB.ComboBox CBO_TRIM_FL 
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
      Height          =   315
      ItemData        =   "ACB5031C.frx":0024
      Left            =   3865
      List            =   "ACB5031C.frx":0034
      TabIndex        =   23
      Top             =   855
      Width           =   720
   End
   Begin VB.TextBox TXT_HTM_CD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15540
      MaxLength       =   1
      TabIndex        =   22
      Text            =   "Y"
      Top             =   120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox CBO_PROD_CD 
      BackColor       =   &H00C0FFFF&
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
      ItemData        =   "ACB5031C.frx":0044
      Left            =   3865
      List            =   "ACB5031C.frx":0051
      TabIndex        =   19
      Text            =   "PP"
      Top             =   75
      Width           =   720
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   7530
      Left            =   45
      TabIndex        =   17
      Top             =   2010
      Width           =   15240
      _Version        =   393216
      _ExtentX        =   26882
      _ExtentY        =   13282
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
      MaxCols         =   31
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACB5031C.frx":0061
   End
   Begin VB.TextBox Text_size_knd_name_IN 
      Enabled         =   0   'False
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
      Left            =   16845
      TabIndex        =   14
      Tag             =   "钢种"
      Top             =   1335
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txt_Trim_NAME 
      Enabled         =   0   'False
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
      Left            =   16845
      TabIndex        =   13
      Tag             =   "钢种"
      Top             =   1725
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txt_Trim_fl 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   16500
      MaxLength       =   1
      TabIndex        =   12
      Top             =   1725
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt_SizeKnd_s 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   16500
      MaxLength       =   2
      TabIndex        =   11
      Top             =   1335
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox text_cur_inv 
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
      Left            =   6420
      TabIndex        =   8
      Top             =   75
      Width           =   1200
   End
   Begin VB.TextBox text_cur_inv_code 
      Alignment       =   2  'Center
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
      Left            =   6030
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "起始区"
      Top             =   75
      Width           =   375
   End
   Begin VB.TextBox txt_dst_inv_code 
      Alignment       =   2  'Center
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
      Left            =   6030
      MaxLength       =   2
      TabIndex        =   6
      Tag             =   "目标区"
      Top             =   465
      Width           =   375
   End
   Begin VB.TextBox txt_dst_inv 
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
      Left            =   6420
      TabIndex        =   5
      Tag             =   "工 厂"
      Top             =   465
      Width           =   1200
   End
   Begin VB.TextBox txt_car_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6030
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "车辆号"
      Top             =   855
      Width           =   1590
   End
   Begin VB.TextBox txt_stdspec_s 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9060
      TabIndex        =   1
      Top             =   855
      Width           =   2895
   End
   Begin VB.TextBox txt_stlgrd_s 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Top             =   9180
      Visible         =   0   'False
      Width           =   1800
   End
   Begin CSTextLibCtl.sidbEdit txt_wid_min_s 
      Height          =   315
      Left            =   13215
      TabIndex        =   2
      Top             =   465
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
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
      NumDecDigits    =   0
      NumIntDigits    =   9
      ShowZero        =   0   'False
      MaxValue        =   99999999
      MinValue        =   -99999999
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk_min_s 
      Height          =   315
      Left            =   13215
      TabIndex        =   3
      Top             =   75
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
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
      NumDecDigits    =   0
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   4830
      Tag             =   "移 送 工 厂"
      Top             =   465
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "目标区"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   4830
      Top             =   75
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "起始区"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   4830
      Top             =   855
      Width           =   1170
      _ExtentX        =   2064
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
      ForeColor       =   16711680
   End
   Begin CSTextLibCtl.sidbEdit txt_wid_max_s 
      Height          =   315
      Left            =   14220
      TabIndex        =   9
      Top             =   465
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
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
      NumDecDigits    =   0
      NumIntDigits    =   9
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk_max_s 
      Height          =   315
      Left            =   14220
      TabIndex        =   10
      Top             =   75
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
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
      NumDecDigits    =   0
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   15525
      Top             =   1335
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "定尺"
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   7860
      Top             =   855
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "标准号"
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
   Begin InDate.ULabel ULabel23 
      Height          =   315
      Left            =   15525
      Top             =   1725
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "切边"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   12225
      Top             =   75
      Width           =   960
      _ExtentX        =   1693
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   12225
      Top             =   465
      Width           =   960
      _ExtentX        =   1693
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
   Begin InDate.ULabel ULabel12 
      Height          =   330
      Left            =   12225
      Top             =   855
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
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
      Left            =   7860
      Top             =   75
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "转库日期"
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
   Begin InDate.UDate MOV_DATE_FR 
      Height          =   315
      Left            =   9060
      TabIndex        =   15
      Tag             =   "移库日期"
      Top             =   75
      Width           =   1440
      _ExtentX        =   2540
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
      BackColor       =   12648447
   End
   Begin InDate.UDate MOV_DATE_TO 
      Height          =   315
      Left            =   10515
      TabIndex        =   16
      Tag             =   "移库日期"
      Top             =   75
      Width           =   1455
      _ExtentX        =   2566
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
      BackColor       =   12648447
   End
   Begin CSTextLibCtl.sidbEdit txt_len_max_s 
      Height          =   315
      Left            =   14220
      TabIndex        =   18
      Top             =   855
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
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
      NumDecDigits    =   0
      NumIntDigits    =   9
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   2670
      Top             =   75
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "产品"
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
      Left            =   7860
      Top             =   465
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "接收日期"
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
   Begin InDate.UDate RCV_DATE_FR 
      Height          =   315
      Left            =   9060
      TabIndex        =   20
      Tag             =   "移库日期"
      Top             =   465
      Width           =   1440
      _ExtentX        =   2540
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
   Begin InDate.UDate RCV_DATE_TO 
      Height          =   315
      Left            =   10515
      TabIndex        =   21
      Tag             =   "移库日期"
      Top             =   465
      Width           =   1455
      _ExtentX        =   2566
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   12225
      Top             =   1200
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "定尺"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   2670
      Top             =   855
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "切边"
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   855
      Left            =   60
      TabIndex        =   27
      Top             =   1080
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   1508
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.OptionButton OPT_D 
         BackColor       =   &H00E0E0E0&
         Caption         =   "明细"
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
         Left            =   1350
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   1110
      End
      Begin VB.OptionButton OPT_A 
         BackColor       =   &H00E0E0E0&
         Caption         =   "汇总"
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
         Left            =   180
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1110
      End
   End
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   2670
      Top             =   1620
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "物料号"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   5940
      Top             =   1620
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      Caption         =   " 按物料号可查询转库履历信息"
      Alignment       =   0
      BackColor       =   14804173
      BackgroundStyle =   1
      BorderEffect    =   0
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
      ForeColor       =   0
   End
   Begin Threed.SSCheck SSCHK 
      Height          =   405
      Left            =   10320
      TabIndex        =   32
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "在途中"
   End
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   4830
      Top             =   1200
      Width           =   1170
      _ExtentX        =   2064
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   7860
      Top             =   1200
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "订单号"
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
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   2670
      Top             =   1200
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "进程状态"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   2670
      Top             =   480
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Caption         =   "仓库"
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
   Begin CSTextLibCtl.sidbEdit txt_len_min_s 
      Height          =   315
      Left            =   13215
      TabIndex        =   39
      Top             =   855
      Width           =   975
      _Version        =   262145
      _ExtentX        =   1720
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
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
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
      NumDecDigits    =   0
      NumIntDigits    =   9
      ShowZero        =   0   'False
      MaxValue        =   99999999
      MinValue        =   -99999999
      Undo            =   0
      Data            =   0
   End
End
Attribute VB_Name = "ACB5031C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name   Common
'-- Program Name      Insert Moving result
'-- Program ID        ACB4031C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HWANG.J.D
'-- Coder             HWANG.J.D
'-- Date              2005.8.24
'-- Description
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
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_MV_NO = 1
Const SS1_MV_LST_NO = 2
Const SS1_MAT_NO = 3
Const SS1_PROD_CD = 4
Const SS1_APLY_STDSPEC = 5
Const SS1_THK = 6
Const SS1_WID = 7
Const SS1_LEN = 8
Const SS1_CNT = 9
Const SS1_WGT = 10
Const SS1_PROD_GRD = 11
Const SS1_SIZE_KND = 12
Const SS1_TRIM_FL = 13
Const SS1_CUST_CD = 14
Const SS1_CUST_NAME = 15
Const SS1_FR_INV = 16
Const SS1_FR_INV_LOC = 17
Const SS1_TO_INV = 18
Const SS1_TO_INV_LOC = 19
Const SS1_MOVE_DATE = 20
Const SS1_MOVE_TIME = 21
Const SS1_MOVE_CAR_NO = 23
Const SS1_MOVE_EMP = 24
Const SS1_RCV_DATE = 25
Const SS1_RCV_TIME = 26
Const SS1_RCV_EMP = 27
Const SS1_ORD = 28

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
       
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_dst_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(MOV_DATE_FR, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(MOV_DATE_TO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(RCV_DATE_FR, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(RCV_DATE_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_car_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_stdspec_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_SIZE_KND, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_thk_min_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_thk_max_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_wid_min_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_wid_max_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_len_min_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_len_max_s, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_TRIM_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_PROD_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_HTM_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_AD_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(SSCHK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_CUST_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Text_PROC_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(text_cur_inv_1, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB5031C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, SS1_MV_NO, True)
    Me.KeyPreview = True

End Sub

Private Sub cmd_Clear_Click()
    Call Gp_Ms_Cls(Mc1("rControl"))
    'cbo_prod_grd.Text = ""
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet
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
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gf_Sp_Cls(sc1)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    text_cur_inv_1.Text = "00"
    CBO_PROD_CD.Text = "PP"
    TXT_HTM_CD.Text = "N"
    TXT_AD_CD.Text = "A"
    
    RCV_DATE_FR.RawData = ""
    RCV_DATE_TO.RawData = ""
    
    Call OPT_A_Click
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
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
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
        If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            
            Call MenuTool_ReSet
            
            ss1.MaxRows = 0
            CBO_PROD_CD.Text = "PP"
            text_cur_inv_code.Text = "00"
            MOV_DATE_FR.RawData = ""
            MOV_DATE_TO.RawData = ""
            RCV_DATE_FR.RawData = ""
            RCV_DATE_TO.RawData = ""
            
        End If
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

On Error Resume Next

      
'    If SSCHK.Value = 0 And Trim(text_cur_inv_code.Text) = "" And Trim(txt_dst_inv_code.Text) = "" Then
'        Call Gp_MsgBoxDisplay("必须输入起始库或目标库其中之一")
'        Exit Sub
'    End If
    
    If Trim(text_cur_inv_code.Text) = Trim(txt_dst_inv_code.Text) Then
       txt_dst_inv_code.Text = ""
    End If
    
    If Len(txt_mat_no) = 10 Or Len(txt_mat_no) = 12 Or Len(txt_mat_no) = 14 Then
        OPT_D.Value = True
    End If
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
    
        Call Gp_Sp_EvenRowBackcolor(ss1)
        If TXT_AD_CD = "D" Then
           Call Sp_AutoSum_D
        Else
           Call Sp_AutoSum_A
        End If
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
    
    End If

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


Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub Option_ORD_FL_Y_Click()

End Sub

Private Sub OPT_A_Click()
    ss1.MaxRows = 0
    Call Gp_Sp_ColHidden(ss1, SS1_MAT_NO, True)
    Call Gp_Sp_ColHidden(ss1, SS1_APLY_STDSPEC, True)
    Call Gp_Sp_ColHidden(ss1, SS1_THK, True)
    Call Gp_Sp_ColHidden(ss1, SS1_WID, True)
    Call Gp_Sp_ColHidden(ss1, SS1_LEN, True)
    Call Gp_Sp_ColHidden(ss1, SS1_PROD_GRD, True)
    Call Gp_Sp_ColHidden(ss1, SS1_SIZE_KND, True)
    Call Gp_Sp_ColHidden(ss1, SS1_TRIM_FL, True)
    Call Gp_Sp_ColHidden(ss1, SS1_CUST_CD, True)
    Call Gp_Sp_ColHidden(ss1, SS1_CUST_NAME, True)
    Call Gp_Sp_ColHidden(ss1, SS1_FR_INV_LOC, True)
    Call Gp_Sp_ColHidden(ss1, SS1_TO_INV_LOC, True)
    Call Gp_Sp_ColHidden(ss1, SS1_ORD, True)
    TXT_AD_CD.Text = "A"
End Sub

Private Sub OPT_D_Click()
    ss1.MaxRows = 0
    Call Gp_Sp_ColHidden(ss1, SS1_MAT_NO, False)
    Call Gp_Sp_ColHidden(ss1, SS1_APLY_STDSPEC, False)
    Call Gp_Sp_ColHidden(ss1, SS1_THK, False)
    Call Gp_Sp_ColHidden(ss1, SS1_WID, False)
    Call Gp_Sp_ColHidden(ss1, SS1_LEN, False)
    Call Gp_Sp_ColHidden(ss1, SS1_PROD_GRD, False)
    Call Gp_Sp_ColHidden(ss1, SS1_SIZE_KND, False)
    Call Gp_Sp_ColHidden(ss1, SS1_TRIM_FL, False)
    Call Gp_Sp_ColHidden(ss1, SS1_CUST_CD, False)
    Call Gp_Sp_ColHidden(ss1, SS1_CUST_NAME, False)
    Call Gp_Sp_ColHidden(ss1, SS1_FR_INV_LOC, False)
    Call Gp_Sp_ColHidden(ss1, SS1_TO_INV_LOC, False)
    Call Gp_Sp_ColHidden(ss1, SS1_ORD, False)
    TXT_AD_CD.Text = "D"
End Sub

Private Sub OPT_N_Click()
    TXT_HTM_CD.Text = "N"
End Sub

Private Sub OPT_Y_Click()
    TXT_HTM_CD.Text = "Y"
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub


Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ss1.MaxRows > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub



'Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "B0005"
'        DD.rControl.Add Item:=text_prod_cd
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'    End If
'
'End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_Change()
  Dim S_CODE As String
   
   If text_cur_inv_1.Text = "00" Then
   
      S_CODE = "C0023"
   ElseIf text_cur_inv_1.Text = "WG" Then
   
      S_CODE = "C0026"
   End If
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, S_CODE, text_cur_inv_code.Text, 2)
    Else
      text_cur_inv.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Dim S_CODE As String
   
   If text_cur_inv_1.Text = "00" Then
   
      S_CODE = "C0023"
   ElseIf text_cur_inv_1.Text = "WG" Then
   
      S_CODE = "C0026"
   End If
   
     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = S_CODE

        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
End Sub

Private Sub txust_cdt_dst_inv_code_DblClick()

    Call txt_dst_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cust_cd_DblClick()

    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_dst_inv_code_DblClick()
    Call txt_dst_inv_code_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then
    
        If cbo_ord_item.Text <> "" Then Exit Sub
        
        txt_ord_no.Text = StrConv(txt_ord_no.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)

    Else
        cbo_ord_item.Clear
    End If

End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_CUST_CD

        DD.nameType = "1"
        Call Gf_Customer_DD(M_CN1, KeyCode)

    End If

End Sub



Private Sub txt_dst_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

   Dim S_CODE As String
   If text_cur_inv_1.Text = "00" Then
   
      S_CODE = "C0023"
   ElseIf text_cur_inv_1.Text = "WG" Then
   
      S_CODE = "C0026"
   End If
     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = S_CODE

        DD.rControl.Add Item:=txt_dst_inv_code
        DD.rControl.Add Item:=txt_dst_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
End Sub

Private Sub txt_car_no_DblClick()

    Call txt_car_no_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_car_no_KeyUp(KeyCode As Integer, Shift As Integer)

    If ULabel10.Caption <> "车辆号" Then Exit Sub
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_car_no
  '      DD.rControl.Add Item:=txt_fac_name

        DD.nameType = "2"

        Call Gf_CAR_NO_DD(M_CN1, KeyCode)

    End If

End Sub

Public Function Gf_CAR_NO_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "A"        'Apply Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT CAR_NO, CAR_KND,CAR_WGT_MAX,CAR_WGT_AVE,CAR_CMP_CD,Gf_Comnnamefind('H0002',CAR_CMP_CD) AS CAR_CMP_NAME FROM  HP_CAR_IMF "
    '    DD.sQuery = DD.sQuery + " WHERE "
        DD.sWhere = " WHERE CAR_NO like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        
    Else
    

    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing
    
End Function

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
'        .Buttons(8).Enabled = False                 'Row Delete
'        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub

Private Sub txt_SizeKnd_s_Change()
    If Len(Trim(txt_SizeKnd_s.Text)) = txt_SizeKnd_s.MaxLength Then
        Text_size_knd_name_IN.Text = Gf_ComnNameFind(M_CN1, "B0043", txt_SizeKnd_s.Text, 2)
        Exit Sub
    Else
        Text_size_knd_name_IN.Text = ""
    End If
End Sub

Private Sub txt_SizeKnd_s_DblClick()

    Call txt_SizeKnd_s_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_SizeKnd_s_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0043"

        DD.rControl.Add Item:=txt_SizeKnd_s

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_stdspec_s_DblClick()

    Call txt_stdspec_s_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_Change()
    If Len(Trim(txt_Trim_fl.Text)) = txt_Trim_fl.MaxLength Then
        txt_Trim_NAME.Text = Gf_ComnNameFind(M_CN1, "B0021", txt_Trim_fl.Text, 2)
        txt_Trim_fl.Text = Trim(txt_Trim_fl.Text)
        Exit Sub
    Else
        txt_Trim_NAME.Text = ""
        txt_Trim_fl.Text = ""
    End If

End Sub

Private Sub txt_trim_fl_DblClick()

    Call txt_trim_fl_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_trim_fl_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0021"

        DD.rControl.Add Item:=txt_Trim_fl

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If

End Sub

Private Sub txt_stdspec_s_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_s

        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_stlgrd_s_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd_s
        
        DD.nameType = "1"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
    End If
End Sub

Private Sub Sp_AutoSum_D()

    Dim dValue(2) As Double
    Dim dValueTotal(2) As Double
    Dim lngLstNum As Long
    Dim iCount As Integer
    Dim x As Integer
    Dim strProdTag As String
    Dim iRow As Integer
    Dim iCurRow As Integer
    Dim iRow2 As Integer
    Dim bLoop As Boolean
     
    bLoop = True
    dValueTotal(0) = 0
    dValueTotal(1) = 0
    lngLstNum = 0
    
    With ss1
        iRow2 = 1
        If .MaxRows < 2 Then Exit Sub
        While bLoop

            If iRow2 > .MaxRows Then
                bLoop = False
                GoTo LastSum
            End If
            'bLoop = False
            iRow = iRow2
            dValue(0) = 0
            dValue(1) = 0
            
            For iCurRow = iRow To .MaxRows
                .Col = 2: .Row = iCurRow: strProdTag = .Text
                .Row = iCurRow + 1
                If .Text = strProdTag Then
                    iRow2 = iRow2 + 1
                Else
                    lngLstNum = lngLstNum + 1
                    .MaxRows = .MaxRows + 1

                    .Row = iCurRow + 1
                    .Action = SS_ACTION_INSERT_ROW
                    .Col = 0: .Text = "∑"
                    Call .AddCellSpan(1, .Row, 3, 1)
                    .Col = 1: .Text = strProdTag & " 合计"
                    For iCount = 1 To .MaxCols
                        .Col = iCount
                        If .CellType = SS_CELL_TYPE_COMBOBOX Then .Value = 0
                    Next iCount
                    Call Gp_Sp_RowColor(ss1, .Row, vbBlue, &HC0FFFF)

                    'dValue = Sp_SumAbove(ss1, 7, iRow, irow2)
                    
                    dValue(0) = Sp_SumAbove(ss1, 10, iRow, iRow2)
                    dValue(1) = Sp_SumAbove(ss1, 11, iRow, iRow2)
                    dValueTotal(0) = dValueTotal(0) + dValue(0)
                    dValueTotal(1) = dValueTotal(1) + dValue(1)
                    .Row = iCurRow + 1
                    .Col = 10: .Value = IIf(dValue(0) > 0, dValue(0), "")
                    .Col = 11: .Value = IIf(dValue(1) > 0, dValue(1), "")
                    iRow2 = iRow2 + 2
                    Exit For
                End If
            Next iCurRow
        Wend
LastSum:
            .MaxRows = .MaxRows + 1

            .Row = .MaxRows
            .Action = SS_ACTION_INSERT_ROW
            .Col = 0: .Text = "∑"
            Call .AddCellSpan(1, .Row, 3, 1)
            .Col = 1: .Text = "合计码单数量/总量"
            .Col = 4: .Text = CStr(lngLstNum)
            For iCount = 1 To .MaxCols
                .Col = iCount
                If .CellType = SS_CELL_TYPE_COMBOBOX Then .Value = 0
            Next iCount
            Call Gp_Sp_RowColor(ss1, .Row, vbBlue, vbYellow)
 
            .Col = 10: .Value = IIf(dValue(0) > 0, dValueTotal(0), "")
            .Col = 11: .Value = IIf(dValue(1) > 0, dValueTotal(1), "")
            iRow2 = iRow2 + 2
    End With

End Sub

Private Sub Sp_AutoSum_A()

    Dim iCount As Integer
    
    Dim sTotnum As Double
    Dim sTotwgt As Double

    With ss1
    
        If .MaxRows <= 1 Then
           Exit Sub
        End If
        
        For iCount = 1 To .MaxRows
            .Row = iCount:            .Col = SS1_CNT:            sTotnum = sTotnum + .Text
            .Row = iCount:            .Col = SS1_WGT:            sTotwgt = sTotwgt + .Text
        Next iCount
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows:            .Col = SS1_MV_LST_NO:      .Text = "汇总"
        .Row = .MaxRows:            .Col = SS1_CNT:            .Text = sTotnum
        .Row = .MaxRows:            .Col = SS1_WGT:            .Text = sTotwgt
        
    End With
    
    Call Gp_Sp_RowColor(ss1, ss1.MaxRows, vbBlue, &HC0FFFF)
    
End Sub

Private Function Sp_SumAbove(ByVal SS As Variant, ByVal iCol As Long, ByVal iRow1, ByVal iRow2) As Double
    Dim dSum As Double
    Dim iCount As Integer

    dSum = 0

    With SS
        If iRow1 > iRow2 Then
            Sp_SumAbove = 0
            Exit Function
        End If
        If iRow2 > .MaxRows Then iRow2 = .MaxRows
        If iRow2 < 2 Then
            Sp_SumAbove = 0
            Exit Function
        End If
        .Col = iCol
        For iCount = iRow1 To iRow2
            .Row = iCount
            If .CellType = SS_CELL_TYPE_NUMBER And .Text <> "" Then
                dSum = dSum + .Value
            End If
        Next iCount

    End With
    Sp_SumAbove = dSum
End Function



