VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AEC2010C 
   BackColor       =   &H00C0C0C0&
   Caption         =   "确定炼钢作业生产管制指示_AEC2010C"
   ClientHeight    =   9375
   ClientLeft      =   720
   ClientTop       =   225
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   15195
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_prc_line2 
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
      ItemData        =   "AEC2010C.frx":0000
      Left            =   7875
      List            =   "AEC2010C.frx":0002
      TabIndex        =   47
      Tag             =   "炉座号"
      Top             =   75
      Width           =   705
   End
   Begin VB.ComboBox cbo_prc_line1 
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
      ItemData        =   "AEC2010C.frx":0004
      Left            =   7170
      List            =   "AEC2010C.frx":0006
      TabIndex        =   28
      Tag             =   "炉座号"
      Top             =   75
      Width           =   705
   End
   Begin Threed.SSFrame Frame2 
      Height          =   435
      Left            =   13500
      TabIndex        =   16
      Top             =   60
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   767
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin Threed.SSOption Opt_InqBof 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "转炉标准"
         Value           =   -1
      End
      Begin Threed.SSOption Opt_InqCcm 
         Height          =   285
         Left            =   1350
         TabIndex        =   18
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "连铸标准"
      End
   End
   Begin VB.TextBox txt_proc_fl 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   14340
      MaxLength       =   2
      TabIndex        =   15
      Tag             =   "工厂"
      Top             =   210
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt_prc_line 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   13830
      MaxLength       =   2
      TabIndex        =   14
      Tag             =   "工厂"
      Top             =   240
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox cbo_prc_line 
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
      ItemData        =   "AEC2010C.frx":0008
      Left            =   6465
      List            =   "AEC2010C.frx":000A
      TabIndex        =   13
      Tag             =   "炉座号"
      Top             =   75
      Width           =   705
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   1905
      MaxLength       =   2
      TabIndex        =   12
      Tag             =   "工厂"
      Top             =   80
      Width           =   465
   End
   Begin VB.TextBox txt_plt_name 
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
      Left            =   2370
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   11
      Tag             =   "工厂"
      Top             =   80
      Width           =   1920
   End
   Begin Threed.SSCommand cmd_Time 
      Height          =   330
      Left            =   12285
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   9300
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "时间调整"
   End
   Begin VB.TextBox txt_chprt 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   1350
      TabIndex        =   8
      Top             =   9300
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.TextBox txt_sendprt 
      BackColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   6
      Top             =   9300
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txt_to_seq 
      Height          =   345
      Left            =   6540
      TabIndex        =   5
      Top             =   9300
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox txt_from_seq 
      Height          =   345
      Left            =   4500
      TabIndex        =   4
      Top             =   9300
      Visible         =   0   'False
      Width           =   1800
   End
   Begin Threed.SSCommand cmd_send 
      Height          =   330
      Left            =   14220
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   9300
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&下达指示"
   End
   Begin VB.TextBox txt_heat_mana_no 
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
      Left            =   2850
      MaxLength       =   8
      TabIndex        =   0
      Top             =   9300
      Visible         =   0   'False
      Width           =   1230
   End
   Begin Threed.SSCommand cmd_change 
      Height          =   330
      Left            =   13260
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9300
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "指示调整"
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8325
      Left            =   90
      TabIndex        =   2
      Top             =   990
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   14684
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AEC2010C.frx":000C
      Begin FPSpread.vaSpread ss1 
         Height          =   5010
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   4890
         _Version        =   393216
         _ExtentX        =   8625
         _ExtentY        =   8837
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   21
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEC2010C.frx":00DE
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3255
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   5070
         Width           =   4890
         _Version        =   393216
         _ExtentX        =   8625
         _ExtentY        =   5741
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEC2010C.frx":0C54
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   5040
         Left            =   4950
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   0
         Width           =   5310
         _Version        =   393216
         _ExtentX        =   9366
         _ExtentY        =   8890
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   21
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEC2010C.frx":16CF
      End
      Begin FPSpread.vaSpread ss5 
         Height          =   5040
         Left            =   10320
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   0
         Width           =   4830
         _Version        =   393216
         _ExtentX        =   8520
         _ExtentY        =   8890
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   21
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEC2010C.frx":2245
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   3225
         Left            =   4950
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   5100
         Width           =   5310
         _Version        =   393216
         _ExtentX        =   9366
         _ExtentY        =   5689
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEC2010C.frx":2DBB
      End
      Begin FPSpread.vaSpread ss6 
         Height          =   3225
         Left            =   10320
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   5100
         Width           =   4830
         _Version        =   393216
         _ExtentX        =   8520
         _ExtentY        =   5689
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEC2010C.frx":3834
      End
   End
   Begin CSTextLibCtl.sidbEdit sdb_from 
      Height          =   285
      Left            =   8130
      TabIndex        =   9
      Top             =   9315
      Visible         =   0   'False
      Width           =   465
      _Version        =   262145
      _ExtentX        =   820
      _ExtentY        =   503
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
      StartText.x     =   2
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      NumDecDigits    =   0
      NumIntDigits    =   8
      MaxValue        =   9999
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_to 
      Height          =   285
      Left            =   8610
      TabIndex        =   10
      Top             =   9315
      Visible         =   0   'False
      Width           =   465
      _Version        =   262145
      _ExtentX        =   820
      _ExtentY        =   503
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
      StartText.x     =   2
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      NumDecDigits    =   0
      NumIntDigits    =   8
      MaxValue        =   9999
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   135
      Top             =   75
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      Caption         =   "工厂"
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
      Left            =   4665
      Top             =   75
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      Caption         =   "炉座号(左/中/右)"
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
   Begin Threed.SSFrame Frame1 
      Height          =   465
      Left            =   120
      TabIndex        =   19
      Top             =   450
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_chg_prc_line 
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
         ItemData        =   "AEC2010C.frx":42AD
         Left            =   5370
         List            =   "AEC2010C.frx":42AF
         TabIndex        =   27
         Tag             =   "炉座号"
         Top             =   90
         Width           =   615
      End
      Begin Threed.SSOption opt_cancel 
         Height          =   285
         Left            =   7920
         TabIndex        =   20
         Top             =   120
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "取消"
      End
      Begin Threed.SSOption opt_mltcd_change 
         Height          =   285
         Left            =   7260
         TabIndex        =   21
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "工序流程变更"
      End
      Begin Threed.SSOption opt_line_change 
         Height          =   285
         Left            =   6030
         TabIndex        =   22
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "炉座号变更"
      End
      Begin Threed.SSOption opt_change 
         Height          =   285
         Left            =   210
         TabIndex        =   23
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "顺序调整"
      End
      Begin Threed.SSOption opt_del 
         Height          =   285
         Left            =   1590
         TabIndex        =   24
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "删除"
      End
      Begin Threed.SSOption opt_time 
         Height          =   285
         Left            =   8130
         TabIndex        =   25
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "时间调整"
      End
      Begin Threed.SSOption opt_send 
         Height          =   285
         Left            =   9780
         TabIndex        =   26
         Top             =   120
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "下达"
      End
      Begin Threed.SSOption opt_rsltdel 
         Height          =   285
         Left            =   8790
         TabIndex        =   36
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "实绩删除"
      End
      Begin Threed.SSOption opt_ins 
         Height          =   285
         Left            =   2520
         TabIndex        =   40
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "生产管制指示"
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   465
      Left            =   4440
      TabIndex        =   29
      Top             =   450
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox cbo_heat_no2 
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
         Left            =   3705
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   33
         Tag             =   "终止炉号"
         Top             =   60
         Width           =   990
      End
      Begin VB.TextBox txt_tgt_heat_no 
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
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   32
         Tag             =   "终止炉号"
         Top             =   60
         Width           =   990
      End
      Begin VB.TextBox cbo_heat_mana_no 
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
         Left            =   1290
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         Tag             =   "起始炉号"
         Top             =   60
         Width           =   990
      End
      Begin Threed.SSOption opt_from 
         Height          =   315
         Left            =   180
         TabIndex        =   30
         Top             =   75
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "起始炉号"
      End
      Begin Threed.SSOption opt_to 
         Height          =   315
         Left            =   2580
         TabIndex        =   34
         Top             =   75
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "终止炉号"
      End
      Begin Threed.SSOption opt_target 
         Height          =   315
         Left            =   5010
         TabIndex        =   35
         Top             =   75
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "目标炉号"
      End
   End
   Begin Threed.SSPanel SSPrtn 
      Height          =   300
      Left            =   14460
      TabIndex        =   37
      Top             =   660
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   529
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "返送"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPpdt 
      Height          =   300
      Left            =   13770
      TabIndex        =   38
      Top             =   660
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   529
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "生产中"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPsend 
      Height          =   300
      Left            =   13080
      TabIndex        =   39
      Top             =   660
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   529
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "已下达"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin InDate.ULabel ULabel9 
      Height          =   225
      Left            =   9000
      Top             =   120
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   397
      Caption         =   "重点订单用红色字体显示"
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
      ForeColor       =   255
   End
   Begin Threed.SSCommand cmd_heat_fl 
      Height          =   330
      Left            =   11640
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "修改炉次标记"
      BevelWidth      =   3
   End
End
Attribute VB_Name = "AEC2010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       System
'-- Sub_System Name
'-- Program Name
'-- Program ID        AEC2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.6.23
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
Public Complete As Boolean
Public P_MODE As String             'CHANGE PRC_LINE = 'U', MOVE = 'M', DELETE = 'D', CANCEL = 'C', SEND = 'L',   TIME = 'T'
Public p_cur_prd As Integer
Public iProd As String
Public iRet  As String
Public p_up_down As String
Public Chg_Lf    As Boolean         'LF Change Check
Public Chg_VD    As Boolean         'VD Change Check
Public Chg_RH    As Boolean         'RH Change Check
Public Ref_FL    As Boolean
Public sAut As String
'Public CS As String

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

Dim pContro3 As New Collection      'Master Primary Key Collection
Dim nContro3 As New Collection      'Master Necessary Collection
Dim mContro3 As New Collection      'Master Maxlength check Collection
Dim iContro3 As New Collection      'Master Insert Collection
Dim rContro3 As New Collection      'Master Refer Collection
Dim cContro3 As New Collection      'Master Copy Collection
Dim aContro3 As New Collection      'Master -> Spread Collection
Dim lContro3 As New Collection      'Master Lock Collection

Dim pContro4 As New Collection      'Master Primary Key Collection
Dim nContro4 As New Collection      'Master Necessary Collection
Dim mContro4 As New Collection      'Master Maxlength check Collection
Dim iContro4 As New Collection      'Master Insert Collection
Dim rContro4 As New Collection      'Master Refer Collection
Dim cContro4 As New Collection      'Master Copy Collection
Dim aContro4 As New Collection      'Master -> Spread Collection
Dim lContro4 As New Collection      'Master Lock Collection

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

Dim pColumn5 As New Collection      'Spread Primary Key Collection
Dim nColumn5 As New Collection      'Spread necessary Column Collection
Dim mColumn5 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn5 As New Collection      'Spread Insert Column Collection
Dim aColumn5 As New Collection      'Master -> Spread Column Collection
Dim lColumn5 As New Collection      'Spread Lock Column Collection

Dim pColumn6 As New Collection      'Spread Primary Key Collection
Dim nColumn6 As New Collection      'Spread necessary Column Collection
Dim mColumn6 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn6 As New Collection      'Spread Insert Column Collection
Dim aColumn6 As New Collection      'Master -> Spread Column Collection
Dim lColumn6 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection
Dim Mc4 As New Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Sc4 As New Collection           'Spread Collection
Dim Sc5 As New Collection           'Spread Collection
Dim Sc6 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS_IMP_CONT = 16             '重点订单红色标记  2013-11-15 by CaoLei  15->16

Dim errMsg   As String
Dim iSelRow  As Integer
Dim txt_AFT_Prc_line As String
Dim txt_AFT_SS_Col As Integer
Dim txt_AFT_SS_Row As Integer
Dim iCount As Integer

Private Sub Form_Define()

    Dim iCol As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(cbo_prc_line, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(Opt_InqBof, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(txt_heat_mana_no, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
       
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    Call Gp_Ms_Collection(cbo_prc_line1, "p", "n", " ", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
       Call Gp_Ms_Collection(Opt_InqBof, "p", " ", " ", " ", " ", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    
    'MASTER Collection
    Mc3.Add Item:=pContro3, Key:="pControl"
    Mc3.Add Item:=nContro3, Key:="nControl"
    Mc3.Add Item:=mContro3, Key:="mControl"
    Mc3.Add Item:=iContro3, Key:="iControl"
    Mc3.Add Item:=rContro3, Key:="rControl"
    Mc3.Add Item:=cContro3, Key:="cControl"
    Mc3.Add Item:=aContro3, Key:="aControl"
    Mc3.Add Item:=lContro3, Key:="lControl"
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
    Call Gp_Ms_Collection(cbo_prc_line2, "p", "n", " ", " ", "r", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
       Call Gp_Ms_Collection(Opt_InqBof, "p", " ", " ", " ", " ", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
    
    'MASTER Collection
    Mc4.Add Item:=pContro4, Key:="pControl"
    Mc4.Add Item:=nContro4, Key:="nControl"
    Mc4.Add Item:=mContro4, Key:="mControl"
    Mc4.Add Item:=iContro4, Key:="iControl"
    Mc4.Add Item:=rContro4, Key:="rControl"
    Mc4.Add Item:=cContro4, Key:="cControl"
    Mc4.Add Item:=aContro4, Key:="aControl"
    Mc4.Add Item:=lContro4, Key:="lControl"
     
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To 2
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
           Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    For iCol = 4 To 21
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    
    If Opt_InqBof.Value = True Then
        Sc1.Add Item:="AEC2010C.P_REFER1", Key:="P-R"
    ElseIf Opt_InqCcm.Value = True Then
        Sc1.Add Item:="AEC2010C.P_REFER4", Key:="P-R"
    End If
    
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To 19
        Call Gp_Sp_Collection(SS2, iCol, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    
    'Spread_Collection
    sc2.Add Item:=SS2, Key:="Spread"
    sc2.Add Item:="AEC2010C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=SS2.MaxCols, Key:="Last"
   
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To 21
        Call Gp_Sp_Collection(SS3, iCol, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Next iCol
    
    'Spread_Collection
    Sc3.Add Item:=SS3, Key:="Spread"
    
    If Opt_InqBof.Value = True Then
        Sc3.Add Item:="AEC2010C.P_REFER1", Key:="P-R"
    ElseIf Opt_InqCcm.Value = True Then
        Sc3.Add Item:="AEC2010C.P_REFER4", Key:="P-R"
    End If
    
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=3, Key:="First"
    Sc3.Add Item:=SS3.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To 19
        Call Gp_Sp_Collection(ss4, iCol, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Next iCol
    
    'Spread_Collection
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:="AEC2010C.P_REFER2", Key:="P-R"
    Sc4.Add Item:=pColumn4, Key:="pColumn"
    Sc4.Add Item:=nColumn4, Key:="nColumn"
    Sc4.Add Item:=aColumn4, Key:="aColumn"
    Sc4.Add Item:=mColumn4, Key:="mColumn"
    Sc4.Add Item:=iColumn4, Key:="iColumn"
    Sc4.Add Item:=lColumn4, Key:="lColumn"
    Sc4.Add Item:=1, Key:="First"
    Sc4.Add Item:=ss4.MaxCols, Key:="Last"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To 21
        Call Gp_Sp_Collection(ss5, iCol, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Next iCol
    
    'Spread_Collection
    Sc5.Add Item:=ss5, Key:="Spread"
    
    If Opt_InqBof.Value = True Then
        Sc5.Add Item:="AEC2010C.P_REFER1", Key:="P-R"
    ElseIf Opt_InqCcm.Value = True Then
        Sc5.Add Item:="AEC2010C.P_REFER4", Key:="P-R"
    End If
    
    Sc5.Add Item:=pColumn5, Key:="pColumn"
    Sc5.Add Item:=nColumn5, Key:="nColumn"
    Sc5.Add Item:=aColumn5, Key:="aColumn"
    Sc5.Add Item:=mColumn5, Key:="mColumn"
    Sc5.Add Item:=iColumn5, Key:="iColumn"
    Sc5.Add Item:=lColumn5, Key:="lColumn"
    Sc5.Add Item:=3, Key:="First"
    Sc5.Add Item:=ss5.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To 19
        Call Gp_Sp_Collection(ss6, iCol, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Next iCol
    
    'Spread_Collection
    Sc6.Add Item:=ss6, Key:="Spread"
    Sc6.Add Item:="AEC2010C.P_REFER2", Key:="P-R"
    Sc6.Add Item:=pColumn6, Key:="pColumn"
    Sc6.Add Item:=nColumn6, Key:="nColumn"
    Sc6.Add Item:=aColumn6, Key:="aColumn"
    Sc6.Add Item:=mColumn6, Key:="mColumn"
    Sc6.Add Item:=iColumn6, Key:="iColumn"
    Sc6.Add Item:=lColumn6, Key:="lColumn"
    Sc6.Add Item:=1, Key:="First"
    Sc6.Add Item:=ss6.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    Proc_Sc.Add Item:=Sc4, Key:="Sc4"
    Proc_Sc.Add Item:=Sc5, Key:="Sc5"
    Proc_Sc.Add Item:=Sc6, Key:="Sc6"
    
    Me.KeyPreview = True
    Me.Opt_InqBof.BackColor = &HE0E0E0
    Me.Opt_InqCcm.BackColor = &HE0E0E0
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 5, True)
    Call Gp_Sp_ColHidden(ss1, 16, True)
    Call Gp_Sp_ColHidden(ss1, 17, True)
    Call Gp_Sp_ColHidden(ss1, 18, True)
    Call Gp_Sp_ColHidden(ss1, 19, True)
    Call Gp_Sp_ColHidden(ss1, 20, True)
    Call Gp_Sp_ColHidden(ss1, 21, True)
    
    Call Gp_Sp_ColHidden(SS3, 5, True)
    Call Gp_Sp_ColHidden(SS3, 16, True)
    Call Gp_Sp_ColHidden(SS3, 17, True)
    Call Gp_Sp_ColHidden(SS3, 18, True)
    Call Gp_Sp_ColHidden(SS3, 19, True)
    Call Gp_Sp_ColHidden(SS3, 20, True)
    Call Gp_Sp_ColHidden(SS3, 21, True)
    
    Call Gp_Sp_ColHidden(ss5, 5, True)
    Call Gp_Sp_ColHidden(ss5, 16, True)
    Call Gp_Sp_ColHidden(ss5, 17, True)
    Call Gp_Sp_ColHidden(ss5, 18, True)
    Call Gp_Sp_ColHidden(ss5, 19, True)
    Call Gp_Sp_ColHidden(ss5, 20, True)
    Call Gp_Sp_ColHidden(ss5, 21, True)
    
    Call Gp_Sp_ColHidden(SS2, 1, True)     'SEQ_NO
    Call Gp_Sp_ColHidden(ss4, 1, True)     'SEQ_NO
    Call Gp_Sp_ColHidden(ss6, 1, True)     'SEQ_NO

End Sub

Private Sub cbo_heat_mana_no_Change()

    Dim iRow As Integer
    Dim sColor As String
    
    With ss1
        .Col = 1
        For iRow = 1 To .MaxRows
            .Row = iRow
            If .BackColor = &HFF Then
               .Col = 4
                sColor = .BackColor
               .Col = 1: .Col2 = 2
               .BackColor = sColor
            End If
            
        Next
    End With
    
    With SS3
        .Col = 1
        For iRow = 1 To .MaxRows
            .Row = iRow
            If .BackColor = &HFF Then
               .Col = 4
                sColor = .BackColor
               .Col = 1: .Col2 = 2
               .BackColor = sColor
            End If
            
        Next
    End With
    
    With ss5
        .Col = 1
        For iRow = 1 To .MaxRows
            .Row = iRow
            If .BackColor = &HFF Then
               .Col = 4
                sColor = .BackColor
               .Col = 1: .Col2 = 2
               .BackColor = sColor
            End If
            
        Next
    End With
        
End Sub

Private Sub cbo_prc_line_Change()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub cbo_prc_line_Click()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub cbo_prc_line1_Change()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub cbo_prc_line1_Click()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub cbo_prc_line2_Change()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref

End Sub

Private Sub cbo_prc_line2_Click()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref

End Sub

Private Sub cmd_heat_fl_Click()

   Call HEAT_DESIGN_FL_CHAGE
   Call Form_Ref

End Sub

Private Sub HEAT_DESIGN_FL_CHAGE()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As adodb.Command
    Dim iCount As Integer
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    For iCount = 1 To Sc1.Item("Spread").MaxRows
    
        ss1.Row = iCount
        ss1.Col = 0
        
        If ss1.Text = "Update" Then
        
            
            sQuery = "{call AEC2010C.P_HEAT_FL_CHAGE ("
            
            'HEAT_MANA_NO
            ss1.Col = 1
            sQuery = sQuery + "'" + ss1.Text + "',"
            
            'PRC_LINE
            ss1.Col = 2
            sQuery = sQuery + "'" + ss1.Text + "',"
            
            'HEAT_DESIGN_FL
            ss1.Col = 3
            sQuery = sQuery + "'" + ss1.Text + "',?)}"
            
            
            'Ado Setting
            M_CN1.CursorLocation = adUseServer
            Set adoCmd = New adodb.Command
            
            adoCmd.CommandType = adCmdText
            Set adoCmd.ActiveConnection = M_CN1
            
            adoCmd.CommandText = sQuery
            
            adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
            
            adoCmd.Execute , , adExecuteNoRecords
            
            'Process Error Check
            If adoCmd("arg_e_msg") <> "" Then
                ret_Result_ErrMsg = "Error Mesg : " & adoCmd("arg_e_msg")
                Screen.MousePointer = vbDefault
                Call Gp_MsgBoxDisplay(ret_Result_ErrMsg)
                Set adoCmd = Nothing
                Exit Sub
            End If
            
            Set adoCmd = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
            
        End If
    
    Next iCount
    
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
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

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    Call Gp_Ms_NeceColor(Mc4("nControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc4.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc5.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc6.Item("Spread"), False)
    
'    Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc4.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc5.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc6.Item("Spread"))
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(Sc4)
    Call Gf_Sp_Cls(Sc5)
    Call Gf_Sp_Cls(Sc6)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "E-System.INI", Me.Name)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc4.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc5.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc5.Item("Spread"), "E-System.INI", Me.Name)
    
    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    Ref_FL = False
    
    cbo_prc_line.Clear
    cbo_prc_line.AddItem "1"
    cbo_prc_line.AddItem "2"
    cbo_prc_line.AddItem "3"
    cbo_prc_line.ListIndex = 0
    
    cbo_prc_line1.Clear
    cbo_prc_line1.AddItem "1"
    cbo_prc_line1.AddItem "2"
    cbo_prc_line1.AddItem "3"
    cbo_prc_line1.ListIndex = 1
    
    cbo_prc_line2.Clear
    cbo_prc_line2.AddItem "1"
    cbo_prc_line2.AddItem "2"
    cbo_prc_line2.AddItem "3"
    cbo_prc_line2.ListIndex = 2
    
    cbo_chg_prc_line.Clear
    cbo_chg_prc_line.AddItem "1"
    cbo_chg_prc_line.AddItem "2"
    cbo_chg_prc_line.AddItem "3"
    cbo_chg_prc_line.ListIndex = 0
    
    txt_PRC_LINE.Text = "2"
    txt_proc_fl.Text = ""
    
    Ref_FL = True
    
    If Mid(sAuthority, 1, 1) = "1" Then
        Call Form_Ref
    End If
    
    Screen.MousePointer = vbDefault
    
    '20140122
   If sUserID = "1BY1002" Or sUserID = "1BY1003" Or sUserID = "1BY1004" Or sUserID = "1BY1005" Then
          opt_del.Enabled = False
   End If
   '20140122
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim iCnt As Integer
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "E-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc4.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc5.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc6.Item("Spread"), "E-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
    
    Set pContro3 = Nothing
    Set nContro3 = Nothing
    Set iContro3 = Nothing
    Set rContro3 = Nothing
    Set cContro3 = Nothing
    Set aContro3 = Nothing
    Set lContro3 = Nothing
    Set mContro3 = Nothing
    
    Set pContro4 = Nothing
    Set nContro4 = Nothing
    Set iContro4 = Nothing
    Set rContro4 = Nothing
    Set cContro4 = Nothing
    Set aContro4 = Nothing
    Set lContro4 = Nothing
    Set mContro4 = Nothing
    
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
    
    Set iColumn5 = Nothing
    Set pColumn5 = Nothing
    Set lColumn5 = Nothing
    Set nColumn5 = Nothing
    Set mColumn5 = Nothing
    Set aColumn5 = Nothing
    
    Set iColumn6 = Nothing
    Set pColumn6 = Nothing
    Set lColumn6 = Nothing
    Set nColumn6 = Nothing
    Set mColumn6 = Nothing
    Set aColumn6 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Mc4 = Nothing
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Sc4 = Nothing
    Set Sc5 = Nothing
    Set Sc6 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
    
        If Gf_Sp_Cls(Sc1) Then
        
            Ref_FL = False
            
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gf_Sp_Cls(Sc3)
            Call Gf_Sp_Cls(Sc4)
            Call Gf_Sp_Cls(Sc5)
            Call Gf_Sp_Cls(Sc6)
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call MenuTool_ReSet
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            txt_heat_mana_no.Enabled = True
            rContro1(1).SetFocus
            txt_plt.Text = "B1"
            Call txt_plt_KeyUp(0, 0)
            
            cbo_heat_mana_no.Text = ""
            cbo_heat_mana_no.Enabled = False
            cbo_heat_no2.Text = ""
            cbo_heat_no2.Enabled = False
            txt_tgt_heat_no.Text = ""
            txt_tgt_heat_no.Enabled = False
            txt_proc_fl.Text = ""
            
'            SSPsend.Visible = False
'            SSPpdt.Visible = False
'            SSPrtn.Visible = False
            
            cbo_chg_prc_line.Text = ""
            opt_line_change.Value = False
            opt_mltcd_change.Value = False
            opt_cancel.Value = False
            opt_change.Value = False
            opt_del.Value = False
            opt_time.Value = False
            opt_send.Value = False
            
            opt_from.Enabled = False
            opt_to.Enabled = False
            opt_target.Enabled = False
            
            Ref_FL = True
            cbo_prc_line.ListIndex = 0
            cbo_prc_line1.ListIndex = 1
            cbo_prc_line2.ListIndex = 2
        End If
        
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sQuery As String
    Dim PGM_ID As String
    Dim Ref_FL As String
    
    If Mid(sAuthority, 1, 1) = "0" Then
        Exit Sub
    End If
    
    PGM_ID = "AEC2010C"
    cbo_chg_prc_line.Text = ""
    cbo_heat_mana_no.Text = ""
    cbo_heat_no2.Text = ""
    txt_tgt_heat_no.Text = ""
    
    opt_line_change.Value = False
    opt_mltcd_change.Value = False
    opt_cancel.Value = False
    opt_change.Value = False
    opt_del.Value = False
    opt_time.Value = False
    opt_send.Value = False
    opt_ins.Value = False
    
    opt_from.Enabled = False
    opt_to.Enabled = False
    opt_target.Enabled = False
    
    If Not opt_line_change Then
        opt_line_change.BackColor = &HE0E0E0
    End If
    
    opt_cancel.BackColor = &HE0E0E0
    opt_del.BackColor = &HE0E0E0
    opt_mltcd_change.BackColor = &HE0E0E0
    opt_change.BackColor = &HE0E0E0
    opt_time.BackColor = &HE0E0E0
    opt_ins.BackColor = &HE0E0E0
    
    Ref_FL = "0"
'    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc3, Mc3, Mc3("nControl"), Mc3("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc5, Mc4, Mc4("nControl"), Mc4("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Ref_FL = "1" Then
            
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Gp_Sp_EvenRowBackcolor(Sc1.Item("Spread"))
        Call Gf_Sp_Cls(sc2)
        
        'Call Gf_Sp_Refer(M_CN1, Sc3, Mc3, Mc3("nControl"), Mc3("mControl"))
        Call Gp_Sp_EvenRowBackcolor(Sc3.Item("Spread"))
        Call Gf_Sp_Cls(Sc4)
        
        'Call Gf_Sp_Refer(M_CN1, Sc5, Mc4, Mc4("nControl"), Mc4("mControl"))
        Call Gp_Sp_EvenRowBackcolor(Sc5.Item("Spread"))
        Call Gf_Sp_Cls(Sc6)
        
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(9).Enabled = False
        ss1.OperationMode = OperationModeNormal
        SS3.OperationMode = OperationModeNormal
        ss5.OperationMode = OperationModeNormal
        
'            SSPsend.Visible = True
'            SSPpdt.Visible = True
'            SSPrtn.Visible = True
        
'        sQuery = "select upd from zp_authority where emp_id = '" + sUserID + "' and pgmid = '" + PGM_ID + "'"
'        sAut = Gf_FloatFind(M_CN1, sQuery)
        
        If Mid(sAuthority, 3, 1) = "1" Then
            Frame1.Enabled = True
        Else
            Frame1.Enabled = False
        End If

'            Call Spread_Color_Setting(ss1)
'            Call Spread_Color_Setting(ss3)
       
        If opt_line_change Or opt_mltcd_change Then
        
            With ss1

                For iRow = iSelRow To .MaxRows
                    .Row = iRow
                    .Col = 19
                    If .Text <> "Y" Then
                    
                        If opt_line_change Then
                            .BlockMode = True
                            .Col = 2:    .Col2 = 2
                            .Row = iRow: .ROW2 = iRow
                            .BackColor = &HC0FFEE
                            .Lock = False
                            .BlockMode = False
                        End If
                        
                        If opt_mltcd_change Then
                            .BlockMode = True
                            .Col = 9:    .Col2 = 9
                            .Row = iRow: .ROW2 = iRow
                            .BackColor = &HC0FFEE
                            .BlockMode = False
                        End If
                    End If
                Next
                
            End With
            
        End If
    
    Else
        Call Gf_Sp_Cls(sc2)
        Call Gf_Sp_Cls(Sc3)
        Call Gf_Sp_Cls(Sc4)
        Call Gf_Sp_Cls(Sc5)
        Call Gf_Sp_Cls(Sc6)
    End If
            
End Sub

Public Sub Spread_Color_Setting(oSpr As vaSpread)

    Dim iRow As Integer
    Dim iCol As Integer
    
    With oSpr
    
        .Col = 1
        
        For iRow = 1 To .MaxRows
            .Row = iRow:  .Col = 19
            If .Text = "Y" Then
                For iCol = 1 To .MaxCols
                    .Col = iCol
                    .BackColor = SSPsend.BackColor
                Next iCol
            ElseIf .Text = "" Then
                Exit For
            End If
        Next iRow
          
        For iRow = 1 To .MaxRows
            .Col = 16:  .Row = iRow
           
            If Trim(.Text) <> "" Then
               p_cur_prd = iRow
               For iCol = 1 To .MaxCols
                   .Col = iCol
                   .BackColor = SSPpdt.BackColor
               Next
            End If
          
            .Col = 17
            .Row = iRow
          
            If .Text > "0" Then
               For iCol = 1 To .MaxCols
                   .Col = iCol
                   .BackColor = SSPrtn.BackColor
               Next
            End If
            
        Next
        
    End With
    
End Sub

Public Sub Spread_Forzens_Setting()
    
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Spread_Can()

    Dim sPrcLine As String
    Dim sMltProc As String
    Dim iRow     As Integer
    
    For iRow = lBlkrow1 To lBlkrow2
        With ss1
            .Row = iRow
            .Col = 0:     .Text = ""
            .Col = 20:    sPrcLine = .Text
            .Col = 21:    sMltProc = .Text
            
            .Col = 2:    .Text = sPrcLine
            .Col = 9:    .Text = sMltProc
        End With
    Next iRow
    
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Active_Spread, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Pro()

    Dim sTemp, Ret_Steel, Not_Ret As String
    Dim iCnt, I, j As Integer
    Dim sQuery As String
    Dim OutParam(1, 4) As Variant

    If ss1.MaxRows = 0 And SS3.MaxRows = 0 And ss5.MaxRows = 0 Then Exit Sub
    
'    If AKN2030C.opt_line_change.Enabled = False And _
'       AKN2030C.opt_mltcd_change.Enabled = False And _
'       AKN2030C.opt_cancel.Value = False And _
'       AKN2030C.opt_change.Value = False And _
'       AKN2030C.opt_del.Value = False And _
'       AKN2030C.opt_time.Value = False And _
'       AKN2030C.opt_send.Value = False Then
'       MsgBox "请选择您要做的操作！", vbCritical, "系统提示信息"
'       Exit Sub
'    End If
    
    errMsg = ""
    
'    If P_MODE = "C" Then
'        Call DataCompareLow
'    ElseIf P_MODE <> "X" Then
'        Call DataCompareHigh
'    End If
    
    If errMsg <> "" Then
       MsgBox errMsg, vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    'Line Status CHECK (added by KIM.SUNG.HO 2010.03.20)
'    If P_MODE <> "P" Then
'        If Not Line_Status_Chk Then Exit Sub
'    End If


    If P_MODE = "M" Or P_MODE = "D" Or P_MODE = "I" Then

        iCnt = Gf_FloatFind(M_CN1, "SELECT COUNT(*) FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ BETWEEN LEAST('" & cbo_heat_mana_no.Text & _
                                                                                                         "','" & cbo_heat_no2.Text & _
                                                                                                         "','" & txt_tgt_heat_no.Text & _
                                                                                                         "')  AND " & _
                                                                                                 "GREATEST('" & cbo_heat_mana_no.Text & _
                                                                                                         "','" & cbo_heat_no2.Text & _
                                                                                                         "','" & txt_tgt_heat_no.Text & _
                                                                                                         "') " & "AND PRC_STS = 'A' ")
                                                                                                
        If iCnt <> 0 Then
            MsgBox "某炉号已下达生产管制,咨询MES组!!!", vbCritical, "系统提示信息"
            Exit Sub
        End If

    End If
    
    '-----------------------------------------------------------------------
    
    Select Case P_MODE
    
            Case "U"
            
                With ss1
                     For I = 1 To .MaxRows
                         .Col = 2
                         .Row = I
                         If .Text <> "1" And .Text <> "2" And .Text <> "3" Then
                             MsgBox "请正确输入炉座号！", vbInformation, "系统提示信息"
                             Exit Sub
                         End If
                     Next I
                End With
                
'                If Gf_MessConfirm("确定要变更作业指示的炉座号吗？", "W", "系统提示信息确认") Then
'                   Call Plc_Line_Change
'                Else
'                   Exit Sub
'                End If

           Case "L"
           
                If Len(Trim(cbo_heat_mana_no.Text)) <> 8 Or Len(Trim(cbo_heat_no2.Text)) <> 8 Then
                    MsgBox "请正确输入起始炉号和终止炉号！", vbCritical, "系统提示信息"
                    Exit Sub
                End If
                
                If Gf_MessConfirm("确定要下达从炉号 <" + cbo_heat_mana_no.Text + "> 到炉号 <" + cbo_heat_no2.Text + "> 的作业指示吗？", "W", "系统提示信息确认") Then
                    Call Gp_Schedul_Send
                Else
                    Exit Sub
                End If
                
           Case "M"
           
                If Len(Trim(cbo_heat_mana_no.Text)) <> 8 Or Len(Trim(cbo_heat_no2.Text)) <> 8 Or Len(Trim(txt_tgt_heat_no.Text)) <> 8 Then
                    MsgBox "请正确输入起始炉号、终止炉号及目标炉号！", vbCritical, "系统提示信息"
                    Exit Sub
                End If

              '  If cbo_heat_mana_no.Text = cbo_heat_no2.Text Then
              '     MsgBox "起始炉号和目标炉号相同，无需调整！", vbInformation, "系统提示信息"
              '     cbo_heat_mana_no.Text = ""
              '     cbo_heat_no2.Text = ""
              '     Exit Sub
              '  End If
                If Gf_MessConfirm("确定要将炉号 <" + cbo_heat_mana_no.Text + "> 到 <" + cbo_heat_no2 + "> 调整到 <" + txt_tgt_heat_no + ">之后吗？", "W", "系统提示信息确认") Then
                    Call Gp_Schedul_Send
                Else
                    Exit Sub
                End If
                
           Case "C"
                
                If Len(Trim(cbo_heat_mana_no.Text)) <> 8 Then
                    MsgBox "请选择您要取消的起始炉号！", vbCritical, "系统提示信息"
                    Exit Sub
                End If
                
                If Gf_MessConfirm("确定要从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消作业指示吗？", "W", "系统提示信息确认") Then
                   Call Gp_Schedul_Send
                Else
                   Exit Sub
                End If
                   
           Case "P"
                
                If Len(Trim(cbo_heat_mana_no.Text)) <> 8 Then
                    MsgBox "请选择您要实绩删除的起始炉号！", vbCritical, "系统提示信息"
                    Exit Sub
                End If
                
                If Gf_MessConfirm("确定要从炉号 <" + cbo_heat_mana_no.Text + "> 开始实绩删除作业指示吗？", "W", "系统提示信息确认") Then
                   Call Gp_Schedul_Send
                Else
                   Exit Sub
                End If
                   
           Case "D"
           
                If Len(Trim(cbo_heat_mana_no.Text)) <> 8 Or Len(Trim(cbo_heat_no2.Text)) <> 8 Then
                    MsgBox "请正确输入起始炉号和终止炉号！", vbCritical, "系统提示信息"
                    Exit Sub
                End If
                
                With ss1
                     For I = 1 To .MaxRows
                         .Col = 1
                         .Row = I
                         If .Text >= cbo_heat_mana_no.Text And .Text <= cbo_heat_no2.Text Then
                             .Col = 3
                             If .BackColor = SSPrtn.BackColor Then
                                 Ret_Steel = "Y"
                             ElseIf .BackColor <> SSPrtn.BackColor Then
                                 Not_Ret = "Y"
                             End If
                         ElseIf .Text > cbo_heat_no2.Text Then
                             Exit For
                         End If
                     Next I
                End With
                
                If Ret_Steel = "Y" And Not_Ret = "Y" Then
                    MsgBox "不能同时删除返送和非返送的炉号！" + vbCrLf + vbCrLf + "请分次删除返送的炉号和非返送的炉号！", vbInformation, "系统提示信息"
                    Exit Sub
                Else
                    If Gf_MessConfirm("确定要删除从炉号 <" + cbo_heat_mana_no.Text + "> 到炉号 <" + cbo_heat_no2 + "> 的作业指示吗？", "W", "系统提示信息确认") Then
                         Call Gp_Schedul_Send
                    Else
                         Exit Sub
                    End If
                End If
                
           Case "I"
           
                If Len(Trim(cbo_heat_mana_no.Text)) <> 8 Or Len(Trim(cbo_heat_no2.Text)) <> 8 Then
                    MsgBox "请正确输入起始炉号和终止炉号！", vbCritical, "系统提示信息"
                    Exit Sub
                End If
                
                If Gf_MessConfirm("确定要指示从炉号 <" + cbo_heat_mana_no.Text + "> 到炉号 <" + cbo_heat_no2 + "> 的作业指示吗？", "W", "系统提示信息确认") Then
                     Call Gp_Schedul_Send
                Else
                     Exit Sub
                End If
                
           Case "T"
           
                If Gf_MessConfirm("您确定要将作业指示中的时间信息调整到更准确吗？", "W", "系统提示信息确认") Then
                    Call Gp_Schedul_Send
                Else
                    Exit Sub
                End If
    End Select
    
    If Trim(errMsg) = "" Then Call Form_Ref
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub Search_Last_Line()

    Dim iRow        As Integer
    
    iSelRow = 1
    
    With ss1
    
        For iRow = .MaxRows To 1 Step -1
            .Row = iRow
            .Col = 4
            If .BackColor = SSPpdt.BackColor Then
                iSelRow = iRow
                Exit Sub
            End If
        Next
        
    End With

End Sub

Private Sub Search_Last_Line3()

    Dim iRow        As Integer
    
    iSelRow = 1
    
    With SS3
    
        For iRow = .MaxRows To 1 Step -1
            .Row = iRow
            .Col = 4
            If .BackColor = SSPpdt.BackColor Then
                iSelRow = iRow
                Exit Sub
            End If
        Next
        
    End With

End Sub

Private Sub Search_Last_Line5()

    Dim iRow        As Integer
    
    iSelRow = 1
    
    With ss5
    
        For iRow = .MaxRows To 1 Step -1
            .Row = iRow
            .Col = 4
            If .BackColor = SSPpdt.BackColor Then
                iSelRow = iRow
                Exit Sub
            End If
        Next
        
    End With

End Sub


Public Function Sf_Sp_ProceExist() As Integer

    Dim iRow        As Integer
    Dim sColor      As String
    
    Sf_Sp_ProceExist = 0
    
    With ss1
    
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 0
            If Trim(.Text) = "Input" Or Trim(.Text) = "Update" Or Trim(.Text) = "Delete" Then
                Sf_Sp_ProceExist = 1
            End If
            
            .Col = 4
             sColor = .BackColor
             
             .Col = 2:   .Col2 = 2
             .BackColor = sColor
             
             .Col = 9:   .Col2 = 9
             .BackColor = sColor
        Next
        
    End With
    
    MDIMain.MenuTool.Buttons(9).Enabled = False
    
End Function

Private Sub opt_ins_Click(Value As Integer)

    If Sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_ins.Value = True

    P_MODE = "I"
    
    opt_from.Enabled = False
    opt_to.Enabled = True
    opt_target.Enabled = False
   
    cbo_heat_mana_no.Text = ""
    cbo_heat_no2.Text = ""
    txt_tgt_heat_no.Text = ""
    opt_to.Value = True
    
    opt_line_change.BackColor = &HE0E0E0
    opt_cancel.BackColor = &HE0E0E0
    opt_del.BackColor = &HE0E0E0
    opt_mltcd_change.BackColor = &HE0E0E0
    opt_change.BackColor = &HE0E0E0
    opt_time.BackColor = &HE0E0E0
    opt_rsltdel.BackColor = &HE0E0E0
    opt_ins.BackColor = &HC0FFFF
    
    ss1.BlockMode = True
    ss1.Col = -1
    ss1.Row = -1
    ss1.Lock = True
    ss1.BlockMode = False

End Sub

Private Sub opt_line_change_Click(Value As Integer)

    Dim iRow        As Integer
    
    If Sf_Sp_ProceExist() > 0 Then Call Form_Ref:  opt_line_change.Value = True
    MDIMain.MenuTool.Buttons(9).Enabled = True
    
    P_MODE = "U"
    
    cbo_chg_prc_line.ListIndex = cbo_prc_line1.ListIndex
    
    opt_from.Enabled = False
    opt_to.Enabled = False
    opt_target.Enabled = False
    opt_from.Value = False
    cbo_heat_mana_no.Text = ""
    cbo_heat_no2.Text = ""
    txt_tgt_heat_no.Text = ""
    
    Call Search_Last_Line
    
    opt_line_change.BackColor = &HC0FFFF
    opt_cancel.BackColor = &HE0E0E0
    opt_del.BackColor = &HE0E0E0
    opt_mltcd_change.BackColor = &HE0E0E0
    opt_change.BackColor = &HE0E0E0
    opt_time.BackColor = &HE0E0E0
    opt_rsltdel.BackColor = &HE0E0E0
    
    With ss1
    
        For iRow = iSelRow To .MaxRows
            .Row = iRow
            .Col = 19
            If .Text <> "Y" Then
                .BlockMode = True
                .Col = 2:    .Col2 = 2
                .Row = iRow: .ROW2 = iRow
                .BackColor = &HC0FFEE
                .Lock = False
                .BlockMode = False
            End If
        Next
        
    End With
    
End Sub

Private Sub opt_mltcd_change_Click(Value As Integer)

    Dim iRow        As Integer
    
    If Sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_mltcd_change.Value = True
    MDIMain.MenuTool.Buttons(9).Enabled = True
    
    P_MODE = "X"
    
    opt_from.Enabled = False
    opt_to.Enabled = False
    opt_target.Enabled = False
    opt_from.Value = False
    cbo_heat_mana_no.Text = ""
    cbo_heat_no2.Text = ""
    txt_tgt_heat_no.Text = ""
    
    opt_line_change.BackColor = &HE0E0E0
    opt_cancel.BackColor = &HE0E0E0
    opt_del.BackColor = &HE0E0E0
    opt_mltcd_change.BackColor = &HC0FFFF
    opt_change.BackColor = &HE0E0E0
    opt_time.BackColor = &HE0E0E0
    opt_rsltdel.BackColor = &HE0E0E0
    
    With ss1
    
        For iRow = iSelRow To .MaxRows
            .Row = iRow
            .Col = 19
            If .Text <> "Y" Then
                .BlockMode = True
                .Col = 9:    .Col2 = 9
                .Row = iRow: .ROW2 = iRow
                .BackColor = &HC0FFEE
'                .Lock = False
                .BlockMode = False
            End If
        Next
        
    End With
    
End Sub

Private Sub opt_cancel_Click(Value As Integer)

    Dim iRow       As Integer
    Dim sColor, TT As String
    
    If Sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_cancel.Value = True
    
    P_MODE = "C"
    
    opt_from.Enabled = True
    opt_to.Enabled = False
    opt_target.Enabled = False
    cbo_heat_no2.Text = ""
    txt_tgt_heat_no.Text = ""
    opt_from.Value = True

    ss1.BlockMode = True
    ss1.Col = -1
    ss1.Row = -1
    ss1.Lock = True
    ss1.BlockMode = False
    
    opt_line_change.BackColor = &HE0E0E0
    opt_cancel.BackColor = &HC0FFFF
    opt_del.BackColor = &HE0E0E0
    opt_mltcd_change.BackColor = &HE0E0E0
    opt_change.BackColor = &HE0E0E0
    opt_time.BackColor = &HE0E0E0
    opt_rsltdel.BackColor = &HE0E0E0

    With ss1
    
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 4
             sColor = .BackColor
             .Col = 1:   .Col2 = 2
             .BackColor = sColor
            If (.Text = cbo_heat_mana_no.Text) And (sColor = SSPsend.BackColor Or sColor = SSPrtn.BackColor) Then
               .BackColor = &HFF
               TT = "A"
            End If
        Next
        
        If TT <> "A" Then
           cbo_heat_mana_no.Text = ""
        End If
        
    End With
    
End Sub

Private Sub opt_change_Click(Value As Integer)

    If Sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_change.Value = True
    
    P_MODE = "M"
    
    opt_from.Enabled = True
    opt_to.Enabled = True
    opt_target.Enabled = True
    opt_from.Value = True
    cbo_heat_mana_no.Text = ""
    cbo_heat_no2.Text = ""
'    Opt_InqBof.Value = True
'    Opt_InqCcm.Value = False
    opt_line_change.BackColor = &HE0E0E0
    opt_cancel.BackColor = &HE0E0E0
    opt_del.BackColor = &HE0E0E0
    opt_mltcd_change.BackColor = &HE0E0E0
    opt_change.BackColor = &HC0FFFF
    opt_time.BackColor = &HE0E0E0
    opt_rsltdel.BackColor = &HE0E0E0
    opt_ins.BackColor = &HE0E0E0

    ss1.BlockMode = True
    ss1.Col = -1
    ss1.Row = -1
    ss1.Lock = True
    ss1.BlockMode = False
    
End Sub

Private Sub opt_del_Click(Value As Integer)

    If Sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_del.Value = True

    P_MODE = "D"
    
    opt_from.Enabled = True
    opt_to.Enabled = True
    opt_target.Enabled = False
   
    cbo_heat_mana_no.Text = ""
    cbo_heat_no2.Text = ""
    txt_tgt_heat_no.Text = ""
    opt_from.Value = True
    
    opt_line_change.BackColor = &HE0E0E0
    opt_cancel.BackColor = &HE0E0E0
    opt_del.BackColor = &HC0FFFF
    opt_mltcd_change.BackColor = &HE0E0E0
    opt_change.BackColor = &HE0E0E0
    opt_time.BackColor = &HE0E0E0
    opt_rsltdel.BackColor = &HE0E0E0
    opt_ins.BackColor = &HE0E0E0
    
    ss1.BlockMode = True
    ss1.Col = -1
    ss1.Row = -1
    ss1.Lock = True
    ss1.BlockMode = False
    
End Sub

Private Sub opt_rsltdel_Click(Value As Integer)

    Dim iRow       As Integer
    Dim sColor, TT As String
    
    If Sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_rsltdel.Value = True
    
    P_MODE = "P"
    
    opt_from.Enabled = True
    opt_to.Enabled = False
    opt_target.Enabled = False
    cbo_heat_no2.Text = ""
    txt_tgt_heat_no.Text = ""
    opt_from.Value = True

    ss1.BlockMode = True
    ss1.Col = -1
    ss1.Row = -1
    ss1.Lock = True
    ss1.BlockMode = False
    
    opt_line_change.BackColor = &HE0E0E0
    opt_cancel.BackColor = &HE0E0E0
    opt_del.BackColor = &HE0E0E0
    opt_mltcd_change.BackColor = &HE0E0E0
    opt_change.BackColor = &HE0E0E0
    opt_time.BackColor = &HE0E0E0
    opt_rsltdel.BackColor = &HC0FFFF

    With ss1
    
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 4
             sColor = .BackColor
             .Col = 1:   .Col2 = 2
             .BackColor = sColor
            If (.Text = cbo_heat_mana_no.Text) And (sColor = SSPsend.BackColor Or sColor = SSPrtn.BackColor) Then
               .BackColor = &HFF
               TT = "A"
            End If
        Next
        
        If TT <> "A" Then
           cbo_heat_mana_no.Text = ""
        End If
        
    End With
    
End Sub

Private Sub opt_send_Click(Value As Integer)

    Dim iRow As Integer
    Dim sColor As String
    
    If Sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_send.Value = True
    
    P_MODE = "L"
    
    opt_line_change.BackColor = &HE0E0E0
    opt_cancel.BackColor = &HE0E0E0
    opt_del.BackColor = &HE0E0E0
    opt_mltcd_change.BackColor = &HE0E0E0
    opt_change.BackColor = &HE0E0E0
    opt_time.BackColor = &HE0E0E0
    opt_rsltdel.BackColor = &HE0E0E0
    
    opt_from.Enabled = True
    opt_to.Enabled = True
    opt_target.Enabled = False
    opt_to.Value = True
    cbo_heat_mana_no.Text = ""
    cbo_heat_no2.Text = ""
    txt_tgt_heat_no = ""
    
    ss1.BlockMode = True
    ss1.Col = -1
    ss1.Row = -1
    ss1.Lock = True
    ss1.BlockMode = False
    
    With ss1
    
        .Col = 1
        For iRow = 1 To .MaxRows
            .Row = iRow
            If .BackColor = &HFF Then
               .Col = 4
                sColor = .BackColor
                .Col = 1:   .Col2 = 2
                .BackColor = sColor
            End If
            If .BackColor <> SSPsend.BackColor And .BackColor <> SSPpdt.BackColor And .BackColor <> SSPrtn.BackColor And cbo_heat_mana_no = "" Then
                sdb_from.Text = .Row
                cbo_heat_mana_no.Text = .Text
            End If
        Next
        
    End With
    
End Sub

Private Sub opt_from_Click(Value As Integer)

    cbo_heat_mana_no.Enabled = True
    cbo_heat_no2.Enabled = False
    txt_tgt_heat_no.Enabled = False
    
End Sub

Private Sub opt_target_Click(Value As Integer)

    txt_tgt_heat_no.Enabled = True
    cbo_heat_mana_no.Enabled = False
    cbo_heat_no2.Enabled = False
    
End Sub

Private Sub opt_time_Click(Value As Integer)

    P_MODE = "T"
    opt_from.Enabled = False
    opt_to.Enabled = False

    opt_line_change.BackColor = &HE0E0E0
    opt_cancel.BackColor = &HE0E0E0
    opt_del.BackColor = &HE0E0E0
    opt_mltcd_change.BackColor = &HE0E0E0
    opt_change.BackColor = &HE0E0E0
    opt_time.BackColor = &HC0FFFF
    opt_rsltdel.BackColor = &HE0E0E0
    
End Sub

Private Sub opt_to_Click(Value As Integer)

    cbo_heat_no2.Enabled = True
    cbo_heat_mana_no.Enabled = False
    txt_tgt_heat_no.Enabled = False
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor, sHeat, sTemp As String
    Dim sChgPrcLine          As String
    Dim sL2SendFL            As String
    Dim dHeat_edt_seq As Double
    Dim sHeat_Mana_No As String
    Dim sHeat_Design_Fl As String
    
    Set Active_Spread = Me.ss1
    If Row <= 0 Then Exit Sub
    
    If Col = 3 Then
        
        With ss1
            .Row = Row
            .Col = 0
           If UCase(Trim(ss1.Text)) = "UPDATE" Then
               .Text = ""
               .Col = Col
               .Text = ""
           Else
               .Text = "Update"
           End If
            
        End With
        
        Exit Sub

    End If
    
    ss1.Row = Row
    ss1.Col = 1
    txt_heat_mana_no.Text = ss1.Text
    ss1.Col = 21
    dHeat_edt_seq = ss1.Value
    
    Call Gp_Sp_Sort(Sc1.Item("Spread"), Col, Row)
    
    sChgPrcLine = cbo_chg_prc_line.Text

'    If opt_line_change And Col = 2 Then
'
'        With ss1
'            .ROW = ROW
'            .Col = Col
'            If .BackColor = &HC0FFEE Then
'                .Col = 0
'                If UCase(Trim(ss1.Text)) = "UPDATE" Then
'                   ' .Text = ""
'                    .Col = 2
'                    .Text = Trim(cbo_prc_line.Text)
'                Else
'                   ' .Text = "Update"
'                    .Col = 2
'                    .Text = sChgPrcLine
'                End If
'            End If
'        End With
'
'    End If
    
'    If opt_mltcd_change And Col = 7 Then
'
'        ss1.Row = Row
'        ss1.Col = 17
'        sL2SendFL = ss1.Text
'
'        ss1.Col = 14
'        txt_proc_fl.Text = ss1.Text
'
'        If sL2SendFL = "Y" Or txt_proc_fl.Text = "B" Then Exit Sub
'
'        Unload AKN2031C
'
'        ss1.Col = 4
'        AKN2031C.txt_stlgrd.Text = ss1.Text
'        ss1.Col = 7
'        AKN2031C.txt_mlt_prc_cd.Text = ss1.Text
'        AKN2031C.txt_plt.Text = txt_plt.Text
'
'        AKN2031C.Show 1
'        AKN2031C.ZOrder (0)
'
'        Exit Sub
'
'    End If

    lBlkrow1 = Row
    lBlkrow2 = Row
    
    If ss1.MaxRows < 1 Then
        Call Gf_Sp_Cls(sc2)
        Exit Sub
    End If
    
    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    sHeat_Mana_No = Gf_CodeFind(M_CN1, "SELECT CHG_HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeat_edt_seq)

    sHeat = txt_heat_mana_no.Text
    txt_heat_mana_no.Text = sHeat_Mana_No
    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    txt_heat_mana_no.Text = sHeat
    
    SS2.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(SS2)

    With ss1
        For iRow1 = .ActiveRow To .MaxRows
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            With SS2
              .Col = 3
              For iRow2 = 1 To .MaxRows
                  .Row = iRow2
                  'If stemp <> "" And stemp <> Left(.Text, 8) Then
                        If Left(.Text, 8) = sHeat Then
                           For iCol = 1 To .MaxCols
                               .Col = iCol
                               .BackColor = sColor

                           Next iCol
                           sTemp = .Text
                        End If

                        If sTemp <> "" And sTemp <> Left(.Text, 8) Then
                           sTemp = ""
                           Exit For
                        End If
                  'End If
              .Col = 3
              Next iRow2
            End With

        Next iRow1
    
    End With
    
    With SS2

        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount

             '重点订单红色标记 2013-11-15  by  CaoLei
            SS2.Row = .Row:          SS2.Col = SS_IMP_CONT
            If SS2.Text = "Y" Then
                 Call Gp_Sp_BlockColor(SS2, 1, .MaxRows, .Row, .Row, &HFF&)
            End If

        Next iCount

    End With
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim iRow, iCnt As Integer
    Dim sColor, M_TEMP As String
    
    If opt_cancel.Value = False And opt_change.Value = False _
                                And opt_del.Value = False _
                                And opt_send.Value = False _
                                And opt_line_change = False _
                                And opt_ins = False _
                                And opt_rsltdel = False Then
        Exit Sub
    End If
            
    ss1.Row = Row
    ss1.Col = 1
    txt_heat_mana_no.Text = ss1.Text
    
    If opt_line_change And Col = 2 Then
    
        'Line Status CHECK (added by KIM.SUNG.HO 2010.03.20)
        'If Not Line_Status_Chk Then Exit Sub
        '-----------------------------------------------------------------------
        
        ss1.Col = Col
        txt_AFT_SS_Col = Col
        ss1.Row = Row
        txt_AFT_SS_Row = Row
        ss1.Text = cbo_chg_prc_line.Text
        ss1.Col = 1
        Call Plc_Line_Change(Trim(ss1.Text))
        Exit Sub
       
    End If
    
    Call Search_Last_Line
    
    With ss1
    
        .Row = .ActiveRow
        .Col = 1
        
        If opt_from.Value = True Then
        
           cbo_heat_mana_no.Text = .Text
           
           If P_MODE = "C" Or P_MODE = "D" Then
           
                For iRow = 1 To .MaxRows
            
                   .Col = 4
                   .Row = iRow
                    sColor = .BackColor
                   .Col = 2: .BackColor = sColor
                   
                   .Col = 1
                   
                    If .Text <> cbo_heat_mana_no.Text Then
                    
                       .BackColor = sColor
                       
                    ElseIf .Text = cbo_heat_mana_no.Text Then
                        
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            .Row = iSelRow
                            .Col = 1
                            MsgBox "炉号(" & .Text & ")正常生产中，不能取消！", vbInformation, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
                        End If
                        
                        If P_MODE = "D" And .BackColor = SSPsend.BackColor Then
                            MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再删除！", vbCritical, "系统提示信息"
                            opt_cancel.Value = True
                        End If
                        
                        If P_MODE = "C" And .BackColor <> SSPsend.BackColor Then
                            MsgBox "您所选择的炉号不能做取消操作！", vbCritical, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
'                        ElseIf P_MODE = "C" And .ROW > p_cur_prd And .ROW <= p_cur_prd + 1 Then
'                           MsgBox "正在等待生产的第一炉不能做取消操作！", vbCritical, "系统提示信息"
'                           cbo_heat_mana_no.Text = ""
'                           Exit Sub
                        End If
                          
                        .BackColor = &HFF&
                    
                    End If
                
                Next
                
           ElseIf P_MODE = "I" Then
           
                Call Gp_Sp_EvenRowBackcolor(ss1)
                'Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 1, Row, , &HFF80FF)
                
                .Row = Row
                .Col = 1
                cbo_heat_no2.Text = .Text
                
                For iRow = 1 To Row
                    .Row = iRow
                    .Col = 1
                    .BackColor = &HFF&
                    
                    If .Row = 1 Then
                        cbo_heat_mana_no.Text = .Text
                    End If
                    
                Next iRow
    
                
           ElseIf P_MODE = "P" Then
           
                For iRow = 1 To .MaxRows
            
                   .Col = 4
                   .Row = iRow
                    sColor = .BackColor
                   .Col = 2: .BackColor = sColor
                   
                   .Col = 1
                   
                    If .Text <> cbo_heat_mana_no.Text Then
                    
                       .BackColor = sColor
                       
                    ElseIf .Text = cbo_heat_mana_no.Text Then
                        
                        .Row = iRow
                        .Col = 16
                        If .Text <> "B" Then
                            .Col = 1
                            MsgBox "炉号(" & .Text & ")不生产中，不能实绩删除！", vbInformation, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
                        End If
                        
                        .Col = 1
                        .BackColor = &HFF&
                    
                    End If
                
                Next
                
                
           ElseIf P_MODE = "M" Then
                
                For iRow = 1 To .MaxRows
                    
                    .Col = 4
                    .Row = iRow
                     sColor = .BackColor
                    .Col = 2: .BackColor = sColor
                    
                    .Col = 1
                    If .Text <> cbo_heat_mana_no.Text Then
                       .BackColor = sColor
                    ElseIf .Text = cbo_heat_mana_no.Text Then
                    
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            M_TEMP = "T"
                        End If
                        
                        If .BackColor = SSPsend.BackColor Then
                            If (P_MODE = "M") Or (P_MODE = "D") Then
                                M_TEMP = "T"
                            End If
                        End If
                        
                        .BackColor = &HFF&
                        
                    End If
                Next
        
                If M_TEMP = "T" Then
                    MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                    cbo_heat_mana_no.Text = ""
                    opt_cancel.Value = True
                End If
            
                For iRow = 1 To .MaxRows
                   .Row = iRow
                   .Col = 16
                    iProd = .Text
                   .Col = 17
                    iRet = .Text
                   .Col = 1
                    If cbo_heat_mana_no.Text = .Text Then
'                       If iRow <= p_cur_prd + 1 And iRow > p_cur_prd And iRet <> "1" Then
'                          MsgBox "正在等待生产的第一个炉号不能调整！", vbInformation, "系统提示信息"
'                          cbo_heat_mana_no.Text = ""
'                          Exit Sub
'                       Else
                          If iProd <> "" And iRet = "0" Then
                             MsgBox "该炉号正常生产中，不能调整！", vbInformation, "系统提示信息"
                             cbo_heat_mana_no.Text = ""
                             Exit Sub
                          Else
                            .Col = 1
                             
                             cbo_heat_no2.Text = ""
                          End If
'                       End If
                    End If
                Next
                
           End If
           
           If P_MODE = "D" Or P_MODE = "L" Or P_MODE = "M" Then
                opt_to.Value = True
           End If
           
        ElseIf opt_to.Value = True Then
        
            If cbo_heat_mana_no.Text <> "" And Mid(cbo_heat_mana_no, 3, 1) <> Mid(.Text, 3, 1) Then
                Exit Sub
            End If
            
            cbo_heat_no2.Text = .Text
           
            If cbo_heat_no2.BackColor = SSPsend.BackColor Then
                MsgBox "请先从炉号 <" + cbo_heat_no2.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                opt_cancel.Value = True
                Exit Sub
            End If
           
            If P_MODE = "M" Or P_MODE = "L" Or P_MODE = "D" Then
           
               For iRow = 1 To .MaxRows
                   .Col = 1
                   .Row = iRow
                      
'                   If cbo_heat_no2.Text = .Text Then
'                      If (iRow <= p_cur_prd + 1) And iRet = "0" And iRow > p_cur_prd Then
'                         MsgBox "正在等待生产的第一个炉号不能调整！", vbInformation, "系统提示信息"
'                         Exit Sub
'                      End If
'                   End If
                    
                   If .Text = cbo_heat_no2.Text Then
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            M_TEMP = "T"
                        End If
                        
                        If .BackColor = SSPsend.BackColor Then
                            M_TEMP = "T"
                        End If
                        .BackColor = &HFF&
                    End If
                    
                Next
        
                If M_TEMP = "T" Then
                    MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                    opt_cancel.Value = True
                    Exit Sub
                End If
                
            ElseIf P_MODE = "I" Then
           
                Call Gp_Sp_EvenRowBackcolor(ss1)
                'Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 1, Row, , &HFF80FF)
                
                .Row = Row
                .Col = 1
                cbo_heat_no2.Text = .Text
                
                For iRow = 1 To Row
                    .Row = iRow
                    .Col = 1
                    .BackColor = &HFF&
                    
                    If .Row = 1 Then
                        cbo_heat_mana_no.Text = .Text
                    End If
                    
                Next iRow
                
           End If
           
           If P_MODE = "M" Then
                opt_target.Value = True
           End If
           
        ElseIf opt_target.Value = True Then
        
            If cbo_heat_mana_no.Text <> "" And Mid(cbo_heat_no2, 3, 1) <> Mid(.Text, 3, 1) Then
                Exit Sub
            End If
            
            txt_tgt_heat_no.Text = .Text
            
            If P_MODE = "M" Or P_MODE = "D" Then
            
                If txt_tgt_heat_no.Text >= cbo_heat_mana_no.Text And txt_tgt_heat_no.Text <= cbo_heat_no2.Text Then
                    MsgBox "目标炉号在选定的起始和最终炉号内！", vbInformation, "系统提示信息"
                    Exit Sub
                End If
                
               For iRow = 1 To .MaxRows
               
                   .Row = iRow
                   .Col = 16
                    iProd = .Text
                   .Col = 17
                    iRet = .Text
                   .Col = 1
                   
                    If txt_tgt_heat_no.Text = .Text Then
                        .Col = 19
                        If .Text = "Y" Then
                            MsgBox "请先从炉号 <" + txt_tgt_heat_no.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                            txt_tgt_heat_no.Text = ""
                            Exit Sub
                        End If
                        
                        If ((iProd <> "" And iRet = "0") Or iSelRow >= iRow) And iSelRow <> 1 Then
                            MsgBox "该炉号正常生产中，不能调整！", vbInformation, "系统提示信息"
                            txt_tgt_heat_no.Text = ""
                            Exit Sub
                        End If
                       
                        .BackColor = &HFF8080
                        
                    End If
                    
                Next
                
            End If
            
        End If
        
        If cbo_heat_mana_no <> "" And cbo_heat_no2 <> "" Then
        
              For iRow = 1 To .MaxRows
                  .Row = iRow
                  .Col = 4
                  sColor = .BackColor
                  .Col = 2: .BackColor = sColor
                  
                  .Col = 1
                  If (.Text >= cbo_heat_mana_no.Text And .Text <= cbo_heat_no2.Text) Or (.Text = txt_tgt_heat_no.Text) Then
                     .BackColor = &HFF&
                  Else
                     .BackColor = sColor
                  End If
                  
             Next
             
        End If
        
    End With
    
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.SS2
    
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If
    
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.SS2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor, sHeat, sTemp As String
    Dim sChgPrcLine          As String
    Dim dHeat_edt_seq As Double
    Dim sHeat_Mana_No As String
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.SS3
    If Row <= 0 Then Exit Sub
    
    If SS3.MaxRows < 1 Then
        Call Gf_Sp_Cls(Sc4)
        Exit Sub
    End If
    
    SS3.Row = Row
    SS3.Col = 1
    txt_heat_mana_no.Text = SS3.Text
    SS3.Col = 21
    dHeat_edt_seq = SS3.Value
    
    sHeat_Mana_No = Gf_CodeFind(M_CN1, "SELECT CHG_HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeat_edt_seq)

    sHeat = txt_heat_mana_no.Text
    txt_heat_mana_no.Text = sHeat_Mana_No
    Call Gf_Sp_Refer(M_CN1, Sc4, Mc2, , , False)
    txt_heat_mana_no.Text = sHeat
    
    ss4.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss4)

    With SS3
    
        For iRow1 = .ActiveRow To .MaxRows
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            
            With ss4
            
                .Col = 3
                For iRow2 = 1 To .MaxRows
                    .Row = iRow2
                    If Left(.Text, 8) = sHeat Then
                       For iCol = 1 To .MaxCols
                           .Col = iCol
                           .BackColor = sColor
                       Next iCol
                       sTemp = .Text
                    End If
                    
                    If sTemp <> "" And sTemp <> Left(.Text, 8) Then
                       sTemp = ""
                       Exit For
                    End If
                    .Col = 3
                Next iRow2
                
            End With

        Next iRow1
        
    End With
    
    With ss4

        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount

             '重点订单红色标记 2013-11-15  by  CaoLei
            ss4.Row = .Row:          ss4.Col = SS_IMP_CONT
            If SS2.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss4, 1, .MaxRows, .Row, .Row, &HFF&)
            End If

        Next iCount

    End With
    
End Sub

Private Sub ss3_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim iRow, iCnt As Integer
    Dim sColor, M_TEMP As String
    
    If opt_cancel.Value = False And opt_change.Value = False _
                                And opt_del.Value = False _
                                And opt_send.Value = False _
                                And opt_line_change = False _
                                And opt_ins = False _
                                And opt_rsltdel = False Then
        Exit Sub
    End If
            
    SS3.Row = Row
    SS3.Col = 1
    txt_heat_mana_no.Text = SS3.Text
    
    If opt_line_change And Col = 2 Then
    
        'Line Status CHECK (added by KIM.SUNG.HO 2010.03.20)
        'If Not Line_Status_Chk Then Exit Sub
        '-----------------------------------------------------------------------
        
        SS3.Col = Col
        txt_AFT_SS_Col = Col
        SS3.Row = Row
        txt_AFT_SS_Row = Row
        SS3.Text = cbo_chg_prc_line.Text
        SS3.Col = 1
        Call Plc_Line_Change(Trim(SS3.Text))
        Exit Sub
       
    End If
    
    Call Search_Last_Line3
    
    With SS3
    
        .Row = .ActiveRow
        .Col = 1
        
        If opt_from.Value = True Then
        
           cbo_heat_mana_no.Text = .Text
           
           If P_MODE = "C" Or P_MODE = "D" Then
           
                For iRow = 1 To .MaxRows
            
                   .Col = 4
                   .Row = iRow
                    sColor = .BackColor
                   .Col = 2: .BackColor = sColor
                   
                   .Col = 1
                   
                    If .Text <> cbo_heat_mana_no.Text Then
                    
                       .BackColor = sColor
                       
                    ElseIf .Text = cbo_heat_mana_no.Text Then
                        
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            .Row = iSelRow
                            .Col = 1
                            MsgBox "炉号(" & .Text & ")正常生产中，不能取消！", vbInformation, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
                        End If
                        
                        If P_MODE = "D" And .BackColor = SSPsend.BackColor Then
                            MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再删除！", vbCritical, "系统提示信息"
                            opt_cancel.Value = True
                        End If
                        
                        If P_MODE = "C" And .BackColor <> SSPsend.BackColor Then
                            MsgBox "您所选择的炉号不能做取消操作！", vbCritical, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
'                        ElseIf P_MODE = "C" And .ROW > p_cur_prd And .ROW <= p_cur_prd + 1 Then
'                           MsgBox "正在等待生产的第一炉不能做取消操作！", vbCritical, "系统提示信息"
'                           cbo_heat_mana_no.Text = ""
'                           Exit Sub
                        End If
                          
                        .BackColor = &HFF&
                    
                    End If
                
                Next
                
           ElseIf P_MODE = "I" Then
           
                Call Gp_Sp_EvenRowBackcolor(SS3)
                'Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, 1, Row, , &HFF80FF)
                
                .Row = Row
                .Col = 1
                cbo_heat_no2.Text = .Text
                
                For iRow = 1 To Row
                    .Row = iRow
                    .Col = 1
                    .BackColor = &HFF&
                    
                    If .Row = 1 Then
                        cbo_heat_mana_no.Text = .Text
                    End If
                    
                Next iRow
    
                
           ElseIf P_MODE = "P" Then
           
                For iRow = 1 To .MaxRows
            
                   .Col = 4
                   .Row = iRow
                    sColor = .BackColor
                   .Col = 2: .BackColor = sColor
                   
                   .Col = 1
                   
                    If .Text <> cbo_heat_mana_no.Text Then
                    
                       .BackColor = sColor
                       
                    ElseIf .Text = cbo_heat_mana_no.Text Then
                        
                        .Row = iRow
                        .Col = 16
                        If .Text <> "B" Then
                            .Col = 1
                            MsgBox "炉号(" & .Text & ")不生产中，不能实绩删除！", vbInformation, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
                        End If
                        
                        .Col = 1
                        .BackColor = &HFF&
                    
                    End If
                
                Next
                
                
           ElseIf P_MODE = "M" Then
                
                For iRow = 1 To .MaxRows
                    
                    .Col = 4
                    .Row = iRow
                     sColor = .BackColor
                    .Col = 2: .BackColor = sColor
                    
                    .Col = 1
                    If .Text <> cbo_heat_mana_no.Text Then
                       .BackColor = sColor
                    ElseIf .Text = cbo_heat_mana_no.Text Then
                    
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            M_TEMP = "T"
                        End If
                        
                        If .BackColor = SSPsend.BackColor Then
                            If (P_MODE = "M") Or (P_MODE = "D") Then
                                M_TEMP = "T"
                            End If
                        End If
                        
                        .BackColor = &HFF&
                        
                    End If
                Next
        
                If M_TEMP = "T" Then
                    MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                    cbo_heat_mana_no.Text = ""
                    opt_cancel.Value = True
                End If
            
                For iRow = 1 To .MaxRows
                   .Row = iRow
                   .Col = 16
                    iProd = .Text
                   .Col = 17
                    iRet = .Text
                   .Col = 1
                    If cbo_heat_mana_no.Text = .Text Then
'                       If iRow <= p_cur_prd + 1 And iRow > p_cur_prd And iRet <> "1" Then
'                          MsgBox "正在等待生产的第一个炉号不能调整！", vbInformation, "系统提示信息"
'                          cbo_heat_mana_no.Text = ""
'                          Exit Sub
'                       Else
                          If iProd <> "" And iRet = "0" Then
                             MsgBox "该炉号正常生产中，不能调整！", vbInformation, "系统提示信息"
                             cbo_heat_mana_no.Text = ""
                             Exit Sub
                          Else
                            .Col = 1
                             
                             cbo_heat_no2.Text = ""
                          End If
'                       End If
                    End If
                Next
                
            End If
           
            If P_MODE = "D" Or P_MODE = "L" Or P_MODE = "M" Then
                opt_to.Value = True
            End If
           
        ElseIf opt_to.Value = True Then
        
            If cbo_heat_mana_no.Text <> "" And Mid(cbo_heat_mana_no, 3, 1) <> Mid(.Text, 3, 1) Then
                Exit Sub
            End If
            
            cbo_heat_no2.Text = .Text
           
            If cbo_heat_no2.BackColor = SSPsend.BackColor Then
                MsgBox "请先从炉号 <" + cbo_heat_no2.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                opt_cancel.Value = True
                Exit Sub
            End If
           
           If P_MODE = "M" Or P_MODE = "L" Or P_MODE = "D" Then
           
               For iRow = 1 To .MaxRows
                   .Col = 1
                   .Row = iRow
                      
'                   If cbo_heat_no2.Text = .Text Then
'                      If (iRow <= p_cur_prd + 1) And iRet = "0" And iRow > p_cur_prd Then
'                         MsgBox "正在等待生产的第一个炉号不能调整！", vbInformation, "系统提示信息"
'                         Exit Sub
'                      End If
'                   End If
                    
                   If .Text = cbo_heat_no2.Text Then
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            M_TEMP = "T"
                        End If
                        
                        If .BackColor = SSPsend.BackColor Then
                            M_TEMP = "T"
                        End If
                        .BackColor = &HFF&
                    End If
                    
                Next
        
                If M_TEMP = "T" Then
                    MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                    opt_cancel.Value = True
                    Exit Sub
                End If
                
            ElseIf P_MODE = "I" Then
           
                Call Gp_Sp_EvenRowBackcolor(SS3)
                'Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, 1, Row, , &HFF80FF)
                
                .Row = Row
                .Col = 1
                cbo_heat_no2.Text = .Text
                
                For iRow = 1 To Row
                    .Row = iRow
                    .Col = 1
                    .BackColor = &HFF&
                    
                    If .Row = 1 Then
                        cbo_heat_mana_no.Text = .Text
                    End If
                    
                Next iRow
                
           End If
           
           If P_MODE = "M" Then
                opt_target.Value = True
           End If
           
        ElseIf opt_target.Value = True Then
        
            If cbo_heat_mana_no.Text <> "" And Mid(cbo_heat_no2, 3, 1) <> Mid(.Text, 3, 1) Then
                Exit Sub
            End If
            
            txt_tgt_heat_no.Text = .Text
            
            If P_MODE = "M" Or P_MODE = "D" Then
            
                If txt_tgt_heat_no.Text >= cbo_heat_mana_no.Text And txt_tgt_heat_no.Text <= cbo_heat_no2.Text Then
                    MsgBox "目标炉号在选定的起始和最终炉号内！", vbInformation, "系统提示信息"
                    Exit Sub
                End If
                
               For iRow = 1 To .MaxRows
               
                   .Row = iRow
                   .Col = 16
                    iProd = .Text
                   .Col = 17
                    iRet = .Text
                   .Col = 1
                   
                    If txt_tgt_heat_no.Text = .Text Then
                        .Col = 19
                        If .Text = "Y" Then
                            MsgBox "请先从炉号 <" + txt_tgt_heat_no.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                            txt_tgt_heat_no.Text = ""
                            Exit Sub
                        End If
                        
                        If ((iProd <> "" And iRet = "0") Or iSelRow >= iRow) And iSelRow <> 1 Then
                            MsgBox "该炉号正常生产中，不能调整！", vbInformation, "系统提示信息"
                            txt_tgt_heat_no.Text = ""
                            Exit Sub
                        End If
                       
                        .BackColor = &HFF8080
                        
                    End If
                    
                Next
                
            End If
            
        End If
        
        If cbo_heat_mana_no <> "" And cbo_heat_no2 <> "" Then
        
              For iRow = 1 To .MaxRows
                  .Row = iRow
                  .Col = 4
                  sColor = .BackColor
                  .Col = 2: .BackColor = sColor
                  
                  .Col = 1
                  If (.Text >= cbo_heat_mana_no.Text And .Text <= cbo_heat_no2.Text) Or (.Text = txt_tgt_heat_no.Text) Then
                     .BackColor = &HFF&
                  Else
                     .BackColor = sColor
                  End If
                  
             Next
             
        End If
        
    End With

End Sub

Private Sub ss3_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.SS3
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss4
    
End Sub

Private Sub ss5_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss5_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor, sHeat, sTemp As String
    Dim sChgPrcLine          As String
    Dim dHeat_edt_seq As Double
    Dim sHeat_Mana_No As String
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss5
    If Row <= 0 Then Exit Sub
    
    If ss5.MaxRows < 1 Then
        Call Gf_Sp_Cls(Sc6)
        Exit Sub
    End If
    
    ss5.Row = Row
    ss5.Col = 1
    txt_heat_mana_no.Text = ss5.Text
    ss5.Col = 21
    dHeat_edt_seq = ss5.Value
    
    sHeat_Mana_No = Gf_CodeFind(M_CN1, "SELECT CHG_HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeat_edt_seq)

    sHeat = txt_heat_mana_no.Text
    txt_heat_mana_no.Text = sHeat_Mana_No
    Call Gf_Sp_Refer(M_CN1, Sc6, Mc2, , , False)
    txt_heat_mana_no.Text = sHeat
    
    ss6.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss6)

    With ss5
    
        For iRow1 = .ActiveRow To .MaxRows
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            
            With ss6
            
                .Col = 3
                For iRow2 = 1 To .MaxRows
                    .Row = iRow2
                    If Left(.Text, 8) = sHeat Then
                       For iCol = 1 To .MaxCols
                           .Col = iCol
                           .BackColor = sColor
                       Next iCol
                       sTemp = .Text
                    End If
                    
                    If sTemp <> "" And sTemp <> Left(.Text, 8) Then
                       sTemp = ""
                       Exit For
                    End If
                    .Col = 3
                Next iRow2
                
            End With

        Next iRow1
        
    End With
    
    With ss6

        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount

             '重点订单红色标记 2013-11-15  by  CaoLei
            ss6.Row = .Row:          ss6.Col = SS_IMP_CONT
            If ss6.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss6, 1, .MaxRows, .Row, .Row, &HFF&)
            End If

        Next iCount

    End With
    
End Sub

Private Sub ss5_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim iRow, iCnt As Integer
    Dim sColor, M_TEMP As String
    
    If opt_cancel.Value = False And opt_change.Value = False _
                                And opt_del.Value = False _
                                And opt_send.Value = False _
                                And opt_line_change = False _
                                And opt_ins = False _
                                And opt_rsltdel = False Then
        Exit Sub
    End If
            
    ss5.Row = Row
    ss5.Col = 1
    txt_heat_mana_no.Text = ss5.Text
    
    If opt_line_change And Col = 2 Then
    
        'Line Status CHECK (added by KIM.SUNG.HO 2010.03.20)
        'If Not Line_Status_Chk Then Exit Sub
        '-----------------------------------------------------------------------
        
        ss5.Col = Col
        txt_AFT_SS_Col = Col
        ss5.Row = Row
        txt_AFT_SS_Row = Row
        ss5.Text = cbo_chg_prc_line.Text
        ss5.Col = 1
        Call Plc_Line_Change(Trim(ss5.Text))
        Exit Sub
       
    End If
    
    Call Search_Last_Line5
    
    With ss5
    
        .Row = .ActiveRow
        .Col = 1
        
        If opt_from.Value = True Then
        
           cbo_heat_mana_no.Text = .Text
           
           If P_MODE = "C" Or P_MODE = "D" Then
           
                For iRow = 1 To .MaxRows
            
                   .Col = 4
                   .Row = iRow
                    sColor = .BackColor
                   .Col = 2: .BackColor = sColor
                   
                   .Col = 1
                   
                    If .Text <> cbo_heat_mana_no.Text Then
                    
                       .BackColor = sColor
                       
                    ElseIf .Text = cbo_heat_mana_no.Text Then
                        
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            .Row = iSelRow
                            .Col = 1
                            MsgBox "炉号(" & .Text & ")正常生产中，不能取消！", vbInformation, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
                        End If
                        
                        If P_MODE = "D" And .BackColor = SSPsend.BackColor Then
                            MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再删除！", vbCritical, "系统提示信息"
                            opt_cancel.Value = True
                        End If
                        
                        If P_MODE = "C" And .BackColor <> SSPsend.BackColor Then
                            MsgBox "您所选择的炉号不能做取消操作！", vbCritical, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
'                        ElseIf P_MODE = "C" And .ROW > p_cur_prd And .ROW <= p_cur_prd + 1 Then
'                           MsgBox "正在等待生产的第一炉不能做取消操作！", vbCritical, "系统提示信息"
'                           cbo_heat_mana_no.Text = ""
'                           Exit Sub
                        End If
                          
                        .BackColor = &HFF&
                    
                    End If
                
                Next
                
           ElseIf P_MODE = "I" Then
           
                Call Gp_Sp_EvenRowBackcolor(ss5)
                'Call Gp_Sp_BlockColor(ss5, 1, ss5.MaxCols, 1, Row, , &HFF80FF)
                
                .Row = Row
                .Col = 1
                cbo_heat_no2.Text = .Text
                
                For iRow = 1 To Row
                    .Row = iRow
                    .Col = 1
                    .BackColor = &HFF&
                    
                    If .Row = 1 Then
                        cbo_heat_mana_no.Text = .Text
                    End If
                    
                Next iRow
    
                
           ElseIf P_MODE = "P" Then
           
                For iRow = 1 To .MaxRows
            
                   .Col = 4
                   .Row = iRow
                    sColor = .BackColor
                   .Col = 2: .BackColor = sColor
                   
                   .Col = 1
                   
                    If .Text <> cbo_heat_mana_no.Text Then
                    
                       .BackColor = sColor
                       
                    ElseIf .Text = cbo_heat_mana_no.Text Then
                        
                        .Row = iRow
                        .Col = 16
                        If .Text <> "B" Then
                            .Col = 1
                            MsgBox "炉号(" & .Text & ")不生产中，不能实绩删除！", vbInformation, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
                        End If
                        
                        .Col = 1
                        .BackColor = &HFF&
                    
                    End If
                
                Next
                
                
           ElseIf P_MODE = "M" Then
                
                For iRow = 1 To .MaxRows
                    
                    .Col = 4
                    .Row = iRow
                     sColor = .BackColor
                    .Col = 2: .BackColor = sColor
                    
                    .Col = 1
                    If .Text <> cbo_heat_mana_no.Text Then
                       .BackColor = sColor
                    ElseIf .Text = cbo_heat_mana_no.Text Then
                    
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            M_TEMP = "T"
                        End If
                        
                        If .BackColor = SSPsend.BackColor Then
                            If (P_MODE = "M") Or (P_MODE = "D") Then
                                M_TEMP = "T"
                            End If
                        End If
                        
                        .BackColor = &HFF&
                        
                    End If
                Next
        
                If M_TEMP = "T" Then
                    MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                    cbo_heat_mana_no.Text = ""
                    opt_cancel.Value = True
                End If
            
                For iRow = 1 To .MaxRows
                   .Row = iRow
                   .Col = 16
                    iProd = .Text
                   .Col = 17
                    iRet = .Text
                   .Col = 1
                    If cbo_heat_mana_no.Text = .Text Then
'                       If iRow <= p_cur_prd + 1 And iRow > p_cur_prd And iRet <> "1" Then
'                          MsgBox "正在等待生产的第一个炉号不能调整！", vbInformation, "系统提示信息"
'                          cbo_heat_mana_no.Text = ""
'                          Exit Sub
'                       Else
                          If iProd <> "" And iRet = "0" Then
                             MsgBox "该炉号正常生产中，不能调整！", vbInformation, "系统提示信息"
                             cbo_heat_mana_no.Text = ""
                             Exit Sub
                          Else
                            .Col = 1
                             
                             cbo_heat_no2.Text = ""
                          End If
'                       End If
                    End If
                Next
                
           End If
           
           If P_MODE = "D" Or P_MODE = "L" Or P_MODE = "M" Then
                opt_to.Value = True
           End If
           
        ElseIf opt_to.Value = True Then
        
            If cbo_heat_mana_no.Text <> "" And Mid(cbo_heat_mana_no, 3, 1) <> Mid(.Text, 3, 1) Then
                Exit Sub
            End If
            
            cbo_heat_no2.Text = .Text
           
            If cbo_heat_no2.BackColor = SSPsend.BackColor Then
                MsgBox "请先从炉号 <" + cbo_heat_no2.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                opt_cancel.Value = True
                Exit Sub
            End If
           
           If P_MODE = "M" Or P_MODE = "L" Or P_MODE = "D" Then
           
               For iRow = 1 To .MaxRows
                   .Col = 1
                   .Row = iRow
                      
'                   If cbo_heat_no2.Text = .Text Then
'                      If (iRow <= p_cur_prd + 1) And iRet = "0" And iRow > p_cur_prd Then
'                         MsgBox "正在等待生产的第一个炉号不能调整！", vbInformation, "系统提示信息"
'                         Exit Sub
'                      End If
'                   End If
                    
                   If .Text = cbo_heat_no2.Text Then
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            M_TEMP = "T"
                        End If
                        
                        If .BackColor = SSPsend.BackColor Then
                            M_TEMP = "T"
                        End If
                        .BackColor = &HFF&
                    End If
                    
                Next
        
                If M_TEMP = "T" Then
                    MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                    opt_cancel.Value = True
                    Exit Sub
                End If
                
            ElseIf P_MODE = "I" Then
           
                Call Gp_Sp_EvenRowBackcolor(ss5)
                'Call Gp_Sp_BlockColor(ss5, 1, ss5.MaxCols, 1, Row, , &HFF80FF)
                
                .Row = Row
                .Col = 1
                cbo_heat_no2.Text = .Text
                
                For iRow = 1 To Row
                    .Row = iRow
                    .Col = 1
                    .BackColor = &HFF&
                    
                    If .Row = 1 Then
                        cbo_heat_mana_no.Text = .Text
                    End If
                    
                Next iRow
                
            End If
           
            If P_MODE = "M" Then
                opt_target.Value = True
            End If
           
        ElseIf opt_target.Value = True Then
        
            If cbo_heat_mana_no.Text <> "" And Mid(cbo_heat_no2, 3, 1) <> Mid(.Text, 3, 1) Then
                Exit Sub
            End If
            
            txt_tgt_heat_no.Text = .Text
            
            If P_MODE = "M" Or P_MODE = "D" Then
            
                If txt_tgt_heat_no.Text >= cbo_heat_mana_no.Text And txt_tgt_heat_no.Text <= cbo_heat_no2.Text Then
                    MsgBox "目标炉号在选定的起始和最终炉号内！", vbInformation, "系统提示信息"
                    Exit Sub
                End If
                
               For iRow = 1 To .MaxRows
               
                   .Row = iRow
                   .Col = 16
                    iProd = .Text
                   .Col = 17
                    iRet = .Text
                   .Col = 1
                   
                    If txt_tgt_heat_no.Text = .Text Then
                        .Col = 19
                        If .Text = "Y" Then
                            MsgBox "请先从炉号 <" + txt_tgt_heat_no.Text + "> 开始取消已下达到二级的作业指示，然后再调整！", vbCritical, "系统提示信息"
                            txt_tgt_heat_no.Text = ""
                            Exit Sub
                        End If
                        
                        If ((iProd <> "" And iRet = "0") Or iSelRow >= iRow) And iSelRow <> 1 Then
                            MsgBox "该炉号正常生产中，不能调整！", vbInformation, "系统提示信息"
                            txt_tgt_heat_no.Text = ""
                            Exit Sub
                        End If
                       
                        .BackColor = &HFF8080
                        
                    End If
                    
                Next
                
            End If
            
        End If
        
        If cbo_heat_mana_no <> "" And cbo_heat_no2 <> "" Then
        
              For iRow = 1 To .MaxRows
                  .Row = iRow
                  .Col = 4
                  sColor = .BackColor
                  .Col = 2: .BackColor = sColor
                  
                  .Col = 1
                  If (.Text >= cbo_heat_mana_no.Text And .Text <= cbo_heat_no2.Text) Or (.Text = txt_tgt_heat_no.Text) Then
                     .BackColor = &HFF&
                  Else
                     .BackColor = sColor
                  End If
                  
             Next
             
        End If
        
    End With

End Sub

Private Sub ss5_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss5
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

End Sub

Private Sub ss6_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss6
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_name
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
        
        If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_name.Text = ""
        End If
        
    End If

End Sub

Public Sub Plc_Line_Change(sHeatManaNo As String)

    Dim sHeatEdtSeq       As String
    Dim sChgNoNew         As String
    Dim OutParam(1, 4)    As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery            As String
    Dim P_TYPE            As String
    Dim adoCmd            As adodb.Command
    
    Set AdoRs = New adodb.Recordset
        
    sQuery = "         SELECT HEAT_EDT_SEQ   " & vbCrLf
    sQuery = sQuery & "  FROM EP_CHARGE_IDX       " & vbCrLf
    sQuery = sQuery & " WHERE HEAT_MANA_NO      = '" & sHeatManaNo & "'   " & vbCrLf
    sQuery = sQuery & "   AND PRC_STS           =    'A'                " & vbCrLf
            
    AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
          
    If AdoRs.BOF Or AdoRs.EOF Then
        Call MsgBox("不可以变！==> 作业中", vbInformation, "系统提示信息")
        ss1.Col = txt_AFT_SS_Col
        ss1.Row = txt_AFT_SS_Row
        ss1.Text = cbo_chg_prc_line.Text
        Exit Sub
    Else
        sHeatEdtSeq = AdoRs.Fields(0)
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
'--------
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
     
    sQuery = "{call AFZ9110P ('" + sHeatEdtSeq + "','" + cbo_chg_prc_line.Text + "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    M_CN1.BeginTrans
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        errMsg = sErrMessg
        M_CN1.RollbackTrans
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        
        ss1.Col = txt_AFT_SS_Col
        ss1.Row = txt_AFT_SS_Row
        ss1.Text = cbo_chg_prc_line.Text
        Exit Sub
    Else
'        If P_MODE = "L" Then
'           Call MsgBox("作业指示已成功下达！", vbInformation, "系统提示信息")
'           Call Form_Ref
'        ElseIf P_MODE = "T" Then
'           Call MsgBox("作业指示中的时间信息已调整成功！请及时将调整后的时间下达给二级系统！", vbInformation, "系统提示信息")
'           Call Form_Ref
'        End If
    End If
    
    M_CN1.CommitTrans
    Set adoCmd = Nothing
'--------
    If Trim(errMsg) = "" Then Call Form_Ref

    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub

Public Sub ChageNo_Search(lSeq As Long, sChargeNo As String, sPlcLine As String, sStrChrNo1 As String, sStrChrNo2 As String)
    
    Dim sQuery      As String
    Dim sChrNoOrg   As String
    Dim sChgNoNew   As String
    
    'NOT USED
    Exit Sub
    
    
    Set AdoRs = New adodb.Recordset
        
    sQuery = "         SELECT MAX(HEAT_MANA_NO)   " & vbCrLf
    sQuery = sQuery & "  FROM EP_CHARGE_IDX       " & vbCrLf
    sQuery = sQuery & " WHERE SUBSTR(HEAT_MANA_NO,1,3) = '" & Left(sChargeNo, 3) & "' " & vbCrLf
    sQuery = sQuery & "   AND CHG_SEQ <   " & lSeq
    
    AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly

    sChrNoOrg = AdoRs(0) & ""
    
    If Trim(sChrNoOrg) = "" Then
        sChgNoNew = Left(sChargeNo, 3) & "00001"
    Else
        sChgNoNew = Left(sChrNoOrg, 3) & Format(Val(Mid(sChrNoOrg, 4, 5) & "") + 1, "00000")
    End If
    
    AdoRs.Close
       
    If Trim(cbo_prc_line.Text) = "2" Then
        sStrChrNo2 = sChgNoNew
    Else
        sStrChrNo1 = sChgNoNew
    End If
    
    sQuery = "         SELECT MAX(HEAT_MANA_NO)            " & vbCrLf
    sQuery = sQuery & "  FROM EP_CHARGE_IDX                " & vbCrLf
    sQuery = sQuery & " WHERE PRC_LINE = '" & sPlcLine & "'" & vbCrLf
    
    AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly

    sChrNoOrg = AdoRs(0) & ""
    
    If Trim(sChrNoOrg) = "" Then
        sChgNoNew = Left(sChargeNo, 2) & sPlcLine & "00001"
    Else
        sChgNoNew = Left(sChrNoOrg, 3) & Format(Val(Mid(sChrNoOrg, 4, 5) & "") + 1, "00000")
    End If
    
    AdoRs.Close
    
    If Trim(cbo_prc_line.Text) = "2" Then
        sStrChrNo1 = sChgNoNew
    Else
        sStrChrNo2 = sChgNoNew
    End If
    
End Sub

Public Sub DataCompareHigh()

    Dim sQuery      As String
    Dim sProdFl     As String
    Dim sL2Send     As String
    Dim dChkSeq     As Double
    Dim dSeverSeq   As Double
    Dim iRow        As Long
    
    dChkSeq = 0
    For iRow = ss1.MaxRows To 1 Step -1
        ss1.Row = iRow
        ss1.Col = 15
        sProdFl = Trim(ss1.Text)
        
        ss1.Col = 19
        sL2Send = Trim(ss1.Text)
        
        If sL2Send = "Y" Or sProdFl >= "B" Then
            ss1.Col = 14
            dChkSeq = Val(ss1.Value & "")
            Exit For
        End If
    Next iRow
    
    Set AdoRs = New adodb.Recordset
    
    If dChkSeq = 0 Then
        sQuery = "         SELECT  MAX(CHG_SEQ)    " & vbCrLf
        sQuery = sQuery & "  FROM  EP_CHARGE_IDX   " & vbCrLf
        ss1.Row = 1:   ss1.Col = 14
        sQuery = sQuery & " WHERE  CHG_SEQ         >=    " & Val(ss1.Value & "") & vbCrLf
        sQuery = sQuery & "   AND  PRC_LINE         = '" & Trim(cbo_prc_line.Text) & "'" & vbCrLf
        sQuery = sQuery & "   AND (NVL(L2_SEND,'N') = 'Y'" & vbCrLf
        sQuery = sQuery & "    OR  PRC_STS         >= 'B') " & vbCrLf
    Else
        sQuery = "         SELECT  MAX(CHG_SEQ)    " & vbCrLf
        sQuery = sQuery & "  FROM  EP_CHARGE_IDX   " & vbCrLf
        sQuery = sQuery & " WHERE (NVL(L2_SEND,'N') = 'Y'  " & vbCrLf
        sQuery = sQuery & "    OR  PRC_STS         >= 'B') " & vbCrLf
        sQuery = sQuery & "   AND  PRC_LINE         = '" & Trim(cbo_prc_line.Text) & "'" & vbCrLf
    End If
    
    AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly

    dSeverSeq = Val(AdoRs(0) & "")
    
    AdoRs.Close
    
    If dChkSeq = 0 Then
        If dSeverSeq <> 0 Then
            errMsg = "再查询以后重新处理吧..！"
        End If
    Else
        If dChkSeq <> dSeverSeq Then
            errMsg = "再查询以后重新处理吧..！"
        End If
    End If
    
End Sub

Public Sub DataCompareLow()

    Dim sQuery      As String
    Dim dChkSeq     As Double
    Dim dSeverSeq   As Double
    Dim iRow        As Long
    
    dChkSeq = 0
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 16
        
        If Trim(ss1.Text) = "B" Then
            ss1.Col = 14
            dChkSeq = Val(ss1.Value & "")
            Exit For
        End If
    Next iRow
    
    Set AdoRs = New adodb.Recordset
    
    sQuery = "         SELECT  MIN(CHG_SEQ)    " & vbCrLf
    sQuery = sQuery & "  FROM  EP_CHARGE_IDX   " & vbCrLf
    sQuery = sQuery & " WHERE  PRC_STS  =  'B' " & vbCrLf
    sQuery = sQuery & "   AND  PRC_LINE = '" & Trim(cbo_prc_line.Text) & "'" & vbCrLf
    
    AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly

    dSeverSeq = Val(AdoRs(0) & "")
    
    AdoRs.Close
    
    If dChkSeq <> dSeverSeq Then
        errMsg = "再查询以后重新处理吧..！"
    End If
    
End Sub

Public Sub Gp_Schedul_Send()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim P_TYPE  As String
    Dim adoCmd As adodb.Command
    Dim lRow As Integer
    Dim dHeatSeq_Str, dHeatSeq_Stp, dHeatSeq_Tgt As Double
    Dim sHeatNo_Str, sHeatNo_Stp, sHeatNo_Tgt As String
    
    P_TYPE = "S"
    
    Screen.MousePointer = vbHourglass
    
    If Mid(cbo_heat_mana_no, 3, 1) = "1" Then
    
        For lRow = 1 To ss1.MaxRows
        
            ss1.Row = lRow
            ss1.Col = 1
            
            If cbo_heat_mana_no.Text = ss1.Text Then
                ss1.Col = 21
                dHeatSeq_Str = ss1.Value
                sHeatNo_Str = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeatSeq_Str)
            End If
            
            ss1.Col = 1
            If cbo_heat_no2.Text = ss1.Text Then
                ss1.Col = 21
                dHeatSeq_Stp = ss1.Value
                sHeatNo_Stp = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeatSeq_Stp)
            End If
            
            If P_MODE = "M" Then
            
                ss1.Col = 1
                If txt_tgt_heat_no.Text = ss1.Text Then
                    ss1.Col = 21
                    dHeatSeq_Tgt = ss1.Value
                    sHeatNo_Tgt = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeatSeq_Tgt)
                End If
            
            End If
        
        Next lRow
    
    ElseIf Mid(cbo_heat_mana_no, 3, 1) = "2" Then
    
        For lRow = 1 To SS3.MaxRows
        
            SS3.Row = lRow
            SS3.Col = 1
            
            If cbo_heat_mana_no.Text = SS3.Text Then
                SS3.Col = 21
                dHeatSeq_Str = SS3.Value
                sHeatNo_Str = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeatSeq_Str)
            End If
            
            SS3.Col = 1
            If cbo_heat_no2.Text = SS3.Text Then
                SS3.Col = 21
                dHeatSeq_Stp = SS3.Value
                sHeatNo_Stp = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeatSeq_Stp)
            End If
            
            If P_MODE = "M" Then
            
                SS3.Col = 1
                If txt_tgt_heat_no.Text = SS3.Text Then
                    SS3.Col = 21
                    dHeatSeq_Tgt = SS3.Value
                    sHeatNo_Tgt = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeatSeq_Tgt)
                End If
            
            End If
        
        Next lRow
        
    ElseIf Mid(cbo_heat_mana_no, 3, 1) = "3" Then
    
        For lRow = 1 To ss5.MaxRows
        
            ss5.Row = lRow
            ss5.Col = 1
            
            If cbo_heat_mana_no.Text = ss5.Text Then
                ss5.Col = 21
                dHeatSeq_Str = ss5.Value
                sHeatNo_Str = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeatSeq_Str)
            End If
            
            ss5.Col = 1
            If cbo_heat_no2.Text = ss5.Text Then
                ss5.Col = 21
                dHeatSeq_Stp = ss5.Value
                sHeatNo_Stp = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeatSeq_Stp)
            End If
            
            If P_MODE = "M" Then
            
                ss5.Col = 1
                If txt_tgt_heat_no.Text = ss5.Text Then
                    ss5.Col = 21
                    dHeatSeq_Tgt = ss5.Value
                    sHeatNo_Tgt = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & dHeatSeq_Tgt)
                End If
            
            End If
        
        Next lRow
    
    Else
    
        Exit Sub
       
    End If
    
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AFZ1000P ('" + P_MODE + "','" + P_TYPE + "','" + sHeatNo_Str + "', '" + sHeatNo_Stp + "','" + sHeatNo_Tgt + "' ,'" + sUserID + "' ,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    M_CN1.BeginTrans
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        errMsg = sErrMessg
        M_CN1.RollbackTrans
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
    Else
        If P_MODE = "L" Then
           Call MsgBox("作业指示已成功下达！", vbInformation, "系统提示信息")
           Call Form_Ref
        ElseIf P_MODE = "T" Then
           Call MsgBox("作业指示中的时间信息已调整成功！请及时将调整后的时间下达给二级系统！", vbInformation, "系统提示信息")
           Call Form_Ref
        End If
    End If
    
    M_CN1.CommitTrans
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub

Public Sub Mlt_Proc_Change()

    Dim sQuery      As String
    Dim sChargeNo   As String
    Dim sToChargeNo As String
    Dim sMltProc    As String
    Dim sCcmLine    As String
    Dim sCcmUpdFl   As String
    Dim sProcFl     As String
    Dim sCcmLn      As String
    Dim dWgt        As Double
    Dim iRow        As Integer
    
    errMsg = ""
    
    On Error GoTo Mlt_Proc_Change_ERROR

    Screen.MousePointer = vbHourglass
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 0
        If UCase(Trim(ss1.Text)) = "UPDATE" Then
            ss1.Col = 1
            sChargeNo = Trim(ss1.Text)
            
            ss1.Col = 7
            dWgt = Val(ss1.Text & "")
            
            ss1.Col = 9
            sMltProc = Trim(ss1.Text)
            
            If Len(sMltProc) = 0 Then
                errMsg = "请正确输入工序流程！"
                GoTo Mlt_Proc_Change_ERROR
            End If
            
            sCcmLine = Right(Trim(ss1.Text), 1)
            sCcmUpdFl = ""
            
            ss1.Col = 16
            sProcFl = Trim(ss1.Text)
            
            ss1.Col = 21
            If sCcmLine <> Right(Trim(ss1.Text), 1) Then sCcmUpdFl = "Y"
                 
            'sToChargeNo = AKN2031C.txt_heat_mana_no
            'sCcmLn = AKN2031C.txt_ccm_after
            
            If (bf_OPT_LF1 <> af_OPT_LF1) Or (bf_CHK_VD <> af_CHK_VD) Or (bf_CHK_RH <> af_CHK_RH) Then
                sProcFl = "C"
            End If
            
            Call Mlt_Proc_Exec(sChargeNo, sToChargeNo, sMltProc, sCcmLine, sCcmUpdFl, dWgt, sProcFl, sCcmLn)
            
        End If
        
    Next iRow
        
    If sErrMessg = "" Then Call MsgBox("作业指示已成功工序流程变更！", vbInformation, "系统提示信息")
    
    Screen.MousePointer = vbDefault
    Exit Sub

Mlt_Proc_Change_ERROR:
    
    Screen.MousePointer = vbDefault
    errMsg = Err.Description & sQuery & errMsg
    
End Sub

Public Sub Mlt_Proc_Exec(sChargeNo As String, sToChargeNo As String, sMltProc As String, sCcmLine As String, sCcmUpdFl As String, dWgt As Double, sProcFle As String, sCcmLn As String)

    On Error GoTo Mlt_Proc_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim adoCmd As adodb.Command
    Dim YN_LF As String
    Dim YN_VD As String
    Dim YN_RH As String
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
     
'    If Chg_Lf <> AKN2031C.opt_lf(1).VALUE Then
'       YN_LF = "Y"
'    Else
'       YN_LF = "N"
'    End If
'
'    If AKN2031C.chk_VD.VALUE = True Then
'       YN_VD = "Y"
'    Else
'       YN_VD = "N"
'    End If
'
'    If AKN2031C.chk_RH.VALUE = True Then
'       YN_RH = "Y"
'    Else
'       YN_RH = "N"
'    End If
        
    If Trim(sProcFle) = "B" Then
        sQuery = "{call AFZ1063P ('" & sChargeNo & "', '" & sToChargeNo & "', '" & sCcmLn & "','" & sMltProc & "',?)}"
    ElseIf Trim(sProcFle) = "C" Then
        sQuery = "{call AFZ1065P ('" & sChargeNo & "', '" & sMltProc & "', '" & YN_LF & "', '" & YN_VD & "', '" & YN_RH & "',?)}"
    Else
        sQuery = "{call AFZ1064P ('" & sChargeNo & "', '" & sToChargeNo & "', '" & sCcmLn & "','" & sMltProc & "',?)}"
    End If
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    M_CN1.BeginTrans
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        errMsg = sErrMessg
        M_CN1.RollbackTrans
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    M_CN1.CommitTrans
    Set adoCmd = Nothing
    Exit Sub

Mlt_Proc_Exec_ERROR:

    Set adoCmd = Nothing
    errMsg = Err.Description & sQuery
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(9).Enabled = False                  'Row Cancel
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Function Line_Status_Chk() As Boolean

    On Error GoTo Line_Status_Chk_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg, errMsg As String
    Dim sQuery As String
    Dim adoCmd As adodb.Command
    
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adInteger
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1
    
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
     
    sQuery = "{call AEC2010C.P_LINE_CHECK (?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Line_Status_Chk = False
        Exit Function
        
    End If
    
    Set adoCmd = Nothing
    Line_Status_Chk = True
    Exit Function

Line_Status_Chk_ERROR:

    Set adoCmd = Nothing
    sErrMessg = Err.Description & sQuery
    Err.Raise Err.Number, Err.Description & sQuery
    Line_Status_Chk = False
    
End Function

