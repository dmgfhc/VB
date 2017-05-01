VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKN2030C 
   BackColor       =   &H00C0C0C0&
   Caption         =   "炼钢作业指示调整及下达_AKN2030C"
   ClientHeight    =   10035
   ClientLeft      =   420
   ClientTop       =   1500
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_prc_line2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AKN2030C.frx":0000
      Left            =   2460
      List            =   "AKN2030C.frx":0002
      TabIndex        =   40
      Tag             =   "炉座号"
      Top             =   135
      Width           =   600
   End
   Begin VB.ComboBox cbo_prc_line1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AKN2030C.frx":0004
      Left            =   1875
      List            =   "AKN2030C.frx":0006
      TabIndex        =   28
      Tag             =   "炉座号"
      Top             =   135
      Width           =   600
   End
   Begin Threed.SSFrame Frame2 
      Height          =   465
      Left            =   75
      TabIndex        =   16
      Top             =   540
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin Threed.SSOption Opt_InqBof 
         Height          =   285
         Left            =   225
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
         Left            =   1650
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
      Left            =   3225
      MaxLength       =   2
      TabIndex        =   15
      Tag             =   "工厂"
      Top             =   30
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
      Left            =   2745
      MaxLength       =   2
      TabIndex        =   14
      Tag             =   "工厂"
      Top             =   30
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox cbo_prc_line 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AKN2030C.frx":0008
      Left            =   1290
      List            =   "AKN2030C.frx":000A
      TabIndex        =   13
      Tag             =   "炉座号"
      Top             =   135
      Width           =   600
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
      Left            =   16590
      MaxLength       =   2
      TabIndex        =   12
      Tag             =   "工厂"
      Top             =   345
      Visible         =   0   'False
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
      Left            =   17055
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   11
      Tag             =   "工厂"
      Top             =   345
      Visible         =   0   'False
      Width           =   1260
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
      Height          =   8235
      Left            =   45
      TabIndex        =   2
      Top             =   1080
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   14526
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AKN2030C.frx":000C
      Begin FPSpread.vaSpread ss1 
         Height          =   4080
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   5835
         _Version        =   393216
         _ExtentX        =   10292
         _ExtentY        =   7197
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2030C.frx":00DE
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   4080
         Left            =   5895
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   0
         Width           =   5430
         _Version        =   393216
         _ExtentX        =   9578
         _ExtentY        =   7197
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2030C.frx":0DF0
      End
      Begin FPSpread.vaSpread ss5 
         Height          =   4080
         Left            =   11385
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   0
         Width           =   3825
         _Version        =   393216
         _ExtentX        =   6747
         _ExtentY        =   7197
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2030C.frx":1AD6
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4095
         Left            =   0
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   4140
         Width           =   5835
         _Version        =   393216
         _ExtentX        =   10292
         _ExtentY        =   7223
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   20
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2030C.frx":27BC
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   4095
         Left            =   5895
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   4140
         Width           =   5430
         _Version        =   393216
         _ExtentX        =   9578
         _ExtentY        =   7223
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   20
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2030C.frx":32C1
      End
      Begin FPSpread.vaSpread ss6 
         Height          =   4095
         Left            =   11385
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   4140
         Width           =   3825
         _Version        =   393216
         _ExtentX        =   6747
         _ExtentY        =   7223
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   20
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2030C.frx":3DA9
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
      Left            =   15465
      Top             =   345
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
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
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   75
      Top             =   135
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "炉座号"
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
      Left            =   3120
      TabIndex        =   19
      Top             =   60
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_chg_prc_line 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "AKN2030C.frx":4891
         Left            =   1500
         List            =   "AKN2030C.frx":4893
         TabIndex        =   27
         Tag             =   "炉座号"
         Top             =   60
         Width           =   675
      End
      Begin Threed.SSOption opt_cancel 
         Height          =   285
         Left            =   2775
         TabIndex        =   20
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "指示取消"
      End
      Begin Threed.SSOption opt_mltcd_change 
         Height          =   285
         Left            =   10770
         TabIndex        =   21
         Top             =   -30
         Visible         =   0   'False
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
         Left            =   255
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
         Left            =   4485
         TabIndex        =   23
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
         Caption         =   "顺序调整"
      End
      Begin Threed.SSOption opt_del 
         Height          =   285
         Left            =   6195
         TabIndex        =   24
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
         Caption         =   "钢水返送"
      End
      Begin Threed.SSOption opt_time 
         Height          =   285
         Left            =   7905
         TabIndex        =   25
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
         Caption         =   "时间调整"
      End
      Begin Threed.SSOption opt_send 
         Height          =   285
         Left            =   9615
         TabIndex        =   26
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
         Caption         =   "指示下达"
      End
      Begin Threed.SSOption opt_rsltdel 
         Height          =   285
         Left            =   10770
         TabIndex        =   39
         Top             =   210
         Visible         =   0   'False
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
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   465
      Left            =   3120
      TabIndex        =   29
      Top             =   540
      Width           =   12135
      _ExtentX        =   21405
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   4005
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   36
         Tag             =   "终止炉号"
         Top             =   80
         Width           =   1110
      End
      Begin VB.TextBox txt_tgt_heat_no 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   6510
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   32
         Tag             =   "终止炉号"
         Top             =   80
         Width           =   1110
      End
      Begin VB.TextBox cbo_heat_mana_no 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1500
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         Tag             =   "起始炉号"
         Top             =   80
         Width           =   1110
      End
      Begin Threed.SSOption opt_from 
         Height          =   315
         Left            =   255
         TabIndex        =   30
         Top             =   75
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
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
         Caption         =   "起始炉号"
      End
      Begin Threed.SSPanel SSPrtn 
         Height          =   420
         Left            =   10680
         TabIndex        =   33
         Top             =   15
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         Height          =   420
         Left            =   9270
         TabIndex        =   34
         Top             =   15
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   16761087
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         Height          =   420
         Left            =   7860
         TabIndex        =   35
         Top             =   15
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
      Begin Threed.SSOption opt_to 
         Height          =   315
         Left            =   2775
         TabIndex        =   37
         Top             =   75
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
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
         Caption         =   "终止炉号"
      End
      Begin Threed.SSOption opt_target 
         Height          =   315
         Left            =   5280
         TabIndex        =   38
         Top             =   75
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
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
         Caption         =   "目标炉号"
      End
   End
End
Attribute VB_Name = "AKN2030C"
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
'-- Program ID        AKN2030C
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
Dim sc1 As New Collection           'Spread Collection
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

Dim errMsg   As String
Dim iSelRow  As Integer
Dim txt_AFT_Prc_line As String
Dim txt_AFT_SS_Col As Integer
Dim txt_AFT_SS_Row As Integer

Dim lHeat_Edt_Seq_Fr As Long
Dim lHeat_Edt_Seq_To As Long
Dim lHeat_Edt_Seq_Ta As Long

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
          Call Gp_Ms_Collection(Opt_InqBof, "p", " ", " ", " ", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    
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
    For iCol = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFN2031C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFN2031C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
   
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss3.MaxCols
        Call Gp_Sp_Collection(ss3, iCol, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Next iCol
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AFN2031C.P_REFER1", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=3, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss4.MaxCols
        Call Gp_Sp_Collection(ss4, iCol, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Next iCol
    
    'Spread_Collection
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:="AFN2031C.P_REFER2", Key:="P-R"
    Sc4.Add Item:=pColumn4, Key:="pColumn"
    Sc4.Add Item:=nColumn4, Key:="nColumn"
    Sc4.Add Item:=aColumn4, Key:="aColumn"
    Sc4.Add Item:=mColumn4, Key:="mColumn"
    Sc4.Add Item:=iColumn4, Key:="iColumn"
    Sc4.Add Item:=lColumn4, Key:="lColumn"
    Sc4.Add Item:=1, Key:="First"
    Sc4.Add Item:=ss4.MaxCols, Key:="Last"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss5.MaxCols
        Call Gp_Sp_Collection(ss5, iCol, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Next iCol
    
    'Spread_Collection
    Sc5.Add Item:=ss5, Key:="Spread"
    Sc5.Add Item:="AFN2031C.P_REFER1", Key:="P-R"
    Sc5.Add Item:=pColumn5, Key:="pColumn"
    Sc5.Add Item:=nColumn5, Key:="nColumn"
    Sc5.Add Item:=aColumn5, Key:="aColumn"
    Sc5.Add Item:=mColumn5, Key:="mColumn"
    Sc5.Add Item:=iColumn5, Key:="iColumn"
    Sc5.Add Item:=lColumn5, Key:="lColumn"
    Sc5.Add Item:=3, Key:="First"
    Sc5.Add Item:=ss5.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss6.MaxCols
        Call Gp_Sp_Collection(ss6, iCol, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Next iCol
    
    'Spread_Collection
    Sc6.Add Item:=ss6, Key:="Spread"
    Sc6.Add Item:="AFN2031C.P_REFER2", Key:="P-R"
    Sc6.Add Item:=pColumn6, Key:="pColumn"
    Sc6.Add Item:=nColumn6, Key:="nColumn"
    Sc6.Add Item:=aColumn6, Key:="aColumn"
    Sc6.Add Item:=mColumn6, Key:="mColumn"
    Sc6.Add Item:=iColumn6, Key:="iColumn"
    Sc6.Add Item:=lColumn6, Key:="lColumn"
    Sc6.Add Item:=1, Key:="First"
    Sc6.Add Item:=ss6.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    Proc_Sc.Add Item:=Sc4, Key:="Sc4"
    Proc_Sc.Add Item:=Sc5, Key:="Sc5"
    Proc_Sc.Add Item:=Sc6, Key:="Sc6"

    Me.KeyPreview = True
    Me.Opt_InqBof.BackColor = &HE0E0E0
    Me.Opt_InqCcm.BackColor = &HE0E0E0
    Me.BackColor = &HE0E0E0
    

    Call Gp_Sp_ColHidden(ss1, 22, True)    'Heat_edt_seq
    
    Call Gp_Sp_ColHidden(ss2, 1, True)     'SEQ_NO
    Call Gp_Sp_ColHidden(ss2, 18, True)    'STATUS
    
    'Call Gp_Sp_ColHidden(ss3, 16, True)   'l2_send y/n
    Call Gp_Sp_ColHidden(ss3, 22, True)    'Heat_edt_seq
    
    Call Gp_Sp_ColHidden(ss4, 1, True)     'SEQ_NO
    Call Gp_Sp_ColHidden(ss4, 18, True)    'STATUS
    
    Call Gp_Sp_ColHidden(ss5, 22, True)    'Heat_edt_seq
    
    Call Gp_Sp_ColHidden(ss6, 1, True)     'SEQ_NO
    Call Gp_Sp_ColHidden(ss6, 18, True)    'STATUS
    
'    Call Gp_Sp_ColHidden(ss1, 12, True)   '序列号
'    Call Gp_Sp_ColHidden(ss3, 12, True)   '序列号
'    Call Gp_Sp_ColHidden(ss1, 13, True)   '钢种组
'    Call Gp_Sp_ColHidden(ss3, 13, True)   '钢种组

End Sub

Private Sub cbo_heat_mana_no_Change()

    Dim iRow As Integer
    Dim sColor As String
    
    With ss1
        .Col = 1
        For iRow = 1 To .MaxRows
            .Row = iRow
            If .BackColor = &HFF Then
               .Col = 3
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

    Dim sQuery As String
    Dim bDyanmic_start As Boolean
    Dim Dynamic_Slab As String

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    'Check Dynamic Slab Cutting Job
    Dynamic_Slab = "SC1"
    sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
    
    If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
    
        Dynamic_Slab = "SC2"
        sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
    
        If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
        
            Dynamic_Slab = "SC3"
            sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
    
            If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
                bDyanmic_start = True
            Else
                bDyanmic_start = False
            End If
        
        Else
            bDyanmic_start = False
        End If
    
    Else
        bDyanmic_start = False
    End If
    
'    Call Dynamic_Slab_ScreenSet(bDyanmic_start)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    Call Gp_Ms_NeceColor(Mc4("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc4.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc5.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc6.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc4.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc5.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc6.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(Sc4)
    Call Gf_Sp_Cls(Sc5)
    Call Gf_Sp_Cls(Sc6)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc4.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc5.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc6.Item("Spread"), "K-System.INI", Me.Name)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "K-System.INI", Me.Name)
    
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
    
    txt_prc_line.Text = "2"
    txt_proc_fl.Text = ""
    
    Ref_FL = True

    Call Form_Ref
    
'20140122
   If sUserID = "1BY1002" Or sUserID = "1BY1003" Or sUserID = "1BY1004" Or sUserID = "1BY1005" Then
          opt_del.Enabled = False
   End If
'20140122
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "K-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc4.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc5.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc6.Item("Spread"), "K-System.INI", Me.Name)
    
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
    Set sc1 = Nothing
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
    
        If Gf_Sp_Cls(sc1) Then
        
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
            
            SSPsend.Visible = False
            SSPpdt.Visible = False
            SSPrtn.Visible = False
            
            cbo_chg_prc_line.Text = ""
            
            opt_line_change.Value = False
            opt_mltcd_change.Value = False
            opt_cancel.Value = False
            opt_change.Value = False
            opt_del.Value = False
            opt_rsltdel.Value = False
            opt_time.Value = False
            opt_send.Value = False
            
            opt_line_change.BackColor = &HE0E0E0
            opt_mltcd_change.BackColor = &HE0E0E0
            opt_cancel.BackColor = &HE0E0E0
            opt_change.BackColor = &HE0E0E0
            opt_del.BackColor = &HE0E0E0
            opt_rsltdel.BackColor = &HE0E0E0
            opt_time.BackColor = &HE0E0E0
            opt_send.BackColor = &HE0E0E0
            
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
    
    PGM_ID = "AKN2030C"
    cbo_chg_prc_line.Text = ""
    cbo_heat_mana_no.Text = ""
    cbo_heat_no2.Text = ""
    txt_tgt_heat_no.Text = ""
    Ref_FL = "0"
    
    opt_line_change.Value = False
    opt_mltcd_change.Value = False
    opt_cancel.Value = False
    opt_change.Value = False
    opt_del.Value = False
    opt_time.Value = False
    opt_send.Value = False
    opt_rsltdel.Value = False
    
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
    
'    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
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
        Call Gp_Sp_EvenRowBackcolor(sc1.Item("Spread"))
        Call Gp_Sp_EvenRowBackcolor(Sc3.Item("Spread"))
        Call Gp_Sp_EvenRowBackcolor(Sc5.Item("Spread"))
        
        Call Gf_Sp_Cls(sc2)
        Call Gf_Sp_Cls(Sc4)
        Call Gf_Sp_Cls(Sc6)
        
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(9).Enabled = False
        
        ss1.OperationMode = OperationModeNormal
        ss3.OperationMode = OperationModeNormal
        ss5.OperationMode = OperationModeNormal
        
        SSPsend.Visible = True
        SSPpdt.Visible = True
        SSPrtn.Visible = True
        
'        sQuery = "select upd from zp_authority where EMP_ID = '" + sUserID + "' and pgmid = '" + PGM_ID + "'"
'20140122
        sQuery = "SELECT MAX(T) FROM (SELECT MAX(Z.UPD) T FROM ZP_AUTHORITY Z, ZP_EMPGRP P WHERE Z.EMP_ID = P.GROUP_ID AND Z.PGMID = '" + PGM_ID + "' AND P.EMP_ID = '" + sUserID + "' UNION SELECT DECODE((SELECT MIN(GROUP_ID) FROM ZP_EMPGRP WHERE EMP_ID = '" + sUserID + "'),'000000','1', '0') T FROM DUAL)"
'20140122
        sAut = Gf_FloatFind(M_CN1, sQuery)
        
        If sAut = "1" Then
            Frame1.Enabled = True
        Else
            Frame1.Enabled = False
        End If

        Call Spread_Color_Setting(ss1)
        Call Spread_Color_Setting(ss3)
        Call Spread_Color_Setting(ss5)
       
        If opt_line_change Or opt_mltcd_change Then
        
            With ss1

                For iRow = iSelRow To .MaxRows
                    .Row = iRow
                    .Col = 19
                    If .Text <> "Y" Then
                    
                        If opt_line_change Then
                            .BlockMode = True
                            .Col = 2:    .Col2 = 2
                            .Row = iRow: .Row2 = iRow
                            .BackColor = &HC0FFEE
                            .Lock = False
                            .BlockMode = False
                        End If
                        
                        If opt_mltcd_change Then
                            .BlockMode = True
                            .Col = 9:    .Col2 = 9
                            .Row = iRow: .Row2 = iRow
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
    
        If oSpr.Name = "ss2" Or oSpr.Name = "ss4" Or oSpr.Name = "ss6" Then
        
            For iRow = 1 To .MaxRows
            
                .Row = iRow:  .Col = 18
                
                If .Text = "B" Then
                
                    For iCol = 1 To .MaxCols
                        .Col = iCol
                        .BackColor = SSPpdt.BackColor
                    Next iCol
                    
                ElseIf .Text = "Y" Then
                
                    For iCol = 1 To .MaxCols
                        .Col = iCol
                        .BackColor = SSPsend.BackColor
                    Next iCol
                    
                End If
                
            Next iRow
            
            Exit Sub
        
        End If
        
        .Col = 1
        
        For iRow = 1 To .MaxRows
            .Row = iRow:  .Col = 19
            
            If .Text = "Y" Then
                For iCol = 1 To .MaxCols
                    .Col = iCol
                    .BackColor = SSPsend.BackColor
                Next iCol
            End If
            
        Next iRow
        
        For iRow = 1 To .MaxRows
            .Row = iRow:  .Col = 24
            
            If .Text = "Y" Then
                For iCol = 1 To .MaxCols
                    .Col = iCol
                    .ForeColor = BLUE
                Next iCol
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
        
'        For iRow = 1 To .MaxRows
'            .Row = iRow:  .Col = 21
'
'            If .Text = "Y" Then
'                For iCol = 1 To .MaxCols
'                    .Col = iCol
'                    .ForeColor = &HFF
'
'                Next iCol
'            End If
'
'        Next iRow
        
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
            .Col = 19:    sPrcLine = .Text
            .Col = 20:    sMltProc = .Text
            
            .Col = 2:    .Text = sPrcLine
            .Col = 8:    .Text = sMltProc
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
    Dim sHeat_Mana_No As String
    Dim iCnt, i, j As Integer
    Dim sQuery As String
    Dim OutParam(1, 4) As Variant

    If ss1.MaxRows = 0 Then Exit Sub
    
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
    
    If P_MODE <> "L" Then
        If cbo_heat_mana_no.Text <> "" Then
        
            sHeat_Mana_No = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & lHeat_Edt_Seq_Fr)
            
            If cbo_heat_mana_no.Text <> sHeat_Mana_No Then
                errMsg = "炉号已变更，请重新查询界面后再操作..!!!"
            End If
            
        End If
    
    End If
    
    If cbo_heat_no2.Text <> "" Then
    
        sHeat_Mana_No = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & lHeat_Edt_Seq_To)
        
        If cbo_heat_no2.Text <> sHeat_Mana_No Then
            errMsg = "炉号已变更，请重新查询界面后再操作..!!!"
        End If
        
    End If
    
    If txt_tgt_heat_no.Text <> "" Then
    
        sHeat_Mana_No = Gf_CodeFind(M_CN1, "SELECT HEAT_MANA_NO FROM EP_CHARGE_IDX WHERE HEAT_EDT_SEQ = " & lHeat_Edt_Seq_Ta)
        
        If txt_tgt_heat_no.Text <> sHeat_Mana_No Then
            errMsg = "炉号已变更，请重新查询界面后再操作..!!!"
        End If
        
    End If
    
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
    If P_MODE <> "P" Then
        If Not Line_Status_Chk Then Exit Sub
    End If
    '-----------------------------------------------------------------------
    
    Select Case P_MODE
    
            Case "X"
            
                If (bf_CHK_VD <> af_CHK_VD) Or (bf_CHK_RH <> af_CHK_RH) Then
                   
                   sQuery = "{call AFN2031C.P_REFER4('" & Trim(AKN2031C.txt_heat_no_fr) & "',?)}"
                
                    'Ado Setting
                    M_CN1.CursorLocation = adUseServer
                    Set adoCmd = New ADODB.Command
                    
                    adoCmd.CommandType = adCmdText
                    Set adoCmd.ActiveConnection = M_CN1
                    OutParam(1, 1) = "arg_e_msg"
                    OutParam(1, 2) = adVarChar
                    OutParam(1, 3) = adParamOutput
                    OutParam(1, 4) = 256
                    
                    adoCmd.CommandText = sQuery
                    
                    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
                    
                    adoCmd.Execute , , adExecuteNoRecords
                    
                    'Process Error Check
                    If adoCmd("arg_e_msg") = "1" Or adoCmd("arg_e_msg") = "2" Then
                        If Not Gf_MessConfirm("可能会影响质量,您确定换吗?", , "系统提示信息") Then
                            Set adoCmd = Nothing
                            Exit Sub
                        End If
                    End If
                    
                    Set adoCmd = Nothing
                    
                End If
                
                Call Mlt_Proc_Change
                
            Case "U"
            
                With ss1
                     For i = 1 To .MaxRows
                         .Col = 2
                         .Row = i
                         If .Text <> "1" And .Text <> "2" Then
                             MsgBox "请正确输入炉座号！", vbInformation, "系统提示信息"
                             Exit Sub
                         End If
                     Next i
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
                
           Case "B"
                
                If Len(Trim(cbo_heat_mana_no.Text)) <> 8 Then
                    MsgBox "请选择您要返送的起始炉号！", vbCritical, "系统提示信息"
                    Exit Sub
                End If
                
                If Gf_MessConfirm("确定要从炉号 <" + cbo_heat_mana_no.Text + "> 开始返送作业指示吗？", "W", "系统提示信息确认") Then
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
                     For i = 1 To .MaxRows
                         .Col = 1
                         .Row = i
                         If .Text >= cbo_heat_mana_no.Text And .Text <= cbo_heat_no2.Text Then
                             .Col = 4
                             If .BackColor = SSPrtn.BackColor Then
                                 Ret_Steel = "Y"
                             ElseIf .BackColor <> SSPrtn.BackColor Then
                                 Not_Ret = "Y"
                             End If
                         ElseIf .Text > cbo_heat_no2.Text Then
                             Exit For
                         End If
                     Next i
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

Public Function sf_Sp_ProceExist() As Integer

    Dim iRow        As Integer
    Dim sColor      As String
    
    sf_Sp_ProceExist = 0
    
    With ss1
    
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 0
            If Trim(.Text) = "Input" Or Trim(.Text) = "Update" Or Trim(.Text) = "Delete" Then
                sf_Sp_ProceExist = 1
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

Private Sub opt_line_change_Click(Value As Integer)

    Dim iRow        As Integer
    
    If sf_Sp_ProceExist() > 0 Then Call Form_Ref:  opt_line_change.Value = True
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
                .Row = iRow: .Row2 = iRow
                .BackColor = &HC0FFEE
                .Lock = False
                .BlockMode = False
            End If
        Next
        
    End With
    
End Sub

Private Sub opt_mltcd_change_Click(Value As Integer)

    Dim iRow        As Integer
    
    If sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_mltcd_change.Value = True
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
                .Row = iRow: .Row2 = iRow
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
    
    If sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_cancel.Value = True
    
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

    If sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_change.Value = True
    
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

    ss1.BlockMode = True
    ss1.Col = -1
    ss1.Row = -1
    ss1.Lock = True
    ss1.BlockMode = False
    
End Sub

Private Sub opt_del_Click(Value As Integer)

    If sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_del.Value = True

    P_MODE = "B"
    
    opt_from.Enabled = True
    opt_to.Enabled = False
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
    
    ss1.BlockMode = True
    ss1.Col = -1
    ss1.Row = -1
    ss1.Lock = True
    ss1.BlockMode = False
    
End Sub

Private Sub opt_rsltdel_Click(Value As Integer)

    Dim iRow       As Integer
    Dim sColor, TT As String
    Dim sQuery As String
    Dim Dynamic_Slab As String

    'Check Dynamic Slab Cutting Job
    Dynamic_Slab = "SC1"
    sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
    
    If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
    
        Dynamic_Slab = "SC2"
        sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
    
        If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
        
            Dynamic_Slab = "SC3"
            sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
    
            If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
            
                MsgBox "Restart this Screen..!!", vbCritical, "系统提示信息"
                opt_rsltdel.Value = False
                Exit Sub
            
            End If
        
        End If
    
    End If
    
    If sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_rsltdel.Value = True
    
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
    
    If sf_Sp_ProceExist() > 0 Then Call Form_Ref: opt_send.Value = True
    
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
    
    Dim i As Integer
    Dim iRow As Integer


    Set Active_Spread = Me.ss1
    If Row <= 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 1
    txt_heat_mana_no.Text = ss1.Text
    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)
    
    sChgPrcLine = cbo_chg_prc_line.Text


    
    If opt_mltcd_change And Col = 8 Then
    
        ss1.Row = Row
        ss1.Col = 19
        sL2SendFL = ss1.Text
        
        ss1.Col = 16
        txt_proc_fl.Text = ss1.Text
                
        If sL2SendFL = "Y" Or txt_proc_fl.Text = "B" Then Exit Sub
        
        AKN2031C.txt_heat_no.Text = txt_heat_mana_no.Text
        AKN2031C.txt_heat_no_fr.Text = txt_heat_mana_no.Text
        AKN2031C.txt_heat_no_to.Text = txt_heat_mana_no.Text
        
        ss1.Col = 2
        AKN2031C.txt_bof_proc.Text = ss1.Text

        ss1.Col = 5
        AKN2031C.txt_Stlgrd.Text = ss1.Text
        
        ss1.Col = 9
        AKN2031C.txt_mlt_prc_cd.Text = ss1.Text
        
        Load AKN2031C
        AKN2031C.Show 1
        
'        AKN2031C.ZOrder (0)
        
        Exit Sub
        
    End If

    lBlkrow1 = Row
    lBlkrow2 = Row
    
    If ss1.MaxRows < 1 Then
    
        Call Gf_Sp_Cls(sc2)
        Exit Sub
    
    End If
    
    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub

    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    ss2.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss2)
    Call Spread_Color_Setting(ss2)
    
     For iRow = 1 To ss2.MaxRows
    
               ss2.Row = iRow
               ss2.Col = 3
                If ss2.Text = "Y" Then
                  For i = 1 To ss2.MaxCols
                       ss2.Col = i
                       ss2.ForeColor = &HC000&
                  Next
                End If
                
                
               ss2.Row = iRow
               ss2.Col = 20
                If ss2.Text = "Y" Then
                  For i = 1 To ss2.MaxCols
                       ss2.Col = i
                       ss2.ForeColor = &HFF&
                  Next
                End If

      
     Next iRow
    

    With ss1
        For iRow1 = .ActiveRow To .MaxRows
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            With ss2
              .Col = 4
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
              .Col = 4
              Next iRow2
            End With

        Next iRow1
    
    End With
    
End Sub

Private Sub ss1_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss1.MaxCols
        ss3.ColWidth(iCol) = ss1.ColWidth(iCol)
        ss5.ColWidth(iCol) = ss1.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim iRow, iCnt As Integer
    Dim sColor, M_TEMP As String
    
    If opt_cancel.Value = False And opt_change.Value = False _
                                And opt_del.Value = False _
                                And opt_send.Value = False _
                                And opt_line_change = False _
                                And opt_rsltdel = False Then
        Exit Sub
    End If
            
    ss1.Row = Row
    ss1.Col = 1
    txt_heat_mana_no.Text = ss1.Text
    
    If opt_line_change And Col = 2 Then
    
        'Line Status CHECK (added by KIM.SUNG.HO 2010.03.20)
        If Not Line_Status_Chk Then Exit Sub
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
           
           .Col = 22
           .Row = .ActiveRow
           
           lHeat_Edt_Seq_Fr = .Text
           
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
                
           ElseIf P_MODE = "B" Then
           
                For iRow = 1 To .MaxRows
            
                   .Col = 4
                   .Row = iRow
                    sColor = .BackColor
                   .Col = 2: .BackColor = sColor
                   
                   .Col = 1
                   
                    If .Text < cbo_heat_mana_no.Text Then
                    
                       .BackColor = sColor
                       
                    ElseIf .Text >= cbo_heat_mana_no.Text Then
                        
                        If iSelRow >= iRow And iSelRow <> 1 Then
                            .Row = iSelRow
                            .Col = 1
                            MsgBox "炉号(" & .Text & ")正常生产中，不能返送！", vbInformation, "系统提示信息"
                            cbo_heat_mana_no.Text = ""
                            Exit Sub
                        End If
                        
                        If .BackColor = SSPsend.BackColor Then
                            MsgBox "请先从炉号 <" + cbo_heat_mana_no.Text + "> 开始取消已下达到二级的作业指示，然后再返送！", vbCritical, "系统提示信息"
                            opt_cancel.Value = True
                            Exit Sub
                        End If
                        
                        .BackColor = &HFF&
                    
                    End If
                
                Next
                
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
        
           cbo_heat_no2.Text = .Text
           
           .Col = 22
           .Row = .ActiveRow
           
           lHeat_Edt_Seq_To = .Text
           
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
                
           End If
           
           If P_MODE = "M" Then
                opt_target.Value = True
           End If
           
        ElseIf opt_target.Value = True Then
        
            txt_tgt_heat_no.Text = .Text
            
            .Col = 22
            .Row = .ActiveRow
            
           lHeat_Edt_Seq_Ta = .Text
           
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
    
    Set Active_Spread = Me.ss2
    
End Sub

Private Sub ss2_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss2.MaxCols
        ss4.ColWidth(iCol) = ss2.ColWidth(iCol)
        ss6.ColWidth(iCol) = ss2.ColWidth(iCol)
    Next iCol

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
        Set Active_Spread = Me.ss2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor, sHeat, sTemp As String
    Dim sChgPrcLine          As String
    
    Dim i As Integer
    Dim iRow As Integer


    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss3
    If Row <= 0 Then Exit Sub
    
    If ss3.MaxRows < 1 Then
        Call Gf_Sp_Cls(Sc4)
        Exit Sub
    End If
    
    ss3.Row = Row
    ss3.Col = 1
    txt_heat_mana_no.Text = ss3.Text
    
    Call Gf_Sp_Refer(M_CN1, Sc4, Mc2, , , False)
    ss4.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss4)
    Call Spread_Color_Setting(ss4)
    
    
    For iRow = 1 To ss4.MaxRows
    
               ss4.Row = iRow
               ss4.Col = 3
                If ss4.Text = "Y" Then
                  For i = 1 To ss4.MaxCols
                       ss4.Col = i
                       ss4.ForeColor = &HC000&
                  Next
                End If
                
               ss4.Row = iRow
               ss4.Col = 20
                If ss4.Text = "Y" Then
                  For i = 1 To ss4.MaxCols
                       ss4.Col = i
                       ss4.ForeColor = &HFF&
                  Next
                End If
                
        
     Next iRow

    
    
   
    With ss3
    
        For iRow1 = .ActiveRow To .MaxRows
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            
            With ss4
            
                .Col = 4
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
                    .Col = 4
                Next iRow2
                
            End With

        Next iRow1
        
    End With
    
    
     

    
End Sub

Private Sub ss3_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss3.MaxCols
        ss1.ColWidth(iCol) = ss3.ColWidth(iCol)
        ss5.ColWidth(iCol) = ss3.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss4
    
End Sub

Private Sub ss4_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss4.MaxCols
        ss2.ColWidth(iCol) = ss4.ColWidth(iCol)
        ss6.ColWidth(iCol) = ss4.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss5_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss5.MaxCols
        ss1.ColWidth(iCol) = ss5.ColWidth(iCol)
        ss3.ColWidth(iCol) = ss5.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss6_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss6
    
End Sub

Private Sub ss5_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor, sHeat, sTemp As String
    Dim sChgPrcLine          As String
    
    Dim i As Integer
    Dim iRow As Integer

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
    
    Call Gf_Sp_Refer(M_CN1, Sc6, Mc2, , , False)
    ss6.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss6)
    Call Spread_Color_Setting(ss6)
    
    For iRow = 1 To ss6.MaxRows
    
               ss6.Row = iRow
               ss6.Col = 3
                If ss6.Text = "Y" Then
                  For i = 1 To ss6.MaxCols
                       ss6.Col = i
                       ss6.ForeColor = &HC000&
                  Next
                End If
                
               ss6.Row = iRow
               ss6.Col = 20
                If ss6.Text = "Y" Then
                  For i = 1 To ss6.MaxCols
                       ss6.Col = i
                       ss6.ForeColor = &HFF&
                  Next
                End If
     Next iRow
    

    With ss5
    
        For iRow1 = .ActiveRow To .MaxRows
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            
            With ss6
            
                .Col = 4
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
                    .Col = 4
                Next iRow2
                
            End With

        Next iRow1
        
    End With

End Sub

Private Sub ss6_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss6.MaxCols
        ss2.ColWidth(iCol) = ss6.ColWidth(iCol)
        ss4.ColWidth(iCol) = ss6.ColWidth(iCol)
    Next iCol

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
    Dim adoCmd            As ADODB.Command
    
    P_TYPE = "S"
    
    Set AdoRs = New ADODB.Recordset
        
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
     
     sQuery = "{call AFZ1000P ('" + P_MODE + "','" + P_TYPE + "','" & sHeatEdtSeq & "', '" + cbo_chg_prc_line.Text + "' ,'','" + sUserID + "' ,?)}"
    'sQuery = "{call AFZ9110P ('" + sHeatEdtSeq + "','" + cbo_chg_prc_line.Text + "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
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
    If Trim(ret_Result_ErrMsg) = "" Then Call Form_Ref

    Screen.MousePointer = vbDefault
    Exit Sub
    
End Sub

Public Sub ChageNo_Search(lSeq As Long, sChargeNo As String, sPlcLine As String, sStrChrNo1 As String, sStrChrNo2 As String)
    
    Dim sQuery      As String
    Dim sChrNoOrg   As String
    Dim sChgNoNew   As String
    
    'NOT USED
    Exit Sub
    
    
    Set AdoRs = New ADODB.Recordset
        
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
        ss1.Col = 16
        sProdFl = Trim(ss1.Text)
        
        ss1.Col = 19
        sL2Send = Trim(ss1.Text)
        
        If sL2Send = "Y" Or sProdFl >= "B" Then
            ss1.Col = 14
            dChkSeq = Val(ss1.Value & "")
            Exit For
        End If
    Next iRow
    
    Set AdoRs = New ADODB.Recordset
    
    If dChkSeq = 0 Then
        sQuery = "         SELECT  MAX(CHG_SEQ)    " & vbCrLf
        sQuery = sQuery & "  FROM  EP_CHARGE_IDX   " & vbCrLf
        ss1.Row = 1:   ss1.Col = 13
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
    
    Set AdoRs = New ADODB.Recordset
    
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
    Dim adoCmd As ADODB.Command
    P_TYPE = "S"
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
     
    sQuery = "{call AFZ1000P ('" + P_MODE + "','" + P_TYPE + "','" + cbo_heat_mana_no + "', '" + cbo_heat_no2 + "','" + txt_tgt_heat_no + "' ,'" + sUserID + "' ,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
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
                 
            sToChargeNo = AKN2031C.txt_heat_mana_no.Text
            sCcmLn = AKN2031C.txt_ccm_after
            
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
    Dim adoCmd As ADODB.Command
    Dim ret_Result_ErrMsg As String
    Dim YN_LF As String
    Dim YN_VD As String
    Dim YN_RH As String
    Dim sQuery As String
    Dim Current_Ccm As String
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
     
    If Chg_Lf <> AKN2031C.opt_lf(1).Value Then
       YN_LF = "Y"
    Else
       YN_LF = "N"
    End If
    
    If AKN2031C.chk_VD.Value = True Then
       YN_VD = "Y"
    Else
       YN_VD = "N"
    End If
    
    If AKN2031C.chk_RH.Value = True Then
       YN_RH = "Y"
    Else
       YN_RH = "N"
    End If
    
        
    If Trim(sProcFle) = "B" Then
        
        sQuery = "{call AFZ1063P ('" & sChargeNo & "', '" & sToChargeNo & "', '" & sCcmLn & "','" & sMltProc & "',?)}"
        
    ElseIf Trim(sProcFle) = "C" Then
    
        sQuery = "{call AFZ1065P ('" & sChargeNo & "', '" & sMltProc & "', '" & YN_LF & "', '" & YN_VD & "', '" & YN_RH & "',?)}"
    
    Else
    
        sQuery = "{call AFZ1064P ('" & sChargeNo & "', '" & sToChargeNo & "', '" & sCcmLn & "','" & sMltProc & "',?)}"
    
    End If
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
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
    Dim adoCmd As ADODB.Command
    
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adInteger
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1
    
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
     
    sQuery = "{call AFN2031C.P_LINE_CHECK (?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
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
'
'Private Sub Dynamic_Slab_ScreenSet(bStart As Boolean)
'
'    If bStart Then
'
'        opt_rsltdel.Visible = False
'
'        Frame1.Width = 10275
'        Frame1.Left = 4890
'
'        opt_line_change.Left = 255
'        cbo_chg_prc_line.Left = 1470
'        opt_mltcd_change.Left = 2400
'        opt_cancel.Left = 4260
'        opt_change.Left = 5280
'        opt_del.Left = 6765
'        opt_time.Left = 7815
'        opt_send.Left = 9270
'
'    Else
'
'        opt_rsltdel.Visible = True
'
'        Frame1.Width = 10815
'        Frame1.Left = 4350
'
'        opt_line_change.Left = 255
'        cbo_chg_prc_line.Left = 1470
'        opt_mltcd_change.Left = 2310
'        opt_cancel.Left = 4080
'        opt_change.Left = 4965
'        opt_del.Left = 6315
'        opt_rsltdel.Left = 7260
'        opt_time.Left = 8625
'        opt_send.Left = 9945
'
'    End If
'
'    Me.Refresh
'
'End Sub
