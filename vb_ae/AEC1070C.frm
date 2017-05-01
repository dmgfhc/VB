VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AEC1070C 
   Caption         =   "连浇炉数编制结果修改_AEC1070C"
   ClientHeight    =   9255
   ClientLeft      =   495
   ClientTop       =   2115
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15315
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_prc_line3 
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
      Left            =   4770
      MaxLength       =   2
      TabIndex        =   47
      Tag             =   "工厂"
      Top             =   345
      Visible         =   0   'False
      Width           =   465
   End
   Begin SSSplitter.SSSplitter SSSplitter2 
      Height          =   7785
      Left            =   60
      TabIndex        =   44
      Top             =   1440
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   13732
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AEC1070C.frx":0000
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   6480
         TabIndex        =   55
         Top             =   5760
         Visible         =   0   'False
         Width           =   270
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   4680
         Left            =   0
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   0
         Width           =   15225
         _Version        =   393216
         _ExtentX        =   26855
         _ExtentY        =   8255
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
         MaxCols         =   17
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEC1070C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3045
         Left            =   0
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   4740
         Width           =   15225
         _Version        =   393216
         _ExtentX        =   26855
         _ExtentY        =   5371
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
         MaxCols         =   18
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AEC1070C.frx":0A5A
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   480
      Left            =   435
      TabIndex        =   38
      Top             =   930
      Width           =   1710
      Begin VB.ComboBox cbo_chg_prc_line 
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
         ItemData        =   "AEC1070C.frx":146C
         Left            =   3240
         List            =   "AEC1070C.frx":146E
         TabIndex        =   48
         Tag             =   "炉/机号"
         Top             =   120
         Visible         =   0   'False
         Width           =   675
      End
      Begin Threed.SSOption opt_line_change 
         Height          =   255
         Left            =   1740
         TabIndex        =   42
         Top             =   150
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "炉座号变更"
      End
      Begin Threed.SSOption opt_mltcd_change 
         Height          =   255
         Left            =   150
         TabIndex        =   43
         Top             =   150
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "工序流程变更"
      End
   End
   Begin VB.TextBox txt_prc_line1 
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
      Left            =   4095
      MaxLength       =   2
      TabIndex        =   36
      Tag             =   "工厂"
      Text            =   "1"
      Top             =   345
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox cbo_prc_line 
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
      ItemData        =   "AEC1070C.frx":1470
      Left            =   3510
      List            =   "AEC1070C.frx":1472
      TabIndex        =   35
      Tag             =   "炉/机号"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txt_prc_line2 
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
      Left            =   4350
      MaxLength       =   2
      TabIndex        =   34
      Tag             =   "工厂"
      Top             =   345
      Visible         =   0   'False
      Width           =   465
   End
   Begin Threed.SSCommand cmd_cast1 
      Height          =   420
      Left            =   11175
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   30
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   128
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
      Caption         =   "连浇炉数编制"
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmd_roll1 
      Height          =   420
      Left            =   11140
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   990
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16576
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
      Caption         =   "编制轧辊单位"
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmd_confirm 
      Height          =   420
      Left            =   13905
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   990
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16576
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
      Caption         =   "指示确定"
      BevelWidth      =   3
   End
   Begin VB.TextBox txt_heat_mana_no1 
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
      Left            =   14910
      MaxLength       =   8
      TabIndex        =   5
      Tag             =   "HEAT_MANA_NO"
      Top             =   9195
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox txt_heat_mana_no 
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
      Left            =   7200
      MaxLength       =   8
      TabIndex        =   3
      Tag             =   "炉次管理号"
      Top             =   1050
      Visible         =   0   'False
      Width           =   915
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
      Height          =   315
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   120
      Width           =   1140
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
      Height          =   315
      Left            =   885
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   120
      Width           =   465
   End
   Begin CSTextLibCtl.sidbEdit sdb_heat_edt_seq 
      Height          =   315
      Left            =   6300
      TabIndex        =   2
      Tag             =   "炉次编制号"
      Top             =   1050
      Visible         =   0   'False
      Width           =   870
      _Version        =   262145
      _ExtentX        =   1535
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
      RawData         =   ""
      Text            =   " 0"
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
      NumIntDigits    =   5
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_prc_line 
      Height          =   315
      Left            =   9525
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   240
      _Version        =   262145
      _ExtentX        =   423
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      RawData         =   "1"
      Text            =   " 1"
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   16
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
      NumIntDigits    =   5
      Undo            =   0
      Data            =   1
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   150
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   10230
      Top             =   9195
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "炉次编制号"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   12630
      Top             =   9195
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "炉次管理号"
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   420
      Left            =   4245
      TabIndex        =   9
      Top             =   510
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   741
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_move 
         Height          =   285
         Left            =   75
         TabIndex        =   10
         Top             =   90
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "移动"
         Value           =   -1
      End
      Begin Threed.SSOption opt_split 
         Height          =   285
         Left            =   1785
         TabIndex        =   11
         Top             =   90
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "分开"
      End
      Begin Threed.SSOption opt_unification 
         Height          =   285
         Left            =   930
         TabIndex        =   12
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "统合"
      End
      Begin Threed.SSOption opt_delete 
         Height          =   285
         Left            =   2610
         TabIndex        =   13
         Top             =   90
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "删除"
      End
   End
   Begin Threed.SSPanel pnl_first 
      Height          =   420
      Left            =   7800
      TabIndex        =   14
      Top             =   510
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   741
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_target 
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
         Left            =   3750
         MaxLength       =   50
         TabIndex        =   26
         Tag             =   "工厂"
         Top             =   45
         Width           =   1200
      End
      Begin VB.TextBox txt_to 
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
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   25
         Tag             =   "工厂"
         Top             =   45
         Width           =   1200
      End
      Begin VB.TextBox txt_from 
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
         Left            =   600
         MaxLength       =   50
         TabIndex        =   24
         Tag             =   "工厂"
         Top             =   45
         Width           =   1200
      End
      Begin Threed.SSOption opt_top 
         Height          =   285
         Left            =   5070
         TabIndex        =   15
         Top             =   60
         Visible         =   0   'False
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "前"
      End
      Begin Threed.SSOption opt_bottom 
         Height          =   285
         Left            =   5625
         TabIndex        =   16
         Top             =   60
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "后"
         Value           =   -1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "对象"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   1800
         TabIndex        =   18
         Top             =   195
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "目标"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3270
         TabIndex        =   17
         Top             =   135
         Width           =   420
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   420
      Left            =   150
      TabIndex        =   20
      Top             =   510
      Width           =   3650
      _ExtentX        =   6429
      _ExtentY        =   741
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_charge 
         Height          =   285
         Left            =   900
         TabIndex        =   21
         Top             =   90
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "炉次"
      End
      Begin Threed.SSOption opt_slab 
         Height          =   285
         Left            =   2820
         TabIndex        =   22
         Top             =   90
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "板坯"
      End
      Begin Threed.SSOption opt_cast 
         Height          =   285
         Left            =   1665
         TabIndex        =   23
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "连浇炉数"
      End
      Begin Threed.SSOption opt_search 
         Height          =   285
         Left            =   105
         TabIndex        =   28
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "查询"
      End
   End
   Begin Threed.SSCommand cmd_OK 
      Height          =   420
      Left            =   14310
      TabIndex        =   27
      Top             =   510
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确定"
      BevelWidth      =   3
   End
   Begin CSTextLibCtl.sidbEdit sdb_plan_wgt 
      Height          =   315
      Left            =   5790
      TabIndex        =   29
      Tag             =   "炉次编制号"
      Top             =   120
      Width           =   960
      _Version        =   262145
      _ExtentX        =   1693
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
      RawData         =   ""
      Text            =   " 0"
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
      NumIntDigits    =   8
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   4740
      Top             =   120
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "计划材"
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
      Left            =   7050
      Top             =   120
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "非计划材"
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
   Begin CSTextLibCtl.sidbEdit sdb_nonplan_wgt 
      Height          =   315
      Left            =   8085
      TabIndex        =   32
      Tag             =   "炉次编制号"
      Top             =   120
      Width           =   960
      _Version        =   262145
      _ExtentX        =   1693
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
      RawData         =   ""
      Text            =   " 0"
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
      NumIntDigits    =   8
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSCommand cmd_Manual_Ordering 
      Height          =   420
      Left            =   9810
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   30
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   12582912
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
      Caption         =   "炼钢紧急编制"
      BevelWidth      =   3
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   2625
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "炉/机号"
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
   Begin Threed.SSCommand cmd_timeSet 
      Height          =   420
      Left            =   13890
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   30
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   8421376
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
      Caption         =   "设定时间"
      BevelWidth      =   3
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   375
      Left            =   165
      TabIndex        =   39
      Top             =   1020
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   0
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
   End
   Begin Threed.SSCommand MltCD_Changed 
      Height          =   390
      Left            =   2220
      TabIndex        =   40
      Top             =   1020
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   688
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "修改工艺路径"
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmd_lenchg 
      Height          =   390
      Left            =   4140
      TabIndex        =   41
      Top             =   1020
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   688
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
      Caption         =   "余坯长度变更确定"
      BevelWidth      =   3
   End
   Begin FPSpread.vaSpread ss3 
      Height          =   315
      Left            =   8250
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   270
      _Version        =   393216
      _ExtentX        =   476
      _ExtentY        =   556
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
      MaxCols         =   17
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AEC1070C.frx":1474
   End
   Begin FPSpread.vaSpread ss4 
      Height          =   315
      Left            =   8580
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   270
      _Version        =   393216
      _ExtentX        =   476
      _ExtentY        =   556
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
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AEC1070C.frx":1D2A
   End
   Begin FPSpread.vaSpread ss5 
      Height          =   315
      Left            =   8910
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   270
      _Version        =   393216
      _ExtentX        =   476
      _ExtentY        =   556
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
      MaxCols         =   17
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AEC1070C.frx":24E1
   End
   Begin FPSpread.vaSpread ss6 
      Height          =   315
      Left            =   9240
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   270
      _Version        =   393216
      _ExtentX        =   476
      _ExtentY        =   556
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
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AEC1070C.frx":2D97
   End
   Begin Threed.SSCommand cmd_sample 
      Height          =   420
      Left            =   12520
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   990
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   12583104
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
      Caption         =   "按取样信息"
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmd_recast 
      Height          =   420
      Left            =   12540
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   30
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   741
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
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
      Caption         =   "连浇再编制"
      BevelWidth      =   3
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "吨"
      Height          =   270
      Index           =   1
      Left            =   9075
      TabIndex        =   31
      Top             =   165
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "吨"
      Height          =   270
      Index           =   0
      Left            =   6780
      TabIndex        =   30
      Top             =   165
      Width           =   195
   End
End
Attribute VB_Name = "AEC1070C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       DAILY SCHEDULE
'-- Sub_System Name
'-- Program Name
'-- Program ID        AEC1070C
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

Public Complete As Boolean          'Move Status Setting

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
Dim Mc3 As New Collection           'Master Collection
Dim Mc4 As New Collection           'Master Collection

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

Dim sLoc        As String
Dim P_PLT       As String           'PLT
Dim P_LINE      As Integer          'LINE = '1'
Dim P_STATUS    As String           'DAILY = 'D', INSTRUCTION = 'I'
Dim P_MODE      As String           'MOVE = 'M',  SPLIT = 'S', UNIFICATION = 'U', DELETE = 'D'
Dim P_UNIT      As String           'PLATE = 'P', SLAB = 'S',  CHARGE = 'H', CAST = 'C', ROLL = 'R'
Dim P_POSITION  As String           'TOP = 'T',   BOTTOM = 'B'

Public strCCM_CD1      As String    'Current mlt_cd
Public strCCM_CD1_Pre  As String    'Pre_Select mlt_cd
Public strCCM_CD2      As String
Public strCCM_CD3      As String
Public lngCurRow       As Long      'Current Select Row
Public lngPreRow       As Long      'Pre_Select Row
Public intCount        As Integer   'Count of Selected Row Number

Public intSeq          As Long

Public ROW1            As Long      'Current Select Row  一次操作插入三块余坯
Public ROW2            As Long      'Current Select Row
Public ROW3            As Long      'Current Select Row

Public strSlabNo1      As String    'strSlabNo1
Public strSlabNo2      As String    'strSlabNo2
Public strSlabNo3      As String

Dim oActiveSp As Object             'Active Spread Sheet

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(txt_prc_line1, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(sdb_heat_edt_seq, "p", " ", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_heat_mana_no, "p", " ", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
            Call Gp_Ms_Collection(txt_from, " ", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
              Call Gp_Ms_Collection(txt_to, " ", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_target, " ", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(cbo_prc_line, " ", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(cbo_chg_prc_line, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
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
    Call Gp_Ms_Collection(txt_heat_mana_no1, "p", " ", " ", " ", "r", " ", "l", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    
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
             Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
       Call Gp_Ms_Collection(txt_prc_line2, "p", "n", " ", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
        Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    Call Gp_Ms_Collection(sdb_heat_edt_seq, "p", " ", " ", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    Call Gp_Ms_Collection(txt_heat_mana_no, "p", " ", " ", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
            Call Gp_Ms_Collection(txt_from, " ", " ", " ", " ", " ", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
              Call Gp_Ms_Collection(txt_to, " ", " ", " ", " ", " ", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
          Call Gp_Ms_Collection(txt_target, " ", " ", " ", " ", " ", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
       'Call Gp_Ms_Collection(sdb_prc_line, "p", "n", "m", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    
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
             Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
       Call Gp_Ms_Collection(txt_prc_line3, "p", "n", " ", " ", "r", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
        Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
    Call Gp_Ms_Collection(sdb_heat_edt_seq, "p", " ", " ", " ", "r", " ", "l", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
    Call Gp_Ms_Collection(txt_heat_mana_no, "p", " ", " ", " ", "r", " ", "l", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
            Call Gp_Ms_Collection(txt_from, " ", " ", " ", " ", " ", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
              Call Gp_Ms_Collection(txt_to, " ", " ", " ", " ", " ", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
          Call Gp_Ms_Collection(txt_target, " ", " ", " ", " ", " ", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
       'Call Gp_Ms_Collection(sdb_prc_line, "p", "n", "m", " ", "r", " ", "l", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
    
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
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AEC1070C.P_SMODIFY", Key:="P-M"
    Sc1.Add Item:="AEC1070C.P_REFER1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(SS2, 1, " ", "p", "n", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(SS2, 2, " ", "p", "n", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)  ' 钢种
     Call Gp_Sp_Collection(SS2, 3, " ", "p", "n", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)  ' 板坯厚度
     Call Gp_Sp_Collection(SS2, 4, " ", "p", "n", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)  ' 板坯宽度
     Call Gp_Sp_Collection(SS2, 5, " ", "p", "n", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)  ' 板坯长度
     Call Gp_Sp_Collection(SS2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(SS2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(SS2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(SS2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(SS2, 18, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    
    'Spread_Collection
    sc2.Add Item:=SS2, Key:="Spread"
    sc2.Add Item:="AEC1070C.P_SMODIFY1", Key:="P-M"
    sc2.Add Item:="AEC1070C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=SS2.MaxCols, Key:="Last"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(SS3, 1, "p", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3, False)
     Call Gp_Sp_Collection(SS3, 2, "p", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3, False)
     Call Gp_Sp_Collection(SS3, 3, "p", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3, False)
     Call Gp_Sp_Collection(SS3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(SS3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(SS3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(SS3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(SS3, 8, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3, False)
     Call Gp_Sp_Collection(SS3, 9, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(SS3, 10, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3, False)
    Call Gp_Sp_Collection(SS3, 11, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(SS3, 12, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(SS3, 13, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(SS3, 14, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(SS3, 15, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(SS3, 16, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(SS3, 17, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    'Spread_Collection
    Sc3.Add Item:=SS3, Key:="Spread"
    Sc3.Add Item:="AEC1070C.P_SMODIFY", Key:="P-M"
    Sc3.Add Item:="AEC1070C.P_REFER1", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=3, Key:="First"
    Sc3.Add Item:=SS3.MaxCols, Key:="Last"
    
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
    Call Gp_Sp_Collection(ss4, 13, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 14, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 15, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    
    'Spread_Collection
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:="AEC1070C.P_REFER2", Key:="P-R"
    Sc4.Add Item:=pColumn4, Key:="pColumn"
    Sc4.Add Item:=nColumn4, Key:="nColumn"
    Sc4.Add Item:=aColumn4, Key:="aColumn"
    Sc4.Add Item:=mColumn4, Key:="mColumn"
    Sc4.Add Item:=iColumn4, Key:="iColumn"
    Sc4.Add Item:=lColumn4, Key:="lColumn"
    Sc4.Add Item:=1, Key:="First"
    Sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss5, 1, "p", " ", " ", "i", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5, False)
     Call Gp_Sp_Collection(ss5, 2, "p", " ", " ", "i", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5, False)
     Call Gp_Sp_Collection(ss5, 3, "p", " ", " ", "i", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5, False)
     Call Gp_Sp_Collection(ss5, 4, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 5, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 6, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 7, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 8, " ", " ", " ", "i", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5, False)
     Call Gp_Sp_Collection(ss5, 9, " ", " ", " ", "i", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 10, " ", " ", " ", "i", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5, False)
    Call Gp_Sp_Collection(ss5, 11, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 12, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 13, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 14, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 15, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 16, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 17, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    
    'Spread_Collection
    Sc5.Add Item:=ss5, Key:="Spread"
    Sc5.Add Item:="AEC1070C.P_SMODIFY", Key:="P-M"
    Sc5.Add Item:="AEC1070C.P_REFER1", Key:="P-R"
    Sc5.Add Item:=pColumn5, Key:="pColumn"
    Sc5.Add Item:=nColumn5, Key:="nColumn"
    Sc5.Add Item:=aColumn5, Key:="aColumn"
    Sc5.Add Item:=mColumn5, Key:="mColumn"
    Sc5.Add Item:=iColumn5, Key:="iColumn"
    Sc5.Add Item:=lColumn5, Key:="lColumn"
    Sc5.Add Item:=3, Key:="First"
    Sc5.Add Item:=ss5.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss6, 1, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 2, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 3, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 4, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 5, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 6, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 7, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 8, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 9, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 10, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 11, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 12, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 13, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 14, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 15, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    
    'Spread_Collection
    Sc6.Add Item:=ss6, Key:="Spread"
    Sc6.Add Item:="AEC1070C.P_REFER2", Key:="P-R"
    Sc6.Add Item:=pColumn6, Key:="pColumn"
    Sc6.Add Item:=nColumn6, Key:="nColumn"
    Sc6.Add Item:=aColumn6, Key:="aColumn"
    Sc6.Add Item:=mColumn6, Key:="mColumn"
    Sc6.Add Item:=iColumn6, Key:="iColumn"
    Sc6.Add Item:=lColumn6, Key:="lColumn"
    Sc6.Add Item:=1, Key:="First"
    Sc6.Add Item:=ss6.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc"
    Proc_Sc.Add Item:=Sc4, Key:="Sc4"
    Proc_Sc.Add Item:=Sc6, Key:="Sc6"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cbo_prc_line_Change()

    If cbo_prc_line.Text = "1# BOF" Then
        txt_prc_line1.Text = "1"
        txt_prc_line2.Text = "2"
        txt_prc_line3.Text = "3"
        strCCM_CD1 = ""
        strCCM_CD2 = ""
        strCCM_CD3 = ""
    ElseIf cbo_prc_line.Text = "2# BOF" Then
        txt_prc_line1.Text = "2"
        txt_prc_line2.Text = "1"
        txt_prc_line3.Text = "3"
        strCCM_CD1 = ""
        strCCM_CD2 = ""
        strCCM_CD3 = ""
    ElseIf cbo_prc_line.Text = "3# BOF" Then
        txt_prc_line1.Text = "3"
        txt_prc_line2.Text = "1"
        txt_prc_line3.Text = "2"
        strCCM_CD1 = ""
        strCCM_CD2 = ""
        strCCM_CD3 = ""
    ElseIf cbo_prc_line.Text = "1# CCM" Then
        txt_prc_line1.Text = "4"
        txt_prc_line2.Text = "5"
        txt_prc_line3.Text = "6"
        strCCM_CD1 = ""
        strCCM_CD2 = ""
        strCCM_CD3 = ""
    ElseIf cbo_prc_line.Text = "2# CCM" Then
        txt_prc_line1.Text = "5"
        txt_prc_line2.Text = "4"
        txt_prc_line3.Text = "6"
        strCCM_CD1 = ""
        strCCM_CD2 = ""
        strCCM_CD3 = ""
    Else
        txt_prc_line1.Text = "6"
        txt_prc_line2.Text = "4"
        txt_prc_line3.Text = "5"
        strCCM_CD1 = ""
        strCCM_CD2 = ""
        strCCM_CD3 = ""
    End If
    
    'Call Form_Ref
    lngCurRow = 0
    
End Sub

Private Sub cbo_prc_line_Click()

    If cbo_prc_line.Text = "1# BOF" Then
        txt_prc_line1.Text = "1"
        txt_prc_line2.Text = "2"
        txt_prc_line3.Text = "3"
    ElseIf cbo_prc_line.Text = "2# BOF" Then
        txt_prc_line1.Text = "2"
        txt_prc_line2.Text = "1"
        txt_prc_line3.Text = "3"
    ElseIf cbo_prc_line.Text = "3# BOF" Then
        txt_prc_line1.Text = "3"
        txt_prc_line2.Text = "1"
        txt_prc_line3.Text = "2"
    ElseIf cbo_prc_line.Text = "1# CCM" Then
        txt_prc_line1.Text = "4"
        txt_prc_line2.Text = "5"
        txt_prc_line3.Text = "6"
    ElseIf cbo_prc_line.Text = "2# CCM" Then
        txt_prc_line1.Text = "5"
        txt_prc_line2.Text = "4"
        txt_prc_line3.Text = "6"
    Else
        txt_prc_line1.Text = "6"
        txt_prc_line2.Text = "4"
        txt_prc_line3.Text = "5"
    End If
    
    Call Form_Ref
    
End Sub

Private Sub cmd_cast1_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    Dim sCcm_line As String
    
    Dim adoCmd As adodb.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    '---------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------
    ' 3#LINE  CALL AEC1000P
    '---------------------------------------------------------------------------------------
    If cbo_prc_line.Text = "1# CCM" Or cbo_prc_line.Text = "1# BOF" Then
        sCcm_line = "1"
    ElseIf cbo_prc_line.Text = "2# CCM" Or cbo_prc_line.Text = "2# BOF" Then
        sCcm_line = "2"
    ElseIf cbo_prc_line.Text = "3# CCM" Or cbo_prc_line.Text = "3# BOF" Then
        sCcm_line = "3"
    End If
    
    sQuery = "{call AEC1000P ('" + sCcm_line + "','" + sUserID + "',?)}"
    
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
        Call Form_Ref
    Else
        Call Gp_MsgBoxDisplay("连浇炉数编制完了..!!", "I")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_confirm_Click()

    Complete = False

    'If ss1.MaxRows = 0 Then Exit Sub
       
    Load Ins_Confirm
    
    Ins_Confirm.P_MODE = "C"           'CAST
    Ins_Confirm.P_PLT = txt_plt.Text   'PLT
    Ins_Confirm.P_LINE = "1"           'LINE
    
    Ins_Confirm.P_CurrentCol = 2
    
    If cbo_prc_line.Text = "1# CCM" Or cbo_prc_line.Text = "1# BOF" Then
        Set Active_Spread = Me.ss1
        Call Ins_Confirm.Gp_Combo_Add(ss1)
        Call Ins_Confirm.Gp_Combo_Add2(SS3)
        Call Ins_Confirm.Gp_Combo_Add3(ss5)
    ElseIf cbo_prc_line.Text = "2# CCM" Or cbo_prc_line.Text = "2# BOF" Then
        Set Active_Spread = Me.ss1
        Call Ins_Confirm.Gp_Combo_Add(SS3)
        Call Ins_Confirm.Gp_Combo_Add2(ss1)
        Call Ins_Confirm.Gp_Combo_Add3(ss5)
    ElseIf cbo_prc_line.Text = "3# CCM" Or cbo_prc_line.Text = "3# BOF" Then
        Set Active_Spread = Me.ss1
        Call Ins_Confirm.Gp_Combo_Add(SS3)
        Call Ins_Confirm.Gp_Combo_Add2(ss5)
        Call Ins_Confirm.Gp_Combo_Add3(ss1)
    End If
    
    Ins_Confirm.Show 1
    
    'If Complete Then
    '    Call Form_Ref
    'End If
    
End Sub

Private Sub cmd_lenchg_Click()

On Error GoTo Process_Exec_ERROR

    Dim iRow As Long
    Dim sSLAB_NO As String
    Dim sSLAB_LEN As Double
    
    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    sSLAB_NO = ""
    sSLAB_LEN = 0
    
    If SS2.MaxRows <= 0 Then Exit Sub
        
    For iRow = 1 To SS2.MaxRows
                
        SS2.Row = iRow
        SS2.Col = 0
        
        If SS2.Text = "Update" Then
           SS2.Col = 1:      sSLAB_NO = SS2.Text
           SS2.Col = 5:      sSLAB_LEN = SS2.Value
           Exit For
        End If
        
    Next iRow
    
    If sSLAB_NO = "" Then Exit Sub
    
    If sSLAB_LEN < 1700 Or sSLAB_LEN > 17600 Then
       Call Gp_MsgBoxDisplay(" 板坯长度错误，请确认", "I")
       Exit Sub
    End If
    
    Dim adoCmd As adodb.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEC1080P('" & sSLAB_NO & "'," & sSLAB_LEN & ",?)}"
    
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("板坯长度变更成功..!!", "I")
        Call Form_Ref
    End If
    
    cmd_roll1.Enabled = False
    
    Set adoCmd = Nothing
    
    For iRow = 1 To ss1.MaxRows
                
        ss1.Row = iRow
        ss1.Col = 1
        
        If ss1.Text = Mid(sSLAB_NO, 1, 8) Then
           Call ss1_Click(1, iRow)
        End If
        
    Next iRow
    
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_Manual_Ordering_Click()

    Load AEC0000C
    AEC0000C.Show 1
    
End Sub

Private Sub cmd_recast_Click()
On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    Dim sCcm_line As String
    
    Dim adoCmd As adodb.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    '---------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------
    ' 3#LINE  CALL AEC1000P
    '---------------------------------------------------------------------------------------
    If cbo_prc_line.Text = "1# CCM" Or cbo_prc_line.Text = "1# BOF" Then
        sCcm_line = "1"
    ElseIf cbo_prc_line.Text = "2# CCM" Or cbo_prc_line.Text = "2# BOF" Then
        sCcm_line = "2"
    ElseIf cbo_prc_line.Text = "3# CCM" Or cbo_prc_line.Text = "3# BOF" Then
        sCcm_line = "3"
    End If
    
    sQuery = "{call AEC1005P ('" + sCcm_line + "','" + sUserID + "',?)}"
    
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("连浇炉数再编制结束..!!", "I")
        Call Form_Ref
    End If
    
    cmd_roll1.Enabled = False
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_roll1_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As adodb.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    'sQuery = "SELECT COUNT(*) FROM EP_SLAB_EDT WHERE HCR_FL = 'H' "
    'iCount = Gf_FloatFind(M_CN1, sQuery)
    
    'If iCount > 0 Then  'HCR
        sQuery = "{call AEC2030P ('" + txt_plt.Text + "','1',?)}"
    'Else                'CCR
    '    sQuery = "{call AED1040P ('" + txt_plt.Text + "','1',?)}"
    'End If
    
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("编制轧辊单位完了..!!", "I")
        '------------------------------------------------
        '20060714 HJD Modified
        '------------------------------------------------
        'Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub cmd_sample_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As adodb.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEC8000P ('" + sUserID + "',?)}"
                
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("按取样信息完了..!!", "I")
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)

End Sub

Private Sub cmd_timeSet_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As adodb.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEC1040P (?)}"
    
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("时间设定完毕!!", "I")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
'    MDIMain.MenuTool.Buttons(4).Enabled = False  保存
'    MDIMain.MenuTool.Buttons(7).Enabled = False  新增行
'    MDIMain.MenuTool.Buttons(8).Enabled = False
'    MDIMain.MenuTool.Buttons(9).Enabled = False  取消行
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

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
    
    'UPDATE AUTHORITY
    If Mid(sAuthority, 3, 1) <> "1" Then
        opt_search.Enabled = False
        opt_charge.Enabled = False
        opt_cast.Enabled = False
        opt_slab.Enabled = False
        cmd_cast1.Enabled = False
        cmd_recast.Enabled = False
        cmd_roll1.Enabled = False
        cmd_sample.Enabled = False
        cmd_confirm.Enabled = False
        cmd_lenchg.Enabled = False  ' 长度变更确认      杨猛 0809041139
    End If

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
'    MDIMain.MenuTool.Buttons(4).Enabled = False
'    MDIMain.MenuTool.Buttons(7).Enabled = False
'    MDIMain.MenuTool.Buttons(8).Enabled = False
'    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc4.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc5.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc6.Item("Spread"), False)
    
'    Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
'    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
'    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc4.Item("Spread"))
'    Call Gp_Sp_ReadOnlySet(Sc5.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc6.Item("Spread"))
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(Sc4)
    Call Gf_Sp_Cls(Sc5)
    Call Gf_Sp_Cls(Sc6)
    
    Call Gp_Spl_SizeGet(SSSplitter2, "E-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc4.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc5.Item("Spread"), "E-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc6.Item("Spread"), "E-System.INI", Me.Name)
    
    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    cbo_prc_line.Clear
    cbo_prc_line.AddItem "1# CCM"
    cbo_prc_line.AddItem "2# CCM"
    cbo_prc_line.AddItem "3# CCM"
'    cbo_prc_line.AddItem "1# BOF"
'    cbo_prc_line.AddItem "2# BOF"
'    cbo_prc_line.AddItem "3# BOF"
    cbo_prc_line.ListIndex = 0
    txt_prc_line1.Text = "4"
    txt_prc_line2.Text = "5"
    txt_prc_line3.Text = "6"
    
    cbo_chg_prc_line.Clear
    cbo_chg_prc_line.AddItem "1"
    cbo_chg_prc_line.AddItem "2"
    cbo_chg_prc_line.AddItem "3"
    
    P_MODE = "M"
    P_POSITION = "T"
    strCCM_CD1 = ""
    strCCM_CD2 = ""
    strCCM_CD3 = ""
    lngCurRow = 0
    Call chk_ss1_UnChecked
    '炼钢紧急编制 AUTHORIT CHECK
    If Mid(Gf_Pgm_Authority("AEC0000C"), 2, 1) <> "1" Then
        cmd_Manual_Ordering.Enabled = False
    End If

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter2, "E-System.INI", Me.Name)
    
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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
        intSeq = 0
        If Gf_Sp_Cls(Sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gf_Sp_Cls(Sc3)
            Call Gf_Sp_Cls(Sc4)
            Call Gf_Sp_Cls(Sc5)
            Call Gf_Sp_Cls(Sc6)
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
'            MDIMain.MenuTool.Buttons(4).Enabled = False
'            MDIMain.MenuTool.Buttons(7).Enabled = False
'            MDIMain.MenuTool.Buttons(8).Enabled = False
'            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            rContro1(1).SetFocus
            txt_plt.Text = "B1"
            Call txt_plt_KeyUp(0, 0)
            cbo_prc_line.ListIndex = 0
            txt_prc_line1.Text = "4"
            txt_prc_line2.Text = "5"
            txt_prc_line3.Text = "6"
            cbo_chg_prc_line.Text = ""
            sLoc = "F"
            txt_from.BackColor = &HC0FFFF
            txt_to.BackColor = &H80000005
            txt_target.BackColor = &H80000005
            opt_search.Value = True
            Call chk_ss1_UnChecked
        End If
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sTemp As String
    Dim sQuery As String
    Dim iRow As Integer
    Dim iCount As Integer
    Dim AdoRs As New adodb.Recordset
    
    cmd_confirm.Visible = True
    intCount = 0
    intSeq = 0
    
    Call Gf_Sp_Cls(Sc1)
    SS2.MaxRows = 0
'    Call Gf_Sp_Cls(Sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(Sc4)
    Call Gf_Sp_Cls(Sc5)
    Call Gf_Sp_Cls(Sc6)
    
    'Call chk_ss1_UnChecked
    
'    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Or _
       Gf_Sp_Refer(M_CN1, Sc3, Mc3, Mc3("nControl"), Mc3("mControl"), False) Or _
       Gf_Sp_Refer(M_CN1, Sc5, Mc4, Mc4("nControl"), Mc4("mControl"), False) Then
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'        MDIMain.MenuTool.Buttons(4).Enabled = False
'        MDIMain.MenuTool.Buttons(7).Enabled = False
'        MDIMain.MenuTool.Buttons(8).Enabled = False
'        MDIMain.MenuTool.Buttons(9).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
        
        txt_from.Text = ""
        txt_to.Text = ""
        txt_target.Text = ""
        sdb_plan_wgt.Value = 0
        sdb_nonplan_wgt.Value = 0
        
        Call Spread_Color_Setting(ss1)
        Call Spread_Color_Setting(SS3)
        Call Spread_Color_Setting(ss5)
        
        ss1.OperationMode = OperationModeNormal
        SS3.OperationMode = OperationModeNormal
        ss5.OperationMode = OperationModeNormal
        
'        Call Gp_Sp_EvenRowBackcolor(Sc1.Item("Spread"))
'        Call Gp_Sp_EvenRowBackcolor(Sc3.Item("Spread"))

'        sQuery = sQuery & "SELECT COUNT(*) FROM EP_SLAB_EDT WHERE SLAB_EDT_FL = '1' "
'        iCount = Gf_FloatFind(M_CN1, sQuery)

        sQuery = "SELECT        SUM(DECODE(SLAB_EDT_FL,'1',1,0))   EDT_CNT       "
        sQuery = sQuery & "    ,SUM(DECODE(ORD_FL,'1',SLAB_WGT,0)) PLAN_WGT      "
        sQuery = sQuery & "    ,SUM(DECODE(ORD_FL,'2',SLAB_WGT,0)) NON_PLAN_WGT  "
        sQuery = sQuery & "FROM EP_SLAB_EDT  "
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
        
        If Not AdoRs.EOF Then
        
            iCount = Val(AdoRs.Fields("EDT_CNT") & "")
            sdb_plan_wgt.Value = Val(AdoRs.Fields("PLAN_WGT") & "")
            sdb_nonplan_wgt.Value = Val(AdoRs.Fields("NON_PLAN_WGT") & "")
            
        End If
        
        AdoRs.Close
        Set AdoRs = Nothing
    
'        If iCount > 0 Then  'HCR
'            cmd_confirm.Visible = True
'        Else                'CCR
'            cmd_confirm.Visible = False
'        End If
        
    End If
    
End Sub

Public Sub Spread_Color_Setting(oSpr As vaSpread)

    Dim iRow  As Integer
    Dim sTemp As String
    
    With oSpr
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 2
            If iRow = 1 Then sTemp = .Text
            
            If sTemp <> .Text Then
                sTemp = .Text
                Call Gp_Sp_BlockColor(oSpr, 1, .MaxCols, iRow, iRow, , &HFFC0FF)
            End If
        Next iRow
    End With
    
End Sub

Public Sub Form_Pro()

'    检查数据
    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Or Gf_Sp_Process(M_CN1, Sc3, Mc3) Or Gf_Sp_Process(M_CN1, sc2, Mc2) Or Gf_Sp_Process(M_CN1, Sc5, Mc4) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'        MDIMain.MenuTool.Buttons(4).Enabled = False
'        MDIMain.MenuTool.Buttons(7).Enabled = False
'        MDIMain.MenuTool.Buttons(8).Enabled = False
'        MDIMain.MenuTool.Buttons(9).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
        'Call Prc_Changed
        Call Form_Ref
    End If
    
End Sub

'Public Sub Form_Ins()
'
'    Call Gp_Sp_Ins(Proc_Sc("Sc"))
'
'
'End Sub
'
'Public Sub Spread_Cpy()
'
'    Call Gp_Sp_Copy(Proc_Sc("Sc"))
'
'End Sub
'
'Public Sub Spread_Pst()
'
'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
'
'End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()
    
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
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

Private Sub MltCD_Changed_Click()

   If opt_line_change.Value Then
'         If Trim(txt_from.Text) = "" Or Trim(txt_to.Text) = "" _
'            Or Val(txt_from.Text) > Val(txt_to.Text) Then
'             Call Gp_MsgBoxDisplay("请选择要插入工作炉座的正确炉次范围", "", "错误提示")
'             Exit Sub
'        End If
'        If MltCD_Changed.Caption = "插入工作炉座号" Then
'            MltCD_Changed.Caption = "撤销工作炉座插入"
'        Else
'            MltCD_Changed.Caption = "插入工作炉座号"
'        End If
        'Call Bof_Insert
        
        If cbo_chg_prc_line.Text = "" Then
            Exit Sub
        End If
        
        Call Bof_Switch
   Else
    
        Load Mltcd_Change
        
        Set Active_Spread = ss1
        Mltcd_Change.txt_MLT_PROC_CD_ORG.Text = strCCM_CD1
      
        strCCM_CD1_Pre = ""
        strCCM_CD1 = ""
        lngCurRow = 0
        lngPreRow = 0
    
        Mltcd_Change.Show 1
        
        Call Spread_Color_Setting(ss1)
        Call Spread_Color_Setting(SS3)
        Call Spread_Color_Setting(ss5)
        
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim sTemp As String
    Dim iRow As Integer
    
    'Call Gp_Sp_Sort(Sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If Row < 1 Then Exit Sub
    
    If ss1.MaxRows < 1 Then
        txt_heat_mana_no1.Text = ""
        Call Gf_Sp_Cls(sc2)
        Exit Sub
    End If
    
    '---2006-07-14 ADD By HJD------------------------------------------------
    If lngCurRow <> Row Then
        If intCount = 0 Then strCCM_CD1 = "": lngCurRow = 0
        strCCM_CD1_Pre = strCCM_CD1
        ss1.Row = Row: ss1.Col = 10: strCCM_CD1 = Trim(ss1.Text)
        'strCCM_CD1 = Mid(ss1.Text, InStr(1, ss1.Text, "BF"), 3)
        lngPreRow = lngCurRow
        lngCurRow = Row
    End If

'    If opt_line_change.Value = True And Col = 8 And cbo_chg_prc_line.Text <> "" Then
'
'        With ss1
'            .Row = Row
'            .Col = 8
'            .Text = cbo_chg_prc_line.Text     'IIf(.Text = "1", "2", "1")
'            .Col = 9:
'
'            'Replace(.Text, IIf(.Text Like "BC1*", "BC1", "BC2"), IIf(.Text Like "BC1*", "BC2", "BC1"), 1)
'
'            If .Text Like "BC1*" Then
'                .Text = Replace(.Text, "BC1", "BC" & cbo_chg_prc_line.Text, 1)
'            ElseIf .Text Like "BC2*" Then
'                .Text = Replace(.Text, "BC2", "BC" & cbo_chg_prc_line.Text, 1)
'            Else
'                .Text = Replace(.Text, "BC3", "BC" & cbo_chg_prc_line.Text, 1)
'            End If
'
'        End With
'
'        Call BOF_Changed(ss1)
'
'    End If
    
    If opt_mltcd_change.Value = True Then
    
        ss1.Col = 0: ss1.Row = Row
        
        If ss1.Text <> "Selected" Then
            ss1.Col = 10
            If ss1.Text <> strCCM_CD1_Pre And Trim(strCCM_CD1_Pre) <> "" Then lngCurRow = lngPreRow: strCCM_CD1 = strCCM_CD1_Pre: Call Gp_MsgBoxDisplay("冶炼工艺不一样"): Exit Sub
            ss1.Col = 0
            ss1.Text = "Selected"
            intCount = intCount + 1
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
        Else
            ss1.Text = ""
            intCount = intCount - 1
            ss1.Col = 3
            If ss1.Text = 1 And Row <> 1 Then
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFC0FF)
            Else
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
            End If
            
        End If
        
    End If
    
    '-----------------------------------------------------------------------
    If opt_charge.Value = True Or opt_cast.Value = True Then Exit Sub
    
    ss1.Row = Row
'    If Gf_Sp_ProceExist(Sc2.Item("Spread")) Then Exit Sub
    
    ss1.Col = 1
    txt_heat_mana_no1.Text = ss1.Text
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    intSeq = 0
    
    SS2.OperationMode = OperationModeNormal
    
    If Row = ss1.MaxRows Then
    
        For iRow = 1 To SS2.MaxRows
                    
            SS2.Row = iRow
            SS2.Col = 1
            
            If iRow = 1 Then sTemp = Mid(SS2.Text, 1, 8)
            
            If sTemp <> Mid(SS2.Text, 1, 8) Then
                sTemp = Mid(SS2.Text, 1, 8)
                Call Gp_Sp_BlockColor(SS2, 1, SS2.MaxCols, iRow, iRow, , &HFFC0FF)
            End If
                     
            SS2.Col = 9
            If SS2.Text = "2" Then
                SS2.Col = 5:       SS2.Col2 = 5
                SS2.Row = iRow:    SS2.ROW2 = iRow
                SS2.Lock = False
                Call Gp_Sp_BlockColor(SS2, 5, 5, iRow, iRow, , &HFFC0FF)
            End If
            
        Next iRow
        
    End If

End Sub

Private Sub ss1_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = Col1 To Col2
    
        SS3.ColWidth(iCol) = ss1.ColWidth(iCol)
        ss5.ColWidth(iCol) = ss1.ColWidth(iCol)
        
    Next iCol
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row < 1 Then Exit Sub

   '--------------------------Modified by hjd----2006-7-31----------------------------------------------------------
    If opt_charge.Value = True Or opt_cast.Value = True Or opt_line_change.Value = True Then
    
        If sLoc = "A" And (opt_split.Value = True Or opt_unification.Value = True) Then
            ss1.Col = 1
        ElseIf opt_charge.Value = True Or opt_line_change.Value = True Then
            ss1.Col = 1
        Else
            ss1.Col = 2
        End If
        ss1.Row = Row
        Call Location_edit(ss1.Text)
        
    End If
    
End Sub

'Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)
'    '---------------------2006-07-14 ADD By HJD------------------------------------------------------
''    Dim strPrcCD As String
''    If Row < 1 Or ss1.MaxRows < 1 Or _
''       opt_mltcd_change.Value = False Then Exit Sub
''    With ss1
''        .Row = Row
'''        .Col = 0
'''        .Text = "Update"
''        .Col = Col: strPrcCD = .Text
''        .Col = 8
''        .Text = IIf(strPrcCD Like "BC1*", "1", "2")
''        .Col = 9
''        If Mid(.Text, InStr(1, .Text, "BF"), 3) <> strCCM_CD1 Then
''            .Text = Replace(.Text, Mid(.Text, InStr(1, .Text, "BF"), 3), strCCM_CD1, 1)
''        End If
''    End With
'   '----------------------------------------------------------------------------------------------------
'
'End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
'    Call Gp_Sp_UpdateMake(Sc1.Item("Spread"), Mode)
    Call Gp_Sp_UpdateMake(Sc1.Item("Spread"), Mode)
    
End Sub

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    
    If RowChanged Then
        Call BOF_Changed(ss1)
    End If
    
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

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

End Sub

Private Sub ss2_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = Col1 To Col2
    
        ss4.ColWidth(iCol) = SS2.ColWidth(iCol)
        ss6.ColWidth(iCol) = SS2.ColWidth(iCol)
        
    Next iCol

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row < 1 Then Exit Sub
    
    If opt_slab.Value = True Then
        SS2.Row = Row
        SS2.Col = 1
        Call Location_edit(SS2.Text)
    End If

End Sub

Private Sub Location_edit(ByVal sMatNo As String)

    Select Case sLoc
        Case "F"
            txt_from.Text = sMatNo
            If txt_to.Enabled = True Then
                Call txt_to_Click
            Else
                Call txt_target_Click
            End If
            txt_from.BackColor = &H80000005
        Case "T"
            txt_to.Text = sMatNo
            txt_to.BackColor = &H80000005
            If txt_target.Enabled = True Then
                Call txt_target_Click
            Else
                Call txt_to_Click
            End If
        Case "A"
            txt_target.Text = sMatNo
            txt_target.BackColor = &H80000005
    End Select
    
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(sc2.Item("Spread"), Mode)
    End If
    
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If sc2.Item("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(sc2)
    End If

    If Shift = 0 Then sc2.Item("Spread").EditMode = True

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

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)

    Dim sTemp As String
    Dim iRow As Integer
    
    'Call Gp_Sp_Sort(Sc3.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If Row < 1 Then Exit Sub
    
    If SS3.MaxRows < 1 Then
        txt_heat_mana_no1.Text = ""
        Call Gf_Sp_Cls(Sc4)
        Exit Sub
    End If
    
    '----------------2006-07-14 ADD By HJD--------------------------------------------------------------------
    If lngCurRow <> Row Then
        SS3.Row = Row: SS3.Col = 10: strCCM_CD2 = Trim(SS3.Text)
        'strCCM_CD2 = Mid(ss3.Text, InStr(1, ss3.Text, "BF"), 3)
        lngCurRow = Row
    End If
    
    If opt_line_change.Value = True And Col = 8 And cbo_chg_prc_line.Text <> "" Then

        With SS3
            .Row = Row
                .Col = 8
                .Text = cbo_chg_prc_line.Text     'IIf(.Text = "1", "2", "1")
                .Col = 10

                '.Text = Replace(.Text, IIf(.Text Like "BC1*", "BC1", "BC2"), IIf(.Text Like "BC1*", "BC2", "BC1"), 1)

                If .Text Like "BC1*" Then
                    .Text = Replace(.Text, "BC1", "BC" & cbo_chg_prc_line.Text, 1)
                ElseIf .Text Like "BC2*" Then
                    .Text = Replace(.Text, "BC2", "BC" & cbo_chg_prc_line.Text, 1)
                Else
                    .Text = Replace(.Text, "BC3", "BC" & cbo_chg_prc_line.Text, 1)
                End If
        End With

        Call BOF_Changed(SS3)

    End If
    
'    If opt_mltcd_change.Value = True Then
'        ss3.Col = 0: ss3.Row = Row
'        If ss3.Text <> "Selected" Then
'            ss3.Text = "Selected"
'            Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, Row, Row, , &HFFFF80)
'        Else
'            ss3.Text = ""
'            Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, Row, Row)
'        End If
'
'    End If

    
    '-----------------------------------------------------------------------------------------------------------
    If opt_charge.Value = True Or opt_cast.Value = True Then Exit Sub
    
    SS3.Row = Row
    'If Gf_Sp_ProceExist(Sc4.Item("Spread")) Then Exit Sub
    
    SS3.Col = 1
    txt_heat_mana_no1.Text = SS3.Text
    
    Call Gf_Sp_Refer(M_CN1, Sc4, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    ss4.OperationMode = OperationModeNormal
    'Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc")("Spread"))
    
    For iRow = 1 To ss4.MaxRows
                
        ss4.Row = iRow
        ss4.Col = 1
        
        If iRow = 1 Then sTemp = Mid(ss4.Text, 1, 8)
        
        If sTemp <> Mid(ss4.Text, 1, 8) Then
            sTemp = Mid(ss4.Text, 1, 8)
            Call Gp_Sp_BlockColor(ss4, 1, ss4.MaxCols, iRow, iRow, , &HFFC0FF)
        End If
    
    Next iRow

End Sub

Private Sub ss3_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = Col1 To Col2
    
        ss1.ColWidth(iCol) = SS3.ColWidth(iCol)
        ss5.ColWidth(iCol) = SS3.ColWidth(iCol)
        
    Next iCol

End Sub

Private Sub ss3_EditChange(ByVal Col As Long, ByVal Row As Long)

    '---------------------2006-07-14 ADD By HJD------------------------------------------------------
    Dim strPrcCD As String
    
    If Row < 1 Or ss1.MaxRows < 1 Or opt_mltcd_change.Value = False Or Col <> 9 Or cbo_chg_prc_line.Text = "" Then Exit Sub
       
    With SS3
        .Row = Row
        .Col = Col: strPrcCD = .Text
        .Col = 8
        
        If strPrcCD Like "BC1*" Then      '.Text = IIf(strPrcCD Like "BC1*", "1", "2")
            .Text = "1"
        ElseIf strPrcCD Like "BC2*" Then
            .Text = "2"
        ElseIf strPrcCD Like "BC3*" Then
            .Text = "3"
        End If
        
        .Col = 10
        If Mid(.Text, InStr(1, .Text, "BF"), 3) <> strCCM_CD2 Then
            '.Text = Replace(.Text, Mid(.Text, InStr(1, .Text, "BF"), 3), strCCM_CD2, 1)
            
            If InStr(1, .Text, "BF1") <> 0 Then
                .Text = Replace(.Text, "BF1", "BF" & cbo_chg_prc_line.Text, 1)
            ElseIf InStr(1, .Text, "BF2") <> 0 Then
                .Text = Replace(.Text, "BF2", "BF" & cbo_chg_prc_line.Text, 1)
            ElseIf InStr(1, .Text, "BF3") <> 0 Then
                .Text = Replace(.Text, "BF3", "BF" & cbo_chg_prc_line.Text, 1)
            End If
        
        End If
    
    End With
    
  '----------------------------------------------------------------------------------------------------
  
End Sub

'Private Sub ss3_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'    If RowChanged Then
'        Call BOF_Changed(ss3)
'    End If
'End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss4_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = Col1 To Col2
    
        SS2.ColWidth(iCol) = ss4.ColWidth(iCol)
        ss6.ColWidth(iCol) = ss4.ColWidth(iCol)
        
    Next iCol

End Sub

Private Sub ss5_Click(ByVal Col As Long, ByVal Row As Long)

    Dim sTemp As String
    Dim iRow As Integer
    
    'Call Gp_Sp_Sort(Sc5.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If Row < 1 Then Exit Sub
    
    If ss5.MaxRows < 1 Then
        txt_heat_mana_no1.Text = ""
        Call Gf_Sp_Cls(Sc6)
        Exit Sub
    End If
    
    '----------------2006-07-14 ADD By HJD--------------------------------------------------------------------
    If lngCurRow <> Row Then
        ss5.Row = Row: ss5.Col = 10: strCCM_CD3 = Trim(ss5.Text)
        'strCCM_CD3 = Mid(ss3.Text, InStr(1, ss3.Text, "BF"), 3)
        lngCurRow = Row
    End If
    
    If opt_line_change.Value = True And Col = 8 And cbo_chg_prc_line.Text <> "" Then
        
        With ss5
            .Row = Row
            .Col = 8
            .Text = cbo_chg_prc_line.Text     'IIf(.Text = "1", "2", "1")
            .Col = 10:
            
            '.Text = Replace(.Text, IIf(.Text Like "BC1*", "BC1", "BC2"), IIf(.Text Like "BC1*", "BC2", "BC1"), 1)
                
            If .Text Like "BC1*" Then
                .Text = Replace(.Text, "BC1", "BC" & cbo_chg_prc_line.Text, 1)
            ElseIf .Text Like "BC2*" Then
                .Text = Replace(.Text, "BC2", "BC" & cbo_chg_prc_line.Text, 1)
            Else
                .Text = Replace(.Text, "BC3", "BC" & cbo_chg_prc_line.Text, 1)
            End If
            
        End With
        
        Call BOF_Changed(ss5)
        
    End If
    
'    If opt_mltcd_change.Value = True Then
'        ss3.Col = 0: ss3.Row = Row
'        If ss3.Text <> "Selected" Then
'            ss3.Text = "Selected"
'            Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, Row, Row, , &HFFFF80)
'        Else
'            ss3.Text = ""
'            Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, Row, Row)
'        End If
'
'    End If

    
    '-----------------------------------------------------------------------------------------------------------
    If opt_charge.Value = True Or opt_cast.Value = True Then Exit Sub
    
    ss5.Row = Row
    If Gf_Sp_ProceExist(Sc6.Item("Spread")) Then Exit Sub
    
    ss5.Col = 1
    txt_heat_mana_no1.Text = ss5.Text
    
    Call Gf_Sp_Refer(M_CN1, Sc6, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    'Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc")("Spread"))
    
    For iRow = 1 To ss6.MaxRows
                
        ss6.Row = iRow
        ss6.Col = 1
        
        If iRow = 1 Then sTemp = Mid(ss6.Text, 1, 8)
        
        If sTemp <> Mid(ss6.Text, 1, 8) Then
            sTemp = Mid(ss6.Text, 1, 8)
            Call Gp_Sp_BlockColor(ss6, 1, ss6.MaxCols, iRow, iRow, , &HFFC0FF)
        End If
    
    Next iRow

End Sub

Private Sub ss5_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = Col1 To Col2
    
        ss1.ColWidth(iCol) = ss5.ColWidth(iCol)
        SS3.ColWidth(iCol) = ss5.ColWidth(iCol)
        
    Next iCol

End Sub

Private Sub ss5_EditChange(ByVal Col As Long, ByVal Row As Long)

    '---------------------2006-07-14 ADD By HJD------------------------------------------------------
    Dim strPrcCD As String
    
    If Row < 1 Or ss1.MaxRows < 1 Or opt_mltcd_change.Value = False Or Col <> 9 Or cbo_chg_prc_line.Text = "" Then Exit Sub
       
    With ss5
        .Row = Row
        .Col = Col: strPrcCD = .Text
        .Col = 8
        
        If strPrcCD Like "BC1*" Then      '.Text = IIf(strPrcCD Like "BC1*", "1", "2")
            .Text = "1"
        ElseIf strPrcCD Like "BC2*" Then
            .Text = "2"
        ElseIf strPrcCD Like "BC3*" Then
            .Text = "3"
        End If
        
        .Col = 10
        If Mid(.Text, InStr(1, .Text, "BF"), 3) <> strCCM_CD3 Then
            '.Text = Replace(.Text, Mid(.Text, InStr(1, .Text, "BF"), 3), strCCM_CD3, 1)
            If InStr(1, .Text, "BF1") <> 0 Then
                .Text = Replace(.Text, "BF1", "BF" & cbo_chg_prc_line.Text, 1)
            ElseIf InStr(1, .Text, "BF2") <> 0 Then
                .Text = Replace(.Text, "BF2", "BF" & cbo_chg_prc_line.Text, 1)
            ElseIf InStr(1, .Text, "BF3") <> 0 Then
                .Text = Replace(.Text, "BF3", "BF" & cbo_chg_prc_line.Text, 1)
            End If
        End If
    End With
    
  '----------------------------------------------------------------------------------------------------

End Sub

Private Sub ss6_Click(ByVal Col As Long, ByVal Row As Long)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss6_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = Col1 To Col2
    
        SS2.ColWidth(iCol) = ss6.ColWidth(iCol)
        ss4.ColWidth(iCol) = ss6.ColWidth(iCol)
        
    Next iCol

End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
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

Private Sub opt_search_Click(Value As Integer)

    If opt_search.Value Then
        P_UNIT = ""
        Call Prod_Button_Edit
    End If
    
End Sub

Private Sub opt_charge_Click(Value As Integer)

    If opt_charge.Value Then
        P_UNIT = "H"
        Call Prod_Button_Edit
    End If
    
End Sub

Private Sub opt_cast_Click(Value As Integer)

    If opt_cast.Value Then
        P_UNIT = "C"
        Call Prod_Button_Edit
    End If
    
End Sub

Private Sub opt_slab_Click(Value As Integer)

    If opt_slab.Value Then
        P_UNIT = "S"
        Call Prod_Button_Edit
    End If
    
End Sub

Private Sub Prod_Button_Edit()

    sLoc = ""
    opt_search.ForeColor = &H808080
    opt_charge.ForeColor = &H808080
    opt_cast.ForeColor = &H808080
    opt_slab.ForeColor = &H808080
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    opt_move.Enabled = True
    opt_unification.Enabled = True
    opt_split.Enabled = True
    opt_delete.Enabled = True
    cmd_OK.Enabled = True
    
    Select Case P_UNIT
        Case "H"    'Charge
            opt_charge.ForeColor = &HFF&
            opt_unification.Enabled = False
            opt_split.Enabled = False
            Call chk_ss1_UnChecked
        Case "C"    'Cast
            opt_cast.ForeColor = &HFF&
            opt_delete.Enabled = False
            Call chk_ss1_UnChecked
        Case "S"    'Slab
            opt_slab.ForeColor = &HFF&
            opt_unification.Enabled = False
            opt_split.Enabled = False
'            opt_delete.Enabled = False
            Call chk_ss1_UnChecked
        Case Else   'Search
            opt_search.ForeColor = &HFF&
            txt_from.Text = ""
            txt_to.Text = ""
            txt_target.Text = ""
            opt_move.Enabled = False
            opt_unification.Enabled = False
            opt_split.Enabled = False
            opt_delete.Enabled = False
            cmd_OK.Enabled = False
    End Select
    opt_move.Value = True
    
End Sub

Private Sub opt_move_Click(Value As Integer)

    If opt_move.Value = True Then
        P_MODE = "M"
        Call Process_Button_Edit
    Else
        opt_move.ForeColor = &H808080
    End If

End Sub

Private Sub opt_unification_Click(Value As Integer)

    If opt_unification.Value = True Then
        P_MODE = "U"
        Call Process_Button_Edit
    Else
        opt_unification.ForeColor = &H808080
    End If

End Sub

Private Sub opt_split_Click(Value As Integer)

    If opt_split.Value = True Then
        P_MODE = "S"
        Call Process_Button_Edit
    Else
        opt_split.ForeColor = &H808080
    End If

End Sub

Private Sub opt_delete_Click(Value As Integer)

    If opt_delete.Value = True Then
        P_MODE = "D"
        Call Process_Button_Edit
    Else
        opt_move.ForeColor = &H808080
    End If
    
End Sub

Private Sub Process_Button_Edit()

    sLoc = ""
    opt_move.ForeColor = &H808080
    opt_unification.ForeColor = &H808080
    opt_split.ForeColor = &H808080
    opt_delete.ForeColor = &H808080
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    txt_to.Enabled = False
    txt_target.Enabled = False
    
    opt_bottom.Enabled = True
    opt_top.Enabled = True
    opt_top.Value = True
    
    Select Case P_MODE
        Case "M"    'Move
            opt_move.ForeColor = &HFF&
            txt_to.Enabled = True
            txt_target.Enabled = True
        Case "U"    'Unification
            opt_unification.ForeColor = &HFF&
            txt_target.Enabled = True
        Case "S"    'Split
            opt_split.ForeColor = &HFF&
            txt_target.Enabled = True
        Case "D"   'Delete
            opt_delete.ForeColor = &HFF&
            txt_to.Enabled = True
            opt_top.Enabled = False
            opt_bottom.Enabled = False
    End Select
    
    Call txt_from_Click
    
End Sub

Private Sub opt_top_Click(Value As Integer)

    If opt_top.Value = True Then
        opt_top.ForeColor = &HFF&
        opt_bottom.ForeColor = &H808080
        P_POSITION = "T"
    Else
        opt_top.ForeColor = &H808080
        P_POSITION = "B"
    End If

End Sub

Private Sub opt_bottom_Click(Value As Integer)

    If opt_bottom.Value = True Then
        opt_bottom.ForeColor = &HFF&
        opt_top.ForeColor = &H808080
        P_POSITION = "B"
    Else
        opt_bottom.ForeColor = &H808080
        P_POSITION = "T"
    End If
    
End Sub

Private Sub txt_from_Click()

    sLoc = "F"
    txt_from.BackColor = &HC0FFFF
    txt_to.BackColor = &H80000005
    txt_target.BackColor = &H80000005
    
End Sub



Private Sub txt_to_Click()

    sLoc = "T"
    txt_to.BackColor = &HC0FFFF
    txt_from.BackColor = &H80000005
    txt_target.BackColor = &H80000005
    
End Sub

Private Sub txt_target_Click()

    sLoc = "A"
    txt_target.BackColor = &HC0FFFF
    txt_from.BackColor = &H80000005
    txt_to.BackColor = &H80000005
    
End Sub

Private Sub txt_from_Change()

    If Trim(txt_to.Text) = "" And txt_to.Enabled = False Then
        txt_to.Text = txt_from.Text
    End If
    
End Sub

Private Sub txt_to_Change()

    Dim sTemp As String
    
    If Trim(txt_to.Text) = "" Then Exit Sub
    
    If Trim(txt_from.Text) > Trim(txt_to.Text) Then
        sTemp = Trim(txt_to.Text)
        txt_to.Text = Trim(txt_from.Text)
        txt_from.Text = sTemp
    End If
    
End Sub

Private Sub txt_target_Change()

    If Trim(txt_target.Text) = "" Then Exit Sub
    
    If Trim(txt_from.Text) <= Trim(txt_target.Text) And _
       Trim(txt_target.Text) <= Trim(txt_to.Text) And _
       opt_split.Value = False And opt_unification.Value = False Then
        Call Gp_MsgBoxDisplay("Value of Target item is between from and to..")
        txt_target.Text = ""
    End If
    
End Sub

Private Sub Cmd_Ok_Click()
    
    If txt_from.Text = "" Or txt_to.Text = "" Or txt_target.Text = "" Then
        If P_MODE = "D" Then
            If txt_from.Text = "" Or txt_to.Text = "" Then
                Call Gp_MsgBoxDisplay("Must input Value of From, To item")
            End If
        Else
            Call Gp_MsgBoxDisplay("Must input From, To, Value of Target item")
            Exit Sub
        End If
    End If
    
    If Trim(txt_from.Text) <= Trim(txt_target.Text) And _
       Trim(txt_target.Text) <= Trim(txt_to.Text) And _
       opt_split.Value = False And opt_unification.Value = False Then
        Call Gp_MsgBoxDisplay("Value of Target item is between from and to..")
        Exit Sub
    End If
    
'    If P_UNIT = "S" Then
'            If txt_prc_line1.Text = "4" Then
'                sdb_prc_line = "1"
'            ElseIf txt_prc_line1.Text = "5" Then
'                sdb_prc_line = "2"
'            ElseIf txt_prc_line1.Text = "6" Then
'                sdb_prc_line = "3"
'            End If
'    End If
    
    
    Call Gp_Process_Exec
    
End Sub

Public Sub Gp_Process_Exec()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As adodb.Command
    
    P_PLT = txt_plt.Text
    P_STATUS = "D"
    P_LINE = sdb_prc_line.Value
        
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEZ5000P ('" + P_PLT + "','" + Trim(str(P_LINE)) + "','" + P_STATUS + "','" + P_MODE + "','" + P_UNIT + "','" + Trim(txt_from.Text) + "','"
    sQuery = sQuery + Trim(txt_to.Text) + "','" + Trim(txt_target.Text) + "','" + P_POSITION + "','" + sUserID + "',?)}"
    
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
    Else
        If P_UNIT = "S" Then
            Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
            Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc")("Spread"))
        Else
            Call Form_Ref
        End If
        txt_from.BackColor = &HC0FFFF
        txt_to.BackColor = &H80000005
        txt_target.BackColor = &H80000005
        txt_from.Text = ""
        txt_to.Text = ""
        txt_target.Text = ""
        sLoc = "F"
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

'------------------------------------------------------------------------------
'Add by HJD
'Date:2006-07-14
'Content:Change BOF NO. & Way of Process
'------------------------------------------------------------------------------
Private Sub opt_line_change_Click(Value As Integer)

    Dim iRow        As Integer
    
    If Sf_Sp_ProceExist(ss1) > 0 Or Sf_Sp_ProceExist(SS3) > 0 Or Sf_Sp_ProceExist(ss5) > 0 Then Call Form_Ref: opt_line_change.Value = True
    
    'MDIMain.MenuTool.Buttons(4).Enabled = True
    'MltCD_Changed.Enabled = opt_mltcd_change.Value
    MltCD_Changed.Enabled = opt_line_change.Value
    MltCD_Changed.Caption = "切换工作炉座号"
    txt_from.Enabled = False: txt_to.Enabled = False: txt_target.Enabled = False
    
    With ss1
    
        For iRow = 1 To .MaxRows
            .BlockMode = True
            .Col = 8:    .Col2 = 8
            .Row = iRow: .ROW2 = iRow
            .BackColor = &HC0FFEE
            .Lock = True
            .BlockMode = False
        Next
        
    End With
    
    With SS3
    
        For iRow = 1 To .MaxRows
            .BlockMode = True
            .Col = 8:    .Col2 = 8
            .Row = iRow: .ROW2 = iRow
            .BackColor = &HC0FFEE
            .Lock = True
            .BlockMode = False
        Next
        
    End With
            
    With ss5
    
        For iRow = 1 To .MaxRows
            .BlockMode = True
            .Col = 8:    .Col2 = 8
            .Row = iRow: .ROW2 = iRow
            .BackColor = &HC0FFEE
            .Lock = True
            .BlockMode = False
        Next
        
    End With
            
End Sub

Private Sub opt_mltcd_change_Click(Value As Integer)
    
     MltCD_Changed.Enabled = opt_mltcd_change.Value
     MltCD_Changed.Caption = "变更工艺路径"
     txt_from.Enabled = False: txt_to.Enabled = False: txt_target.Enabled = False
    'MDIMain.MenuTool.Buttons(4).Enabled = True
    
End Sub

Public Function Sf_Sp_ProceExist(ByVal sPname As Variant) As Integer

    Dim iRow        As Integer
    Dim sColor      As String

    Sf_Sp_ProceExist = 0

    With sPname
    
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 0
            If Trim(.Text) = "Input" Or Trim(.Text) = "Update" Or Trim(.Text) = "Delete" Then
                Sf_Sp_ProceExist = 1
            End If

            .Col = 3
             sColor = .BackColor

             .Col = 8:   .Col2 = 8
             .BackColor = sColor

             .Col = 10:   .Col2 = 10
             .BackColor = sColor
             .Lock = True
        Next
        
    End With

    MDIMain.MenuTool.Buttons(9).Enabled = False
    
End Function

Private Sub Chk_ss1_Click(Value As Integer)
      
    If Chk_ss1.Value = -1 Then
        Frame1.Enabled = True
        cbo_chg_prc_line.Enabled = False
        opt_mltcd_change.ForeColor = &HFF&
        opt_line_change.ForeColor = &HFF&
        opt_mltcd_change.Enabled = True
        opt_line_change.Enabled = True
        opt_search.Value = True
        cbo_chg_prc_line.Enabled = True
        txt_from.Enabled = False: txt_to.Enabled = False: txt_target.Enabled = False
    Else
        'Chk_ss1.Value = ssCBUnchecked
        Frame1.Enabled = False
        opt_mltcd_change.ForeColor = &H808080
        opt_line_change.ForeColor = &H808080
        opt_mltcd_change.Value = False
        opt_line_change.Value = False
        opt_mltcd_change.Enabled = False
        opt_line_change.Enabled = False
        cbo_chg_prc_line.Enabled = False
        cbo_chg_prc_line.Text = ""
        MltCD_Changed.Enabled = opt_line_change.Value
        'MDIMain.MenuTool.Buttons(4).Enabled = False
        Call Form_Ref
    End If
    
End Sub

Private Sub chk_ss1_UnChecked()

    Chk_ss1.Value = ssCBUnchecked
    Frame1.Enabled = False
    opt_mltcd_change.ForeColor = &H808080
    opt_line_change.ForeColor = &H808080
    opt_mltcd_change.Value = False
    opt_line_change.Value = False
    MltCD_Changed.Enabled = opt_mltcd_change.Value
    MltCD_Changed.Caption = "变更工艺路径"
    'MDIMain.MenuTool.Buttons(4).Enabled = False
    
End Sub

Private Sub Prc_Changed()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As adodb.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEC1070P (?)}"
    
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("工艺路径更新完毕!!", "I")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub BOF_Changed(ByVal sPname As Variant)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As adodb.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "E_CODE"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 2
    
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    sQuery = "{call AEC1070C.P_SMODIFY (  'U', "
    
    With sPname
        .Row = lngCurRow
        .Col = 1: sQuery = sQuery + "'" + Trim(.Text) + "',"
        
        .Col = 2: sQuery = sQuery + "'" + Trim(.Text) + "',"
        
        .Col = 3: sQuery = sQuery + "'" + Trim(.Text) + "',"
        
        .Col = 8: sQuery = sQuery + "'" + Trim(.Text) + "',"
        
        .Col = 10: sQuery = sQuery + "'" + Trim(.Text) + "',?,?)}"
    End With
    
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
    If Trim(adoCmd("E_CODE")) <> "0" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        'Call Gp_MsgBoxDisplay("工艺路径更新完毕!!", "I")
        Call Form_Ref
        Call Sp_Col_Set
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub Bof_Insert()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As adodb.Command
    
    'If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEC1070C.P_BOF_INSERT ( '" + Trim(txt_from.Text) + "','" + Trim(txt_to.Text) + "',?)}"
  
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay("Error Mesg : " & ret_Result_ErrMsg)
        Set adoCmd = Nothing
        Exit Sub
    Else
        Call Gp_MsgBoxDisplay("插入工作炉座号更新完毕!!", "I")
        Call Form_Ref
        Call Sp_Col_Set
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub Bof_Switch()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim sTemp As String
    Dim adoCmd As adodb.Command
    
    If ss1.MaxRows = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    ss1.Row = ss1.ActiveRow: ss1.Col = 1
    
    sQuery = "SELECT HEAT_EDT_SEQ FROM EP_CHARGE_EDT WHERE HEAT_MANA_NO = '" + Trim(ss1.Text) + "'"
    sTemp = Gf_CodeFind(M_CN1, sQuery)
    If sTemp = "" Or sTemp = "FAIL" Then Exit Sub
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call AEZ6015P( '" + Trim(sTemp) + "','','','','',?)}"
    'sQuery = "{call AEZ6015P( '0','','','','',?)}"
  
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay("Error Mesg : " & ret_Result_ErrMsg)
        Set adoCmd = Nothing
        Exit Sub
    Else
        Call Gp_MsgBoxDisplay("炉座号更新完毕!!", "I")
        Call Form_Ref
        Call Sp_Col_Set
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
    
End Sub

Private Sub Sp_Col_Set()

    Dim iRow As Long
    
    If opt_line_change.Value = True Then
    
        With ss1
        
            For iRow = 1 To .MaxRows
                .BlockMode = True
                .Col = 8:    .Col2 = 8
                .Row = iRow: .ROW2 = iRow
                .BackColor = &HC0FFEE
                .Lock = True
                .BlockMode = False
            Next
            
        End With
        
        With SS3
        
            For iRow = 1 To .MaxRows
                .BlockMode = True
                .Col = 8:    .Col2 = 8
                .Row = iRow: .ROW2 = iRow
                .BackColor = &HC0FFEE
                .Lock = True
                .BlockMode = False
            Next
            
        End With
        
        With ss5
        
            For iRow = 1 To .MaxRows
                .BlockMode = True
                .Col = 8:    .Col2 = 8
                .Row = iRow: .ROW2 = iRow
                .BackColor = &HC0FFEE
                .Lock = True
                .BlockMode = False
            Next
            
        End With
        
    End If
   
End Sub

Private Sub Sp_Select(ByVal sPname As Variant)

    Dim I As Long
    
    With sPname
        
        If .MaxRows < 1 Then Exit Sub
        If .SelBlockRow < 1 Then Exit Sub
        
        For I = .SelBlockRow To .SelBlockRow2
            .Row = I
            .Col = 0
            
            If Trim(.Text) = "" Then
                .Text = "Selected"
            End If
        Next I
        
    End With
    
End Sub


Public Sub Form_Ins()
    
   Call Gp_Sp_Copy(Proc_Sc("Sc"))
   Call Gp_Sp_Paste(Proc_Sc("Sc"))
   
        intSeq = intSeq + 1
        
        SS2.Row = SS2.ActiveRow
        SS2.Col = 18
        SS2.Text = intSeq
        
'ss2 插入的板坯 必须同炉次 检查
' Dim iRow As Long
'
'
'   With ss2
'
'        For iRow = 1 To .MaxRows
'            .Col = 16
'            If .Text <> "" Then
'               If strSlabNo1 = "" Then
'                   .Col = 1
'                   strSlabNo1 = .Text
'               ElseIf strSlabNo2 = "" Then
'                   .Col = 1
'                   strSlabNo2 = .Text
'               ElseIf strSlabNo3 = "" Then
'                   .Col = 1
'                   strSlabNo3 = .Text
'               End If
'
'               If strSlabNo1 <> "" And strSlabNo2 <> "" And strSlabNo3 <> "" Then
'                    Call Gp_MsgBoxDisplay("最多增加三块余坯!!", "I")
'               End If
'              .Col = 16
'            End If
'        Next
'    End With
'
'    If strSlabNo1 <> "" And strSlabNo2 <> "" Then
'       If Left(strSlabNo1, 8) <> Left(strSlabNo2, 8) Then
'           Call Gp_MsgBoxDisplay("余坯追加必须同炉次!!", "I")
'       End If
'    End If
'
'
'    If strSlabNo1 <> "" And strSlabNo3 <> "" Then
'       If Left(strSlabNo1, 8) <> Left(strSlabNo3, 8) Then
'           Call Gp_MsgBoxDisplay("余坯追加必须同炉次!!", "I")
'       End If
'    End If
    
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 44)
    
End Sub

