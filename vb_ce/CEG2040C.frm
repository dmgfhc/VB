VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form CEG2040C 
   Caption         =   "加热炉均衡查询/调整_CEG2040C"
   ClientHeight    =   8400
   ClientLeft      =   195
   ClientTop       =   2250
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   15315
   WindowState     =   2  'Maximized
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
      Left            =   2010
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   60
      Width           =   2835
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
      Left            =   1515
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   60
      Width           =   465
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   210
      Top             =   60
      Width           =   1275
      _ExtentX        =   2249
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   675
      Left            =   1830
      Top             =   450
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   1191
      Caption         =   ""
      Alignment       =   1
      BackColor       =   12640511
      BackgroundStyle =   1
      BorderEffect    =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_time1 
      Height          =   315
      Left            =   9180
      TabIndex        =   4
      Top             =   420
      Width           =   1140
      _Version        =   262145
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   128
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt1 
      Height          =   315
      Left            =   7740
      TabIndex        =   3
      Top             =   420
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   5670
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "加热炉 #1"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_cnt1 
      Height          =   315
      Left            =   6990
      TabIndex        =   2
      Top             =   420
      Width           =   720
      _Version        =   262145
      _ExtentX        =   1270
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   16711680
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_time2 
      Height          =   315
      Left            =   9180
      TabIndex        =   5
      Top             =   780
      Width           =   1140
      _Version        =   262145
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   128
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt2 
      Height          =   315
      Left            =   7740
      TabIndex        =   6
      Top             =   780
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   1
      Left            =   5670
      Top             =   780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "加热炉 #2,3"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_cnt2 
      Height          =   315
      Left            =   6990
      TabIndex        =   7
      Top             =   780
      Width           =   720
      _Version        =   262145
      _ExtentX        =   1270
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   16711680
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_time3 
      Height          =   315
      Left            =   14040
      TabIndex        =   8
      Top             =   420
      Width           =   1140
      _Version        =   262145
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   128
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt3 
      Height          =   315
      Left            =   12600
      TabIndex        =   9
      Top             =   420
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   10530
      Top             =   420
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "共通"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_cnt3 
      Height          =   315
      Left            =   11850
      TabIndex        =   10
      Top             =   420
      Width           =   720
      _Version        =   262145
      _ExtentX        =   1270
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   16711680
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8010
      Left            =   60
      TabIndex        =   11
      Top             =   1170
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   14129
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "CEG2040C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   3885
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   15210
         _Version        =   393216
         _ExtentX        =   26829
         _ExtentY        =   6853
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
         MaxCols         =   0
         MaxRows         =   20
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         SpreadDesigner  =   "CEG2040C.frx":0052
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   4065
         Left            =   0
         TabIndex        =   12
         Top             =   3945
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   7170
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "CEG2040C.frx":0446
         Begin Threed.SSPanel SSPanel1 
            Height          =   570
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   15210
            _ExtentX        =   26829
            _ExtentY        =   1005
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_fur_line 
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
               Left            =   10200
               TabIndex        =   20
               Tag             =   "工厂"
               Top             =   120
               Visible         =   0   'False
               Width           =   465
            End
            Begin Threed.SSOption opt_fur_no1 
               Height          =   285
               Left            =   1830
               TabIndex        =   18
               Top             =   150
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "加热炉 #1"
               Value           =   -1
            End
            Begin VB.TextBox txt_stlgrd 
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
               Left            =   13470
               MaxLength       =   11
               TabIndex        =   15
               Top             =   120
               Visible         =   0   'False
               Width           =   1275
            End
            Begin CSTextLibCtl.sidbEdit sdb_slab_thk 
               Height          =   315
               Left            =   10680
               TabIndex        =   16
               Top             =   120
               Visible         =   0   'False
               Width           =   1365
               _Version        =   262145
               _ExtentX        =   2408
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               Text            =   " 0.00"
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
               NumIntDigits    =   4
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel6 
               Height          =   315
               Left            =   270
               Top             =   120
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               Caption         =   "变更加热炉"
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
               ForeColor       =   0
            End
            Begin Threed.SSOption opt_fur_no2 
               Height          =   285
               Left            =   3210
               TabIndex        =   19
               Top             =   150
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "加热炉 #2,3"
            End
            Begin CSTextLibCtl.sidbEdit sdb_slab_wid 
               Height          =   315
               Left            =   12060
               TabIndex        =   22
               Top             =   120
               Visible         =   0   'False
               Width           =   1365
               _Version        =   262145
               _ExtentX        =   2408
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               Text            =   " 0.00"
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
               NumIntDigits    =   4
               Undo            =   0
               Data            =   0
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   3465
            Left            =   0
            TabIndex        =   14
            Top             =   600
            Width           =   15210
            _Version        =   393216
            _ExtentX        =   26829
            _ExtentY        =   6112
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
            MaxCols         =   32
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CEG2040C.frx":0498
         End
      End
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   0
      Left            =   6990
      Top             =   60
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   556
      Caption         =   "块数"
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
      Index           =   2
      Left            =   7740
      Top             =   60
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      Caption         =   "重量"
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
      Index           =   3
      Left            =   9180
      Top             =   60
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "加热炉时间"
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   10530
      Top             =   780
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "轧制时间"
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
   Begin CSTextLibCtl.sidbEdit sdb_mill_time 
      Height          =   315
      Left            =   11850
      TabIndex        =   21
      Top             =   780
      Width           =   1140
      _Version        =   262145
      _ExtentX        =   2011
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   128
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSCommand cmd_slab_design_change 
      Height          =   405
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   720
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "设计标准变更"
      BevelWidth      =   3
   End
End
Attribute VB_Name = "CEG2040C"
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
'-- Program Name      FUR LINE CHANGE
'-- Program ID        CEG2040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2007.10.24
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

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Active_Row As Integer           'Active Row
Dim Active_Col As Integer           'Active Col

Private Sub Form_Define()

    Dim iCol As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_cnt1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_wgt1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_slab_time1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_cnt2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_wgt2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_slab_time2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_cnt3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_wgt3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(sdb_slab_time3, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_fur_line, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_slab_thk, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_slab_wid, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
         
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, False)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, False)
    
    For iCol = 5 To 32
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CEG2040C.P_REFER2", Key:="P-R"
    sc2.Add Item:="CEG2040C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:="CEG2040C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, SpreadHeader + (ss1.RowHeaderCols - 1), True)

End Sub

Public Sub Sp_Setting()

    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 3)) = 16
    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 2)) = 5
    ss1.MaxCols = 0

End Sub

Private Sub cmd_slab_design_change_Click()

    If ss1.MaxCols - 1 = Active_Col Or ss1.MaxCols - 2 = Active_Col Or _
       ss1.MaxCols - 3 = Active_Col Or ss1.MaxCols = Active_Col Then Exit Sub
    
    If Active_Col Mod 4 <> 2 And Active_Col Mod 4 <> 1 Then Exit Sub
    
    Load Slab_Design_Change

    If Active_Row > 0 And Active_Col > 0 Then
    
        ss1.Row = 0
        ss1.Col = Active_Col
        Slab_Design_Change.sdb_slab_thk_fr.Value = ss1.Text
        Slab_Design_Change.sdb_slab_thk_to.Value = ss1.Text
    
        ss1.Row = Active_Row
        ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 2)
        Slab_Design_Change.sdb_slab_wid_fr.Value = ss1.Value
        Slab_Design_Change.sdb_slab_wid_to.Value = ss1.Value
    
        ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 1)
        Slab_Design_Change.txt_stlgrd.Text = ss1.Text
        ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 3)
        Slab_Design_Change.txt_stlgrd_name.Text = ss1.Text
    
        Slab_Design_Change.sdb_slab_len_fr.Value = 0
        Slab_Design_Change.sdb_slab_len_to.Value = 9999999
        
        Slab_Design_Change.sdb_asroll_cnt_fr.Value = 0
        Slab_Design_Change.sdb_asroll_cnt_to.Value = 99
        
        Slab_Design_Change.sdb_prod_thk_fr.Value = 0
        Slab_Design_Change.sdb_prod_thk_to.Value = 9999.99
        
        Slab_Design_Change.sdb_prod_wid_fr.Value = 0
        Slab_Design_Change.sdb_prod_wid_to.Value = 9999.99
        
        Slab_Design_Change.sdb_prod_len_fr.Value = 0
        Slab_Design_Change.sdb_prod_len_to.Value = 9999999.9
        
        Slab_Design_Change.cbo_prod_cnt.Text = "0"
        
        Slab_Design_Change.txt_fur_line.Text = IIf(Active_Col Mod 4 = 1, "2", "1")
        
        If Active_Col Mod 4 = 1 Then
            Slab_Design_Change.opt_fur_no2.Value = True
            Slab_Design_Change.opt_fur_no1.Value = False
        Else
            Slab_Design_Change.opt_fur_no1.Value = True
            Slab_Design_Change.opt_fur_no2.Value = False
        End If
        
    End If
    
    Slab_Design_Change.P_MODE = "W"
    
    Slab_Design_Change.Show 1

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

    Dim sStatus As String
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    If Mid(sAuthority, 3, 1) <> "1" Then
        cmd_slab_design_change.Enabled = False
    End If
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    'Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Sp_Setting
    
    sStatus = Gf_CodeFind(M_CN1, "SELECT PRC_STS FROM ZP_JOB WHERE PLT = 'C3' AND MAIN_PGMID = 'CEG2000P'")
    If sStatus <> "C" Then
        ULabel1.Caption = "坯料使用计划进行中...!!"
    Else
        sStatus = Gf_CodeFind(M_CN1, "SELECT PRC_STS FROM ZP_JOB WHERE PLT = 'C3' AND MAIN_PGMID = 'CEH1000P'")
        If sStatus <> "C" Then
            ULabel1.Caption = "坯料分段作业指示进行中...!!"
            cmd_slab_design_change.Enabled = False
        Else
            sStatus = Gf_CodeFind(M_CN1, "SELECT PRC_STS FROM ZP_JOB WHERE PLT = 'C3' AND MAIN_PGMID = 'CED1000P'")
            If sStatus <> "C" Then
                ULabel1.Caption = "轧钢工序计划进行中...!!"
                cmd_slab_design_change.Enabled = False
            Else
                ULabel1.Caption = "没有可作业对象...!!"
                sAuthority = "0000"
                cmd_slab_design_change.Enabled = False
            End If
        End If
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet

    txt_plt.Text = "C3"
    Call txt_plt_KeyUp(0, 0)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "E-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "E-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Spl_SizeSet(SSSplitter1, "E-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
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
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
        Call Gf_Sp_Cls(sc2)
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        txt_plt.Text = "C3"
        Call txt_plt_KeyUp(0, 0)
        opt_fur_no1.Value = True
        ss1.MaxCols = 0
        Active_Row = 0
        Active_Col = 0
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Ref()

    Dim sQuery1 As String   'Header Display
    Dim sQuery2 As String   'Data Display
    Dim sMessg As String
    
    'Header Display
    sQuery1 = "SELECT  DISTINCT  SLAB_THK "
    sQuery1 = sQuery1 + "  FROM  EP_SLAB_EDT3 "
    sQuery1 = sQuery1 + " ORDER  BY SLAB_THK ASC "
    
    'Data Display
    sQuery2 = " {call CEG2040C.P_DATA ()} "
     
    sMessg = Gf_Ms_NeceCheck(nControl)
    If sMessg = "OK" Then
    
        sMessg = Gf_Ms_NeceCheck2(mControl)
        If sMessg = "OK" Then

            'Header Display
            Call Sp_Header_Refer1(ss1, sQuery1)      'Header Display
        
            'Data Display
            If Sp_Data_Refer1(ss1, sQuery2) Then     'SLAB Data Display
                ss1.OperationMode = OperationModeNormal
                opt_fur_no1.Value = True
                Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
                Active_Row = 0
                Active_Col = 0
                sdb_slab_cnt1.Value = Gf_FloatFind(M_CN1, "SELECT COUNT(SLAB_EDT_SEQ) FROM EP_SLAB_EDT3 WHERE FUR_LINE = '1' ")
                sdb_slab_cnt2.Value = Gf_FloatFind(M_CN1, "SELECT COUNT(SLAB_EDT_SEQ) FROM EP_SLAB_EDT3 WHERE FUR_LINE = '2' ")
                sdb_slab_cnt3.Value = Gf_FloatFind(M_CN1, "SELECT COUNT(SLAB_EDT_SEQ) FROM EP_SLAB_EDT3 WHERE FUR_LINE_AVA = '*' ")
                sdb_slab_wgt1.Value = Gf_FloatFind(M_CN1, "SELECT SUM(NVL(SLAB_WGT,0)) FROM EP_SLAB_EDT3 WHERE FUR_LINE = '1' ")
                sdb_slab_wgt2.Value = Gf_FloatFind(M_CN1, "SELECT SUM(NVL(SLAB_WGT,0)) FROM EP_SLAB_EDT3 WHERE FUR_LINE = '2' ")
                sdb_slab_wgt3.Value = Gf_FloatFind(M_CN1, "SELECT SUM(NVL(SLAB_WGT,0)) FROM EP_SLAB_EDT3 WHERE FUR_LINE_AVA = '*' ")
                sdb_slab_time1.Value = Gf_FloatFind(M_CN1, "SELECT SUM(FUR_TME) FROM EP_SLAB_EDT3 WHERE FUR_LINE = '1'")
                sdb_slab_time2.Value = Gf_FloatFind(M_CN1, "SELECT SUM(FUR_TME2) FROM EP_SLAB_EDT3 WHERE FUR_LINE = '2'")
                sdb_slab_time3.Value = Gf_FloatFind(M_CN1, "SELECT SUM(DECODE(FUR_LINE,'1',FUR_TME,'2',FUR_TME2)) FROM EP_SLAB_EDT3 WHERE FUR_LINE_AVA = '*'")
                sdb_mill_time.Value = Gf_FloatFind(M_CN1, "SELECT SUM(MILL_TME)/60 FROM EP_SLAB_EDT3 ")
            End If
            
        Else
            Call Gp_MsgBoxDisplay(Trim(sMessg) + "长度不正确", "I")
        End If
    
    Else
        Call Gp_MsgBoxDisplay(Trim(sMessg) + "必须输入", "I")
    End If

End Sub

Public Sub Form_Pro()
    
    If Gf_Sp_Process(M_CN1, sc2, Mc2) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Form_Ref
    End If
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, sc2)
      
End Sub

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

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_fur_no1_Click(Value As Integer)

    If opt_fur_no1.Value Then
        opt_fur_no1.ForeColor = &HFF&
        opt_fur_no2.ForeColor = &H80000012
    Else
        opt_fur_no1.ForeColor = &H80000012
    End If
    
End Sub

Private Sub opt_fur_no2_Click(Value As Integer)
    
    If opt_fur_no2.Value Then
        opt_fur_no2.ForeColor = &HFF&
        opt_fur_no1.ForeColor = &H80000012
    Else
        opt_fur_no2.ForeColor = &H80000012
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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If Row < 1 Or ss1.MaxRows = Row Then
        Active_Row = 0
        Active_Col = 0
        Exit Sub
    End If
    
    If ss1.MaxCols - 1 = Col Or ss1.MaxCols - 2 = Col Or ss1.MaxCols - 3 = Col Or ss1.MaxCols = Col Then
        Active_Row = 0
        Active_Col = 0
        Exit Sub
    End If
    
    If Col Mod 4 <> 2 And Col Mod 4 <> 1 Then
        Active_Row = 0
        Active_Col = 0
        Exit Sub
    End If
    
    Active_Row = Row
    Active_Col = Col
    
    txt_fur_line.Text = Trim(Str(Col Mod 4))
    ss1.Row = 0
    ss1.Col = Col
    sdb_slab_thk.Value = Val(ss1.Text)
    
    ss1.Row = Row
    ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 2)
    sdb_slab_wid.Value = ss1.Value
    
    ss1.Row = Row
    ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 1)
    
    txt_stlgrd.Text = ss1.Text
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MenuTool_ReSet
    If txt_fur_line.Text = "1" Then
        opt_fur_no2.Value = True
    Else
        opt_fur_no1.Value = True
    End If
    ss2.OperationMode = OperationModeNormal

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If Row < 1 Then Exit Sub
    
    ss2.Col = 0
    ss2.Row = Row
    
    If ss2.Text = "" Then
        ss2.Col = 2
        If ss2.Text <> "*" Then
            Call Gp_MsgBoxDisplay("不能变更加热炉", "I", Me.Caption)
            Exit Sub
        End If
        
        ss2.Col = 0:    ss2.Text = "Update"
        ss2.Col = 4
        If opt_fur_no1.Value Then
            ss2.Text = "1"
        Else
            ss2.Text = "2"
        End If
        
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row, , &HFFFF80)
    Else
        ss2.Col = 0:    ss2.Text = ""
        ss2.Col = 4:    ss2.Text = ""
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row)
    End If

End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If
    
End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If
    
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
        Exit Sub
        
    End If

    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_name.Text = ""
    End If

End Sub

Public Function Sp_Header_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Header_Refer1 = True
        
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer1 = False
            '.ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 4
            For iCol = 0 To .MaxCols - 1 Step 4
            
                For iColCnt = 1 To 4
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 2)
                    .Col = iCol + iColCnt
                    
                    If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(0, iCnt))
                    End If
                    
                    .ColWidth(iCol + iColCnt) = 6
    
                    .Col = iCol + iColCnt: .Col2 = iCol + iColCnt
                    .Row = 1: .Row2 = -1
                    .BlockMode = True
                    .CellType = 13      'SS_CELL_TYPE_NUMBER
                    .TypeNumberDecPlaces = 0
                    .TypeNumberMax = 999999999
                    .TypeNumberMin = 0
                    .TypeNumberShowSep = True
                    .TypeNumberLeadingZero = TypeLeadingZeroYes

                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
                    .BlockMode = False
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 1)
                    .Col = iCol + iColCnt
                    
                    Select Case iColCnt
                        Case 1
                            .Text = "#1"
                        Case 2
                            .Text = "#2,#3"
                        Case 3
                            .Text = "共通"
                        Case 4
                            .Text = "合计"
                    End Select
                    
                    If iColCnt = 4 Then
                        Call Gp_Sp_ColHidden(ss1, .Col, True)
                    End If
                    
                Next iColCnt
                
                iCnt = iCnt + 1
                
            Next iCol
            
            '合计 Col
            For iColCnt = 1 To 4
                
                .MaxCols = .MaxCols + 1
                .Col = .MaxCols
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                .Text = "合计(t)"
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                
                Select Case iColCnt
                    Case 1
                        .Text = "#1"
                    Case 2
                        .Text = "#2,#3"
                    Case 3
                        .Text = "共通"
                    Case 4
                        .Text = "合计"
                End Select
                    
                .ColWidth(.Col) = 8
                    
                .Col = .MaxCols: .Col2 = .MaxCols
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 999999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .TypeVAlign = TypeVAlignCenter
                .BlockMode = False
                
            Next iColCnt
            
        End If
        
        .BlockMode = True
        .Col = .MaxCols:  .Col2 = .MaxCols
        .Row = 1: .Row2 = -1
        .ForeColor = &HFF&  '&H00FF0000&
        .BlockMode = False
        
        For iColCnt = 4 To .MaxCols - 4 Step 4
            .BlockMode = True
            .Col = iColCnt:  .Col2 = iColCnt
            .Row = 1: .Row2 = -1
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iColCnt
        
        .BlockMode = True
        .Row = 0
        .Col = 1
        .Row2 = 0
        .Col2 = -1
        .RowMerge = MergeAlways
        '.ColMerge = MergeAlways
        .BlockMode = False
        
        .ReDraw = True
        .Refresh
        
        Screen.MousePointer = vbDefault
        
    End With
        
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim iBas As Integer
    Dim iCot As Integer
    
    Dim sCol_a As String
    Dim sCol_b As String
    Dim sRText As String
    Dim sStlgrd As String
    
    Dim ColSum As Double
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Data_Refer1 = True
        .ReDraw = False
        .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer1 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)

                If iCnt = 0 Or sRText <> Trim(ArrayRecords(2, iCnt)) Or sStlgrd <> Trim(ArrayRecords(1, iCnt)) Then
                    sRText = ArrayRecords(2, iCnt)
                    sStlgrd = ArrayRecords(1, iCnt)
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = SpreadHeader + (.RowHeaderCols - 3)
                    .Text = Trim(ArrayRecords(0, iCnt))
                    .Col = SpreadHeader + (.RowHeaderCols - 2)
                    .Text = Trim(ArrayRecords(2, iCnt))
                    .Col = SpreadHeader + (.RowHeaderCols - 1)
                    .Text = Trim(ArrayRecords(1, iCnt))
                    .Col = .MaxCols
                    .Text = sRText
                End If

                .Row = SpreadHeader + (.ColHeaderRows - 2)
                
                For iCol = 1 To .MaxCols Step 4
                
                    .Col = iCol
                    
                    If .Text = Trim(ArrayRecords(3, iCnt)) Then

                        .Row = .MaxRows
                        
                        If VarType(ArrayRecords(4, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(4, iCnt)) <> 0 Then
                                .Text = Trim(ArrayRecords(4, iCnt))
                            Else
                                .Text = ""
                            End If
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(5, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(5, iCnt)) <> 0 Then
                                .Text = Trim(ArrayRecords(5, iCnt))
                            Else
                                .Text = ""
                            End If
                        End If
                        
                        .Col = iCol + 2
                        If VarType(ArrayRecords(6, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(6, iCnt)) <> 0 Then
                                .Text = Trim(ArrayRecords(6, iCnt))
                            Else
                                .Text = ""
                            End If
                        End If
                        
                        .Col = iCol + 3
                        If VarType(ArrayRecords(7, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(7, iCnt)) <> 0 Then
                                .Text = Trim(ArrayRecords(7, iCnt))
                            Else
                                .Text = ""
                            End If
                        End If
                        
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 0
        .Text = "合计(t)"
        
        Call Gp_Sp_EvenRowBackcolor(sPname, 1)
        
        .BlockMode = True
        .Row = .MaxRows:  .Row2 = .MaxRows
        .Col = 1: .Col2 = -1
        .ForeColor = &HFF&
        .BlockMode = False
        
        For iCol = 4 To .MaxCols - 4 Step 4
            .BlockMode = True
            .Col = iCol:  .Col2 = iCol
            .Row = .MaxRows: .Row2 = .MaxRows
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iCol
        
        'Column Sum
        For iCol = 1 To .MaxCols
        
            .Col = iCol
            
            If .Col <= 26 Then
                sCol_a = Chr(.Col + 64)
                .Formula = "sum(" + sCol_a + "1:" + sCol_a & .MaxRows - 1 & ")"
            Else
                iCot = Int(((.Col - 1) / 26))
                iBas = 26 * iCot
                sCol_a = Chr((.Col - iBas) + 64)
                sCol_b = Chr(iCot + 64)
                .Formula = "sum(" + sCol_b + sCol_a + "1:" + sCol_b + sCol_a & .MaxRows - 1 & ")"
            End If
            
        Next iCol
            
        'Row Sum
        For iRow = 1 To .MaxRows
        
            .Row = iRow
            
            ColSum = 0
            For iCol = 1 To .MaxCols - 4 Step 4
            
                .Col = iCol
                If .Text <> "" Then
                    ColSum = ColSum + .Value
                End If
                
            Next iCol
            
            .Col = .MaxCols - 3
            
            If ColSum <> 0 Then
                .Value = ColSum
            End If
            
            ColSum = 0
            For iCol = 2 To .MaxCols - 4 Step 4
            
                .Col = iCol
                If .Text <> "" Then
                    ColSum = ColSum + .Value
                End If
                
            Next iCol
            
            .Col = .MaxCols - 2
            If ColSum <> 0 Then
                .Value = ColSum
            End If
            
            ColSum = 0
            For iCol = 3 To .MaxCols - 4 Step 4
            
                .Col = iCol
                If .Text <> "" Then
                    ColSum = ColSum + .Value
                End If
                
            Next iCol
            
            .Col = .MaxCols - 1
            If ColSum <> 0 Then
                .Value = ColSum
            End If
            
            ColSum = 0
            For iCol = 4 To .MaxCols - 4 Step 4
            
                .Col = iCol
                If .Text <> "" Then
                    ColSum = ColSum + .Value
                End If
                
            Next iCol
            
            .Col = .MaxCols
            If ColSum <> 0 Then
                .Value = ColSum
            End If
            
        Next iRow
        
        .ReDraw = True
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub
