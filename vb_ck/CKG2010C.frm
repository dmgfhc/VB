VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form CKG2010C 
   Caption         =   "指示调整_CKG2010C"
   ClientHeight    =   9615
   ClientLeft      =   420
   ClientTop       =   1665
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel1 
      Height          =   435
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   767
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_move 
         Height          =   330
         Left            =   2085
         TabIndex        =   14
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
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
         Caption         =   "调 整"
      End
      Begin Threed.SSOption opt_delete 
         Height          =   330
         Left            =   3060
         TabIndex        =   15
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "删 除"
      End
      Begin Threed.SSOption opt_sent 
         Height          =   330
         Left            =   135
         TabIndex        =   16
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "发 送"
      End
      Begin Threed.SSOption opt_cancel 
         Height          =   330
         Left            =   1110
         TabIndex        =   17
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "取 消"
      End
      Begin Threed.SSOption opt_return 
         Height          =   330
         Left            =   4005
         TabIndex        =   31
         Top             =   60
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "返 送"
      End
   End
   Begin VB.TextBox txt_search_slabno 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   315
      Left            =   11970
      MaxLength       =   10
      TabIndex        =   30
      ToolTipText     =   "回车检索"
      Top             =   540
      Width           =   1230
   End
   Begin VB.TextBox target_y 
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
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   29
      Tag             =   "工厂"
      Top             =   9300
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox to_y 
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
      Left            =   1590
      MaxLength       =   2
      TabIndex        =   28
      Tag             =   "工厂"
      Top             =   9300
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox from_y 
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
      Left            =   570
      MaxLength       =   2
      TabIndex        =   27
      Tag             =   "工厂"
      Top             =   9300
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txt_to 
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
      Left            =   7995
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   540
      Width           =   1230
   End
   Begin VB.TextBox txt_target 
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
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   540
      Width           =   1230
   End
   Begin VB.TextBox txt_from 
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
      Left            =   6375
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   540
      Width           =   1230
   End
   Begin VB.TextBox TXT_PLT 
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
      Left            =   1500
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   540
      Width           =   540
   End
   Begin VB.TextBox TXT_PLT_NAME 
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
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   540
      Width           =   2880
   End
   Begin Threed.SSPanel SSPsend 
      Height          =   315
      Left            =   13245
      TabIndex        =   2
      Top             =   540
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "已下达"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPpdt 
      Height          =   315
      Left            =   14235
      TabIndex        =   3
      Top             =   540
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "生产中"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8370
      Left            =   60
      TabIndex        =   4
      Top             =   915
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   14764
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
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
      TabCaption(0)   =   "轧制指示"
      TabPicture(0)   =   "CKG2010C.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSSplitter1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "母板、钢板指示"
      TabPicture(1)   =   "CKG2010C.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSSplitter2"
      Tab(1).ControlCount=   1
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   7995
         Left            =   -75000
         TabIndex        =   5
         Top             =   300
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   14102
         _Version        =   196609
         SplitterBarWidth=   4
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   1
         BackColor       =   16761087
         PaneTree        =   "CKG2010C.frx":0038
         Begin FPSpread.vaSpread ss3 
            Height          =   7965
            Left            =   15
            TabIndex        =   6
            Top             =   15
            Width           =   7545
            _Version        =   393216
            _ExtentX        =   13309
            _ExtentY        =   14049
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
            SpreadDesigner  =   "CKG2010C.frx":008A
         End
         Begin FPSpread.vaSpread ss4 
            Height          =   7965
            Left            =   7620
            TabIndex        =   7
            Top             =   15
            Width           =   7500
            _Version        =   393216
            _ExtentX        =   13229
            _ExtentY        =   14049
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
            SpreadDesigner  =   "CKG2010C.frx":0658
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   7995
         Left            =   0
         TabIndex        =   8
         Top             =   300
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   14102
         _Version        =   196609
         BorderStyle     =   1
         PaneTree        =   "CKG2010C.frx":0B25
         Begin FPSpread.vaSpread ss1 
            Height          =   7965
            Left            =   15
            TabIndex        =   9
            Top             =   15
            Width           =   15105
            _Version        =   393216
            _ExtentX        =   26644
            _ExtentY        =   14049
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
<<<<<<< .mine
            MaxCols         =   61
=======
            MaxCols         =   66
>>>>>>> .r7591
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CKG2010C.frx":0B57
         End
      End
   End
   Begin CSTextLibCtl.sidbEdit SDB_SLAB_EDT_SEQ 
      Height          =   315
      Left            =   3090
      TabIndex        =   10
      Tag             =   "炉次编制号"
      Top             =   540
      Visible         =   0   'False
      Width           =   375
      _Version        =   262145
      _ExtentX        =   661
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
      RawData         =   ""
      Text            =   " 0"
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
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_PRC_LINE 
      Height          =   315
      Left            =   3510
      TabIndex        =   11
      Top             =   540
      Visible         =   0   'False
      Width           =   180
      _Version        =   262145
      _ExtentX        =   317
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
      Left            =   60
      Top             =   540
      Width           =   1410
      _ExtentX        =   2487
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   4980
      Top             =   540
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "起始板坯号"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   9285
      Top             =   540
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "目标板坯号"
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
      Left            =   7605
      Top             =   540
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Caption         =   "->"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   435
      Left            =   5010
      TabIndex        =   21
      Top             =   60
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   767
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_target 
         Height          =   330
         Left            =   5640
         TabIndex        =   22
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
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
         Caption         =   "目标板坯号"
      End
      Begin Threed.SSOption opt_from 
         Height          =   330
         Left            =   1320
         TabIndex        =   23
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "起始板坯号"
      End
      Begin Threed.SSOption opt_to 
         Height          =   330
         Left            =   2970
         TabIndex        =   24
         Top             =   60
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
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
         Caption         =   "->"
      End
      Begin Threed.SSCommand cmd_roll_mana 
         Height          =   375
         Left            =   9030
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   30
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   661
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
         Caption         =   "辊期编制"
      End
      Begin Threed.SSCommand cmd_abnormal_send 
         Height          =   375
         Left            =   7830
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   30
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   661
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
         Caption         =   "强制发送"
      End
   End
   Begin VB.TextBox TXT_MPLATE_NO 
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
      Left            =   10155
      MaxLength       =   12
      TabIndex        =   12
      Tag             =   "炉次管理号"
      Top             =   75
      Visible         =   0   'False
      Width           =   1395
   End
   Begin Threed.SSPanel SSOrd 
      Height          =   315
      Left            =   15270
      TabIndex        =   32
      Top             =   540
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   0
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "多订单"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin CSTextLibCtl.sidbEdit SDB_MPLATE_EDT_SEQ 
      Height          =   315
      Left            =   90
      TabIndex        =   33
      Tag             =   "板坯编制号"
      Top             =   540
      Visible         =   0   'False
      Width           =   375
      _Version        =   262145
      _ExtentX        =   661
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
      RawData         =   ""
      Text            =   " 0"
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
      NumIntDigits    =   8
      Undo            =   0
      Data            =   0
   End
End
Attribute VB_Name = "CKG2010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      指示调整
'-- Program ID        CKG2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2007.11.19
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
Dim Mode As String

'Public Complete As Boolean           'Move Status Setting

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

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

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
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection

Dim Sc1 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_PRC_STS = 5
Const SS1_SLAB_NO = 6
Const SS1_OVER_FL = 11
Const SS1_PROD_DATE = 18
Const SS1_ORD_NO = 20
Const SS1_HCR_FL = 25
Const SS1_RH = 40 '20130326 LICHAO
Const SS1_L2_SEND = 48
Const SS1_ORD_CNT = 53
Const SS1_SLAB_EDT_SEQ = 55
Const SS1_URGNT_FL = 56  '紧急订单
Const SS3_MP_NO = 2
Const SS3_SLAB_EDT_SEQ = 8
Const SS3_MPLATE_EDT_SEQ = 9

Dim sSlab_Edt_Seq_Fr As String
Dim sSlab_Edt_Seq_To As String
Dim sSlab_Edt_Seq_Tg As String

Private Sub Form_Define()
        
    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

             Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(SDB_SLAB_EDT_SEQ, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    Call Gp_Ms_Collection(SDB_SLAB_EDT_SEQ, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
    
      Call Gp_Ms_Collection(SDB_SLAB_EDT_SEQ, "p", " ", " ", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    Call Gp_Ms_Collection(SDB_MPLATE_EDT_SEQ, "p", " ", " ", " ", "r", " ", "l", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    
    'MASTER Collection
    Mc3.Add Item:=pContro3, Key:="pControl"
    Mc3.Add Item:=nContro3, Key:="nControl"
    Mc3.Add Item:=mContro3, Key:="mControl"
    Mc3.Add Item:=iContro3, Key:="iControl"
    Mc3.Add Item:=rContro3, Key:="rControl"
    Mc3.Add Item:=cContro3, Key:="cControl"
    Mc3.Add Item:=aContro3, Key:="aControl"
    Mc3.Add Item:=lContro3, Key:="lControl"
    
    For i = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, i, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next i
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="CKG2010C.P_REFER1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    For i = 1 To ss3.MaxCols
        Call Gp_Sp_Collection(ss3, i, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Next i
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="CKG2010C.P_REFER3", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    For i = 1 To ss4.MaxCols
        Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Next i
    
    'Spread_Collection
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="CKG2010C.P_REFER4", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    'Call Gp_Sp_ColHidden(ss1, 43, True)
    Call Gp_Sp_ColHidden(ss3, 8, True)
    Call Gp_Sp_ColHidden(ss3, 9, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cmd_abnormal_send_Click()

    If Trim(txt_to) <> "" Then
        If MsgBox("确定要强制下达到 '" + txt_to + "' 的作业指示吗？", vbOKCancel, "指示下达确定") = vbOK Then
            If Gf_Mc_Authority(sAuthority, Mc1) Then
                If Gp_Process_Exec("A") = "" Then
                    MsgBox ("作业指示下达完毕 ！")
                    Call Form_Ref
                End If
            End If
        End If
    Else
        MsgBox ("请选择要强制下达的板坯号 ！")
    End If
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
    End With

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
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    Call Gp_Sp_Setting(sc4.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc4.Item("Spread"))
   
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(sc4)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc4.Item("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Spl_SizeGet(SSSplitter2, "G-System.INI", Me.Name, "W")
    
    txt_plt.Text = "C3"
    
    Call txt_plt_KeyUp(0, 0)
    If sAuthority = "1000" Then
       cmd_abnormal_send.Enabled = False
       cmd_roll_mana.Enabled = False
    End If
    
    opt_from.Enabled = False
    opt_to.Enabled = False
    opt_target.Enabled = False
        
    Screen.MousePointer = vbDefault
    
    txt_search_slabno.Text = "检索板坯号"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc4.Item("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Spl_SizeSet(SSSplitter2, "G-System.INI", Me.Name)
    
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
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
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
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Sc1 = Nothing
    Set Sc3 = Nothing
    Set sc4 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Sc3) And Gf_Sp_Cls(sc4) Then
        
        If Gf_Sp_Cls(Sc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            MDIMain.MenuTool.Buttons(4).Enabled = True
            txt_plt.Text = "C3"
            Call txt_plt_KeyUp(0, 0)
            opt_sent.Value = False
            opt_cancel.Value = False
            opt_move.Value = False
            opt_delete.Value = False
            opt_return.Value = False
            opt_from.Value = False
            opt_to.Value = False
            opt_target.Value = False
            opt_from.Enabled = False
            opt_to.Enabled = False
            opt_target.Enabled = False
            opt_sent.ForeColor = &H808080
            opt_move.ForeColor = &H808080
            opt_delete.ForeColor = &H808080
            opt_return.ForeColor = &H808080
            opt_cancel.ForeColor = &H808080
            opt_from.ForeColor = &H808080
            opt_to.ForeColor = &H808080
            opt_target.ForeColor = &H808080
            txt_from = ""
            from_y.Text = ""
            txt_to = ""
            to_y.Text = ""
            txt_target = ""
            target_y.Text = ""
            txt_MPLATE_NO = ""
            sSlab_Edt_Seq_Fr = 0
            sSlab_Edt_Seq_To = 0
            sSlab_Edt_Seq_Tg = 0
        End If
        
        With MDIMain.MenuTool
            .Buttons(7).Enabled = False                 'Row Insert
            .Buttons(8).Enabled = False                 'Row Delete
            .Buttons(9).Enabled = False                 'Row Cancel
            .Buttons(11).Enabled = False                'Copy
            .Buttons(12).Enabled = False                'Paste
            .Buttons(14).Enabled = False                'Excel
        End With
        
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sTemp As String
    Dim sL2_Send As String
    Dim sSlab_No As String
    Dim sPrc_Sts As String
    Dim sProd_Date As String
    Dim sHcr_Fl As String
    Dim iOrd_cnt As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    Dim cBackColor As Variant
    Dim sUrgnt_Fl As String
    Dim sOVER_FL As String

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
    
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       ss1.OperationMode = OperationModeNormal
       Call Gf_Sp_Cls(Sc3)
       Call Gf_Sp_Cls(sc4)
    
        With MDIMain.MenuTool
            .Buttons(7).Enabled = False                 'Row Insert
            .Buttons(8).Enabled = False                 'Row Delete
            .Buttons(9).Enabled = False                 'Row Cancel
            .Buttons(11).Enabled = False                'Copy
            .Buttons(12).Enabled = False                'Paste
            .Buttons(14).Enabled = True                 'Excel
        End With
        
        sSlab_Edt_Seq_Fr = 0
        sSlab_Edt_Seq_To = 0
        sSlab_Edt_Seq_Tg = 0

    End If
    
    With ss1
    
        For iRow = 1 To .MaxRows
            
            .ROW = iRow
            .Col = SS1_SLAB_NO:    sSlab_No = Trim(.Text)
            .Col = SS1_L2_SEND:    sL2_Send = Trim(.Text)
            .Col = SS1_PRC_STS:    sPrc_Sts = Trim(.Text)
            .Col = SS1_ORD_CNT:    iOrd_cnt = .Value
            .Col = SS1_PROD_DATE:  sProd_Date = Trim(.Text)
            .Col = SS1_HCR_FL:     sHcr_Fl = Trim(.Text)
            .Col = SS1_URGNT_FL:   sUrgnt_Fl = Trim(.Text)
            .Col = SS1_OVER_FL:    sOVER_FL = Trim(.Text)
            
            
            If sPrc_Sts = "B" Then
                Call Gp_Sp_RowColor(ss1, iRow, , SSPpdt.BackColor)
            Else
                If sL2_Send = "Y" Then
                    Call Gp_Sp_RowColor(ss1, iRow, , SSPsend.BackColor)
                End If
            End If
                
            If iOrd_cnt > 1 Then
                Call Gp_Sp_RowColor(ss1, iRow, , SSOrd.BackColor)
            End If
            
            .Col = 1
            cBackColor = .BackColor
            
            If sProd_Date = "" And sHcr_Fl = "H" Then
                Call Gp_Sp_CellColor(ss1, SS1_SLAB_NO, iRow, vbRed, cBackColor)
            End If
             '紧急订单绿色显示 add by liqian 2012-08-16
            If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, SS1_SLAB_NO, SS1_SLAB_NO, iRow, iRow, &HC000&)
                     Call Gp_Sp_BlockColor(ss1, SS1_ORD_NO, SS1_ORD_NO, iRow, iRow, &HC000&)
                     Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, iRow, iRow, &HC000&)
            End If
            
            If sOVER_FL <> "" Then
                Call Gp_Sp_BlockColor(ss1, SS1_OVER_FL, SS1_OVER_FL, iRow, iRow, &HFF&)
            End If
            
        Next iRow
        
        Call .SetActiveCell(1, .MaxRows)

    End With
            
End Sub

Public Sub Form_Pro()

    Dim mResult As String
    Dim sMsg As String

    Mode = ""

    If opt_move = True Then
        If Not ((from_y.Text = "Y" And to_y.Text = "Y" And target_y.Text = "Y") _
        Or (from_y.Text = "" And to_y.Text = "" And target_y.Text = "")) Then
            MsgBox ("已下达的指示和未下达的指示不能混在一起调整！")
            Exit Sub
        End If
    End If

    If opt_sent = True Then
    
        Mode = "L"
    
        If txt_to <> "" Then
            If MsgBox("确定要下达到 '" + txt_to + "' 的作业指示吗？", vbOKCancel, "指示下达确定") = vbOK Then
                If Gp_Process_Exec = "" Then
                    MsgBox ("作业指示下达完毕 ！")
                    Call Form_Ref
                End If
            End If
        
        Else
            MsgBox ("请选择目标板坯号 ！")
        End If
    
    End If

    If opt_cancel = True Then
        
        Mode = "C"
    
        If txt_from <> "" Then
            If MsgBox("确定取消从 '" + txt_from + "' 的作业指示吗？", vbOKCancel, "指示取消确定") = vbOK Then
                If Gp_Process_Exec = "" Then
                    MsgBox ("取消指示完毕 ！")
                    Call Form_Ref
                End If
            End If
        Else
            MsgBox ("请选择起始板坯号 ！")
        End If
    End If
 
    If opt_move = True Then
        
        Mode = "M"
    
        If txt_from.Text <> "" And txt_to.Text <> "" And txt_target.Text <> "" Then  '顺序变更
            sMsg = "确定要把板坯从(" + txt_from.Text + ")->(" + txt_to.Text + ")" + "调整到板坯(" + txt_target.Text + ")后边吗？"
        Else
            sMsg = "必须输入起始板坯号、终止板坯号和目标板坯号！"
            Call Gp_MsgBoxDisplay(sMsg)
            Exit Sub
        End If
    
        mResult = MsgBox(sMsg, vbYesNo)
        
        If mResult = vbYes Then
            If Gp_Process_Exec = "" Then
                MsgBox ("作业指示调整完毕 ！")
                Call Form_Ref
            End If
        End If
 
    End If
 
    If opt_delete = True Then
    
        Mode = "X"
        
        If txt_from.Text = "" Then
            sMsg = "必须输入起始板坯号！"
            Call Gp_MsgBoxDisplay(sMsg)
            Exit Sub
        End If
        
        sMsg = "确定要删除选定板坯(" + txt_from.Text + ")" + ")吗？"
        
        If txt_to.Text <> "" Then
            sMsg = "确定要删除选定板坯(" + txt_from.Text + ")->(" + txt_to.Text + ")吗？"
        End If
        
        mResult = MsgBox(sMsg, vbYesNo)
        
        If mResult = vbYes Then
           If Gp_Process_Exec = "" Then
              MsgBox ("作业指示删除完毕 ！")
              Call Form_Ref
           End If
        End If
    End If
 
    If opt_return = True Then
    
        Mode = "B"
        
        If txt_from.Text = "" Then
            sMsg = "必须输入起始板坯号！"
            Call Gp_MsgBoxDisplay(sMsg)
            Exit Sub
        End If
        
        sMsg = "确定要返送选定板坯(" + txt_from.Text + ")" + ")吗？"
        
        If txt_to.Text <> "" Then
            sMsg = "确定要返送选定板坯(" + txt_from.Text + ")->(" + txt_to.Text + ")吗？"
        End If
        
        mResult = MsgBox(sMsg, vbYesNo)
        
        If mResult = vbYes Then
           If Gp_Process_Exec = "" Then
              MsgBox ("作业指示返送完毕 ！")
              Call Form_Ref
           End If
        End If
    End If
 
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
    End With
    
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

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Private Sub opt_cancel_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_cancel.Value = True Then
        opt_cancel.ForeColor = &HFF&
        opt_sent.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_return.ForeColor = &H808080
        opt_from.Enabled = True
        opt_from.Value = True
        opt_to.Enabled = False
        opt_target.Enabled = False
    Else
        opt_cancel.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_delete_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_delete.Value = True Then
    
        opt_delete.ForeColor = &HFF&
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_return.ForeColor = &H808080
        opt_from.Enabled = True
        opt_from.Value = True
        opt_to.Enabled = True
        opt_target.Enabled = False
    Else
        opt_delete.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_from_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_from.Value = True Then
        opt_from.ForeColor = &HFF&
        opt_to.ForeColor = &H808080
        opt_target.ForeColor = &H808080
    Else
        opt_from.ForeColor = &H808080
    End If
    
End Sub

Private Sub opt_move_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_move.Value = True Then
        opt_move.ForeColor = &HFF&
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_return.ForeColor = &H808080
        opt_from.Enabled = True
        opt_from.Value = True
        opt_to.Enabled = True
        opt_target.Enabled = True
    Else
        opt_move.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_return_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_return.Value = True Then
    
        opt_return.ForeColor = &HFF&
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_from.Enabled = True
        opt_from.Value = True
        opt_to.Enabled = True
        opt_target.Enabled = False
    Else
        opt_return.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0

End Sub

Private Sub opt_sent_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_sent.Value = True Then
        opt_sent.ForeColor = &HFF&
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_return.ForeColor = &H808080
        opt_from.Enabled = False
        opt_to.Enabled = True
        opt_to.Value = True
        opt_target.Enabled = False
    Else
        opt_sent.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_target_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_target.Value = True Then
        opt_target.ForeColor = &HFF&
        opt_from.ForeColor = &H808080
        opt_to.ForeColor = &H808080
    Else
        opt_target.ForeColor = &H808080
    End If
    
End Sub

Private Sub opt_to_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_to.Value = True Then
        opt_to.ForeColor = &HFF&
        opt_from.ForeColor = &H808080
        opt_target.ForeColor = &H808080
    Else
        opt_to.ForeColor = &H808080
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    Dim C, M As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    Dim SEND_SLAB As String

    If Gf_Sp_Change(Proc_Sc, Sc1) Then
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0
    End If
    
    If ROW < 1 Then Exit Sub
    
    If ss1.MaxRows < 1 Then
        Call Gf_Sp_Cls(Sc3)
        Call Gf_Sp_Cls(sc4)
        Exit Sub
    End If
    
    ss1.ROW = ROW
    ss1.Col = SS1_SLAB_NO

    If opt_from.Value = True Then
       txt_from.Text = ss1.Text
       ss1.Col = SS1_L2_SEND
       from_y.Text = ss1.Text
       
       ss1.Col = SS1_SLAB_EDT_SEQ
       sSlab_Edt_Seq_Fr = ss1.Text
       
    ElseIf opt_to.Value = True Then
       txt_to.Text = ss1.Text
       ss1.Col = SS1_L2_SEND
       to_y.Text = ss1.Text
       
       ss1.Col = SS1_SLAB_EDT_SEQ
       sSlab_Edt_Seq_To = ss1.Text
       
    ElseIf opt_target.Value = True Then
       txt_target.Text = ss1.Text
       ss1.Col = SS1_L2_SEND
       target_y.Text = ss1.Text
       
       ss1.Col = SS1_SLAB_EDT_SEQ
       sSlab_Edt_Seq_Tg = ss1.Text
       
    End If

    '检查L2_SEND是否已为'Y'
    ss1.Col = SS1_L2_SEND
    If (opt_sent.Value = True Or opt_delete.Value = True) Then
        If ss1.Text = "Y" Then
            ss1.Col = SS1_SLAB_NO
            MsgBox ("板坯号 " + ss1.Text + " 作业指示已下达！")
            
            If opt_from.Value = True Then
                txt_from.Text = ""
                from_y.Text = ""
                sSlab_Edt_Seq_Fr = 0
                
            ElseIf opt_to.Value = True Then
                txt_to.Text = ""
                to_y.Text = ""
                sSlab_Edt_Seq_To = 0
                
            ElseIf opt_target.Value = True Then
                txt_target.Text = ""
                target_y.Text = ""
                sSlab_Edt_Seq_Tg = 0
                
            End If
            Exit Sub
        End If
    End If
    
           
    '检查是否已入炉
    ss1.Col = SS1_PRC_STS
    If ss1.Text = "B" And (opt_sent.Value = True Or opt_move.Value = True Or opt_cancel.Value = True Or opt_delete.Value = True) Then
        ss1.Col = SS1_SLAB_NO
        MsgBox ("板坯号 " + ss1.Text + " 已入炉，不能调整！")
        
        If opt_from.Value = True Then
           txt_from.Text = ""
           from_y.Text = ""
           sSlab_Edt_Seq_Fr = 0
           
        ElseIf opt_to.Value = True Then
           txt_to.Text = ""
           to_y.Text = ""
           sSlab_Edt_Seq_To = 0
           
        ElseIf opt_target.Value = True Then
           txt_target.Text = ""
           target_y.Text = ""
           sSlab_Edt_Seq_Tg = 0
           
        End If
        
        Exit Sub
    End If
                       
    If opt_sent = False And opt_cancel = False And opt_move = False And opt_delete = False And opt_return = False Then
        
        ss1.ROW = ROW
        
        ss1.Col = SS1_SLAB_NO
        txt_to.Text = ss1.Text
        
        If Len(Trim(txt_to.Text)) = 10 Then
            ss1.Col = SS1_SLAB_EDT_SEQ
            SDB_SLAB_EDT_SEQ.Value = ss1.Text
            Call Gf_Sp_Refer(M_CN1, Sc3, Mc2, , , False)
            ss3.OperationMode = OperationModeNormal
        End If
        
        txt_to.Text = ""
        
    End If
        
End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal ROW As Long)

    Dim P As Integer
    Dim iRow As Integer
    Dim iCol As Integer

    If ss3.ROW < 1 Then Exit Sub
    
    If Gf_Sp_Change(Proc_Sc, Sc3) Then
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0
    End If
    
    If ss3.MaxRows < 1 Then
        txt_MPLATE_NO.Text = ""
        Call Gf_Sp_Cls(sc4)
        Exit Sub
    End If
    
    ss3.ROW = ROW
    ss3.Col = SS3_SLAB_EDT_SEQ
    SDB_SLAB_EDT_SEQ.Value = ss3.Text
    ss3.Col = SS3_MPLATE_EDT_SEQ
    SDB_MPLATE_EDT_SEQ.Value = ss3.Text
    
    Call Gf_Sp_Refer(M_CN1, sc4, Mc3, Mc3("nControl"), Mc3("mControl"), False)

    ss4.OperationMode = OperationModeNormal

End Sub

Private Sub ss3_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss4_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal ROW As Long)

    If Gf_Sp_Change(Proc_Sc, sc4) Then
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0
    End If

End Sub

Private Sub SSPanel1_Click()
    
    opt_sent.Value = False
    opt_cancel.Value = False
    opt_move.Value = False
    opt_delete.Value = False
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    opt_sent.ForeColor = &H808080
    opt_move.ForeColor = &H808080
    opt_delete.ForeColor = &H808080
    opt_cancel.ForeColor = &H808080
    opt_from.ForeColor = &H808080
    opt_to.ForeColor = &H808080
    opt_target.ForeColor = &H808080
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=TXT_PLT_NAME
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        TXT_PLT_NAME.Text = ""
    End If

End Sub

Private Sub cmd_roll_mana_Click()

    Dim sMsg As String
    Dim mResult As String
    
    If txt_target.Text <> "" Then
       sMsg = "确定从板坯（" + txt_target.Text + "）进行辊期编制吗？"
       mResult = MsgBox(sMsg, vbYesNo)
       
       If mResult = vbYes Then
           If Gp_Process_Exec("R") = "" Then
              MsgBox ("辊期编制完毕 ！")
              Call Form_Ref
           End If
       End If
       
    End If
    
End Sub

Private Sub txt_search_slabno_Click()

    If txt_search_slabno.Text = "检索板坯号" Then
        txt_search_slabno.Text = ""
    End If
    
End Sub

Private Sub txt_search_slabno_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer

    If KeyCode = 13 Then
        
        For i = 1 To ss1.MaxRows
            ss1.ROW = i
            ss1.Col = SS1_SLAB_NO
            If ss1.Text = Trim(txt_search_slabno.Text) Then
                Call ss1.SetActiveCell(SS1_SLAB_NO, i)
            End If
        Next i
    End If
    
End Sub

Public Function Gp_Process_Exec(Optional Process_Type As String) As String

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    Dim sSlab_Seq_Fr As String
    Dim sSlab_Seq_To As String
    Dim sSlab_Seq_Tg As String
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sSlab_Seq_Fr = sSlab_Edt_Seq_Fr
    sSlab_Seq_To = sSlab_Edt_Seq_To
    sSlab_Seq_Tg = sSlab_Edt_Seq_Tg
    
    Select Case Process_Type
    
        Case "A"
        
            sQuery = "{call CGG2046P ('" + sSlab_Seq_To + "',?)}"
            
        Case "R"
        
            sQuery = "{call BKG2050P ('" + "C3" + "','" + "1" + "','" + sSlab_Seq_Tg + "',?)}"
            
        Case Else
        
            Select Case Mode
            
                Case "L"
                
                    sQuery = "{call CGG2042P ('A','" + sSlab_Seq_To + "',?)}"
                
                Case "C"
                
                    sQuery = "{call CGG2041P ('" + sSlab_Seq_Fr + "','" + "M" + "',?)}"
                
                Case "M"
                
                    sQuery = "{call CGG2044P ('M','" + sSlab_Seq_Fr + "','" + sSlab_Seq_To + "','" + sSlab_Seq_Tg + "',?)}"
                
                Case "X"
                
                    sQuery = "{call AFZ1000P ('X','M','" + sSlab_Seq_Fr + "','" + sSlab_Seq_To + "','"
                    sQuery = sQuery + sSlab_Seq_Tg + "','" + sUserID + "',?)}"
                
                Case "B"
                
                    sQuery = "{call AFZ1301P ('C3','" + sSlab_Seq_Fr + "','" + sSlab_Seq_To + "',?)}"
                
            End Select
        
    End Select
        
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Screen.MousePointer = vbDefault
        Set adoCmd = Nothing
        Gp_Process_Exec = sErrMessg
        Exit Function
    End If
    
    Set adoCmd = Nothing
    Gp_Process_Exec = ""
    Screen.MousePointer = vbDefault
    Exit Function

Process_Exec_ERROR:
    Set adoCmd = Nothing
    Gp_Process_Exec = "Process_Exec_ERROR"
    Screen.MousePointer = vbDefault
    Err.Raise Err.Number, Err.Description & sQuery
    
End Function
