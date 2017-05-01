VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AFE5010C 
   Caption         =   "精炼原始操作记录_AFE5010C"
   ClientHeight    =   9225
   ClientLeft      =   285
   ClientTop       =   2220
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_from_heat 
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
      Left            =   10335
      MaxLength       =   8
      TabIndex        =   38
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox txt_to_heat 
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
      Left            =   11835
      MaxLength       =   8
      TabIndex        =   37
      Top             =   120
      Width           =   1200
   End
   Begin VB.ComboBox cbo_group 
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
      ItemData        =   "AFE5010C.frx":0000
      Left            =   8160
      List            =   "AFE5010C.frx":0002
      TabIndex        =   7
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cbo_shift 
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
      ItemData        =   "AFE5010C.frx":0004
      Left            =   5970
      List            =   "AFE5010C.frx":0006
      TabIndex        =   1
      Tag             =   "班次"
      Top             =   120
      Width           =   735
   End
   Begin InDate.ULabel ULabel 
      Height          =   315
      Left            =   4905
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "班次"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   7095
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "班别"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
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
   Begin InDate.UDate dtp_date 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.74
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   240
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8685
      Left            =   60
      TabIndex        =   2
      Top             =   510
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   15319
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "LF/VD/RH操作原始记录"
      TabPicture(0)   =   "AFE5010C.frx":0008
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LF_ID(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LF_ID(12)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LF_ID(11)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LF_ID(9)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LF_ID(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LF_plantemp(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LF_ID(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LF_plantemp(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_emp_name2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_main2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_emp_name1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_main1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_lf_cov_life1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_lf_cov_life2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "SSSplitter1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_vd_cov_life"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_main3"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_emp_name3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_main4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt_emp_name4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_rh_cov_life"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_lf_cov_life3"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmd_RHPRINT"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmd_VDPRINT"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmd_LFPRINT"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "LF/VD/RH成分"
      TabPicture(1)   =   "AFE5010C.frx":0024
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_rhchem"
      Tab(1).Control(1)=   "cmd_vdchem"
      Tab(1).Control(2)=   "cmd_lfchem"
      Tab(1).Control(3)=   "SSSplitter2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "LF/VD/RH原辅材料"
      TabPicture(2)   =   "AFE5010C.frx":0040
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSSplitter3"
      Tab(2).ControlCount=   1
      Begin Threed.SSCommand cmd_LFPRINT 
         Height          =   330
         Left            =   10860
         TabIndex        =   11
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LF记录打印"
      End
      Begin Threed.SSCommand cmd_VDPRINT 
         Height          =   330
         Left            =   12270
         TabIndex        =   12
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "VD记录打印"
      End
      Begin Threed.SSCommand cmd_RHPRINT 
         Height          =   330
         Left            =   13650
         TabIndex        =   24
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "RH记录打印"
      End
      Begin VB.TextBox txt_lf_cov_life3 
         Alignment       =   1  'Right Justify
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
         Left            =   4710
         TabIndex        =   40
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txt_rh_cov_life 
         Alignment       =   1  'Right Justify
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
         Left            =   7800
         TabIndex        =   34
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txt_emp_name4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   14040
         TabIndex        =   33
         Text            =   " "
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_main4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12780
         TabIndex        =   32
         Text            =   " "
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin SSSplitter.SSSplitter SSSplitter3 
         Height          =   8280
         Left            =   -74940
         TabIndex        =   30
         Top             =   345
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   14605
         _Version        =   196609
         SplitterBarWidth=   4
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "AFE5010C.frx":005C
         Begin FPSpread.vaSpread ss7 
            Height          =   3030
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   15000
            _Version        =   393216
            _ExtentX        =   26458
            _ExtentY        =   5345
            _StockProps     =   64
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   40
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFE5010C.frx":00CE
         End
         Begin FPSpread.vaSpread ss8 
            Height          =   1980
            Left            =   0
            TabIndex        =   35
            Top             =   3090
            Width           =   15000
            _Version        =   393216
            _ExtentX        =   26458
            _ExtentY        =   3492
            _StockProps     =   64
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   40
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFE5010C.frx":12A0
         End
         Begin FPSpread.vaSpread ss9 
            Height          =   3150
            Left            =   0
            TabIndex        =   36
            Top             =   5130
            Width           =   15000
            _Version        =   393216
            _ExtentX        =   26458
            _ExtentY        =   5556
            _StockProps     =   64
            ColsFrozen      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   40
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFE5010C.frx":2472
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   7815
         Left            =   -74940
         TabIndex        =   26
         Top             =   810
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   13785
         _Version        =   196609
         SplitterBarWidth=   4
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "AFE5010C.frx":3644
         Begin FPSpread.vaSpread ss4 
            Height          =   2775
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   15000
            _Version        =   393216
            _ExtentX        =   26458
            _ExtentY        =   4895
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
            MaxCols         =   37
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFE5010C.frx":36B6
         End
         Begin FPSpread.vaSpread ss5 
            Height          =   2220
            Left            =   0
            TabIndex        =   28
            Top             =   2835
            Width           =   15000
            _Version        =   393216
            _ExtentX        =   26458
            _ExtentY        =   3916
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
            MaxCols         =   37
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFE5010C.frx":4532
         End
         Begin FPSpread.vaSpread ss6 
            Height          =   2700
            Left            =   0
            TabIndex        =   29
            Top             =   5115
            Width           =   15000
            _Version        =   393216
            _ExtentX        =   26458
            _ExtentY        =   4762
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
            MaxCols         =   37
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFE5010C.frx":516E
         End
      End
      Begin VB.TextBox txt_emp_name3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13770
         TabIndex        =   23
         Text            =   " "
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_main3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12480
         TabIndex        =   22
         Text            =   " "
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_vd_cov_life 
         Alignment       =   1  'Right Justify
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
         Left            =   6240
         TabIndex        =   21
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   7815
         Left            =   60
         TabIndex        =   17
         Top             =   810
         Width           =   15000
         _ExtentX        =   26458
         _ExtentY        =   13785
         _Version        =   196609
         SplitterBarWidth=   4
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "AFE5010C.frx":5DAA
         Begin FPSpread.vaSpread ss1 
            Height          =   2265
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   15000
            _Version        =   393216
            _ExtentX        =   26458
            _ExtentY        =   3995
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
            MaxCols         =   18
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFE5010C.frx":5E1C
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   2070
            Left            =   0
            TabIndex        =   19
            Top             =   2325
            Width           =   15000
            _Version        =   393216
            _ExtentX        =   26458
            _ExtentY        =   3651
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
            MaxCols         =   18
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFE5010C.frx":670C
         End
         Begin FPSpread.vaSpread ss3 
            Height          =   3360
            Left            =   0
            TabIndex        =   20
            Top             =   4455
            Width           =   15000
            _Version        =   393216
            _ExtentX        =   26458
            _ExtentY        =   5927
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
            MaxCols         =   27
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AFE5010C.frx":6FFF
         End
      End
      Begin VB.TextBox txt_lf_cov_life2 
         Alignment       =   1  'Right Justify
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
         Left            =   3180
         TabIndex        =   9
         Text            =   " "
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txt_lf_cov_life1 
         Alignment       =   1  'Right Justify
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
         Left            =   1650
         TabIndex        =   8
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox txt_main1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11910
         TabIndex        =   6
         Text            =   " "
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_emp_name1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13260
         TabIndex        =   5
         Text            =   " "
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_main2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12180
         TabIndex        =   4
         Text            =   " "
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txt_emp_name2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   13500
         TabIndex        =   3
         Text            =   " "
         Top             =   390
         Visible         =   0   'False
         Width           =   285
      End
      Begin InDate.ULabel LF_plantemp 
         Height          =   345
         Index           =   1
         Left            =   13080
         Top             =   390
         Visible         =   0   'False
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Caption         =   "主控人"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel LF_ID 
         Height          =   345
         Index           =   1
         Left            =   11670
         Top             =   390
         Visible         =   0   'False
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   609
         Caption         =   "号长"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel LF_plantemp 
         Height          =   315
         Index           =   0
         Left            =   150
         Top             =   420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "炉盖寿命"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin Threed.SSCommand cmd_lfchem 
         Height          =   330
         Left            =   -64140
         TabIndex        =   13
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "LF成分打印"
      End
      Begin Threed.SSCommand cmd_vdchem 
         Height          =   330
         Left            =   -62730
         TabIndex        =   14
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "VD成分打印"
      End
      Begin InDate.ULabel LF_ID 
         Height          =   315
         Index           =   8
         Left            =   5610
         Top             =   420
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Caption         =   "VD"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel LF_ID 
         Height          =   315
         Index           =   9
         Left            =   7140
         Top             =   420
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Caption         =   "RH"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin Threed.SSCommand cmd_rhchem 
         Height          =   330
         Left            =   -61350
         TabIndex        =   25
         Top             =   420
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16744576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "RH成分打印"
      End
      Begin InDate.ULabel LF_ID 
         Height          =   315
         Index           =   11
         Left            =   1020
         Top             =   420
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Caption         =   "#1 LF"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel LF_ID 
         Height          =   315
         Index           =   12
         Left            =   2550
         Top             =   420
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Caption         =   "#2 LF"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel LF_ID 
         Height          =   315
         Index           =   0
         Left            =   4080
         Top             =   420
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Caption         =   "#3 LF"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
   End
   Begin Threed.SSCommand cmd_cover 
      Height          =   375
      Left            =   13845
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "炉盖更换"
   End
   Begin InDate.UDate txt_to_DATE 
      Height          =   315
      Left            =   3075
      TabIndex        =   15
      Tag             =   "终止日期"
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   9255
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "炉号"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   11595
      TabIndex        =   39
      Top             =   195
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2820
      TabIndex        =   16
      Top             =   180
      Width           =   255
   End
End
Attribute VB_Name = "AFE5010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name
'-- Sub_System Name
'-- Program Name
'-- Program ID        AFE501C
'-- Document No       Q-00-0010(Specification)
'-- Designer          yuan wei
'-- Coder             yuan wei
'-- Date              2004.11.16
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
Dim nColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection

Dim pColumn5 As New Collection      'Spread Primary Key Collection
Dim nColumn5 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn5 As New Collection      'Spread necessary Column Collection
Dim mColumn5 As New Collection      'Spread Insert Column Collection
Dim aColumn5 As New Collection      'Master -> Spread Column Collection
Dim lColumn5 As New Collection      'Spread Lock Column Collection

Dim pColumn6 As New Collection      'Spread Primary Key Collection
Dim nColumn6 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn6 As New Collection      'Spread necessary Column Collection
Dim mColumn6 As New Collection      'Spread Insert Column Collection
Dim aColumn6 As New Collection      'Master -> Spread Column Collection
Dim lColumn6 As New Collection      'Spread Lock Column Collection

Dim pColumn7 As New Collection      'Spread Primary Key Collection
Dim nColumn7 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn7 As New Collection      'Spread necessary Column Collection
Dim mColumn7 As New Collection      'Spread Insert Column Collection
Dim aColumn7 As New Collection      'Master -> Spread Column Collection
Dim lColumn7 As New Collection      'Spread Lock Column Collection

Dim pColumn8 As New Collection      'Spread Primary Key Collection
Dim nColumn8 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn8 As New Collection      'Spread necessary Column Collection
Dim mColumn8 As New Collection      'Spread Insert Column Collection
Dim aColumn8 As New Collection      'Master -> Spread Column Collection
Dim lColumn8 As New Collection      'Spread Lock Column Collection

Dim pColumn9 As New Collection      'Spread Primary Key Collection
Dim nColumn9 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn9 As New Collection      'Spread necessary Column Collection
Dim mColumn9 As New Collection      'Spread Insert Column Collection
Dim aColumn9 As New Collection      'Master -> Spread Column Collection
Dim lColumn9 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Sc4 As New Collection           'Spread Collection
Dim Sc5 As New Collection           'Spread Collection
Dim Sc6 As New Collection           'Spread Collection
Dim Sc7 As New Collection           'Spread Collection
Dim Sc8 As New Collection           'Spread Collection
Dim Sc9 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(dtp_date, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_to_DATE, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_from_heat, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_to_heat, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(cbo_group, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(cbo_shift, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_main1, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_emp_name1, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_main2, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_emp_name2, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_main3, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_emp_name3, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_main4, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_emp_name4, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_lf_cov_life1, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_lf_cov_life2, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_lf_cov_life3, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_vd_cov_life, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_rh_cov_life, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:="AFE5010C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
     Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
     Call Gp_Sp_Collection(ss3, 1, "p", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 16, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 17, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 18, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 19, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 20, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 21, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 22, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 23, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 24, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 25, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     
     Call Gp_Sp_Collection(ss4, 1, "p", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 5, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 6, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 7, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 8, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     Call Gp_Sp_Collection(ss4, 9, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 10, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 11, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 12, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 13, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 14, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 15, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 16, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 17, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 18, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 19, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 20, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 21, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 22, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 23, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 24, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 25, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 26, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 27, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 28, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 29, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 30, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 31, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 32, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 33, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 34, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 35, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 36, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 37, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
     
     Call Gp_Sp_Collection(ss5, 1, "p", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 2, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 3, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 4, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 5, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 6, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 7, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 8, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
     Call Gp_Sp_Collection(ss5, 9, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 10, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 11, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 12, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 13, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 14, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 15, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 16, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 17, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 18, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 19, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 20, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 21, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 22, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 23, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 24, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 25, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 26, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 27, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 28, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 29, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 30, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 31, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 32, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 33, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 34, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 35, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 36, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 37, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    
     Call Gp_Sp_Collection(ss6, 1, "p", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 2, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 3, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 4, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 5, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 6, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 7, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 8, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
     Call Gp_Sp_Collection(ss6, 9, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 10, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 11, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 12, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 13, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 14, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 15, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 16, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 17, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 18, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 19, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 20, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 21, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 22, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 23, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 24, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 25, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 26, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 27, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 28, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 29, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 30, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 31, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 32, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 33, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 34, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 35, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 36, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 37, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    
     Call Gp_Sp_Collection(ss7, 1, "p", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
     Call Gp_Sp_Collection(ss7, 2, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
     Call Gp_Sp_Collection(ss7, 3, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
     Call Gp_Sp_Collection(ss7, 4, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
     Call Gp_Sp_Collection(ss7, 5, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
     Call Gp_Sp_Collection(ss7, 6, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
     Call Gp_Sp_Collection(ss7, 7, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
     Call Gp_Sp_Collection(ss7, 8, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
     Call Gp_Sp_Collection(ss7, 9, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 10, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 11, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 12, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 13, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 14, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 15, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 16, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 17, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 18, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 19, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 20, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 21, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 22, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 23, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 24, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 25, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 26, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 27, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 28, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 29, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 30, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 31, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 32, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 33, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 34, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 35, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 36, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 37, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 38, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 39, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 40, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
     
     Call Gp_Sp_Collection(ss8, 1, "p", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
     Call Gp_Sp_Collection(ss8, 2, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
     Call Gp_Sp_Collection(ss8, 3, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
     Call Gp_Sp_Collection(ss8, 4, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
     Call Gp_Sp_Collection(ss8, 5, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
     Call Gp_Sp_Collection(ss8, 6, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
     Call Gp_Sp_Collection(ss8, 7, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
     Call Gp_Sp_Collection(ss8, 8, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
     Call Gp_Sp_Collection(ss8, 9, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 10, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 11, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 12, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 13, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 14, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 15, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 16, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 17, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 18, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 19, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 20, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 21, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 22, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 23, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 24, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 25, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 26, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 27, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 28, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 29, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 30, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 31, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 32, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 33, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 34, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 35, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 36, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 37, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 38, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 39, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 40, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
     
     Call Gp_Sp_Collection(ss9, 1, "p", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
     Call Gp_Sp_Collection(ss9, 2, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
     Call Gp_Sp_Collection(ss9, 3, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
     Call Gp_Sp_Collection(ss9, 4, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
     Call Gp_Sp_Collection(ss9, 5, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
     Call Gp_Sp_Collection(ss9, 6, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
     Call Gp_Sp_Collection(ss9, 7, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
     Call Gp_Sp_Collection(ss9, 8, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
     Call Gp_Sp_Collection(ss9, 9, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 10, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 11, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 12, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 13, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 14, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 15, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 16, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 17, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 18, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 19, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 20, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 21, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 22, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 23, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 24, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 25, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 26, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 27, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 28, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 29, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 30, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 31, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 32, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 33, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 34, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 35, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 36, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 37, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 38, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 39, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
    Call Gp_Sp_Collection(ss9, 40, " ", " ", " ", " ", " ", " ", pColumn9, nColumn9, mColumn9, iColumn9, aColumn9, lColumn9)
     'Spread_Collection
     
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFE5010C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFE5010C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AFE5010C.P_REFER3", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:="AFE5010C.P_REFER4", Key:="P-R"
    Sc4.Add Item:=pColumn4, Key:="pColumn"
    Sc4.Add Item:=nColumn4, Key:="nColumn"
    Sc4.Add Item:=aColumn4, Key:="aColumn"
    Sc4.Add Item:=mColumn4, Key:="mColumn"
    Sc4.Add Item:=iColumn4, Key:="iColumn"
    Sc4.Add Item:=lColumn4, Key:="lColumn"
    Sc4.Add Item:=1, Key:="First"
    Sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    Sc5.Add Item:=ss5, Key:="Spread"
    Sc5.Add Item:="AFE5010C.P_REFER5", Key:="P-R"
    Sc5.Add Item:=pColumn5, Key:="pColumn"
    Sc5.Add Item:=nColumn5, Key:="nColumn"
    Sc5.Add Item:=aColumn5, Key:="aColumn"
    Sc5.Add Item:=mColumn5, Key:="mColumn"
    Sc5.Add Item:=iColumn5, Key:="iColumn"
    Sc5.Add Item:=lColumn5, Key:="lColumn"
    Sc5.Add Item:=1, Key:="First"
    Sc5.Add Item:=ss5.MaxCols, Key:="Last"
    
    Sc6.Add Item:=ss6, Key:="Spread"
    Sc6.Add Item:="AFE5010C.P_REFER6", Key:="P-R"
    Sc6.Add Item:=pColumn6, Key:="pColumn"
    Sc6.Add Item:=nColumn6, Key:="nColumn"
    Sc6.Add Item:=aColumn6, Key:="aColumn"
    Sc6.Add Item:=mColumn6, Key:="mColumn"
    Sc6.Add Item:=iColumn6, Key:="iColumn"
    Sc6.Add Item:=lColumn6, Key:="lColumn"
    Sc6.Add Item:=1, Key:="First"
    Sc6.Add Item:=ss6.MaxCols, Key:="Last"
    
    Sc7.Add Item:=ss7, Key:="Spread"
    Sc7.Add Item:="AFE5010C.P_REFER7", Key:="P-R"
    Sc7.Add Item:=pColumn7, Key:="pColumn"
    Sc7.Add Item:=nColumn7, Key:="nColumn"
    Sc7.Add Item:=aColumn7, Key:="aColumn"
    Sc7.Add Item:=mColumn7, Key:="mColumn"
    Sc7.Add Item:=iColumn7, Key:="iColumn"
    Sc7.Add Item:=lColumn7, Key:="lColumn"
    Sc7.Add Item:=1, Key:="First"
    Sc7.Add Item:=ss7.MaxCols, Key:="Last"
    
    Sc8.Add Item:=ss8, Key:="Spread"
    Sc8.Add Item:="AFE5010C.P_REFER8", Key:="P-R"
    Sc8.Add Item:=pColumn8, Key:="pColumn"
    Sc8.Add Item:=nColumn8, Key:="nColumn"
    Sc8.Add Item:=aColumn8, Key:="aColumn"
    Sc8.Add Item:=mColumn8, Key:="mColumn"
    Sc8.Add Item:=iColumn8, Key:="iColumn"
    Sc8.Add Item:=lColumn8, Key:="lColumn"
    Sc8.Add Item:=1, Key:="First"
    Sc8.Add Item:=ss8.MaxCols, Key:="Last"
    
    Sc9.Add Item:=ss9, Key:="Spread"
    Sc9.Add Item:="AFE5010C.P_REFER9", Key:="P-R"
    Sc9.Add Item:=pColumn9, Key:="pColumn"
    Sc9.Add Item:=nColumn9, Key:="nColumn"
    Sc9.Add Item:=aColumn9, Key:="aColumn"
    Sc9.Add Item:=mColumn9, Key:="mColumn"
    Sc9.Add Item:=iColumn9, Key:="iColumn"
    Sc9.Add Item:=lColumn9, Key:="lColumn"
    Sc9.Add Item:=1, Key:="First"
    Sc9.Add Item:=ss9.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    Proc_Sc.Add Item:=Sc4, Key:="Sc4"
    Proc_Sc.Add Item:=Sc5, Key:="Sc5"
    Proc_Sc.Add Item:=Sc6, Key:="Sc6"
    Proc_Sc.Add Item:=Sc7, Key:="Sc7"
    Proc_Sc.Add Item:=Sc8, Key:="Sc8"
    Proc_Sc.Add Item:=Sc9, Key:="Sc9"
    
    ss1.Col = 18
   ' ss2.Col = 17
    'ss1.ColHidden = True
     ss1.ColHidden = False
    ' ss2.ColHidden = False
    
    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    iSumCnt = 1
    
    'Sum Column Setting
    iSumCol.Add Item:=5
  '  iSumCol.Add Item:=17
    
        
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
    
    cbo_group.AddItem ("A")
    cbo_group.AddItem ("B")
    cbo_group.AddItem ("C")
    cbo_group.AddItem ("D")
    
    cbo_shift.AddItem ("1")
    cbo_shift.AddItem ("2")
    cbo_shift.AddItem ("3")
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
        
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc4")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc4"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc4")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc5")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc5"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc5")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc6")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc6"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc6")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc7")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc7"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc7")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc8")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc8"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc8")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc9")("Spread"), False)
    Call Gf_Sp_Cls(Proc_Sc("Sc9"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc9")("Spread"), "F-System.INI", Me.Name)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "F-System.INI", Me.Name, "H")
    Call Gp_Spl_SizeGet(SSSplitter2, "F-System.INI", Me.Name, "H")
    Call Gp_Spl_SizeGet(SSSplitter3, "F-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc4.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc5.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc6.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc7.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc8.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc9.Item("Spread"))
    
    Call Menu_Setting
    
    dtp_date.RawData = Format(Now, "YYYYMMDD")
    txt_to_DATE.RawData = Format(Now, "YYYYMMDD")
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub Menu_Setting()
    
    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Row cancel
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Row cancel.
    
    ss1.Row = 0:    ss1.Col = 0:    ss1.Text = "LF"
    ss2.Row = 0:    ss2.Col = 0:    ss2.Text = "VD"
    ss3.Row = 0:    ss3.Col = 0:    ss3.Text = "RH"
    ss4.Row = 0:    ss4.Col = 0:    ss4.Text = "LF"
    ss5.Row = 0:    ss5.Col = 0:    ss5.Text = "VD"
    ss6.Row = 0:    ss6.Col = 0:    ss6.Text = "RH"
    ss7.Row = 0:    ss7.Col = 0:    ss7.Text = "LF"
    ss8.Row = 0:    ss8.Col = 0:    ss8.Text = "VD"
    ss9.Row = 0:    ss9.Col = 0:    ss9.Text = "RH"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call Gp_Spl_SizeSet(SSSplitter1, "F-System.INI", Me.Name)
    Call Gp_Spl_SizeSet(SSSplitter2, "F-System.INI", Me.Name)
    Call Gp_Spl_SizeSet(SSSplitter3, "F-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc4")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc5")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc6")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc7")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc8")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc9")("Spread"), "F-System.INI", Me.Name)
    
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
    
    Set iColumn7 = Nothing
    Set pColumn7 = Nothing
    Set lColumn7 = Nothing
    Set nColumn7 = Nothing
    Set mColumn7 = Nothing
    Set aColumn7 = Nothing
    
    Set iColumn8 = Nothing
    Set pColumn8 = Nothing
    Set lColumn8 = Nothing
    Set nColumn8 = Nothing
    Set mColumn8 = Nothing
    Set aColumn8 = Nothing
    
    Set iColumn9 = Nothing
    Set pColumn9 = Nothing
    Set lColumn9 = Nothing
    Set nColumn9 = Nothing
    Set mColumn9 = Nothing
    Set aColumn9 = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Sc4 = Nothing
    Set Sc5 = Nothing
    Set Sc6 = Nothing
    Set Sc7 = Nothing
    Set Sc8 = Nothing
    Set Sc9 = Nothing
    Set iSumCol = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    Call Gf_Sp_Cls(Proc_Sc("Sc4"))
    Call Gf_Sp_Cls(Proc_Sc("Sc5"))
    Call Gf_Sp_Cls(Proc_Sc("Sc6"))
    Call Gf_Sp_Cls(Proc_Sc("Sc7"))
    Call Gf_Sp_Cls(Proc_Sc("Sc8"))
    Call Gf_Sp_Cls(Proc_Sc("Sc9"))
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    
    rControl(1).SetFocus
   
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    dtp_date.RawData = Format(Now, "YYYYMMDD")
    txt_to_DATE.RawData = Format(Now, "YYYYMMDD")
    
    dtp_date.SetFocus
End Sub

Public Sub Form_Ref()
    Dim sQuery As String
    Dim sMesg As String
    Dim iWgtsum As Double
    Dim iRawMat()  As Double
     
    On Error GoTo Refer_Err
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc1").Item("Spread")) Then Exit Sub
      
    If dtp_date.RawData = "" Or txt_to_DATE.RawData = "" Then
        Call Gp_MsgBoxDisplay("请正确输入日期！")
        Exit Sub
    End If
    
    If Gf_Ms_Refer(M_CN1, Mc1, Nothing, Nothing, False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    sQuery = Gf_Ms_MakeQuery(Proc_Sc("Sc1").Item("P-R"), "R", pControl)
    If Gf_Total_Display(M_CN1, Proc_Sc("Sc1"), sQuery, iDupCnt, iSumCnt, iSumCol, False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
    
    sQuery = Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", pControl)
    If Gf_Total_Display(M_CN1, Proc_Sc("Sc2"), sQuery, iDupCnt, iSumCnt, iSumCol, False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
       
    sQuery = Gf_Ms_MakeQuery(Proc_Sc("Sc3").Item("P-R"), "R", pControl)
    If Gf_Total_Display(M_CN1, Proc_Sc("Sc3"), sQuery, iDupCnt, iSumCnt, iSumCol, False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
        
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc4"), Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Call Gp_Sp_EvenRowBackcolor(Sc4.Item("Spread"))
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc5"), Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Call Gp_Sp_EvenRowBackcolor(Sc5.Item("Spread"))
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
     
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc6"), Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Call Gp_Sp_EvenRowBackcolor(Sc6.Item("Spread"))
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
        
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc7"), Mc1, , , False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc8"), Mc1, , , False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc9"), Mc1, , , False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
    
    Call Spread_Total_Display(ss7)
    Call Spread_Total_Display(ss8)
    Call Spread_Total_Display(ss9)
    Call Gp_Sp_EvenRowBackcolor(Sc7.Item("Spread"), 1)
    Call Gp_Sp_EvenRowBackcolor(Sc8.Item("Spread"), 1)
    Call Gp_Sp_EvenRowBackcolor(Sc9.Item("Spread"), 1)
          
'    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
    
    Exit Sub
     
Refer_Err:

End Sub

Public Sub Spread_Total_Display(oSpr As vaSpread)
    Dim iRawMat()   As Double
    Dim iRow        As Integer
    Dim iCol        As Integer
    
    ReDim iRawMat(1 To oSpr.MaxCols)
    
    With oSpr
        If .MaxRows <> 0 Then
            .MaxRows = .MaxRows + 1
            Call Gp_Sp_BlockColor(oSpr, 1, .MaxCols, .MaxRows, .MaxRows, BLACK, &HE6E6FF)
            
            For iRow = 1 To .MaxRows - 1
                .Row = iRow
                For iCol = 4 To .MaxCols
                    .Col = iCol
                    iRawMat(iCol) = iRawMat(iCol) + Val(.VALUE & "")
                Next
            Next
            
            .Row = .MaxRows
            .Col = 1
            .Text = "合  计"
            For iCol = 4 To .MaxCols
                .Col = iCol
                .VALUE = iRawMat(iCol)
            Next
        End If
    End With
    
End Sub

Public Sub Form_Pro()

    Call Menu_Setting
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
    ss1.Col = 0:    ss1.Row = 0
    If Right(ss1.Text, 1) = "◎" Then Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ss2.Col = 0:    ss2.Row = 0
    If Right(ss2.Text, 1) = "◎" Then Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ss3.Col = 0:    ss3.Row = 0
    If Right(ss3.Text, 1) = "◎" Then Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ss4.Col = 0:    ss4.Row = 0
    If Right(ss4.Text, 1) = "◎" Then Call Gp_Sp_Excel(Me, Proc_Sc("Sc4")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ss5.Col = 0:    ss5.Row = 0
    If Right(ss5.Text, 1) = "◎" Then Call Gp_Sp_Excel(Me, Proc_Sc("Sc5")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ss6.Col = 0:    ss6.Row = 0
    If Right(ss6.Text, 1) = "◎" Then Call Gp_Sp_Excel(Me, Proc_Sc("Sc6")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ss7.Col = 0:    ss7.Row = 0
    If Right(ss7.Text, 1) = "◎" Then Call Gp_Sp_Excel(Me, Proc_Sc("Sc7")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ss8.Col = 0:    ss8.Row = 0
    If Right(ss8.Text, 1) = "◎" Then Call Gp_Sp_Excel(Me, Proc_Sc("Sc8")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ss9.Col = 0:    ss9.Row = 0
    If Right(ss9.Text, 1) = "◎" Then Call Gp_Sp_Excel(Me, Proc_Sc("Sc9")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End Sub

Public Sub Form_Exit()

    Unload Me
    
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

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

 Private Sub ss2_LostFocus()

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

 Private Sub ss4_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss5_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

 Private Sub ss5_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss6_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

 Private Sub ss6_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss7_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

 Private Sub ss7_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
Private Sub ss8_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

 Private Sub ss8_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss9_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

 Private Sub ss9_LostFocus()

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

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row < 1 Or ss1.MaxRows = Row Then Exit Sub
    
    Load LF_MODIF
    LF_MODIF.P_DATE = dtp_date.Text
    LF_MODIF.P_SHIFT = cbo_shift.Text
    LF_MODIF.P_GROUP = cbo_group.Text
    LF_MODIF.Show 1
    
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Or ss2.MaxRows = Row Then Exit Sub
    
    Load VD_MODIF
    VD_MODIF.P_DATE = dtp_date.Text
    VD_MODIF.P_SHIFT = cbo_shift.Text
    VD_MODIF.P_GROUP = cbo_group.Text
    VD_MODIF.Show
End Sub

Private Sub ss3_DBLClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row < 1 Or ss3.MaxRows = Row Then Exit Sub
    
    Load RH_MODIF
    RH_MODIF.P_DATE = dtp_date.Text
    RH_MODIF.P_SHIFT = cbo_shift.Text
    RH_MODIF.P_GROUP = cbo_group.Text
    RH_MODIF.Show
End Sub

Private Sub Header_Edit(oSpr As vaSpread)
    ss1.Col = 0:    ss1.Row = 0:  ss1.Text = Replace(ss1.Text, "◎", "")
    ss2.Col = 0:    ss2.Row = 0:  ss2.Text = Replace(ss2.Text, "◎", "")
    ss3.Col = 0:    ss3.Row = 0:  ss3.Text = Replace(ss3.Text, "◎", "")
    ss4.Col = 0:    ss4.Row = 0:  ss4.Text = Replace(ss4.Text, "◎", "")
    ss5.Col = 0:    ss5.Row = 0:  ss5.Text = Replace(ss5.Text, "◎", "")
    ss6.Col = 0:    ss6.Row = 0:  ss6.Text = Replace(ss6.Text, "◎", "")
    ss7.Col = 0:    ss7.Row = 0:  ss7.Text = Replace(ss7.Text, "◎", "")
    ss8.Col = 0:    ss8.Row = 0:  ss8.Text = Replace(ss8.Text, "◎", "")
    ss9.Col = 0:    ss9.Row = 0:  ss9.Text = Replace(ss9.Text, "◎", "")
    
    oSpr.Col = 0:   oSpr.Row = 0: oSpr.Text = oSpr.Text & "◎"
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc1")("Spread"), Col, Row)
    Call Header_Edit(ss1)
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc2")("Spread"), Col, Row)
    Call Header_Edit(ss2)
End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc3")("Spread"), Col, Row)
    Call Header_Edit(ss3)
End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc4")("Spread"), Col, Row)
    Call Header_Edit(ss4)
End Sub

Private Sub ss5_Click(ByVal Col As Long, ByVal Row As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc5")("Spread"), Col, Row)
    Call Header_Edit(ss5)
End Sub

Private Sub ss6_Click(ByVal Col As Long, ByVal Row As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc6")("Spread"), Col, Row)
    Call Header_Edit(ss6)
End Sub

Private Sub ss7_Click(ByVal Col As Long, ByVal Row As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc7")("Spread"), Col, Row)
    Call Header_Edit(ss7)
End Sub

Private Sub ss8_Click(ByVal Col As Long, ByVal Row As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc8")("Spread"), Col, Row)
    Call Header_Edit(ss8)
End Sub

Private Sub ss9_Click(ByVal Col As Long, ByVal Row As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc9")("Spread"), Col, Row)
    Call Header_Edit(ss9)
End Sub

Private Sub cmd_cover_Click()
    Load COVER_CH
    COVER_CH.Show
End Sub

