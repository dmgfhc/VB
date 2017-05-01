VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKN2040C 
   BackColor       =   &H00C0C0C0&
   Caption         =   "连铸作业指示调整及下达界面_AKN2040C"
   ClientHeight    =   9255
   ClientLeft      =   555
   ClientTop       =   1905
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   15930
      Top             =   330
   End
   Begin VB.TextBox txt_ccm_line 
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
      Left            =   16830
      MaxLength       =   1
      TabIndex        =   26
      Tag             =   "工厂"
      Top             =   1890
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt_plt 
      Alignment       =   2  'Center
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
      Left            =   15480
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "工厂"
      Top             =   2250
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt_plt_name 
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
      Left            =   15945
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "工厂"
      Top             =   2250
      Visible         =   0   'False
      Width           =   1290
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8055
      Left            =   45
      TabIndex        =   1
      Top             =   1170
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   14208
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AKN2040C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   2820
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   5250
         _Version        =   393216
         _ExtentX        =   9260
         _ExtentY        =   4974
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
         MaxCols         =   15
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2040C.frx":00D2
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   5175
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2880
         Width           =   5250
         _Version        =   393216
         _ExtentX        =   9260
         _ExtentY        =   9128
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
         MaxCols         =   22
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2040C.frx":0978
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   2820
         Left            =   5310
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   0
         Width           =   4830
         _Version        =   393216
         _ExtentX        =   8520
         _ExtentY        =   4974
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
         MaxCols         =   15
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2040C.frx":1513
      End
      Begin FPSpread.vaSpread ss5 
         Height          =   2820
         Left            =   10200
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   0
         Width           =   5070
         _Version        =   393216
         _ExtentX        =   8943
         _ExtentY        =   4974
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
         MaxCols         =   15
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2040C.frx":1DA7
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   5175
         Left            =   5310
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2880
         Width           =   4830
         _Version        =   393216
         _ExtentX        =   8520
         _ExtentY        =   9128
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
         MaxCols         =   22
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2040C.frx":264D
      End
      Begin FPSpread.vaSpread ss6 
         Height          =   5175
         Left            =   10200
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2880
         Width           =   5070
         _Version        =   393216
         _ExtentX        =   8943
         _ExtentY        =   9128
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
         MaxCols         =   22
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2040C.frx":31D2
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   15480
      Top             =   1890
      Visible         =   0   'False
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSFrame Frame1 
      Height          =   465
      Left            =   2145
      TabIndex        =   4
      Top             =   90
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      Enabled         =   0   'False
      ShadowStyle     =   1
      Begin Threed.SSOption opt_cancel 
         Height          =   285
         Left            =   285
         TabIndex        =   5
         Top             =   90
         Width           =   1170
         _ExtentX        =   2064
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
         Caption         =   "取消下达"
      End
      Begin Threed.SSOption opt_ccm 
         Height          =   285
         Left            =   6930
         TabIndex        =   6
         Top             =   -120
         Visible         =   0   'False
         Width           =   570
         _ExtentX        =   1005
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
         Caption         =   "连铸机号变"
      End
      Begin Threed.SSOption opt_change 
         Height          =   285
         Left            =   1755
         TabIndex        =   7
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
         Left            =   3240
         TabIndex        =   8
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
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
         Caption         =   "删除"
      End
      Begin Threed.SSOption opt_send 
         Height          =   285
         Left            =   6795
         TabIndex        =   9
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   4210688
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
      Begin Threed.SSOption opt_plan 
         Height          =   285
         Left            =   4305
         TabIndex        =   21
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "计划名"
      End
      Begin Threed.SSOption opt_subst 
         Height          =   285
         Left            =   5550
         TabIndex        =   33
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
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
         Caption         =   "替换炉"
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   465
      Left            =   2145
      TabIndex        =   10
      Top             =   570
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_move_line 
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
         ItemData        =   "AKN2040C.frx":3D1C
         Left            =   8010
         List            =   "AKN2040C.frx":3D1E
         TabIndex        =   32
         Tag             =   "炉座号"
         Top             =   80
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txt_to 
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
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "终止炉号"
         Top             =   80
         Width           =   1170
      End
      Begin VB.TextBox txt_target 
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
         Left            =   6795
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   13
         Tag             =   "终止炉号"
         Top             =   80
         Width           =   1170
      End
      Begin VB.TextBox txt_from 
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
         Left            =   1545
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "起始炉号"
         Top             =   80
         Width           =   1170
      End
      Begin Threed.SSOption opt_from 
         Height          =   315
         Left            =   270
         TabIndex        =   11
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   2940
         TabIndex        =   15
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
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
         Left            =   5550
         TabIndex        =   16
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
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
      Begin Threed.SSPanel SSPpdt 
         Height          =   420
         Left            =   10275
         TabIndex        =   34
         Top             =   15
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "计划下达"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPsend 
         Height          =   420
         Left            =   8850
         TabIndex        =   35
         Top             =   15
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   16761024
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "计划锁定"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   420
         Left            =   11700
         TabIndex        =   36
         Top             =   15
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "钢种变更"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSFrame Frame2 
      Height          =   465
      Left            =   75
      TabIndex        =   19
      Top             =   90
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      Enabled         =   0   'False
      ShadowStyle     =   1
      Begin Threed.SSOption opt_charge 
         Height          =   285
         Left            =   255
         TabIndex        =   0
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
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
         Caption         =   "炉次"
         Value           =   -1
      End
      Begin Threed.SSOption opt_slab 
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
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
         Caption         =   "板坯"
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   465
      Left            =   9990
      TabIndex        =   22
      Top             =   90
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin Threed.SSOption opt_move 
         Height          =   285
         Left            =   285
         TabIndex        =   23
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
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
         Caption         =   "移动"
      End
      Begin Threed.SSOption opt_unif 
         Height          =   285
         Left            =   1342
         TabIndex        =   24
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
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
         Caption         =   "统合"
      End
      Begin Threed.SSOption opt_split 
         Height          =   285
         Left            =   2400
         TabIndex        =   25
         Top             =   90
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
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
         Caption         =   "分开"
      End
   End
   Begin Threed.SSCommand cmd_manual_order 
      Height          =   465
      Left            =   13440
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   90
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   8388736
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
      Caption         =   "非计划材编制"
   End
   Begin Threed.SSFrame Frame3 
      Height          =   465
      Left            =   75
      TabIndex        =   37
      Top             =   570
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin Threed.SSOption opt_time_on 
         Height          =   345
         Left            =   255
         TabIndex        =   38
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
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
         Caption         =   "动态"
         Value           =   -1
      End
      Begin Threed.SSOption opt_time_off 
         Height          =   345
         Left            =   1080
         TabIndex        =   39
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
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
         Caption         =   "静态"
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   45
      X2              =   15200
      Y1              =   1125
      Y2              =   1125
   End
End
Attribute VB_Name = "AKN2040C"
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
'-- Program ID        AKN2040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              20011.1.18
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

Public Select_Spread As Object      'Spread Object

Public Select_Spread_Row As Integer 'Active Spread Row

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

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

Dim P_TYPE As String                'Charge('C'), Slab('S')
Dim P_MODE As String                'L2 Cancel('C'), Change('M'), Delete('D'), Plan_Name('P'), CCM Change('X'), L2 Send('L')
Dim P_ACTM As String                'Move('M'), Unification('U'), Split('S')

Dim P_Fr_Edt_Seq As Long            'Slab_Edt_Seq (From)
Dim P_To_Edt_Seq As Long            'Slab_Edt_Seq (To)
Dim P_Tr_Edt_Seq As Long            'Slab_Edt_Seq (Target)

Private Sub Form_Define()
        
    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_ccm_line, "p", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To SS1.MaxCols
     Call Gp_Sp_Collection(SS1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    'Spread_Collection
    sc1.Add Item:=SS1, Key:="Spread"
    sc1.Add Item:="AFN2040C.P_REFER1", Key:="P-R"
    
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=SS1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFN2040C.P_REFER2", Key:="P-R"
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
    Sc3.Add Item:="AFN2040C.P_REFER1", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss4.MaxCols
        Call Gp_Sp_Collection(ss4, iCol, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Next iCol
    
    'Spread_Collection
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:="AFN2040C.P_REFER2", Key:="P-R"
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
    Sc5.Add Item:="AFN2040C.P_REFER1", Key:="P-R"
    Sc5.Add Item:=pColumn5, Key:="pColumn"
    Sc5.Add Item:=nColumn5, Key:="nColumn"
    Sc5.Add Item:=aColumn5, Key:="aColumn"
    Sc5.Add Item:=mColumn5, Key:="mColumn"
    Sc5.Add Item:=iColumn5, Key:="iColumn"
    Sc5.Add Item:=lColumn5, Key:="lColumn"
    Sc5.Add Item:=1, Key:="First"
    Sc5.Add Item:=ss5.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss6.MaxCols
        Call Gp_Sp_Collection(ss6, iCol, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Next iCol
    
    'Spread_Collection
    Sc6.Add Item:=ss6, Key:="Spread"
    Sc6.Add Item:="AFN2040C.P_REFER2", Key:="P-R"
    Sc6.Add Item:=pColumn6, Key:="pColumn"
    Sc6.Add Item:=nColumn6, Key:="nColumn"
    Sc6.Add Item:=aColumn6, Key:="aColumn"
    Sc6.Add Item:=mColumn6, Key:="mColumn"
    Sc6.Add Item:=iColumn6, Key:="iColumn"
    Sc6.Add Item:=lColumn6, Key:="lColumn"
    Sc6.Add Item:=1, Key:="First"
    Sc6.Add Item:=ss6.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    Proc_Sc.Add Item:=Sc4, Key:="Sc4"
    Proc_Sc.Add Item:=Sc5, Key:="Sc5"
    Proc_Sc.Add Item:=Sc6, Key:="Sc6"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(SS1, 2, True)
    Call Gp_Sp_ColHidden(SS1, 5, True)
    Call Gp_Sp_ColHidden(SS1, 15, True)
    
    Call Gp_Sp_ColHidden(ss2, 1, True)
    Call Gp_Sp_ColHidden(ss2, 19, True)
    Call Gp_Sp_ColHidden(ss2, 20, True)
    Call Gp_Sp_ColHidden(ss2, 21, True)
    
    Call Gp_Sp_ColHidden(ss3, 2, True)
    Call Gp_Sp_ColHidden(ss3, 5, True)
    Call Gp_Sp_ColHidden(ss3, 15, True)
    
    Call Gp_Sp_ColHidden(ss4, 1, True)
    Call Gp_Sp_ColHidden(ss4, 19, True)
    Call Gp_Sp_ColHidden(ss4, 20, True)
    Call Gp_Sp_ColHidden(ss4, 21, True)
    
    Call Gp_Sp_ColHidden(ss5, 2, True)
    Call Gp_Sp_ColHidden(ss5, 5, True)
    Call Gp_Sp_ColHidden(ss5, 15, True)
    
    Call Gp_Sp_ColHidden(ss6, 1, True)
    Call Gp_Sp_ColHidden(ss6, 19, True)
    Call Gp_Sp_ColHidden(ss6, 20, True)
    Call Gp_Sp_ColHidden(ss6, 21, True)
    
End Sub

Private Sub cbo_move_line_Click()

    txt_target.Text = ""
    
End Sub

Private Sub cmd_manual_order_Click()

    Load AKN2041C
    AKN2041C.Show 1
    
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
    
    Call Gp_Spl_SizeGet(SSSplitter1, "K-System.INI", Me.Name)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc4.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc5.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc6.Item("Spread"), "K-System.INI", Me.Name)
    
    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    cbo_move_line.Clear
    cbo_move_line.AddItem "1"
    cbo_move_line.AddItem "2"
    cbo_move_line.AddItem "3"
    cbo_move_line.Text = ""
    
    opt_from.Enabled = False
    opt_to.Enabled = False
    opt_target.Enabled = False
    
    txt_target.Left = 6885
    cbo_move_line.Visible = False
    
    P_TYPE = "C"
    Call Form_Ref
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

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
        
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gf_Sp_Cls(Sc3)
            Call Gf_Sp_Cls(Sc4)
            Call Gf_Sp_Cls(Sc5)
            Call Gf_Sp_Cls(Sc6)
            
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call MenuTool_ReSet
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            txt_plt.Text = "B1"
            Call txt_plt_KeyUp(0, 0)
            
            opt_charge.Value = True
            
            opt_from.Value = False
            opt_to.Value = False
            opt_target.Value = False
            
            opt_from.Enabled = False
            opt_to.Enabled = False
            opt_target.Enabled = False
            
            txt_from.Text = ""
            txt_to.Text = ""
            txt_target.Text = ""
            
            cbo_move_line.Text = ""
            
            P_Fr_Edt_Seq = 0
            P_To_Edt_Seq = 0
            P_Tr_Edt_Seq = 0
            
        End If
        
    End If
    
End Sub

Public Sub Form_Ref()

    Dim Ref_FL As String
    Dim sQuery As String
    Dim Dynamic_Slab As String
    
    txt_ccm_line.Text = "1"
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    Dynamic_Slab = "SC1"
    sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
        
    If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
        Call Gp_Sp_Scolor(SS1, "Y")
        Call Gp_Sp_Scolor(ss2, "Y")
    Else
        Call Gp_Sp_Scolor(SS1, "N")
        Call Gp_Sp_Scolor(ss2, "N")
    End If
    
    txt_ccm_line.Text = "2"
    If Gf_Sp_Refer(M_CN1, Sc3, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc4, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If

    Dynamic_Slab = "SC2"
    sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
        
    If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
        Call Gp_Sp_Scolor(ss3, "Y")
        Call Gp_Sp_Scolor(ss4, "Y")
    Else
        Call Gp_Sp_Scolor(ss3, "N")
        Call Gp_Sp_Scolor(ss4, "N")
    End If
    
    txt_ccm_line.Text = "3"
    If Gf_Sp_Refer(M_CN1, Sc5, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc6, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    Dynamic_Slab = "SC3"
    sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
        
    If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
        Call Gp_Sp_Scolor(ss5, "Y")
        Call Gp_Sp_Scolor(ss6, "Y")
    Else
        Call Gp_Sp_Scolor(ss5, "N")
        Call Gp_Sp_Scolor(ss6, "N")
    End If
    
    If Ref_FL = "1" Then
    
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        
        SS1.OperationMode = OperationModeNormal
        ss2.OperationMode = OperationModeNormal
        ss3.OperationMode = OperationModeNormal
        ss4.OperationMode = OperationModeNormal
        ss5.OperationMode = OperationModeNormal
        ss6.OperationMode = OperationModeNormal
        
        Call Spread_Color_Setting(SS1)
        Call Spread_Color_Setting(ss2)
        Call Spread_Color_Setting(ss3)
        Call Spread_Color_Setting(ss4)
        Call Spread_Color_Setting(ss5)
        Call Spread_Color_Setting(ss6)
        
        P_Fr_Edt_Seq = 0
        P_To_Edt_Seq = 0
        P_Tr_Edt_Seq = 0
        
    End If
    
    opt_charge.Value = True
            
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

End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Active_Spread, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Pro()

On Error GoTo Process_Error

    Dim OutParam(1, 4) As Variant
    Dim errMsg As String
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    Dim sFrom_No As String
    Dim sTo_No As String
    Dim sTarget_No As String
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
     
    'P_TYPE ('C':Charge, 'S':Slab)
    'P_MODE ('C':取消下达, 'M':顺序调整, 'D':删除, 'P':计划名', 'E':替换炉，'X':连铸机号变, 'L':下达)
    'P_ACTM ('M':移动, 'U':统合, 'S':分开)
    
    If opt_subst.Value Then Exit Sub
    
    sFrom_No = txt_from.Text
    sTo_No = txt_to.Text
    sTarget_No = txt_target.Text
    
    If P_TYPE = "S" Then   ' CHANGE SLAB_NO --> SLAB_EDT_SEQ
    
        Select Case P_MODE
            
            Case "C"
            
                sFrom_No = P_Fr_Edt_Seq
            
            Case "M"
            
                sFrom_No = P_Fr_Edt_Seq
                sTo_No = P_To_Edt_Seq
                sTarget_No = P_Tr_Edt_Seq
            
            Case "D"
            
                sFrom_No = P_Fr_Edt_Seq
                sTo_No = P_To_Edt_Seq
            
            Case "P"
            
                If P_ACTM = "S" Then
                    sTarget_No = P_Tr_Edt_Seq
                End If
            
        End Select
    
    End If
    
    sQuery = "{call AFZ4000P ('" & txt_plt.Text & "','" & P_TYPE & "', '" & P_MODE & "', '" & P_ACTM & "','" & sFrom_No & "','" & sTo_No & "','" & sTarget_No & "','" & cbo_move_line.Text & "','" & sUserID & "',?)}"
    
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
        Exit Sub
    End If
    
    M_CN1.CommitTrans
    Set adoCmd = Nothing
    
    Call Form_Ref
    
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    sErrMessg = Err.Description & sQuery
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub

Public Sub Spread_Del()
    
End Sub

Private Sub opt_cancel_Click(Value As Integer)

    If opt_cancel.Value = False Then Exit Sub
    
    P_MODE = "C"
    P_ACTM = ""
    
    cbo_move_line.Text = ""
    
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_move.Value = False
    opt_unif.Value = False
    opt_split.Value = False
    
    opt_move.Enabled = False
    opt_unif.Enabled = False
    opt_split.Enabled = False
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    txt_target.Left = 6885
    cbo_move_line.Visible = False
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
    opt_from.Enabled = True
    opt_to.Enabled = False
    opt_target.Enabled = False
    
    opt_from.Value = True
    
    If opt_charge Then
        opt_from.Caption = "起始炉号"
        opt_to.Caption = "终止炉号"
        opt_target.Caption = "目标炉号"
    Else
        opt_from.Caption = "起始坯号"
        opt_to.Caption = "终止坯号"
        opt_target.Caption = "目标坯号"
    End If
    
    Call Gp_Sp_SubstColor(SS1, "N")
    Call Gp_Sp_SubstColor(ss3, "N")
    Call Gp_Sp_SubstColor(ss5, "N")
    
End Sub

Private Sub opt_ccm_Click(Value As Integer)

    If opt_ccm.Value = False Then Exit Sub
    
    P_MODE = "X"
    P_ACTM = ""
    
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_move.Value = False
    opt_unif.Value = False
    opt_split.Value = False
    
    opt_move.Enabled = False
    opt_unif.Enabled = False
    opt_split.Enabled = False
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    txt_target.Left = 6885
    cbo_move_line.Visible = False
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
    opt_from.Enabled = True
    
    If opt_charge Then
        opt_to.Enabled = False
    Else
        opt_to.Enabled = True
    End If
    
    opt_target.Enabled = False
    
    opt_from.Value = True
    
    If opt_charge Then
        opt_from.Caption = "起始炉号"
        opt_to.Caption = "终止炉号"
        opt_target.Caption = "目标炉号"
    Else
        opt_from.Caption = "起始坯号"
        opt_to.Caption = "终止坯号"
        opt_target.Caption = "目标坯号"
    End If
    
End Sub

Private Sub opt_change_Click(Value As Integer)

    If opt_change.Value = False Then Exit Sub
    
    P_MODE = "M"
    P_ACTM = ""
    
    cbo_move_line.Text = ""
    
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_move.Value = False
    opt_unif.Value = False
    opt_split.Value = False
    
    opt_move.Enabled = False
    opt_unif.Enabled = False
    opt_split.Enabled = False
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    If opt_charge Then
        txt_target.Left = 6795
        cbo_move_line.Visible = True
    End If
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
    opt_from.Enabled = True
    opt_to.Enabled = True
    opt_target.Enabled = True
    
    opt_from.Value = True
    
    opt_from.Caption = "起始坯号"
    opt_to.Caption = "终止坯号"
    opt_target.Caption = "目标坯号"
    
    Call Gp_Sp_SubstColor(SS1, "N")
    Call Gp_Sp_SubstColor(ss3, "N")
    Call Gp_Sp_SubstColor(ss5, "N")

End Sub

Private Sub opt_charge_Click(Value As Integer)

    If opt_charge.Value = False Then Exit Sub
    
    P_TYPE = "C"
    P_MODE = ""
    P_ACTM = ""
    
    opt_cancel.Value = False
    opt_change.Value = False
    opt_del.Value = False
    opt_plan.Value = False
    opt_ccm.Value = False
    opt_send.Value = False
    opt_subst.Value = False
    
    opt_del.Enabled = False
    opt_subst.Enabled = True
    
    cbo_move_line.Text = ""

    opt_move.Value = False
    opt_unif.Value = False
    opt_split.Value = False
    
    opt_move.Enabled = False
    opt_unif.Enabled = False
    opt_split.Enabled = False
    
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_from.Enabled = False
    opt_to.Enabled = False
    opt_target.Enabled = False
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    txt_target.Left = 6885
    cbo_move_line.Visible = False
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
    opt_from.Caption = "起始炉号"
    opt_to.Caption = "终止炉号"
    opt_target.Caption = "目标炉号"
    
    Call Gp_Sp_SubstColor(SS1, "N")
    Call Gp_Sp_SubstColor(ss3, "N")
    Call Gp_Sp_SubstColor(ss5, "N")
    
End Sub

Private Sub opt_del_Click(Value As Integer)

    If opt_del.Value = False Then Exit Sub
    
    P_MODE = "D"
    P_ACTM = ""
    
    cbo_move_line.Text = ""
    
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_move.Value = False
    opt_unif.Value = False
    opt_split.Value = False
    
    opt_move.Enabled = False
    opt_unif.Enabled = False
    opt_split.Enabled = False
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    txt_target.Left = 6885
    cbo_move_line.Visible = False
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
    opt_from.Enabled = True
    opt_to.Enabled = True
    opt_target.Enabled = False
    
    opt_from.Value = True
    
    opt_from.Caption = "起始坯号"
    opt_to.Caption = "终止坯号"
    opt_target.Caption = "目标坯号"
    
    Call Gp_Sp_SubstColor(SS1, "N")
    Call Gp_Sp_SubstColor(ss3, "N")
    Call Gp_Sp_SubstColor(ss5, "N")

End Sub

Private Sub opt_move_Click(Value As Integer)

    If opt_move.Value = False Then Exit Sub
    
    P_ACTM = "M"
    
    opt_to.Enabled = True
    opt_target.Enabled = True
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
End Sub

Private Sub opt_plan_Click(Value As Integer)

    If opt_plan.Value = False Then Exit Sub
    
    P_MODE = "P"
    P_ACTM = ""
    
    cbo_move_line.Text = ""
    
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_move.Value = True
    opt_unif.Value = False
    opt_split.Value = False
    
    opt_move.Enabled = True
    opt_unif.Enabled = True
    opt_split.Enabled = True
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    txt_target.Left = 6885
    cbo_move_line.Visible = False
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
    opt_from.Enabled = True
    opt_to.Enabled = True
    opt_target.Enabled = True
    
    opt_from.Value = True
    
    opt_from.Caption = "起始计划名"
    opt_to.Caption = "终止计划名"
    opt_target.Caption = "目标计划名"
    
    Call Gp_Sp_SubstColor(SS1, "N")
    Call Gp_Sp_SubstColor(ss3, "N")
    Call Gp_Sp_SubstColor(ss5, "N")
    
End Sub

Private Sub opt_send_Click(Value As Integer)

    If opt_send.Value = False Then Exit Sub
    
    P_MODE = "L"
    P_ACTM = ""
    
    cbo_move_line.Text = ""
    
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_move.Value = False
    opt_unif.Value = False
    opt_split.Value = False
    
    opt_move.Enabled = False
    opt_unif.Enabled = False
    opt_split.Enabled = False
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    txt_target.Left = 6885
    cbo_move_line.Visible = False
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
    opt_from.Enabled = False
    opt_to.Enabled = False
    opt_target.Enabled = False
    
    Call Gp_Sp_SubstColor(SS1, "N")
    Call Gp_Sp_SubstColor(ss3, "N")
    Call Gp_Sp_SubstColor(ss5, "N")
    
End Sub

Private Sub opt_slab_Click(Value As Integer)

    If opt_slab.Value = False Then Exit Sub
    
    P_TYPE = "S"
    P_MODE = ""
    P_ACTM = ""
    
    opt_cancel.Value = False
    opt_change.Value = False
    opt_del.Value = False
    opt_plan.Value = False
    opt_ccm.Value = False
    opt_send.Value = False
    opt_subst.Value = False
    
    opt_change.Enabled = True
    opt_del.Enabled = True
    opt_subst.Enabled = False
    
    cbo_move_line.Text = ""

    opt_move.Value = False
    opt_unif.Value = False
    opt_split.Value = False
    
    opt_move.Enabled = False
    opt_unif.Enabled = False
    opt_split.Enabled = False
    
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_from.Enabled = False
    opt_to.Enabled = False
    opt_target.Enabled = False
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    txt_target.Left = 6885
    cbo_move_line.Visible = False
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
    opt_from.Caption = "起始坯号"
    opt_to.Caption = "终止坯号"
    opt_target.Caption = "目标坯号"
    
    Call Gp_Sp_SubstColor(SS1, "N")
    Call Gp_Sp_SubstColor(ss3, "N")
    Call Gp_Sp_SubstColor(ss5, "N")
    
End Sub

Private Sub opt_split_Click(Value As Integer)

    If opt_split.Value = False Then Exit Sub
    
    P_ACTM = "S"
    
    opt_to.Enabled = False
    opt_target.Enabled = True
    
    If opt_charge Then
        opt_target.Caption = "目标炉号"
    Else
        opt_target.Caption = "目标坯号"
    End If
    
    opt_from.Value = True
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
End Sub

Private Sub opt_subst_Click(Value As Integer)

    If opt_subst.Value = False Then Exit Sub
    
    P_MODE = "E"
    P_ACTM = ""
    
    cbo_move_line.Text = ""
    
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_move.Value = False
    opt_unif.Value = False
    opt_split.Value = False
    
    opt_move.Enabled = False
    opt_unif.Enabled = False
    opt_split.Enabled = False
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    txt_target.Left = 6885
    cbo_move_line.Visible = False
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
    opt_from.Enabled = False
    opt_to.Enabled = False
    opt_target.Enabled = False
    
    Call Gp_Sp_SubstColor(SS1, "Y")
    Call Gp_Sp_SubstColor(ss3, "Y")
    Call Gp_Sp_SubstColor(ss5, "Y")

End Sub

Private Sub opt_time_off_Click(Value As Integer)

    If opt_time_off.Value = False Then Exit Sub
    
    opt_charge.Value = True
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmd_manual_order.Enabled = True
    opt_time_on.Value = False
    Timer1.Enabled = False
    
End Sub

Private Sub opt_time_on_Click(Value As Integer)

    If opt_time_on.Value = False Then Exit Sub
    
    opt_charge.Value = True
    Frame1.Enabled = False
    Frame2.Enabled = False
    opt_time_off.Value = False
    cmd_manual_order.Enabled = False
    Timer1.Enabled = True
    
End Sub

Private Sub opt_unif_Click(Value As Integer)

    If opt_unif.Value = False Then Exit Sub
    
    P_ACTM = "U"
    
    opt_to.Enabled = True
    opt_target.Enabled = False
    
    opt_from.Value = True
    
    txt_from.Text = ""
    txt_to.Text = ""
    txt_target.Text = ""
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To SS1.MaxCols
        ss3.ColWidth(iCol) = SS1.ColWidth(iCol)
        ss5.ColWidth(iCol) = SS1.ColWidth(iCol)
    Next iCol
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_charge = False Then Exit Sub
    
    With SS1
    
        .Row = Row
        
        .Col = 15
        If .Text = "Y" Then
            
            If opt_plan Then
            
                If opt_move Then Exit Sub
            
            Else
                Exit Sub
            End If
            
        End If
        
        If opt_from Then
        
            If opt_plan Then
                .Col = 7
                txt_from.Text = .Text
                If opt_split Then
                    txt_to.Text = .Text
                End If
            Else
                .Col = 8
                txt_from.Text = .Text
                If opt_ccm Then
                    txt_to.Text = .Text
                End If
            End If
            
            If opt_to.Enabled Then
                opt_to.Value = True
            ElseIf opt_target.Enabled Then
                opt_target.Value = True
            End If
            
        ElseIf opt_to Then
        
            If opt_plan Then
                .Col = 7
                txt_to.Text = .Text
            Else
                .Col = 8
                txt_to.Text = .Text
            End If
        
            If opt_target.Enabled Then
                opt_target.Value = True
            End If
        
        ElseIf opt_target Then
            
            If opt_plan Then
                If opt_split Then
                    .Col = 8
                    txt_target.Text = .Text
                Else
                    .Col = 7
                    txt_target.Text = .Text
                End If
            Else
                .Col = 8
                txt_target.Text = .Text
            End If
            
            cbo_move_line.Text = ""
            
        End If
        
    End With
    
End Sub

Private Sub ss2_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss2.MaxCols
        ss4.ColWidth(iCol) = ss2.ColWidth(iCol)
        ss6.ColWidth(iCol) = ss2.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_slab = False Then Exit Sub
    
    With ss2
    
        .Row = Row
        
        .Col = 19
        If .Text = "Y" Then Exit Sub
        
        If opt_from Then
        
            If opt_plan Then
                .Col = 6
                txt_from.Text = .Text
                If opt_split Then
                    txt_to.Text = .Text
                End If
            Else
                .Col = 8
                txt_from.Text = .Text
            End If
            
            If opt_to.Enabled Then
                opt_to.Value = True
            ElseIf opt_target.Enabled Then
                opt_target.Value = True
            End If
            
            .Col = 20
            P_Fr_Edt_Seq = .Value
            
        ElseIf opt_to Then
        
            If opt_plan Then
                .Col = 6
                txt_to.Text = .Text
            Else
                .Col = 8
                txt_to.Text = .Text
            End If
        
            If opt_target.Enabled Then
                opt_target.Value = True
            End If
            
            .Col = 20
            P_To_Edt_Seq = .Value
        
        ElseIf opt_target Then
            
            If opt_plan Then
                If opt_split Then
                    .Col = 8
                    txt_target.Text = .Text
                Else
                    .Col = 6
                    txt_target.Text = .Text
                End If
            Else
                .Col = 8
                txt_target.Text = .Text
            End If
            
            .Col = 20
            P_Tr_Edt_Seq = .Value
            
            cbo_move_line.Text = ""
            
        End If
        
    End With
    
End Sub

Private Sub ss3_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss3.MaxCols
        SS1.ColWidth(iCol) = ss3.ColWidth(iCol)
        ss5.ColWidth(iCol) = ss3.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss3_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_charge = False Then Exit Sub
    
    With ss3
    
        .Row = Row
        
        .Col = 15
        If .Text = "Y" Then
            
            If opt_plan Then
            
                If opt_move Then Exit Sub
            
            Else
                Exit Sub
            End If
            
        End If
        
        If opt_from Then
        
            If opt_plan Then
                .Col = 7
                txt_from.Text = .Text
                If opt_split Then
                    txt_to.Text = .Text
                End If
            Else
                .Col = 8
                txt_from.Text = .Text
                If opt_ccm Then
                    txt_to.Text = .Text
                End If
            End If
            
            If opt_to.Enabled Then
                opt_to.Value = True
            ElseIf opt_target.Enabled Then
                opt_target.Value = True
            End If
            
        ElseIf opt_to Then
        
            If opt_plan Then
                .Col = 7
                txt_to.Text = .Text
            Else
                .Col = 8
                txt_to.Text = .Text
            End If
        
            If opt_target.Enabled Then
                opt_target.Value = True
            End If
        
        ElseIf opt_target Then
            
            If opt_plan Then
                If opt_split Then
                    .Col = 8
                    txt_target.Text = .Text
                Else
                    .Col = 7
                    txt_target.Text = .Text
                End If
            Else
                .Col = 8
                txt_target.Text = .Text
            End If
            
            cbo_move_line.Text = ""
            
        End If
        
    End With
    
End Sub

Private Sub ss4_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss4.MaxCols
        ss2.ColWidth(iCol) = ss4.ColWidth(iCol)
        ss6.ColWidth(iCol) = ss4.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss4_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_slab = False Then Exit Sub
    
    With ss4
    
        .Row = Row
        
        .Col = 19
        If .Text = "Y" Then Exit Sub
        
        If opt_from Then
        
            If opt_plan Then
                .Col = 6
                txt_from.Text = .Text
                If opt_split Then
                    txt_to.Text = .Text
                End If
            Else
                .Col = 8
                txt_from.Text = .Text
            End If
            
            If opt_to.Enabled Then
                opt_to.Value = True
            ElseIf opt_target.Enabled Then
                opt_target.Value = True
            End If
            
            .Col = 20
            P_Fr_Edt_Seq = .Value
            
        ElseIf opt_to Then
        
            If opt_plan Then
                .Col = 6
                txt_to.Text = .Text
            Else
                .Col = 8
                txt_to.Text = .Text
            End If
        
            If opt_target.Enabled Then
                opt_target.Value = True
            End If
            
            .Col = 20
            P_To_Edt_Seq = .Value
        
        ElseIf opt_target Then
            
            If opt_plan Then
                If opt_split Then
                    .Col = 8
                    txt_target.Text = .Text
                Else
                    .Col = 6
                    txt_target.Text = .Text
                End If
            Else
                .Col = 8
                txt_target.Text = .Text
            End If
            
            .Col = 20
            P_Tr_Edt_Seq = .Value
            
            cbo_move_line.Text = ""
            
        End If
        
    End With
    
End Sub

Private Sub ss5_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss5.MaxCols
        SS1.ColWidth(iCol) = ss5.ColWidth(iCol)
        ss3.ColWidth(iCol) = ss5.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss5_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_charge = False Then Exit Sub
    
    With ss5
    
        .Row = Row
        
        .Col = 15
        If .Text = "Y" Then
            
            If opt_plan Then
            
                If opt_move Then Exit Sub
            
            Else
                Exit Sub
            End If
            
        End If
        
        If opt_from Then
        
            If opt_plan Then
                .Col = 7
                txt_from.Text = .Text
                If opt_split Then
                    txt_to.Text = .Text
                End If
            Else
                .Col = 8
                txt_from.Text = .Text
                If opt_ccm Then
                    txt_to.Text = .Text
                End If
            End If
            
            If opt_to.Enabled Then
                opt_to.Value = True
            ElseIf opt_target.Enabled Then
                opt_target.Value = True
            End If
            
        ElseIf opt_to Then
        
            If opt_plan Then
                .Col = 7
                txt_to.Text = .Text
            Else
                .Col = 8
                txt_to.Text = .Text
            End If
        
            If opt_target.Enabled Then
                opt_target.Value = True
            End If
        
        ElseIf opt_target Then
            
            If opt_plan Then
                If opt_split Then
                    .Col = 8
                    txt_target.Text = .Text
                Else
                    .Col = 7
                    txt_target.Text = .Text
                End If
            Else
                .Col = 8
                txt_target.Text = .Text
            End If
            
            cbo_move_line.Text = ""
            
        End If
        
    End With
    
End Sub

Private Sub ss6_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss6.MaxCols
        ss2.ColWidth(iCol) = ss6.ColWidth(iCol)
        ss4.ColWidth(iCol) = ss6.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss6_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_slab = False Then Exit Sub
    
    With ss6
    
        .Row = Row
        
        .Col = 19
        If .Text = "Y" Then Exit Sub
        
        If opt_from Then
        
            If opt_plan Then
                .Col = 6
                txt_from.Text = .Text
                If opt_split Then
                    txt_to.Text = .Text
                End If
            Else
                .Col = 8
                txt_from.Text = .Text
            End If
            
            If opt_to.Enabled Then
                opt_to.Value = True
            ElseIf opt_target.Enabled Then
                opt_target.Value = True
            End If
            
            .Col = 20
            P_Fr_Edt_Seq = .Value
            
        ElseIf opt_to Then
        
            If opt_plan Then
                .Col = 6
                txt_to.Text = .Text
            Else
                .Col = 8
                txt_to.Text = .Text
            End If
        
            If opt_target.Enabled Then
                opt_target.Value = True
            End If
            
            .Col = 20
            P_To_Edt_Seq = .Value
        
        ElseIf opt_target Then
            
            If opt_plan Then
                If opt_split Then
                    .Col = 8
                    txt_target.Text = .Text
                Else
                    .Col = 6
                    txt_target.Text = .Text
                End If
            Else
                .Col = 8
                txt_target.Text = .Text
            End If
            
            .Col = 20
            P_Tr_Edt_Seq = .Value
            
            cbo_move_line.Text = ""
            
        End If
        
    End With
    
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss4_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss5_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss6_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
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
    
    Set Active_Spread = Me.SS1
    Set Select_Spread = Me.SS1
    
    If Not opt_subst.Value Then Exit Sub
    
    If Col = 8 Then
        
        ss3.Row = Row
        SS1.Col = 14  'L2-CCM-SEND
        If SS1.Text = "" Then
        
            SS1.Col = 15  'L2-LCOK
            If SS1.Text = "" Then
            
                Select_Spread_Row = Row
                Load AKN2042C
                AKN2042C.Show 1
                
            End If
            
        End If
        
    End If
    
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss2
    
End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss3
    Set Select_Spread = Me.ss3
    
    If Not opt_subst.Value Then Exit Sub
    
    If Col = 8 Then
    
        ss3.Row = Row
        ss3.Col = 14  'L2-CCM-SEND
        If ss3.Text = "" Then
        
            ss3.Col = 15  'L2-LCOK
            If ss3.Text = "" Then
                Select_Spread_Row = Row
                Load AKN2042C
                AKN2042C.Show 1
            End If
            
        End If
        
    End If
    
End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss4
    
End Sub

Private Sub ss5_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss5
    Set Select_Spread = Me.ss5
    
    If Not opt_subst.Value Then Exit Sub
    
    If Col = 8 Then
    
        ss3.Row = Row
        ss5.Col = 14  'L2-CCM-SEND
        If ss5.Text = "" Then
        
            ss5.Col = 15  'L2-LCOK
            If ss5.Text = "" Then
                Select_Spread_Row = Row
                Load AKN2042C
                AKN2042C.Show 1
            End If
            
        End If
        
    End If
    
End Sub

Private Sub ss6_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss6
    
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

Private Sub ss5_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss6_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub Timer1_Timer()

    Call Form_Ref
    
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

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(9).Enabled = False                  'Row Cancel
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Sub Spread_Color_Setting(oSpr As vaSpread)

    Dim iRow As Long
    Dim sPlan_Name As String
    Dim sAct_Stlgrd_Grp As String
    Dim sAct_Stlgrd As String
    
    With oSpr
    
        If oSpr.Name = "ss1" Or oSpr.Name = "ss3" Or oSpr.Name = "ss5" Then
    
            For iRow = 1 To .MaxRows
                
                .Row = iRow
                
                .Col = 7  'PLAN_NAME
                
                If iRow = 1 Then
                
                    sPlan_Name = .Text
                    
                    Call Gp_Sp_Bold(oSpr, "N", .Row)
                    
                    .Col = 15  'L2-LCOK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                    Else
                        .Col = 14  'L2-CCM-SEND
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                        End If
                        
                        .Col = 1
                        If .Text <> "" Then
                            sAct_Stlgrd_Grp = .Text
                            .Col = 2
                            sAct_Stlgrd = .Text
                            
'                            .Col = 4
'                            If sAct_Stlgrd_Grp = "Z" Then
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
'                                Else
                                    .Col = 5
                                    If sAct_Stlgrd <> .Text Then
                                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
                                    End If
                               
'                                End If
'
'                            Else
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
'                                End If
'
'                            End If
                            
                        End If
                        
                    End If
                
                ElseIf sPlan_Name <> .Text Then
                    
                    sPlan_Name = .Text
                    
                    Call Gp_Sp_Bold(oSpr, "Y", .Row)
                    
                    .Col = 15  'L2-LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFC0C0)
                    Else
                        .Col = 14  'L2-CCM-SEND
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HC0FFFF)
                        Else
                            Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF)
                        End If
                        
                        .Col = 1
                        If .Text <> "" Then
                            sAct_Stlgrd_Grp = .Text
                            .Col = 2
                            sAct_Stlgrd = .Text
                            
'                            .Col = 4
'                            If sAct_Stlgrd_Grp = "Z" Then
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFFFC0)
'                                Else
                                    .Col = 5
                                    If sAct_Stlgrd <> .Text Then
                                        Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFFFC0)
                                    End If
                               
'                                End If
'
'                            Else
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFFFC0)
'                                End If
'
'                            End If
                            
                        End If
                    
                    End If
                    
                Else
                    
                    Call Gp_Sp_Bold(oSpr, "N", .Row)
                    
                    .Col = 15  'L2-LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                    Else
                        .Col = 14  'L2-CCM-SEND
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                        End If
                        
                        .Col = 1
                        If .Text <> "" Then
                            sAct_Stlgrd_Grp = .Text
                            .Col = 2
                            sAct_Stlgrd = .Text
                            
'                            .Col = 4
'                            If sAct_Stlgrd_Grp = "Z" Then
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
'                                Else
                                    .Col = 5
                                    If sAct_Stlgrd <> .Text Then
                                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
                                    End If
                               
'                                End If
'
'                            Else
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
'                                End If
'
'                            End If
                            
                        End If
                        
                    End If
                
                End If
                
            Next iRow
          
        Else
        
            For iRow = 1 To .MaxRows
                
                .Row = iRow
                
                .Col = 6  'PLAN_NAME
                
                If iRow = 1 Then
                
                    sPlan_Name = .Text
                
                    Call Gp_Sp_Bold(oSpr, "N", iRow)
                    
                    .Col = 19  'LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                    Else
                        .Col = 18  'L2
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                        End If
                    End If
                
                ElseIf sPlan_Name <> .Text Then
                    
                    sPlan_Name = .Text
                    
                    Call Gp_Sp_Bold(oSpr, "Y", .Row)
                
                    .Col = 19  'LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFC0C0)
                    Else
                        .Col = 18  'L2
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HC0FFFF)
                        Else
                            Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF)
                        End If
                    End If
                    
                Else
                    
                    Call Gp_Sp_Bold(oSpr, "N", .Row)
                    
                    .Col = 19  'LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                    Else
                        .Col = 18  'L2
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                        End If
                    End If
                
                End If
                
                .Row = iRow
                .Col = 21  'insert program-id
                
                If .Text <> "" Then
                    .Col = 8: .Col2 = 8
                    .Row = iRow: .Row2 = iRow
                    
                    .BlockMode = True
                    .ForeColor = vbRed
                    .BlockMode = False
                End If
                
            Next iRow
        
        End If
        
        .RowHeight(-1) = 12.54
          
    End With
    
End Sub

Private Sub Gp_Sp_Scolor(sPname As Variant, sColType As String)

    With sPname
    
        .Row = 0: .Row2 = 0
        .Col = 0: .Col2 = 0
        
        .BlockMode = True
        
        .CellType = SS_CELL_TYPE_STATIC_TEXT
        .TypeHAlign = SS_CELL_H_ALIGN_CENTER
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .TypeTextWordWrap = True
        
        .BackColor = &HE1E4CD
        
        If sColType = "N" Then
            .ForeColor = vbRed
        Else
            .ForeColor = vbBlack
        End If
        
        .BlockMode = False
        
    End With
    
End Sub

Private Sub Gp_Sp_SubstColor(sPname As Variant, sColType As String)

    Dim iRow As Long
    Dim sAct_Stlgrd_Grp As String
    Dim sAct_Stlgrd As String
    Dim vBackColor As Variant
    
    With sPname
    
        For iRow = 1 To .MaxRows
            
            .Row = iRow
            
            .Col = 15  'L2-LOCK
            
            If .Text = "" Then
            
                .Col = 14  'L2-CCM-SEND
                If .Text = "" Then
                
                    .Col = 1
                    vBackColor = .BackColor
                    
                    .Row = iRow
                    .Row2 = iRow
                    .Col = 8
                    .Col2 = 8
        
                    .BlockMode = True
        
                    If sColType = "Y" Then
                        .BackColor = &HC0FFEE
                    Else
                        .BackColor = vBackColor
                    End If
        
                    .BlockMode = False
            
                End If
                
            End If

        Next iRow
    
    End With
    
End Sub

Private Sub Gp_Sp_Bold(sPname As Variant, sType As String, iRow As Long)

    With sPname
    
        .Row = iRow: .Row2 = iRow
        .Col = 1: .Col2 = .MaxCols
        
        .BlockMode = True
        
        If sType = "N" Then
            .FontBold = False
        Else
            .FontBold = True
        End If
        
        .BlockMode = False
        
    End With
    
End Sub
