VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGE2080C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "钢卷库库存现状查询_AGE2080C"
   ClientHeight    =   9360
   ClientLeft      =   750
   ClientTop       =   1350
   ClientWidth     =   14205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin Threed.SSOption opt_area 
      Height          =   330
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   135
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      ActiveColors    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "A 区"
   End
   Begin Threed.SSOption opt_area 
      Height          =   330
      Index           =   1
      Left            =   4625
      TabIndex        =   3
      Top             =   135
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      ActiveColors    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "B 区"
   End
   Begin Threed.SSOption opt_area 
      Height          =   330
      Index           =   2
      Left            =   5530
      TabIndex        =   4
      Top             =   135
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      ActiveColors    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "C 区"
   End
   Begin Threed.SSOption opt_area 
      Height          =   330
      Index           =   3
      Left            =   6435
      TabIndex        =   5
      Top             =   135
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      ActiveColors    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "D 区"
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   555
      Left            =   1755
      TabIndex        =   16
      Top             =   8865
      Visible         =   0   'False
      Width           =   1245
      _Version        =   393216
      _ExtentX        =   2196
      _ExtentY        =   979
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
      MaxCols         =   7
      SpreadDesigner  =   "AGE2080C.frx":0000
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   555
      Left            =   315
      TabIndex        =   15
      Top             =   8865
      Visible         =   0   'False
      Width           =   1410
      _Version        =   393216
      _ExtentX        =   2487
      _ExtentY        =   979
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
      MaxCols         =   2
      SpreadDesigner  =   "AGE2080C.frx":021A
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8205
      LargeChange     =   7
      Left            =   14670
      Min             =   1
      TabIndex        =   14
      Top             =   630
      Value           =   1
      Visible         =   0   'False
      Width           =   285
   End
   Begin Threed.SSPanel pan_addr 
      Height          =   1170
      Index           =   0
      Left            =   315
      TabIndex        =   7
      Top             =   630
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2064
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16744576
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSRibbon rib_area 
      Height          =   375
      Index           =   0
      Left            =   1860
      TabIndex        =   0
      Top             =   135
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1跨"
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   12600
      Top             =   180
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Caption         =   "钢卷总数"
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
   Begin Threed.SSRibbon rib_area 
      Height          =   375
      Index           =   1
      Left            =   8250
      TabIndex        =   2
      Top             =   135
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2跨"
   End
   Begin CSTextLibCtl.sidbEdit sdb_coil_cnt 
      Height          =   315
      Left            =   13770
      TabIndex        =   6
      Top             =   180
      Width           =   870
      _Version        =   262145
      _ExtentX        =   1535
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
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSPanel pan_addr 
      Height          =   1170
      Index           =   1
      Left            =   315
      TabIndex        =   8
      Top             =   1800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2064
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16744576
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pan_addr 
      Height          =   1170
      Index           =   2
      Left            =   315
      TabIndex        =   9
      Top             =   2970
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2064
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16744576
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pan_addr 
      Height          =   1170
      Index           =   3
      Left            =   315
      TabIndex        =   10
      Top             =   4140
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2064
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16744576
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pan_addr 
      Height          =   1170
      Index           =   4
      Left            =   315
      TabIndex        =   11
      Top             =   5310
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2064
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16744576
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pan_addr 
      Height          =   1170
      Index           =   5
      Left            =   315
      TabIndex        =   12
      Top             =   6480
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2064
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16744576
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel pan_addr 
      Height          =   1170
      Index           =   6
      Left            =   315
      TabIndex        =   13
      Top             =   7650
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2064
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16744576
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   1170
      Left            =   1845
      TabIndex        =   17
      Top             =   1800
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   2064
      _Version        =   196609
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   81
         Left            =   9045
         TabIndex        =   79
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   80
         Left            =   8595
         TabIndex        =   78
         Top             =   150
         Width           =   375
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   41
         Left            =   9300
         TabIndex        =   77
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   79
         Left            =   8145
         TabIndex        =   76
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   78
         Left            =   7680
         TabIndex        =   75
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   77
         Left            =   7245
         TabIndex        =   74
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   76
         Left            =   6795
         TabIndex        =   73
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   75
         Left            =   6345
         TabIndex        =   72
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   74
         Left            =   5895
         TabIndex        =   71
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   73
         Left            =   5445
         TabIndex        =   70
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   72
         Left            =   4995
         TabIndex        =   69
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   71
         Left            =   4545
         TabIndex        =   68
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   70
         Left            =   4095
         TabIndex        =   67
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   69
         Left            =   3645
         TabIndex        =   66
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   68
         Left            =   3195
         TabIndex        =   65
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   67
         Left            =   2745
         TabIndex        =   64
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   66
         Left            =   2295
         TabIndex        =   63
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   65
         Left            =   1845
         TabIndex        =   62
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   64
         Left            =   1395
         TabIndex        =   61
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   63
         Left            =   945
         TabIndex        =   60
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   62
         Left            =   495
         TabIndex        =   59
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   61
         Left            =   9255
         TabIndex        =   58
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   60
         Left            =   8820
         TabIndex        =   57
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   59
         Left            =   8370
         TabIndex        =   56
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   58
         Left            =   7920
         TabIndex        =   55
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   57
         Left            =   7470
         TabIndex        =   54
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   56
         Left            =   7020
         TabIndex        =   53
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   55
         Left            =   6570
         TabIndex        =   52
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   54
         Left            =   6120
         TabIndex        =   51
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   53
         Left            =   5670
         TabIndex        =   50
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   52
         Left            =   5220
         TabIndex        =   49
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   51
         Left            =   4770
         TabIndex        =   48
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   50
         Left            =   4320
         TabIndex        =   47
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   49
         Left            =   3870
         TabIndex        =   46
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   48
         Left            =   3420
         TabIndex        =   45
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   47
         Left            =   2970
         TabIndex        =   44
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   46
         Left            =   2520
         TabIndex        =   43
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   45
         Left            =   2070
         TabIndex        =   42
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   44
         Left            =   1620
         TabIndex        =   41
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   43
         Left            =   1170
         TabIndex        =   40
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   42
         Left            =   720
         TabIndex        =   39
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   41
         Left            =   270
         TabIndex        =   38
         Top             =   495
         Width           =   375
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   40
         Left            =   8895
         TabIndex        =   37
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   39
         Left            =   8445
         TabIndex        =   36
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   38
         Left            =   7995
         TabIndex        =   35
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   37
         Left            =   7545
         TabIndex        =   34
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   36
         Left            =   7080
         TabIndex        =   33
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   35
         Left            =   6645
         TabIndex        =   32
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   34
         Left            =   6195
         TabIndex        =   31
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   33
         Left            =   5760
         TabIndex        =   30
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   32
         Left            =   5295
         TabIndex        =   29
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   31
         Left            =   4845
         TabIndex        =   28
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   30
         Left            =   4395
         TabIndex        =   27
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   29
         Left            =   3990
         TabIndex        =   26
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   27
         Left            =   3075
         TabIndex        =   25
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   26
         Left            =   2595
         TabIndex        =   24
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   25
         Left            =   2145
         TabIndex        =   23
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   24
         Left            =   1695
         TabIndex        =   22
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   23
         Left            =   1215
         TabIndex        =   21
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   22
         Left            =   825
         TabIndex        =   20
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   28
         Left            =   3525
         TabIndex        =   19
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   21
         Left            =   330
         TabIndex        =   18
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   81
         Left            =   9000
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   79
         Left            =   8100
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   78
         Left            =   7650
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   77
         Left            =   7200
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   76
         Left            =   6750
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   75
         Left            =   6300
         Shape           =   3  'Circle
         Top             =   90
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   74
         Left            =   5850
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   73
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   72
         Left            =   4950
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   71
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   70
         Left            =   4050
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   69
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   68
         Left            =   3150
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   67
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   66
         Left            =   2250
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   65
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   64
         Left            =   1350
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   63
         Left            =   900
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   62
         Left            =   450
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   61
         Left            =   9210
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   60
         Left            =   8775
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   59
         Left            =   8325
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   58
         Left            =   7875
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   57
         Left            =   7420
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   56
         Left            =   6975
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   55
         Left            =   6525
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   54
         Left            =   6075
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   53
         Left            =   5625
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   52
         Left            =   5175
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   51
         Left            =   4725
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   50
         Left            =   4275
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   49
         Left            =   3825
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   48
         Left            =   3375
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   47
         Left            =   2925
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   46
         Left            =   2475
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   45
         Left            =   2025
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   44
         Left            =   1575
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   43
         Left            =   1125
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   42
         Left            =   675
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   41
         Left            =   225
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   80
         Left            =   8550
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin Threed.SSPanel SSPanel3 
      Height          =   1170
      Left            =   1845
      TabIndex        =   80
      Top             =   2970
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   2064
      _Version        =   196609
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   122
         Left            =   9045
         TabIndex        =   142
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   121
         Left            =   8595
         TabIndex        =   141
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   120
         Left            =   8145
         TabIndex        =   140
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   119
         Left            =   7680
         TabIndex        =   139
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   118
         Left            =   7245
         TabIndex        =   138
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   117
         Left            =   6795
         TabIndex        =   137
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   116
         Left            =   6345
         TabIndex        =   136
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   115
         Left            =   5895
         TabIndex        =   135
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   114
         Left            =   5445
         TabIndex        =   134
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   113
         Left            =   4995
         TabIndex        =   133
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   112
         Left            =   4545
         TabIndex        =   132
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   111
         Left            =   4095
         TabIndex        =   131
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   110
         Left            =   3645
         TabIndex        =   130
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   109
         Left            =   3195
         TabIndex        =   129
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   108
         Left            =   2745
         TabIndex        =   128
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   107
         Left            =   2295
         TabIndex        =   127
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   106
         Left            =   1845
         TabIndex        =   126
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   105
         Left            =   1395
         TabIndex        =   125
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   104
         Left            =   945
         TabIndex        =   124
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   103
         Left            =   495
         TabIndex        =   123
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   102
         Left            =   9255
         TabIndex        =   122
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   101
         Left            =   8820
         TabIndex        =   121
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   100
         Left            =   8370
         TabIndex        =   120
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   99
         Left            =   7920
         TabIndex        =   119
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   98
         Left            =   7470
         TabIndex        =   118
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   97
         Left            =   7020
         TabIndex        =   117
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   96
         Left            =   6570
         TabIndex        =   116
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   95
         Left            =   6120
         TabIndex        =   115
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   94
         Left            =   5670
         TabIndex        =   114
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   93
         Left            =   5220
         TabIndex        =   113
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   92
         Left            =   4770
         TabIndex        =   112
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   91
         Left            =   4320
         TabIndex        =   111
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   90
         Left            =   3870
         TabIndex        =   110
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   89
         Left            =   3420
         TabIndex        =   109
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   88
         Left            =   2970
         TabIndex        =   108
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   87
         Left            =   2520
         TabIndex        =   107
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   86
         Left            =   2070
         TabIndex        =   106
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   85
         Left            =   1620
         TabIndex        =   105
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   84
         Left            =   1170
         TabIndex        =   104
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   83
         Left            =   720
         TabIndex        =   103
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   82
         Left            =   270
         TabIndex        =   102
         Top             =   495
         Width           =   375
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   62
         Left            =   9300
         TabIndex        =   101
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   61
         Left            =   8895
         TabIndex        =   100
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   60
         Left            =   8445
         TabIndex        =   99
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   59
         Left            =   7995
         TabIndex        =   98
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   58
         Left            =   7545
         TabIndex        =   97
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   57
         Left            =   7080
         TabIndex        =   96
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   56
         Left            =   6645
         TabIndex        =   95
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   55
         Left            =   6195
         TabIndex        =   94
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   54
         Left            =   5760
         TabIndex        =   93
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   53
         Left            =   5295
         TabIndex        =   92
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   52
         Left            =   4845
         TabIndex        =   91
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   51
         Left            =   4395
         TabIndex        =   90
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   50
         Left            =   3990
         TabIndex        =   89
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   48
         Left            =   3075
         TabIndex        =   88
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   47
         Left            =   2595
         TabIndex        =   87
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   46
         Left            =   2145
         TabIndex        =   86
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   45
         Left            =   1695
         TabIndex        =   85
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   44
         Left            =   1215
         TabIndex        =   84
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   43
         Left            =   825
         TabIndex        =   83
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   49
         Left            =   3525
         TabIndex        =   82
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   42
         Left            =   330
         TabIndex        =   81
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   122
         Left            =   9000
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   121
         Left            =   8550
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000008&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   82
         Left            =   225
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   83
         Left            =   675
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   84
         Left            =   1125
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   85
         Left            =   1575
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   86
         Left            =   2025
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   87
         Left            =   2475
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   88
         Left            =   2925
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   89
         Left            =   3375
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   90
         Left            =   3825
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   91
         Left            =   4275
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   92
         Left            =   4725
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   93
         Left            =   5175
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   94
         Left            =   5625
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   95
         Left            =   6075
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   96
         Left            =   6525
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   97
         Left            =   6975
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   98
         Left            =   7420
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   99
         Left            =   7875
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   100
         Left            =   8325
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   101
         Left            =   8775
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   102
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   103
         Left            =   450
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   104
         Left            =   900
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   105
         Left            =   1350
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   106
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   107
         Left            =   2250
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   108
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   109
         Left            =   3150
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   110
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   111
         Left            =   4050
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   112
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   113
         Left            =   4950
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   114
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   115
         Left            =   5850
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   116
         Left            =   6300
         Shape           =   3  'Circle
         Top             =   90
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   117
         Left            =   6750
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   118
         Left            =   7200
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   119
         Left            =   7650
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   120
         Left            =   8100
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin Threed.SSPanel SSPanel4 
      Height          =   1170
      Left            =   1845
      TabIndex        =   143
      Top             =   4140
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   2064
      _Version        =   196609
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   163
         Left            =   9045
         TabIndex        =   205
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   162
         Left            =   8595
         TabIndex        =   204
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   161
         Left            =   8145
         TabIndex        =   203
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   160
         Left            =   7680
         TabIndex        =   202
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   159
         Left            =   7245
         TabIndex        =   201
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   158
         Left            =   6795
         TabIndex        =   200
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   157
         Left            =   6345
         TabIndex        =   199
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   156
         Left            =   5895
         TabIndex        =   198
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   155
         Left            =   5445
         TabIndex        =   197
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   154
         Left            =   4995
         TabIndex        =   196
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   153
         Left            =   4545
         TabIndex        =   195
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   152
         Left            =   4095
         TabIndex        =   194
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   151
         Left            =   3645
         TabIndex        =   193
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   150
         Left            =   3195
         TabIndex        =   192
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   149
         Left            =   2745
         TabIndex        =   191
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   148
         Left            =   2295
         TabIndex        =   190
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   147
         Left            =   1845
         TabIndex        =   189
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   146
         Left            =   1395
         TabIndex        =   188
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   145
         Left            =   945
         TabIndex        =   187
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   144
         Left            =   495
         TabIndex        =   186
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   143
         Left            =   9255
         TabIndex        =   185
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   142
         Left            =   8820
         TabIndex        =   184
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   141
         Left            =   8370
         TabIndex        =   183
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   140
         Left            =   7920
         TabIndex        =   182
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   139
         Left            =   7470
         TabIndex        =   181
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   138
         Left            =   7020
         TabIndex        =   180
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   137
         Left            =   6570
         TabIndex        =   179
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   136
         Left            =   6120
         TabIndex        =   178
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   135
         Left            =   5670
         TabIndex        =   177
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   134
         Left            =   5220
         TabIndex        =   176
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   133
         Left            =   4770
         TabIndex        =   175
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   132
         Left            =   4320
         TabIndex        =   174
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   131
         Left            =   3870
         TabIndex        =   173
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   130
         Left            =   3420
         TabIndex        =   172
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   129
         Left            =   2970
         TabIndex        =   171
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   128
         Left            =   2520
         TabIndex        =   170
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   127
         Left            =   2070
         TabIndex        =   169
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   126
         Left            =   1620
         TabIndex        =   168
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   125
         Left            =   1170
         TabIndex        =   167
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   124
         Left            =   720
         TabIndex        =   166
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   123
         Left            =   270
         TabIndex        =   165
         Top             =   495
         Width           =   375
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   83
         Left            =   9300
         TabIndex        =   164
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   82
         Left            =   8895
         TabIndex        =   163
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   81
         Left            =   8445
         TabIndex        =   162
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   80
         Left            =   7995
         TabIndex        =   161
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   79
         Left            =   7545
         TabIndex        =   160
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   78
         Left            =   7080
         TabIndex        =   159
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   77
         Left            =   6645
         TabIndex        =   158
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   76
         Left            =   6195
         TabIndex        =   157
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   75
         Left            =   5760
         TabIndex        =   156
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   74
         Left            =   5295
         TabIndex        =   155
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   73
         Left            =   4845
         TabIndex        =   154
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   72
         Left            =   4395
         TabIndex        =   153
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   71
         Left            =   3990
         TabIndex        =   152
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   69
         Left            =   3075
         TabIndex        =   151
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   68
         Left            =   2595
         TabIndex        =   150
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   67
         Left            =   2145
         TabIndex        =   149
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   66
         Left            =   1695
         TabIndex        =   148
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   65
         Left            =   1215
         TabIndex        =   147
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   64
         Left            =   825
         TabIndex        =   146
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   70
         Left            =   3525
         TabIndex        =   145
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   63
         Left            =   330
         TabIndex        =   144
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   163
         Left            =   9000
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   162
         Left            =   8550
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000008&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   123
         Left            =   225
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   124
         Left            =   675
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   125
         Left            =   1125
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   126
         Left            =   1575
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   127
         Left            =   2025
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   128
         Left            =   2475
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   129
         Left            =   2925
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   130
         Left            =   3375
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   131
         Left            =   3825
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   132
         Left            =   4275
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   133
         Left            =   4725
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   134
         Left            =   5175
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   135
         Left            =   5625
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   136
         Left            =   6075
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   137
         Left            =   6525
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   138
         Left            =   6975
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   139
         Left            =   7420
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   140
         Left            =   7875
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   141
         Left            =   8325
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   142
         Left            =   8775
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   143
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   144
         Left            =   450
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   145
         Left            =   900
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   146
         Left            =   1350
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   147
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   148
         Left            =   2250
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   149
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   150
         Left            =   3150
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   151
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   152
         Left            =   4050
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   153
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   154
         Left            =   4950
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   155
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   156
         Left            =   5850
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   157
         Left            =   6300
         Shape           =   3  'Circle
         Top             =   90
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   158
         Left            =   6750
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   159
         Left            =   7200
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   160
         Left            =   7650
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   161
         Left            =   8100
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin Threed.SSPanel SSPanel5 
      Height          =   1170
      Left            =   1845
      TabIndex        =   206
      Top             =   5310
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   2064
      _Version        =   196609
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   204
         Left            =   9045
         TabIndex        =   268
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   203
         Left            =   8595
         TabIndex        =   267
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   202
         Left            =   8145
         TabIndex        =   266
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   201
         Left            =   7680
         TabIndex        =   265
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   200
         Left            =   7245
         TabIndex        =   264
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   199
         Left            =   6795
         TabIndex        =   263
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   198
         Left            =   6345
         TabIndex        =   262
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   197
         Left            =   5895
         TabIndex        =   261
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   196
         Left            =   5445
         TabIndex        =   260
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   195
         Left            =   4995
         TabIndex        =   259
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   194
         Left            =   4545
         TabIndex        =   258
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   193
         Left            =   4095
         TabIndex        =   257
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   192
         Left            =   3645
         TabIndex        =   256
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   191
         Left            =   3180
         TabIndex        =   255
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   190
         Left            =   2745
         TabIndex        =   254
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   189
         Left            =   2295
         TabIndex        =   253
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   188
         Left            =   1845
         TabIndex        =   252
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   187
         Left            =   1395
         TabIndex        =   251
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   186
         Left            =   945
         TabIndex        =   250
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   185
         Left            =   495
         TabIndex        =   249
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   184
         Left            =   9255
         TabIndex        =   248
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   183
         Left            =   8820
         TabIndex        =   247
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   182
         Left            =   8370
         TabIndex        =   246
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   181
         Left            =   7920
         TabIndex        =   245
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   180
         Left            =   7470
         TabIndex        =   244
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   179
         Left            =   7020
         TabIndex        =   243
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   178
         Left            =   6570
         TabIndex        =   242
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   177
         Left            =   6120
         TabIndex        =   241
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   176
         Left            =   5670
         TabIndex        =   240
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   175
         Left            =   5220
         TabIndex        =   239
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   174
         Left            =   4770
         TabIndex        =   238
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   173
         Left            =   4320
         TabIndex        =   237
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   172
         Left            =   3870
         TabIndex        =   236
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   171
         Left            =   3420
         TabIndex        =   235
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   170
         Left            =   2970
         TabIndex        =   234
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   169
         Left            =   2520
         TabIndex        =   233
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   168
         Left            =   2070
         TabIndex        =   232
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   167
         Left            =   1620
         TabIndex        =   231
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   166
         Left            =   1170
         TabIndex        =   230
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   165
         Left            =   720
         TabIndex        =   229
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   164
         Left            =   270
         TabIndex        =   228
         Top             =   495
         Width           =   375
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   204
         Left            =   9000
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   203
         Left            =   8550
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   164
         Left            =   225
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   165
         Left            =   675
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   166
         Left            =   1125
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   167
         Left            =   1575
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   168
         Left            =   2025
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   169
         Left            =   2475
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   170
         Left            =   2925
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   171
         Left            =   3375
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   172
         Left            =   3825
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   173
         Left            =   4275
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   174
         Left            =   4725
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   175
         Left            =   5175
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   176
         Left            =   5625
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   177
         Left            =   6075
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   178
         Left            =   6525
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   179
         Left            =   6975
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   180
         Left            =   7420
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   181
         Left            =   7875
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   182
         Left            =   8325
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   183
         Left            =   8775
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   184
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   185
         Left            =   450
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   186
         Left            =   900
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   187
         Left            =   1350
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   188
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   189
         Left            =   2250
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   190
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   191
         Left            =   3150
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   192
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   193
         Left            =   4050
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   194
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   195
         Left            =   4950
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   196
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   197
         Left            =   5850
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   198
         Left            =   6300
         Shape           =   3  'Circle
         Top             =   90
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   199
         Left            =   6750
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   200
         Left            =   7200
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   201
         Left            =   7650
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   202
         Left            =   8100
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   104
         Left            =   9300
         TabIndex        =   227
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   103
         Left            =   8895
         TabIndex        =   226
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   102
         Left            =   8445
         TabIndex        =   225
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   101
         Left            =   7995
         TabIndex        =   224
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   100
         Left            =   7545
         TabIndex        =   223
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   99
         Left            =   7080
         TabIndex        =   222
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   98
         Left            =   6645
         TabIndex        =   221
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   97
         Left            =   6195
         TabIndex        =   220
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   96
         Left            =   5760
         TabIndex        =   219
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   95
         Left            =   5295
         TabIndex        =   218
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   94
         Left            =   4845
         TabIndex        =   217
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   93
         Left            =   4395
         TabIndex        =   216
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   92
         Left            =   3990
         TabIndex        =   215
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   90
         Left            =   3075
         TabIndex        =   214
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   89
         Left            =   2595
         TabIndex        =   213
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   88
         Left            =   2145
         TabIndex        =   212
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   87
         Left            =   1695
         TabIndex        =   211
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   86
         Left            =   1215
         TabIndex        =   210
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   85
         Left            =   825
         TabIndex        =   209
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   91
         Left            =   3525
         TabIndex        =   208
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   84
         Left            =   330
         TabIndex        =   207
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
   End
   Begin Threed.SSPanel SSPanel6 
      Height          =   1170
      Left            =   1845
      TabIndex        =   269
      Top             =   6480
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   2064
      _Version        =   196609
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   245
         Left            =   9045
         TabIndex        =   331
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   244
         Left            =   8595
         TabIndex        =   330
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   243
         Left            =   8145
         TabIndex        =   329
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   242
         Left            =   7680
         TabIndex        =   328
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   241
         Left            =   7245
         TabIndex        =   327
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   240
         Left            =   6795
         TabIndex        =   326
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   239
         Left            =   6345
         TabIndex        =   325
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   238
         Left            =   5895
         TabIndex        =   324
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   237
         Left            =   5445
         TabIndex        =   323
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   236
         Left            =   4995
         TabIndex        =   322
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   235
         Left            =   4545
         TabIndex        =   321
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   234
         Left            =   4095
         TabIndex        =   320
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   233
         Left            =   3645
         TabIndex        =   319
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   232
         Left            =   3180
         TabIndex        =   318
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   231
         Left            =   2745
         TabIndex        =   317
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   230
         Left            =   2295
         TabIndex        =   316
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   229
         Left            =   1845
         TabIndex        =   315
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   228
         Left            =   1395
         TabIndex        =   314
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   227
         Left            =   945
         TabIndex        =   313
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   226
         Left            =   495
         TabIndex        =   312
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   225
         Left            =   9255
         TabIndex        =   311
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   224
         Left            =   8820
         TabIndex        =   310
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   223
         Left            =   8370
         TabIndex        =   309
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   222
         Left            =   7920
         TabIndex        =   308
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   221
         Left            =   7470
         TabIndex        =   307
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   220
         Left            =   7020
         TabIndex        =   306
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   219
         Left            =   6570
         TabIndex        =   305
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   218
         Left            =   6120
         TabIndex        =   304
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   217
         Left            =   5670
         TabIndex        =   303
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   216
         Left            =   5220
         TabIndex        =   302
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   215
         Left            =   4770
         TabIndex        =   301
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   214
         Left            =   4320
         TabIndex        =   300
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   213
         Left            =   3870
         TabIndex        =   299
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   212
         Left            =   3420
         TabIndex        =   298
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   211
         Left            =   2970
         TabIndex        =   297
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   210
         Left            =   2520
         TabIndex        =   296
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   209
         Left            =   2070
         TabIndex        =   295
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   208
         Left            =   1620
         TabIndex        =   294
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   207
         Left            =   1170
         TabIndex        =   293
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   206
         Left            =   720
         TabIndex        =   292
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   205
         Left            =   270
         TabIndex        =   291
         Top             =   495
         Width           =   375
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   125
         Left            =   9300
         TabIndex        =   290
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   124
         Left            =   8895
         TabIndex        =   289
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   123
         Left            =   8445
         TabIndex        =   288
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   122
         Left            =   7995
         TabIndex        =   287
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   121
         Left            =   7545
         TabIndex        =   286
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   120
         Left            =   7080
         TabIndex        =   285
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   119
         Left            =   6645
         TabIndex        =   284
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   118
         Left            =   6195
         TabIndex        =   283
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   117
         Left            =   5760
         TabIndex        =   282
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   116
         Left            =   5295
         TabIndex        =   281
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   115
         Left            =   4845
         TabIndex        =   280
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   114
         Left            =   4395
         TabIndex        =   279
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   113
         Left            =   3990
         TabIndex        =   278
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   111
         Left            =   3075
         TabIndex        =   277
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   110
         Left            =   2595
         TabIndex        =   276
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   109
         Left            =   2145
         TabIndex        =   275
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   108
         Left            =   1695
         TabIndex        =   274
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   107
         Left            =   1215
         TabIndex        =   273
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   106
         Left            =   825
         TabIndex        =   272
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   112
         Left            =   3525
         TabIndex        =   271
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   105
         Left            =   375
         TabIndex        =   270
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   245
         Left            =   9000
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   244
         Left            =   8550
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   205
         Left            =   225
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   206
         Left            =   675
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   207
         Left            =   1125
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   208
         Left            =   1575
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   209
         Left            =   2025
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   210
         Left            =   2475
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   211
         Left            =   2925
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   212
         Left            =   3375
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   213
         Left            =   3825
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   214
         Left            =   4275
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   215
         Left            =   4725
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   216
         Left            =   5175
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   217
         Left            =   5625
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   218
         Left            =   6075
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   219
         Left            =   6525
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   220
         Left            =   6975
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   221
         Left            =   7420
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   222
         Left            =   7875
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   223
         Left            =   8325
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   224
         Left            =   8775
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   225
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   226
         Left            =   450
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   227
         Left            =   900
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   228
         Left            =   1350
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   229
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   230
         Left            =   2250
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   231
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   232
         Left            =   3150
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   233
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   234
         Left            =   4050
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   235
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   236
         Left            =   4950
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   237
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   238
         Left            =   5850
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   239
         Left            =   6300
         Shape           =   3  'Circle
         Top             =   90
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   240
         Left            =   6750
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   241
         Left            =   7200
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   242
         Left            =   7650
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   243
         Left            =   8100
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin Threed.SSPanel SSPanel7 
      Height          =   1170
      Left            =   1845
      TabIndex        =   332
      Top             =   7650
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   2064
      _Version        =   196609
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   286
         Left            =   9045
         TabIndex        =   394
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   285
         Left            =   8595
         TabIndex        =   393
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   284
         Left            =   8160
         TabIndex        =   392
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   283
         Left            =   7680
         TabIndex        =   391
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   282
         Left            =   7245
         TabIndex        =   390
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   281
         Left            =   6795
         TabIndex        =   389
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   280
         Left            =   6345
         TabIndex        =   388
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   279
         Left            =   5895
         TabIndex        =   387
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   278
         Left            =   5445
         TabIndex        =   386
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   277
         Left            =   4995
         TabIndex        =   385
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   276
         Left            =   4545
         TabIndex        =   384
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   275
         Left            =   4095
         TabIndex        =   383
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   274
         Left            =   3645
         TabIndex        =   382
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   273
         Left            =   3180
         TabIndex        =   381
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   272
         Left            =   2745
         TabIndex        =   380
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   271
         Left            =   2295
         TabIndex        =   379
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   270
         Left            =   1845
         TabIndex        =   378
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   269
         Left            =   1395
         TabIndex        =   377
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   268
         Left            =   945
         TabIndex        =   376
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   267
         Left            =   465
         TabIndex        =   375
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   266
         Left            =   9255
         TabIndex        =   374
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   265
         Left            =   8820
         TabIndex        =   373
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   264
         Left            =   8370
         TabIndex        =   372
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   263
         Left            =   7920
         TabIndex        =   371
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   262
         Left            =   7470
         TabIndex        =   370
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   261
         Left            =   7020
         TabIndex        =   369
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   260
         Left            =   6570
         TabIndex        =   368
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   259
         Left            =   6120
         TabIndex        =   367
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   258
         Left            =   5670
         TabIndex        =   366
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   257
         Left            =   5220
         TabIndex        =   365
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   256
         Left            =   4770
         TabIndex        =   364
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   255
         Left            =   4320
         TabIndex        =   363
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   254
         Left            =   3870
         TabIndex        =   362
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   253
         Left            =   3420
         TabIndex        =   361
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   252
         Left            =   2970
         TabIndex        =   360
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   251
         Left            =   2520
         TabIndex        =   359
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   250
         Left            =   2070
         TabIndex        =   358
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   249
         Left            =   1620
         TabIndex        =   357
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   248
         Left            =   1170
         TabIndex        =   356
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   247
         Left            =   720
         TabIndex        =   355
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   246
         Left            =   270
         TabIndex        =   354
         Top             =   495
         Width           =   375
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   146
         Left            =   9300
         TabIndex        =   353
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   145
         Left            =   8895
         TabIndex        =   352
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   144
         Left            =   8445
         TabIndex        =   351
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   143
         Left            =   7995
         TabIndex        =   350
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   142
         Left            =   7545
         TabIndex        =   349
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   141
         Left            =   7080
         TabIndex        =   348
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   140
         Left            =   6645
         TabIndex        =   347
         Top             =   945
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   139
         Left            =   6195
         TabIndex        =   346
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   138
         Left            =   5760
         TabIndex        =   345
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   137
         Left            =   5295
         TabIndex        =   344
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   136
         Left            =   4845
         TabIndex        =   343
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   135
         Left            =   4395
         TabIndex        =   342
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   134
         Left            =   3990
         TabIndex        =   341
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   132
         Left            =   3075
         TabIndex        =   340
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   131
         Left            =   2595
         TabIndex        =   339
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   130
         Left            =   2145
         TabIndex        =   338
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   129
         Left            =   1695
         TabIndex        =   337
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   128
         Left            =   1215
         TabIndex        =   336
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   127
         Left            =   825
         TabIndex        =   335
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   133
         Left            =   3525
         TabIndex        =   334
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   126
         Left            =   375
         TabIndex        =   333
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   286
         Left            =   9000
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   285
         Left            =   8550
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   246
         Left            =   225
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   247
         Left            =   675
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   248
         Left            =   1125
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   249
         Left            =   1575
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   250
         Left            =   2025
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   251
         Left            =   2475
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   252
         Left            =   2925
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   253
         Left            =   3375
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   254
         Left            =   3825
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   255
         Left            =   4275
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   256
         Left            =   4725
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   257
         Left            =   5175
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   258
         Left            =   5625
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   259
         Left            =   6075
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   260
         Left            =   6525
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   261
         Left            =   6975
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   262
         Left            =   7420
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   263
         Left            =   7875
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   264
         Left            =   8325
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   265
         Left            =   8775
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   266
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   267
         Left            =   450
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   268
         Left            =   900
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   269
         Left            =   1350
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   270
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   271
         Left            =   2250
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   272
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   273
         Left            =   3150
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   274
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   275
         Left            =   4050
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   276
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   277
         Left            =   4950
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   278
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   279
         Left            =   5850
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   280
         Left            =   6300
         Shape           =   3  'Circle
         Top             =   90
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   281
         Left            =   6750
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   282
         Left            =   7200
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   283
         Left            =   7650
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   284
         Left            =   8100
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   1170
      Left            =   1845
      TabIndex        =   395
      Top             =   630
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   2064
      _Version        =   196609
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
      BorderWidth     =   1
      BevelInner      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   40
         Left            =   9045
         TabIndex        =   457
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   39
         Left            =   8595
         TabIndex        =   456
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   38
         Left            =   8160
         TabIndex        =   455
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   37
         Left            =   7680
         TabIndex        =   454
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   36
         Left            =   7245
         TabIndex        =   453
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   35
         Left            =   6795
         TabIndex        =   452
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   34
         Left            =   6345
         TabIndex        =   451
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   33
         Left            =   5895
         TabIndex        =   450
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   32
         Left            =   5445
         TabIndex        =   449
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   31
         Left            =   4995
         TabIndex        =   448
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   30
         Left            =   4545
         TabIndex        =   447
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   29
         Left            =   4095
         TabIndex        =   446
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   28
         Left            =   3645
         TabIndex        =   445
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   27
         Left            =   3180
         TabIndex        =   444
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   26
         Left            =   2745
         TabIndex        =   443
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   25
         Left            =   2295
         TabIndex        =   442
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   24
         Left            =   1845
         TabIndex        =   441
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   23
         Left            =   1395
         TabIndex        =   440
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   22
         Left            =   945
         TabIndex        =   439
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   21
         Left            =   495
         TabIndex        =   438
         Top             =   135
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   20
         Left            =   9270
         TabIndex        =   437
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   19
         Left            =   8820
         TabIndex        =   436
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   18
         Left            =   8370
         TabIndex        =   435
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   17
         Left            =   7920
         TabIndex        =   434
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   16
         Left            =   7470
         TabIndex        =   433
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   15
         Left            =   7020
         TabIndex        =   432
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   14
         Left            =   6570
         TabIndex        =   431
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   13
         Left            =   6120
         TabIndex        =   430
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   12
         Left            =   5670
         TabIndex        =   429
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   11
         Left            =   5220
         TabIndex        =   428
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   10
         Left            =   4770
         TabIndex        =   427
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   9
         Left            =   4320
         TabIndex        =   426
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   8
         Left            =   3870
         TabIndex        =   425
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   7
         Left            =   3420
         TabIndex        =   424
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   6
         Left            =   2970
         TabIndex        =   423
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   5
         Left            =   2520
         TabIndex        =   422
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   4
         Left            =   2070
         TabIndex        =   421
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   3
         Left            =   1620
         TabIndex        =   420
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   2
         Left            =   1170
         TabIndex        =   419
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   1
         Left            =   720
         TabIndex        =   418
         Top             =   495
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   360
         Index           =   0
         Left            =   270
         TabIndex        =   417
         Top             =   495
         Width           =   375
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   40
         Left            =   9000
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   39
         Left            =   8550
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   0
         Left            =   225
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   1
         Left            =   675
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   2
         Left            =   1125
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   3
         Left            =   1575
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   4
         Left            =   2025
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   5
         Left            =   2475
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   6
         Left            =   2925
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   7
         Left            =   3375
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   8
         Left            =   3825
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   9
         Left            =   4275
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   10
         Left            =   4725
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   11
         Left            =   5175
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   12
         Left            =   5625
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   13
         Left            =   6075
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   14
         Left            =   6525
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   15
         Left            =   6975
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   16
         Left            =   7420
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   17
         Left            =   7875
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   18
         Left            =   8325
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   19
         Left            =   8775
         Shape           =   3  'Circle
         Top             =   460
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   20
         Left            =   9240
         Shape           =   3  'Circle
         Top             =   465
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   21
         Left            =   450
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   22
         Left            =   900
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   23
         Left            =   1350
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   24
         Left            =   1800
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   25
         Left            =   2250
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   26
         Left            =   2700
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   27
         Left            =   3150
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   28
         Left            =   3600
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   29
         Left            =   4050
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   30
         Left            =   4500
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   31
         Left            =   4950
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   32
         Left            =   5400
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   33
         Left            =   5850
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   34
         Left            =   6300
         Shape           =   3  'Circle
         Top             =   90
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   35
         Left            =   6750
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   36
         Left            =   7200
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   37
         Left            =   7650
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Shape S_F1 
         BackColor       =   &H00000000&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   465
         Index           =   38
         Left            =   8100
         Shape           =   3  'Circle
         Top             =   75
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   20
         Left            =   9300
         TabIndex        =   416
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   19
         Left            =   8895
         TabIndex        =   415
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   18
         Left            =   8445
         TabIndex        =   414
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   17
         Left            =   7995
         TabIndex        =   413
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   16
         Left            =   7545
         TabIndex        =   412
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   15
         Left            =   7080
         TabIndex        =   411
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   14
         Left            =   6645
         TabIndex        =   410
         Top             =   945
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   13
         Left            =   6195
         TabIndex        =   409
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   12
         Left            =   5760
         TabIndex        =   408
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   11
         Left            =   5295
         TabIndex        =   407
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   10
         Left            =   4845
         TabIndex        =   406
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   9
         Left            =   4395
         TabIndex        =   405
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   8
         Left            =   3990
         TabIndex        =   404
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   6
         Left            =   3075
         TabIndex        =   403
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   5
         Left            =   2595
         TabIndex        =   402
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   4
         Left            =   2145
         TabIndex        =   401
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   3
         Left            =   1695
         TabIndex        =   400
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   2
         Left            =   1215
         TabIndex        =   399
         Top             =   930
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   1
         Left            =   825
         TabIndex        =   398
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   7
         Left            =   3525
         TabIndex        =   397
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label L_F1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Index           =   0
         Left            =   375
         TabIndex        =   396
         Top             =   930
         Visible         =   0   'False
         Width           =   135
      End
   End
End
Attribute VB_Name = "AGE2080C"
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
'-- Program Name      钢卷库存现状查询
'-- Program ID        AGE2080C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2003.7.23
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

Public sOrd_no As String
Public Active_CForm As String       'Form Active
Public Form_Wid As Double
Public Form_Len As Double

Dim rib_prev As Integer
Dim opt_prev As Integer
Dim ss1_curr As Long

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

       Call Gp_Ms_Collection(sdb_coil_cnt, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc2.Add Item:=ss2, Key:="Spread"

    rib_prev = 0
    opt_prev = 0
    ss1_curr = 0
    rib_area(0).Value = True
    opt_area(0).Value = True
    
    VScroll1.SmallChange = 1
    VScroll1.LargeChange = 7
    VScroll1.Value = 1
    VScroll1.Max = 0
    
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
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set rControl = Nothing
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    rib_area(0).Value = True
    opt_area(0).Value = True
    VScroll1.Max = 0
    
    Call Screen_Cls
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
End Sub

Public Sub Form_Ref()

    Dim sQuery As String
    Dim sMesg As String
    Dim sArea As String
    Dim sAddr As String
    Dim icount As Long
    Dim ddd As Long
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    
    If sMesg = "OK" Then
    
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then
        
            sAddr = Trim(Str(rib_prev + 1))
            
            Select Case opt_prev
                    Case 0
                        sArea = "A"
                    Case 1
                        sArea = "B"
                    Case 2
                        sArea = "C"
                    Case 3
                        sArea = "D"
            End Select
        
            ss1_curr = 1
            VScroll1.Max = 0
            sdb_coil_cnt.Value = 0
            Call Screen_Cls
            
            'COIL YARD STANDARD
            sQuery = "         SELECT LOCATION, NVL(MAX_CNT,0) "
            sQuery = sQuery + "  FROM FP_STDYARD "
            sQuery = sQuery + " WHERE SUBSTR(LOCATION,1,3) = 'C" + sAddr + sArea + "'"
            sQuery = sQuery + " ORDER BY LOCATION "
                     
            If Gf_Only_Display(M_CN1, sc1, sQuery) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                MDIMain.MenuTool.Buttons(14).Enabled = False
                
                sdb_coil_cnt.Value = Gf_FloatFind(M_CN1, _
                                     "SELECT COUNT(*) FROM GP_COILYARD WHERE SUBSTR(YARD_ADDR,2,2) = '" _
                                      + sAddr + sArea + "' AND NVL(COIL_NO,' ') <> ' ' ")
                                      
                If sArea = "B" Then
                   sdb_coil_cnt.Value = sdb_coil_cnt.Value + 18
                End If
                
                VScroll1.Max = Gf_FloatFind(M_CN1, _
                                     "SELECT COUNT(*) FROM FP_STDYARD WHERE SUBSTR(LOCATION,1,3) = 'C" _
                                      + sAddr + sArea + "' ")
            Else
                Exit Sub
            End If
        
            'COIL YARD
            sQuery = "         SELECT B.YARD_ADDR, B.COIL_LAYER, B.COIL_SEQ, A.INDIA, A.OUTDIA, A.WGT, B.COIL_NO "
            sQuery = sQuery + "  FROM GP_COIL A, GP_COILYARD B, FP_STDYARD C "
            sQuery = sQuery + " WHERE SUBSTR(B.YARD_ADDR,2,2) = '" + sAddr + sArea + "'"
            sQuery = sQuery + "   AND NVL(B.COIL_NO,' ') <> ' ' "
            sQuery = sQuery + "   AND B.COIL_NO = A.COIL_NO(+) "
            sQuery = sQuery + "   AND B.YARD_ADDR = C.LOCATION(+) "
            sQuery = sQuery + " ORDER BY B.YARD_ADDR, B.COIL_LAYER, B.COIL_SEQ "
                     
            Call Gf_Only_Display(M_CN1, sc2, sQuery, , False)
            
            'FIRST START
            Call Screen_Display(ss1_curr)
            
        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
    If sAddr = "1" And sArea = "B" Then
    
       sdb_coil_cnt.Value = sdb_coil_cnt.Value - 18
    
       For icount = 215 To 218
           Label1(icount).Visible = False
           S_F1(icount).Visible = False
       Next icount
       For icount = 235 To 239
           Label1(icount).Visible = False
           S_F1(icount).Visible = False
       Next icount
       
       For icount = 256 To 259
           Label1(icount).Visible = False
           S_F1(icount).Visible = False
       Next icount
       For icount = 276 To 280
           Label1(icount).Visible = False
           S_F1(icount).Visible = False
       Next icount
       
       For icount = 115 To 118
           L_F1(icount).Visible = False
       Next icount
       For icount = 136 To 139
           L_F1(icount).Visible = False
       Next icount
       
    End If
 
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub



Private Sub opt_area_Click(Index As Integer, Value As Integer)

    If opt_area(Index) Then
        opt_area(Index).ForeColor = &HFF&
        If Index <> opt_prev Then
            opt_area(opt_prev).ForeColor = &H80000012
        End If
        opt_prev = Index
    End If

    
End Sub

Private Sub pan_addr_Click(Index As Integer)

    If pan_addr(Index).ForeColor = &HFF8080 Then Exit Sub
    
    Load AGE2010C
    
    AGE2010C.sOth = "AGE2080C"
    AGE2010C.TXT_S_YARD_ADDR = Trim(pan_addr(Index).Caption)
    
    AGE2010C.Show
    AGE2010C.SetFocus
    
End Sub

Private Sub rib_area_Click(Index As Integer, Value As Integer)

    If rib_area(Index) Then
        rib_area(Index).BackColor = &H80FFFF
        rib_area(Index).ForeColor = &HFF&
        If Index <> rib_prev Then
            rib_area(rib_prev).BackColor = &HE0E0E0
            rib_area(rib_prev).ForeColor = &H80000012
        End If
    Else
        rib_area(Index).BackColor = &HE0E0E0
        rib_area(Index).ForeColor = &H80000012
    End If
    
    rib_prev = Index
    
    opt_area(opt_prev).Value = False
    opt_area(opt_prev).ForeColor = &H80000012
    
    opt_area(0).Enabled = True
    opt_area(0).Value = True
    opt_area(1).Enabled = True
    opt_area(2).Enabled = True
    opt_area(3).Enabled = True
    
End Sub

Public Sub Screen_Display(Curr_Row As Long)

    Dim lCnt As Long
    Dim lCount As Long
    
    Dim sArea As String
    Dim sAddr As String
    Dim sTemp As String
    
    If Curr_Row = 0 Then Curr_Row = 1
    If Curr_Row + 6 > ss1.MaxRows Then Exit Sub
    
    Call Screen_Cls
    
    sAddr = Trim(Str(rib_prev + 1))
    
    Select Case opt_prev
        Case 0
            sArea = "A"
        Case 1
            sArea = "B"
        Case 2
            sArea = "C"
        Case 3
            sArea = "D"
    End Select
        
    Screen.MousePointer = vbHourglass
    
    lCount = 0
    For lCnt = Curr_Row To Curr_Row + 6
        ss1.Row = lCnt
        ss1.Col = 1
        pan_addr(lCount).Caption = ss1.Text    'YARD ADDR
        pan_addr(lCount).ForeColor = &HFF8080
        ss1.Col = 2
        If ss1.Text = "" Then ss1.Text = "0"
        Call MaxCoil_Display(lCount, Val(ss1.Text))
        lCount = lCount + 1
        ss1_curr = lCnt - 1
    Next lCnt
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub MaxCoil_Display(lPanNum As Long, Max_Coil As Long)

    Dim lCnt As Long
    Dim lCoilMax As Long
    Dim lLayer1 As Long
    Dim lLayer2 As Long
    
    Select Case lPanNum
        Case 0
            lCoilMax = 0
            lLayer1 = 0
            lLayer2 = 21
        Case 1
            lCoilMax = 21
            lLayer1 = 41
            lLayer2 = 62
        Case 2
            lCoilMax = 42
            lLayer1 = 82
            lLayer2 = 103
        Case 3
            lCoilMax = 63
            lLayer1 = 123
            lLayer2 = 144
        Case 4
            lCoilMax = 84
            lLayer1 = 164
            lLayer2 = 185
        Case 5
            lCoilMax = 105
            lLayer1 = 205
            lLayer2 = 226
        Case 6
            lCoilMax = 126
            lLayer1 = 246
            lLayer2 = 267
    End Select
    
    'Max Coil Seq
    For lCnt = lCoilMax To lCoilMax + Max_Coil - 1
        L_F1(lCnt).Visible = True
    Next lCnt
    
    'Coil Yard Layer 1
    For lCnt = lLayer1 To lLayer1 + Max_Coil - 1
        S_F1(lCnt).Visible = True
        S_F1(lCnt).BorderStyle = 3
        S_F1(lCnt).FillColor = &HE0E0E0
        S_F1(lCnt).BorderColor = &H808080
    Next lCnt

    'Coil Yard Layer 2
    For lCnt = lLayer2 To lLayer2 + Max_Coil - 2
        S_F1(lCnt).Visible = True
        S_F1(lCnt).BorderStyle = 3
        S_F1(lCnt).FillColor = &HE0E0E0
        S_F1(lCnt).BorderColor = &H808080
    Next lCnt
    
    Call Coil_Display(pan_addr(lPanNum).Caption, lLayer1, lLayer2, lPanNum)
    
End Sub

Public Sub Coil_Display(Addr_No As String, lLayer1 As Long, lLayer2 As Long, lPanNum As Long)

    Dim lCnt As Long
    Dim sSeq As String
    Dim sWgt As String
    Dim sIndia As String
    Dim sOutdia As String
    Dim sCoil_No As String

    For lCnt = 1 To ss2.MaxRows
        
        ss2.Row = lCnt
        ss2.Col = 1
        
        If Trim(Addr_No) = ss2.Text Then
        
            pan_addr(lPanNum).ForeColor = &HFF0000
            ss2.Col = 3     'Seq
            sSeq = ss2.Text
            ss2.Col = 4     'INDIA
            sIndia = ss2.Text
            ss2.Col = 5     'OUTDIA
            sOutdia = ss2.Text
            ss2.Col = 6     'WGT
            sWgt = ss2.Text
            ss2.Col = 7     'Coil_No
            sCoil_No = ss2.Text
        
            ss2.Col = 2
            If Trim(ss2.Text) = "1" Then  ' Layer 1
                S_F1(lLayer1 + Val(sSeq) - 1).BorderStyle = 1
                S_F1(lLayer1 + Val(sSeq) - 1).FillColor = &HC0C0FF
                S_F1(lLayer1 + Val(sSeq) - 1).BorderColor = &H80000008
                Label1(lLayer1 + Val(sSeq) - 1).ToolTipText = "Coil No. " & _
                                                              sCoil_No & " (" & _
                                                              sIndia & " / " & _
                                                              sOutdia & " / " & _
                                                              sWgt & ")"
            Else
                S_F1(lLayer2 + Val(sSeq) - 1).BorderStyle = 1
                S_F1(lLayer2 + Val(sSeq) - 1).FillColor = &HC0C0FF
                S_F1(lLayer2 + Val(sSeq) - 1).BorderColor = &H80000008
                Label1(lLayer2 + Val(sSeq) - 1).ToolTipText = "Coil No. " & _
                                                              sCoil_No & " (" & _
                                                              sIndia & " / " & _
                                                              sOutdia & " / " & _
                                                              sWgt & ")"
            End If
        
        End If
    
    Next lCnt

End Sub

Public Sub Screen_Cls()

    Dim lCount As Long
    
    For lCount = 0 To 286
        S_F1(lCount).Visible = False
        Label1(lCount).ToolTipText = ""
    Next lCount
    
    For lCount = 0 To 146
        L_F1(lCount).Visible = False
    Next lCount
    
    For lCount = 0 To 6
        pan_addr(lCount).Caption = ""
        pan_addr(lCount).ForeColor = &HFF8080
    Next lCount
    
End Sub

Private Sub VScroll1_Change()

    Call Screen_Display(VScroll1.Value)
    
End Sub
