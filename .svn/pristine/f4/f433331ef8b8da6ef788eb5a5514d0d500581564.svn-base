VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGA2090C 
   Caption         =   "外来板坯实绩录入_CGA2090C"
   ClientHeight    =   6075
   ClientLeft      =   660
   ClientTop       =   1710
   ClientWidth     =   13605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   13605
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand comm_slab 
      Height          =   405
      Left            =   150
      TabIndex        =   58
      Top             =   3480
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确认"
   End
   Begin Threed.SSCommand cmd_get_info 
      Height          =   675
      Left            =   5220
      TabIndex        =   57
      Top             =   720
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1191
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确认"
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   615
      Left            =   150
      TabIndex        =   50
      Top             =   30
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   1085
      _Version        =   196609
      BackColor       =   14737632
      Caption         =   "入库分类"
      Begin VB.TextBox txt_InPltCo 
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
         Left            =   4515
         MaxLength       =   6
         TabIndex        =   54
         Top             =   210
         Width           =   1110
      End
      Begin VB.TextBox txt_InPltCoDesc 
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
         Height          =   315
         Left            =   5610
         MaxLength       =   11
         TabIndex        =   53
         Top             =   210
         Width           =   1920
      End
      Begin Threed.SSOption opt_BuyCo 
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   210
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   609
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
         Caption         =   "老炼钢"
      End
      Begin Threed.SSOption opt_BuyCo 
         Height          =   345
         Index           =   1
         Left            =   1230
         TabIndex        =   52
         Top             =   210
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
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
         Caption         =   "外购坯"
      End
      Begin InDate.ULabel lab_InPltCo 
         Height          =   315
         Left            =   3510
         Top             =   210
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         Caption         =   "外购公司"
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
      Begin Threed.SSOption opt_BuyCo 
         Height          =   345
         Index           =   2
         Left            =   2310
         TabIndex        =   56
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   609
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
         Caption         =   "委托加工"
      End
   End
   Begin VB.TextBox txt_cen 
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
      Left            =   4440
      MaxLength       =   2
      TabIndex        =   49
      Top             =   1080
      Width           =   675
   End
   Begin VB.TextBox txt_old_slabno 
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
      Left            =   6570
      MaxLength       =   10
      TabIndex        =   47
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_seq 
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
      Left            =   4440
      MaxLength       =   4
      TabIndex        =   3
      Top             =   690
      Width           =   675
   End
   Begin VB.TextBox txt_heat 
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
      Left            =   2910
      MaxLength       =   1
      TabIndex        =   2
      Top             =   690
      Width           =   375
   End
   Begin VB.TextBox txt_mon 
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
      Left            =   1740
      MaxLength       =   2
      TabIndex        =   1
      Top             =   690
      Width           =   375
   End
   Begin VB.TextBox txt_car_no 
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
      Left            =   11700
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1470
      Width           =   915
   End
   Begin VB.TextBox txt_act_stlgrd_dec 
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
      Left            =   2640
      MaxLength       =   11
      TabIndex        =   42
      Top             =   1470
      Width           =   1485
   End
   Begin VB.TextBox txt_new_slab_no 
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
      Left            =   1410
      MaxLength       =   8
      TabIndex        =   41
      Top             =   1080
      Width           =   1875
   End
   Begin VB.TextBox txt_act_stlgrd 
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
      Left            =   1400
      MaxLength       =   11
      TabIndex        =   4
      Top             =   1470
      Width           =   1245
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3900
      Width           =   14610
      _Version        =   393216
      _ExtentX        =   25770
      _ExtentY        =   8493
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   13
      MaxRows         =   20
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "CGA2090C.frx":0000
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   12720
      Top             =   1470
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "入库块数"
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
   Begin InDate.ULabel ULabel7 
      Height          =   1485
      Left            =   150
      Top             =   1950
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2619
      Caption         =   "转炉成分(%)"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   6615
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Si"
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   5325
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "S"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   4035
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "P"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   2730
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Mn"
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
      Height          =   315
      Left            =   1395
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "C"
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
   Begin InDate.ULabel ULabel13 
      Height          =   315
      Left            =   7920
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Ceq"
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
   Begin InDate.ULabel ULabel14 
      Height          =   315
      Left            =   9210
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Nb"
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
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   10455
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Cu"
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   11760
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "V"
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
      Left            =   6615
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "W"
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
      Left            =   5325
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Ti"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   4035
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Mo"
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   2730
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Alt"
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   1395
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Ni"
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
   Begin InDate.ULabel ULabel19 
      Height          =   315
      Left            =   7920
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "B"
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
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   9210
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Re"
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
   Begin InDate.ULabel ULabel21 
      Height          =   315
      Left            =   10455
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Pb"
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
   Begin InDate.ULabel ULabel22 
      Height          =   315
      Left            =   11760
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Ca"
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
   Begin InDate.ULabel ULabel24 
      Height          =   315
      Left            =   6615
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Mg"
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
   Begin InDate.ULabel ULabel25 
      Height          =   315
      Left            =   5325
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Zr"
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
   Begin InDate.ULabel ULabel26 
      Height          =   315
      Left            =   4035
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Als"
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
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   2730
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "O"
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
   Begin InDate.ULabel ULabel28 
      Height          =   315
      Left            =   1395
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "H"
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
   Begin InDate.ULabel ULabel29 
      Height          =   315
      Left            =   7920
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Sn"
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
   Begin InDate.ULabel ULabel30 
      Height          =   315
      Left            =   9210
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "As"
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
   Begin InDate.ULabel ULabel31 
      Height          =   315
      Left            =   10455
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Co"
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
   Begin InDate.ULabel ULabel32 
      Height          =   315
      Left            =   11760
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "TE"
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
      Left            =   150
      Top             =   1470
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "钢种"
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
   Begin InDate.ULabel ULabel34 
      Height          =   315
      Left            =   4305
      Top             =   1470
      Width           =   660
      _ExtentX        =   1164
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
   Begin InDate.ULabel ULabel23 
      Height          =   315
      Left            =   9360
      Top             =   1470
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   556
      Caption         =   "重量"
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
   Begin InDate.ULabel ULabel33 
      Height          =   315
      Left            =   11040
      Top             =   1470
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   556
      Caption         =   "车号"
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
   Begin InDate.ULabel ULabel35 
      Height          =   315
      Left            =   150
      Top             =   690
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "转炉号(老)"
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
   Begin InDate.ULabel ULabel36 
      Height          =   315
      Left            =   13080
      Top             =   1950
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Cr"
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
   Begin InDate.ULabel ULabel37 
      Height          =   315
      Left            =   13080
      Top             =   2340
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "N"
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
   Begin InDate.ULabel ULabel38 
      Height          =   315
      Left            =   13080
      Top             =   2730
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "BI"
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
   Begin CSTextLibCtl.sidbEdit txt_wid 
      Height          =   315
      Left            =   6600
      TabIndex        =   6
      Top             =   1470
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      DataProperty    =   1
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
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_len 
      Height          =   315
      Left            =   8250
      TabIndex        =   7
      Top             =   1470
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      DataProperty    =   1
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
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_wgt 
      Height          =   315
      Left            =   10020
      TabIndex        =   8
      Top             =   1470
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_c 
      Height          =   315
      Left            =   1740
      TabIndex        =   11
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ni 
      Height          =   315
      Left            =   1740
      TabIndex        =   21
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_h 
      Height          =   315
      Left            =   1740
      TabIndex        =   31
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   7
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_mn 
      Height          =   315
      Left            =   3075
      TabIndex        =   12
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_alt 
      Height          =   315
      Left            =   3075
      TabIndex        =   22
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_o 
      Height          =   315
      Left            =   3075
      TabIndex        =   32
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   7
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_p 
      Height          =   315
      Left            =   4380
      TabIndex        =   13
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_mo 
      Height          =   315
      Left            =   4380
      TabIndex        =   23
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_als 
      Height          =   315
      Left            =   4380
      TabIndex        =   33
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_s 
      Height          =   315
      Left            =   5670
      TabIndex        =   14
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ti 
      Height          =   315
      Left            =   5670
      TabIndex        =   24
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_zr 
      Height          =   315
      Left            =   5670
      TabIndex        =   34
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_si 
      Height          =   315
      Left            =   6960
      TabIndex        =   15
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_w 
      Height          =   315
      Left            =   6960
      TabIndex        =   25
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_mg 
      Height          =   315
      Left            =   6960
      TabIndex        =   35
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ceq 
      Height          =   315
      Left            =   8250
      TabIndex        =   16
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   3
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_b 
      Height          =   315
      Left            =   8250
      TabIndex        =   26
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_sn 
      Height          =   315
      Left            =   8250
      TabIndex        =   36
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_nb 
      Height          =   315
      Left            =   9540
      TabIndex        =   17
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_re 
      Height          =   315
      Left            =   9540
      TabIndex        =   27
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_as 
      Height          =   315
      Left            =   9540
      TabIndex        =   37
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_cu 
      Height          =   315
      Left            =   10800
      TabIndex        =   18
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_pb 
      Height          =   315
      Left            =   10800
      TabIndex        =   28
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_co 
      Height          =   315
      Left            =   10800
      TabIndex        =   38
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_v 
      Height          =   315
      Left            =   12090
      TabIndex        =   19
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   4
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ca 
      Height          =   315
      Left            =   12090
      TabIndex        =   29
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_te 
      Height          =   315
      Left            =   12090
      TabIndex        =   39
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_cr 
      Height          =   315
      Left            =   13410
      TabIndex        =   20
      Top             =   1950
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
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
      FmtThousands    =   0
      FmtControl      =   1
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_n 
      Height          =   315
      Left            =   13410
      TabIndex        =   30
      Top             =   2340
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   7
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_bi 
      Height          =   315
      Left            =   13410
      TabIndex        =   40
      Top             =   2730
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_slabcnt 
      Height          =   315
      Left            =   13650
      TabIndex        =   10
      Top             =   1470
      Width           =   645
      _Version        =   262145
      _ExtentX        =   1138
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
      DataProperty    =   1
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
      NumIntDigits    =   1
      MaxValue        =   20
      MinValue        =   1
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_thk 
      Height          =   315
      Left            =   4965
      TabIndex        =   5
      Top             =   1470
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
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
      DataProperty    =   1
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
      NumIntDigits    =   3
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   150
      Top             =   1080
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "转炉号(新)"
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
   Begin InDate.ULabel ULabel39 
      Height          =   315
      Left            =   5940
      Top             =   1470
      Width           =   660
      _ExtentX        =   1164
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
   Begin InDate.ULabel ULabel40 
      Height          =   315
      Left            =   7590
      Top             =   1470
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   556
      Caption         =   "长度"
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
   Begin InDate.ULabel ULabel42 
      Height          =   315
      Left            =   5325
      Top             =   3120
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Ta"
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
   Begin InDate.ULabel ULabel43 
      Height          =   315
      Left            =   4035
      Top             =   3120
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Se"
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
   Begin InDate.ULabel ULabel44 
      Height          =   315
      Left            =   2730
      Top             =   3120
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Zn"
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
   Begin InDate.ULabel ULabel45 
      Height          =   315
      Left            =   1395
      Top             =   3120
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Sb"
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
   Begin CSTextLibCtl.sidbEdit txt_sb 
      Height          =   315
      Left            =   1740
      TabIndex        =   43
      Top             =   3120
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_zn 
      Height          =   315
      Left            =   3075
      TabIndex        =   44
      Top             =   3120
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_se 
      Height          =   315
      Left            =   4380
      TabIndex        =   45
      Top             =   3120
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_ta 
      Height          =   315
      Left            =   5670
      TabIndex        =   46
      Top             =   3120
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0000"
      Text            =   " 0.0000"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   4
      NumIntDigits    =   5
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel41 
      Height          =   315
      Left            =   6615
      Top             =   3120
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "Pcm"
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
   Begin CSTextLibCtl.sidbEdit txt_pcm 
      Height          =   315
      Left            =   6960
      TabIndex        =   48
      Top             =   3120
      Width           =   855
      _Version        =   262145
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.000"
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
      DataProperty    =   1
      FocusSelect     =   -1  'True
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   3
      MaxValue        =   20
      MinValue        =   10
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel46 
      Height          =   315
      Left            =   1410
      Top             =   690
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   556
      Caption         =   "月"
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
   Begin InDate.ULabel ULabel47 
      Height          =   315
      Left            =   2160
      Top             =   690
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   556
      Caption         =   "转炉号"
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
   Begin InDate.ULabel ULabel48 
      Height          =   315
      Left            =   3330
      Top             =   690
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "生产顺序号"
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
   Begin InDate.ULabel ULabel49 
      Height          =   315
      Left            =   3330
      Top             =   1080
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "生产年度"
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
   Begin VB.TextBox txt_OldSlabNo2 
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
      Left            =   1410
      MaxLength       =   20
      TabIndex        =   55
      Top             =   690
      Visible         =   0   'False
      Width           =   3045
   End
   Begin CSTextLibCtl.sitxEdit txt_inyard_time 
      Height          =   315
      Left            =   12135
      TabIndex        =   59
      Tag             =   "缺号时"
      Top             =   3120
      Width           =   2130
      _Version        =   262145
      _ExtentX        =   3757
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "____-__-__ __:__:__"
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
      Mask            =   "____-__-__ __:__:__"
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   14
      ValidateMask    =   0   'False
   End
   Begin InDate.ULabel ULabel50 
      Height          =   315
      Left            =   10455
      Top             =   3120
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Caption         =   "来料入库时间"
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
End
Attribute VB_Name = "CGA2090C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      板坯外卖信息录入
'-- Program ID        CGA2090C
'-- Designer          SHIN.C.S
'-- Coder             SHIN.C.S
'-- Date              2007.07.26
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

Dim bySlabNo As String
Dim firstFL As String
Dim SALENO As String


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(txt_new_slab_no, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_InPltCo, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_InPltCoDesc, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_OldSlabNo2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_THK, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_WID, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_LEN, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_WGT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_car_no, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_act_stlgrd, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_act_stlgrd_dec, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_slabcnt, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_mon, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_heat, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_SEQ, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

            '成分
             Call Gp_Ms_Collection(txt_c, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_mn, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_P, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_s, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_si, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_ceq, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_nb, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_cu, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_v, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_cr, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ni, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_alt, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_mo, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ti, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_w, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_b, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_re, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_pb, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ca, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_n, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_H, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_o, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_als, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_zr, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_mg, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_sn, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_as, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_co, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_te, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_bi, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_SB, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_zn, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_se, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_ta, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_pcm, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_inyard_time, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:="CGA2090C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="CGA2090C.P_REFER", Key:="P-R"
    
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

    
  
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
'    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGA2090C.P_MODIFY1", Key:="P-M"
    sc1.Add Item:="CGA2090C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 7, True)
    Call Gp_Sp_ColHidden(ss1, 10, True)
    Call Gp_Sp_ColHidden(ss1, 13, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cmd_get_info_Click()
Dim Mon As String
Dim HeatNo As String
Dim sQuery As String


    If Len(Trim(txt_new_slab_no)) = 8 Then
        
        sQuery = "          SELECT MAX(SLAB_NO) "
        sQuery = sQuery & "   FROM NISCO.FP_SLAB "
        sQuery = sQuery & "  WHERE SLAB_NO LIKE '" & Trim(txt_new_slab_no) & "%' AND MOTHER_SLAB IS NULL"
    
        bySlabNo = Gf_CodeFind(M_CN1, sQuery)
        If Len(bySlabNo) = 0 Then
           bySlabNo = txt_new_slab_no & "00"
        End If
        
        Call Get_OldSms_Info
        
        comm_slab.Enabled = True
       
    End If
End Sub

Private Sub comm_slab_Click()
Dim I, j As Integer
Dim NEWSLABNO As String
Dim tmThk, tmWid, tmLen As String
Dim tmWgt As Double

    If Len(txt_act_stlgrd) = 0 Then
       Call Gp_MsgBoxDisplay("请输入钢种..!", "G", "")
       Exit Sub
    End If
    
    If TXT_THK = 0 Or TXT_WID = 0 Or TXT_LEN = 0 Then
       Call Gp_MsgBoxDisplay("请输入板坯尺寸..!", "G", "")
       Exit Sub
    End If
    
    If TXT_WGT = 0 Then
       Call Gp_MsgBoxDisplay("请输入板坯重量..!", "G", "")
       Exit Sub
    End If
    
    If txt_slabcnt = 0 Then
       Call Gp_MsgBoxDisplay("请输入板坯切割块数 ..!", "G", "")
       Exit Sub
    End If
 
    If opt_BuyCo(1).Value = True Or opt_BuyCo(2).Value = True Then
       If Trim(txt_OldSlabNo2) = "" Then
            Call Gp_MsgBoxDisplay("请输入原来板坯号码..!", "G", "")
            Exit Sub
        End If
        
        If Len(Trim(txt_new_slab_no)) <> 8 Then
            Call Gp_MsgBoxDisplay("请输入新板坯号码..!", "G", "")
            Exit Sub
        End If
        
        If Mid(Trim(txt_new_slab_no), 3, 1) <> "A" And Mid(Trim(txt_new_slab_no), 3, 1) <> "B" Then
            Call Gp_MsgBoxDisplay("请输入新板坯号码..!", "G", "")
            Exit Sub
        End If
        
        If Len(Trim(txt_InPltCo)) <> 6 Then
            Call Gp_MsgBoxDisplay("请输入外购公司代吗..!", "G", "")
            Exit Sub
        End If
        
        
    End If
    
    ss1.MaxRows = txt_slabcnt
    
    If opt_BuyCo(1).Value = True Or opt_BuyCo(2).Value = True Then
       bySlabNo = "00"
    End If
    
    For I = 1 To txt_slabcnt
        ss1.Row = I
        ss1.Col = 1
        If opt_BuyCo(1).Value = True Or opt_BuyCo(2).Value = True Then
           If I < 10 Then
              bySlabNo = "0" + CStr((bySlabNo + 1))
           Else
              bySlabNo = bySlabNo + 1
           End If
           
            NEWSLABNO = SALENO + bySlabNo
            
        Else
            NEWSLABNO = Mid(bySlabNo, 1, 4) & CStr(Mid(bySlabNo, 5, 6) + I)
            
            If Len(Mid(NEWSLABNO, 5, 6)) = 5 Then
               NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "0" & Mid(NEWSLABNO, 5, 5)
            ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 4 Then
               NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "00" & Mid(NEWSLABNO, 5, 5)
            ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 3 Then
               NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "000" & Mid(NEWSLABNO, 5, 5)
            End If
        End If
        '''''MODIFIED BY GUOLI AT 20080301
        If txt_new_slab_no.Enabled = True Then
        ss1.Text = NEWSLABNO
        End If

        ss1.Col = 2
        ss1.Text = txt_old_slabno

        ss1.Col = 3
        ss1.Text = TXT_THK
        tmThk = TXT_THK
        
        ss1.Col = 4
        ss1.Text = TXT_WID
        tmWid = TXT_WID

        ss1.Col = 5
        ss1.Text = TXT_LEN
        tmLen = TXT_LEN
        
        ss1.Col = 6
        ss1.Text = TXT_WGT / txt_slabcnt
        tmWgt = tmWgt + CDbl(ss1.Text)

        ss1.Col = 7
        ss1.Text = ((CDbl(tmThk) * CDbl(tmWid) * CDbl(tmLen)) * 7.85) / 1000000000

        ss1.Col = 8
        ss1.Text = ""
        
        ss1.Col = 9
        ss1.Text = txt_act_stlgrd

        ss1.Col = 10
        ss1.Text = ""

        ss1.Col = 11
        ss1.Text = sUserID
        
        ss1.Col = 12
        ss1.Text = txt_car_no
        
        ss1.Col = 13
        ss1.Text = txt_InPltCo
        
        ss1.Row = I
        ss1.Col = 0
        If txt_new_slab_no.Enabled = True Then
            ss1.Text = "Input"
        Else
            ss1.Text = "Update"
        End If
    Next I
    
    Call WGT_CAL

End Sub

Private Sub NewKey_Creation()
Dim Mon As String
Dim HeatNo As String
Dim sQuery As String
Dim HeatSeq As String


        If opt_BuyCo(1).Value = True Or opt_BuyCo(2).Value = True Then
            If Len(txt_cen) = 2 Then
                If txt_cen <> "07" And txt_cen <> "08" And txt_cen <> "09" And txt_cen <> "10" And txt_cen <> "11" And txt_cen <> "12" Then
                    MsgBox "请确认年...!"
                    txt_cen.SetFocus
                    Exit Sub
                End If
            End If
            
            sQuery = "          SELECT MAX(SUBSTR(SLAB_NO,1,8)) "
            sQuery = sQuery & "   FROM NISCO.FP_SLAB "
            
            If opt_BuyCo(1).Value = True Then
                sQuery = sQuery & "  WHERE SUBSTR(SLAB_NO,3,1)= 'A'"
            Else
                sQuery = sQuery & "  WHERE SUBSTR(SLAB_NO,3,1)= 'B'"
            End If
            
            SALENO = Gf_CodeFind(M_CN1, sQuery)
            
            If Len(SALENO) = 0 Then
                If opt_BuyCo(1).Value = True Then
                    SALENO = txt_cen & "A" & "00001"
                Else
                    SALENO = txt_cen & "B" & "00001"
                End If
            Else
                HeatSeq = Mid(SALENO, 4, 5) + 1
                
                If Len(HeatSeq) = 1 Then
                   HeatSeq = "0000" & HeatSeq
                ElseIf Len(HeatSeq) = 2 Then
                   HeatSeq = "000" & HeatSeq
                ElseIf Len(HeatSeq) = 3 Then
                   HeatSeq = "00" & HeatSeq
                ElseIf Len(HeatSeq) = 4 Then
                   HeatSeq = "0" & HeatSeq
                End If
                
                If opt_BuyCo(1).Value = True Then
                    SALENO = txt_cen & "A" & HeatSeq
                Else
                    SALENO = txt_cen & "B" & HeatSeq
                End If
            End If
        
            txt_new_slab_no = SALENO
            Exit Sub
        End If
        
        If Len(txt_cen) = 2 Then
            If txt_cen <> "07" And txt_cen <> "08" And txt_cen <> "09" And txt_cen <> "10" And txt_cen <> "11" And txt_cen <> "12" Then
                MsgBox "请确认年...!"
                txt_cen.SetFocus
                Exit Sub
            End If
        End If
            
        If Len(txt_mon) = 2 Then
            If txt_mon = "01" Or txt_mon = "02" Or txt_mon = "03" Or txt_mon = "04" Or txt_mon = "05" Or txt_mon = "06" _
                              Or txt_mon = "07" Or txt_mon = "08" Or txt_mon = "09" Or txt_mon = "10" Or txt_mon = "11" Or txt_mon = "12" Then
            Else
                MsgBox "请确认月...!"
                txt_mon.SetFocus
                Exit Sub
            End If
        End If
        
        If Len(txt_heat) = 1 Then
            If txt_heat = "1" Or txt_heat = "2" Or txt_heat = "3" Then
            Else
                MsgBox "请确认转炉号(不是 1,2,3) ...!"
                txt_heat.SetFocus
              
                Exit Sub
            End If
        End If
        
        'Month check
        If txt_mon = "10" Then
           Mon = "A"
        ElseIf txt_mon = "11" Then
           Mon = "B"
        ElseIf txt_mon = "12" Then
           Mon = "C"
        Else
           Mon = Mid(txt_mon, 2, 1)
        End If
        
        'Heat no check
        If txt_heat = "1" Then
           HeatNo = "4"
        ElseIf txt_heat = "2" Then
           HeatNo = "5"
        ElseIf txt_heat = "3" Then
           HeatNo = "6"
        End If
        
        txt_old_slabno = txt_mon & "-" & txt_heat & "-" & TXT_SEQ
        txt_new_slab_no = txt_cen & HeatNo & Mon & TXT_SEQ
        
        If Len(Trim(txt_new_slab_no)) = 8 Then
            Call cmd_get_info_Click
        End If

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
    Call Form_Activate
    
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
    opt_BuyCo(0).Value = True
    lab_InPltCo.Visible = False
    txt_InPltCo.Visible = False
    txt_InPltCoDesc.Visible = False
    
    txt_cen.Text = Gf_CodeFind(M_CN1, "SELECT SUBSTR(TO_CHAR(SYSDATE,'YYYY'),3,2) FROM DUAL")
       
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
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

Public Sub Spread_Can()
  
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    Call WGT_CAL
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        Call MenuTool_ReSet
    End If
    
    txt_cen.Text = Gf_CodeFind(M_CN1, "SELECT SUBSTR(TO_CHAR(SYSDATE,'YYYY'),3,2) FROM DUAL")
    txt_act_stlgrd_dec = ""
    comm_slab.Enabled = True
    
End Sub

Public Sub Form_Ref()
Dim ForCnt As Integer
Dim ChkVal As Integer

    firstFL = ""
    

    If Len(Trim(txt_new_slab_no)) < 8 Then
       MsgBox "确认Heat_no 不够8位"
       Exit Sub
    End If
  
'    If Mid(txt_new_slab_no, 3, 1) = "A" Then
'       opt_BuyCo(1).Value = True
'    ElseIf Mid(txt_new_slab_no, 3, 1) = "G" Then
'        opt_BuyCo(2).Value = True
'    End If
  
    Call Gf_Sp_Cls(Proc_Sc("SC"))
    
    If Gf_Ms_Refer(M_CN1, Mc1, , , False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        comm_slab.Enabled = False
    End If
    Call CHEMISTRY_DISP
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Nothing, Mc1("mControl")) Then
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        firstFL = "Y"
        comm_slab.Enabled = False
    End If
    
        
    ChkVal = 0
    For ForCnt = 1 To ss1.MaxRows
        ss1.Row = ForCnt
        ss1.Col = 8
        If ss1.Text = "使用完" Then
           ChkVal = ChkVal + 1
        End If
    Next ForCnt
    
    If ChkVal > 0 Then
        Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows, True)
    End If
    
    If ss1.MaxRows > 0 Then
       txt_slabcnt = ss1.MaxRows
    End If

End Sub
Public Sub WGT_CAL()

    Dim tmThk As Double
    Dim tmWid As Double
    Dim tmLen As Double
    Dim tempWgt As Double
    Dim tot_cal_total As Double
    Dim cal_wgt As Double
    Dim sub_wgt As Double
    Dim tmp_rat As Double
    Dim tot_rate As Double
    
    Dim I As Integer

    cal_wgt = 0
    sub_wgt = 0

    If ss1.MaxRows < 1 Then Exit Sub
    
    For I = 1 To ss1.MaxRows
         ss1.Row = I
         ss1.Col = 3
             tmThk = ss1.Value
         ss1.Col = 4
             tmWid = ss1.Value
         ss1.Col = 5
             tmLen = ss1.Value
         ss1.Col = 7
         ss1.Text = Round((tmThk * tmWid * tmLen * 7.85) / 1000000000, 3)
         cal_wgt = cal_wgt + Round((tmThk * tmWid * tmLen * 7.85) / 1000000000, 3)
    Next I
    
    sub_wgt = CDbl(TXT_WGT)
    tot_rate = 0
    
    For I = 1 To ss1.MaxRows
         ss1.Row = I
         ss1.Col = 0
         If UCase(ss1.Text) <> "DELETE" Then
             ss1.Row = I
             ss1.Col = 7
                 tmp_rat = ss1.Text / cal_wgt
                 tot_rate = tot_rate + tmp_rat
             ss1.Col = 6
             ss1.Text = Round(TXT_WGT * tmp_rat, 3)
         End If
    Next I
    
    For I = 1 To ss1.MaxRows - 1
         ss1.Row = I
         ss1.Col = 0
         If UCase(ss1.Text) <> "DELETE" Then
             ss1.Row = I
             ss1.Col = 7
                 tmp_rat = ss1.Text / cal_wgt
             ss1.Col = 6
             ss1.Text = Round(TXT_WGT * tmp_rat / tot_rate, 3)
             sub_wgt = sub_wgt - Round(TXT_WGT * tmp_rat / tot_rate, 3)
         End If
    Next I
    
    ss1.Row = ss1.MaxRows
    ss1.Col = 6
    ss1.Text = sub_wgt
    
    For I = 1 To ss1.MaxRows
         ss1.Row = I
         ss1.Col = 0
         If UCase(ss1.Text) = "" Then
            ss1.Text = "Update"
         End If
    Next I
       
       
End Sub

Public Sub Form_Pro()

    Dim ForCnt As Integer
    Dim ChkVal As Integer
    Dim I As Integer
    Dim NEWSLABNO As String
    
    If opt_BuyCo(1).Value = True Or opt_BuyCo(2).Value = True Then
       Call NewKey_Creation
       bySlabNo = "00"
    End If
    
    For I = 1 To txt_slabcnt
    
        ss1.Row = I
        ss1.Col = 1
        If opt_BuyCo(1).Value = True Or opt_BuyCo(2).Value = True Then
           If I < 10 Then
              bySlabNo = "0" + CStr((bySlabNo + 1))
           Else
              bySlabNo = bySlabNo + 1
           End If
           
           NEWSLABNO = SALENO + bySlabNo
           
'        Else
'            NEWSLABNO = Mid(bySlabNo, 1, 4) & CStr(Mid(bySlabNo, 5, 6) + I)
'
'            If Len(Mid(NEWSLABNO, 5, 6)) = 5 Then
'               NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "0" & Mid(NEWSLABNO, 5, 5)
'            ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 4 Then
'               NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "00" & Mid(NEWSLABNO, 5, 5)
'            ElseIf Len(Mid(NEWSLABNO, 5, 6)) = 3 Then
'               NEWSLABNO = Mid(NEWSLABNO, 1, 4) & "000" & Mid(NEWSLABNO, 5, 5)
''            End If
        End If
'            If txt_new_slab_no.Enabled = True Then
'                ss1.Text = NEWSLABNO
'            End If
        
        
    Next I
    
    
    If ss1.Row = 0 Or ss1.Row = -999 Then
        MsgBox "请确认子板坯号....!"
        Exit Sub
    End If
    If Len(txt_car_no) = 0 Then
       Call Gp_MsgBoxDisplay("请输入车辆号..!", "G", "")
       txt_car_no.SetFocus
       Exit Sub
    End If
    
    If Len(Trim(txt_act_stlgrd.Text)) < 10 Then
        MsgBox "请确认钢种代吗....!"
        Exit Sub
    End If
    
    If txt_c = 0 Then
        MsgBox "请确认成分(C)....!"
        Exit Sub
    End If
    
    If txt_mn = 0 Then
        MsgBox "请确认成分(MN)....!"
        Exit Sub
    End If
    
    If TXT_P = 0 Then
        MsgBox "请确认成分(P)....!"
        Exit Sub
    End If
    
    If txt_s = 0 Then
        MsgBox "请确认成分(S)....!"
        Exit Sub
    End If
    
    If txt_si = 0 Then
        MsgBox "请确认成分(SI)....!"
        Exit Sub
    End If
    
'    If txt_ceq = 0 Then
'        MsgBox "请确认成分(CEQ)....!"
'        Exit Sub
'    End If
'
'    If txt_nb = 0 Then
'        MsgBox "请确认成分(NB)....!"
'        Exit Sub
'    End If
'
'    If txt_cu = 0 Then
'        MsgBox "请确认成分(CU)....!"
'        Exit Sub
'    End If
'
'    If txt_v = 0 Then
'        MsgBox "请确认成分(V)....!"
'        Exit Sub
'    End If
'
'    If txt_cr = 0 Then
'        MsgBox "请确认成分(CR)....!"
'        Exit Sub
'    End If
'
'    If txt_ni = 0 Then
'        MsgBox "请确认成分(CR)....!"
'        Exit Sub
'    End If
    
    Call Gf_Ms_Process(M_CN1, Mc1, sAuthority)
    'Call CHEMISTRY_DISP
    Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1)
    
    ChkVal = 0
    For ForCnt = 1 To ss1.MaxRows
        ss1.Row = ForCnt
        ss1.Col = 8
        If ss1.Text = "使用完" Then
           ChkVal = ChkVal + 1
        End If
    Next ForCnt
    
    If ChkVal > 0 Then
        Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows, True)
    End If
    
    Call MenuTool_ReSet
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    
    With ss1
        .Row = .ActiveRow
        .Col = 8
        .Text = sUserID
    End With

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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

Public Sub Spread_Del()
    
    If ss1.SelBlockRow2 = ss1.MaxRows Then
        Call Gp_Sp_Del(Proc_Sc("SC"))
        
        Call WGT_CAL
    End If
    
End Sub

Private Sub opt_BuyCo_Click(Index As Integer, Value As Integer)

    Dim memory_heat_no As String
    
    memory_heat_no = txt_new_slab_no
    Call Form_Cls
    'txt_new_slab_no = memory_heat_no
    
    If Index = 0 Then
        lab_InPltCo.Visible = False
        txt_InPltCo.Visible = False
        txt_InPltCoDesc.Visible = False
        txt_mon.Visible = True
        txt_heat.Visible = True
        TXT_SEQ.Visible = True
        ULabel46.Visible = True
        ULabel47.Visible = True
        ULabel48.Visible = True
        txt_OldSlabNo2.Visible = False
        SSFrame1.Width = 3525
    Else
        lab_InPltCo.Visible = True
        txt_InPltCo.Visible = True
        txt_InPltCoDesc.Visible = True
        txt_mon.Visible = False
        txt_heat.Visible = False
        TXT_SEQ.Visible = False
        ULabel46.Visible = False
        ULabel47.Visible = False
        ULabel48.Visible = False
        txt_OldSlabNo2.Visible = True
        SSFrame1.Width = 7725
    End If

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)


   If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
         
    Call WGT_CAL
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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

Private Sub txt_act_stlgrd_Change()

    If Len(Trim(txt_act_stlgrd.Text)) < 10 Then txt_act_stlgrd_dec.Text = ""
    
End Sub

Private Sub txt_act_stlgrd_DblClick()

    Call txt_act_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_act_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        'txt_act_stlgrd.Text = ""
        DD.rControl.Add Item:=txt_act_stlgrd
        DD.rControl.Add Item:=txt_act_stlgrd_dec

        Call Gf_Stlgrd_DD(M_CN1, vbKeyF4)
        
    Else
    
        If Len(Trim(txt_act_stlgrd.Text)) >= 10 Then
            txt_act_stlgrd_dec.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_act_stlgrd.Text))
        Else
            txt_act_stlgrd_dec.Text = ""
        End If

    End If
    
    If txt_act_stlgrd_dec.Text <> "" Then
        
        If opt_BuyCo(2).Value = True Then
        
            txt_c.Value = Gf_FloatFind(M_CN1, "SELECT CHEM_COMP_TGT FROM QP_NISCO_CHEM WHERE STLGRD = '" & txt_act_stlgrd.Text & "' AND CHEM_COMP_CD = 'C' ")
            txt_mn.Value = Gf_FloatFind(M_CN1, "SELECT CHEM_COMP_TGT FROM QP_NISCO_CHEM WHERE STLGRD = '" & txt_act_stlgrd.Text & "' AND CHEM_COMP_CD = 'Mn' ")
            TXT_P.Value = Gf_FloatFind(M_CN1, "SELECT CHEM_COMP_TGT FROM QP_NISCO_CHEM WHERE STLGRD = '" & txt_act_stlgrd.Text & "' AND CHEM_COMP_CD = 'P' ")
            txt_s.Value = Gf_FloatFind(M_CN1, "SELECT CHEM_COMP_TGT FROM QP_NISCO_CHEM WHERE STLGRD = '" & txt_act_stlgrd.Text & "' AND CHEM_COMP_CD = 'S' ")
            txt_si.Value = Gf_FloatFind(M_CN1, "SELECT CHEM_COMP_TGT FROM QP_NISCO_CHEM WHERE STLGRD = '" & txt_act_stlgrd.Text & "' AND CHEM_COMP_CD = 'Si' ")
            
        End If
        
    End If
    
End Sub

Private Sub Get_OldSms_Info()

    Dim sQuery As String
    Dim iRowCount As Long
    Dim ArrayRecords As Variant
    Dim AdoRs As ADODB.Recordset
    Dim tmHeatNo As String

    'Db Connection Check
    If M_CN1.State = 0 Then
        If GF_DbConnect = False Then Exit Sub
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    sQuery = "          SELECT HEAT_NO "
    sQuery = sQuery & "   FROM NISCO.FP_CHARGE "
    sQuery = sQuery & "  WHERE HEAT_NO LIKE '" & Trim(UCase(txt_new_slab_no)) & "%'"
    
    tmHeatNo = Gf_FloatFind(M_CN1, sQuery)
    If Trim(tmHeatNo) <> "0" Then
        Exit Sub
    End If
    
    sQuery = "SELECT * FROM NISCO.GP_OLDSMSCHEMIF  WHERE NEW_HEAT_NO = '" & Trim(txt_new_slab_no) & "'"
     
    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        ArrayRecords = AdoRs.GetRows
        
        For iRowCount = 0 To UBound(ArrayRecords, 2)
        
       txt_act_stlgrd_dec = ArrayRecords(5, iRowCount)
            TXT_THK.Value = ArrayRecords(6, iRowCount)
            TXT_WID.Value = ArrayRecords(7, iRowCount)
            TXT_LEN.Value = ArrayRecords(8, iRowCount)
            TXT_WGT.Value = ArrayRecords(9, iRowCount)
       txt_slabcnt.Value = ArrayRecords(10, iRowCount)
             
         txt_car_no.Text = IIf(IsNull(ArrayRecords(13, iRowCount)), 0, ArrayRecords(13, iRowCount))
             txt_c.Value = IIf(IsNull(ArrayRecords(14, iRowCount)), 0, ArrayRecords(14, iRowCount))
            txt_si.Value = IIf(IsNull(ArrayRecords(15, iRowCount)), 0, ArrayRecords(15, iRowCount))
            txt_mn.Value = IIf(IsNull(ArrayRecords(16, iRowCount)), 0, ArrayRecords(16, iRowCount))
             TXT_P.Value = IIf(IsNull(ArrayRecords(17, iRowCount)), 0, ArrayRecords(17, iRowCount))
             txt_s.Value = IIf(IsNull(ArrayRecords(18, iRowCount)), 0, ArrayRecords(18, iRowCount))
            txt_cu.Value = IIf(IsNull(ArrayRecords(19, iRowCount)), 0, ArrayRecords(19, iRowCount))
           txt_alt.Value = IIf(IsNull(ArrayRecords(20, iRowCount)), 0, ArrayRecords(20, iRowCount))
           txt_als.Value = IIf(IsNull(ArrayRecords(21, iRowCount)), 0, ArrayRecords(21, iRowCount))
             txt_b.Value = IIf(IsNull(ArrayRecords(22, iRowCount)), 0, ArrayRecords(22, iRowCount))
            txt_ni.Value = IIf(IsNull(ArrayRecords(23, iRowCount)), 0, ArrayRecords(23, iRowCount))
            txt_cr.Value = IIf(IsNull(ArrayRecords(24, iRowCount)), 0, ArrayRecords(24, iRowCount))
            txt_mo.Value = IIf(IsNull(ArrayRecords(25, iRowCount)), 0, ArrayRecords(25, iRowCount))
             txt_w.Value = IIf(IsNull(ArrayRecords(26, iRowCount)), 0, ArrayRecords(26, iRowCount))
            txt_ti.Value = IIf(IsNull(ArrayRecords(27, iRowCount)), 0, ArrayRecords(27, iRowCount))
             txt_v.Value = IIf(IsNull(ArrayRecords(28, iRowCount)), 0, ArrayRecords(28, iRowCount))
            txt_zr.Value = IIf(IsNull(ArrayRecords(29, iRowCount)), 0, ArrayRecords(29, iRowCount))
            txt_pb.Value = IIf(IsNull(ArrayRecords(30, iRowCount)), 0, ArrayRecords(30, iRowCount))
            txt_sn.Value = IIf(IsNull(ArrayRecords(31, iRowCount)), 0, ArrayRecords(31, iRowCount))
            txt_as.Value = IIf(IsNull(ArrayRecords(32, iRowCount)), 0, ArrayRecords(32, iRowCount))
            txt_ca.Value = IIf(IsNull(ArrayRecords(33, iRowCount)), 0, ArrayRecords(33, iRowCount))
            txt_co.Value = IIf(IsNull(ArrayRecords(34, iRowCount)), 0, ArrayRecords(34, iRowCount))
            txt_mg.Value = IIf(IsNull(ArrayRecords(35, iRowCount)), 0, ArrayRecords(35, iRowCount))
            txt_te.Value = IIf(IsNull(ArrayRecords(36, iRowCount)), 0, ArrayRecords(36, iRowCount))
            txt_bi.Value = IIf(IsNull(ArrayRecords(37, iRowCount)), 0, ArrayRecords(37, iRowCount))
            txt_SB.Value = IIf(IsNull(ArrayRecords(38, iRowCount)), 0, ArrayRecords(38, iRowCount))
            txt_zn.Value = IIf(IsNull(ArrayRecords(39, iRowCount)), 0, ArrayRecords(39, iRowCount))
            txt_nb.Value = IIf(IsNull(ArrayRecords(40, iRowCount)), 0, ArrayRecords(40, iRowCount))
           txt_ceq.Value = IIf(IsNull(ArrayRecords(41, iRowCount)), 0, ArrayRecords(41, iRowCount))
            txt_re.Value = IIf(IsNull(ArrayRecords(42, iRowCount)), 0, ArrayRecords(42, iRowCount))
            txt_ta.Value = IIf(IsNull(ArrayRecords(43, iRowCount)), 0, ArrayRecords(43, iRowCount))
             txt_n.Value = IIf(IsNull(ArrayRecords(44, iRowCount)), 0, ArrayRecords(44, iRowCount))
             TXT_H.Value = IIf(IsNull(ArrayRecords(45, iRowCount)), 0, ArrayRecords(45, iRowCount))
             txt_o.Value = IIf(IsNull(ArrayRecords(46, iRowCount)), 0, ArrayRecords(46, iRowCount))
            txt_se.Value = IIf(IsNull(ArrayRecords(47, iRowCount)), 0, ArrayRecords(47, iRowCount))
           txt_pcm.Value = IIf(IsNull(ArrayRecords(48, iRowCount)), 0, ArrayRecords(48, iRowCount))

        Next iRowCount
        
        AdoRs.Close
        Set AdoRs = Nothing
        If Trim(txt_act_stlgrd_dec) <> "" Then
            sQuery = "SELECT Gf_Stlgrd_CODE('" & Trim(txt_act_stlgrd_dec) & "') FROM DUAL"
            txt_act_stlgrd = Gf_CodeFind(M_CN1, sQuery)
        End If
        
        Call comm_slab_Click
    
    End If

End Sub

Private Sub CHEMISTRY_DISP()

    Dim sQuery As String
    Dim iRowCount As Long
    Dim ArrayRecords As Variant
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If M_CN1.State = 0 Then
        If GF_DbConnect = False Then Exit Sub
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    sQuery = "SELECT ELEMENT_CD, ELEMENT_VAL FROM NISCO.FP_CHEMISTRY WHERE HEAT_NO = '" & Trim(txt_new_slab_no) & "' ORDER BY SEQ_NO "
     

    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        ArrayRecords = AdoRs.GetRows
        
        For iRowCount = 0 To UBound(ArrayRecords, 2)
        
            Select Case ArrayRecords(0, iRowCount)

                    Case "C"
                        txt_c.Value = ArrayRecords(1, iRowCount)
                    Case "Mn"
                        txt_mn.Value = ArrayRecords(1, iRowCount)
                    Case "P"
                        TXT_P.Value = ArrayRecords(1, iRowCount)
                    Case "S"
                        txt_s.Value = ArrayRecords(1, iRowCount)
                    Case "Si"
                        txt_si.Value = ArrayRecords(1, iRowCount)
                    Case "Ceq"
                        txt_ceq.Value = ArrayRecords(1, iRowCount)
                    Case "Nb"
                        txt_nb.Value = ArrayRecords(1, iRowCount)
                    Case "Cu"
                        txt_cu.Value = ArrayRecords(1, iRowCount)
                    Case "V"
                        txt_v.Value = ArrayRecords(1, iRowCount)
                    Case "Cr"
                        txt_cr.Value = ArrayRecords(1, iRowCount)
                    Case "Ni"
                        txt_ni.Value = ArrayRecords(1, iRowCount)
                    Case "Alt"
                        txt_alt.Value = ArrayRecords(1, iRowCount)
                    Case "Mo"
                        txt_mo.Value = ArrayRecords(1, iRowCount)
                    Case "Ti"
                        txt_ti.Value = ArrayRecords(1, iRowCount)
                    Case "W"
                        txt_w.Value = ArrayRecords(1, iRowCount)
                    Case "B"
                        txt_b.Value = ArrayRecords(1, iRowCount)
                    Case "Re"
                        txt_re.Value = ArrayRecords(1, iRowCount)
                    Case "Pb"
                        txt_pb.Value = ArrayRecords(1, iRowCount)
                    Case "Ca"
                        txt_ca.Value = ArrayRecords(1, iRowCount)
                    Case "N"
                        txt_n.Value = ArrayRecords(1, iRowCount)
                    Case "H"
                        TXT_H.Value = ArrayRecords(1, iRowCount)
                    Case "O"
                        txt_o.Value = ArrayRecords(1, iRowCount)
                    Case "Als"
                        txt_als.Value = ArrayRecords(1, iRowCount)
                    Case "Zr"
                        txt_zr.Value = ArrayRecords(1, iRowCount)
                    Case "Mg"
                        txt_mg.Value = ArrayRecords(1, iRowCount)
                    Case "Sn"
                        txt_sn.Value = ArrayRecords(1, iRowCount)
                    Case "As"
                        txt_as.Value = ArrayRecords(1, iRowCount)
                    Case "Co"
                        txt_co.Value = ArrayRecords(1, iRowCount)
                    Case "Te"
                        txt_te.Value = ArrayRecords(1, iRowCount)
                    Case "Bi"
                        txt_bi.Value = ArrayRecords(1, iRowCount)
                    Case "Sb"
                        txt_SB.Value = ArrayRecords(1, iRowCount)
                    Case "Zn"
                        txt_zn.Value = ArrayRecords(1, iRowCount)
                    Case "Se"
                        txt_se.Value = ArrayRecords(1, iRowCount)
                    Case "Ta"
                        txt_ta.Value = ArrayRecords(1, iRowCount)
                    Case "Pcm"
                        txt_pcm.Value = ArrayRecords(1, iRowCount)

             End Select
        Next iRowCount
        
    End If
    
    sQuery = "          SELECT STEEL_NET_WGT "
    sQuery = sQuery & "   FROM NISCO.FP_CHARGE "
    sQuery = sQuery & "  WHERE HEAT_NO LIKE '" & Trim(UCase(txt_new_slab_no)) & "%'"

    TXT_WGT.Value = 0
    TXT_WGT.Value = Gf_FloatFind(M_CN1, sQuery)
    txt_slabcnt = ss1.MaxRows
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    If txt_mon <> "" And txt_heat <> "" Then Call cmd_get_info_Click
End Sub

Private Sub txt_car_no_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Integer
    For I = 1 To txt_slabcnt
        ss1.Row = I
        ss1.Col = 12
        ss1.Text = txt_car_no
    Next I
End Sub

Private Sub txt_cen_Change()
    If Len(txt_cen) = 2 And (opt_BuyCo(1).Value = True Or opt_BuyCo(2).Value = True) Then
        Call NewKey_Creation
        'txt_mon.SetFocus
    End If
End Sub

Private Sub txt_heat_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(txt_heat) = 1 Then
        
        TXT_SEQ.SetFocus
        Call NewKey_Creation
    End If
End Sub

Private Sub txt_InPltCo_Change()
    If Len(Trim(txt_InPltCo.Text)) = 0 Then txt_InPltCoDesc.Text = ""
End Sub

Private Sub txt_InPltCo_DblClick()
    Call txt_InPltCo_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_InPltCo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_InPltCo
        DD.rControl.Add Item:=txt_InPltCoDesc

        DD.nameType = "1"

        If opt_BuyCo(1).Value = True Then
            Call Gf_Customer_DD2(M_CN1, KeyCode, "P")
        ElseIf opt_BuyCo(2).Value = True Then
            Call Gf_Customer_DD2(M_CN1, KeyCode, "B")
        End If
        
        If Trim(txt_OldSlabNo2) = "" And Trim(txt_new_slab_no) = "" Then
            Call NewKey_Creation
        End If
        
    End If

End Sub

Private Sub txt_mon_Change()
    If Len(txt_mon) = 2 Then
        Call NewKey_Creation
        txt_heat.SetFocus
    End If
    
End Sub

Private Sub txt_OldSlabNo2_Change()
     txt_old_slabno = txt_OldSlabNo2
End Sub

Private Sub txt_seq_Change()
    If Len(TXT_SEQ) = 4 Then
        Call NewKey_Creation
        txt_act_stlgrd.SetFocus
    End If
End Sub

Private Sub txt_seq_LostFocus()
    If Len(txt_mon) = 2 And Len(txt_heat) = 1 Then
        If Len(TXT_SEQ) < 4 Then
            MsgBox "请确认顺序号码 ...!"
            TXT_SEQ.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txt_slabcnt_Change()
If opt_BuyCo(1).Value = True Or opt_BuyCo(2).Value = True Then
    TXT_WGT = TXT_THK.Value * TXT_WID.Value * TXT_LEN.Value * txt_slabcnt * 7.85 / 1000000000
End If
End Sub

Private Sub txt_wgt_Change()
    ss1.MaxRows = 0
'    If ss1.MaxRows > 0 Then
'        Call WGT_CAL
'    End If
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub
