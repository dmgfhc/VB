VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGE2010C 
   Caption         =   "钢卷垛位变更及查询界面_AGE2010C"
   ClientHeight    =   8115
   ClientLeft      =   420
   ClientTop       =   1785
   ClientWidth     =   13950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_location3 
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
      Left            =   13695
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   91
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txt_location2 
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
      Left            =   12225
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   90
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txt_location1 
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
      Left            =   10740
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   89
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox TXT_T_YARD_ADDR 
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
      Left            =   4725
      MaxLength       =   7
      TabIndex        =   2
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox TXT_COIL_NO 
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
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   1600
   End
   Begin VB.TextBox TXT_S_YARD_ADDR 
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
      MaxLength       =   7
      TabIndex        =   1
      Top             =   120
      Width           =   990
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7620
      Left            =   45
      TabIndex        =   42
      Top             =   1665
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   13441
      _Version        =   196609
      PaneTree        =   "AGE2010C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   7560
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   7470
         _Version        =   393216
         _ExtentX        =   13176
         _ExtentY        =   13335
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   16
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGE2010C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   7560
         Left            =   7590
         TabIndex        =   4
         Top             =   30
         Width           =   7545
         _Version        =   393216
         _ExtentX        =   13309
         _ExtentY        =   13335
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   16
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGE2010C.frx":1BD1
      End
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   6330
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "钢卷号"
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
      Left            =   390
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "起始垛位"
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
      Left            =   3375
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "目的垛位"
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
   Begin Threed.SSCommand cmd_Loc_Search 
      Height          =   315
      Left            =   9435
      TabIndex        =   92
      Top             =   120
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      Caption         =   "垛位查询"
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   40
      Left            =   14490
      TabIndex        =   88
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   39
      Left            =   14130
      TabIndex        =   87
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   38
      Left            =   13800
      TabIndex        =   84
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   37
      Left            =   13455
      TabIndex        =   83
      Top             =   945
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   40
      Left            =   14130
      Shape           =   3  'Circle
      Top             =   855
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   39
      Left            =   14475
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   38
      Left            =   13800
      Shape           =   3  'Circle
      Top             =   855
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   18
      Left            =   13980
      TabIndex        =   22
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   17
      Left            =   13665
      TabIndex        =   21
      Top             =   1215
      Width           =   345
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   45
      X1              =   15000
      X2              =   15000
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   34
      X1              =   14670
      X2              =   14670
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   37
      Left            =   5850
      TabIndex        =   54
      Top             =   930
      Width           =   345
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   33
      X1              =   7380
      X2              =   7380
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   40
      Left            =   6900
      TabIndex        =   86
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   39
      Left            =   6540
      TabIndex        =   85
      Top             =   930
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   40
      Left            =   6870
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   39
      Left            =   6540
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   60
      X2              =   15165
      Y1              =   555
      Y2              =   570
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   36
      Left            =   13110
      TabIndex        =   82
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   35
      Left            =   12765
      TabIndex        =   81
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   34
      Left            =   12435
      TabIndex        =   80
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   33
      Left            =   12090
      TabIndex        =   79
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   32
      Left            =   11745
      TabIndex        =   78
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   31
      Left            =   11415
      TabIndex        =   77
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   30
      Left            =   11070
      TabIndex        =   76
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   29
      Left            =   10725
      TabIndex        =   75
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   28
      Left            =   10380
      TabIndex        =   74
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   27
      Left            =   10035
      TabIndex        =   73
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   26
      Left            =   9690
      TabIndex        =   72
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   25
      Left            =   9345
      TabIndex        =   71
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   24
      Left            =   9000
      TabIndex        =   70
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   23
      Left            =   8655
      TabIndex        =   69
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   22
      Left            =   8280
      TabIndex        =   68
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   21
      Left            =   7965
      TabIndex        =   67
      Top             =   945
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   20
      Left            =   14670
      TabIndex        =   66
      Top             =   1215
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   37
      Left            =   13455
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   19
      Left            =   14340
      TabIndex        =   65
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   38
      Left            =   6210
      TabIndex        =   64
      Top             =   930
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   38
      Left            =   6210
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   36
      Left            =   13110
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   35
      Left            =   12765
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   34
      Left            =   12420
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   32
      Left            =   11730
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   33
      Left            =   12075
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   31
      Left            =   11385
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   30
      Left            =   11055
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   29
      Left            =   10710
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   28
      Left            =   10365
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   27
      Left            =   10020
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   26
      Left            =   9675
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   25
      Left            =   9330
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   24
      Left            =   8985
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   23
      Left            =   8640
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   22
      Left            =   8295
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   21
      Left            =   7950
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   20
      Left            =   14670
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   19
      Left            =   14340
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   44
      X1              =   11580
      X2              =   11580
      Y1              =   1425
      Y2              =   1515
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   43
      X1              =   11925
      X2              =   11925
      Y1              =   1425
      Y2              =   1515
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   41
      X1              =   12270
      X2              =   12270
      Y1              =   1425
      Y2              =   1515
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   40
      X1              =   12615
      X2              =   12615
      Y1              =   1425
      Y2              =   1515
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   39
      X1              =   12960
      X2              =   12960
      Y1              =   1425
      Y2              =   1515
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   38
      X1              =   13305
      X2              =   13305
      Y1              =   1425
      Y2              =   1515
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   37
      X1              =   13650
      X2              =   13650
      Y1              =   1425
      Y2              =   1515
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   36
      X1              =   13995
      X2              =   13995
      Y1              =   1425
      Y2              =   1515
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   35
      X1              =   14340
      X2              =   14340
      Y1              =   1425
      Y2              =   1515
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   10905
      TabIndex        =   63
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   29
      Left            =   3105
      TabIndex        =   62
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   30
      Left            =   3450
      TabIndex        =   61
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   31
      Left            =   3810
      TabIndex        =   60
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   32
      Left            =   4140
      TabIndex        =   59
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   33
      Left            =   4485
      TabIndex        =   58
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   34
      Left            =   4830
      TabIndex        =   57
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   35
      Left            =   5175
      TabIndex        =   56
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   36
      Left            =   5520
      TabIndex        =   55
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   3645
      TabIndex        =   53
      Top             =   1230
      Width           =   345
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   42
      X1              =   15765
      X2              =   15765
      Y1              =   1335
      Y2              =   1425
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   420
      Index           =   98
      Left            =   15345
      Shape           =   3  'Circle
      Top             =   975
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   99
      Left            =   15390
      TabIndex        =   52
      Top             =   1065
      Width           =   375
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   10
      X1              =   15345
      X2              =   15345
      Y1              =   1335
      Y2              =   1425
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   21
      X1              =   11235
      X2              =   11235
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   10
      Left            =   3645
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   3990
      TabIndex        =   51
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   4335
      TabIndex        =   50
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   4695
      TabIndex        =   49
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   5040
      TabIndex        =   48
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   5385
      TabIndex        =   47
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   16
      Left            =   5715
      TabIndex        =   46
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   17
      Left            =   6045
      TabIndex        =   45
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   18
      Left            =   6375
      TabIndex        =   44
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   19
      Left            =   6750
      TabIndex        =   43
      Top             =   1230
      Width           =   345
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   32
      X1              =   3645
      X2              =   3645
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   31
      X1              =   3990
      X2              =   3990
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   30
      X1              =   4335
      X2              =   4335
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   29
      X1              =   4680
      X2              =   4680
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   28
      X1              =   5025
      X2              =   5025
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   27
      X1              =   5370
      X2              =   5370
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   26
      X1              =   5700
      X2              =   5700
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   25
      X1              =   6060
      X2              =   6060
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   24
      X1              =   6375
      X2              =   6375
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   23
      X1              =   6720
      X2              =   6720
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   22
      X1              =   7050
      X2              =   7050
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   20
      X1              =   10875
      X2              =   10875
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   19
      X1              =   10545
      X2              =   10545
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   18
      X1              =   10200
      X2              =   10200
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   17
      X1              =   9855
      X2              =   9855
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   16
      X1              =   9495
      X2              =   9495
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   15
      X1              =   9165
      X2              =   9165
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   14
      X1              =   8820
      X2              =   8820
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   13
      X1              =   8475
      X2              =   8475
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   12
      X1              =   8130
      X2              =   8130
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   11
      X1              =   7815
      X2              =   7815
      Y1              =   1455
      Y2              =   1545
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   9
      X1              =   3300
      X2              =   3300
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   8
      X1              =   2955
      X2              =   2955
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   7
      X1              =   2610
      X2              =   2610
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   6
      X1              =   2265
      X2              =   2265
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   5
      X1              =   1920
      X2              =   1920
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   4
      X1              =   1575
      X2              =   1575
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   3
      X1              =   1230
      X2              =   1230
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   2
      X1              =   885
      X2              =   885
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   1
      X1              =   540
      X2              =   540
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      Index           =   0
      X1              =   210
      X2              =   210
      Y1              =   1440
      Y2              =   1530
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   1
      X1              =   7740
      X2              =   15090
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      Index           =   0
      X1              =   150
      X2              =   7440
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   28
      Left            =   2760
      TabIndex        =   41
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   27
      Left            =   2430
      TabIndex        =   40
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   26
      Left            =   2085
      TabIndex        =   39
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   25
      Left            =   1740
      TabIndex        =   38
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   24
      Left            =   1395
      TabIndex        =   37
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   23
      Left            =   1050
      TabIndex        =   36
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   22
      Left            =   705
      TabIndex        =   35
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   21
      Left            =   360
      TabIndex        =   34
      Top             =   930
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   20
      Left            =   7050
      TabIndex        =   33
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   3300
      TabIndex        =   32
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   8
      Left            =   2955
      TabIndex        =   31
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   2610
      TabIndex        =   30
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   2265
      TabIndex        =   29
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   1920
      TabIndex        =   28
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   1575
      TabIndex        =   27
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1230
      TabIndex        =   26
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   885
      TabIndex        =   25
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   540
      TabIndex        =   24
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_F1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   195
      TabIndex        =   23
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   16
      Left            =   13320
      TabIndex        =   20
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   12975
      TabIndex        =   19
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   12630
      TabIndex        =   18
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   12285
      TabIndex        =   17
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   11940
      TabIndex        =   16
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   11595
      TabIndex        =   15
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   11250
      TabIndex        =   14
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   7800
      TabIndex        =   5
      Top             =   1215
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   195
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   10560
      TabIndex        =   13
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   10215
      TabIndex        =   12
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   9870
      TabIndex        =   11
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   9525
      TabIndex        =   10
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   9180
      TabIndex        =   9
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   8835
      TabIndex        =   8
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   8490
      TabIndex        =   7
      Top             =   1215
      Width           =   345
   End
   Begin VB.Label L_T1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   8145
      TabIndex        =   6
      Top             =   1215
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   18
      Left            =   13995
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   17
      Left            =   13650
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   16
      Left            =   13305
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   15
      Left            =   12960
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   14
      Left            =   12615
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   13
      Left            =   12270
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   12
      Left            =   11925
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   11
      Left            =   11580
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   10
      Left            =   11235
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   8
      Left            =   10545
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   10200
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   9855
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   9510
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   9165
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   8820
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   8475
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   8130
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   7785
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   28
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   27
      Left            =   2415
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   26
      Left            =   2070
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   25
      Left            =   1725
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   24
      Left            =   1380
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   23
      Left            =   1035
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   22
      Left            =   690
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   21
      Left            =   345
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   20
      Left            =   7050
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   9
      Left            =   3300
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   8
      Left            =   2955
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   2610
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   2265
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   1575
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   1230
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   885
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   540
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   11
      Left            =   3990
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   12
      Left            =   4335
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   13
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   14
      Left            =   5025
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   15
      Left            =   5370
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   16
      Left            =   5715
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   17
      Left            =   6045
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   18
      Left            =   6375
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   19
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   29
      Left            =   3105
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   30
      Left            =   3450
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   31
      Left            =   3795
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   32
      Left            =   4140
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   33
      Left            =   4485
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   34
      Left            =   4830
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   35
      Left            =   5175
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   36
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_F1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   37
      Left            =   5865
      Shape           =   3  'Circle
      Top             =   870
      Width           =   345
   End
   Begin VB.Shape S_T1 
      BackColor       =   &H00000000&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   9
      Left            =   10890
      Shape           =   3  'Circle
      Top             =   1155
      Width           =   345
   End
End
Attribute VB_Name = "AGE2010C"
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
'-- Program Name      钢卷入库，垛位变更及库存查询界面
'-- Program ID        AGE2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
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
Public sOth As String

Dim ScoilNo As String
Dim Click_YN As Boolean

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl2 As New Collection     'Master Primary Key Collection
Dim nControl2 As New Collection     'Master Necessary Collection
Dim mControl2 As New Collection     'Master Maxlength check Collection
Dim iControl2 As New Collection     'Master Insert Collection
Dim rControl2 As New Collection     'Master Refer Collection
Dim cControl2 As New Collection     'Master Copy Collection
Dim aControl2 As New Collection     'Master -> Spread Collection
Dim lControl2 As New Collection     'Master Lock Collection

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

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(TXT_COIL_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_S_YARD_ADDR, "p", " ", " ", " ", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
           Call Gp_Ms_Collection(TXT_COIL_NO, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(TXT_T_YARD_ADDR, "p", " ", " ", " ", "r", "a", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     
    'MASTER Collection
'    Mc1.Add Item:="AGE2010C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AGE2010C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
    
   ' Mc2.Add Item:="AGE2010C.P_MODIFY", Key:="P-M"
    Mc2.Add Item:="AGE2010C.P_REFER", Key:="P-R"
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", "i", "a", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGE2010C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:="AGE2010C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AGE2010C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGE2010C.P_SMODIFY", Key:="P-M"
    sc2.Add Item:="AGE2010C.P_SREFER", Key:="P-R"
    sc2.Add Item:="AGE2010C.P_ONEROW", Key:="P-O"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 14, True)
    Call Gp_Sp_ColHidden(ss2, 14, True)
    Call Gp_Sp_ColHidden(ss1, 15, True)
    Call Gp_Sp_ColHidden(ss2, 15, True)
    
    Click_YN = False
    
End Sub

Private Sub cmd_Loc_Search_Click()
    Dim OutParam(3, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    If Trim(TXT_COIL_NO.Text) = "" Then
        Call Gp_MsgBoxDisplay("必须输入钢卷号", "", "错误提示")
        Exit Sub
    End If
    
    On Error Resume Next

    Screen.MousePointer = vbHourglass
    
    txt_location1.Text = ""
    txt_location2.Text = ""
    txt_location3.Text = ""
        
    'Return loaction1 Parameter
    OutParam(1, 1) = "arg_loaction1"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 10

    'Return loaction2 Parameter
    OutParam(2, 1) = "arg_loaction2"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 10
    
    'Return loaction3 Parameter
    OutParam(3, 1) = "arg_loaction3"
    OutParam(3, 2) = adVarChar
    OutParam(3, 3) = adParamOutput
    OutParam(3, 4) = 10
        
    sQuery = "{call AFL2010P ('HC','" & Trim(TXT_COIL_NO.Text) & "',?,?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(3, 1), OutParam(3, 2), OutParam(3, 3), OutParam(3, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If Left(adoCmd("arg_loaction1"), 3) = "NOT" Then
        Call Gp_MsgBoxDisplay("垛位查询失败，请确认")
    Else
        txt_location1.Text = adoCmd("arg_loaction1")
        txt_location2.Text = adoCmd("arg_loaction2")
        txt_location3.Text = adoCmd("arg_loaction3")
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

    If sOth <> "" Then
       Call Form_Ref
       sOth = ""
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Dim i As Integer

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "G-System.INI", Me.Name)
    
    For i = 0 To 40
       S_F1(i).Visible = False
       S_T1(i).Visible = False
       L_F1(i).Visible = False
       L_T1(i).Visible = False
    Next i
       
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) And Gf_Sp_ProceExist(Proc_Sc("Sc2")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "G-System.INI", Me.Name)
    
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

Public Sub Spread_Can()

'    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1) And Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2) Then
'        Call Gp_Pic1_Display(Proc_Sc("Sc1")("Spread"))
'        Call Gp_Pic2_Display(Proc_Sc("Sc2")("Spread"))
'        Call DISPLAY_LOCK(ss1)
'    End If
Call Form_Ref

End Sub

Public Sub Form_Cls()

    Dim i As Integer
    
    If Gf_Sp_Cls(Proc_Sc("Sc1")) And Gf_Sp_Cls(Proc_Sc("Sc2")) Then
        Call Gp_Pic1_Display(Proc_Sc("Sc1")("Spread"))
        Call Gp_Pic2_Display(Proc_Sc("Sc2")("Spread"))
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("rControl"), False)
        Call Gp_Ms_ControlLock(Mc2("rControl"), False)
        Click_YN = False
        For i = 0 To 40
            S_F1(i).Visible = False
            S_T1(i).Visible = False
            L_F1(i).Visible = False
            L_T1(i).Visible = False
        Next i
        
        txt_location1.Text = ""
        txt_location2.Text = ""
        txt_location3.Text = ""

        rControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

   Dim i As Integer
   Dim iRowCount As Integer

On Error GoTo Refer_Err

    Dim SMESG As String

    If Gf_Sp_ProceExist(Proc_Sc("Sc1").Item("Spread")) Or Gf_Sp_ProceExist(Proc_Sc("Sc2").Item("Spread")) Then Exit Sub
    
    If Trim(TXT_T_YARD_ADDR.Text) = "C0A0101" Then
       SMESG = "钢卷入库移送线地址( C0A0101 )不能作为目的垛位 ！！！"
       Call Gp_MsgBoxDisplay(SMESG)
       TXT_T_YARD_ADDR.Text = ""
       Exit Sub
    End If

    If Trim(TXT_COIL_NO.Text) <> "" Then
    
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1) Then
                                        
                Call Gf_Ms_Refer(M_CN1, Mc1)
                Call Gp_Ms_ControlLock(Mc1("rControl"), True)
'                Call Gp_Ms_ControlLock(Mc2("rControl"), True)
                Call Gp_Pic1_Display(Proc_Sc("Sc1")("Spread"))
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call DISPLAY_LOCK(ss1)
                Exit Sub
            End If
    End If
    
    If Trim(TXT_S_YARD_ADDR.Text) <> "" Then
    
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1) Then
                
'                Call Gp_Ms_ControlLock(Mc2("rControl"), True)
                Call Gp_Pic1_Display(Proc_Sc("Sc1")("Spread"))
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call DISPLAY_LOCK(ss1)
            End If
    End If
    
    If Trim(TXT_T_YARD_ADDR.Text) <> "" Then
    
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2) Then
                Click_YN = True
                Call Gp_Ms_ControlLock(Mc2("rControl"), True)
                Call Gp_Pic2_Display(Proc_Sc("Sc2")("Spread"))
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call DISPLAY_LOCK(ss2)
            End If
    End If
    
    If Trim(TXT_S_YARD_ADDR.Text) = "C0A0101" Then
       For i = 21 To 29
           L_F1(i).Visible = False
           S_F1(i).Visible = False
       Next i
    End If
    
'    If Trim(TXT_T_YARD_ADDR.Text) = "C0A0101" Then
'       For i = 21 To 29
'           L_T1(i).Visible = False
'           S_T1(i).Visible = False
'       Next i
'    End If

    If Trim(TXT_S_YARD_ADDR.Text) = "C1B0602" Or Trim(TXT_S_YARD_ADDR.Text) = "C1B0702" Then
       For i = 10 To 13
           L_F1(i).Visible = False
           S_F1(i).Visible = False
       Next i
       For i = 30 To 34
           L_F1(i).Visible = False
           S_F1(i).Visible = False
       Next i
       With ss1
            For iRowCount = 7 To 11
                  
                     .Row = iRowCount
                     .Col = 3:          .Col2 = 3
                     .Row = iRowCount:  .Row2 = iRowCount
                     .BlockMode = True
                     .Lock = True
                     .BlockMode = False
             
             Next iRowCount
            For iRowCount = 28 To 31
                  
                     .Row = iRowCount
                     .Col = 3:          .Col2 = 3
                     .Row = iRowCount:  .Row2 = iRowCount
                     .BlockMode = True
                     .Lock = True
                     .BlockMode = False
             
             Next iRowCount
        End With
    End If
    
    If Trim(TXT_T_YARD_ADDR.Text) = "C1B0602" Or Trim(TXT_T_YARD_ADDR.Text) = "C1B0702" Then
       For i = 10 To 13
           L_T1(i).Visible = False
           S_T1(i).Visible = False
       Next i
       For i = 30 To 34
           L_T1(i).Visible = False
           S_T1(i).Visible = False
       Next i
       With ss2
            For iRowCount = 7 To 11
                  
                     .Row = iRowCount
                     .Col = 3:          .Col2 = 3
                     .Row = iRowCount:  .Row2 = iRowCount
                     .BlockMode = True
                     .Lock = True
                     .BlockMode = False
             
             Next iRowCount
            For iRowCount = 28 To 31
                  
                     .Row = iRowCount
                     .Col = 3:          .Col2 = 3
                     .Row = iRowCount:  .Row2 = iRowCount
                     .BlockMode = True
                     .Lock = True
                     .BlockMode = False
             
             Next iRowCount
        End With
    End If

    
Refer_Err:

End Sub

Public Sub Form_Pro()

   Dim i As Integer
   Dim iRowCount As Integer
    
    If Trim(TXT_COIL_NO.Text) <> "" Then
    
            If Gf_Mill_Process(M_CN1, Proc_Sc("Sc1"), Mc1, , "C") Then
            
                Call Gp_Pic1_Display(Proc_Sc("Sc1")("Spread"))
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call DISPLAY_LOCK(ss1)
                Exit Sub
                
            End If
    End If
    
    If Trim(TXT_S_YARD_ADDR.Text) <> "" Then
    
            If Gf_Mill_Process(M_CN1, Proc_Sc("Sc1"), Mc1, , "C") Then
                
                Call Gp_Pic1_Display(Proc_Sc("Sc1")("Spread"))
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call DISPLAY_LOCK(ss1)
                
            End If
    End If
    
    If Trim(TXT_T_YARD_ADDR.Text) <> "" Then
            If Gf_Mill_Process(M_CN1, Proc_Sc("Sc2"), Mc2, , "C") Then
              
                Call Gp_Pic2_Display(Proc_Sc("Sc2")("Spread"))
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call DISPLAY_LOCK(ss2)
            End If
    End If
    
    If Trim(TXT_S_YARD_ADDR.Text) = "C0A0101" Then
       For i = 21 To 29
           L_F1(i).Visible = False
           S_F1(i).Visible = False
       Next i
    End If
    
    If Trim(TXT_T_YARD_ADDR.Text) = "C0A0101" Then
       For i = 21 To 29
           L_T1(i).Visible = False
           S_T1(i).Visible = False
       Next i
    End If
    
    If Trim(TXT_S_YARD_ADDR.Text) = "C1B0602" Or Trim(TXT_S_YARD_ADDR.Text) = "C1B0702" Then
       For i = 10 To 13
           L_F1(i).Visible = False
           S_F1(i).Visible = False
       Next i
       For i = 30 To 34
           L_F1(i).Visible = False
           S_F1(i).Visible = False
       Next i
       With ss1
            For iRowCount = 7 To 11
                  
                     .Row = iRowCount
                     .Col = 3:          .Col2 = 3
                     .Row = iRowCount:  .Row2 = iRowCount
                     .BlockMode = True
                     .Lock = True
                     .BlockMode = False
             
             Next iRowCount
            For iRowCount = 28 To 31
                  
                     .Row = iRowCount
                     .Col = 3:          .Col2 = 3
                     .Row = iRowCount:  .Row2 = iRowCount
                     .BlockMode = True
                     .Lock = True
                     .BlockMode = False
             
             Next iRowCount
        End With
    End If
    
    If Trim(TXT_T_YARD_ADDR.Text) = "C1B0602" Or Trim(TXT_T_YARD_ADDR.Text) = "C1B0702" Then
       For i = 10 To 13
           L_T1(i).Visible = False
           S_T1(i).Visible = False
       Next i
       For i = 30 To 34
           L_T1(i).Visible = False
           S_T1(i).Visible = False
       Next i
       With ss2
            For iRowCount = 7 To 11
                  
                     .Row = iRowCount
                     .Col = 3:          .Col2 = 3
                     .Row = iRowCount:  .Row2 = iRowCount
                     .BlockMode = True
                     .Lock = True
                     .BlockMode = False
             
             Next iRowCount
            For iRowCount = 28 To 31
                  
                     .Row = iRowCount
                     .Col = 3:          .Col2 = 3
                     .Row = iRowCount:  .Row2 = iRowCount
                     .BlockMode = True
                     .Lock = True
                     .BlockMode = False
             
             Next iRowCount
        End With
    End If
    
End Sub

Public Sub Form_Ins()

End Sub

Public Sub Spread_Cpy()
    
End Sub

Public Sub Spread_Pst()
    
End Sub

Public Sub Spread_ColumnsSort()
    
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
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
    If Trim(TXT_COIL_NO.Text) <> "" Then
       Call Gp_Sp_Del(Proc_Sc("Sc1"))
       Exit Sub
    End If
    If Trim(TXT_S_YARD_ADDR.Text) <> "" Then
       Call Gp_Sp_Del(Proc_Sc("Sc1"))
    End If
    If Trim(TXT_T_YARD_ADDR.Text) <> "" Then
       Call Gp_Sp_Del(Proc_Sc("Sc2"))
    End If

End Sub

Private Sub L_F1_Click(Index As Integer)

   Dim iCount As Integer
   Dim iRowCount As Long
   Dim iColcount As Long
   Dim iLayer As String
   Dim iSeq As String
   Dim iText As String

   Select Case S_F1(Index).FillColor
          Case &HE0E0E0          'BLUE
               Exit Sub
          Case &HFF&       'RED
          
               For iCount = 0 To Index - 1
                   If S_F1(iCount).FillColor = &HFFFF& Then  'YELLOW
                      Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1)
                      Call DISPLAY_LOCK(ss1)
                      S_F1(iCount).FillColor = &HFF&
                      Exit For
                      Exit Sub
                   End If
               Next iCount
               For iCount = Index + 1 To 40
                   If S_F1(iCount).FillColor = &HFFFF& Then
                      Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1)
                      Call DISPLAY_LOCK(ss1)
                      S_F1(iCount).FillColor = &HFF&
                      Exit For
                      Exit Sub
                   End If
               Next iCount
               'S_F1(Index).FillColor = &HFFFF&
               
               If Val(Mid(Index, 1, 2)) < 21 Then
                  Select Case Index
                         Case 0
                         If S_F1(21).FillColor = &HFF& Then
                             Call Gp_MsgBoxDisplay("请先移动上层钢卷再移动此钢卷        ", "W", " 注意了")
                            Exit Sub
                         End If
                         Case 20
                         If S_F1(40).FillColor = &HFF& Then
                             Call Gp_MsgBoxDisplay("请先移动上层钢卷再移动此钢卷        ", "W", " 注意了")
                            Exit Sub
                         End If
                         Case Else
                         If S_F1(Index + 20).FillColor = &HFF& Or S_F1(Index + 21).FillColor = &HFF& Then
                             Call Gp_MsgBoxDisplay("请先移动上层钢卷再移动此钢卷        ", "W", " 注意了")
                            Exit Sub
                         End If
                  End Select
                  iLayer = 1
                  iSeq = Index + 1
                  For iRowCount = 0 To ss1.MaxRows - 1

                      iText = ""
                      ss1.Row = iRowCount + 1
                      For iColcount = 1 To 2
                          ss1.Col = iColcount
                          iText = Trim(iText) + Trim(ss1.Text)
                      Next iColcount
                      
                      If iText = iLayer + iSeq Then
                         For iColcount = 1 To 3 'ss1.MaxCols
                         ss1.Col = iColcount
                         ss1.BackColor = &HFF& '&HFFFF&
                         Next iColcount
                         ss1.Col = 0
                         ss1.Text = "Delete"
                         ss1.Col = 3
                         ScoilNo = ss1.Text

                      End If
                  
                  Next iRowCount
                  For iRowCount = 0 To ss2.MaxRows - 1

                      iText = ""
                      ss2.Row = iRowCount + 1
                      ss2.Col = 0
                      If ss2.Text = "Input" Then
                         ss2.Col = 3
                         ss2.Text = ScoilNo
                      End If
                                        
                  Next iRowCount
                  
               Else
                  iLayer = 2
                  iSeq = Index - 20
                  For iRowCount = 0 To ss1.MaxRows - 1

                      iText = ""
                      ss1.Row = iRowCount + 1
                      For iColcount = 1 To 2
                          ss1.Col = iColcount
                          iText = Trim(iText) + Trim(ss1.Text)
                      Next iColcount
                      If iText = iLayer + iSeq Then
                         For iColcount = 1 To 3 'ss1.MaxCols
                         ss1.Col = iColcount
                         ss1.BackColor = &HFF& ' &HFFFF&
                         Next iColcount
                         ss1.Col = 0
                         ss1.Text = "Delete"
                         ss1.Col = 3
                         ScoilNo = ss1.Text

                      End If
                  
                  Next iRowCount
                  For iRowCount = 0 To ss2.MaxRows - 1

                      iText = ""
                      ss2.Row = iRowCount + 1
                      ss2.Col = 0
                      If ss2.Text = "Input" Then
                         ss2.Col = 3
                         ss2.Text = ScoilNo
                      End If
                                        
                  Next iRowCount
               End If
               
               S_F1(Index).FillColor = &HFFFF&

          Case &HFFFF&
               S_F1(Index).FillColor = &HFF&
               Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1)
               Call DISPLAY_LOCK(ss1)
               For iCount = 0 To 40
                   If S_T1(iCount).FillColor = &HFFFF& Then  'YELLOW
                      S_T1(iCount).FillColor = &HE0E0E0
                      Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2)
                      Call DISPLAY_LOCK(ss2)
                   End If
               Next iCount
          End Select

End Sub

Private Sub L_T1_Click(Index As Integer)

   Dim iCount As Integer
   Dim iRowCount As Long
   Dim iColcount As Long
   Dim iLayer As String
   Dim iSeq As String
   Dim iText As String
   Dim iSign As Boolean
   
   If Click_YN = False Then Exit Sub
   
   For iCount = 0 To 40
        If S_F1(iCount).FillColor = &HFFFF& Then
            iSign = True
            Exit For
        End If
   Next iCount

   If iSign = False Then
       Exit Sub
   End If

   Select Case S_T1(Index).FillColor
          Case &HFF&       'RED
               Exit Sub
          Case &HE0E0E0          'BLUE
          
               For iCount = 0 To Index - 1
                   If S_T1(iCount).FillColor = &HFFFF& Then  'YELLOW
                      
                      Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2)
                       Call DISPLAY_LOCK(ss2)
                      S_T1(iCount).FillColor = &HE0E0E0
                      Exit For
                      Exit Sub
                   End If
               Next iCount
               For iCount = Index + 1 To 40
                   If S_T1(iCount).FillColor = &HFFFF& Then
                      Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2)
                       Call DISPLAY_LOCK(ss2)
                      S_T1(iCount).FillColor = &HE0E0E0
                      Exit For
                      Exit Sub
                   End If
               Next iCount
              ' S_T1(Index).FillColor = &HFFFF&
               
               If Val(Mid(Index, 1, 2)) < 21 Then
                  iLayer = 1
                  iSeq = Index + 1
                  For iRowCount = 0 To ss2.MaxRows - 1

                      iText = ""
                      ss2.Row = iRowCount + 1
                      For iColcount = 1 To 2
                          ss2.Col = iColcount
                          iText = Trim(iText) + Trim(ss2.Text)
                      Next iColcount
                      
                      If iText = iLayer + iSeq Then
                         For iColcount = 1 To 3 'ss2.MaxCols
                         ss2.Col = iColcount
                         ss2.BackColor = &HFF& '&HFFFF&
                         Next iColcount
                         ss2.Col = 0
                         ss2.Text = "Input"
                         ss2.Col = 3
                         ss2.Text = ScoilNo
                      End If
                  
                  Next iRowCount
               Else
                  If S_T1(Index - 21).FillColor = &HE0E0E0 Or S_T1(Index - 20).FillColor = &HE0E0E0 Then
                     Call Gp_MsgBoxDisplay("请逐层堆放钢卷          ", "W", " 注意了")
                     Exit Sub
                  End If
                  iLayer = 2
                  iSeq = Index - 20
                  For iRowCount = 0 To ss2.MaxRows - 1

                      iText = ""
                      ss2.Row = iRowCount + 1
                      For iColcount = 1 To 2
                          ss2.Col = iColcount
                          iText = Trim(iText) + Trim(ss2.Text)
                      Next iColcount
                      If iText = iLayer + iSeq Then
                         For iColcount = 1 To 3 'ss2.MaxCols
                         ss2.Col = iColcount
                         ss2.BackColor = &HFF& '&HFFFF&
                         Next iColcount
                         ss2.Col = 0
                         ss2.Text = "Input"
                         ss2.Col = 3
                         ss2.Text = ScoilNo
                      End If
                  
                  Next iRowCount
               End If
               
               S_T1(Index).FillColor = &HFFFF&

          Case &HFFFF&
               S_T1(Index).FillColor = &HE0E0E0
               Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2)
                Call DISPLAY_LOCK(ss2)
    End Select
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    If ss1.MaxRows < 1 Then Exit Sub
    
    With ss1
         If Col = 3 Then
            .Row = Row
            .Col = 3
            If Trim(.Text) = "" Then
               .Col = 0
               .Text = "Input"
            Else
               .Col = 0
               .Text = "Update"
            End If
         End If
    End With
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

End Sub
Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC1")("Spread"), Mode)
     
    End If
    
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
    If ss2.MaxRows < 1 Then Exit Sub
    
    With ss2
         If Col = 3 Then
            .Row = Row
            .Col = 3
            If Trim(.Text) = "" Then
               .Col = 0
               .Text = "Input"
            Else
               .Col = 0
               .Text = "Update"
            End If
         End If
    End With
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC2")("Spread"), Mode)
       
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc1")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc1"))
    End If

    If Shift = 0 Then Proc_Sc("Sc1")("Spread").EditMode = True

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

Public Sub Gp_Pic1_Display(sPname As Variant, Optional Cls As Boolean = True)

On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim iText As String
    Dim iValue As String
    Dim CNT As Integer
    
    ss1.Row = ss1.MaxRows
    ss1.Col = 14
    CNT = Val(ss1.Text)
    
    If Cls = True Then
        For iCount = 0 To CNT - 1
            S_F1(iCount).Visible = True
            L_F1(iCount).Visible = True
            S_F1(iCount).FillColor = &HE0E0E0
        Next iCount
        For iCount = 0 To CNT - 2
            S_F1(iCount + 21).Visible = True
            L_F1(iCount + 21).Visible = True
            S_F1(iCount + 21).FillColor = &HE0E0E0
        Next
    End If
    
    With sPname
             
         For iRowCount = 0 To .MaxRows - 1
              iText = ""
             .Row = iRowCount + 1
              For iColcount = 1 To 2
                  .Col = iColcount
                  iText = Trim(iText) + Trim(.Text)
              Next iColcount
              .Col = 3
              iValue = Trim(.Text)
              
              Select Case iText
                     Case "11"
                          If iValue <> "" Then
                          S_F1(0).FillColor = &HFF&
                          End If
                     Case "12"
                          If iValue <> "" Then
                          S_F1(1).FillColor = &HFF&
                          End If
                     Case "13"
                          If iValue <> "" Then
                          S_F1(2).FillColor = &HFF&
                          End If
                     Case "14"
                          If iValue <> "" Then
                          S_F1(3).FillColor = &HFF&
                          End If
                     Case "15"
                          If iValue <> "" Then
                          S_F1(4).FillColor = &HFF&
                          End If
                     Case "16"
                          If iValue <> "" Then
                          S_F1(5).FillColor = &HFF&
                          End If
                     Case "17"
                          If iValue <> "" Then
                          S_F1(6).FillColor = &HFF&
                          End If
                     Case "18"
                          If iValue <> "" Then
                          S_F1(7).FillColor = &HFF&
                          End If
                     Case "19"
                          If iValue <> "" Then
                          S_F1(8).FillColor = &HFF&
                          End If
                     Case "110"
                          If iValue <> "" Then
                          S_F1(9).FillColor = &HFF&
                          End If
                     Case "111"
                          If iValue <> "" Then
                          S_F1(10).FillColor = &HFF&
                          End If
                     Case "112"
                          If iValue <> "" Then
                          S_F1(11).FillColor = &HFF&
                          End If
                     Case "113"
                          If iValue <> "" Then
                          S_F1(12).FillColor = &HFF&
                          End If
                     Case "114"
                          If iValue <> "" Then
                          S_F1(13).FillColor = &HFF&
                          End If
                     Case "115"
                          If iValue <> "" Then
                          S_F1(14).FillColor = &HFF&
                          End If
                     Case "116"
                          If iValue <> "" Then
                          S_F1(15).FillColor = &HFF&
                          End If
                     Case "117"
                          If iValue <> "" Then
                          S_F1(16).FillColor = &HFF&
                          End If
                     Case "118"
                          If iValue <> "" Then
                          S_F1(17).FillColor = &HFF&
                          End If
                     Case "119"
                          If iValue <> "" Then
                          S_F1(18).FillColor = &HFF&
                          End If
                     Case "120"
                          If iValue <> "" Then
                          S_F1(19).FillColor = &HFF&
                          End If
                    Case "121"
                          If iValue <> "" Then
                          S_F1(20).FillColor = &HFF&
                          End If
                    Case "21"
                          If iValue <> "" Then
                          S_F1(21).FillColor = &HFF&
                          End If
                     Case "22"
                          If iValue <> "" Then
                          S_F1(22).FillColor = &HFF&
                          End If
                     Case "23"
                          If iValue <> "" Then
                          S_F1(23).FillColor = &HFF&
                          End If
                     Case "24"
                          If iValue <> "" Then
                          S_F1(24).FillColor = &HFF&
                          End If
                     Case "25"
                          If iValue <> "" Then
                          S_F1(25).FillColor = &HFF&
                          End If
                     Case "26"
                          If iValue <> "" Then
                          S_F1(26).FillColor = &HFF&
                          End If
                     Case "27"
                          If iValue <> "" Then
                          S_F1(27).FillColor = &HFF&
                          End If
                     Case "28"
                          If iValue <> "" Then
                          S_F1(28).FillColor = &HFF&
                          End If
                     Case "29"
                          If iValue <> "" Then
                          S_F1(29).FillColor = &HFF&
                          End If
                     Case "210"
                          If iValue <> "" Then
                          S_F1(30).FillColor = &HFF&
                          End If
                     Case "211"
                          If iValue <> "" Then
                          S_F1(31).FillColor = &HFF&
                          End If
                     Case "212"
                          If iValue <> "" Then
                          S_F1(32).FillColor = &HFF&
                          End If
                     Case "213"
                          If iValue <> "" Then
                          S_F1(33).FillColor = &HFF&
                          End If
                     Case "214"
                          If iValue <> "" Then
                          S_F1(34).FillColor = &HFF&
                          End If
                     Case "215"
                          If iValue <> "" Then
                          S_F1(35).FillColor = &HFF&
                          End If
                     Case "216"
                          If iValue <> "" Then
                          S_F1(36).FillColor = &HFF&
                          End If
                     Case "217"
                          If iValue <> "" Then
                          S_F1(37).FillColor = &HFF&
                          End If
                     Case "218"
                          If iValue <> "" Then
                          S_F1(38).FillColor = &HFF&
                          End If
                     Case "219"
                          If iValue <> "" Then
                          S_F1(39).FillColor = &HFF&
                          End If
                     Case "220"
                          If iValue <> "" Then
                          S_F1(40).FillColor = &HFF&
                          End If

              End Select
              
         Next iRowCount
 
    End With
    
SpreadDisplay_Error:

End Sub

Public Sub Gp_Pic2_Display(sPname As Variant, Optional Cls As Boolean = True)

On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim iText As String
    Dim iValue As String
    Dim CNT  As Integer
    
    ss2.Row = ss2.MaxRows
    ss2.Col = 14
    CNT = Val(ss2.Text)
    
    If Cls = True Then
        For iCount = 0 To CNT - 1
            S_T1(iCount).Visible = True
            S_T1(iCount).FillColor = &HE0E0E0
            L_T1(iCount).Visible = True
        Next iCount
          For iCount = 0 To CNT - 2
            S_T1(iCount + 21).Visible = True
            S_T1(iCount + 21).FillColor = &HE0E0E0
            L_T1(iCount + 21).Visible = True
        Next iCount
      
        
    End If
    
    With sPname
             
         For iRowCount = 0 To .MaxRows - 1
              iText = ""
             .Row = iRowCount + 1
              For iColcount = 1 To 2
                  .Col = iColcount
                  iText = Trim(iText) + Trim(.Text)
              Next iColcount
              .Col = 3
              iValue = Trim(.Text)
              
              Select Case iText
                     Case "11"
                          If iValue <> "" Then
                          S_T1(0).FillColor = &HFF&
                          End If
                     Case "12"
                          If iValue <> "" Then
                          S_T1(1).FillColor = &HFF&
                          End If
                     Case "13"
                          If iValue <> "" Then
                          S_T1(2).FillColor = &HFF&
                          End If
                     Case "14"
                          If iValue <> "" Then
                          S_T1(3).FillColor = &HFF&
                          End If
                     Case "15"
                          If iValue <> "" Then
                          S_T1(4).FillColor = &HFF&
                          End If
                     Case "16"
                          If iValue <> "" Then
                          S_T1(5).FillColor = &HFF&
                          End If
                     Case "17"
                          If iValue <> "" Then
                          S_T1(6).FillColor = &HFF&
                          End If
                     Case "18"
                          If iValue <> "" Then
                          S_T1(7).FillColor = &HFF&
                          End If
                     Case "19"
                          If iValue <> "" Then
                          S_T1(8).FillColor = &HFF&
                          End If
                     Case "110"
                          If iValue <> "" Then
                          S_T1(9).FillColor = &HFF&
                          End If
                     Case "111"
                          If iValue <> "" Then
                          S_T1(10).FillColor = &HFF&
                          End If
                     Case "112"
                          If iValue <> "" Then
                          S_T1(11).FillColor = &HFF&
                          End If
                     Case "113"
                          If iValue <> "" Then
                          S_T1(12).FillColor = &HFF&
                          End If
                     Case "114"
                          If iValue <> "" Then
                          S_T1(13).FillColor = &HFF&
                          End If
                     Case "115"
                          If iValue <> "" Then
                          S_T1(14).FillColor = &HFF&
                          End If
                     Case "116"
                          If iValue <> "" Then
                          S_T1(15).FillColor = &HFF&
                          End If
                     Case "117"
                          If iValue <> "" Then
                          S_T1(16).FillColor = &HFF&
                          End If
                     Case "118"
                          If iValue <> "" Then
                          S_T1(17).FillColor = &HFF&
                          End If
                     Case "119"
                          If iValue <> "" Then
                          S_T1(18).FillColor = &HFF&
                          End If
                     Case "120"
                          If iValue <> "" Then
                          S_T1(19).FillColor = &HFF&
                          End If
                    Case "121"
                          If iValue <> "" Then
                          S_T1(20).FillColor = &HFF&
                          End If
                    Case "21"
                          If iValue <> "" Then
                          S_T1(21).FillColor = &HFF&
                          End If
                     Case "22"
                          If iValue <> "" Then
                          S_T1(22).FillColor = &HFF&
                          End If
                     Case "23"
                          If iValue <> "" Then
                          S_T1(23).FillColor = &HFF&
                          End If
                     Case "24"
                          If iValue <> "" Then
                          S_T1(24).FillColor = &HFF&
                          End If
                     Case "25"
                          If iValue <> "" Then
                          S_T1(25).FillColor = &HFF&
                          End If
                     Case "26"
                          If iValue <> "" Then
                          S_T1(26).FillColor = &HFF&
                          End If
                     Case "27"
                          If iValue <> "" Then
                          S_T1(27).FillColor = &HFF&
                          End If
                     Case "28"
                          If iValue <> "" Then
                          S_T1(28).FillColor = &HFF&
                          End If
                     Case "29"
                          If iValue <> "" Then
                          S_T1(29).FillColor = &HFF&
                          End If
                     Case "210"
                          If iValue <> "" Then
                          S_T1(30).FillColor = &HFF&
                          End If
                     Case "211"
                          If iValue <> "" Then
                          S_T1(31).FillColor = &HFF&
                          End If
                     Case "212"
                          If iValue <> "" Then
                          S_T1(32).FillColor = &HFF&
                          End If
                     Case "213"
                          If iValue <> "" Then
                          S_T1(33).FillColor = &HFF&
                          End If
                     Case "214"
                          If iValue <> "" Then
                          S_T1(34).FillColor = &HFF&
                          End If
                     Case "215"
                          If iValue <> "" Then
                          S_T1(35).FillColor = &HFF&
                          End If
                     Case "216"
                          If iValue <> "" Then
                          S_T1(36).FillColor = &HFF&
                          End If
                     Case "217"
                          If iValue <> "" Then
                          S_T1(37).FillColor = &HFF&
                          End If
                     Case "218"
                          If iValue <> "" Then
                          S_T1(38).FillColor = &HFF&
                          End If
                     Case "219"
                          If iValue <> "" Then
                          S_T1(39).FillColor = &HFF&
                          End If
                     Case "220"
                          If iValue <> "" Then
                          S_T1(40).FillColor = &HFF&
                          End If
                          
              End Select
              
         Next iRowCount
 
    End With
    
SpreadDisplay_Error:

End Sub

Private Sub txt_location1_DblClick()
    TXT_T_YARD_ADDR.Text = txt_location1.Text
    Call Form_Ref
End Sub

Private Sub txt_location2_DblClick()
    TXT_T_YARD_ADDR.Text = txt_location2.Text
    Call Form_Ref
End Sub

Private Sub txt_location3_DblClick()
    TXT_T_YARD_ADDR.Text = txt_location3.Text
    Call Form_Ref
End Sub

Private Sub TXT_S_YARD_ADDR_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "F0009"
        TXT_S_YARD_ADDR.Text = "C"
        DD.rControl.Add Item:=TXT_S_YARD_ADDR
              
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

End Sub
Private Sub TXT_T_YARD_ADDR_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "F0009"
        TXT_T_YARD_ADDR.Text = "C"
        DD.rControl.Add Item:=TXT_T_YARD_ADDR
               
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

End Sub

Private Sub TXT_T_YARD_ADDR_GotFocus()

   TXT_COIL_NO.Text = ""
   
End Sub

Private Sub DISPLAY_LOCK(sPname As Variant)

    Dim iText1 As String
    Dim iText2 As String
    Dim iRowCount As Integer
    Dim iRow As Integer
    Dim iMaxup As String
    Dim iMax, iCount As Integer


    With sPname
    
        .Col = 14
        .Row = 1
        iMaxup = .Text
        iMax = .MaxRows
        iCount = iMax - Val(iMaxup)
        
        For iRowCount = 1 To iCount
        
             iRow = iRowCount
             
        '        .Row = iRow
        '        .Col = 3:        .Col2 = 3
        '        .Row = iRow:     .Row2 = iRow
        '        .BlockMode = True
        '        .Lock = False
        '        .BlockMode = False
                
             iText1 = ""
             iText2 = ""
             .Col = 3
             .Row = iRowCount + iCount
             iText1 = .Text
             .Row = iRowCount + iCount + 1
             iText2 = .Text
             
             If Trim(iText1) = "" Or Trim(iText2) = "" Then
             
                .Row = iRow
                .Col = 3:        .Col2 = 3
                .Row = iRow:     .Row2 = iRow
                .BlockMode = True
                .Lock = True
                .BlockMode = False
             
             End If
        
        Next iRowCount
    End With

End Sub
