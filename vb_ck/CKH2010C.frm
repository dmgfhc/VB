VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CKH2010C 
   Caption         =   "轧钢生产线进程现状界面_CKH2010C"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   290.471
   ScaleMode       =   0  'User
   ScaleWidth      =   607.688
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_CB_LINE 
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
      ItemData        =   "CKH2010C.frx":0000
      Left            =   19530
      List            =   "CKH2010C.frx":0010
      TabIndex        =   33
      Tag             =   "加热炉号"
      Text            =   "1#"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox CBO_RHF_LINE 
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
      ItemData        =   "CKH2010C.frx":0024
      Left            =   17085
      List            =   "CKH2010C.frx":0034
      TabIndex        =   32
      Tag             =   "加热炉号"
      Text            =   "2#"
      Top             =   3975
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_slabcnt2 
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
      Left            =   5205
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   780
      Width           =   705
   End
   Begin VB.TextBox txt_slabcnt1 
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
      Left            =   1695
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   780
      Width           =   705
   End
   Begin VB.TextBox txt_mill 
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
      Left            =   18510
      TabIndex        =   29
      Top             =   3240
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.TextBox txt_PLATE_LEN_DS2 
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
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   8040
      Width           =   1005
   End
   Begin VB.TextBox txt_PLATE_NO_DS2 
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
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txt_SURF_GRD_DS2 
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
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   " "
      Top             =   8040
      Width           =   795
   End
   Begin VB.TextBox txt_MPLATE_LEN_R 
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
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   8760
      Width           =   1260
   End
   Begin VB.TextBox txt_MPLATE_NO_R 
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
      Left            =   4185
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   8760
      Width           =   1530
   End
   Begin VB.TextBox txt_MPLATE_LEN_L 
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
      Left            =   8940
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   8760
      Width           =   1260
   End
   Begin VB.TextBox txt_MPLATE_NO_L 
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
      Left            =   7395
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   8760
      Width           =   1530
   End
   Begin VB.TextBox txt_MPLATE_LEN_RST 
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   8040
      Width           =   1260
   End
   Begin VB.TextBox txt_MPLATE_NO_RST 
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
      Left            =   5655
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   8040
      Width           =   1530
   End
   Begin VB.TextBox txt_SLAB_SIZE3 
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
      Left            =   7935
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4140
      Width           =   1875
   End
   Begin VB.TextBox TXT_SLAB_NO3 
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
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4140
      Width           =   1260
   End
   Begin VB.TextBox txt_SLAB_SIZE2 
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
      Left            =   4635
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4140
      Width           =   1890
   End
   Begin VB.TextBox TXT_SLAB_NO2 
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
      Left            =   3375
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   4140
      Width           =   1260
   End
   Begin VB.TextBox txt_UST_DEC 
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
      Left            =   14415
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4140
      Width           =   840
   End
   Begin VB.TextBox txt_PLATE_NO 
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
      Left            =   12690
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4140
      Width           =   1710
   End
   Begin VB.TextBox txt_SURF_GRD_DS1 
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
      Left            =   3030
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   8760
      Width           =   795
   End
   Begin VB.TextBox txt_PLATE_NO_DS1 
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
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   8760
      Width           =   1935
   End
   Begin VB.TextBox txt_MPLATE_NO 
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
      Left            =   9975
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4140
      Width           =   1590
   End
   Begin VB.TextBox txt_MPLATE_LEN 
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
      Left            =   11574
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4140
      Width           =   945
   End
   Begin VB.TextBox txt_PLATE_LEN_DS1 
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
      Left            =   2025
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   8760
      Width           =   1005
   End
   Begin VB.TextBox txt_ONC_CNT2 
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
      Left            =   17760
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3270
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.TextBox txt_ONC_CNT1 
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
      Left            =   1380
      TabIndex        =   5
      Top             =   4620
      Width           =   765
   End
   Begin VB.TextBox TXT_SLAB_NO1 
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
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4140
      Width           =   1260
   End
   Begin VB.TextBox txt_SLAB_SIZE1 
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
      Left            =   1335
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4140
      Width           =   1875
   End
   Begin VB.TextBox txt_INF_CNT1 
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
      Left            =   8880
      TabIndex        =   2
      Top             =   780
      Width           =   765
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "动态进程"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   6390
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "静态进程"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   7740
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   1950
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   2370
      Top             =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   7605
      Top             =   780
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "入炉块数"
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
      Left            =   75
      Top             =   3825
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      Caption         =   "粗轧:板坯号    厚 * 宽 * 长  "
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   15690
      Top             =   3270
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      Caption         =   "2#冷床 | 母板张数"
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
   Begin InDate.ULabel ULabel18 
      Height          =   315
      Left            =   90
      Top             =   8445
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      Caption         =   "钢板号(1#定尺剪) |   长   | 等级"
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   9984
      Top             =   3825
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      Caption         =   "分段剪: 母板号   长度"
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
      Left            =   12690
      Top             =   3825
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      Caption         =   "    UST钢板号   |  等级"
      Alignment       =   0
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   2565
      Left            =   75
      TabIndex        =   14
      Top             =   1170
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   4524
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      PaneTree        =   "CKH2010C.frx":0048
      Begin FPSpread.vaSpread ss4 
         Height          =   2565
         Left            =   11280
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   0
         Width           =   3915
         _Version        =   393216
         _ExtentX        =   6906
         _ExtentY        =   4524
         _StockProps     =   64
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
         MaxCols         =   3
         MaxRows         =   12
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CKH2010C.frx":00DA
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   2565
         Left            =   0
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   0
         Width           =   3540
         _Version        =   393216
         _ExtentX        =   6244
         _ExtentY        =   4524
         _StockProps     =   64
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
         MaxCols         =   3
         MaxRows         =   12
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CKH2010C.frx":04F1
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   2565
         Left            =   3600
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   0
         Width           =   3870
         _Version        =   393216
         _ExtentX        =   6826
         _ExtentY        =   4524
         _StockProps     =   64
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
         MaxCols         =   3
         MaxRows         =   12
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CKH2010C.frx":0908
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   2565
         Left            =   7530
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   0
         Width           =   3690
         _Version        =   393216
         _ExtentX        =   6509
         _ExtentY        =   4524
         _StockProps     =   64
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
         MaxCols         =   3
         MaxRows         =   12
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CKH2010C.frx":0D03
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   3378
      Top             =   3825
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      Caption         =   "精轧:板坯号    厚 * 宽 * 长  "
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
      Left            =   6681
      Top             =   3825
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      Caption         =   "热矫直:板坯号   厚 * 宽 * 长  "
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
   Begin SSSplitter.SSSplitter Sp1 
      Height          =   2565
      Left            =   75
      TabIndex        =   19
      Top             =   5010
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   4524
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      PaneTree        =   "CKH2010C.frx":111A
      Begin FPSpread.vaSpread ss5 
         Height          =   2565
         Left            =   0
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   0
         Width           =   3600
         _Version        =   393216
         _ExtentX        =   6350
         _ExtentY        =   4524
         _StockProps     =   64
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
         MaxCols         =   4
         MaxRows         =   12
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CKH2010C.frx":11AC
      End
      Begin FPSpread.vaSpread ss6 
         Height          =   2565
         Left            =   3660
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   0
         Width           =   3810
         _Version        =   393216
         _ExtentX        =   6720
         _ExtentY        =   4524
         _StockProps     =   64
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
         MaxCols         =   4
         MaxRows         =   12
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CKH2010C.frx":1623
      End
      Begin FPSpread.vaSpread ss7 
         Height          =   2565
         Left            =   7530
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Width           =   3705
         _Version        =   393216
         _ExtentX        =   6535
         _ExtentY        =   4524
         _StockProps     =   64
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
         MaxCols         =   4
         MaxRows         =   12
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CKH2010C.frx":1AAC
      End
      Begin FPSpread.vaSpread ss8 
         Height          =   2565
         Left            =   11295
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   0
         Width           =   3900
         _Version        =   393216
         _ExtentX        =   6879
         _ExtentY        =   4524
         _StockProps     =   64
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
         MaxCols         =   4
         MaxRows         =   12
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "CKH2010C.frx":1F35
      End
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   7395
      Top             =   8445
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      Caption         =   "左切剪: 母板号    长度"
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   4185
      Top             =   8445
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      Caption         =   "右切剪: 母板号    长度"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   90
      Top             =   7710
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      Caption         =   "钢板号(2#定尺剪) |   长   | 等级"
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
      Left            =   90
      Top             =   4620
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "母板张数"
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
      Left            =   5655
      Top             =   7710
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      Caption         =   "圆盘剪: 母板号    长度"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   135
      Top             =   780
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "库存炼钢坯数"
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
      Left            =   3645
      Top             =   780
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      Caption         =   "库存板卷坯数"
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
      Left            =   15795
      Top             =   3975
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "加热炉号"
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
   Begin InDate.ULabel ULabel15 
      Height          =   315
      Left            =   18240
      Top             =   3960
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "冷床号"
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
   Begin Threed.SSCommand Cmd_Edit 
      Height          =   360
      Left            =   11400
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   330
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "进程刷新"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      X1              =   44.152
      X2              =   105.16
      Y1              =   4.469
      Y2              =   4.469
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   14340
      Picture         =   "CKH2010C.frx":23BE
      Stretch         =   -1  'True
      Top             =   135
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   105
      Picture         =   "CKH2010C.frx":2B74
      Stretch         =   -1  'True
      Top             =   135
      Width           =   915
   End
End
Attribute VB_Name = "CKH2010C"
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
'-- Program Name      轧钢生产线进程现状界面
'-- Program ID        CKH2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2007.10.29
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
Public sDateTime As String          'Active Form Authority Setting

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

Dim pColumn7 As New Collection      'Spread Primary Key Collection
Dim nColumn7 As New Collection      'Spread necessary Column Collection
Dim mColumn7 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn7 As New Collection      'Spread Insert Column Collection
Dim aColumn7 As New Collection      'Master -> Spread Column Collection
Dim lColumn7 As New Collection      'Spread Lock Column Collection

Dim pColumn8 As New Collection      'Spread Primary Key Collection
Dim nColumn8 As New Collection      'Spread necessary Column Collection
Dim mColumn8 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn8 As New Collection      'Spread Insert Column Collection
Dim aColumn8 As New Collection      'Master -> Spread Column Collection
Dim lColumn8 As New Collection      'Spread Lock Column Collection

Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
Dim sc5 As New Collection           'Spread Collection
Dim sc6 As New Collection           'Spread Collection
Dim sc7 As New Collection           'Spread Collection
Dim sc8 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

   'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "MSheet"
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       
         Call Gp_Ms_Collection(txt_INF_CNT1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_SLAB_NO1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_SLAB_SIZE1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_SLAB_NO2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_SLAB_SIZE2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_SLAB_NO3, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_SLAB_SIZE3, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_MPLATE_NO, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_MPLATE_LEN, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_PLATE_NO, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_UST_DEC, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_ONC_CNT1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_ONC_CNT2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
           
    Call Gp_Ms_Collection(txt_MPLATE_NO_RST, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(txt_MPLATE_LEN_RST, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_MPLATE_NO_L, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_MPLATE_LEN_L, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_MPLATE_NO_R, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_MPLATE_LEN_R, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_PLATE_NO_DS1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_PLATE_LEN_DS1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_SURF_GRD_DS1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_PLATE_NO_DS2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_PLATE_LEN_DS2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(txt_SURF_GRD_DS2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_slabcnt1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_slabcnt2, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    'MASTER Collection
     Mc1.Add Item:="CKH2010C.P_REFER1", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
  
         Call Gp_Ms_Collection(CBO_RHF_LINE, "p", "n", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(CBO_CB_LINE, "p", "n", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         
    'MASTER Collection
     Mc2.Add Item:="CKH2010C.P_REFER2", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
  
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    
    Call Gp_Sp_Collection(ss5, 1, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 2, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 3, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 4, " ", " ", " ", " ", " ", " ", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    
    Call Gp_Sp_Collection(ss6, 1, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 2, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 3, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Call Gp_Sp_Collection(ss6, 4, " ", " ", " ", " ", " ", " ", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    
    Call Gp_Sp_Collection(ss7, 1, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 2, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 3, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    Call Gp_Sp_Collection(ss7, 4, " ", " ", " ", " ", " ", " ", pColumn7, nColumn7, mColumn7, iColumn7, aColumn7, lColumn7)
    
    Call Gp_Sp_Collection(ss8, 1, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 2, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 3, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    Call Gp_Sp_Collection(ss8, 4, " ", " ", " ", " ", " ", " ", pColumn8, nColumn8, mColumn8, iColumn8, aColumn8, lColumn8)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="CKH2010C.P_SREFER1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=2, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="CKH2010C.P_SREFER2", Key:="P-R"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=2, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
    
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="CKH2010C.P_SREFER3", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=2, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="CKH2010C.P_SREFER4", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=2, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc4, Key:="Sc4"
    
    sc5.Add Item:=ss5, Key:="Spread"
    sc5.Add Item:="CKH2010C.P_SREFER5", Key:="P-R"
    sc5.Add Item:=pColumn5, Key:="pColumn"
    sc5.Add Item:=nColumn5, Key:="nColumn"
    sc5.Add Item:=aColumn5, Key:="aColumn"
    sc5.Add Item:=mColumn5, Key:="mColumn"
    sc5.Add Item:=iColumn5, Key:="iColumn"
    sc5.Add Item:=lColumn5, Key:="lColumn"
    sc5.Add Item:=2, Key:="First"
    sc5.Add Item:=ss5.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc5, Key:="Sc5"
    
    sc6.Add Item:=ss6, Key:="Spread"
    sc6.Add Item:="CKH2010C.P_SREFER6", Key:="P-R"
    sc6.Add Item:=pColumn6, Key:="pColumn"
    sc6.Add Item:=nColumn6, Key:="nColumn"
    sc6.Add Item:=aColumn6, Key:="aColumn"
    sc6.Add Item:=mColumn6, Key:="mColumn"
    sc6.Add Item:=iColumn6, Key:="iColumn"
    sc6.Add Item:=lColumn6, Key:="lColumn"
    sc6.Add Item:=2, Key:="First"
    sc6.Add Item:=ss6.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc6, Key:="Sc6"
    
    sc7.Add Item:=ss7, Key:="Spread"
    sc7.Add Item:="CKH2010C.P_SREFER7", Key:="P-R"
    sc7.Add Item:=pColumn7, Key:="pColumn"
    sc7.Add Item:=nColumn7, Key:="nColumn"
    sc7.Add Item:=aColumn7, Key:="aColumn"
    sc7.Add Item:=mColumn7, Key:="mColumn"
    sc7.Add Item:=iColumn7, Key:="iColumn"
    sc7.Add Item:=lColumn7, Key:="lColumn"
    sc7.Add Item:=2, Key:="First"
    sc7.Add Item:=ss7.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc7, Key:="Sc7"
    
    sc8.Add Item:=ss8, Key:="Spread"
    sc8.Add Item:="CKH2010C.P_SREFER8", Key:="P-R"
    sc8.Add Item:=pColumn8, Key:="pColumn"
    sc8.Add Item:=nColumn8, Key:="nColumn"
    sc8.Add Item:=aColumn8, Key:="aColumn"
    sc8.Add Item:=mColumn8, Key:="mColumn"
    sc8.Add Item:=iColumn8, Key:="iColumn"
    sc8.Add Item:=lColumn8, Key:="lColumn"
    sc8.Add Item:=2, Key:="First"
    sc8.Add Item:=ss8.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc8, Key:="Sc8"

      
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub CBO_CB_LINE_Click()

  If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc4"), Mc2, , , False) Then
     Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
  End If
       
       txt_ONC_CNT1.Text = ss4.MaxRows
  
End Sub

Private Sub CBO_RHF_LINE_Click()

  If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc2, , , False) Then
     Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
  End If
  
     txt_INF_CNT1.Text = ss1.MaxRows
  
End Sub

Private Sub Cmd_Edit_Click()

    Call Form_Ref1
    
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.StatusBar1.Panels(1) = "Message : "
    
    Call Form_Ref1
'    Call Form_Ref2
    
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_ControlLock(Mc2("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc4")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc5")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc6")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc7")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc8")("Spread"))
    
    ss1.ColWidth(0) = 3
    ss2.ColWidth(0) = 3
    ss3.ColWidth(0) = 3
    ss4.ColWidth(0) = 3
    ss5.ColWidth(0) = 3
    ss6.ColWidth(0) = 3
    ss7.ColWidth(0) = 3
    ss8.ColWidth(0) = 3

    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    Call Gf_Sp_Cls(Proc_Sc("Sc4"))
    Call Gf_Sp_Cls(Proc_Sc("Sc5"))
    Call Gf_Sp_Cls(Proc_Sc("Sc6"))
    Call Gf_Sp_Cls(Proc_Sc("Sc7"))
    Call Gf_Sp_Cls(Proc_Sc("Sc8"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "CG-System.INI", Me.Name)
    ss1.RowHeight(0) = 18
    ss1.Col = -1: ss1.Col2 = -1
    ss1.ROW = 0: ss1.Row2 = 0
    ss1.FontSize = 9
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "CG-System.INI", Me.Name)
    ss2.RowHeight(0) = 18
    ss2.Col = -1: ss2.Col2 = -1
    ss2.ROW = 0: ss2.Row2 = 0
    ss2.FontSize = 9

    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "CG-System.INI", Me.Name)
    ss3.RowHeight(0) = 18
    ss3.Col = -1: ss3.Col2 = -1
    ss3.ROW = 0: ss3.Row2 = 0
    ss3.FontSize = 9

    Call Gp_Sp_ColGet(Proc_Sc("Sc4")("Spread"), "CG-System.INI", Me.Name)
    ss4.RowHeight(0) = 18
    ss4.Col = -1: ss4.Col2 = -1
    ss4.ROW = 0: ss4.Row2 = 0
    ss4.FontSize = 9

    Call Gp_Sp_ColGet(Proc_Sc("Sc5")("Spread"), "CG-System.INI", Me.Name)
    ss5.RowHeight(0) = 18
    ss5.Col = -1: ss5.Col2 = -1
    ss5.ROW = 0: ss5.Row2 = 0
    ss5.FontSize = 9
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc6")("Spread"), "CG-System.INI", Me.Name)
    ss6.RowHeight(0) = 18
    ss6.Col = -1: ss6.Col2 = -1
    ss6.ROW = 0: ss6.Row2 = 0
    ss6.FontSize = 9
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc7")("Spread"), "CG-System.INI", Me.Name)
    ss7.RowHeight(0) = 18
    ss7.Col = -1: ss7.Col2 = -1
    ss7.ROW = 0: ss7.Row2 = 0
    ss7.FontSize = 9
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc8")("Spread"), "CG-System.INI", Me.Name)
    ss8.RowHeight(0) = 18
    ss8.Col = -1: ss8.Col2 = -1
    ss8.ROW = 0: ss8.Row2 = 0
    ss8.FontSize = 9
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc4")("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc5")("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc6")("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc7")("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc8")("Spread"), "CG-System.INI", Me.Name)
    
    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
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
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set Sc3 = Nothing
    Set sc4 = Nothing
    Set sc5 = Nothing
    Set sc6 = Nothing
    Set sc7 = Nothing
    Set sc8 = Nothing
    
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    Call Gp_Ms_ControlLock(Mc2("pControl"), False)
    Option1.SetFocus
    
End Sub

Public Sub Form_Ref()
    
     Call Form_Ref1
    
End Sub

Public Sub Form_Ref1()

  If Gf_Ms_Outpara(M_CN1, Mc1) And Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, , , False) And _
                                   Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc1, , , False) And _
                                   Gf_Sp_Refer(M_CN1, Proc_Sc("Sc3"), Mc1, , , False) And _
                                   Gf_Sp_Refer(M_CN1, Proc_Sc("Sc4"), Mc1, , , False) And _
                                   Gf_Sp_Refer(M_CN1, Proc_Sc("Sc5"), Mc1, , , False) And _
                                   Gf_Sp_Refer(M_CN1, Proc_Sc("Sc6"), Mc1, , , False) And _
                                   Gf_Sp_Refer(M_CN1, Proc_Sc("Sc7"), Mc1, , , False) And _
                                   Gf_Sp_Refer(M_CN1, Proc_Sc("Sc8"), Mc1, , , False) Then
     Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
  End If
  
     txt_ONC_CNT1.Text = ss5.MaxRows + ss6.MaxRows
     txt_INF_CNT1.Text = ss1.MaxRows + ss2.MaxRows + ss3.MaxRows + ss4.MaxRows
  
End Sub

Public Sub Form_Ref2()

'  If Gf_Ms_Outpara(M_CN1, Mc2) Then
'     Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'  End If
  
End Sub

Private Sub Option1_Click()

  Timer1.Enabled = True
  Timer2.Enabled = True
  
End Sub

Private Sub Option2_Click()

    Timer1.Enabled = False
    Timer2.Enabled = False
  
End Sub

Private Sub Timer1_Timer()

    Call Form_Ref1
    
End Sub

Private Sub Timer2_Timer()

'   Dim link_mill As Long
'   link_mill = Val(txt_mill.Text)
'
'   Call Form_Ref2
'
'   If link_mill = Val(txt_mill.Text) Then
'       Line_1.Visible = False
'       Line1.Visible = True
'       Line2.Visible = True
'       Line3.BorderColor = &HFF00FF
'       Line4.Visible = True
''      Line_1.BorderColor = &HFF00FF
'   Else
'       Line_1.Visible = True
'       Line1.Visible = False
'       Line2.Visible = False
'       Line3.BorderColor = &H8000&
'       Line4.Visible = False
''      Line_1.BorderColor = &HC000&
'   End If
     
End Sub




