VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGB2010C 
   Caption         =   "加热炉装炉作业实绩查询及修改界面_CGB2010C"
   ClientHeight    =   9045
   ClientLeft      =   810
   ClientTop       =   1830
   ClientWidth     =   13680
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000011&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_tmpseq 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txt_EntCan 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   " "
      Top             =   720
      Visible         =   0   'False
      Width           =   210
   End
   Begin Threed.SSCheck SSCheck1 
      Height          =   255
      Left            =   435
      TabIndex        =   39
      Top             =   720
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   450
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   "装炉作业"
   End
   Begin TabDlg.SSTab tab1 
      Height          =   5115
      Left            =   165
      TabIndex        =   29
      Top             =   4050
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   9022
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabMaxWidth     =   3528
      BackColor       =   14737632
      TabCaption(0)   =   "装炉等待"
      TabPicture(0)   =   "CGB2010C.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSPpdt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ULabel1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ss1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt_search_slabno"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "装炉实绩"
      TabPicture(1)   =   "CGB2010C.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ss2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ULabel5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt_RstFormDate"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txt_RstToDate"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
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
         Left            =   2015
         MaxLength       =   10
         TabIndex        =   53
         ToolTipText     =   "回车检索"
         Top             =   405
         Width           =   1365
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3855
         Left            =   -74910
         TabIndex        =   30
         Top             =   390
         Width           =   14700
         _Version        =   393216
         _ExtentX        =   25929
         _ExtentY        =   6800
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
         MaxCols         =   10
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2010C.frx":0038
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   3855
         Left            =   -74910
         TabIndex        =   31
         Top             =   390
         Width           =   14700
         _Version        =   393216
         _ExtentX        =   25929
         _ExtentY        =   6800
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
         MaxCols         =   10
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2010C.frx":1B97
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   3855
         Left            =   -74910
         TabIndex        =   32
         Top             =   390
         Width           =   14700
         _Version        =   393216
         _ExtentX        =   25929
         _ExtentY        =   6800
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
         MaxCols         =   10
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2010C.frx":36F6
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4065
         Left            =   -74910
         TabIndex        =   33
         Top             =   780
         Width           =   14790
         _Version        =   393216
         _ExtentX        =   26088
         _ExtentY        =   7170
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
         MaxCols         =   16
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2010C.frx":5255
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   4260
         Left            =   120
         TabIndex        =   34
         Top             =   765
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   7514
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
         MaxCols         =   25
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2010C.frx":5A44
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   -74895
         Top             =   420
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "装炉时间"
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
      Begin CSTextLibCtl.sitxEdit txt_RstFormDate 
         Height          =   315
         Left            =   -73500
         TabIndex        =   37
         Tag             =   "装炉时间"
         Top             =   420
         Width           =   1800
         _Version        =   262145
         _ExtentX        =   3175
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __-__-__"
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
         Mask            =   "____-__-__ __:__"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin CSTextLibCtl.sitxEdit txt_RstToDate 
         Height          =   315
         Left            =   -71670
         TabIndex        =   38
         Tag             =   "装炉时间"
         Top             =   420
         Width           =   1770
         _Version        =   262145
         _ExtentX        =   3122
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __-__-__"
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
         Mask            =   "____-__-__ __:__"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   120
         Top             =   405
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         Caption         =   "检索板坯号"
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
      Begin Threed.SSPanel SSPpdt 
         Height          =   315
         Left            =   5130
         TabIndex        =   57
         Top             =   390
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   255
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "当月以前交货订单"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSCheck sc3 
      Height          =   255
      Left            =   8550
      TabIndex        =   8
      Top             =   720
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   450
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "缺号板坯"
   End
   Begin InDate.ULabel ULabel29 
      Height          =   315
      Left            =   15480
      Top             =   3660
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "尾部宽度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel28 
      Height          =   315
      Left            =   15480
      Top             =   3150
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "头部宽度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   15480
      Top             =   2655
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "厚度"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CSTextLibCtl.sidbEdit SDB_TAIL_SLAB_WID 
      Height          =   315
      Left            =   17010
      TabIndex        =   9
      Top             =   3615
      Visible         =   0   'False
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
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
      NumDecDigits    =   2
      NumIntDigits    =   4
      ShowZero        =   0   'False
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_HEAD_SLAB_WID 
      Height          =   315
      Left            =   17010
      TabIndex        =   10
      Top             =   3150
      Visible         =   0   'False
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
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
      NumDecDigits    =   2
      NumIntDigits    =   4
      ShowZero        =   0   'False
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_SLAB_REJ_THK 
      Height          =   315
      Left            =   17010
      TabIndex        =   11
      Top             =   2655
      Visible         =   0   'False
      Width           =   1215
      _Version        =   262145
      _ExtentX        =   2143
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
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
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
      NumDecDigits    =   2
      NumIntDigits    =   4
      ShowZero        =   0   'False
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   720
      Left            =   165
      TabIndex        =   3
      Top             =   105
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   1270
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_RollingSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9490
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   " "
         Top             =   165
         Width           =   2955
      End
      Begin VB.TextBox TXT_SLAB_SIZE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4995
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   " "
         Top             =   165
         Width           =   2805
      End
      Begin VB.TextBox txt_SlabNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   35
         Top             =   150
         Width           =   1635
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   285
         Top             =   165
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "板坯号"
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
         Left            =   3600
         Top             =   165
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "板坯规格"
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
         Left            =   8100
         Top             =   165
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "轧制规格"
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
   Begin Threed.SSFrame sf1 
      Height          =   3225
      Left            =   165
      TabIndex        =   15
      Top             =   795
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   5689
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.OptionButton opt_rhf 
         BackColor       =   &H00E0E0E0&
         Caption         =   "四号炉"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   3
         Left            =   4980
         TabIndex        =   52
         Top             =   870
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   2550
         TabIndex        =   47
         Top             =   150
         Width           =   3555
         Begin Threed.SSOption opt_EntCan 
            Height          =   435
            Index           =   0
            Left            =   210
            TabIndex        =   48
            Top             =   0
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   767
            _Version        =   196609
            ForeColor       =   0
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
            Caption         =   "1:装炉"
         End
         Begin Threed.SSOption opt_EntCan 
            Height          =   435
            Index           =   1
            Left            =   1290
            TabIndex        =   49
            Top             =   0
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   767
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
            Caption         =   "2:装炉取消"
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4980
         TabIndex        =   43
         Top             =   1290
         Width           =   3045
         Begin Threed.SSOption opt_RHF_ROW 
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   44
            Top             =   0
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   196609
            ForeColor       =   0
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
            Caption         =   "0-单排"
         End
         Begin Threed.SSOption opt_RHF_ROW 
            Height          =   285
            Index           =   1
            Left            =   1132
            TabIndex        =   45
            Top             =   0
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
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
            Caption         =   "1-左"
         End
         Begin Threed.SSOption opt_RHF_ROW 
            Height          =   285
            Index           =   2
            Left            =   2115
            TabIndex        =   46
            Top             =   0
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
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
            Caption         =   "2-右"
         End
         Begin Threed.SSOption opt_RHF_ROW 
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   54
            Top             =   300
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   196609
            ForeColor       =   0
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
            Caption         =   "3-中"
         End
         Begin Threed.SSOption opt_RHF_ROW 
            Height          =   285
            Index           =   4
            Left            =   1132
            TabIndex        =   55
            Top             =   300
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
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
            Caption         =   "4-长左"
         End
         Begin Threed.SSOption opt_RHF_ROW 
            Height          =   285
            Index           =   5
            Left            =   2115
            TabIndex        =   56
            Top             =   300
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   503
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
            Caption         =   "5-长右"
         End
      End
      Begin VB.TextBox txt_func 
         Alignment       =   2  'Center
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
         Left            =   6450
         Locked          =   -1  'True
         TabIndex        =   28
         Tag             =   "炉座号"
         Text            =   " "
         Top             =   840
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.OptionButton opt_rhf 
         BackColor       =   &H00E0E0E0&
         Caption         =   "一号炉"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   1680
         MaskColor       =   &H8000000F&
         TabIndex        =   26
         Top             =   870
         Width           =   1095
      End
      Begin VB.OptionButton opt_rhf 
         BackColor       =   &H00E0E0E0&
         Caption         =   "二号炉"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   2780
         TabIndex        =   25
         Top             =   870
         Width           =   1095
      End
      Begin VB.OptionButton opt_rhf 
         BackColor       =   &H00E0E0E0&
         Caption         =   "三号炉"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   2
         Left            =   3880
         TabIndex        =   24
         Top             =   870
         Width           =   1095
      End
      Begin VB.TextBox TXT_SHIFT 
         Alignment       =   2  'Center
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
         Left            =   270
         MaxLength       =   1
         TabIndex        =   20
         Top             =   2820
         Width           =   705
      End
      Begin VB.TextBox TXT_GROUP 
         Alignment       =   2  'Center
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
         Left            =   975
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   19
         Top             =   2820
         Width           =   705
      End
      Begin VB.TextBox TXT_EMP 
         Alignment       =   2  'Center
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   18
         Top             =   2820
         Width           =   1335
      End
      Begin VB.TextBox TXT_RHF_CH_ROW 
         Alignment       =   2  'Center
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
         Left            =   4005
         TabIndex        =   17
         Tag             =   "布料方式"
         Text            =   " "
         Top             =   1305
         Width           =   915
      End
      Begin VB.TextBox txt_Status 
         Alignment       =   2  'Center
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
         Left            =   1645
         Locked          =   -1  'True
         TabIndex        =   16
         Tag             =   "装炉/取消"
         Text            =   " "
         Top             =   195
         Width           =   855
      End
      Begin CSTextLibCtl.sitxEdit TXT_RHF_CH_TIME 
         Height          =   315
         Left            =   1645
         TabIndex        =   21
         Tag             =   "装炉时间"
         Top             =   2070
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
      Begin CSTextLibCtl.sidbEdit SDB_CHARGE_TEMP 
         Height          =   315
         Left            =   1645
         TabIndex        =   1
         Tag             =   "温度"
         Top             =   1305
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
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
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
         ShowZero        =   0   'False
         MaxValue        =   9999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   270
         Top             =   2070
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "装炉时间"
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
         Left            =   2640
         Top             =   1305
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "布料方式"
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
         Left            =   270
         Top             =   1305
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "温度(℃)"
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
         Left            =   270
         Top             =   195
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         Caption         =   "装炉/取消"
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
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   270
         Top             =   2490
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "班次"
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
         Left            =   975
         Top             =   2490
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "班别"
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
         Left            =   1680
         Top             =   2490
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "作业人员"
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
      Begin CSTextLibCtl.sitxEdit TXT_RHF_CH_TIME_UPD 
         Height          =   315
         Left            =   5565
         TabIndex        =   22
         Tag             =   "装炉时间"
         Top             =   2070
         Visible         =   0   'False
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
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   4020
         Top             =   2070
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         Caption         =   "装炉时间修正"
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
         Left            =   270
         Top             =   840
         Width           =   1350
         _ExtentX        =   2381
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
      Begin Threed.SSPanel SSP4 
         Height          =   315
         Left            =   3390
         TabIndex        =   58
         Top             =   2790
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "重点订单"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSFrame sf3 
      Height          =   3225
      Left            =   8295
      TabIndex        =   4
      Top             =   795
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5689
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_COMFRM 
         Height          =   330
         Left            =   1800
         TabIndex        =   27
         Text            =   " "
         Top             =   1740
         Width           =   555
      End
      Begin VB.TextBox TXT_REASON_NAME 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox TXT_REASON_CD 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1020
         Width           =   645
      End
      Begin VB.TextBox TXT_REJ_EMP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   7
         Top             =   2820
         Width           =   1335
      End
      Begin VB.TextBox TXT_REJ_GROUP 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   975
         MaxLength       =   1
         TabIndex        =   6
         Top             =   2820
         Width           =   705
      End
      Begin VB.TextBox TXT_REJ_SHIFT 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2820
         Width           =   705
      End
      Begin InDate.ULabel ULabel26 
         Height          =   315
         Left            =   270
         Top             =   315
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "缺号时间"
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
         Left            =   270
         Top             =   1020
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "缺号代码"
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
      Begin CSTextLibCtl.sitxEdit TXT_REJ_OCCR_TIME 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Tag             =   "缺号时"
         Top             =   315
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
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   270
         Top             =   2490
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "班次"
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
      Begin InDate.ULabel ULabel41 
         Height          =   315
         Left            =   975
         Top             =   2490
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         Caption         =   "班别"
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
         Left            =   1680
         Top             =   2490
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "作业人员"
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
      Begin InDate.ULabel ULabel3 
         Height          =   330
         Left            =   270
         Top             =   1740
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         Caption         =   "缺号板坯确定"
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
      Begin Threed.SSOption opt_ORDER 
         Height          =   285
         Index           =   0
         Left            =   2430
         TabIndex        =   41
         Top             =   1770
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   196609
         ForeColor       =   0
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
         Caption         =   "1:订单"
      End
      Begin Threed.SSOption opt_ORDER 
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   42
         Top             =   1770
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
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
         Caption         =   "2:余材"
      End
   End
   Begin VB.Label Label49 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13995
      TabIndex        =   14
      Top             =   3645
      Width           =   255
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13995
      TabIndex        =   13
      Top             =   2655
      Width           =   255
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13995
      TabIndex        =   12
      Top             =   3180
      Width           =   255
   End
End
Attribute VB_Name = "CGB2010C"
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
'-- Program Name      加热炉装炉实绩查询及修改界面
'-- Program ID        CGB2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          SHIN.C.S
'-- Coder             SHIN.C.S
'-- Date              2007.7.23
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public FormType As String            'Form Type
Public Toolbar_St As String          'Active Form ToolBar Setting
Public sAuthority As String          'Active Form Authority Setting
Public sDateTime As String           'Active Form Time Setting
Public sQuery_Rt As String

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

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

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
Dim Mc3 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_SLAB_NO = 1
Const SS1_SLAB_SIZE = 7
Const SS1_MILL_SIZE = 10
Const SS1_DEL_TO_DATE = 18
Const SS1_URGNT_FL = 20
Const SS1_IMP_CONT = 21
Const SS2_IMP_CONT = 16
Const SS2_SLAB_NO = 1
 
Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(txt_SlabNo, "p", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(TXT_SLAB_SIZE, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_RollingSize, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
             Call Gp_Ms_Collection(txt_Status, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
               Call Gp_Ms_Collection(txt_func, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(SDB_CHARGE_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_RHF_CH_ROW, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(TXT_RHF_CH_TIME, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(TXT_RHF_CH_TIME_UPD, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(txt_Shift, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(TXT_GROUP, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
                Call Gp_Ms_Collection(TXT_EMP, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          
    'MASTER Collection
     Mc1.Add Item:="CGB2010C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="CGB2010C.P_REFER3", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
          
          
        Call Gp_Ms_Collection(txt_SlabNo, "p", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
 Call Gp_Ms_Collection(TXT_REJ_OCCR_TIME, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(TXT_REASON_CD, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(TXT_COMFRM, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(txt_Shift, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_GROUP, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(TXT_EMP, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                 
     Mc2.Add Item:="CGB2010C.P_MODIFY3", Key:="P-M"
     Mc2.Add Item:="CGB2010C.P_REFER4", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
     
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
 
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGB2010C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
          
    Call Gp_Ms_Collection(txt_RstFormDate, "p", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
      Call Gp_Ms_Collection(txt_RstToDate, "p", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    
    'MASTER Collection
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"
   
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGB2010C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxRows, Key:="Last"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    ss1.ROW = 0
    ss1.Col = 0
    ss1.Text = "◎"
    ss2.ROW = 0
    ss2.Col = 0
    ss2.Text = ""
     
End Sub


Private Sub opt_EntCan_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       txt_Status.Text = "1"
       opt_EntCan(0).ForeColor = &HFF&       'Red color
       opt_EntCan(1).ForeColor = &H80000012    'Black color
    ElseIf Index = 1 Then
       txt_Status.Text = "2"
       opt_EntCan(1).ForeColor = &HFF&       'Red color
       opt_EntCan(0).ForeColor = &H80000012    'Black color
    End If
End Sub

Private Sub opt_ORDER_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       TXT_COMFRM.Text = "1"
       opt_ORDER(0).ForeColor = &HFF&       'Red color
       opt_ORDER(1).ForeColor = &H80000012    'Black color
    ElseIf Index = 1 Then
       TXT_COMFRM.Text = "2"
       opt_ORDER(1).ForeColor = &HFF&       'Red color
       opt_ORDER(0).ForeColor = &H80000012    'Black color
    End If
End Sub

Private Sub opt_rhf_Click(Index As Integer)
    If Index = 0 Then
        txt_func.Text = "1"
        opt_rhf(0).ForeColor = &HFF&
        opt_rhf(1).ForeColor = &H80000011
        opt_rhf(2).ForeColor = &H80000011
        opt_rhf(3).ForeColor = &H80000011
    ElseIf Index = 1 Then
        txt_func.Text = "2"
        opt_rhf(1).ForeColor = &HFF&
        opt_rhf(0).ForeColor = &H80000011
        opt_rhf(2).ForeColor = &H80000011
        opt_rhf(3).ForeColor = &H80000011
    ElseIf Index = 2 Then
        txt_func.Text = "3"
        opt_rhf(2).ForeColor = &HFF&
        opt_rhf(0).ForeColor = &H80000011
        opt_rhf(1).ForeColor = &H80000011
        opt_rhf(3).ForeColor = &H80000011
   ElseIf Index = 3 Then
        txt_func.Text = "4"
        opt_rhf(3).ForeColor = &HFF&
        opt_rhf(0).ForeColor = &H80000011
        opt_rhf(1).ForeColor = &H80000011
        opt_rhf(2).ForeColor = &H80000011
    End If
End Sub

Private Sub opt_RHF_ROW_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       TXT_RHF_CH_ROW = "0"
       opt_RHF_ROW(0).ForeColor = &HFF&       'Red color
       opt_RHF_ROW(1).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(2).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(3).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(4).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(5).ForeColor = &H80000012    'Black color
    ElseIf Index = 1 Then
       TXT_RHF_CH_ROW = "1"
       opt_RHF_ROW(1).ForeColor = &HFF&       'Red color
       opt_RHF_ROW(0).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(2).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(3).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(4).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(5).ForeColor = &H80000012    'Black color
    ElseIf Index = 2 Then
       TXT_RHF_CH_ROW = "2"
       opt_RHF_ROW(2).ForeColor = &HFF&       'Red color
       opt_RHF_ROW(0).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(1).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(3).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(4).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(5).ForeColor = &H80000012    'Black color
    ElseIf Index = 3 Then
       TXT_RHF_CH_ROW = "3"
       opt_RHF_ROW(3).ForeColor = &HFF&       'Red color
       opt_RHF_ROW(0).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(1).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(2).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(4).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(5).ForeColor = &H80000012    'Black color
    ElseIf Index = 4 Then
       TXT_RHF_CH_ROW = "4"
       opt_RHF_ROW(4).ForeColor = &HFF&       'Red color
       opt_RHF_ROW(0).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(1).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(2).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(3).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(5).ForeColor = &H80000012    'Black color
    ElseIf Index = 5 Then
       TXT_RHF_CH_ROW = "5"
       opt_RHF_ROW(5).ForeColor = &HFF&       'Red color
       opt_RHF_ROW(0).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(1).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(2).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(3).ForeColor = &H80000012    'Black color
       opt_RHF_ROW(4).ForeColor = &H80000012    'Black color
    End If
End Sub


Private Sub sc3_Click(Value As Integer)

    If sc3.Value = ssCBUnchecked Then
       If SSCheck1.Value = ssCBUnchecked Then
          sc3.Value = ssCBChecked
          sc3.ForeColor = &HFF&
       End If
        Exit Sub
    End If

    txt_EntCan = "2"
    
    'Cancel Data Enbaled is False
    sc3.Value = ssCBChecked
    SSCheck1.Value = ssCBUnchecked
    TXT_REJ_OCCR_TIME.Enabled = True
    TXT_REASON_CD.Enabled = True
    TXT_REASON_NAME.Enabled = True
    TXT_COMFRM.Enabled = True


    TXT_REJ_SHIFT.Enabled = True
    TXT_REJ_GROUP.Enabled = True
    TXT_REJ_EMP.Enabled = True
    
    'Slab Entriy Data Enbaled is True
    txt_Status.Enabled = False


    
    SDB_CHARGE_TEMP.Enabled = False
    TXT_RHF_CH_ROW.Enabled = False

    TXT_RHF_CH_TIME.Enabled = False
    TXT_RHF_CH_TIME_UPD.Enabled = False
    txt_Shift.Enabled = False
    TXT_GROUP.Enabled = False
    TXT_EMP.Enabled = False
    
    opt_ORDER(0).Value = True
    TXT_REJ_SHIFT = Gf_ShiftSet3(M_CN1)
    TXT_REJ_GROUP = Gf_GroupSet(M_CN1, Trim(TXT_REJ_SHIFT), Gf_DTSet(M_CN1, , "X"))
    TXT_REJ_EMP = sUserID
    Call TXT_REJ_OCCR_TIME_DblClick

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    If ROW = 0 And Col = 1 Then
        Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    End If

    If ROW > 0 Then
        If SSCheck1.Value = -1 And opt_EntCan(0).Value = True Then
            ss1.ROW = ROW
            ss1.Col = SS1_SLAB_NO:            txt_SlabNo.Text = ss1.Text
            ss1.Col = SS1_SLAB_SIZE:          TXT_SLAB_SIZE.Text = ss1.Text
            ss1.Col = SS1_MILL_SIZE:          txt_RollingSize.Text = ss1.Text
            txt_tmpseq.Text = ROW
                
        ElseIf sc3.Value = -1 Then
        End If
        
    End If
    
    ss1.ROW = 0
    ss1.Col = 0
    ss1.Text = "◎"
    ss2.ROW = 0
    ss2.Col = 0
    ss2.Text = ""
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If SSCheck1.Value = -1 And opt_EntCan(0).Value = True Then
            ss1.ROW = ROW
            ss1.Col = SS1_SLAB_NO:            txt_SlabNo.Text = ss1.Text
            ss1.Col = SS1_SLAB_SIZE:          TXT_SLAB_SIZE.Text = ss1.Text
            ss1.Col = SS1_MILL_SIZE:          txt_RollingSize.Text = ss1.Text
            
        ElseIf sc3.Value = -1 Then
        End If
        
    End If
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)

    If ROW = 0 And Col = 1 Then
        Call Gp_Sp_Sort(ss2, Col, ROW)
    End If

    If ROW > 0 Then
        If SSCheck1.Value = -1 And opt_EntCan(1).Value = True Then
            ss2.ROW = ROW
            ss2.Col = 1
            txt_SlabNo.Text = ss2.Text
        ElseIf sc3.Value = -1 Then
            TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        
        If Trim(txt_SlabNo.Text) <> "" Then
            Call Gf_Ms_Refer(M_CN1, Mc1, , , True)
        End If
        
    End If
    
    If opt_EntCan(0).Value = True Then
       txt_Status.Text = "1"
    ElseIf opt_EntCan(1).Value = True Then
       txt_Status.Text = "2"
    End If
    
    ss2.ROW = 0
    ss2.Col = 0
    ss2.Text = "◎"
    ss1.ROW = 0
    ss1.Col = 0
    ss1.Text = ""
    
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If SSCheck1.Value = -1 And opt_EntCan(1).Value = True Then
            ss2.ROW = ROW
            ss2.Col = SS1_SLAB_NO:            txt_SlabNo.Text = ss2.Text
        ElseIf sc3.Value = -1 Then
            TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        
        If Trim(txt_SlabNo.Text) <> "" Then
            Call Gf_Ms_Refer(M_CN1, Mc1, , , True)
        End If
    End If
    
    If opt_EntCan(0).Value = True Then
       txt_Status.Text = "1"
    ElseIf opt_EntCan(1).Value = True Then
       txt_Status.Text = "2"
    End If

End Sub

Private Sub SSCheck1_Click(Value As Integer)

    If SSCheck1.Value = ssCBUnchecked Then
       If sc3.Value = ssCBUnchecked Then
          SSCheck1.Value = ssCBChecked
          SSCheck1.ForeColor = &HFF&
       End If
       Exit Sub
    End If
            
    txt_EntCan = "1"
    
    'Cancel Data Enbaled is False
    sc3.Value = ssCBUnchecked
    TXT_REJ_OCCR_TIME.Enabled = False
    TXT_REASON_CD.Enabled = False
    TXT_REASON_NAME.Enabled = False
    TXT_COMFRM.Enabled = False
    TXT_REJ_SHIFT.Enabled = False
    TXT_REJ_GROUP.Enabled = False
    TXT_REJ_EMP.Enabled = False
    
    'Slab Entriy Data Enbaled is True
    txt_Status.Enabled = True

    
    SDB_CHARGE_TEMP.Enabled = True
    TXT_RHF_CH_ROW.Enabled = True
    
    opt_EntCan(0).Value = True
    opt_RHF_ROW(0).Value = True

    TXT_RHF_CH_TIME.Enabled = True
    TXT_RHF_CH_TIME_UPD.Enabled = True
    txt_Shift.Enabled = True
    TXT_GROUP.Enabled = True
    TXT_EMP.Enabled = True
    
    'Call TXT_RHF_CH_TIME_DblClick
    txt_Shift = Gf_ShiftSet3(M_CN1)
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
    TXT_EMP = sUserID

End Sub


Private Sub tab1_Click(PreviousTab As Integer)
    If tab1.Tab = "1" Then
        txt_Shift = Gf_ShiftSet3(M_CN1)
        If txt_Shift = "1" Then
            txt_RstFormDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "000001"
            txt_RstToDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "081459"
        ElseIf txt_Shift = "2" Then
            txt_RstFormDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "081500"
            txt_RstToDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "155959"
        ElseIf txt_Shift = "3" Then
            txt_RstFormDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "160000"
            txt_RstToDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "235959"
        End If
    End If
    
    Call Form_Ref
End Sub

Private Sub TXT_EMP_DblClick()
    Call TXT_EMP_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_EMP_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        TXT_EMP.Text = ""
        DD.rControl.Add Item:=TXT_EMP

        Call Gf_EmpID_DD(M_CN1, vbKeyF4, "1ZB")

        Exit Sub
    End If
End Sub

Private Sub txt_func_Change()
    If txt_func.Text = "1" Then
        opt_rhf(0).Value = True
        opt_rhf(0).ForeColor = &HFF&
        opt_rhf(1).ForeColor = &H80000011
        opt_rhf(2).ForeColor = &H80000011
        opt_rhf(3).ForeColor = &H80000011
    ElseIf txt_func.Text = "2" Then
        opt_rhf(1).Value = True
        opt_rhf(1).ForeColor = &HFF&
        opt_rhf(0).ForeColor = &H80000011
        opt_rhf(2).ForeColor = &H80000011
        opt_rhf(3).ForeColor = &H80000011
    ElseIf txt_func.Text = "3" Then
        opt_rhf(2).Value = True
        opt_rhf(2).ForeColor = &HFF&
        opt_rhf(0).ForeColor = &H80000011
        opt_rhf(1).ForeColor = &H80000011
        opt_rhf(3).ForeColor = &H80000011
    ElseIf txt_func.Text = "4" Then
        opt_rhf(3).Value = True
        opt_rhf(3).ForeColor = &HFF&
        opt_rhf(0).ForeColor = &H80000011
        opt_rhf(1).ForeColor = &H80000011
        opt_rhf(2).ForeColor = &H80000011
    End If
End Sub

Private Sub TXT_REASON_CD_Change()
    If TXT_REASON_CD = "" Then
       TXT_REASON_NAME = ""
    End If
End Sub

Private Sub TXT_REASON_CD_DblClick()
    Call TXT_REASON_CD_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_REASON_CD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "G0001"
        DD.rControl.Add Item:=TXT_REASON_CD
        DD.rControl.Add Item:=TXT_REASON_NAME
    
        DD.nameType = "1"
    
        Call Gf_Common_DD(M_CN1, KeyCode)
    
        Exit Sub
    
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
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_ControlLock(Mc2("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))

    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "CG-System.INI", Me.Name)

    
    txt_Shift = Gf_ShiftSet3(M_CN1)
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(txt_Shift), Gf_DTSet(M_CN1, , "X"))
    TXT_EMP = sUserID
    
    Call SSCheck1_Click(1)
    opt_rhf(0).Value = True
    opt_EntCan(0).Value = True
    opt_RHF_ROW(0).Value = True
    
    tab1.Tab = 0
    Call Form_Ref
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "CG-System.INI", Me.Name)

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
    
    Set pControl3 = Nothing
    Set nControl3 = Nothing
    Set iControl3 = Nothing
    Set rControl3 = Nothing
    Set cControl3 = Nothing
    Set aControl3 = Nothing
    Set lControl3 = Nothing
    Set mControl3 = Nothing
    
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
    Set Mc3 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing

    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()
ss1.ROW = 0: ss1.Col = 0
If ss1.Text = "◎" Then
    Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End If
ss2.ROW = 0: ss2.Col = 0
If ss2.Text = "◎" Then
    Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End If

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

    Dim sMesg As String
    Dim i, j As Integer

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)

    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    Call Gp_Ms_ControlLock(Mc2("pControl"), False)
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = False                'Excel
    End With
       
    For i = 0 To 3
        opt_rhf(i).Value = False
        opt_rhf(i).ForeColor = &H80000011
    Next i

    For j = 0 To 2
        opt_RHF_ROW(j).Value = False
        opt_RHF_ROW(j).ForeColor = &H80000012
    Next j

    opt_EntCan(0).Value = False
    opt_EntCan(0).ForeColor = &H80000012
    opt_EntCan(1).Value = False
    opt_EntCan(1).ForeColor = &H80000012
    
    TXT_EMP.Text = sUserID
        
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     End If

End Sub

Public Sub Form_Ref()

    Dim iRow As Integer
    Dim iCol As Integer
    Dim sCurDate As String
    Dim sDel_To_Date As String
    Dim sUrgnt_Fl As String
    Dim simpcont As String
    Dim simpcont1 As String
    
    sCurDate = Format(Now, "YYYYMM")

    If tab1.Tab = 0 Then
        Call Gf_Sp_Refer(M_CN1, sc1, Mc1)
        ss1.OperationMode = OperationModeNormal
        opt_EntCan(0).Value = True
        ss1.Col = 1
        ss1.ROW = IIf(Val(txt_tmpseq.Text) = 0, 1, Val(txt_tmpseq.Text))
        ss1.SetActiveCell 1, ss1.ROW
        If ss1.Text <> "" Then
            Call ss1_DblClick(1, ss1.ROW)
        End If
        
        With ss1
              For iRow = 1 To .MaxRows
                 .ROW = iRow:             .Col = SS1_DEL_TO_DATE
                 .Col = SS1_URGNT_FL:     sUrgnt_Fl = Trim(.Text)
                  sDel_To_Date = Mid(.Value, 1, 6)
                  If sDel_To_Date < sCurDate Then
                       Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, &HFF&)
                  End If
                   '紧急订单绿色显示 add by liqian 2012-08-15
                  If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, &HC000&)
                  End If
                  
                  .ROW = iRow:
                  .Col = SS1_IMP_CONT:   simpcont = Trim(.Text)
                   If simpcont = "Y" Then
                       Call Gp_Sp_BlockColor(ss1, SS1_SLAB_NO, SS1_SLAB_NO, iRow, iRow, SSP4.BackColor)
                       Call Gp_Sp_BlockColor(ss1, SS1_IMP_CONT, SS1_IMP_CONT, iRow, iRow, SSP4.BackColor)
                   End If
              Next iRow
        End With
        
        With ss2
              For iRow = 1 To .MaxRows
                .ROW = iRow:
                .Col = SS2_IMP_CONT:   simpcont1 = Trim(.Text)
                 If simpcont = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, SS2_SLAB_NO, SS2_SLAB_NO, iRow, iRow, SSP4.BackColor)
                     Call Gp_Sp_BlockColor(ss1, SS2_IMP_CONT, SS2_IMP_CONT, iRow, iRow, SSP4.BackColor)
                 End If
              Next iRow
        End With
        
    ElseIf tab1.Tab = 1 Then
        Call Gf_Sp_Refer(M_CN1, sc2, Mc3, Mc3("nControl"), Mc3("mControl"), False)
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

Public Sub Form_Pro()
Dim sMesg As String
Dim sLoc As String
Dim Temp_no As String
Dim i, j As Integer
    
    If Not Gp_DateCheck(TXT_RHF_CH_TIME) Then
            sMesg = " 请正确输入装炉时间 ！"
            Call Gp_MsgBoxDisplay(sMesg)
            Exit Sub
    End If
    
    If txt_EntCan.Text = "1" Then
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) = False Then Exit Sub
    ElseIf txt_EntCan.Text = "2" Then
        If Gf_Ms_Process(M_CN1, Mc2, sAuthority) = False Then Exit Sub
    End If
    
    TXT_RHF_CH_TIME = ""
    TXT_REJ_OCCR_TIME = ""
    
    If tab1.Tab = 0 Then
        Call Gf_Sp_Refer(M_CN1, sc1, Nothing, Nothing, Nothing)
        Call ss1_DblClick(1, 1)
    ElseIf tab1.Tab = 1 Then
        Call Gf_Sp_Refer(M_CN1, sc2, Mc3, Mc3("nControl"), Mc3("mControl"))
        Call ss2_DblClick(1, 1)
    End If
    
    ''added by guoli at 20100703
    If Trim(txt_func.Text) = "" Then
       For i = 0 To 3
           If opt_rhf(i).Value = True Then
              txt_func.Text = i + 1
           End If
       Next i
    End If
    
    If Trim(TXT_RHF_CH_ROW.Text) = "" Then
       For j = 0 To 2
           If opt_RHF_ROW(j).Value = True Then
              TXT_RHF_CH_ROW.Text = j
           End If
       Next j
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub


Private Sub TXT_REJ_EMP_DblClick()
    Call TXT_REJ_EMP_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_REJ_EMP_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        TXT_REJ_EMP.Text = ""
        DD.rControl.Add Item:=TXT_REJ_EMP

        Call Gf_EmpID_DD(M_CN1, vbKeyF4, "1ZB")

        Exit Sub
    End If
End Sub

Private Sub TXT_RHF_CH_ROW_Change()
    If TXT_RHF_CH_ROW.Text = "0" Then
        opt_RHF_ROW(0).Value = True
        opt_RHF_ROW(0).ForeColor = &HFF&
        opt_RHF_ROW(1).ForeColor = &H80000011
        opt_RHF_ROW(2).ForeColor = &H80000011
    ElseIf TXT_RHF_CH_ROW.Text = "1" Then
        opt_RHF_ROW(1).Value = True
        opt_RHF_ROW(1).ForeColor = &HFF&
        opt_RHF_ROW(0).ForeColor = &H80000011
        opt_RHF_ROW(2).ForeColor = &H80000011
    ElseIf TXT_RHF_CH_ROW.Text = "2" Then
        opt_RHF_ROW(2).Value = True
        opt_RHF_ROW(2).ForeColor = &HFF&
        opt_RHF_ROW(0).ForeColor = &H80000011
        opt_RHF_ROW(1).ForeColor = &H80000011
    End If
End Sub

Private Sub TXT_RHF_CH_TIME_DblClick()

    TXT_RHF_CH_TIME.RawData = Gf_DTSet(M_CN1, , "X")
   
End Sub

Private Sub TXT_REJ_OCCR_TIME_DblClick()

     TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")

End Sub
Private Sub TXT_RHF_CH_TIME_UPD_DblClick()

    TXT_RHF_CH_TIME_UPD.RawData = Gf_DTSet(M_CN1, , "X")
End Sub

Private Sub txt_RstToDate_DblClick()
    txt_RstFormDate.RawData = Gf_DTSet(M_CN1, , "X")
    txt_RstToDate.RawData = Gf_DTSet(M_CN1, , "X")
End Sub
Private Sub txt_search_slabno_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
If KeyCode = 13 Then
       For i = 1 To ss1.MaxRows
           ss1.ROW = i
           ss1.Col = 1
           If ss1.Text = Trim(txt_search_slabno.Text) Then
              Call ss1.SetActiveCell(1, i)
              Exit For
           End If
       Next i
    End If
End Sub
