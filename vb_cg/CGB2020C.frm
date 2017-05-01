VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGB2020C 
   Caption         =   "加热炉出炉作业实绩查询及修改界面_CGB2020C"
   ClientHeight    =   9240
   ClientLeft      =   1680
   ClientTop       =   1695
   ClientWidth     =   15120
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_EntCan 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   " "
      Top             =   690
      Visible         =   0   'False
      Width           =   210
   End
   Begin Threed.SSCheck sChk2 
      Height          =   315
      Left            =   8550
      TabIndex        =   33
      Top             =   660
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   "缺号板坯"
   End
   Begin TabDlg.SSTab tab1 
      Height          =   4335
      Left            =   165
      TabIndex        =   28
      Top             =   4830
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabMaxWidth     =   3528
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
      TabCaption(0)   =   "一号加热炉"
      TabPicture(0)   =   "CGB2020C.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ss1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "二号加热炉"
      TabPicture(1)   =   "CGB2020C.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ss2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "三号加热炉"
      TabPicture(2)   =   "CGB2020C.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ss3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "四号加热炉"
      TabPicture(3)   =   "CGB2020C.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ss5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "加热炉作业实绩"
      TabPicture(4)   =   "CGB2020C.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ss4"
      Tab(4).Control(1)=   "ULabel5"
      Tab(4).Control(2)=   "txt_RstFormDate"
      Tab(4).Control(3)=   "txt_RstToDate"
      Tab(4).ControlCount=   4
      Begin FPSpread.vaSpread ss1 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   29
         Top             =   390
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   6800
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2020C.frx":008C
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3855
         Left            =   120
         TabIndex        =   38
         Top             =   390
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   6800
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
         MaxCols         =   34
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2020C.frx":0B2F
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   39
         Top             =   390
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   6800
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
         MaxCols         =   34
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2020C.frx":182F
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   3525
         Left            =   -74880
         TabIndex        =   45
         Top             =   720
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   6218
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
         MaxCols         =   26
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2020C.frx":2524
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   -74880
         Top             =   360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "出炉时间"
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
         Left            =   -73470
         TabIndex        =   46
         Tag             =   "装炉时间"
         Top             =   360
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
         Left            =   -71640
         TabIndex        =   47
         Tag             =   "装炉时间"
         Top             =   360
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
      Begin FPSpread.vaSpread ss5 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   48
         Top             =   390
         Width           =   14760
         _Version        =   393216
         _ExtentX        =   26035
         _ExtentY        =   6800
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
         MaxCols         =   34
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGB2020C.frx":2FCF
      End
   End
   Begin Threed.SSCheck sChk1 
      Height          =   315
      Left            =   450
      TabIndex        =   0
      Top             =   660
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   "出炉作业"
   End
   Begin Threed.SSFrame sf2 
      Height          =   4005
      Left            =   165
      TabIndex        =   25
      Top             =   795
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   7064
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   2535
         TabIndex        =   42
         Top             =   120
         Width           =   2670
         Begin Threed.SSOption opt_EntCan 
            Height          =   435
            Index           =   0
            Left            =   90
            TabIndex        =   43
            Top             =   60
            Width           =   975
            _ExtentX        =   1720
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
            Caption         =   "1:出炉"
         End
         Begin Threed.SSOption opt_EntCan 
            Height          =   435
            Index           =   1
            Left            =   1245
            TabIndex        =   44
            Top             =   60
            Width           =   1365
            _ExtentX        =   2408
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
            Caption         =   "2:出炉取消"
         End
      End
      Begin VB.TextBox TXT_DIS_UNDIS_IND 
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
         Left            =   1815
         MaxLength       =   1
         TabIndex        =   26
         Text            =   " "
         Top             =   255
         Width           =   630
      End
      Begin VB.TextBox TXT_EMP 
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
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   15
         Top             =   3540
         Width           =   1335
      End
      Begin VB.TextBox TXT_GROUP 
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
         Left            =   990
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   14
         Top             =   3540
         Width           =   705
      End
      Begin VB.TextBox TXT_SHIFT 
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
         Left            =   285
         MaxLength       =   1
         TabIndex        =   13
         Tag             =   "班次"
         Top             =   3540
         Width           =   705
      End
      Begin CSTextLibCtl.sidbEdit SDB_EXP_TEMP 
         Height          =   315
         Left            =   1815
         TabIndex        =   2
         Tag             =   "温度"
         Top             =   1065
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
      Begin CSTextLibCtl.sidbEdit SDB_PRE_TOP_SLAB_TEMP 
         Height          =   315
         Left            =   1815
         TabIndex        =   4
         Tag             =   "预热区上表温度"
         Top             =   1815
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_HT_TOP_SLAB_TEMP 
         Height          =   315
         Left            =   1815
         TabIndex        =   7
         Tag             =   "一号加热区上表温度"
         Top             =   2160
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_SOK_HOT_SLAB_TEMP 
         Height          =   315
         Left            =   1815
         TabIndex        =   21
         Tag             =   "均热区上表温度"
         Top             =   2835
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_PRE_BOT_SLAB_TEMP 
         Height          =   315
         Left            =   3000
         TabIndex        =   5
         Tag             =   "预热区下表温度"
         Top             =   1815
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_HT_BOT_SLAB_TEMP 
         Height          =   315
         Left            =   3000
         TabIndex        =   8
         Tag             =   "一号加热区下表温度"
         Top             =   2160
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_SOK_BOT_SLAB_TEMP 
         Height          =   315
         Left            =   3000
         TabIndex        =   22
         Tag             =   "均热区下表温度"
         Top             =   2835
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_PRE_ZONE_TIME 
         Height          =   315
         Left            =   4185
         TabIndex        =   6
         Tag             =   "预热区驻留时间"
         Top             =   1815
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_HT_ZONE_TIME 
         Height          =   315
         Left            =   4185
         TabIndex        =   9
         Tag             =   "一号加热区驻留时间"
         Top             =   2160
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_SOK_ZONE_TIME 
         Height          =   315
         Left            =   4185
         TabIndex        =   23
         Tag             =   "均热区驻留时间"
         Top             =   2835
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   285
         Top             =   1815
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "预热区"
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
         Left            =   285
         Top             =   2160
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "一号加热区"
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
         Left            =   285
         Top             =   2835
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "均热区"
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
         Left            =   1815
         Top             =   1455
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "TOP 温度"
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
         Left            =   3000
         Top             =   1455
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "BOT 温度"
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
         Left            =   4185
         Top             =   1455
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "驻留时间"
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
         Left            =   285
         Top             =   255
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "出炉/取消"
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
         Left            =   285
         Top             =   690
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "出炉时间"
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
         Left            =   285
         Top             =   1065
         Width           =   1500
         _ExtentX        =   2646
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
      Begin CSTextLibCtl.sitxEdit TXT_DISCHARGE_TIME 
         Height          =   315
         Left            =   1815
         TabIndex        =   1
         Tag             =   "出炉时间"
         Top             =   690
         Width           =   2100
         _Version        =   262145
         _ExtentX        =   3704
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   285
         Top             =   3210
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
      Begin InDate.ULabel ULabel38 
         Height          =   315
         Left            =   1020
         Top             =   3210
         Width           =   675
         _ExtentX        =   1191
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
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   1720
         Top             =   3210
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   2865
         Top             =   1065
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "温度均匀性"
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
      Begin CSTextLibCtl.sidbEdit SDB_PDT_UNI_TEMP 
         Height          =   315
         Left            =   4215
         TabIndex        =   3
         Tag             =   "温度均匀性"
         Top             =   1065
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
      Begin CSTextLibCtl.sidbEdit SDB_HT_TOP_SLAB_TEMP2 
         Height          =   315
         Left            =   1815
         TabIndex        =   10
         Tag             =   "二号加热区上表温度"
         Top             =   2490
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_HT_BOT_SLAB_TEMP2 
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         Tag             =   "二号加热区下表温度"
         Top             =   2490
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
      Begin CSTextLibCtl.sidbEdit SDB_HT_ZONE_TIME2 
         Height          =   315
         Left            =   4185
         TabIndex        =   12
         Tag             =   "二号加热区驻留时间"
         Top             =   2490
         Width           =   1170
         _Version        =   262145
         _ExtentX        =   2064
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
         Left            =   285
         Top             =   2490
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "二号加热区"
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
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2730
         TabIndex        =   27
         Top             =   585
         Width           =   255
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   720
      Left            =   165
      TabIndex        =   24
      Top             =   105
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   1270
      _Version        =   196609
      BackColor       =   14737632
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
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   36
         Top             =   150
         Width           =   1635
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
         Height          =   315
         Left            =   4995
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   " "
         Top             =   180
         Width           =   2805
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
         Top             =   180
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
      Begin Threed.SSPanel SSP1 
         Height          =   285
         Left            =   10200
         TabIndex        =   49
         Top             =   150
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16711935
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "出口订单"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSP4 
         Height          =   315
         Left            =   8790
         TabIndex        =   50
         Top             =   150
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "定制配送"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSP2 
         Height          =   285
         Left            =   11610
         TabIndex        =   51
         Top             =   150
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   196609
         ForeColor       =   0
         BackColor       =   65535
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "堆冷标识"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "mm"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7890
         TabIndex        =   35
         Top             =   225
         Width           =   285
      End
   End
   Begin Threed.SSFrame sf3 
      Height          =   4005
      Left            =   8295
      TabIndex        =   30
      Top             =   795
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7064
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_REJ_SHIFT 
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
         Left            =   270
         MaxLength       =   1
         TabIndex        =   18
         Tag             =   "班次"
         Top             =   3540
         Width           =   705
      End
      Begin VB.TextBox TXT_REJ_GROUP 
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
         Left            =   1005
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   19
         Tag             =   "班别"
         Top             =   3540
         Width           =   705
      End
      Begin VB.TextBox TXT_REJ_EMP 
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
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   20
         Top             =   3540
         Width           =   1335
      End
      Begin VB.TextBox TXT_REASON_CD 
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
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   17
         Tag             =   "缺号代码"
         Top             =   1020
         Width           =   645
      End
      Begin VB.TextBox TXT_REASON_NAME 
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
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox TXT_COMFRM 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   31
         Text            =   " "
         Top             =   1740
         Width           =   645
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
         TabIndex        =   16
         Tag             =   "缺号时间"
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
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   270
         Top             =   3210
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
         Left            =   1005
         Top             =   3210
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
         Left            =   1740
         Top             =   3210
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
         Left            =   2580
         TabIndex        =   40
         Top             =   1770
         Width           =   975
         _ExtentX        =   1720
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
         Left            =   3600
         TabIndex        =   41
         Top             =   1770
         Width           =   1125
         _ExtentX        =   1984
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
End
Attribute VB_Name = "CGB2020C"
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
'-- Program Name      加热炉作业实绩查询及修改界面
'-- Program ID        CGB2020C
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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
Dim sc5 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

'超交货期用红色显示 add by liqian 2012-06-11
Const SS1_STDSPEC = 9
Const SS1_UST_FL = 15
Const SS1_DEL_TO_DATE = 24
Const SS1_URGNT_FL = 25
Const SS1_OVER_FL = 2 '超量标记

Const SS2_SLAB_NO = 1
Const SS2_STDSPEC = 10
Const SS2_UST_FL = 17
Const SS2_DEL_TO_DATE = 27
Const SS2_URGNT_FL = 28
Const SS2_FLAG_FL = 29
Const SS2_EXPORT_FL = 30
Const SS2_OVER_FL = 2 '超量标记
Const SS2_DUILENG = 18

Const SS3_SLAB_NO = 1
Const SS3_STDSPEC = 10
Const SS3_UST_FL = 17
Const SS3_DEL_TO_DATE = 27
Const SS3_URGNT_FL = 28
Const SS3_FLAG_FL = 29
Const SS3_EXPORT_FL = 30
Const SS3_OVER_FL = 2 '超量标记
Const SS3_DUILENG = 18

Const SS5_SLAB_NO = 1
Const SS5_STDSPEC = 10
Const SS5_UST_FL = 17
Const SS5_DEL_TO_DATE = 27
Const SS5_URGNT_FL = 28
Const SS5_FLAG_FL = 29
Const SS5_EXPORT_FL = 30
Const SS5_OVER_FL = 2 '超量标记
Const SS5_DUILENG = 18



Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_SlabNo, "p", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(TXT_SLAB_SIZE, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(TXT_DIS_UNDIS_IND, " ", "n", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   Call Gp_Ms_Collection(TXT_DISCHARGE_TIME, " ", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(SDB_EXP_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(SDB_PDT_UNI_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)

Call Gp_Ms_Collection(SDB_PRE_TOP_SLAB_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(SDB_PRE_BOT_SLAB_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(SDB_PRE_ZONE_TIME, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_HT_TOP_SLAB_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
 Call Gp_Ms_Collection(SDB_HT_BOT_SLAB_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(SDB_HT_ZONE_TIME, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(SDB_HT_TOP_SLAB_TEMP2, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(SDB_HT_BOT_SLAB_TEMP2, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(SDB_HT_ZONE_TIME2, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(SDB_SOK_HOT_SLAB_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(SDB_SOK_BOT_SLAB_TEMP, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(SDB_SOK_ZONE_TIME, " ", " ", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(TXT_SHIFT, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(TXT_GROUP, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
              Call Gp_Ms_Collection(TXT_EMP, " ", " ", " ", "i", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            
    'MASTER Collection
     Mc1.Add Item:="CGB2020C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="CGB2020C.P_SREFER1", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
     
        Call Gp_Ms_Collection(txt_SlabNo, "p", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
 Call Gp_Ms_Collection(TXT_REJ_OCCR_TIME, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(TXT_REASON_CD, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(TXT_COMFRM, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(TXT_REJ_SHIFT, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(TXT_REJ_GROUP, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(TXT_REJ_EMP, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                 
     Mc2.Add Item:="CGB2020C.P_MODIFY3", Key:="P-M"
     Mc2.Add Item:="CGB2020C.P_SREFER2", Key:="P-R"
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
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '超量标记 20150330
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
 
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGB2020C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '超量标记 20150330
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
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '定尺类别 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '是否探伤 add by liqian 2013-04-08
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '超交货期用红色显示 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '超交货期用红色显示 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '是否紧急订单 add by liqian 2012-08-15
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '是否定制配送
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '是否出口订单
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '是否出口订单
   Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '是否出口订单
   Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '是否出口订单
   Call Gp_Sp_Collection(ss2, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '是否出口订单
   Call Gp_Sp_Collection(ss2, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
   'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGB2020C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3) '超量标记 20150330
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3) '定尺类别 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '是否探伤 add by liqian 2013-04-08
   Call Gp_Sp_Collection(ss3, 17, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 18, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 19, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 20, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 21, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 22, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 23, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss3, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '超交货期用红色显示 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss3, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '超交货期用红色显示 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss3, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '是否紧急订单 add by liqian 2012-08-15
   Call Gp_Sp_Collection(ss3, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss3, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss3, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss3, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss3, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss3, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss3, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

   'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="CGB2020C.P_REFER3", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
       'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss5, 1, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 2, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5) '超量标记 20150330
    Call Gp_Sp_Collection(ss5, 3, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 4, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 5, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 6, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 7, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 8, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Call Gp_Sp_Collection(ss5, 9, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 10, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 11, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 12, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 13, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 14, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5) '定尺类别 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss5, 15, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '是否探伤 add by liqian 2013-04-08
   Call Gp_Sp_Collection(ss5, 17, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 18, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 19, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 20, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 21, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 22, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 23, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
   Call Gp_Sp_Collection(ss5, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss5, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '超交货期用红色显示 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss5, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '超交货期用红色显示 add by liqian 2012-06-11
   Call Gp_Sp_Collection(ss5, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '是否紧急订单 add by liqian 2012-08-15
   Call Gp_Sp_Collection(ss5, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss5, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss5, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss5, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss5, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss5, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss3, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

   'Spread_Collection
    sc5.Add Item:=ss5, Key:="Spread"
    sc5.Add Item:="CGB2020C.P_REFER5", Key:="P-R"
    sc5.Add Item:=pColumn5, Key:="pColumn"
    sc5.Add Item:=nColumn5, Key:="nColumn"
    sc5.Add Item:=aColumn5, Key:="aColumn"
    sc5.Add Item:=mColumn5, Key:="mColumn"
    sc5.Add Item:=iColumn5, Key:="iColumn"
    sc5.Add Item:=lColumn5, Key:="lColumn"
    sc5.Add Item:=1, Key:="First"
    sc5.Add Item:=ss5.MaxCols, Key:="Last"
    
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
    Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4) '超量标记 20150330
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
   Call Gp_Sp_Collection(ss4, 16, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 17, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 18, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 19, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 20, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 21, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 22, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 23, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 24, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 25, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
  
   'Spread_Collection
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="CGB2020C.P_REFER4", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"
   

    Proc_Sc.Add Item:=sc1, Key:="Sc"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
End Sub

Private Sub opt_EntCan_Click(Index As Integer, Value As Integer)
    If Index = 0 Then
       TXT_DIS_UNDIS_IND.Text = "1"
       opt_EntCan(0).ForeColor = &HFF&       'Red color
       opt_EntCan(1).ForeColor = &H80000012    'Black color
    ElseIf Index = 1 Then
       TXT_DIS_UNDIS_IND.Text = "2"
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

Private Sub sChk1_Click(Value As Integer)
   
    If sChk1.Value = ssCBUnchecked Then
       If sChk2.Value = ssCBUnchecked Then
          sChk1.Value = ssCBChecked
          sChk1.ForeColor = &HFF&
       End If
       Exit Sub
    End If
            
    txt_EntCan = "1"
    
    sChk2.Value = ssCBUnchecked
    sChk2.ForeColor = &H80000012
    sChk1.ForeColor = &HFF&
    sf3.Enabled = False
    sf2.Enabled = True
    
    TXT_SHIFT = Gf_ShiftSet3(M_CN1)
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(TXT_SHIFT), Gf_DTSet(M_CN1, , "X"))
    TXT_EMP = sUserID
    opt_EntCan(0).Value = True

End Sub

Private Sub sChk2_Click(Value As Integer)
    If sChk2.Value = ssCBUnchecked Then
       If sChk1.Value = ssCBUnchecked Then
          sChk2.Value = ssCBChecked
          sChk2.ForeColor = &HFF&
       End If
       Exit Sub
    End If
    txt_EntCan = "2"
    
    sChk1.Value = ssCBUnchecked
    sChk1.ForeColor = &H80000012
    sChk2.ForeColor = &HFF&
    sf2.Enabled = False
    sf3.Enabled = True
    
    opt_ORDER(0).Value = True
    TXT_REJ_SHIFT = Gf_ShiftSet3(M_CN1)
    TXT_REJ_GROUP = Gf_GroupSet(M_CN1, Trim(TXT_REJ_SHIFT), Gf_DTSet(M_CN1, , "X"))
    TXT_REJ_EMP = sUserID
    Call TXT_REJ_OCCR_TIME_DblClick
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If sChk1.Value = -1 And opt_EntCan(0).Value = True Then
            ss1.ROW = ROW
            ss1.Col = 1
            txt_SlabNo.Text = ss1.Text
            ss1.Col = 10           '9
            TXT_SLAB_SIZE.Text = ss1.Text
                
            'TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        ElseIf sChk2.Value = -1 Then
            'TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        
    End If
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
        
    If ROW > 0 Then
        If sChk1.Value = -1 And opt_EntCan(0).Value = True Then
            ss1.ROW = ROW
            ss1.Col = 1
            txt_SlabNo.Text = ss1.Text
            ss1.Col = 10          '9
            TXT_SLAB_SIZE.Text = ss1.Text
                
            'TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        ElseIf sChk2.Value = -1 Then
            'TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        
    End If
        
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If sChk1.Value = -1 And opt_EntCan(0).Value = True Then
            ss2.ROW = ROW
            ss2.Col = 1
            txt_SlabNo.Text = ss2.Text
            ss2.Col = 11               '9
            TXT_SLAB_SIZE.Text = ss2.Text
                
            'TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        ElseIf sChk2.Value = -1 Then
            'TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        
    End If
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If sChk1.Value = -1 And opt_EntCan(0).Value = True Then
            ss2.ROW = ROW
            ss2.Col = 1
            txt_SlabNo.Text = ss2.Text
            ss2.Col = 11            '9
            TXT_SLAB_SIZE.Text = ss2.Text

            'TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        ElseIf sChk2.Value = -1 Then
            'TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If

    End If

    Dim iRow As Long
    
    With ss2

        iRow = ROW

       If ROW <> 0 Then
   
           Load CGB2021C
    
           .ROW = ROW
    
           .Col = 1: CGB2021C.txt_slab_no = .Text
      
       End If
            
        CGB2021C.Show 1

   End With
          
End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If sChk1.Value = -1 And opt_EntCan(0).Value = True Then
            ss3.ROW = ROW
            ss3.Col = 1
            txt_SlabNo.Text = ss3.Text
            ss3.Col = 11         '9
            TXT_SLAB_SIZE.Text = ss3.Text
                
            'TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        ElseIf sChk2.Value = -1 Then
            'TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        
    End If
End Sub

Private Sub ss3_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If sChk1.Value = -1 And opt_EntCan(0).Value = True Then
            ss3.ROW = ROW
            ss3.Col = 1
            txt_SlabNo.Text = ss3.Text
            ss3.Col = 11             '9
            TXT_SLAB_SIZE.Text = ss3.Text

            'TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        ElseIf sChk2.Value = -1 Then
            'TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If

    End If
    
    Dim iRow As Long
    
    With ss3

        iRow = ROW

       If ROW <> 0 Then
   
           Load CGB2021C
    
           .ROW = ROW
    
           .Col = 1: CGB2021C.txt_slab_no = .Text
      
       End If
            
        CGB2021C.Show 1

   End With
   
End Sub

Private Sub ss5_Click(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If sChk1.Value = -1 And opt_EntCan(0).Value = True Then
            ss5.ROW = ROW
            ss5.Col = 1
            txt_SlabNo.Text = ss5.Text
            ss5.Col = 11          '9
            TXT_SLAB_SIZE.Text = ss5.Text
                
            'TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        ElseIf sChk2.Value = -1 Then
            'TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        
    End If
End Sub

Private Sub ss5_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If sChk1.Value = -1 And opt_EntCan(0).Value = True Then
            ss5.ROW = ROW
            ss5.Col = 1
            txt_SlabNo.Text = ss5.Text
            ss5.Col = 11         '9
            TXT_SLAB_SIZE.Text = ss5.Text
                
            'TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        ElseIf sChk2.Value = -1 Then
            'TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
        End If
        
    End If
    
    Dim iRow As Long
    
    With ss5

        iRow = ROW

       If ROW <> 0 Then
   
           Load CGB2021C
    
           .ROW = ROW
    
           .Col = 1: CGB2021C.txt_slab_no = .Text
      
       End If
            
        CGB2021C.Show 1

   End With
   
End Sub


Private Sub ss4_Click(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If sChk2.Value = -1 Or opt_EntCan(1).Value = True Then
            ss4.ROW = ROW
            ss4.Col = 1
            txt_SlabNo.Text = ss4.Text
            If sChk2.Value = -1 Then
               If Not Gf_Ms_Refer(M_CN1, Mc2, , , False) Then
                  TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
                  TXT_REASON_CD.Text = ""
               End If
            Else
               Call Gf_Ms_Refer(M_CN1, Mc1, , , True)
            End If
        End If
    End If
End Sub

Private Sub ss4_DblClick(ByVal Col As Long, ByVal ROW As Long)
    If ROW > 0 Then
        If sChk2.Value = -1 Or opt_EntCan(1).Value = True Then
            ss4.ROW = ROW
            ss4.Col = 1
            txt_SlabNo.Text = ss4.Text
            If sChk2.Value = -1 Then
               If Not Gf_Ms_Refer(M_CN1, Mc2, , , False) Then
                  TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")
                  TXT_REASON_CD.Text = ""
               End If
            Else
               Call Gf_Ms_Refer(M_CN1, Mc1, , , True)
            End If
        End If
    End If
End Sub

Private Sub tab1_Click(PreviousTab As Integer)
    If tab1.Tab = "4" Then
        TXT_SHIFT = Gf_ShiftSet3(M_CN1)
        If TXT_SHIFT = "1" Then
            txt_RstFormDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "000001"
            txt_RstToDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "075959"
        ElseIf TXT_SHIFT = "2" Then
            txt_RstFormDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "080000"
            txt_RstToDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 8) & "155959"
        ElseIf TXT_SHIFT = "3" Then
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

Private Sub TXT_REASON_CD_Change()
    If TXT_REASON_CD = "" Then
       TXT_REASON_NAME = ""
    Else
       TXT_REASON_NAME = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'G0001' AND CD = '" & TXT_REASON_CD.Text & "' ")
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    Call Gp_Sp_Setting(sc4.Item("Spread"), False)
    Call Gp_Sp_Setting(sc5.Item("Spread"), False)
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(sc4)
    Call Gf_Sp_Cls(sc5)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc4.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc5.Item("Spread"), "CG-System.INI", Me.Name)

    TXT_SHIFT = Gf_ShiftSet3(M_CN1)
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(TXT_SHIFT), Gf_DTSet(M_CN1, , "X"))
    TXT_EMP = sUserID
    opt_EntCan(0).Value = True
    
    Call sChk1_Click(1)
    tab1.Tab = 0
    Call Form_Ref
          
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc4.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc5.Item("Spread"), "CG-System.INI", Me.Name)

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

    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set sc4 = Nothing
    Set sc5 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()

    If tab1.Tab = 0 Then
       Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf tab1.Tab = 1 Then
       Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf tab1.Tab = 2 Then
       Call Gp_Sp_Excel(Me, ss3, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf tab1.Tab = 3 Then
       Call Gp_Sp_Excel(Me, ss5, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    ElseIf tab1.Tab = 4 Then
       Call Gp_Sp_Excel(Me, ss4, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

    Dim sMesg As String

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(sc4)
    Call Gf_Sp_Cls(sc5)

   'Call Gp_SSCheck_Cls(MC("sControl"))
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

    pControl1(1).SetFocus
    
    opt_EntCan(0).Value = False
    opt_EntCan(0).ForeColor = &H80000012
    opt_EntCan(1).Value = False
    opt_EntCan(1).ForeColor = &H80000012
    TXT_DIS_UNDIS_IND.Text = ""
    
    opt_ORDER(0).Value = False
    opt_ORDER(0).ForeColor = &H80000012
    opt_ORDER(1).Value = False
    opt_ORDER(1).ForeColor = &H80000012
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
             
     End If

End Sub

Public Sub Form_Ref()
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sCurDate As String
    Dim sDel_To_Date As String
    Dim sUrgnt_Fl As String
    Dim sUst_Fl As String
    Dim sFlag As String
    Dim sexport As String
    Dim sOver_Fl As String
    
    Dim sFlag1 As String
    Dim sExport1 As String
    
    Dim sFlag2 As String
    Dim sExport2 As String
    
    sCurDate = Format(Now, "YYYYMM")

    If tab1.Tab = 0 Then
        Call Gf_Sp_Refer(M_CN1, sc1, , , , False)
        ss1.Col = 1
        ss1.ROW = 1
'        If ss1.Text <> "" Then
            Call ss1_Click(1, 1)
'        End If
        '超交货期用红色显示 add by liqian 2012-06-11
        With ss1
              For iRow = 1 To .MaxRows
                 .ROW = iRow:             .Col = SS1_DEL_TO_DATE
                  sDel_To_Date = Mid(.Value, 1, 6)
                  .Col = SS1_URGNT_FL:    sUrgnt_Fl = Trim(.Text)
                  If sDel_To_Date < sCurDate Then
                       Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, &HFF&)
                  End If
                   '紧急订单绿色显示 add by liqian 2012-08-15
                  If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, &HC000&)
                  End If
                  '是否探伤 add by liqian 2013-04-08
                  .ROW = iRow:
                  .Col = SS1_UST_FL:   sUst_Fl = Trim(.Text)
                  If sUst_Fl = "是" Then
                     Call Gp_Sp_BlockColor(ss1, SS1_UST_FL, SS1_UST_FL, iRow, iRow, &HFF00FF)
                     Call Gp_Sp_BlockColor(ss1, SS1_STDSPEC, SS1_STDSPEC, iRow, iRow, &HFF00FF)
                  End If
                  '是否超量 add by Lee 2015-03-30
                  .ROW = iRow:
                  .Col = SS1_OVER_FL:   sOver_Fl = Trim(.Text)
                  If sOver_Fl = "*" Then
                     Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, iRow, iRow, SSP2.BackColor)
                  End If
              Next iRow
        End With
    ElseIf tab1.Tab = 1 Then
        Call Gf_Sp_Refer(M_CN1, sc2, , , , False)
'        If ss2.Text <> "" Then
           Call ss2_Click(1, 1)
'        End If
        '超交货期用红色显示 add by liqian 2012-06-11
        With ss2
              For iRow = 1 To .MaxRows
                 .ROW = iRow:             .Col = SS2_DEL_TO_DATE
                  sDel_To_Date = Mid(.Value, 1, 6)
                  .Col = SS2_URGNT_FL:    sUrgnt_Fl = Trim(.Text)
                  If sDel_To_Date < sCurDate Then
                       Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iRow, iRow, &HFF&)
                  End If
                  '紧急订单绿色显示 add by liqian 2012-08-15
                  If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iRow, iRow, &HC000&)
                  End If
                   '是否探伤 add by liqian 2013-04-08
                  .ROW = iRow:
                  .Col = SS2_UST_FL:   sUst_Fl = Trim(.Text)
                  If sUst_Fl = "是" Then
                     Call Gp_Sp_BlockColor(ss2, SS2_UST_FL, SS2_UST_FL, iRow, iRow, &HFF00FF)
                     Call Gp_Sp_BlockColor(ss2, SS2_STDSPEC, SS2_STDSPEC, iRow, iRow, &HFF00FF)
                  End If
                  '是否定制配送
                  .ROW = iRow:
                  .Col = SS2_FLAG_FL: sFlag = Trim(.Text)
                  If sFlag = "Y" Then
                     Call Gp_Sp_BlockColor(ss2, SS2_SLAB_NO, SS2_SLAB_NO, iRow, iRow, SSP4.BackColor)
                  End If
                  '是否出口订单
                  .ROW = iRow:
                  .Col = SS2_EXPORT_FL: sexport = Trim(.Text)
                  If sexport = "Y" Then
                     Call Gp_Sp_BlockColor(ss2, SS2_SLAB_NO, SS2_SLAB_NO, iRow, iRow, SSP1.BackColor)
                  End If
                  '是否超量 add by Lee 2015-03-30
                  .ROW = iRow:
                  .Col = SS2_DUILENG:   sOver_Fl = Trim(.Text)
                  If sOver_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iRow, iRow, , SSP2.BackColor)
                  End If
              Next iRow
        End With
    ElseIf tab1.Tab = 2 Then
        Call Gf_Sp_Refer(M_CN1, Sc3, , , , False)
'        If ss3.Text <> "" Then
            Call ss3_Click(1, 1)
'        End If
        '超交货期用红色显示 add by liqian 2012-06-11
        With ss3
              For iRow = 1 To .MaxRows
                 .ROW = iRow:             .Col = SS3_DEL_TO_DATE
                  sDel_To_Date = Mid(.Value, 1, 6)
                  .Col = SS3_URGNT_FL:    sUrgnt_Fl = Trim(.Text)
                  If sDel_To_Date < sCurDate Then
                       Call Gp_Sp_BlockColor(ss3, 1, .MaxCols, iRow, iRow, &HFF&)
                  End If
                  '紧急订单绿色显示 add by liqian 2012-08-15
                  If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss3, 1, .MaxCols, iRow, iRow, &HC000&)
                  End If
                   '是否探伤 add by liqian 2013-04-08
                  .ROW = iRow:
                  .Col = SS3_UST_FL:   sUst_Fl = Trim(.Text)
                  If sUst_Fl = "是" Then
                     Call Gp_Sp_BlockColor(ss3, SS3_UST_FL, SS3_UST_FL, iRow, iRow, &HFF00FF)
                     Call Gp_Sp_BlockColor(ss3, SS3_STDSPEC, SS3_STDSPEC, iRow, iRow, &HFF00FF)
                  End If
                  '是否定制配送
                  .ROW = iRow:
                  .Col = SS3_FLAG_FL: sFlag = Trim(.Text)
                  If sFlag = "Y" Then
                     Call Gp_Sp_BlockColor(ss3, SS3_SLAB_NO, SS3_SLAB_NO, iRow, iRow, SSP4.BackColor)
                  End If
                  '是否出口订单
                  .ROW = iRow:
                  .Col = SS3_EXPORT_FL: sexport = Trim(.Text)
                  If sexport = "Y" Then
                     Call Gp_Sp_BlockColor(ss3, SS3_SLAB_NO, SS3_SLAB_NO, iRow, iRow, SSP1.BackColor)
                  End If
                  '是否超量 add by Lee 2015-03-30
                  .ROW = iRow:
                  .Col = SS3_DUILENG:   sOver_Fl = Trim(.Text)
                  If sOver_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss3, 1, .MaxCols, iRow, iRow, , SSP2.BackColor)
                  End If
              Next iRow
        End With
    ElseIf tab1.Tab = 3 Then
        Call Gf_Sp_Refer(M_CN1, sc5, , , , False)
'        If ss5.Text <> "" Then
            Call ss5_Click(1, 1)
'        End If
        '超交货期用红色显示 add by liqian 2012-06-11
        With ss5
              For iRow = 1 To .MaxRows
                 .ROW = iRow:             .Col = SS5_DEL_TO_DATE
                  sDel_To_Date = Mid(.Value, 1, 6)
                  .Col = SS5_URGNT_FL:    sUrgnt_Fl = Trim(.Text)
                  If sDel_To_Date < sCurDate Then
                       Call Gp_Sp_BlockColor(ss5, 1, .MaxCols, iRow, iRow, &HFF&)
                  End If
                  '紧急订单绿色显示 add by liqian 2012-08-15
                  If sUrgnt_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss5, 1, .MaxCols, iRow, iRow, &HC000&)
                  End If
                   '是否探伤 add by liqian 2013-04-08
                  .ROW = iRow:
                  .Col = SS5_UST_FL:   sUst_Fl = Trim(.Text)
                  If sUst_Fl = "是" Then
                     Call Gp_Sp_BlockColor(ss5, SS5_UST_FL, SS5_UST_FL, iRow, iRow, &HFF00FF)
                     Call Gp_Sp_BlockColor(ss5, SS5_STDSPEC, SS5_STDSPEC, iRow, iRow, &HFF00FF)
                  End If
                 '是否定制配送
                  .ROW = iRow:
                  .Col = SS5_FLAG_FL: sFlag = Trim(.Text)
                  If sFlag = "Y" Then
                     Call Gp_Sp_BlockColor(ss5, SS5_SLAB_NO, SS5_SLAB_NO, iRow, iRow, SSP4.BackColor)
                  End If
                  '是否出口订单
                  .ROW = iRow:
                  .Col = SS5_EXPORT_FL: sexport = Trim(.Text)
                  If sexport = "Y" Then
                     Call Gp_Sp_BlockColor(ss5, SS5_SLAB_NO, SS5_SLAB_NO, iRow, iRow, SSP1.BackColor)
                  End If
                  '是否超量 add by Lee 2015-03-30
                  .ROW = iRow:
                  .Col = SS5_DUILENG:   sOver_Fl = Trim(.Text)
                  If sOver_Fl = "Y" Then
                     Call Gp_Sp_BlockColor(ss5, 1, .MaxCols, iRow, iRow, , SSP2.BackColor)
                  End If
              Next iRow
        End With
    ElseIf tab1.Tab = 4 Then
        Call Gf_Sp_Refer(M_CN1, sc4, Mc3, Mc3("nControl"), Mc3("mControl"), False)
    End If
    
    
     
End Sub

Public Sub Form_Pro()
Dim sMesg As String
If sChk2.Value = -1 Then
   If Not Gp_DateCheck(TXT_REJ_OCCR_TIME) Then
          sMesg = " 请正确输入缺号时间 ！"
          Call Gp_MsgBoxDisplay(sMesg)
          Exit Sub
   End If
   If TXT_REASON_NAME.Text = "" Then
      sMesg = " 请正确选择缺号代码 ！"
      Call Gp_MsgBoxDisplay(sMesg)
      Exit Sub
   End If
ElseIf sChk1.Value = -1 Then
   If Not Gp_DateCheck(TXT_DISCHARGE_TIME) Then
          sMesg = " 请正确输入出炉时间 ！"
          Call Gp_MsgBoxDisplay(sMesg)
          Exit Sub
   End If
End If
    
    If txt_EntCan.Text = "1" Then
        If Not Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then Exit Sub
    ElseIf txt_EntCan.Text = "2" Then
        If Not Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then Exit Sub
    End If
    
    If tab1.Tab = 0 Then
        Call Gf_Sp_Refer(M_CN1, sc1, Nothing, Nothing, Nothing)
        Call ss1_Click(1, 1)
    ElseIf tab1.Tab = 1 Then
        Call Gf_Sp_Refer(M_CN1, sc2, Nothing, Nothing, Nothing)
        Call ss2_Click(1, 1)
    ElseIf tab1.Tab = 2 Then
        Call Gf_Sp_Refer(M_CN1, Sc3, Nothing, Nothing, Nothing)
        Call ss3_Click(1, 1)
    ElseIf tab1.Tab = 3 Then
        Call Gf_Sp_Refer(M_CN1, sc5, Nothing, Nothing, Nothing)
        Call ss5_Click(1, 1)
    ElseIf tab1.Tab = 4 Then
        Call Gf_Sp_Refer(M_CN1, sc4, Mc3, Mc3("nControl"), Mc3("mControl"))
    End If

    TXT_DISCHARGE_TIME = ""
    TXT_REJ_OCCR_TIME = ""
    '''ADDED BY GUOLI AT 20080326 避免保存后班次 班别 作业人员被清空''''
    TXT_SHIFT = Gf_ShiftSet3(M_CN1)
    TXT_GROUP = Gf_GroupSet(M_CN1, Trim(TXT_SHIFT), Gf_DTSet(M_CN1, , "X"))
    TXT_EMP = sUserID
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
    End With
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub


Private Sub TXT_DISCHARGE_TIME_DblClick()

    TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
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

Private Sub TXT_REJ_OCCR_TIME_DblClick()

     TXT_REJ_OCCR_TIME.RawData = Gf_DTSet(M_CN1, , "X")

End Sub


Private Sub txt_RstFormDate_DblClick()
    txt_RstFormDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 12)
    txt_RstToDate.RawData = Mid(Gf_DTSet(M_CN1, , "X"), 1, 12)
End Sub
