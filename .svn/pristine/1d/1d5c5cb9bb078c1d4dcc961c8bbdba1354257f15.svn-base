VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form EGA1030C 
   Caption         =   "出炉作业实绩查询及修改_EGA1030C"
   ClientHeight    =   9495
   ClientLeft      =   -450
   ClientTop       =   1455
   ClientWidth     =   15240
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   6495
      Left            =   90
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2640
      Width           =   15210
      _Version        =   393216
      _ExtentX        =   26829
      _ExtentY        =   11456
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
      MaxCols         =   36
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "EGA1030C.frx":0000
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   660
      Left            =   90
      TabIndex        =   22
      Top             =   90
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   1164
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_PrcLine 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5190
         MaxLength       =   2
         TabIndex        =   24
         Tag             =   "工厂"
         Top             =   210
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.ComboBox cbo_PrcLine 
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
         ItemData        =   "EGA1030C.frx":0FEE
         Left            =   5460
         List            =   "EGA1030C.frx":0FF0
         TabIndex        =   19
         Tag             =   "炉座号"
         Top             =   180
         Width           =   1635
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
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "工厂"
         Top             =   180
         Width           =   1020
      End
      Begin VB.TextBox txt_Plt 
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
         Left            =   1395
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   17
         Tag             =   "工厂"
         Top             =   180
         Width           =   540
      End
      Begin VB.TextBox txt_iType 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   23
         Tag             =   "工厂"
         Text            =   "1"
         Top             =   210
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.ComboBox cbo_chg_no 
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
         ItemData        =   "EGA1030C.frx":0FF2
         Left            =   9375
         List            =   "EGA1030C.frx":0FF4
         TabIndex        =   20
         Tag             =   "炉座号"
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox TXT_MAT_NO 
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
         Left            =   13275
         TabIndex        =   0
         Top             =   180
         Width           =   1620
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   155
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   4230
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "产线别"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   8160
         Top             =   180
         Width           =   1170
         _ExtentX        =   2064
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   12045
         Top             =   180
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "查询号"
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
   Begin Threed.SSFrame SSFrame4 
      Height          =   1935
      Left            =   90
      TabIndex        =   25
      Top             =   720
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   3413
      _Version        =   196609
      BackColor       =   12632319
      Begin VB.TextBox TXT_DIS_SHIFT 
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
         Left            =   13995
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1005
         Width           =   855
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   8685
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "设定温度"
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
      Begin CSTextLibCtl.sidbEdit txt_HeatTemp 
         Height          =   315
         Left            =   10395
         TabIndex        =   3
         Top             =   180
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
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
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   135
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
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
      Begin CSTextLibCtl.sitxEdit TXT_DISCHARGE_TIME 
         Height          =   315
         Left            =   1860
         TabIndex        =   1
         Tag             =   "出炉时间"
         Top             =   180
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
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   135
         Top             =   1425
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "冷却温度(开/完)"
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   8685
         Top             =   600
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "加热温度"
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
      Begin CSTextLibCtl.sidbEdit txt_DisCharTemp 
         Height          =   315
         Left            =   10395
         TabIndex        =   7
         Top             =   600
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
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
      Begin CSTextLibCtl.sidbEdit TXT_ColStaTemp 
         Height          =   315
         Left            =   1860
         TabIndex        =   13
         Top             =   1425
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
      Begin CSTextLibCtl.sidbEdit TXT_ColEndTemp 
         Height          =   315
         Left            =   2925
         TabIndex        =   14
         Top             =   1425
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   4965
         Top             =   1005
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "加热速率"
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
      Begin CSTextLibCtl.sidbEdit txt_HEAT_RATIO 
         Height          =   315
         Left            =   6675
         TabIndex        =   10
         Top             =   1005
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
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
         NumIntDigits    =   2
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   12285
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "入炉速度"
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
      Begin CSTextLibCtl.sidbEdit txt_SP_CHARGE 
         Height          =   315
         Left            =   13995
         TabIndex        =   4
         Top             =   180
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         MaxValue        =   999.99
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   8685
         Top             =   1005
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "出炉速度"
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
      Begin CSTextLibCtl.sidbEdit txt_SP_DISCHARGE 
         Height          =   315
         Left            =   10395
         TabIndex        =   11
         Top             =   1005
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
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
         NumIntDigits    =   4
         MaxValue        =   9999.99
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   12285
         Top             =   600
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "工艺速度"
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
      Begin CSTextLibCtl.sidbEdit txt_SP_CAL 
         Height          =   315
         Left            =   13995
         TabIndex        =   8
         Top             =   600
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         MaxValue        =   999.99
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   135
         Top             =   600
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "进冷却区时间"
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
      Begin CSTextLibCtl.sitxEdit TXT_COL_IN_TIME 
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Tag             =   "出炉时间"
         Top             =   600
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
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   135
         Top             =   1005
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "出冷却区时间"
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
      Begin CSTextLibCtl.sitxEdit TXT_COL_OUT_TIME 
         Height          =   315
         Left            =   1860
         TabIndex        =   9
         Tag             =   "出炉时间"
         Top             =   1005
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
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   4965
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "加热段时间"
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
      Begin CSTextLibCtl.sidbEdit TXT_REHEAT_DT 
         Height          =   315
         Left            =   6675
         TabIndex        =   2
         Top             =   180
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   4965
         Top             =   600
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "均热段时间"
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
      Begin CSTextLibCtl.sidbEdit TXT_UNIFORM_DT 
         Height          =   315
         Left            =   6675
         TabIndex        =   6
         Top             =   600
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
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
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   4965
         Top             =   1425
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "头部流量"
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
         Left            =   8685
         Top             =   1425
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "尾部流量"
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
      Begin CSTextLibCtl.sidbEdit TXT_LOWER_FLOW 
         Height          =   315
         Left            =   10395
         TabIndex        =   16
         Top             =   1425
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
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
         NumIntDigits    =   2
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   12285
         Top             =   1005
         Width           =   1680
         _ExtentX        =   2963
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
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   12285
         Top             =   1440
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "驻留时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         Enabled         =   0   'False
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
      Begin CSTextLibCtl.sidbEdit TXT_IN_FCE_TM 
         Height          =   315
         Left            =   13995
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
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
         Enabled         =   0   'False
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
      Begin CSTextLibCtl.sidbEdit TXT_UPPER_FLOW 
         Height          =   315
         Left            =   6675
         TabIndex        =   15
         Top             =   1440
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
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
         NumIntDigits    =   2
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
   End
End
Attribute VB_Name = "EGA1030C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   HTM System
'-- Program Name      中板热处理出炉作业实绩查询及修改
'-- Program ID        EGA1030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2010.7.20
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

Const SS1_REHEAT_DT = 5                  '加热段时间
Const SS1_UNIFORM_DT = 6                 '均热段时间
Const SS1_IN_FCE_TM = 7                  '驻留时间
Const SS1_DISCHARGE_TIME = 8             '出炉时间
Const SS1_HeatTemp = 9                   '加热温度
Const SS1_DisCharTemp = 10               '出炉温度
Const SS1_DIS_SHIFT = 11                 '班次
Const SS1_sUserID = 13                   '作业人员
Const SS1_COL_IN_TIME = 14               '进冷却区时间
Const SS1_COL_OUT_TIME = 15              '出冷却区时间
Const SS1_ColStaTemp = 16                '冷却开始温度
Const SS1_ColEndTemp = 17                '冷却结束温度
Const SS1_UPPER_FLOW = 18                '头部流量
Const SS1_LOWER_FLOW = 19                '尾部流量
Const SS1_PLT = 27                       '工厂
Const SS1_PRC_LINE = 28                  '机座号
Const SS1_HEAT_RATIO = 31                '加热速率
Const SS1_SP_CHARGE = 32                 '入炉速度
Const SS1_SP_CAL = 33                    '工艺速度
Const SS1_SP_DISCHARGE = 34              '出炉速度

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_Plt, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_PrcLine, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(cbo_chg_no, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    'MASTER Collection
     Mc1.Add Item:="EGA1030C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
          
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '作业人员
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '厚度
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '宽度
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '长度
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '重量
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '探伤
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '试样
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '炉座号
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) 'PLT
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) 'PRC_LINE
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '备注
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '轧批号
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '加热速率
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '入炉速度
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '工艺速度
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '出炉速度
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '平均温度（计算）
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="EGA1030C.P_REFER1", Key:="P-R"
    sc1.Add Item:="EGA1030C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="EGA1030C.P_SONEROW", Key:="P-O"
    
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 1, True)
    Call Gp_Sp_ColHidden(ss1, SS1_IN_FCE_TM, True)  '驻留时间
    Call Gp_Sp_ColHidden(ss1, SS1_PLT, True)
    Call Gp_Sp_ColHidden(ss1, SS1_PRC_LINE, True)
    Call Gp_Sp_ColHidden(ss1, 35, True)
   
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
End Sub
Public Sub Form_Ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1) Then
        If ss1.MaxRows > 0 Then
            ss1.ROW = 1
            ss1.Col = 1
            Call Gf_Ms_Refer(M_CN1, Mc1, Nothing, Nothing, False)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    End If
            
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Row Cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False                'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False                'Paste
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
End Sub

Private Sub cbo_PrcLine_Click()
If cbo_PrcLine.ListIndex = 0 Then
   txt_PrcLine = "1"
   cbo_chg_no.Clear
   cbo_chg_no.List(0) = 1
   cbo_chg_no.List(1) = 2
   cbo_chg_no.List(2) = 3
   cbo_chg_no.List(3) = 4
Else
    txt_PrcLine = "2"
    cbo_chg_no.Clear
    cbo_chg_no.List(0) = 1
    cbo_chg_no.Text = "1"
    
End If
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    Dim iCol As Integer

    
    If ROW < 0 Then Exit Sub
    
If Col = 0 Then
    If Mid(TXT_DISCHARGE_TIME, 1, 1) <> "2" Then
        MsgBox "请先确认出炉时间......!", vbCritical, "系统提示信息"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
'    If TXT_REHEAT_DT.Value = 0 Then
'        MsgBox "请确定加热段时间....!", vbCritical, "系统提示信息"
'        Screen.MousePointer = vbDefault
'       Exit Sub
'    End If
     
    If txt_HeatTemp.Value = 0 Then
        MsgBox "请确定加热温度....!", vbCritical, "系统提示信息"
        Screen.MousePointer = vbDefault
       Exit Sub
    End If
    
    If txt_DisCharTemp.Value = 0 Then
       MsgBox "请确认出炉温度....!", vbCritical, "系统提示信息"
       Screen.MousePointer = vbDefault
       Exit Sub
    End If

    ss1.ROW = ROW
    ss1.Col = 0
    
    If ss1.Text = "Update" Then
        ss1.Text = ROW
        For iCol = SS1_REHEAT_DT To SS1_ColEndTemp     '加热段时间---冷却结束温度
            ss1.Col = iCol
            ss1.Text = ""
        Next iCol

    Else
        ss1.Text = "Update"
        
        ss1.Col = SS1_REHEAT_DT '加热段时间
        If TXT_REHEAT_DT.Value > 0 Then
           ss1.Text = TXT_REHEAT_DT.Value
        End If
        
        ss1.Col = SS1_UNIFORM_DT '均热段时间
        If TXT_UNIFORM_DT.Value > 0 Then
           ss1.Text = TXT_UNIFORM_DT.Value
        End If
        
        ss1.Col = SS1_IN_FCE_TM '驻留时间
        If TXT_IN_FCE_TM.Value > 0 Then
           ss1.Text = TXT_IN_FCE_TM.Value
        End If
        
        ss1.Col = SS1_DISCHARGE_TIME '出炉时间
        If TXT_DISCHARGE_TIME.RawData <> "" Then
            ss1.Value = TXT_DISCHARGE_TIME.RawData
        End If

        ss1.Col = SS1_HeatTemp '加热温度
        If txt_HeatTemp.Value > 0 Then
            ss1.Text = txt_HeatTemp.Value
        End If
        
        ss1.Col = SS1_DisCharTemp '出炉温度
        If txt_DisCharTemp.Value > 0 Then
            ss1.Text = txt_DisCharTemp.Value
        End If
        
        ss1.Col = SS1_DIS_SHIFT '班次
        If TXT_DIS_SHIFT.Text <> "" Then
            ss1.Text = TXT_DIS_SHIFT.Text
        End If
        
        ss1.Col = SS1_sUserID             '作业人员
        ss1.Text = sUserID
        
        ss1.Col = SS1_COL_IN_TIME '进冷却区时间
        If TXT_COL_IN_TIME.RawData <> "" Then
            ss1.Value = TXT_COL_IN_TIME.RawData
        End If
        
        ss1.Col = SS1_COL_OUT_TIME '出冷却区时间
        If TXT_COL_OUT_TIME.RawData <> "" Then
            ss1.Value = TXT_COL_OUT_TIME.RawData
        End If
        
        ss1.Col = SS1_ColStaTemp '冷却开始温度
        If TXT_ColStaTemp.Value > 0 Then
           ss1.Text = TXT_ColStaTemp.Value
        End If
        
        ss1.Col = SS1_ColEndTemp '冷却结束温度
        If TXT_ColEndTemp.Value > 0 Then
           ss1.Text = TXT_ColEndTemp.Value
        End If
        
        ss1.Col = SS1_UPPER_FLOW '头部流量
        If TXT_UPPER_FLOW.Value > 0 Then
           ss1.Text = TXT_UPPER_FLOW.Value
        End If
        
        ss1.Col = SS1_LOWER_FLOW '尾部流量
        If TXT_LOWER_FLOW.Value > 0 Then
           ss1.Text = TXT_LOWER_FLOW.Value
        End If
                
        ss1.Col = SS1_HEAT_RATIO '加热速率
        ss1.Text = txt_HEAT_RATIO.Value
         
        ss1.Col = SS1_SP_CHARGE '入炉速度
        ss1.Text = txt_SP_CHARGE.Value
        
        ss1.Col = SS1_SP_CAL '工艺速度
        ss1.Text = txt_SP_CAL.Value
        
        ss1.Col = SS1_SP_DISCHARGE '出炉速度
        ss1.Text = txt_SP_DISCHARGE.Value

    End If
End If

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
        MDIMain.MenuTool.Buttons(7).Enabled = False                 'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Delete
        MDIMain.MenuTool.Buttons(9).Enabled = False                 'Row Cancel
        MDIMain.MenuTool.Buttons(11).Enabled = False                'Copy
        MDIMain.MenuTool.Buttons(12).Enabled = False                'Paste

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
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    
    cbo_PrcLine.AddItem "一号线"
    cbo_PrcLine.AddItem "二号线"
    cbo_PrcLine.ListIndex = 1
    
    txt_Plt.Text = "C3"
    
    TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
    TXT_DIS_SHIFT = Gf_ShiftSet(M_CN1, Mid(TXT_DISCHARGE_TIME.RawData, 9, 4))
          
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "EG-System.INI", Me.Name)

    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
      
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
Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

    Dim sMesg As String

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gf_Sp_Cls(sc1)

    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = False                'Excel
    End With
    
    TXT_DISCHARGE_TIME = Gf_DTSet(M_CN1, , "X")
    TXT_DIS_SHIFT = Gf_ShiftSet(M_CN1, Mid(TXT_DISCHARGE_TIME.RawData, 9, 4))
    
    pControl1(1).SetFocus
    
End Sub

Private Sub TXT_DISCHARGE_TIME_Change()
Dim FOR_CNT As Integer

    TXT_DIS_SHIFT = Gf_ShiftSet(M_CN1, Mid(TXT_DISCHARGE_TIME.RawData, 9, 4))
    
    For FOR_CNT = 1 To ss1.MaxRows
        ss1.ROW = FOR_CNT
        ss1.Col = 0
        If ss1.Text = "Input" Or ss1.Text = "Update" Then
            ss1.Col = SS1_DISCHARGE_TIME
            ss1.Text = TXT_DISCHARGE_TIME
            ss1.Col = SS1_sUserID
            ss1.Text = sUserID
        End If
    Next
End Sub

Private Sub TXT_DISCHARGE_TIME_DblClick()
    TXT_DISCHARGE_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub
Private Sub TXT_COL_IN_TIME_DblClick()
    TXT_COL_IN_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub
Private Sub TXT_COL_OUT_TIME_DblClick()
    TXT_COL_OUT_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub
Private Sub txt_Plt_Change()
    If Len(Trim(txt_Plt.Text)) = txt_Plt.MaxLength Then
        TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_Plt.Text), 2)
    Else
        TXT_PLT_NAME.Text = ""
    End If
End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub
