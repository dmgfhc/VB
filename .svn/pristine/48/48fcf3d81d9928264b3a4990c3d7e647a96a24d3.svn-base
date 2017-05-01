VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form DGB1000C 
   Caption         =   "淬火机作业实绩查询及修改_DGB1000C"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   12705
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   7530
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1770
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   13282
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
      MaxCols         =   43
      MaxRows         =   20
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "DGB1000C.frx":0000
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   660
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
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
         Left            =   3375
         MaxLength       =   2
         TabIndex        =   20
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
         ItemData        =   "DGB1000C.frx":0EAC
         Left            =   4980
         List            =   "DGB1000C.frx":0EAE
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
         Left            =   1860
         Locked          =   -1  'True
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
         Left            =   1305
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
         TabIndex        =   16
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
         ItemData        =   "DGB1000C.frx":0EB0
         Left            =   8910
         List            =   "DGB1000C.frx":0EB2
         Locked          =   -1  'True
         TabIndex        =   15
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
         Left            =   12630
         TabIndex        =   14
         Top             =   180
         Width           =   1620
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   90
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
         Left            =   3750
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
         Left            =   7710
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
         Left            =   11400
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
      Height          =   1020
      Left            =   120
      TabIndex        =   21
      Top             =   750
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   1799
      _Version        =   196609
      BackColor       =   12632319
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
         Left            =   13710
         MaxLength       =   1
         TabIndex        =   11
         Top             =   570
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
         Left            =   13710
         MaxLength       =   8
         TabIndex        =   10
         Top             =   105
         Width           =   1155
      End
      Begin VB.ComboBox CBO_COOL_TYPE 
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
         ItemData        =   "DGB1000C.frx":0EB4
         Left            =   8190
         List            =   "DGB1000C.frx":0EB6
         TabIndex        =   6
         Tag             =   "炉座号"
         Top             =   105
         Width           =   1830
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   3345
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "开始温度"
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
         Left            =   12630
         Top             =   570
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
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel36 
         Height          =   315
         Left            =   12630
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "作业人员"
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
         ForeColor       =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_STA_Temp 
         Height          =   315
         Left            =   4410
         TabIndex        =   2
         Top             =   105
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
         FmtThousands    =   0
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
         Left            =   90
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "开始时间"
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
      Begin CSTextLibCtl.sitxEdit TXT_STA_TIME 
         Height          =   315
         Left            =   1170
         TabIndex        =   0
         Tag             =   "出炉时间"
         Top             =   105
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   3345
         Top             =   570
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "结束温度"
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
      Begin CSTextLibCtl.sidbEdit txt_END_Temp 
         Height          =   315
         Left            =   4410
         TabIndex        =   3
         Top             =   570
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
         FmtThousands    =   0
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
         Left            =   5160
         Top             =   570
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "冷却速率"
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
      Begin CSTextLibCtl.sidbEdit TXT_COOL_RATIO 
         Height          =   315
         Left            =   6225
         TabIndex        =   5
         Top             =   570
         Width           =   780
         _Version        =   262145
         _ExtentX        =   1376
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.0"
         Text            =   " 0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   4
         MaxValue        =   0
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   10125
         Top             =   105
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "总水耗"
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
      Begin CSTextLibCtl.sidbEdit txt_TOT_WAT 
         Height          =   315
         Left            =   11535
         TabIndex        =   8
         Top             =   105
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
         MaxValue        =   9999.999
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   90
         Top             =   570
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "结束时间"
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
      Begin CSTextLibCtl.sitxEdit TXT_END_TIME 
         Height          =   315
         Left            =   1170
         TabIndex        =   1
         Tag             =   "出炉时间"
         Top             =   570
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   7110
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "冷却模式"
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
         Left            =   7110
         Top             =   570
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "摆动时间"
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
      Begin CSTextLibCtl.sidbEdit txt_PW_TIME 
         Height          =   315
         Left            =   8190
         TabIndex        =   7
         Top             =   570
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
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   5160
         Top             =   105
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "平均水温"
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
      Begin CSTextLibCtl.sidbEdit txt_AVE_Temp 
         Height          =   315
         Left            =   6225
         TabIndex        =   4
         Top             =   105
         Width           =   780
         _Version        =   262145
         _ExtentX        =   1376
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         MaxValue        =   20
         MinValue        =   10
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   10125
         Top             =   570
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         Caption         =   "钢板运行速度"
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
      Begin CSTextLibCtl.sidbEdit txt_SPEED 
         Height          =   315
         Left            =   11535
         TabIndex        =   9
         Top             =   570
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
         Modified        =   -1  'True
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
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
   End
End
Attribute VB_Name = "DGB1000C"
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
'-- Program Name      淬火机作业实绩查询及修改
'-- Program ID        DGB1000C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2009.10.29
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



Private Sub Form_Define()
Dim I As Integer
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_PrcLine, "p", "n", " ", " ", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(cbo_chg_no, "p", "n", " ", " ", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    'MASTER Collection
     Mc1.Add Item:="DGA1030C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="DGA1030C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
          
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

   For I = 25 To ss1.MaxCols
       Call Gp_Sp_Collection(ss1, I, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_ColHidden(ss1, I, True)
   Next I
   
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="DGB1000C.P_SREFER1", Key:="P-R"
    sc1.Add Item:="DGB1000C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:="DGB1000C.P_SONEROW", Key:="P-O"
    
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
End Sub
Public Sub Form_Ref()

  Dim I As Integer
  Dim iRow As Integer
  Dim iCol As Integer

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If txt_plt <> "C1" And txt_plt <> "C2" Then
         Call Gp_MsgBoxDisplay("只能查询工厂为C1和C2！！！")
         Exit Sub
    End If
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1) Then
'        If ss1.MaxRows > 0 Then
'            ss1.Row = 1
'            ss1.Col = 1
'            Call Gf_Ms_Refer(M_CN1, Mc1, Nothing, Nothing, False)
            'Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            
            
             For iRow = 1 To ss1.MaxRows
    
               ss1.Row = iRow
               ss1.Col = 24
                If ss1.Text = "Y" Then
                  For I = 1 To ss1.MaxCols
                       ss1.Col = I
                       ss1.ForeColor = &HC000&
                  Next
                End If
        
             Next iRow
            
            
'        End If
    End If
            
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Row Cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False                'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False                'Paste
    MDIMain.MenuTool.Buttons(14).Enabled = True                 'Excel
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

   Dim I As Integer
   Dim iRow As Integer
   Dim iCol As Integer

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    
     For iRow = 1 To ss1.MaxRows
    
               ss1.Row = iRow
               ss1.Col = 24
                If ss1.Text = "Y" Then
                  For I = 1 To ss1.MaxCols
                       ss1.Col = I
                       ss1.ForeColor = &HC000&
                  Next
                End If
        
     Next iRow
    
    
End Sub


Private Sub cbo_PrcLine_Click()

  If cbo_PrcLine.ListIndex = 0 Then
      txt_PrcLine = "1"
  ElseIf cbo_PrcLine.ListIndex = 1 Then
      txt_PrcLine = "2"
  ElseIf cbo_PrcLine.ListIndex = 2 Then
      txt_PrcLine = "3"
      
  End If
End Sub



Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim I As Integer
    Dim iRow As Integer
    Dim iCol As Integer

    
    If Row < 0 Then Exit Sub
    
If Col = 0 Then
    
    If TXT_STA_TIME.RawData = "" Or TXT_END_TIME.RawData = "" Then
       MsgBox "请确认开始时间和结束时间!"
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
    
    If txt_END_Temp.Value = 0 Then
       MsgBox "请确认结束温度....!"
       Screen.MousePointer = vbDefault
       Exit Sub
    End If

    ss1.Row = Row
    ss1.Col = 0
    
    If ss1.Text = "Update" Then
        ss1.Text = Row
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
        Call Gp_Sp_BlockColor(ss1, 1, 2, Row, Row, , &HC0FFFF)
        Call Gp_Sp_BlockColor(ss1, 5, 10, Row, Row, , &HC0FFFF)
        For I = 2 To 3
            ss1.Col = I
            ss1.Text = ""
        Next I
        
        For I = 5 To 14
            ss1.Col = I
            ss1.Text = ""
        Next I

    Else
        ss1.Text = "Update"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
        
        ss1.Col = 2
        ss1.Value = TXT_STA_TIME.RawData
        
        ss1.Col = 3
        ss1.Value = TXT_END_TIME.RawData
        
        ss1.Col = 5
        If txt_STA_Temp.Value > 0 Then
           ss1.Text = txt_STA_Temp.Value
        End If
        
        ss1.Col = 6
        If txt_END_Temp.Value > 0 Then
           ss1.Text = txt_END_Temp.Value
        End If
        
        ss1.Col = 7
        If txt_AVE_Temp.Value > 0 Then
            ss1.Text = txt_AVE_Temp.Value
        End If
        
        ss1.Col = 8
        If TXT_COOL_RATIO.Value > 0 Then
           ss1.Text = TXT_COOL_RATIO.Value
        End If
        
        ss1.Col = 9
        ss1.Text = CBO_COOL_TYPE.Text

        ss1.Col = 10
        If txt_PW_TIME.Value > 0 Then
            ss1.Text = txt_PW_TIME.Value
        End If
        
        ss1.Col = 11
        If txt_TOT_WAT.Value > 0 Then
            ss1.Text = txt_TOT_WAT.Value
        End If
        
        ss1.Col = 12
        If TXT_SPEED.Value > 0 Then
            ss1.Text = TXT_SPEED.Value
        End If
        
        ss1.Col = 13
        ss1.Text = Trim(TXT_SHIFT.Text)

        ss1.Col = 15
        ss1.Text = sUserID
        
        ss1.Col = 20
        ss1.Text = cbo_chg_no.Text
        
        ss1.Col = 21
        ss1.Text = txt_plt.Text
        
        ss1.Col = 22
        ss1.Text = txt_PrcLine.Text

    End If
    
     For iRow = 1 To ss1.MaxRows
    
               ss1.Row = iRow
               ss1.Col = 24
                If ss1.Text = "Y" Then
                  For I = 1 To ss1.MaxCols
                       ss1.Col = I
                       ss1.ForeColor = &HC000&
                  Next
                End If
        
      Next iRow
    
    
    
End If

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
        MDIMain.MenuTool.Buttons(7).Enabled = False                 'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Delete
        MDIMain.MenuTool.Buttons(9).Enabled = False                 'Row Cancel
        MDIMain.MenuTool.Buttons(11).Enabled = False                'Copy
        MDIMain.MenuTool.Buttons(12).Enabled = False                'Paste
        MDIMain.MenuTool.Buttons(14).Enabled = True                 'Excel
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
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "DG-System.INI", Me.Name)
    
    cbo_PrcLine.AddItem "一号热处理"
    cbo_PrcLine.AddItem "二号热处理"
    cbo_PrcLine.AddItem "三号热处理"
    cbo_PrcLine.ListIndex = 0
    
    ''''''ADDED BY GUOLI AT 20080904133500''''
    cbo_chg_no.AddItem "1"
    cbo_chg_no.AddItem "2"
    cbo_chg_no.AddItem "3"
    cbo_chg_no.Text = "3"
    
    txt_plt.Text = "C1"
    
    TXT_STA_TIME = Gf_DTSet(M_CN1, , "X")
    TXT_SHIFT = Gf_ShiftSet(M_CN1)
    TXT_EMP = sUserID
    
    CBO_COOL_TYPE.AddItem "1:连续淬火操作"
    CBO_COOL_TYPE.AddItem "2:连续加摆动操作"
    CBO_COOL_TYPE.AddItem "3:常化控冷操作"
    CBO_COOL_TYPE.AddItem "4:常化空过模式"
          
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "DG-System.INI", Me.Name)

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

    Dim SMESG As String

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
    
    TXT_STA_TIME = Gf_DTSet(M_CN1, , "X")
    TXT_SHIFT = Gf_ShiftSet3(M_CN1)
    TXT_EMP = sUserID
    
    txt_plt.Text = "C1"
    
    pControl1(1).SetFocus
    
End Sub

Private Sub TXT_STA_TIME_Change()
Dim for_cnt As Integer

    TXT_SHIFT = Gf_ShiftSet(M_CN1, Mid(TXT_STA_TIME.RawData, 9, 4))
    
    For for_cnt = 1 To ss1.MaxRows
        ss1.Row = for_cnt
        ss1.Col = 0
        If ss1.Text = "Input" Or ss1.Text = "Update" Then
            ss1.Col = 2
            ss1.Text = TXT_STA_TIME
            ss1.Col = 12
            ss1.Text = TXT_SHIFT

            ss1.Col = 15
            ss1.Text = sUserID
        End If
    Next
End Sub

Private Sub TXT_STA_TIME_DblClick()
    TXT_STA_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub

Private Sub TXT_END_TIME_DblClick()
    TXT_END_TIME.RawData = Gf_DTSet(M_CN1, , "X")
End Sub

Private Sub txt_Plt_Change()
    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        TXT_PLT_NAME.Text = ""
    End If
End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

