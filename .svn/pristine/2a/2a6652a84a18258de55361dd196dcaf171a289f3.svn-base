VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0600C 
   Caption         =   " 异钢种替代生产规范维护- AQA0600C"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   2730
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   14715
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   1614
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox Txt_DIFFER_CODE 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   13200
         MaxLength       =   18
         TabIndex        =   17
         Tag             =   "厂别"
         Top             =   600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox Txt_coolhot_load 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   13560
         MaxLength       =   18
         TabIndex        =   16
         Tag             =   "厂别"
         Top             =   600
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox Txt_flg 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   13920
         MaxLength       =   18
         TabIndex        =   15
         Tag             =   "厂别"
         Top             =   720
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.ComboBox Com_DIFFER_CODE 
         Height          =   300
         ItemData        =   "AQA0600C.frx":0000
         Left            =   1440
         List            =   "AQA0600C.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "1"
         Top             =   555
         Width           =   1095
      End
      Begin VB.ComboBox Com_flg 
         Height          =   300
         ItemData        =   "AQA0600C.frx":0020
         Left            =   12000
         List            =   "AQA0600C.frx":002A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "1"
         Top             =   555
         Width           =   1095
      End
      Begin VB.ComboBox Cmb_coolhot_load 
         Height          =   300
         ItemData        =   "AQA0600C.frx":003E
         Left            =   4200
         List            =   "AQA0600C.frx":0048
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "1"
         Top             =   555
         Width           =   1095
      End
      Begin VB.TextBox txt_STLGRD_Detail 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   4200
         MaxLength       =   18
         TabIndex        =   7
         Tag             =   "原始牌号"
         Top             =   120
         Width           =   2115
      End
      Begin VB.TextBox TxtPlt 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1440
         MaxLength       =   18
         TabIndex        =   3
         Tag             =   "工厂"
         Top             =   120
         Width           =   1035
      End
      Begin VB.TextBox TXT_TGT_STDSPEC 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   8160
         MaxLength       =   18
         TabIndex        =   1
         Tag             =   "目标钢种"
         Top             =   120
         Width           =   2115
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   1
         Left            =   6840
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "标准钢种"
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
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Index           =   1
         Left            =   6840
         Top             =   555
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "坯料厚度范围"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   0
         Left            =   240
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "厂别"
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
      Begin CSTextLibCtl.sidbEdit ET_ORD_THK_MIN 
         Height          =   315
         Left            =   12000
         TabIndex        =   2
         Tag             =   "订单厚度"
         Top             =   120
         Width           =   945
         _Version        =   262145
         _ExtentX        =   1667
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
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Index           =   0
         Left            =   10680
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "订单厚度范围"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   10
         Left            =   2880
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "原始牌号"
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
      Begin CSTextLibCtl.sidbEdit ET_SLAB_THK_MIN 
         Height          =   315
         Left            =   8160
         TabIndex        =   8
         Tag             =   "订单厚度"
         Top             =   555
         Width           =   945
         _Version        =   262145
         _ExtentX        =   1667
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit ET_SLAB_THK_MAX 
         Height          =   315
         Left            =   9360
         TabIndex        =   9
         Tag             =   "订单厚度"
         Top             =   555
         Width           =   945
         _Version        =   262145
         _ExtentX        =   1667
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit ET_ORD_THK_MAX 
         Height          =   315
         Left            =   13200
         TabIndex        =   11
         Tag             =   "订单厚度"
         Top             =   120
         Width           =   945
         _Version        =   262145
         _ExtentX        =   1667
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
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   30
         Left            =   240
         Top             =   555
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "区分代码"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   0
         Left            =   2880
         Top             =   555
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "冷/热装"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   33
         Left            =   10680
         Top             =   555
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "状态"
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
         ForeColor       =   16711680
      End
      Begin VB.Label Label4 
         Caption         =   "—"
         Height          =   255
         Left            =   9120
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "—"
         Height          =   255
         Left            =   12960
         TabIndex        =   5
         Top             =   165
         Width           =   255
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   50
      Left            =   3120
      Top             =   120
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      Caption         =   "热处理条件1"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   51
      Left            =   0
      Top             =   0
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      Caption         =   "热处理方法1"
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
   Begin CSTextLibCtl.sidbEdit sidbEdit1 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Tag             =   "订单厚度"
      Top             =   0
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
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
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
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
      FmtThousands    =   0
      FmtControl      =   1
      NumDecDigits    =   1
      NumIntDigits    =   3
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   3660
      Left            =   -240
      TabIndex        =   6
      Top             =   5520
      Width           =   15060
      _Version        =   393216
      _ExtentX        =   26564
      _ExtentY        =   6456
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   18
      MaxRows         =   5
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0600C.frx":0058
   End
   Begin Threed.SSPanel SSPanel2 
      Align           =   1  'Align Top
      Height          =   4560
      Left            =   0
      TabIndex        =   18
      Top             =   915
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   8043
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox Text_MES_STDSPEC 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   6390
         MaxLength       =   18
         TabIndex        =   77
         Tag             =   "目标钢种"
         Top             =   480
         Width           =   1650
      End
      Begin VB.TextBox TxtPlt2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1440
         MaxLength       =   18
         TabIndex        =   42
         Tag             =   "标准号"
         Top             =   120
         Width           =   1035
      End
      Begin VB.TextBox TXT_TGT_STDSPEC2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   7950
         MaxLength       =   18
         TabIndex        =   41
         Tag             =   "目标钢种"
         Top             =   120
         Width           =   2115
      End
      Begin VB.TextBox txt_UST_FL 
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
         Left            =   10800
         MaxLength       =   4
         TabIndex        =   40
         Top             =   1940
         Width           =   975
      End
      Begin VB.TextBox txt_UST_FL_NAME 
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
         Left            =   11760
         TabIndex        =   39
         Top             =   1940
         Width           =   2475
      End
      Begin VB.TextBox txt_CR_CD_Z 
         Height          =   300
         Left            =   10800
         MaxLength       =   1
         TabIndex        =   38
         Top             =   1580
         Width           =   975
      End
      Begin VB.TextBox txt_CR_NAME_Z 
         Enabled         =   0   'False
         Height          =   300
         Left            =   11775
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox txt_COOL_WAY_Z 
         Height          =   300
         Left            =   7485
         MaxLength       =   1
         TabIndex        =   36
         Top             =   2655
         Width           =   615
      End
      Begin VB.TextBox txt_COOL_WAY_NAME_Z 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8130
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2655
         Width           =   975
      End
      Begin VB.TextBox txt_COOL_CTL_TYP_Z 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   7485
         MaxLength       =   1
         TabIndex        =   34
         Top             =   3015
         Width           =   615
      End
      Begin VB.TextBox txt_COOL_CTL_NAME_Z 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3015
         Width           =   975
      End
      Begin VB.TextBox txt_HOT_LVL_USE_Z 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1980
         MaxLength       =   1
         TabIndex        =   32
         Top             =   2655
         Width           =   810
      End
      Begin VB.TextBox txt_HCR_KND_NAME_1 
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
         Left            =   11780
         MaxLength       =   11
         TabIndex        =   31
         Top             =   2640
         Width           =   2475
      End
      Begin VB.TextBox txt_HCR_KND_1 
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
         Left            =   10800
         MaxLength       =   1
         TabIndex        =   30
         Top             =   2655
         Width           =   975
      End
      Begin VB.TextBox txt_MILL_STD_EDT_NO 
         Height          =   300
         Left            =   1800
         MaxLength       =   80
         TabIndex        =   29
         Top             =   4125
         Width           =   9855
      End
      Begin VB.TextBox txt_ins_emp 
         Height          =   270
         Left            =   9960
         TabIndex        =   28
         Top             =   3765
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_DEL_STATE 
         Height          =   315
         Left            =   10800
         TabIndex        =   27
         Top             =   3015
         Width           =   1020
      End
      Begin VB.TextBox txt_STLGRD_Detail2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   3960
         MaxLength       =   18
         TabIndex        =   26
         Tag             =   "标准号"
         Top             =   120
         Width           =   1995
      End
      Begin VB.TextBox txt_DEL_STATE_DETAIL 
         Enabled         =   0   'False
         Height          =   300
         Left            =   11820
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3000
         Width           =   2415
      End
      Begin VB.ComboBox Cmb_coolhot_load2 
         Height          =   300
         ItemData        =   "AQA0600C.frx":085E
         Left            =   3960
         List            =   "AQA0600C.frx":0868
         TabIndex        =   24
         Text            =   $"AQA0600C.frx":0878
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox Com_DIFFER_CODE2 
         Height          =   300
         ItemData        =   "AQA0600C.frx":088F
         Left            =   1440
         List            =   "AQA0600C.frx":0899
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Tag             =   "1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Txt_DIFFER_CODE2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   12480
         MaxLength       =   18
         TabIndex        =   22
         Tag             =   "厂别"
         Top             =   3760
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox Txt_coolhot_load2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   12840
         MaxLength       =   18
         TabIndex        =   21
         Tag             =   "厂别"
         Top             =   3760
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ComboBox Com_flg2 
         Height          =   300
         ItemData        =   "AQA0600C.frx":08AF
         Left            =   13320
         List            =   "AQA0600C.frx":08B9
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Tag             =   "1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Txt_flg2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   13320
         MaxLength       =   18
         TabIndex        =   19
         Tag             =   "厂别"
         Top             =   3760
         Visible         =   0   'False
         Width           =   330
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   1
         Left            =   240
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "厂别"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   2
         Left            =   240
         Top             =   4125
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         Caption         =   "轧钢规范编辑号"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   3
         Left            =   240
         Top             =   840
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "加热温度目标值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   4
         Left            =   9720
         Top             =   1940
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "UST代码"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   5
         Left            =   240
         Top             =   1210
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "均热温度目标"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   6
         Left            =   6480
         Top             =   825
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "加热温度最大值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   7
         Left            =   9720
         Top             =   2655
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "HCR区分"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   8
         Left            =   3360
         Top             =   825
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "加热温度最小值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Index           =   0
         Left            =   6840
         Top             =   120
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         Caption         =   "标准钢种"
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
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Index           =   0
         Left            =   8160
         Top             =   465
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "坯料厚度范围"
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
         ForeColor       =   0
      End
      Begin CSTextLibCtl.sidbEdit ET_ORD_THK_MIN2 
         Height          =   315
         Left            =   12000
         TabIndex        =   43
         Tag             =   "订单厚度"
         Top             =   120
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
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
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Index           =   1
         Left            =   10680
         Top             =   120
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "订单厚度范围"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   13
         Left            =   6480
         Top             =   2295
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "冷却温度最大值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   14
         Left            =   3360
         Top             =   2295
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "冷却温度最小值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   15
         Left            =   6480
         Top             =   2655
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "冷却方法"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   16
         Left            =   240
         Top             =   1920
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "轧制温度目标值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   17
         Left            =   6480
         Top             =   1935
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "轧制温度最大值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   18
         Left            =   3360
         Top             =   1935
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "轧制温度最小值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   19
         Left            =   9720
         Top             =   3390
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "T2压下率"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   20
         Left            =   240
         Top             =   3390
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "T2温度"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   21
         Left            =   3360
         Top             =   3765
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "T1压下率"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   23
         Left            =   240
         Top             =   3765
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "T1温度"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   24
         Left            =   9720
         Top             =   1580
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "控轧代码"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   26
         Left            =   3360
         Top             =   1210
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "均热温度下限"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   27
         Left            =   6480
         Top             =   1210
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "均热温度上限"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   45
         Left            =   240
         Top             =   2655
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   556
         Caption         =   "热矫直机使用与否"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   48
         Left            =   3360
         Top             =   2655
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "冷却温度变化率"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   49
         Left            =   240
         Top             =   2295
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "冷却温度目标值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   31
         Left            =   6480
         Top             =   3015
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         Caption         =   "控制冷却"
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
      Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MAX_Z 
         Height          =   300
         Left            =   8280
         TabIndex        =   44
         Top             =   1935
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_MIN_Z 
         Height          =   300
         Left            =   5160
         TabIndex        =   45
         Top             =   1935
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_TGT_Z 
         Height          =   300
         Left            =   1755
         TabIndex        =   46
         Top             =   2295
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1808
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MAX_Z 
         Height          =   300
         Left            =   8280
         TabIndex        =   47
         Top             =   2295
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_MIN_Z 
         Height          =   300
         Left            =   5160
         TabIndex        =   48
         Top             =   2295
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_MILL_TMP_TGT_Z 
         Height          =   300
         Left            =   1755
         TabIndex        =   49
         Top             =   1935
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1808
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_HEAT_TGT_THK_MIN 
         Height          =   315
         Left            =   5160
         TabIndex        =   50
         Top             =   825
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_HEAT_TGT_THK 
         Height          =   315
         Left            =   1755
         TabIndex        =   51
         Top             =   840
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1808
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_HEAT_TGT_THK_MAX 
         Height          =   315
         Left            =   8280
         TabIndex        =   52
         Tag             =   "订单厚度"
         Top             =   825
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_AVGH_TGT_WID 
         Height          =   315
         Left            =   1755
         TabIndex        =   53
         Top             =   1210
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1808
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_AVGH_TGT_WID_MAX 
         Height          =   315
         Left            =   5160
         TabIndex        =   54
         Top             =   1210
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_AVGH_TGT_WID_MIN 
         Height          =   315
         Left            =   8280
         TabIndex        =   55
         Tag             =   "订单厚度"
         Top             =   1210
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_CR_MILL_TMPT2_Z 
         Height          =   300
         Left            =   1755
         TabIndex        =   56
         Top             =   3390
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1808
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_CR_MILL_TMPT1_Z 
         Height          =   300
         Left            =   1755
         TabIndex        =   57
         Top             =   3765
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1808
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_CR_MILL_RATET2_Z 
         Height          =   315
         Left            =   11400
         TabIndex        =   58
         Top             =   3390
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit ET_ORD_THK_MAX2 
         Height          =   315
         Left            =   13080
         TabIndex        =   59
         Tag             =   "订单厚度"
         Top             =   120
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
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
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   1
         Left            =   9720
         Top             =   3015
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "交货状态"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   12
         Left            =   240
         Top             =   3015
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "加热时间上限"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   22
         Left            =   6480
         Top             =   3765
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "冷床温度目标值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   28
         Left            =   3360
         Top             =   3015
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "加热时间下限"
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
         ForeColor       =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_COOL_BED_TMP_TGT_Z 
         Height          =   300
         Left            =   8280
         TabIndex        =   60
         Top             =   3765
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   11
         Left            =   2880
         Top             =   120
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         Caption         =   "原始牌号"
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Index           =   2
         Left            =   9720
         Top             =   825
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "均热段驻留时间下限"
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
      Begin CSTextLibCtl.sidbEdit ET_SLAB_THK_MIN2 
         Height          =   315
         Left            =   9480
         TabIndex        =   61
         Tag             =   "订单厚度"
         Top             =   465
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit ET_SLAB_THK_MAX2 
         Height          =   315
         Left            =   10560
         TabIndex        =   62
         Tag             =   "订单厚度"
         Top             =   465
         Width           =   825
         _Version        =   262145
         _ExtentX        =   1455
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   0
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Index           =   3
         Left            =   9720
         Top             =   1210
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         Caption         =   "均热段驻留时间上限"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   25
         Left            =   3360
         Top             =   3390
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "T2温度最小值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   29
         Left            =   6480
         Top             =   3390
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         Caption         =   "T2温度最大值"
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
         ForeColor       =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_CR_MILL_TMPT2_MIN 
         Height          =   300
         Left            =   5160
         TabIndex        =   63
         Top             =   3390
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_CR_MILL_TMPT2_MAX 
         Height          =   300
         Left            =   8160
         TabIndex        =   64
         Top             =   3390
         Width           =   945
         _Version        =   262145
         _ExtentX        =   1667
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_COOL_TMP_RATE_Z 
         Height          =   300
         Left            =   5160
         TabIndex        =   65
         Top             =   2655
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_AVG_HEAT_TIME1 
         Height          =   315
         Left            =   11640
         TabIndex        =   66
         Tag             =   "订单厚度"
         Top             =   825
         Width           =   1260
         _Version        =   262145
         _ExtentX        =   2222
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
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_AVG_HEAT_TIME2 
         Height          =   315
         Left            =   11640
         TabIndex        =   67
         Tag             =   "订单厚度"
         Top             =   1210
         Width           =   1260
         _Version        =   262145
         _ExtentX        =   2222
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
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit TXT_HEAT_TIME_MIN 
         Height          =   300
         Left            =   5160
         TabIndex        =   68
         Top             =   3015
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         NumDecDigits    =   2
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_CR_MILL_RATET1_Z 
         Height          =   300
         Left            =   5040
         TabIndex        =   69
         Top             =   3765
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         NumDecDigits    =   2
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit TXT_HEAT_TIME_MAX 
         Height          =   300
         Left            =   1755
         TabIndex        =   70
         Top             =   3015
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1799
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         NumDecDigits    =   2
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   32
         Left            =   2880
         Top             =   480
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "冷/热装"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   34
         Left            =   240
         Top             =   480
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "区分代码"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   9
         Left            =   12000
         Top             =   480
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "状态"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   35
         Left            =   240
         Top             =   1560
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Caption         =   "出钢温度目标值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   36
         Left            =   6480
         Top             =   1575
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "出钢温度最大值"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   37
         Left            =   3360
         Top             =   1575
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         Caption         =   "出钢温度最小值"
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
         ForeColor       =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_TAP_TMP_MAX 
         Height          =   300
         Left            =   8280
         TabIndex        =   73
         Top             =   1575
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_TAP_TMP_MIN 
         Height          =   300
         Left            =   5160
         TabIndex        =   74
         Top             =   1575
         Width           =   810
         _Version        =   262145
         _ExtentX        =   1429
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_TAP_TMP_TGT 
         Height          =   300
         Left            =   1755
         TabIndex        =   75
         Top             =   1575
         Width           =   1020
         _Version        =   262145
         _ExtentX        =   1808
         _ExtentY        =   529
         _StockProps     =   125
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         DataProperty    =   2
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   ""
         StartText.x     =   3
         StartText.y     =   2
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   38
         Left            =   9720
         Top             =   2280
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         Caption         =   "冷却速率"
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
      Begin CSTextLibCtl.sidbEdit txt_COOL_RATE 
         Height          =   315
         Left            =   11400
         TabIndex        =   76
         Top             =   2280
         Width           =   900
         _Version        =   262145
         _ExtentX        =   1587
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
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
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
         FmtThousands    =   0
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel MES_STDSPEC 
         Height          =   315
         Index           =   2
         Left            =   5400
         Top             =   480
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         Caption         =   "MES标准号"
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
      Begin VB.Label Label2 
         Caption         =   "—"
         Height          =   135
         Left            =   12840
         TabIndex        =   72
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "—"
         Height          =   135
         Left            =   10320
         TabIndex        =   71
         Top             =   600
         Width           =   195
      End
   End
End
Attribute VB_Name = "AQA0600C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量设计
'-- Program Name      规范设计结果修改及查询
'-- Program ID        AQB0160C
'-- Document No       Q-00-0010(Specification)
'-- Designer          WANG CHENG
'-- Coder             WANG CHENG
'-- Date              2013.05.09
'-- Description       规范设计结果修改及查询
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER     DATE          EDITOR       DESCRIPTION
'   1.1   2013.05.09    WANG CHENG
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

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Sc1 As New Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long
Dim lBlkcol2 As Long
Dim lBlkrow1 As Long
Dim lBlkrow2 As Long



Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"
    
                  'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
                    Call Gp_Ms_Collection(TxtPlt, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_STLGRD_DETAIL, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_TGT_STDSPEC, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(ET_SLAB_THK_MIN, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(ET_SLAB_THK_MAX, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(Txt_coolhot_load, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(Txt_DIFFER_CODE, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(ET_ORD_THK_MIN, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(ET_ORD_THK_MAX, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(Txt_flg, "P", " ", " ", " ", " ", " ", " ", pControl2, nControl, mControl, iControl, rControl, aControl, lControl)
                    
                   Call Gp_Ms_Collection(TxtPlt2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_STLGRD_Detail2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'        Call Gp_Ms_Collection(txt_STLGRD_Detail2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_TGT_STDSPEC2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(ET_SLAB_THK_MIN2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(ET_SLAB_THK_MAX2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(Txt_coolhot_load2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(Txt_DIFFER_CODE2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(ET_ORD_THK_MIN2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(ET_ORD_THK_MAX2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  Call Gp_Ms_Collection(Txt_flg2, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
          Call Gp_Ms_Collection(txt_HEAT_TGT_THK, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HEAT_TGT_THK_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_HEAT_TGT_THK_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 
          Call Gp_Ms_Collection(txt_AVGH_TGT_WID, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_AVGH_TGT_WID_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_AVGH_TGT_WID_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_AVG_HEAT_TIME1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
        Call Gp_Ms_Collection(txt_MILL_TMP_TGT_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MILL_TMP_MIN_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_MILL_TMP_MAX_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_AVG_HEAT_TIME2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
         
                
        Call Gp_Ms_Collection(txt_COOL_TMP_TGT_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_TMP_MIN_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_TMP_MAX_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_UST_FL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_UST_FL_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
         Call Gp_Ms_Collection(txt_HOT_LVL_USE_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_COOL_TMP_RATE_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_COOL_BED_TMP_TGT_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_HCR_KND_1, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_HCR_KND_NAME_1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
             
         Call Gp_Ms_Collection(TXT_HEAT_TIME_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_HEAT_TIME_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_COOL_WAY_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_COOL_WAY_NAME_Z, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_DEL_STATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_DEL_STATE_DETAIL, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
       Call Gp_Ms_Collection(txt_CR_MILL_TMPT1_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_CR_MILL_RATET1_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_COOL_CTL_TYP_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_COOL_CTL_NAME_Z, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_CR_CD_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_CR_NAME_Z, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
        
       Call Gp_Ms_Collection(txt_CR_MILL_TMPT2_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_CR_MILL_RATET2_Z, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_CR_MILL_TMPT2_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_CR_MILL_TMPT2_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
       Call Gp_Ms_Collection(txt_MILL_STD_EDT_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
               Call Gp_Ms_Collection(txt_ins_emp, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
      Call Gp_Ms_Collection(txt_TAP_TMP_TGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_TAP_TMP_MIN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_TAP_TMP_MAX, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_COOL_RATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Text_MES_STDSPEC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
    
    'MASTER Collection
     Mc1.Add Item:="AQA0600C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQA0600C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     'MASTER Collection
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"

     
     'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, "P", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, "P", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 10, "P", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 12, "P", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0600C.P_REFER1", Key:="P-R"
    Sc1.Add Item:="AQA0600C.P_MODIFY1", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 8, True)
    Call Gp_Sp_ColHidden(ss1, 10, True)
    Call Gp_Sp_ColHidden(ss1, 12, True)
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    

End Sub
'
'Private Sub Cmb_coolhot_load_Change()
'dim
'End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name, False)
'
    sAuthority = "1111"

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
     Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)

    TxtPlt.Text = "C3"
'    Cmb_coolhot_load.AddItem ("冷装")
'    Cmb_coolhot_load.AddItem ("热装")
'    Cmb_coolhot_load.ItemData = "1"
    Cmb_coolhot_load.ListIndex = 0
    Com_flg.ListIndex = 1
    Com_DIFFER_CODE.ListIndex = 0
    
     Cmb_coolhot_load2.ListIndex = 0
    Com_flg2.ListIndex = 1
    Com_DIFFER_CODE2.ListIndex = 0
'    Call Com_Data_Dind
   
    

    Screen.MousePointer = vbDefault
    
    
End Sub

Public Sub Form_Ref()

    Call comcheck
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc2, Mc2("nControl"), Mc2("mControl")) Then
    
'         Call Spread_to_Master(ss1, 1)
         
'        If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'            Call Gf_subMasterLock(Mc1, Trim(txt_Design_STS.Text))
        
'        End If
    End If
'    Call Gp_Sp_RowBackcolor(Proc_Sc("Sc").Item("Spread"))
    Call Gp_Ms_ControlLock(Mc2("pControl"), False)
    Com_DIFFER_CODE2.Enabled = False
    Cmb_coolhot_load2.Enabled = False
    Com_flg2.Enabled = False
    
   
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- ss1_DELETE ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Form_Pro()
     Dim i As Integer
     Dim sMesg As String
     txt_ins_emp.Text = sUserID
     Call comcheck
     If ET_ORD_THK_MIN2.Value > ET_ORD_THK_MAX2.Value Then
          Call Gp_MsgBoxDisplay("订单范围输入不正确")
          Exit Sub
     End If
     
     If ET_SLAB_THK_MIN2.Value > ET_SLAB_THK_MAX2.Value Then
      
          Call Gp_MsgBoxDisplay("板坯厚度范围输入不正确")
          Exit Sub
     End If
       
     If (txt_COOL_TMP_TGT_Z.Value <> 0 Or txt_COOL_TMP_MIN_Z.Value <> 0 Or txt_COOL_TMP_MAX_Z.Value <> 0) And txt_COOL_WAY_Z.Text <> "W" Then
     
       Call Gp_MsgBoxDisplay("请输入冷却方法：水冷")
       Exit Sub
     End If
     
       If (txt_COOL_TMP_TGT_Z.Value = 0 And txt_COOL_TMP_MIN_Z.Value = 0 And txt_COOL_TMP_MAX_Z.Value = 0) And txt_COOL_WAY_Z.Text <> "A" Then
     
       Call Gp_MsgBoxDisplay("请输入冷却方法为：自然冷却")
       Exit Sub
     End If
     
'     If (txt_CR_MILL_TMPT1_Z.Value = 0 And txt_CR_MILL_RATET1_Z.Value = 0) And txt_CR_CD_Z.Text <> "N" Then
'
'       Call Gp_MsgBoxDisplay("请输入控轧方法为：常规控轧")
'       Exit Sub
'     End If
        
     With ss1
        .Col = 0
        For i = 0 To ss1.MaxRows
            .Row = i
           If .Text = "Delete" Then
              If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc2) Then
                   Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        '           Call Gp_Ms_Cls(Mc1("rControl"))
              End If
              Call Form_Ref
              Exit Sub
           End If
        Next i
     End With
     
    sMesg = NullCheck
    If (sMesg <> "ok") Then
    
        Call Gp_MsgBoxDisplay(Trim(sMesg) + "必须输入", "I")
        Exit Sub
        
    End If
    If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
         Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    End If
       
     Call Form_Ref
     
End Sub

Public Sub Form_Cls()
    
     Call Gp_Ms_Cls(Mc1("rControl"))
     Call Gf_Sp_Cls(Proc_Sc("Sc"))
     Call Gp_Ms_Cls(Mc1("pControl"))
     Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     Com_flg.ListIndex = 1

End Sub

Private Function NullCheck() As String

 If (TxtPlt2.Text = "") Then
   NullCheck = TxtPlt2.Tag
   Exit Function
 End If
 
 If (txt_STLGRD_Detail2.Text = "") Then
 
   NullCheck = txt_STLGRD_Detail2.Tag
   Exit Function
 End If
 
 If (TXT_TGT_STDSPEC2.Text = "") Then
    
   NullCheck = TXT_TGT_STDSPEC2.Tag
   Exit Function
 End If
 
 NullCheck = "ok"

End Function

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_Activate --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- ss1_DELETEREMARK ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Public Sub Spread_Can()

    Call GP_ROW_CANCEL(Proc_Sc("SC"))
      
End Sub

Public Sub Form_Ins()
    'Spread Row Insert
    
'    Call Gp_MsgBoxDisplay("行插入")
     Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    Com_DIFFER_CODE2.Enabled = True
    Cmb_coolhot_load2.Enabled = True
    Com_flg2.Enabled = True
     
End Sub


Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
        If Row > 0 Then
             Call comcheck
             Call Spread_to_Master(ss1, Row)
             Call Gp_Ms_ControlLock(Mc1("pControl"), True)
             Com_DIFFER_CODE2.Enabled = False
             Cmb_coolhot_load2.Enabled = False
             Com_flg2.Enabled = False
             Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl"))
        End If
End Sub

Private Sub comcheck()
   If Com_DIFFER_CODE.ListIndex = 0 Then
       Txt_DIFFER_CODE.Text = "0"
    Else
       Txt_DIFFER_CODE.Text = "1"
    End If
    
    If Cmb_coolhot_load.ListIndex = 0 Then
       Txt_coolhot_load.Text = "0"
    Else
       Txt_coolhot_load.Text = "1"
    End If
    
    If Com_flg.ListIndex = 0 Then
       Txt_flg.Text = "0"
    Else
       Txt_flg.Text = "1"
    End If
    
    If Com_DIFFER_CODE2.ListIndex = 0 Then
       Txt_DIFFER_CODE2.Text = "0"
    Else
       Txt_DIFFER_CODE2.Text = "1"
    End If
    
    If Cmb_coolhot_load2.ListIndex = 0 Then
       Txt_coolhot_load2.Text = "0"
    Else
       Txt_coolhot_load2.Text = "1"
    End If
    
    If Com_flg2.ListIndex = 0 Then
       Txt_flg2.Text = "0"
    Else
       Txt_flg2.Text = "1"
    End If
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- ss1_LeaveRow ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
'Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
''     Call Gp_Ms_Cls(Mc1("rControl"))
'     Call Spread_to_Master(ss1, NewRow)
'     Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl"))
'End Sub

'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
        Case "txt_HCR_KND_1"            'HCR分类
            sCode = "C0005"
            Set oCodeName = txt_HCR_KND_NAME_1
                
        Case "txt_CR_CD_Z"                '控制轧制
            sCode = "Q0035"
            Set oCodeName = txt_CR_NAME_Z
                
        Case "txt_COOL_WAY_Z"             '冷却方法
            sCode = "Q0036"
            Set oCodeName = txt_COOL_WAY_NAME_Z
        
         Case "txt_UST_FL"               'UST與否
            sCode = "Q0046"
            Set oCodeName = txt_UST_FL_NAME

        Case "txt_COOL_CTL_TYP_Z"         '控制冷却代码
            sCode = "Q0037"
            Set oCodeName = txt_COOL_CTL_NAME_Z
            
'        Case "TXT_ORG_STDSPEC2"
'            sCode = "STLGRD"
'             Set oCodeName = txt_STLGRD_Detail2
            
'        Case "TXT_ORG_STDSPEC"
'            sCode = "STLGRD"
'            Set oCodeName = txt_STLGRD_Detail
       
        Case "txt_STLGRD_Detail"
            sCode = "STLGRD"
            Set oCodeName = txt_STLGRD_DETAIL
        
         Case "txt_STLGRD_Detail2"
            sCode = "STLGRD"
             Set oCodeName = txt_STLGRD_Detail2
        
        Case "TXT_TGT_STDSPEC2"
            sCode = "STD_STLGRD"
            
        Case "TXT_TGT_STDSPEC"
            sCode = "STD_STLGRD"
            
        Case "TxtPlt"                 '工厂
            sCode = "C0001"
            
        Case "TxtPlt2"                 '工厂
            sCode = "C0001"
            
        Case "txt_DEL_STATE"            '热处理方法
            sCode = "Q0073"
            Set oCodeName = txt_DEL_STATE_DETAIL
            
'        Case "txt_DEL_STATE1"            '热处理方法
'            sCode = "Q0073"
'            Set oCodeName = txt_DEL_STATE_DETAIL1
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Spread_to_Master ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Spread_to_Master(ByVal sp As vaSpread, ByVal iRow As Long)
    
    Dim RowLabel As String

    With sp
    
        If iRow > 0 Then
            .Row = iRow
             
            .Col = 1:  TxtPlt2.Text = .Text
'            .Col = 2:  TXT_ORG_STDSPEC.Text = .Text
            .Col = 2:  txt_STLGRD_Detail2.Text = .Text
            .Col = 3:  TXT_TGT_STDSPEC2.Text = .Text
            .Col = 4:  ET_SLAB_THK_MIN2.RawData = .Text
            .Col = 5:  ET_SLAB_THK_MAX2.RawData = .Text
                       ET_ORD_THK_MIN2.RawData = 0
            .Col = 6:  ET_ORD_THK_MIN2.RawData = .Text
                       ET_ORD_THK_MAX2.RawData = 0
            .Col = 7:  ET_ORD_THK_MAX2.RawData = .Text
            .Col = 8: Cmb_coolhot_load2.ListIndex = .Text
            .Col = 10: Com_DIFFER_CODE2.ListIndex = .Text
            .Col = 12: Com_flg2.ListIndex = .Text
            
        Else
            Exit Sub
        End If
    
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

'Private Sub ET_SLAB_THK_MIN2_Change()
'
'  ET_SLAB_THK_MIN.Text = ET_SLAB_THK_MIN2.Text
'
'End Sub
'
'Private Sub ET_ORD_THK3_Change()
'
'    ET_ORD_THK1.Text = ET_ORD_THK3.Text
'
'End Sub
'
'Private Sub ET_SLAB_THK_MAX2_Change()
'
'    ET_SLAB_THK_MAX.Text = ET_SLAB_THK_MAX2.Text
'
'End Sub
'
'Private Sub ET_ORD_THK2_Change()
'
'    ET_ORD_THK.Text = ""
'    ET_ORD_THK.Text = ET_ORD_THK2.Text
'
'End Sub


'Private Sub txt_DEL_STATE_Change()
'
'  txt_DEL_STATE1.Text = txt_DEL_STATE.Text
'
'End Sub
'
'Private Sub txt_DEL_STATE_DETAIL_Change()
'
'   txt_DEL_STATE_DETAIL1.Text = txt_DEL_STATE_DETAIL.Text
'
'End Sub

'Private Sub TXT_ORG_STDSPEC2_Change()
'
'    TXT_ORG_STDSPEC.Text = TXT_ORG_STDSPEC2.Text
'
'End Sub

'
'Private Sub Com_Data_Dind()
'    Dim sSQL As String
'
'      sSQL = "SELECT DISTINCT(THK_MAX) FROM EP_CCM_CON  ORDER BY THK_MAX ASC"
'
'    Call Gf_ComboAdd(M_CN1, Com_SLAB_THK, sSQL)
'     Call Gf_ComboAdd(M_CN1, Com_SLAB_THK2, sSQL)
'
'End Sub

'Private Sub txt_STLGRD_Detail2_Change()
'
'    txt_STLGRD_Detail.Text = txt_STLGRD_Detail2.Text
'
'End Sub
'
'Private Sub TXT_TGT_STDSPEC2_Change()
'
'    TXT_TGT_STDSPEC.Text = TXT_TGT_STDSPEC2.Text
'
'End Sub
'
'Private Sub TxtPlt2_Change()
'
'   TxtPlt.Text = TxtPlt2.Text
'
'End Sub


'---------------------------------------------------------------------------------------------------------------------------------------------
'---------------------------------------------------设置颜色 ------------------------------------------------------------------------
''---------------------------------------------------------------------------------------------------------------------------------------------
'Public Sub Gp_Sp_RowBackcolor(ByVal sPname As Variant, Optional MaxCnt As Integer = 0)
'
'    Dim i As Integer
'
'    With sPname
'        .ReDraw = False
'
'        For i = 1 To .MaxRows - MaxCnt
'            .Row = i
'
'                .BlockMode = True
'                .Row2 = i
'                .Col = 1: .Col2 = -1
'                .BackColor = &HFFFFFF
'                .BlockMode = False
'
'        Next i
'
'        .ReDraw = True
'        .Refresh
'
'    End With
'
'End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Form_Exc ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub




