VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CGF2031C 
   Caption         =   "轴承座、轴承保养的管理界面_CGF2031C"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   14130
   WindowState     =   2  'Maximized
   Begin Threed.SSCheck sc4 
      Height          =   315
      Left            =   9780
      TabIndex        =   0
      Top             =   975
      Width           =   1230
      _ExtentX        =   2170
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
      Caption         =   "护板修理"
   End
   Begin Threed.SSCheck sc1 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   1005
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
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
      Caption         =   "轧辊"
   End
   Begin Threed.SSCheck sc2 
      Height          =   315
      Left            =   2325
      TabIndex        =   2
      Top             =   1005
      Width           =   1380
      _ExtentX        =   2434
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
      Caption         =   "轴承座加油"
   End
   Begin Threed.SSCheck sc3 
      Height          =   315
      Left            =   6060
      TabIndex        =   3
      Top             =   1005
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "轴承清洗"
   End
   Begin Threed.SSFrame sf3 
      Height          =   7725
      Left            =   5790
      TabIndex        =   4
      Top             =   1125
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_C_IN_EMP 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   6945
         Width           =   1335
      End
      Begin VB.TextBox TXT_C_GROUP 
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
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   6945
         Width           =   705
      End
      Begin VB.TextBox TXT_C_SHIFT 
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
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   6945
         Width           =   705
      End
      Begin VB.TextBox TXT_C_ROLL_MOT 
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
         Left            =   1605
         TabIndex        =   6
         Tag             =   "供货商"
         Top             =   1344
         Width           =   1215
      End
      Begin VB.TextBox TXT_BEAR_NO 
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
         Left            =   1590
         TabIndex        =   5
         Tag             =   "轴承标识号"
         Top             =   852
         Width           =   1215
      End
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   270
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "清洗时间"
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
         Left            =   270
         Top             =   1350
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "清洗方式"
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
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   270
         Top             =   6615
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
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   975
         Top             =   6615
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
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   1680
         Top             =   6615
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
      End
      Begin CSTextLibCtl.sitxEdit UTP_C_ROLL_IN_TIME 
         Height          =   315
         Left            =   1605
         TabIndex        =   10
         Tag             =   "入库时间"
         Top             =   360
         Width           =   1785
         _Version        =   262145
         _ExtentX        =   3149
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __:__"
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
      Begin InDate.ULabel ULabel38 
         Height          =   315
         Left            =   270
         Top             =   855
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "轴承标识号"
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
   End
   Begin Threed.SSFrame sf2 
      Height          =   7725
      Left            =   2070
      TabIndex        =   11
      Top             =   1125
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_B_IN_EMP 
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
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   6945
         Width           =   1335
      End
      Begin VB.TextBox TXT_B_GROUP 
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
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   6945
         Width           =   705
      End
      Begin VB.TextBox TXT_B_SHIFT 
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
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   6945
         Width           =   705
      End
      Begin VB.TextBox TXT_CHOCK_NO 
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
         Left            =   1605
         TabIndex        =   12
         Tag             =   "轴承座标识号"
         Top             =   1365
         Width           =   1215
      End
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   270
         Top             =   870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "换轴承时间"
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
         Left            =   270
         Top             =   2370
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "滑板测量间距"
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
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   225
         Top             =   6615
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
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   945
         Top             =   6615
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
      Begin InDate.ULabel ULabel35 
         Height          =   315
         Left            =   1665
         Top             =   6615
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
      End
      Begin CSTextLibCtl.sitxEdit UTP_B_ROLL_IN_TIME 
         Height          =   315
         Left            =   1605
         TabIndex        =   16
         Tag             =   "入库时间"
         Top             =   870
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
         Text            =   "____-__-__ __:__"
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   270
         Top             =   1365
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "轴承座标识号"
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
      Begin CSTextLibCtl.sidbEdit SDB_B_PLANK_VAL 
         Height          =   315
         Left            =   1605
         TabIndex        =   49
         Tag             =   "轧辊重量"
         Top             =   2370
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         MaxValue        =   99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   270
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "开始使用时间"
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
      Begin CSTextLibCtl.sitxEdit UTP_B_ROLL_USE_TIME 
         Height          =   315
         Left            =   1605
         TabIndex        =   50
         Tag             =   "入库时间"
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
         Text            =   "____-__-__ __:__"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   270
         Top             =   1860
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "滑板测量时间"
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
      Begin CSTextLibCtl.sitxEdit SDB_B_PLANK_DATE 
         Height          =   315
         Left            =   1605
         TabIndex        =   51
         Tag             =   "入库时间"
         Top             =   1860
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
         Text            =   "____-__-__ __:__"
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
   End
   Begin Threed.SSFrame sf1 
      Height          =   7725
      Left            =   210
      TabIndex        =   17
      Top             =   1125
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_R_IN_EMP 
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
         Left            =   1680
         TabIndex        =   23
         Top             =   6945
         Width           =   1335
      End
      Begin VB.TextBox TXT_R_GROUP 
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
         Left            =   975
         TabIndex        =   22
         Top             =   6945
         Width           =   705
      End
      Begin VB.TextBox TXT_R_SHIFT 
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
         TabIndex        =   21
         Top             =   6945
         Width           =   705
      End
      Begin VB.TextBox TXT_ROLL_MATERIAL 
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
         Left            =   1605
         TabIndex        =   20
         Tag             =   "材质代码"
         Top             =   5775
         Width           =   1215
      End
      Begin VB.TextBox TXT_ROLL_MAKER 
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
         Left            =   1605
         TabIndex        =   19
         Tag             =   "供货商"
         Top             =   1344
         Width           =   1215
      End
      Begin VB.TextBox TXT_ROLL_NO 
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
         Left            =   1605
         TabIndex        =   18
         Tag             =   "轧辊标识号"
         Top             =   852
         Width           =   1215
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   24
         Tag             =   "入库辊径"
         Top             =   1830
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_SHLD_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   25
         Tag             =   "辊肩直径"
         Top             =   2325
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_NECK_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   26
         Tag             =   "辊颈直径"
         Top             =   2820
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         FmtControl      =   1
         NumDecDigits    =   2
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_WGT 
         Height          =   315
         Left            =   1605
         TabIndex        =   27
         Tag             =   "轧辊重量"
         Top             =   3315
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         MaxValue        =   999999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_IN_W_HARD 
         Height          =   315
         Left            =   1605
         TabIndex        =   28
         Tag             =   "工作侧硬度"
         Top             =   3810
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         MaxValue        =   99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_IN_C_HARD 
         Height          =   315
         Left            =   1605
         TabIndex        =   29
         Tag             =   "中部硬度"
         Top             =   4290
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         MaxValue        =   99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_IN_D_HARD 
         Height          =   315
         Left            =   1605
         TabIndex        =   30
         Tag             =   "驱动侧硬度"
         Top             =   4785
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         MaxValue        =   99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_IN_AVE_HARD 
         Height          =   315
         Left            =   1605
         TabIndex        =   31
         Tag             =   "平均硬度"
         Top             =   5280
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         MaxValue        =   99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   270
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "入库时间"
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
         Left            =   270
         Top             =   1350
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "供货商"
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   270
         Top             =   1830
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "入库辊径"
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
         Top             =   2325
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "辊肩直径"
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
         Left            =   270
         Top             =   3810
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "工作侧硬度"
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
         Top             =   4290
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "中部硬度"
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
         Left            =   270
         Top             =   2820
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "辊颈直径"
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
         Left            =   270
         Top             =   3315
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "轧辊重量"
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
         Left            =   270
         Top             =   4785
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "驱动侧硬度"
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
         Left            =   270
         Top             =   5280
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "平均硬度"
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
         Left            =   270
         Top             =   5775
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "材质代码"
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
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   270
         Top             =   6615
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
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   975
         Top             =   6615
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
         Top             =   6615
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "作业人员"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
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
      Begin CSTextLibCtl.sitxEdit UTP_ROLL_IN_TIME 
         Height          =   315
         Left            =   1605
         TabIndex        =   32
         Tag             =   "入库时间"
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
         Text            =   "____-__-__ __:__"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   270
         Top             =   855
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "轧辊标识号"
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
      Begin VB.Label Label59 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   2835
         TabIndex        =   36
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   2835
         TabIndex        =   35
         Top             =   1860
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   2835
         TabIndex        =   34
         Top             =   2370
         Width           =   375
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
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
         Left            =   2835
         TabIndex        =   33
         Top             =   3390
         Width           =   375
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   735
      Left            =   210
      TabIndex        =   37
      Top             =   120
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   1296
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.ComboBox CBO_GROUP 
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
         ItemData        =   "CGF2031C.frx":0000
         Left            =   9630
         List            =   "CGF2031C.frx":0010
         TabIndex        =   42
         Top             =   195
         Width           =   735
      End
      Begin VB.ComboBox CBO_SHIFT 
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
         ItemData        =   "CGF2031C.frx":0020
         Left            =   7575
         List            =   "CGF2031C.frx":002D
         TabIndex        =   41
         Top             =   195
         Width           =   735
      End
      Begin VB.TextBox TXT_ROLL_IN_EMP 
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
         Left            =   11835
         MaxLength       =   7
         TabIndex        =   40
         Top             =   195
         Width           =   1215
      End
      Begin VB.ComboBox CBO_PLT 
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
         ItemData        =   "CGF2031C.frx":003A
         Left            =   5040
         List            =   "CGF2031C.frx":0041
         TabIndex        =   39
         Top             =   195
         Width           =   735
      End
      Begin VB.ComboBox CBO_ROLL_NO 
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
         Left            =   1590
         TabIndex        =   38
         Tag             =   "ROLL_NO"
         Top             =   195
         Width           =   1365
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   510
         Top             =   195
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Caption         =   "轧辊号"
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
         Left            =   3930
         Top             =   195
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "工厂代码"
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
         Left            =   6585
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   8640
         Top             =   195
         Width           =   975
         _ExtentX        =   1720
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   10725
         Top             =   195
         Width           =   1095
         _ExtentX        =   1931
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
      End
   End
   Begin Threed.SSFrame sf4 
      Height          =   7725
      Left            =   9510
      TabIndex        =   43
      Top             =   1125
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_PLANK_NO 
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
         Left            =   1605
         TabIndex        =   47
         Tag             =   "轴承标识号"
         Top             =   852
         Width           =   1215
      End
      Begin VB.TextBox TXT_P_SHIFT 
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
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   6945
         Width           =   705
      End
      Begin VB.TextBox TXT_P_GROUP 
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
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   6945
         Width           =   705
      End
      Begin VB.TextBox TXT_P_IN_EMP 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   6945
         Width           =   1335
      End
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   270
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "修理时间"
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
         Left            =   270
         Top             =   6615
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
      Begin InDate.ULabel ULabel45 
         Height          =   315
         Left            =   975
         Top             =   6615
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
      Begin InDate.ULabel ULabel46 
         Height          =   315
         Left            =   1680
         Top             =   6615
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
      End
      Begin CSTextLibCtl.sitxEdit UTP_P_ROLL_IN_TIME 
         Height          =   315
         Left            =   1605
         TabIndex        =   48
         Tag             =   "入库时间"
         Top             =   360
         Width           =   1785
         _Version        =   262145
         _ExtentX        =   3149
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __:__"
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
      Begin InDate.ULabel ULabel47 
         Height          =   315
         Left            =   270
         Top             =   855
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "护板标识号"
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
   End
End
Attribute VB_Name = "CGF2031C"
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
'-- Program Name      轴承座、轴承保养的管理界面
'-- Program ID        CGF2031C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2007.10.10
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
Public sDateTime As String          'Active Form Time Setting
Public sQuery_load As String        'Active Form sQuery Setting

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

Dim pControl3 As New Collection     'Master Primary Key Collection
Dim nControl3 As New Collection     'Master Necessary Collection
Dim mControl3 As New Collection     'Master Maxlength check Collection
Dim iControl3 As New Collection     'Master Insert Collection
Dim rControl3 As New Collection     'Master Refer Collection
Dim cControl3 As New Collection     'Master Copy Collection
Dim aControl3 As New Collection     'Master -> Spread Collection
Dim lControl3 As New Collection     'Master Lock Collection

Dim pControl4 As New Collection     'Master Primary Key Collection
Dim nControl4 As New Collection     'Master Necessary Collection
Dim mControl4 As New Collection     'Master Maxlength check Collection
Dim iControl4 As New Collection     'Master Insert Collection
Dim rControl4 As New Collection     'Master Refer Collection
Dim cControl4 As New Collection     'Master Copy Collection
Dim aControl4 As New Collection     'Master -> Spread Collection
Dim lControl4 As New Collection     'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim Mc4 As New Collection           'Master Collection

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CBO_PLT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(UTP_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ROLL_NO, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_ROLL_DIA, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_ROLL_SHLD_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_ROLL_NECK_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_ROLL_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_W_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_C_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_D_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(SDB_ROLL_IN_AVE_HARD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_ROLL_MATERIAL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_R_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_R_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_R_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

          Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'              Call Gp_Ms_Collection(TXT_PLT, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(UTP_B_ROLL_USE_TIME, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(UTP_B_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_CHOCK_NO, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(SDB_B_PLANK_DATE, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(SDB_B_PLANK_VAL, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_B_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_B_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_B_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           
           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
'              Call Gp_Ms_Collection(TXT_PLT, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
       Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(UTP_C_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_BEAR_NO, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(TXT_C_ROLL_MOT, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_C_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_C_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          Call Gp_Ms_Collection(TXT_C_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          
           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
       Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(UTP_P_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(TXT_PLANK_NO, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
           Call Gp_Ms_Collection(TXT_P_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
           Call Gp_Ms_Collection(TXT_P_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(TXT_P_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)

    'MASTER Collection
     Mc1.Add Item:="CGF2031C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="CGF2031C.P_REFER1", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     Mc2.Add Item:="CGF2031C.P_MODIFY1", Key:="P-M"
     Mc2.Add Item:="CGF2031C.P_REFER1", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
     
     Mc3.Add Item:="CGF2031C.P_MODIFY2", Key:="P-M"
     Mc3.Add Item:="CGF2031C.P_REFER2", Key:="P-R"
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"
     
     Mc4.Add Item:="CGF2031C.P_MODIFY4", Key:="P-M"
     Mc4.Add Item:="CGF2031C.P_REFER4", Key:="P-R"
     Mc4.Add Item:=pControl4, Key:="pControl"
     Mc4.Add Item:=nControl4, Key:="nControl"
     Mc4.Add Item:=mControl4, Key:="mControl"
     Mc4.Add Item:=iControl4, Key:="iControl"
     Mc4.Add Item:=rControl4, Key:="rControl"
     Mc4.Add Item:=cControl4, Key:="cControl"
     Mc4.Add Item:=aControl4, Key:="aControl"
     Mc4.Add Item:=lControl4, Key:="lControl"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
     
     'sQuery_load
'    sQuery_load = "SELECT GOODS_ID FROM (SELECT B.GOODS_ID FROM FP_TRACKIDX A, FP_TRACKDATA B WHERE B.SEQ_NO <= A.LAST_SEQ AND A.FACT_CD = 'C1' " _
'    & "AND A.PRC = 'CD' AND A.PRC_LINE='1' AND A.FACT_CD=B.FACT_CD  AND A.PRC=B.PRC AND A.PRC_LINE=B.PRC_LINE ORDER BY B.SEQ_NO DESC) WHERE ROWNUM<=5"

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

    Dim sQuery_Rt As String
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    TXT_ROLL_IN_EMP = sUserID ' + ":" + sUsername
    CBO_PLT.ListIndex = 0
   
    If CBO_SHIFT.Text <> "1" Or CBO_SHIFT.Text <> "2" Or CBO_SHIFT.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet3(M_CN1)
       End If
       CBO_SHIFT.Text = sShiftSet
    End If
   
    sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK3 WHERE ROLL_STATUS<>'DL'  "
    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
    
    Sc2.ForeColor = &HFF&
    sc1.ForeColor = &H808080
    sc3.ForeColor = &H808080
    sc4.ForeColor = &H808080
    Sc2.Value = ssCBChecked
    sc1.Value = ssCBUnchecked
    sc3.Value = ssCBUnchecked
    sc4.Value = ssCBUnchecked
    sf2.Enabled = True
    SF1.Enabled = False
    sf3.Enabled = False
    SF4.Enabled = False
    ULabel16.Caption = "轴承座号"
    
    Screen.MousePointer = vbDefault

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
    
    Set pControl4 = Nothing
    Set nControl4 = Nothing
    Set iControl4 = Nothing
    Set rControl4 = Nothing
    Set cControl4 = Nothing
    Set aControl4 = Nothing
    Set lControl4 = Nothing
    Set mControl4 = Nothing

    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Mc4 = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))
    Call Gp_Ms_Cls(Mc4("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    sc1.ForeColor = &HFF&
    sc1.Value = ssCBChecked
    Sc2.ForeColor = &H808080
    Sc2.Value = ssCBUnchecked
    sc3.ForeColor = &H808080
    sc3.Value = ssCBUnchecked
    sc4.ForeColor = &H808080
    sc4.Value = ssCBUnchecked
    SF1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
    SF4.Enabled = False
    ULabel16.Caption = "轧辊号"
    CBO_ROLL_NO.Clear
    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL3 WHERE ROLL_STATUS<>'DL'  "
    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

    TXT_ROLL_IN_EMP = sUserID ' + ":" + sUsername
    CBO_PLT.ListIndex = 0
   
    If CBO_SHIFT.Text <> "1" Or CBO_SHIFT.Text <> "2" Or CBO_SHIFT.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet3(M_CN1)
       End If
       CBO_SHIFT.Text = sShiftSet
    End If
    
End Sub

Public Sub Master_Cpy()

'    Call Gf_Ms_Copy(Mc1)
'    Call Gf_Ms_Copy(Mc2)
'    Call Gf_Ms_Copy(Mc3)
'    Call Gf_Ms_Copy(Mc4)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) And Gf_Ms_Paste(M_CN1, Mc2) And Gf_Ms_Paste(M_CN1, Mc3) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       ' Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     End If

End Sub

Public Sub Form_Ref()

    Dim sQuery_Rt As String
    
    Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
           
           Case "R"
              If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
           Case "B"
              If Gf_Ms_Refer(M_CN1, Mc3, Mc3("pControl"), Mc3("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
           Case "C"
              If Gf_Ms_Refer(M_CN1, Mc2, Mc2("pControl"), Mc2("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
           Case "P"
              If Gf_Ms_Refer(M_CN1, Mc4, Mc4("pControl"), Mc4("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
     End Select

End Sub

Public Sub Form_Pro()

    Dim sMesg As String
    
    TXT_ROLL_IN_EMP = sUserID

    Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
       
       Case "R"
          If Not Gp_DateCheck(UTP_ROLL_IN_TIME) Then
              sMesg = " 请正确输入轧辊入库时间 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
          End If
       Case "B"
          If Not Gp_DateCheck(UTP_B_ROLL_IN_TIME) Then
              sMesg = " 请正确输入轴承座入库时间 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc3, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
          End If
       Case "C"
          If Not Gp_DateCheck(UTP_C_ROLL_IN_TIME) Then
              sMesg = " 请正确输入轴承入库时间 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
       Case "P"
          If Not Gp_DateCheck(UTP_P_ROLL_IN_TIME) Then
              sMesg = " 请正确输入护板入库时间 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc4, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
    End Select

End Sub

Public Sub Form_Del()

    Select Case Mid(CBO_ROLL_NO, 1, 1)
       
       Case "R"
          If Not Gf_Ms_Del(M_CN1, Mc1) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
       Case "B"
          If Not Gf_Ms_Del(M_CN1, Mc3) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
       Case "C"
          If Not Gf_Ms_Del(M_CN1, Mc2) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
       Case "P"
          If Not Gf_Ms_Del(M_CN1, Mc4) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
    End Select

End Sub

Private Sub sc1_Click(Value As Integer)
    
    If sc1.Value = ssCBUnchecked Then
       If Sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked Then
          sc1.Value = ssCBChecked
'          ULabel16.Caption = "轧辊号"
       End If
    Exit Sub
    End If
   
  '  If sc1.Value = -1 Then  '-1: ssCBChecked
        sc1.ForeColor = &HFF&
        Sc2.ForeColor = &H808080
        Sc2.Value = ssCBUnchecked
        sc3.ForeColor = &H808080
        sc3.Value = ssCBUnchecked
        sc4.ForeColor = &H808080
        sc4.Value = ssCBUnchecked
        SF1.Enabled = True
        sf2.Enabled = False
        sf3.Enabled = False
        SF4.Enabled = False
        ULabel16.Caption = "轧辊号"
        sQuery_load = "SELECT ROLL_NO FROM GP_ROLL3 WHERE ROLL_STATUS<>'DL'  "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
'    Else
'        sc1.Value = ssCBUnchecked
'        sc2.Value = ssCBChecked
 '   End If
   
End Sub

Private Sub sc2_Click(Value As Integer)
   
    If Sc2.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked Then
          Sc2.Value = ssCBChecked
'          ULabel16.Caption = "轴承座号"
       End If
    Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
        Sc2.ForeColor = &HFF&
        sc1.ForeColor = &H808080
        sc1.Value = ssCBUnchecked
        sc3.ForeColor = &H808080
        sc3.Value = ssCBUnchecked
        sc4.ForeColor = &H808080
        sc4.Value = ssCBUnchecked
        sf2.Enabled = True
        SF1.Enabled = False
        sf3.Enabled = False
        SF4.Enabled = False
        ULabel16.Caption = "轴承座号"
        sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK3    "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub

Private Sub sc3_Click(Value As Integer)
   
    If sc3.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And Sc2.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked Then
          sc3.Value = ssCBChecked
'          ULabel16.Caption = "轴承号"
       End If
    Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
        sc3.ForeColor = &HFF&
        sc1.ForeColor = &H808080
        sc1.Value = ssCBUnchecked
        Sc2.ForeColor = &H808080
        Sc2.Value = ssCBUnchecked
        sc4.ForeColor = &H808080
        sc4.Value = ssCBUnchecked
        sf3.Enabled = True
        SF1.Enabled = False
        sf2.Enabled = False
        SF4.Enabled = False
        ULabel16.Caption = "轴承号"
        sQuery_load = "SELECT BEARING_ID FROM GP_BEARING3    "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub
Private Sub sc4_Click(Value As Integer)
   
    If sc4.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And Sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked Then
          sc4.Value = ssCBChecked
'          ULabel16.Caption = "轴承号"
       End If
    Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
        sc4.ForeColor = &HFF&
        sc1.ForeColor = &H808080
        sc1.Value = ssCBUnchecked
        Sc2.ForeColor = &H808080
        Sc2.Value = ssCBUnchecked
        sc3.ForeColor = &H808080
        sc3.Value = ssCBUnchecked
        SF4.Enabled = True
        SF1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = False
        ULabel16.Caption = "护板号"
        sQuery_load = "SELECT PLANK_NO FROM GP_PLANK3    "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub



Private Sub TXT_C_ROLL_MOT_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0024"
        DD.rControl.Add Item:=TXT_C_ROLL_MOT


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub


Private Sub CBO_ROLL_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0003"
        DD.rControl.Add Item:=CBO_ROLL_NO


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub



Private Sub TXT_ROLL_MATERIAL_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0004"
        DD.rControl.Add Item:=TXT_ROLL_MATERIAL


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub

Private Sub UTP_B_ROLL_USE_TIME_DblClick()

    UTP_B_ROLL_USE_TIME.RawData = Gf_DTSet(M_CN1, "I")
    
End Sub

Private Sub UTP_ROLL_IN_TIME_DblClick()

    UTP_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub
Private Sub SDB_B_PLANK_DATE_DblClick()

    SDB_B_PLANK_DATE.RawData = Gf_DTSet(M_CN1, "I")
    
End Sub

Private Sub UTP_B_ROLL_IN_TIME_DblClick()

    UTP_B_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub

Private Sub UTP_C_ROLL_IN_TIME_DblClick()

    UTP_C_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub
Private Sub UTP_P_ROLL_IN_TIME_DblClick()

    UTP_P_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub


