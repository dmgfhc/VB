VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AGF2031C 
   Caption         =   "工作辊轴承(座)保养管理界面_AGF2031C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin Threed.SSCheck sc4 
      Height          =   315
      Left            =   11160
      TabIndex        =   17
      Top             =   960
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
      TabIndex        =   18
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
      Left            =   1800
      TabIndex        =   19
      Top             =   960
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
      Caption         =   "轴承座维护"
   End
   Begin Threed.SSCheck sc3 
      Height          =   315
      Left            =   6360
      TabIndex        =   20
      Top             =   960
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
      Left            =   5040
      TabIndex        =   21
      Top             =   1080
      Width           =   4815
      _ExtentX        =   8493
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         Left            =   1620
         TabIndex        =   12
         Tag             =   "供货商"
         Top             =   1344
         Width           =   1815
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
         Left            =   1620
         TabIndex        =   11
         Tag             =   "轴承标识号"
         Top             =   852
         Width           =   1815
      End
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   240
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel24 
         Height          =   315
         Left            =   240
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Index           =   0
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit UTP_C_ROLL_IN_TIME 
         Height          =   315
         Left            =   1620
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
         Left            =   240
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Index           =   1
         Left            =   240
         Top             =   1920
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "轴承转向时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit txt_y_gb_time 
         Height          =   315
         Left            =   1620
         TabIndex        =   13
         Tag             =   "入库时间"
         Top             =   1920
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __:__:__"
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
      End
   End
   Begin Threed.SSFrame sf2 
      Height          =   7725
      Left            =   240
      TabIndex        =   25
      Top             =   1080
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chk_a_n_slag 
         Caption         =   "否"
         Height          =   195
         Left            =   2010
         TabIndex        =   57
         Top             =   4050
         Width           =   555
      End
      Begin VB.CheckBox chk_a_y_slag 
         Caption         =   "是"
         Height          =   195
         Left            =   2010
         TabIndex        =   56
         Top             =   3810
         Width           =   555
      End
      Begin VB.TextBox txt_a_slag 
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
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3870
         Width           =   360
      End
      Begin VB.CheckBox chk_b_n_slag 
         Caption         =   "否"
         Height          =   195
         Left            =   2010
         TabIndex        =   55
         Top             =   3450
         Width           =   555
      End
      Begin VB.CheckBox chk_b_y_slag 
         Caption         =   "是"
         Height          =   195
         Left            =   2010
         TabIndex        =   54
         Top             =   3210
         Width           =   555
      End
      Begin VB.TextBox txt_b_slag 
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
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3255
         Width           =   360
      End
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   6
         Tag             =   "轴承座标识号"
         Top             =   840
         Width           =   1815
      End
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   240
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "加油时间"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   240
         Top             =   1380
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "加油量"
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
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
         TabIndex        =   5
         Tag             =   "加油时间"
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   240
         Top             =   855
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_B_ROLL_KG 
         Height          =   315
         Left            =   1605
         TabIndex        =   7
         Tag             =   "加油量"
         Top             =   1380
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   1
         Left            =   240
         Top             =   3255
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "水封更换"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Index           =   1
         Left            =   240
         Top             =   3855
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "油封更换"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit SDB_B_IN_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   58
         Tag             =   "内径"
         Top             =   1860
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_B_OUT_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   59
         Tag             =   "外径"
         Top             =   2730
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   240
         Top             =   1860
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "内径"
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   240
         Top             =   2745
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "不带衬板尺寸"
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
      Begin CSTextLibCtl.sidbEdit SDB_B_BT_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   60
         Tag             =   "外径"
         Top             =   2280
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel61 
         Height          =   315
         Left            =   240
         Top             =   2310
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "带衬板尺寸"
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
      Begin VB.Label Label1 
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
         Left            =   2820
         TabIndex        =   63
         Top             =   1875
         Width           =   375
      End
      Begin VB.Label Label2 
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
         Left            =   2820
         TabIndex        =   62
         Top             =   2790
         Width           =   375
      End
      Begin VB.Label Label3 
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
         Left            =   2820
         TabIndex        =   61
         Top             =   2310
         Width           =   375
      End
   End
   Begin Threed.SSFrame sf1 
      Height          =   7725
      Left            =   210
      TabIndex        =   29
      Top             =   1080
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
         Tag             =   "轧辊标识号"
         Top             =   852
         Width           =   1215
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   36
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
         TabIndex        =   37
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
         TabIndex        =   38
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
         TabIndex        =   39
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
         TabIndex        =   40
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
         TabIndex        =   41
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
         TabIndex        =   42
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
         TabIndex        =   43
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
            Size            =   9.75
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Index           =   0
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Index           =   0
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
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
         TabIndex        =   44
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
            Size            =   9.75
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   3390
         Width           =   375
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   735
      Left            =   210
      TabIndex        =   49
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
         ItemData        =   "AGF2031C.frx":0000
         Left            =   9630
         List            =   "AGF2031C.frx":0010
         TabIndex        =   3
         Tag             =   "班别"
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
         ItemData        =   "AGF2031C.frx":0020
         Left            =   7575
         List            =   "AGF2031C.frx":002D
         TabIndex        =   2
         Tag             =   "班次"
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
         TabIndex        =   4
         Tag             =   "作业人员"
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
         ItemData        =   "AGF2031C.frx":003A
         Left            =   5040
         List            =   "AGF2031C.frx":0041
         TabIndex        =   1
         Tag             =   "工厂代码"
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
         ItemData        =   "AGF2031C.frx":0049
         Left            =   1590
         List            =   "AGF2031C.frx":004B
         TabIndex        =   0
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
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
      Left            =   9840
      TabIndex        =   50
      Top             =   1080
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.ComboBox cbo_plank_mend 
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
         ItemData        =   "AGF2031C.frx":004D
         Left            =   1680
         List            =   "AGF2031C.frx":0057
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
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
         Left            =   1680
         TabIndex        =   15
         Tag             =   "轴承标识号"
         Top             =   852
         Width           =   1815
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   6945
         Width           =   1335
      End
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   240
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit UTP_P_ROLL_IN_TIME 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
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
         Left            =   240
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel26 
         Height          =   315
         Left            =   240
         Top             =   1320
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "修理方式"
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
   End
End
Attribute VB_Name = "AGF2031C"
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
'-- Program ID        AGF2010C
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


Dim pControl5 As New Collection     'Master Primary Key Collection
Dim nControl5 As New Collection     'Master Necessary Collection
Dim mControl5 As New Collection     'Master Maxlength check Collection
Dim iControl5 As New Collection     'Master Insert Collection
Dim rControl5 As New Collection     'Master Refer Collection
Dim cControl5 As New Collection     'Master Copy Collection
Dim aControl5 As New Collection     'Master -> Spread Collection
Dim lControl5 As New Collection     'Master Lock Collection




Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim Mc4 As New Collection           'Master Collection
Dim Mc5 As New Collection           'Master Collection



Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CBO_PLT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(cbo_group, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(UTP_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ROLL_NO, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_ROLL_DIA, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_ROLL_SHLD_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_ROLL_NECK_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_ROLL_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_W_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_C_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_D_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(SDB_ROLL_IN_AVE_HARD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_ROLL_MATERIAL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_R_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_R_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_R_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

          Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(cbo_group, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(UTP_B_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_CHOCK_NO, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(SDB_B_ROLL_KG, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(SDB_B_IN_DIA, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(SDB_B_BT_DIA, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(SDB_B_OUT_DIA, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(txt_b_slag, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(txt_a_slag, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_B_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_B_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_B_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           
           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
'              Call Gp_Ms_Collection(TXT_PLT, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             Call Gp_Ms_Collection(cbo_group, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
       Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(UTP_C_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_BEAR_NO, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(TXT_C_ROLL_MOT, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
         Call Gp_Ms_Collection(txt_y_gb_time, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_C_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_C_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          Call Gp_Ms_Collection(TXT_C_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          
           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(cbo_group, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
       Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(UTP_P_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(TXT_PLANK_NO, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
        Call Gp_Ms_Collection(cbo_plank_mend, " ", " ", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
           Call Gp_Ms_Collection(TXT_P_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
           Call Gp_Ms_Collection(TXT_P_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(TXT_P_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)




'    Call Gp_Ms_Collection(CBO_PIAR_ROLL_NO, "p", " ", " ", " ", "r", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
'       Call Gp_Ms_Collection(TXT_CHOCK_NO1, " ", " ", " ", " ", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
'     Call Gp_Clear_Collection(chk_b_y_slag, "s", sControl)
'     Call Gp_Clear_Collection(chk_b_n_slag, "s", sControl)
'     Call Gp_Clear_Collection(chk_a_y_slag, "s", sControl)
'     Call Gp_Clear_Collection(chk_a_n_slag, "s", sControl)
'
'     MC.Add Item:=sControl, Key:="sControl"

    'MASTER Collection
     Mc1.Add Item:="AGF2031C.P_MODIFY4", Key:="P-M"
     Mc1.Add Item:="AGF2031C.P_REFER4", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     Mc2.Add Item:="AGF2031C.P_MODIFY1", Key:="P-M"
     Mc2.Add Item:="AGF2031C.P_REFER1", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
     
     Mc3.Add Item:="AGF2031C.P_MODIFY2", Key:="P-M"
     Mc3.Add Item:="AGF2031C.P_REFER2", Key:="P-R"
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"
     
     Mc4.Add Item:="AGF2031C.P_MODIFY3", Key:="P-M"
     Mc4.Add Item:="AGF2031C.P_REFER3", Key:="P-R"
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
'Private Sub CBO_ROLL_NO_Change()
'
'   Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
'
'
'
'    Case "B"
'              If Gf_Ms_Refer(M_CN1, Mc3, Mc3("pControl"), Mc3("pControl")) Then
'
'              Else
'                 TXT_BEAR_NO.Text = ""
'
'
'                  End If
'
'
'    Case "C"
'               If Gf_Ms_Refer(M_CN1, Mc2, Mc2("pControl"), Mc2("pControl")) Then
'                 Else
'                 TXT_CHOCK_NO.Text = ""
'              End If
'
'
'    Case "P"
'              If Gf_Ms_Refer(M_CN1, Mc4, Mc4("pControl"), Mc4("pControl")) Then
'               Else
'                 TXT_PLANK_NO.Text = ""
'
'              End If
'
'
'  End Select
'
'
'
'
''End Sub
'
Private Sub CBO_ROLL_NO_Click()
   
  Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
           Case "B"
                TXT_BEAR_NO.Text = Gf_FloatFind(M_CN1, "SELECT BEARING_INPUT_NO FROM GP_BEARING WHERE BEARING_ID = '" + CBO_ROLL_NO.Text + "'")
           Case "C"
                TXT_CHOCK_NO.Text = Gf_FloatFind(M_CN1, "SELECT CHOCK_INPUT_NO FROM GP_CHOCK WHERE CHOCK_ID = '" + CBO_ROLL_NO.Text + "'")
           Case "P"
                TXT_PLANK_NO.Text = Gf_FloatFind(M_CN1, "SELECT PLANK_INPUT_NO FROM GP_PLANK WHERE PLANK_NO = '" + CBO_ROLL_NO.Text + "'")
  End Select

End Sub

Private Sub chk_b_y_slag_Click()
    If chk_b_y_slag.Value = ssCBUnchecked Then
        If chk_b_n_slag.Value = ssCBUnchecked Then
           txt_b_slag.Text = " "
        End If
        Exit Sub
    End If
    
    txt_b_slag.Text = "Y"
    
    chk_b_y_slag.ForeColor = &HFF&
    chk_b_y_slag.Value = ssCBChecked
    
    chk_b_n_slag.ForeColor = &H808080
    chk_b_n_slag.Value = ssCBUnchecked
End Sub
Private Sub chk_b_n_slag_Click()
    If chk_b_n_slag.Value = ssCBUnchecked Then
        If chk_b_y_slag.Value = ssCBUnchecked Then
           txt_b_slag.Text = " "
        End If
        Exit Sub
    End If
    
    txt_b_slag.Text = "N"
    
    chk_b_n_slag.ForeColor = &HFF&
    chk_b_n_slag.Value = ssCBChecked
    
    chk_b_y_slag.ForeColor = &H808080
    chk_b_y_slag.Value = ssCBUnchecked
End Sub

Private Sub chk_a_y_slag_Click()
    If chk_a_y_slag.Value = ssCBUnchecked Then
        If chk_a_n_slag.Value = ssCBUnchecked Then
           txt_a_slag.Text = " "
        End If
        Exit Sub
    End If
    
    txt_a_slag.Text = "Y"
    
    chk_a_y_slag.ForeColor = &HFF&
    chk_a_y_slag.Value = ssCBChecked
    
    chk_a_n_slag.ForeColor = &H808080
    chk_a_n_slag.Value = ssCBUnchecked
End Sub
Private Sub chk_a_n_slag_Click()
    If chk_a_n_slag.Value = ssCBUnchecked Then
        If chk_a_y_slag.Value = ssCBUnchecked Then
           txt_a_slag.Text = " "
        End If
        Exit Sub
    End If
    
    txt_a_slag.Text = "N"
    
    chk_a_n_slag.ForeColor = &HFF&
    chk_a_n_slag.Value = ssCBChecked
    
    chk_a_y_slag.ForeColor = &H808080
    chk_a_y_slag.Value = ssCBUnchecked
End Sub
'Private Sub CBO_ROLL_NO_Change()
'
'    Dim sMesg As String
'
'    If Len(Trim(CBO_ROLL_NO.Text)) = 1 Then
'       Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
'
'          Case "R"
'               sc1.Value = ssCBChecked
'               sc1.ForeColor = &HFF&
'               sc2.ForeColor = &H808080
'               sc2.Value = ssCBUnchecked
'               sc3.ForeColor = &H808080
'               sc3.Value = ssCBUnchecked
'               SF1.Enabled = True
'               sf2.Enabled = False
'               sf3.Enabled = False
'               ULabel16.Caption = "轧辊号"
'               sQuery_load = "SELECT ROLL_NO FROM GP_ROLL WHERE ROLL_STATUS<>'DL'  "
'               Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
'
'          Case "C"
'               sc2.Value = ssCBChecked
'               sc2.ForeColor = &HFF&
'               sc1.ForeColor = &H808080
'               sc1.Value = ssCBUnchecked
'               sc3.ForeColor = &H808080
'               sc3.Value = ssCBUnchecked
'               sf2.Enabled = True
'               SF1.Enabled = False
'               sf3.Enabled = False
'               ULabel16.Caption = "轴承座号"
'               sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK    "
'               Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
'
'          Case "B"
'               sc3.Value = ssCBChecked
'               sc3.ForeColor = &HFF&
'               sc1.ForeColor = &H808080
'               sc1.Value = ssCBUnchecked
'               sc2.ForeColor = &H808080
'               sc2.Value = ssCBUnchecked
'               sf3.Enabled = True
'               SF1.Enabled = False
'               sf2.Enabled = False
'               ULabel16.Caption = "轴承号"
'               sQuery_load = "SELECT BEARING_ID FROM GP_BEARING    "
'               Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
'
'
'          Case Else
'               sMesg = " Must input 'R' or 'B' or 'C' "
'               Call Gp_MsgBoxDisplay(sMesg, "Q")
'        End Select
'    Else
'
'    End If
'
'    If Len(Trim(CBO_ROLL_NO.Text)) > 7 Then
'        sMesg = sMesg + "Len Must < 7   "
'        Call Gp_MsgBoxDisplay(sMesg, "Q")
'    End If
'
'End Sub

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
    
    TXT_ROLL_IN_EMP = sUserID
    TXT_ROLL_IN_EMP.ForeColor = &H80000011
    CBO_PLT.ListIndex = 0
   
    If CBO_SHIFT.Text <> "1" Or CBO_SHIFT.Text <> "2" Or CBO_SHIFT.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet(M_CN1)
       End If
       CBO_SHIFT.Text = sShiftSet
    End If
   
    sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK WHERE ROLL_STATUS<>'DL'  "
    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
    
    sc2.ForeColor = &HFF&
    sc1.ForeColor = &H808080
    sc3.ForeColor = &H808080
    sc4.ForeColor = &H808080
    sc2.Value = ssCBChecked
    sc1.Value = ssCBUnchecked
    sc3.Value = ssCBUnchecked
    sc4.Value = ssCBUnchecked
    sf2.Enabled = True
    sf1.Enabled = False
    sf3.Enabled = False
    sf4.Enabled = False
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
'
'    Set pControl5 = Nothing
'    Set nControl5 = Nothing
'    Set iControl5 = Nothing
'    Set rControl5 = Nothing
'    Set cControl5 = Nothing
'    Set aControl5 = Nothing
'    Set lControl5 = Nothing
'    Set mControl5 = Nothing
    
'    Set pControl6 = Nothing
'    Set nControl6 = Nothing
'    Set iControl6 = Nothing
'    Set rControl6 = Nothing
'    Set cControl6 = Nothing
'    Set aControl6 = Nothing
'    Set lControl6 = Nothing
'    Set mControl6 = Nothing
'
'    Set pControl7 = Nothing
'    Set nControl7 = Nothing
'    Set iControl7 = Nothing
'    Set rControl7 = Nothing
'    Set cControl7 = Nothing
'    Set aControl7 = Nothing
'    Set lControl7 = Nothing
'    Set mControl7 = Nothing
    
'    Set sControl = Nothing
'    Set MC = Nothing

    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Mc4 = Nothing
'    Set Mc5 = Nothing
'    Set Mc6 = Nothing
'    Set Mc7 = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
 
'    Call Gp_SSCheck_Cls(MC("sControl"))
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))
    Call Gp_Ms_Cls(Mc4("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
       UTP_B_ROLL_IN_TIME.Enabled = True
       UTP_C_ROLL_IN_TIME.Enabled = True
       UTP_P_ROLL_IN_TIME.Enabled = True
       
        

'     pControl.SetFocus
'    sc1.ForeColor = &HFF&
'    sc1.Value = ssCBChecked
'    sc2.ForeColor = &H808080
'    sc2.Value = ssCBUnchecked
'    sc3.ForeColor = &H808080
'    sc3.Value = ssCBUnchecked
'    sc4.ForeColor = &H808080
'    sc4.Value = ssCBUnchecked
'    sf1.Enabled = True
'    sf2.Enabled = False
'    sf3.Enabled = False
'    sf4.Enabled = False
'    ULabel16.Caption = "轧辊号"
'    CBO_ROLL_NO.Clear
'    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL WHERE ROLL_STATUS<>'DL'  "
'    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
 

    TXT_ROLL_IN_EMP = sUserID
    TXT_ROLL_IN_EMP.ForeColor = &H80000011
    CBO_PLT.ListIndex = 0
    
    chk_b_y_slag.Value = 0
    chk_b_y_slag.ForeColor = &H80000012
    chk_b_n_slag.Value = 0
    chk_b_n_slag.ForeColor = &H80000012
    chk_a_y_slag.Value = 0
    chk_a_y_slag.ForeColor = &H80000012
    chk_a_n_slag.Value = 0
    chk_a_n_slag.ForeColor = &H80000012

    If CBO_SHIFT.Text <> "1" Or CBO_SHIFT.Text <> "2" Or CBO_SHIFT.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet(M_CN1)
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

    Dim sQuery_Rt  As String
    
'    Call Gp_SSCheck_Cls(MC("sControl"))
    
    Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
           
           Case "R"
              If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
           Case "B"
              If Gf_Ms_Refer(M_CN1, Mc3, Mc3("pControl"), Mc3("pControl")) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Gp_Ms_ControlLock(Mc1("pControl"), True)
'                rControl(1).SetFocus

               

'                If UTP_C_ROLL_IN_TIME.Enabled = True Then
'
'                UTP_C_ROLL_IN_TIME.ForeColor = &H80000011
'
'                If UTP_C_ROLL_IN_TIME.Locked = True Then
'                  UTP_C_ROLL_IN_TIME = False
'
'
           
              End If
           Case "C"
              If Gf_Ms_Refer(M_CN1, Mc2, Mc2("pControl"), Mc2("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Gp_Ms_ControlLock(Mc1("pControl"), True)
                 If txt_b_slag.Text = "Y" Then
                    chk_b_y_slag.Value = 1
                    chk_b_y_slag.ForeColor = &HFF
                 ElseIf txt_b_slag.Text = "N" Then
                    chk_b_n_slag.Value = 1
                    chk_b_n_slag.ForeColor = &HFF
                 End If

                If txt_a_slag.Text = "Y" Then
                   chk_a_y_slag.Value = 1
                   chk_a_y_slag.ForeColor = &HFF
                ElseIf txt_a_slag.Text = "N" Then
                   chk_a_n_slag.Value = 1
                   chk_a_n_slag.ForeColor = &HFF
                End If
              End If
              
              
            Case "P"
              If Gf_Ms_Refer(M_CN1, Mc4, Mc4("pControl"), Mc4("pControl")) Then
               Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
     End Select

    If TXT_ROLL_IN_EMP = "" Then
       TXT_ROLL_IN_EMP = sUserID
       TXT_ROLL_IN_EMP.ForeColor = &H80000011
    End If
End Sub

Public Sub Form_Pro()

    Dim SMESG As String
    
'    CBO_ROLL_NO.Text = Mid(CBO_ROLL_NO.Text, 1, 7)
'    If TXT_ROLL_IN_EMP = "" Then
'       MsgBox "作业人员必须输入", vbCritical, "错误提示"
'       Exit Sub
'    End If
    
    
    
    
   TXT_ROLL_IN_EMP = sUserID

    Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
       
       Case "R"
          If Not Gp_DateCheck(UTP_ROLL_IN_TIME) Then
              SMESG = " 请正确输入轧辊入库时间 ！"
              Call Gp_MsgBoxDisplay(SMESG)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           End If
           
           
       Case "B"
          If Not Gp_DateCheck(UTP_C_ROLL_IN_TIME) Then
              SMESG = " 请正确输入轴承清洗时间 ！"
              Call Gp_MsgBoxDisplay(SMESG)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc3, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
          End If
          
          
       Case "C"
          If Not Gp_DateCheck(UTP_B_ROLL_IN_TIME) Then
              SMESG = " 请正确输入轴承加油时间 ！"
              Call Gp_MsgBoxDisplay(SMESG)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
        
          
       Case "P"
          If Not Gp_DateCheck(UTP_P_ROLL_IN_TIME) Then
              SMESG = " 请正确输入护板入库时间 ！"
              Call Gp_MsgBoxDisplay(SMESG)
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
       If sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked Then
          sc1.Value = ssCBChecked
'          ULabel16.Caption = "轧辊号"
       End If
    Exit Sub
    End If
   
  '  If sc1.Value = -1 Then  '-1: ssCBChecked
        sc1.ForeColor = &HFF&
        sc2.ForeColor = &H808080
        sc2.Value = ssCBUnchecked
        sc3.ForeColor = &H808080
        sc3.Value = ssCBUnchecked
        sc4.ForeColor = &H808080
        sc4.Value = ssCBUnchecked
        sf1.Enabled = True
        sf2.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = False
        ULabel16.Caption = "轧辊号"
        sQuery_load = "SELECT ROLL_NO FROM GP_ROLL WHERE ROLL_STATUS<>'DL'ORDER BY ROLL_NO "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
'    Else
'        sc1.Value = ssCBUnchecked
'        sc2.Value = ssCBChecked
 '   End If
   
End Sub

Private Sub sc2_Click(Value As Integer)
   Call Form_Cls
    If sc2.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked Then
          sc2.Value = ssCBChecked
'          ULabel16.Caption = "轴承座号"
       End If
    Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
        sc2.ForeColor = &HFF&
        sc1.ForeColor = &H808080
        sc1.Value = ssCBUnchecked
        sc3.ForeColor = &H808080
        sc3.Value = ssCBUnchecked
        sc4.ForeColor = &H808080
        sc4.Value = ssCBUnchecked
        sf2.Enabled = True
        sf1.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = False
        ULabel16.Caption = "轴承座号"
        sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK  WHERE CHOCK_ID LIKE 'C2%'  "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub

Private Sub sc3_Click(Value As Integer)
Call Form_Cls
    If sc3.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc2.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked Then
          sc3.Value = ssCBChecked
'          ULabel16.Caption = "轴承号"
       End If
    Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
        sc3.ForeColor = &HFF&
        sc1.ForeColor = &H808080
        sc1.Value = ssCBUnchecked
        sc2.ForeColor = &H808080
        sc2.Value = ssCBUnchecked
        sc4.ForeColor = &H808080
        sc4.Value = ssCBUnchecked
        sf3.Enabled = True
        sf1.Enabled = False
        sf2.Enabled = False
        sf4.Enabled = False
        ULabel16.Caption = "轴承号"
        sQuery_load = "SELECT BEARING_ID FROM GP_BEARING WHERE BEARING_ID LIKE 'B2%'  "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub
Private Sub sc4_Click(Value As Integer)
  Call Form_Cls
    If sc4.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked Then
          sc4.Value = ssCBChecked
'          ULabel16.Caption = "轴承号"
       End If
    Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
        sc4.ForeColor = &HFF&
        sc1.ForeColor = &H808080
        sc1.Value = ssCBUnchecked
        sc2.ForeColor = &H808080
        sc2.Value = ssCBUnchecked
        sc3.ForeColor = &H808080
        sc3.Value = ssCBUnchecked
        sf4.Enabled = True
        sf1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = False
        ULabel16.Caption = "护板号"
        sQuery_load = "SELECT PLANK_NO FROM GP_PLANK    "
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



Private Sub UTP_ROLL_IN_TIME_DblClick()

    UTP_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub

Private Sub UTP_B_ROLL_IN_TIME_DblClick()

    UTP_B_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub

Private Sub UTP_C_ROLL_IN_TIME_DblClick()

    UTP_C_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")
     

End Sub
Private Sub txt_y_gb_time_DblClick()

    txt_y_gb_time.RawData = Gf_DTSet(M_CN1, "I")

End Sub
Private Sub UTP_P_ROLL_IN_TIME_DblClick()

    UTP_P_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub


