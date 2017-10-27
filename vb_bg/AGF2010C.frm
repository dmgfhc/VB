VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AGF2010C 
   Caption         =   "轧辊、轴承(座)入库实绩查询及修改界面_AGF2010C"
   ClientHeight    =   9675
   ClientLeft      =   150
   ClientTop       =   1635
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin Threed.SSCheck sc6 
      Height          =   315
      Left            =   7980
      TabIndex        =   73
      Top             =   5310
      Width           =   1635
      _ExtentX        =   2884
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
      Caption         =   "油膜轴承衬套"
   End
   Begin Threed.SSCheck sc5 
      Height          =   315
      Left            =   4200
      TabIndex        =   72
      Top             =   5310
      Width           =   1560
      _ExtentX        =   2752
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
      Caption         =   "油膜轴承锥套"
   End
   Begin Threed.SSCheck sc4 
      Height          =   315
      Left            =   11730
      TabIndex        =   57
      Top             =   855
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
      Caption         =   "护板"
   End
   Begin Threed.SSCheck sc1 
      Height          =   315
      Left            =   450
      TabIndex        =   43
      Top             =   855
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
      Left            =   4185
      TabIndex        =   44
      Top             =   855
      Width           =   990
      _ExtentX        =   1746
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
      Caption         =   "轴承座"
   End
   Begin Threed.SSCheck sc3 
      Height          =   315
      Left            =   7950
      TabIndex        =   45
      Top             =   855
      Width           =   1320
      _ExtentX        =   2328
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
      Caption         =   "工作辊轴承"
   End
   Begin Threed.SSFrame sf3 
      Height          =   4470
      Left            =   7695
      TabIndex        =   38
      Top             =   870
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   7885
      _Version        =   196609
      BackColor       =   14737632
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
         TabIndex        =   49
         Tag             =   "轴承标识号"
         Top             =   852
         Width           =   1215
      End
      Begin VB.TextBox TXT_C_ROLL_MAKER 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1605
         TabIndex        =   20
         Tag             =   "供货商"
         Top             =   1344
         Width           =   1215
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
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3930
         Width           =   705
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
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3930
         Width           =   705
      End
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
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3930
         Width           =   1335
      End
      Begin CSTextLibCtl.sidbEdit SDB_C_IN_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   21
         Tag             =   "内径"
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_C_OUT_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   22
         Tag             =   "外径"
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_C_ROLL_WID 
         Height          =   315
         Left            =   1605
         TabIndex        =   23
         Tag             =   "宽度"
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel23 
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
      Begin InDate.ULabel ULabel24 
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
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   270
         Top             =   1830
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
         Left            =   270
         Top             =   2325
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "外径"
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
         Left            =   270
         Top             =   2820
         Width           =   1305
         _ExtentX        =   2302
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
      Begin InDate.ULabel ULabel28 
         Height          =   315
         Left            =   300
         Top             =   3600
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
         Left            =   1005
         Top             =   3600
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
         Left            =   1710
         Top             =   3600
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
         TabIndex        =   19
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
      Begin VB.Label Label43 
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
         Left            =   2880
         TabIndex        =   41
         Top             =   2355
         Width           =   330
      End
      Begin VB.Label Label42 
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
         Left            =   2880
         TabIndex        =   40
         Top             =   1830
         Width           =   330
      End
      Begin VB.Label Label21 
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
         Left            =   2880
         TabIndex        =   39
         Top             =   2865
         Width           =   330
      End
   End
   Begin Threed.SSFrame sf2 
      Height          =   4470
      Left            =   3930
      TabIndex        =   35
      Top             =   870
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   7885
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
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   3960
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
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   3960
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
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   3960
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
         TabIndex        =   48
         Tag             =   "轴承座标识号"
         Top             =   795
         Width           =   1215
      End
      Begin VB.TextBox TXT_B_ROLL_MAKER 
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
         TabIndex        =   15
         Tag             =   "供货商"
         Top             =   1260
         Width           =   1215
      End
      Begin VB.TextBox TXT_B_ROLL_MATERIAL 
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
         Tag             =   "材质代码"
         Top             =   3150
         Width           =   1215
      End
      Begin CSTextLibCtl.sidbEdit SDB_B_IN_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   16
         Tag             =   "内径"
         Top             =   1740
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
         TabIndex        =   17
         Tag             =   "外径"
         Top             =   2715
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
      Begin InDate.ULabel ULabel18 
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   270
         Top             =   1260
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   270
         Top             =   1740
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
         Left            =   270
         Top             =   2715
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
         Top             =   3150
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
      Begin CSTextLibCtl.sitxEdit UTP_B_ROLL_IN_TIME 
         Height          =   315
         Left            =   1605
         TabIndex        =   14
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
      Begin InDate.ULabel ULabel37 
         Height          =   315
         Left            =   270
         Top             =   795
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
      Begin InDate.ULabel ULabel41 
         Height          =   315
         Left            =   270
         Top             =   3630
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
      Begin InDate.ULabel ULabel42 
         Height          =   315
         Left            =   990
         Top             =   3630
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
      Begin InDate.ULabel ULabel43 
         Height          =   315
         Left            =   1710
         Top             =   3630
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
      Begin CSTextLibCtl.sidbEdit SDB_B_BT_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   74
         Tag             =   "外径"
         Top             =   2250
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
         Left            =   270
         Top             =   2250
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Left            =   2850
         TabIndex        =   75
         Top             =   2295
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
         Left            =   2850
         TabIndex        =   37
         Top             =   2760
         Width           =   375
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
         Left            =   2850
         TabIndex        =   36
         Top             =   1755
         Width           =   375
      End
   End
   Begin Threed.SSFrame sf1 
      Height          =   8130
      Left            =   180
      TabIndex        =   26
      Top             =   870
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   14340
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_ISSUETALLYNO 
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
         TabIndex        =   61
         Tag             =   "领料单号"
         Top             =   5352
         Width           =   1800
      End
      Begin VB.TextBox txt_MTRLNO 
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
         TabIndex        =   60
         Tag             =   "料号"
         Top             =   5768
         Width           =   1800
      End
      Begin VB.TextBox txt_PLAN_DIA 
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
         TabIndex        =   59
         Tag             =   "限位辊径"
         Top             =   6600
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
         TabIndex        =   47
         Tag             =   "轧辊标识号"
         Top             =   776
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
         TabIndex        =   46
         Tag             =   "供货商"
         Top             =   1192
         Width           =   1215
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
         TabIndex        =   13
         Tag             =   "材质代码"
         Top             =   4936
         Width           =   1215
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
         TabIndex        =   24
         Top             =   7485
         Width           =   705
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
         TabIndex        =   25
         Top             =   7485
         Width           =   705
      End
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
         TabIndex        =   42
         Top             =   7485
         Width           =   1335
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   5
         Tag             =   "入库辊径"
         Top             =   1608
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
         TabIndex        =   6
         Tag             =   "辊肩直径"
         Top             =   2024
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
         TabIndex        =   7
         Tag             =   "辊颈直径"
         Top             =   2440
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
         TabIndex        =   8
         Tag             =   "轧辊重量"
         Top             =   2856
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
         TabIndex        =   9
         Tag             =   "工作侧硬度"
         Top             =   3272
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
         TabIndex        =   10
         Tag             =   "中部硬度"
         Top             =   3688
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
         TabIndex        =   11
         Tag             =   "驱动侧硬度"
         Top             =   4104
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
         TabIndex        =   12
         Tag             =   "平均硬度"
         Top             =   4520
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
         Top             =   1192
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
         Top             =   1608
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
         Top             =   2024
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
         Top             =   3272
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
         Top             =   3688
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
         Top             =   2440
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
         Top             =   2856
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
         Top             =   4104
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
         Top             =   4520
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
         Top             =   4936
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
         Top             =   7155
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
         Top             =   7155
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
         Top             =   7155
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
         TabIndex        =   4
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
         Top             =   776
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
      Begin InDate.ULabel ULabel49 
         Height          =   315
         Left            =   270
         Top             =   5352
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "领料单号"
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
      Begin InDate.ULabel ULabel50 
         Height          =   315
         Left            =   270
         Top             =   5768
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "料号"
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
      Begin InDate.ULabel ULabel51 
         Height          =   315
         Left            =   270
         Top             =   6184
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "单价"
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
      Begin InDate.ULabel ULabel52 
         Height          =   315
         Left            =   270
         Top             =   6600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "限位辊径"
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
      Begin CSTextLibCtl.sidbEdit txt_ROLL_PRICE 
         Height          =   315
         Left            =   1605
         TabIndex        =   84
         Tag             =   "轧辊重量"
         Top             =   6180
         Width           =   1800
         _Version        =   262145
         _ExtentX        =   3175
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         MaxValue        =   999999
         MinValue        =   0
         Undo            =   0
         Data            =   0
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
         Left            =   2790
         TabIndex        =   34
         Top             =   2910
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
         Left            =   2790
         TabIndex        =   33
         Top             =   2080
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
         Left            =   2790
         TabIndex        =   32
         Top             =   1665
         Width           =   375
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
         Left            =   2790
         TabIndex        =   31
         Top             =   2495
         Width           =   375
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   735
      Left            =   180
      TabIndex        =   30
      Top             =   135
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   1296
      _Version        =   196609
      BackColor       =   14737632
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
         ItemData        =   "AGF2010C.frx":0000
         Left            =   1890
         List            =   "AGF2010C.frx":0002
         TabIndex        =   50
         Tag             =   "ROLL_NO"
         Top             =   195
         Width           =   1485
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
         ItemData        =   "AGF2010C.frx":0004
         Left            =   5130
         List            =   "AGF2010C.frx":000B
         TabIndex        =   0
         Tag             =   "工厂代码"
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
         Left            =   12165
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "作业人员"
         Top             =   195
         Width           =   1215
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
         ItemData        =   "AGF2010C.frx":0013
         Left            =   7575
         List            =   "AGF2010C.frx":0020
         TabIndex        =   1
         Tag             =   "班次"
         Top             =   195
         Width           =   735
      End
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
         ItemData        =   "AGF2010C.frx":002D
         Left            =   9840
         List            =   "AGF2010C.frx":003D
         TabIndex        =   2
         Tag             =   "班别"
         Top             =   195
         Width           =   735
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   240
         Top             =   195
         Width           =   1635
         _ExtentX        =   2884
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
         Left            =   4020
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
         Left            =   8850
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
         Left            =   11055
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
      Height          =   4470
      Left            =   11475
      TabIndex        =   51
      Top             =   870
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   7885
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_P_ROLL_MAKER 
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
         TabIndex        =   58
         Tag             =   "供货商"
         Top             =   1350
         Width           =   1215
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
         TabIndex        =   55
         Top             =   3915
         Width           =   1335
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
         TabIndex        =   54
         Top             =   3915
         Width           =   705
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
         Top             =   3915
         Width           =   705
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
         Left            =   1605
         MaxLength       =   3
         TabIndex        =   52
         Tag             =   "轴承标识号"
         Top             =   852
         Width           =   1215
      End
      Begin InDate.ULabel ULabel39 
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
      Begin InDate.ULabel ULabel40 
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
      Begin InDate.ULabel ULabel44 
         Height          =   315
         Left            =   270
         Top             =   3585
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
         Top             =   3585
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
         Top             =   3585
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
         TabIndex        =   56
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
   Begin Threed.SSFrame sf5 
      Height          =   3660
      Left            =   3930
      TabIndex        =   65
      Top             =   5340
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   6456
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_S_ROLL_MAKER 
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
         TabIndex        =   70
         Tag             =   "供货商"
         Top             =   1830
         Width           =   1215
      End
      Begin VB.TextBox TXT_S_NO 
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
         TabIndex        =   69
         Tag             =   "轴承座标识号"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TXT_S_SHIFT 
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
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   3000
         Width           =   705
      End
      Begin VB.TextBox TXT_S_GROUP 
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
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   3000
         Width           =   705
      End
      Begin VB.TextBox TXT_S_IN_EMP 
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
         TabIndex        =   66
         Top             =   3000
         Width           =   1335
      End
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   300
         Top             =   2670
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
         Left            =   1020
         Top             =   2670
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
         Left            =   1740
         Top             =   2670
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
      Begin InDate.ULabel ULabel48 
         Height          =   315
         Left            =   285
         Top             =   570
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
      Begin CSTextLibCtl.sitxEdit UTP_S_ROLL_IN_TIME 
         Height          =   315
         Left            =   1620
         TabIndex        =   71
         Tag             =   "入库时间"
         Top             =   570
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
      Begin InDate.ULabel ULabel53 
         Height          =   315
         Left            =   285
         Top             =   1200
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "轴承序列号"
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
      Begin InDate.ULabel ULabel54 
         Height          =   315
         Left            =   285
         Top             =   1830
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
   End
   Begin Threed.SSFrame sf6 
      Height          =   3660
      Left            =   7695
      TabIndex        =   76
      Top             =   5340
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   6456
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_T_ROLL_MAKER 
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
         TabIndex        =   81
         Tag             =   "供货商"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TXT_T_NO 
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
         TabIndex        =   80
         Tag             =   "轴承座标识号"
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox TXT_T_SHIFT 
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
         Left            =   390
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   2970
         Width           =   705
      End
      Begin VB.TextBox TXT_T_GROUP 
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
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   2970
         Width           =   705
      End
      Begin VB.TextBox TXT_T_IN_EMP 
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
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   2970
         Width           =   1335
      End
      Begin InDate.ULabel ULabel55 
         Height          =   315
         Left            =   390
         Top             =   2640
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
      Begin InDate.ULabel ULabel56 
         Height          =   315
         Left            =   1110
         Top             =   2640
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
      Begin InDate.ULabel ULabel57 
         Height          =   315
         Left            =   1830
         Top             =   2640
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
      Begin InDate.ULabel ULabel58 
         Height          =   315
         Left            =   345
         Top             =   570
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
      Begin CSTextLibCtl.sitxEdit UTP_T_ROLL_IN_TIME 
         Height          =   315
         Left            =   1680
         TabIndex        =   82
         Tag             =   "入库时间"
         Top             =   570
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
      Begin InDate.ULabel ULabel59 
         Height          =   315
         Left            =   345
         Top             =   1170
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "轴承序列号"
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
      Begin InDate.ULabel ULabel60 
         Height          =   315
         Left            =   345
         Top             =   1800
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
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   3660
      Left            =   11475
      TabIndex        =   83
      Top             =   5340
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   6456
      _Version        =   196609
      BackColor       =   14737632
      Begin InDate.ULabel ULabel65 
         Height          =   315
         Left            =   330
         Top             =   570
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "轧辊号:第一位为:R "
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
      Begin InDate.ULabel ULabel66 
         Height          =   315
         Left            =   330
         Top             =   1530
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "轴承座号:第一位为:C"
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
      Begin InDate.ULabel ULabel62 
         Height          =   315
         Left            =   330
         Top             =   2010
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "工作辊轴承号:第一位为:B"
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
      Begin InDate.ULabel ULabel63 
         Height          =   315
         Left            =   330
         Top             =   1050
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "护板号:第一位为:P"
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
      Begin InDate.ULabel ULabel64 
         Height          =   315
         Left            =   330
         Top             =   2490
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "油膜轴承锥套号:第一位为:S"
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
      Begin InDate.ULabel ULabel67 
         Height          =   315
         Left            =   330
         Top             =   2940
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "油膜轴承衬套号:第一位为:T"
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
End
Attribute VB_Name = "AGF2010C"
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
'-- Program Name      轧辊、轴承座和轴承的入库、查询及修改界面
'-- Program ID        AGF2010C
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

Dim pControl6 As New Collection     'Master Primary Key Collection
Dim nControl6 As New Collection     'Master Necessary Collection
Dim mControl6 As New Collection     'Master Maxlength check Collection
Dim iControl6 As New Collection     'Master Insert Collection
Dim rControl6 As New Collection     'Master Refer Collection
Dim cControl6 As New Collection     'Master Copy Collection
Dim aControl6 As New Collection     'Master -> Spread Collection
Dim lControl6 As New Collection     'Master Lock Collection


Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim Mc4 As New Collection           'Master Collection
Dim Mc5 As New Collection           'Master Collection
Dim Mc6 As New Collection           'Master Collection


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Master"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CBO_PLT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(cbo_shift, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(UTP_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ROLL_NO, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_ROLL_DIA, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_ROLL_SHLD_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_ROLL_NECK_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_ROLL_WGT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_W_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_C_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_D_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(SDB_ROLL_IN_AVE_HARD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_ROLL_MATERIAL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_R_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_R_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_R_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     '''added by guoli at 20081229100800 for ERP
     Call Gp_Ms_Collection(txt_ISSUETALLYNO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_MTRLNO, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ROLL_PRICE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_PLAN_DIA, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
          Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'              Call Gp_Ms_Collection(TXT_PLT, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(cbo_shift, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(UTP_B_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_CHOCK_NO, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(TXT_B_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(SDB_B_IN_DIA, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(SDB_B_OUT_DIA, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(SDB_B_BT_DIA, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(TXT_B_ROLL_MATERIAL, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_B_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_B_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_B_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           
           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
'              Call Gp_Ms_Collection(TXT_PLT, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             Call Gp_Ms_Collection(cbo_shift, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
       Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(UTP_C_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_BEAR_NO, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
      Call Gp_Ms_Collection(TXT_C_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          Call Gp_Ms_Collection(SDB_C_IN_DIA, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
         Call Gp_Ms_Collection(SDB_C_OUT_DIA, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        Call Gp_Ms_Collection(SDB_C_ROLL_WID, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_C_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
           Call Gp_Ms_Collection(TXT_C_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          Call Gp_Ms_Collection(TXT_C_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
          
           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(cbo_shift, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
       Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(UTP_P_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(TXT_PLANK_NO, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
      Call Gp_Ms_Collection(TXT_P_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
           Call Gp_Ms_Collection(TXT_P_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
           Call Gp_Ms_Collection(TXT_P_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(TXT_P_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)

           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
             Call Gp_Ms_Collection(cbo_shift, " ", "n", " ", "i", " ", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
             Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
       Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
    Call Gp_Ms_Collection(UTP_S_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
              Call Gp_Ms_Collection(TXT_S_NO, " ", "n", " ", "i", "r", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
      Call Gp_Ms_Collection(TXT_S_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
           Call Gp_Ms_Collection(TXT_S_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
           Call Gp_Ms_Collection(TXT_S_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
          Call Gp_Ms_Collection(TXT_S_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)

           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
             Call Gp_Ms_Collection(cbo_shift, " ", "n", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
             Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
       Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
    Call Gp_Ms_Collection(UTP_T_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
              Call Gp_Ms_Collection(TXT_T_NO, " ", "n", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(TXT_T_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
           Call Gp_Ms_Collection(TXT_T_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
           Call Gp_Ms_Collection(TXT_T_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
          Call Gp_Ms_Collection(TXT_T_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)


    'MASTER Collection
     Mc1.Add Item:="AGF2010C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="AGF2010C.P_REFER1", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     Mc2.Add Item:="AGF2010C.P_MODIFY2", Key:="P-M"
     Mc2.Add Item:="AGF2010C.P_REFER2", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
     
     Mc3.Add Item:="AGF2010C.P_MODIFY3", Key:="P-M"
     Mc3.Add Item:="AGF2010C.P_REFER3", Key:="P-R"
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"
     
     Mc4.Add Item:="AGF2010C.P_MODIFY4", Key:="P-M"
     Mc4.Add Item:="AGF2010C.P_REFER4", Key:="P-R"
     Mc4.Add Item:=pControl4, Key:="pControl"
     Mc4.Add Item:=nControl4, Key:="nControl"
     Mc4.Add Item:=mControl4, Key:="mControl"
     Mc4.Add Item:=iControl4, Key:="iControl"
     Mc4.Add Item:=rControl4, Key:="rControl"
     Mc4.Add Item:=cControl4, Key:="cControl"
     Mc4.Add Item:=aControl4, Key:="aControl"
     Mc4.Add Item:=lControl4, Key:="lControl"
     
     Mc5.Add Item:="AGF2010C.P_MODIFY5", Key:="P-M"
     Mc5.Add Item:="AGF2010C.P_REFER5", Key:="P-R"
     Mc5.Add Item:=pControl5, Key:="pControl"
     Mc5.Add Item:=nControl5, Key:="nControl"
     Mc5.Add Item:=mControl5, Key:="mControl"
     Mc5.Add Item:=iControl5, Key:="iControl"
     Mc5.Add Item:=rControl5, Key:="rControl"
     Mc5.Add Item:=cControl5, Key:="cControl"
     Mc5.Add Item:=aControl5, Key:="aControl"
     Mc5.Add Item:=lControl5, Key:="lControl"
     
     Mc6.Add Item:="AGF2010C.P_MODIFY6", Key:="P-M"
     Mc6.Add Item:="AGF2010C.P_REFER6", Key:="P-R"
     Mc6.Add Item:=pControl6, Key:="pControl"
     Mc6.Add Item:=nControl6, Key:="nControl"
     Mc6.Add Item:=mControl6, Key:="mControl"
     Mc6.Add Item:=iControl6, Key:="iControl"
     Mc6.Add Item:=rControl6, Key:="rControl"
     Mc6.Add Item:=cControl6, Key:="cControl"
     Mc6.Add Item:=aControl6, Key:="aControl"
     Mc6.Add Item:=lControl6, Key:="lControl"
   
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
   
    If cbo_shift.Text <> "1" Or cbo_shift.Text <> "2" Or cbo_shift.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet(M_CN1)
       End If
       cbo_shift.Text = sShiftSet
    End If
   
    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL WHERE ROLL_STATUS<>'DL' ORDER BY ROLL_NO "
    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
    
    sc1.ForeColor = &HFF&
    sc2.ForeColor = &H808080
    sc3.ForeColor = &H808080
    sc4.ForeColor = &H808080
    sc5.ForeColor = &H808080
    sc6.ForeColor = &H808080
    sc1.Value = ssCBChecked
    sc2.Value = ssCBUnchecked
    sc3.Value = ssCBUnchecked
    sc4.Value = ssCBUnchecked
    sc5.Value = ssCBUnchecked
    sc6.Value = ssCBUnchecked
    sf1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
    sf4.Enabled = False
    sf5.Enabled = False
    sf6.Enabled = False
    ULabel16.Caption = "轧辊号"
    
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
    
    Set pControl5 = Nothing
    Set nControl5 = Nothing
    Set iControl5 = Nothing
    Set rControl5 = Nothing
    Set cControl5 = Nothing
    Set aControl5 = Nothing
    Set lControl5 = Nothing
    Set mControl5 = Nothing
    
    Set pControl6 = Nothing
    Set nControl6 = Nothing
    Set iControl6 = Nothing
    Set rControl6 = Nothing
    Set cControl6 = Nothing
    Set aControl6 = Nothing
    Set lControl6 = Nothing
    Set mControl6 = Nothing


    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Mc4 = Nothing
    Set Mc5 = Nothing
    Set Mc6 = Nothing
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
    Call Gp_Ms_Cls(Mc5("rControl"))
    Call Gp_Ms_Cls(Mc6("rControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    sc1.ForeColor = &HFF&
    sc1.Value = ssCBChecked
    sc2.ForeColor = &H808080
    sc2.Value = ssCBUnchecked
    sc3.ForeColor = &H808080
    sc3.Value = ssCBUnchecked
    sc4.ForeColor = &H808080
    sc4.Value = ssCBUnchecked
    sc5.ForeColor = &H808080
    sc5.Value = ssCBUnchecked
    sc6.ForeColor = &H808080
    sc6.Value = ssCBUnchecked

    sf1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
    sf4.Enabled = False
    sf5.Enabled = False
    sf6.Enabled = False

    ULabel16.Caption = "轧辊号"
    CBO_ROLL_NO.Clear
    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL WHERE ROLL_STATUS<>'DL' ORDER BY ROLL_NO "
    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

    TXT_ROLL_IN_EMP = sUserID ' + ":" + sUsername
    CBO_PLT.ListIndex = 0
   
    If cbo_shift.Text <> "1" Or cbo_shift.Text <> "2" Or cbo_shift.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet(M_CN1)
       End If
       cbo_shift.Text = sShiftSet
    End If
    
End Sub

Public Sub Master_Cpy()

'    Call Gf_Ms_Copy(Mc1)
'    Call Gf_Ms_Copy(Mc2)
'    Call Gf_Ms_Copy(Mc3)
'    Call Gf_Ms_Copy(Mc4)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) And Gf_Ms_Paste(M_CN1, Mc2) And Gf_Ms_Paste(M_CN1, Mc3) And Gf_Ms_Paste(M_CN1, Mc4) And Gf_Ms_Paste(M_CN1, Mc5) And Gf_Ms_Paste(M_CN1, Mc6) Then
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
                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
           Case "C"
              If Gf_Ms_Refer(M_CN1, Mc2, Mc2("pControl"), Mc2("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
           Case "P"
              If Gf_Ms_Refer(M_CN1, Mc4, Mc4("pControl"), Mc4("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
            Case "S"
              If Gf_Ms_Refer(M_CN1, Mc5, Mc5("pControl"), Mc5("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
           Case "T"
              If Gf_Ms_Refer(M_CN1, Mc6, Mc6("pControl"), Mc6("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
                  
     End Select

End Sub

Public Sub Form_Pro()

    Dim SMESG As String
    
    TXT_ROLL_IN_EMP = sUserID

    Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
       
       Case "R"
          If Not Gp_DateCheck(UTP_ROLL_IN_TIME) Then
              SMESG = " 请正确输入轧辊入库时间 ！"
              Call Gp_MsgBoxDisplay(SMESG)
              Exit Sub
          End If
          
          ''''added by guoli at 20081229 103600 for ERP'''''''
          If Trim(txt_ISSUETALLYNO.Text) = "" Or Trim(txt_ROLL_PRICE.Text) = "" Then
              SMESG = "领料单号和单价不能为空!"
              Call Gp_MsgBoxDisplay(SMESG)
              Exit Sub
          End If
          
          If Len(txt_ISSUETALLYNO.Text) <> 10 Then
             SMESG = "领料单号应为10位!"
             Call Gp_MsgBoxDisplay(SMESG)
             Exit Sub
          End If
          
          If SDB_ROLL_WGT.Value = 0 Then
              SMESG = "轧辊重量不能为空!"
              Call Gp_MsgBoxDisplay(SMESG)
              Exit Sub
          End If
          ''''''''''''''''''''''''''''''''''''''''''''''''''''
          
          If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
          End If
       Case "B"
          If Not Gp_DateCheck(UTP_C_ROLL_IN_TIME) Then
              SMESG = " 请正确输入轴承座入库时间 ！"
              Call Gp_MsgBoxDisplay(SMESG)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc3, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
          End If
       Case "C"
          If Not Gp_DateCheck(UTP_B_ROLL_IN_TIME) Then
              SMESG = " 请正确输入工作辊轴承入库时间 ！"
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
          
         Case "S"
          If Not Gp_DateCheck(UTP_S_ROLL_IN_TIME) Then
              SMESG = " 请正确输入油膜轴承锥套入库时间 ！"
              Call Gp_MsgBoxDisplay(SMESG)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc5, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
          
        Case "T"
          If Not Gp_DateCheck(UTP_T_ROLL_IN_TIME) Then
              SMESG = " 请正确输入油膜轴承衬套入库时间 ！"
              Call Gp_MsgBoxDisplay(SMESG)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc6, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
          
          
          
    End Select

End Sub

Public Sub Form_Del()

    If Mid(CBO_ROLL_NO, 1, 1) = "R" Then
        MsgBox "轧辊不能删除！", vbCritical, "系统提示信息"
        Exit Sub
    End If

    Select Case Mid(CBO_ROLL_NO, 1, 1)
       Case "R"
            If Not Gf_Ms_Del(M_CN1, Mc1) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            End If
       Case "B"
          If Gf_Ms_Del(M_CN1, Mc3) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
             CBO_ROLL_NO.Enabled = True
             Call Gp_Ms_Cls(Mc3("rControl"))
          End If
         
       Case "C"
          If Gf_Ms_Del(M_CN1, Mc2) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
             CBO_ROLL_NO.Enabled = True
             Call Gp_Ms_Cls(Mc2("rControl"))
          End If
        
       Case "P"
          If Gf_Ms_Del(M_CN1, Mc4) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
             CBO_ROLL_NO.Enabled = True
             Call Gp_Ms_Cls(Mc4("rControl"))
          End If
        
        Case "S"
          If Gf_Ms_Del(M_CN1, Mc5) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
             CBO_ROLL_NO.Enabled = True
             Call Gp_Ms_Cls(Mc5("rControl"))
          End If
        
        Case "T"
          If Gf_Ms_Del(M_CN1, Mc6) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
             CBO_ROLL_NO.Enabled = True
             Call Gp_Ms_Cls(Mc6("rControl"))
          End If
        
     End Select

End Sub

Private Sub sc1_Click(Value As Integer)
    
    CBO_ROLL_NO.Enabled = True
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    If sc1.Value = ssCBUnchecked Then
       If sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked And sc5.Value = ssCBUnchecked And sc6.Value = ssCBUnchecked Then
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
        sc5.ForeColor = &H808080
        sc5.Value = ssCBUnchecked
        sc6.ForeColor = &H808080
        sc6.Value = ssCBUnchecked
        sf1.Enabled = True
        sf2.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = False
        sf5.Enabled = False
        sf6.Enabled = False
        ULabel16.Caption = "轧辊号"
    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL WHERE ROLL_STATUS<>'DL' ORDER BY ROLL_NO "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
'    Else
'        sc1.Value = ssCBUnchecked
'        sc2.Value = ssCBChecked
 '   End If
   
End Sub

Private Sub sc2_Click(Value As Integer)
        
    CBO_ROLL_NO.Enabled = True
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    If sc2.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked And sc5.Value = ssCBUnchecked And sc6.Value = ssCBUnchecked Then
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
        sc5.ForeColor = &H808080
        sc5.Value = ssCBUnchecked
        sc6.ForeColor = &H808080
        sc6.Value = ssCBUnchecked
        
        sf2.Enabled = True
        sf1.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = False
        sf5.Enabled = False
        sf6.Enabled = False
        ULabel16.Caption = "轴承座号"
        sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK WHERE STATUS <> 'DL' ORDER BY CHOCK_ID  "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub

Private Sub sc3_Click(Value As Integer)
      
    CBO_ROLL_NO.Enabled = True
    Call Gp_Ms_Cls(Mc3("rControl"))
      
    If sc3.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc2.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked And sc5.Value = ssCBUnchecked And sc6.Value = ssCBUnchecked Then
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
        sc5.ForeColor = &H808080
        sc5.Value = ssCBUnchecked
        sc6.ForeColor = &H808080
        sc6.Value = ssCBUnchecked
        sf3.Enabled = True
        sf1.Enabled = False
        sf2.Enabled = False
        sf4.Enabled = False
        sf5.Enabled = False
        sf6.Enabled = False
        ULabel16.Caption = "轴承号"
        sQuery_load = "SELECT BEARING_ID FROM GP_BEARING  WHERE STATUS <> 'DL' ORDER BY BEARING_ID "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
        

End Sub
Private Sub sc4_Click(Value As Integer)
        
    CBO_ROLL_NO.Enabled = True
    Call Gp_Ms_Cls(Mc4("rControl"))
        
    If sc4.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc5.Value = ssCBUnchecked And sc6.Value = ssCBUnchecked Then
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
        sc5.ForeColor = &H808080
        sc5.Value = ssCBUnchecked
        sc6.ForeColor = &H808080
        sc6.Value = ssCBUnchecked
        sf4.Enabled = True
        sf1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = False
        sf5.Enabled = False
        sf6.Enabled = False
        ULabel16.Caption = "护板号"
        sQuery_load = "SELECT PLANK_NO FROM GP_PLANK GP_PLANK WHERE SUBSTR(PLANK_NO,1,1)= 'P'  ORDER BY PLANK_NO "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

End Sub
Private Sub sc5_Click(Value As Integer)
        
    CBO_ROLL_NO.Enabled = True
    Call Gp_Ms_Cls(Mc5("rControl"))
    
    If sc5.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked And sc6.Value = ssCBUnchecked Then
          sc5.Value = ssCBChecked
'          ULabel16.Caption = "轴承号"
       End If
    Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
        sc5.ForeColor = &HFF&
        sc1.ForeColor = &H808080
        sc1.Value = ssCBUnchecked
        sc2.ForeColor = &H808080
        sc2.Value = ssCBUnchecked
        sc3.ForeColor = &H808080
        sc3.Value = ssCBUnchecked
        sc4.ForeColor = &H808080
        sc4.Value = ssCBUnchecked
        sc6.ForeColor = &H808080
        sc6.Value = ssCBUnchecked
        sf5.Enabled = True
        sf1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = False
        sf6.Enabled = False
        ULabel16.Caption = "油膜轴承锥套号"
        sQuery_load = "SELECT PLANK_NO FROM GP_PLANK WHERE SUBSTR(PLANK_NO,1,1)= 'S' ORDER BY PLANK_NO "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

End Sub
Private Sub sc6_Click(Value As Integer)
     
    CBO_ROLL_NO.Enabled = True
    Call Gp_Ms_Cls(Mc6("rControl"))

    If sc6.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked And sc5.Value = ssCBUnchecked Then
          sc6.Value = ssCBChecked
'          ULabel16.Caption = "轴承号"
       End If
    Exit Sub
    End If
  
  '  If sc2.Value = -1 Then    '-1: ssCBChecked
        sc6.ForeColor = &HFF&
        sc1.ForeColor = &H808080
        sc1.Value = ssCBUnchecked
        sc2.ForeColor = &H808080
        sc2.Value = ssCBUnchecked
        sc3.ForeColor = &H808080
        sc3.Value = ssCBUnchecked
        sc4.ForeColor = &H808080
        sc4.Value = ssCBUnchecked
        sc5.ForeColor = &H808080
        sc5.Value = ssCBUnchecked
        sf6.Enabled = True
        sf1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = False
        sf5.Enabled = False
        ULabel16.Caption = "油膜轴承衬套号"
        sQuery_load = "SELECT PLANK_NO FROM GP_PLANK WHERE SUBSTR(PLANK_NO,1,1)= 'T' ORDER BY PLANK_NO "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
        
 End Sub

Private Sub TXT_B_ROLL_MAKER_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0005"
        DD.rControl.Add Item:=TXT_B_ROLL_MAKER


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
    
End Sub

Private Sub TXT_B_ROLL_MATERIAL_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0004"
        DD.rControl.Add Item:=TXT_B_ROLL_MATERIAL


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub

Private Sub TXT_P_ROLL_MAKER_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0023"
        DD.rControl.Add Item:=TXT_P_ROLL_MAKER


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub
Private Sub TXT_S_ROLL_MAKER_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0005"
        DD.rControl.Add Item:=TXT_S_ROLL_MAKER


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub
Private Sub TXT_T_ROLL_MAKER_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0005"
        DD.rControl.Add Item:=TXT_T_ROLL_MAKER


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub
Private Sub TXT_C_ROLL_MAKER_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0005"
        DD.rControl.Add Item:=TXT_C_ROLL_MAKER


        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub
Private Sub TXT_ROLL_MAKER_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0005"
        DD.rControl.Add Item:=TXT_ROLL_MAKER


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

Private Sub UTP_ROLL_IN_TIME_DblClick()

    UTP_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

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
Private Sub UTP_S_ROLL_IN_TIME_DblClick()

    UTP_S_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub
Private Sub UTP_T_ROLL_IN_TIME_DblClick()

    UTP_T_ROLL_IN_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub
