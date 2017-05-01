VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CGF2010C 
   Caption         =   "轧辊、轴承座和轴承的入库、查询及修改界面_CGF2010C"
   ClientHeight    =   10230
   ClientLeft      =   150
   ClientTop       =   1635
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   14880
   WindowState     =   2  'Maximized
   Begin Threed.SSCheck sc4 
      Height          =   300
      Left            =   11850
      TabIndex        =   63
      Top             =   855
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   529
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
      Height          =   300
      Left            =   405
      TabIndex        =   52
      Top             =   855
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   529
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
      Height          =   300
      Left            =   4620
      TabIndex        =   53
      Top             =   855
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   529
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
      Height          =   300
      Left            =   8190
      TabIndex        =   54
      Top             =   855
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   529
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
      Caption         =   "轴承"
   End
   Begin Threed.SSFrame sf3 
      Height          =   8400
      Left            =   7920
      TabIndex        =   47
      Top             =   870
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   14817
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
         MaxLength       =   6
         TabIndex        =   56
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
         MaxLength       =   2
         TabIndex        =   26
         Tag             =   "供货商"
         Top             =   1344
         Width           =   1215
      End
      Begin VB.TextBox TXT_C_SHIFT 
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
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   7980
         Width           =   705
      End
      Begin VB.TextBox TXT_C_GROUP 
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
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   7980
         Width           =   705
      End
      Begin VB.TextBox TXT_C_IN_EMP 
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
         TabIndex        =   38
         Top             =   8010
         Width           =   1335
      End
      Begin CSTextLibCtl.sidbEdit SDB_C_IN_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   27
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
         TabIndex        =   28
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
         TabIndex        =   29
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
         Left            =   270
         Top             =   7650
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
         Top             =   7650
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
         Top             =   7650
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
         TabIndex        =   25
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
         Top             =   2865
         Width           =   330
      End
   End
   Begin Threed.SSFrame sf2 
      Height          =   8400
      Left            =   4350
      TabIndex        =   44
      Top             =   870
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   14817
      _Version        =   196609
      BackColor       =   14737632
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
         MaxLength       =   6
         TabIndex        =   55
         Tag             =   "轴承座标识号"
         Top             =   852
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
         MaxLength       =   2
         TabIndex        =   21
         Tag             =   "供货商"
         Top             =   1344
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
         MaxLength       =   1
         TabIndex        =   24
         Tag             =   "材质代码"
         Top             =   2820
         Width           =   1215
      End
      Begin VB.TextBox TXT_B_SHIFT 
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
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   7980
         Width           =   705
      End
      Begin VB.TextBox TXT_B_GROUP 
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   7980
         Width           =   705
      End
      Begin VB.TextBox TXT_B_IN_EMP 
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
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   7980
         Width           =   1335
      End
      Begin CSTextLibCtl.sidbEdit SDB_B_IN_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   22
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
      Begin CSTextLibCtl.sidbEdit SDB_B_OUT_DIA 
         Height          =   315
         Left            =   1605
         TabIndex        =   23
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
      Begin InDate.ULabel ULabel20 
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
      Begin InDate.ULabel ULabel21 
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   270
         Top             =   2820
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
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   225
         Top             =   7650
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
         Top             =   7650
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
         Top             =   7650
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
         TabIndex        =   20
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel42 
         Height          =   315
         Left            =   270
         Top             =   3330
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
         TabIndex        =   65
         Tag             =   "入库时间"
         Top             =   3330
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
         TabIndex        =   46
         Top             =   2370
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
         TabIndex        =   45
         Top             =   1845
         Width           =   375
      End
   End
   Begin Threed.SSFrame sf1 
      Height          =   8400
      Left            =   120
      TabIndex        =   32
      Top             =   870
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   14817
      _Version        =   196609
      BackColor       =   14737632
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
         Height          =   300
         Left            =   1740
         MaxLength       =   12
         TabIndex        =   19
         Tag             =   "料号"
         Top             =   5835
         Width           =   1800
      End
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
         Height          =   300
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "领料单号"
         Top             =   5535
         Width           =   1800
      End
      Begin VB.TextBox txt_sec_treat_mtd 
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
         Left            =   3075
         TabIndex        =   70
         Tag             =   "供货商"
         Top             =   7845
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox TXT_LOC 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "材质代码"
         Top             =   5235
         Width           =   1425
      End
      Begin VB.TextBox TXT_MAKER_NO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1740
         MaxLength       =   25
         TabIndex        =   16
         Tag             =   "制造编号"
         Top             =   4935
         Width           =   2325
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
         Height          =   300
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   71
         Tag             =   "轧辊标识号"
         Top             =   690
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
         Height          =   300
         Left            =   1740
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "供货商"
         Top             =   990
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
         Height          =   300
         Left            =   1740
         MaxLength       =   1
         TabIndex        =   15
         Tag             =   "材质代码"
         Top             =   4635
         Width           =   1215
      End
      Begin VB.TextBox TXT_R_SHIFT 
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
         TabIndex        =   30
         Top             =   7980
         Width           =   720
      End
      Begin VB.TextBox TXT_R_GROUP 
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
         Left            =   1020
         TabIndex        =   31
         Top             =   7980
         Width           =   720
      End
      Begin VB.TextBox TXT_R_IN_EMP 
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
         Left            =   1755
         TabIndex        =   51
         Top             =   7980
         Width           =   1350
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_DIA 
         Height          =   300
         Left            =   1740
         TabIndex        =   7
         Tag             =   "入库辊径"
         Top             =   1290
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_SHLD_DIA 
         Height          =   300
         Left            =   1740
         TabIndex        =   8
         Tag             =   "辊肩直径"
         Top             =   1890
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_NECK_DIA 
         Height          =   300
         Left            =   1740
         TabIndex        =   9
         Tag             =   "辊颈直径"
         Top             =   2190
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_WGT 
         Height          =   300
         Left            =   1740
         TabIndex        =   10
         Tag             =   "轧辊重量"
         Top             =   2490
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         MaxValue        =   999999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_IN_W_HARD 
         Height          =   300
         Left            =   1740
         TabIndex        =   11
         Tag             =   "工作侧硬度"
         Top             =   2790
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         MaxValue        =   99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_IN_C_HARD 
         Height          =   315
         Left            =   1740
         TabIndex        =   12
         Tag             =   "中部硬度"
         Top             =   3390
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
         Left            =   1740
         TabIndex        =   13
         Tag             =   "驱动侧硬度"
         Top             =   3705
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
         Height          =   300
         Left            =   1740
         TabIndex        =   14
         Tag             =   "平均硬度"
         Top             =   4335
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         MaxValue        =   99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel7 
         Height          =   300
         Left            =   285
         Top             =   390
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         Height          =   300
         Left            =   285
         Top             =   990
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         Height          =   300
         Left            =   285
         Top             =   1290
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         Height          =   300
         Left            =   285
         Top             =   1890
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         Height          =   300
         Left            =   285
         Top             =   2790
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "传动侧硬度1"
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
         Height          =   300
         Left            =   285
         Top             =   3390
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         Height          =   300
         Left            =   285
         Top             =   2190
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         Height          =   300
         Left            =   285
         Top             =   2490
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         Left            =   285
         Top             =   3690
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "操作侧硬度1"
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
         Height          =   300
         Left            =   285
         Top             =   4320
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         Height          =   300
         Left            =   285
         Top             =   4620
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
         Left            =   285
         Top             =   7650
         Width           =   720
         _ExtentX        =   1270
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
         Left            =   1020
         Top             =   7650
         Width           =   720
         _ExtentX        =   1270
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
         Left            =   1755
         Top             =   7650
         Width           =   1350
         _ExtentX        =   2381
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
         Height          =   300
         Left            =   1740
         TabIndex        =   5
         Tag             =   "入库时间"
         Top             =   390
         Width           =   1800
         _Version        =   262145
         _ExtentX        =   3175
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         Mask            =   "____-__-__ __:__"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel3 
         Height          =   300
         Left            =   285
         Top             =   690
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
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
      Begin InDate.ULabel ULabel41 
         Height          =   300
         Left            =   285
         Top             =   4920
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "制造编号"
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
      Begin InDate.ULabel ULabel43 
         Height          =   300
         Left            =   285
         Top             =   5220
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "辊架位置"
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
      Begin InDate.ULabel ULabel48 
         Height          =   330
         Left            =   285
         Top             =   6765
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   582
         Caption         =   "二次处理方式"
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
      Begin Threed.SSCheck SSC 
         Height          =   315
         Index           =   0
         Left            =   330
         TabIndex        =   66
         Top             =   7065
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
         Caption         =   "激光修复"
      End
      Begin Threed.SSCheck SSC 
         Height          =   315
         Index           =   1
         Left            =   1650
         TabIndex        =   67
         Top             =   7065
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "堆焊处理"
      End
      Begin Threed.SSCheck SSC 
         Height          =   315
         Index           =   2
         Left            =   330
         TabIndex        =   68
         Top             =   7350
         Width           =   1095
         _ExtentX        =   1931
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
         Caption         =   "激光强化"
      End
      Begin Threed.SSCheck SSC 
         Height          =   315
         Index           =   3
         Left            =   1650
         TabIndex        =   69
         Top             =   7350
         Width           =   1290
         _ExtentX        =   2275
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
         Caption         =   "车削加工"
      End
      Begin InDate.ULabel ULabel49 
         Height          =   300
         Left            =   285
         Top             =   5520
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "领料单号"
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
      Begin InDate.ULabel ULabel50 
         Height          =   300
         Left            =   285
         Top             =   5820
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "料号"
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
      Begin InDate.ULabel ULabel51 
         Height          =   300
         Left            =   285
         Top             =   6120
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "单价"
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
      Begin InDate.ULabel ULabel52 
         Height          =   300
         Left            =   285
         Top             =   6435
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "限位辊径"
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
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_DIA_Y 
         Height          =   300
         Left            =   1740
         TabIndex        =   73
         Top             =   1590
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   5
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel53 
         Height          =   300
         Left            =   285
         Top             =   1590
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "当前辊径"
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
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_W_DIA2 
         Height          =   300
         Left            =   1740
         TabIndex        =   75
         Top             =   3090
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         MaxValue        =   99
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel54 
         Height          =   300
         Left            =   285
         Top             =   3090
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Caption         =   "传动侧硬度2"
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
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_D_DIA2 
         Height          =   315
         Left            =   1740
         TabIndex        =   76
         Top             =   4020
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
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel55 
         Height          =   315
         Left            =   285
         Top             =   4005
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         Caption         =   "操作侧硬度2"
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
      Begin CSTextLibCtl.sidbEdit txt_PLAN_DIA 
         Height          =   315
         Left            =   1740
         TabIndex        =   77
         Tag             =   "限位辊径"
         Top             =   6450
         Width           =   1380
         _Version        =   262145
         _ExtentX        =   2434
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
      Begin CSTextLibCtl.sidbEdit txt_ROLL_PRICE 
         Height          =   315
         Left            =   1740
         TabIndex        =   78
         Tag             =   "轧辊重量"
         Top             =   6150
         Width           =   1380
         _Version        =   262145
         _ExtentX        =   2434
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
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   2940
         TabIndex        =   74
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "元(RMB)"
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
         Left            =   3135
         TabIndex        =   72
         Top             =   6225
         Width           =   720
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
         Height          =   300
         Left            =   2940
         TabIndex        =   43
         Top             =   2550
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
         Height          =   300
         Left            =   2925
         TabIndex        =   42
         Top             =   1935
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
         Height          =   300
         Left            =   2940
         TabIndex        =   41
         Top             =   1335
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
         Height          =   300
         Left            =   2925
         TabIndex        =   40
         Top             =   2220
         Width           =   375
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   735
      Left            =   120
      TabIndex        =   39
      Top             =   135
      Width           =   15015
      _ExtentX        =   26485
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
         Left            =   1590
         TabIndex        =   0
         Tag             =   "ROLL_NO"
         Top             =   195
         Width           =   1365
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
         ItemData        =   "CGF2010C.frx":0000
         Left            =   5040
         List            =   "CGF2010C.frx":0007
         TabIndex        =   1
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
         Left            =   12810
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   4
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
         ItemData        =   "CGF2010C.frx":000F
         Left            =   7575
         List            =   "CGF2010C.frx":001C
         TabIndex        =   2
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
         ItemData        =   "CGF2010C.frx":0029
         Left            =   10140
         List            =   "CGF2010C.frx":0039
         TabIndex        =   3
         Top             =   195
         Width           =   735
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
         Left            =   9150
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
         Left            =   11700
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
      Height          =   8400
      Left            =   11580
      TabIndex        =   57
      Top             =   870
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   14817
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
         MaxLength       =   2
         TabIndex        =   64
         Tag             =   "供货商"
         Top             =   1350
         Width           =   1215
      End
      Begin VB.TextBox TXT_P_IN_EMP 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   7980
         Width           =   1335
      End
      Begin VB.TextBox TXT_P_GROUP 
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
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   7980
         Width           =   705
      End
      Begin VB.TextBox TXT_P_SHIFT 
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
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   7980
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
         TabIndex        =   58
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
         Top             =   7650
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
         Top             =   7650
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
         Top             =   7650
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
         TabIndex        =   62
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
Attribute VB_Name = "CGF2010C"
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
'-- Program ID        CGF2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2007.10.31
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
            Call Gp_Ms_Collection(CBO_GROUP, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(UTP_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ROLL_NO, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_ROLL_DIA, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_ROLL_DIA_Y, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_ROLL_SHLD_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDB_ROLL_NECK_DIA, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDB_ROLL_WGT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_W_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ROLL_W_DIA2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_C_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(SDB_ROLL_IN_D_HARD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ROLL_D_DIA2, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(SDB_ROLL_IN_AVE_HARD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_ROLL_MATERIAL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_MAKER_NO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_R_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_R_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_R_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_loc, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_sec_treat_mtd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     '''added by guoli at 20081229100800 for ERP
     Call Gp_Ms_Collection(txt_ISSUETALLYNO, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_MTRLNO, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_ROLL_PRICE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_PLAN_DIA, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
          Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'              Call Gp_Ms_Collection(TXT_PLT, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(CBO_GROUP, " ", " ", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(UTP_B_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(UTP_B_ROLL_USE_TIME, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_CHOCK_NO, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(TXT_B_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(SDB_B_IN_DIA, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(SDB_B_OUT_DIA, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(TXT_B_ROLL_MATERIAL, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_B_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(TXT_B_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(TXT_B_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           
           Call Gp_Ms_Collection(CBO_ROLL_NO, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
'              Call Gp_Ms_Collection(TXT_PLT, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
             Call Gp_Ms_Collection(CBO_GROUP, " ", " ", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
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
             Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(CBO_GROUP, " ", " ", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
       Call Gp_Ms_Collection(TXT_ROLL_IN_EMP, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(UTP_P_ROLL_IN_TIME, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(TXT_PLANK_NO, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
      Call Gp_Ms_Collection(TXT_P_ROLL_MAKER, " ", " ", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
           Call Gp_Ms_Collection(TXT_P_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
           Call Gp_Ms_Collection(TXT_P_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
          Call Gp_Ms_Collection(TXT_P_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)

    'MASTER Collection
     Mc1.Add Item:="CGF2010C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="CGF2010C.P_REFER1", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     Mc2.Add Item:="CGF2010C.P_MODIFY2", Key:="P-M"
     Mc2.Add Item:="CGF2010C.P_REFER2", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
     
     Mc3.Add Item:="CGF2010C.P_MODIFY3", Key:="P-M"
     Mc3.Add Item:="CGF2010C.P_REFER3", Key:="P-R"
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"
     
     Mc4.Add Item:="CGF2010C.P_MODIFY4", Key:="P-M"
     Mc4.Add Item:="CGF2010C.P_REFER4", Key:="P-R"
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

Private Sub CBO_ROLL_NO_Change()
    If Len(Trim(CBO_ROLL_NO.Text)) = 7 Then
       TXT_ROLL_NO.Text = Mid(CBO_ROLL_NO.Text, 6, 2)
    End If
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

    sc1.ForeColor = &HFF&
    sc2.ForeColor = &H808080
    sc3.ForeColor = &H808080
    sc4.ForeColor = &H808080
    sc1.Value = ssCBChecked
    sc2.Value = ssCBUnchecked
    sc3.Value = ssCBUnchecked
    sc4.Value = ssCBUnchecked
    sf1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
    SF4.Enabled = False
    ULabel16.Caption = "轧辊号"
    
    Screen.MousePointer = vbDefault
    
    CBO_ROLL_NO.Clear
    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL3 WHERE ROLL_STATUS<>'DL' ORDER BY SUBSTR(ROLL_NO,1,1) DESC, SUBSTR(ROLL_NO,6,2) "
    Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)


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
    sc2.ForeColor = &H808080
    sc2.Value = ssCBUnchecked
    sc3.ForeColor = &H808080
    sc3.Value = ssCBUnchecked
    sc4.ForeColor = &H808080
    sc4.Value = ssCBUnchecked
    sf1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
    SF4.Enabled = False
    
    SSC(0).Value = 0
    SSC(1).Value = 0
    SSC(2).Value = 0
    SSC(3).Value = 0
    
    SSC(0).ForeColor = &H80000012
    SSC(1).ForeColor = &H80000012
    SSC(2).ForeColor = &H80000012
    SSC(3).ForeColor = &H80000012
    
    ULabel16.Caption = "轧辊号"
    CBO_ROLL_NO.Clear
    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL3 WHERE ROLL_STATUS<>'DL' ORDER BY SUBSTR(ROLL_NO,1,1) DESC, SUBSTR(ROLL_NO,6,2) "
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
    Dim i As Integer
    If (Mid(Trim(CBO_ROLL_NO.Text), 1, 1) = "J") Or (Mid(Trim(CBO_ROLL_NO.Text), 1, 1) = "C") Then
              If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("pControl")) Then
                If txt_sec_treat_mtd.Text <> "" Then
                   For i = 1 To Len(txt_sec_treat_mtd.Text)
                       SSC(CInt(Mid(txt_sec_treat_mtd.Text, i, 1)) - 1).Value = -1
                       SSC(CInt(Mid(txt_sec_treat_mtd.Text, i, 1)) - 1).ForeColor = &HFF&       'red
                   Next
                End If
                 
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 'Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
    End If
    
    Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
        
           Case "B"
              If Gf_Ms_Refer(M_CN1, Mc3, Mc3("pControl"), Mc3("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
'           Case "C"
'              If Gf_Ms_Refer(M_CN1, Mc2, Mc2("pControl"), Mc2("pControl")) Then
'                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
''                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
'              End If
           Case "P"
              If Gf_Ms_Refer(M_CN1, Mc4, Mc4("pControl"), Mc4("pControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
              End If
     End Select

End Sub

Public Sub Form_Pro()

    Dim sMesg As String
    Dim i As Integer
    
    TXT_ROLL_IN_EMP = sUserID
    
    
'    If sc1.ForeColor = &HFF& Then
'
'       If Mid(CBO_ROLL_NO.Text, 1, 1) <> "J" Then
'          sMesg = " 请输入正确的轧辊号 ！"
'          Call Gp_MsgBoxDisplay(sMesg)
'          Exit Sub
'       End If
'    End If
    
    If sc2.ForeColor = &HFF& Then
       
       If Mid(CBO_ROLL_NO.Text, 1, 1) <> "C" Then
          sMesg = " 请输入正确的轴承座号 ！"
          Call Gp_MsgBoxDisplay(sMesg)
          Exit Sub
       End If
    End If
    
    
    If sc3.ForeColor = &HFF& Then
       
       If Mid(CBO_ROLL_NO.Text, 1, 1) <> "B" Then
          sMesg = " 请输入正确的轴承号 ！"
          Call Gp_MsgBoxDisplay(sMesg)
          Exit Sub
       End If
    End If
    
    
    If sc4.ForeColor = &HFF& Then
       
       If Mid(CBO_ROLL_NO.Text, 1, 1) <> "P" Then
          sMesg = " 请输入正确的护板号 ！"
          Call Gp_MsgBoxDisplay(sMesg)
          Exit Sub
       End If
    End If
    
     
    

    Select Case Mid(Trim(CBO_ROLL_NO.Text), 1, 1)
       
       Case "J"
          If Not Gp_DateCheck(UTP_ROLL_IN_TIME) Then
              sMesg = " 请正确输入轧辊入库时间 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          ''''added by guoli at 20081229 103600 for ERP'''''
          If Trim(txt_ISSUETALLYNO.Text) = "" And Trim(txt_ROLL_PRICE.Text) = "" Then
              sMesg = "领料单号和单价不能同时为空!"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          
          If SDB_ROLL_WGT.Value = 0 Then
              sMesg = "轧辊重量不能为空!"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          ''''''''''''''''''''''''''''''''''''''''''''''''''

          txt_sec_treat_mtd.Text = ""
          For i = 0 To 3
              If SSC(i).Value = -1 Then
                 txt_sec_treat_mtd.Text = txt_sec_treat_mtd.Text & CStr(i + 1)
              End If
          Next

          If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
          End If
       Case "C"
          If Not Gp_DateCheck(UTP_ROLL_IN_TIME) Then
              sMesg = " 请正确输入轧辊入库时间 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          ''''added by guoli at 20081229 103600 for ERP
          If Trim(txt_ISSUETALLYNO.Text) = "" And Trim(txt_ROLL_PRICE.Text) = "" Then
              sMesg = "领料单号和单价不能同时为空!"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If

          txt_sec_treat_mtd.Text = ""
          For i = 0 To 3
              If SSC(i).Value = -1 Then
                 txt_sec_treat_mtd.Text = txt_sec_treat_mtd.Text & CStr(i + 1)
              End If
          Next

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
       
       Case "J"
          CBO_ROLL_NO.Enabled = False
          If Not Gf_Ms_Del(M_CN1, Mc1) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'             Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
          CBO_ROLL_NO.Enabled = True
      
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

    
   CBO_ROLL_NO.Enabled = True
   Call Gp_Ms_Cls(Mc1("rControl"))
    
    If sc1.Value = ssCBUnchecked Then
       If sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked Then
          sc1.Value = ssCBChecked

       End If
    Exit Sub
    End If
   

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
        SF4.Enabled = False
        ULabel16.Caption = "轧辊号"
        sQuery_load = "SELECT ROLL_NO FROM GP_ROLL3 WHERE ROLL_STATUS<>'DL'  "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
  
   
End Sub

Private Sub sc2_Click(Value As Integer)


   CBO_ROLL_NO.Enabled = True
   Call Gp_Ms_Cls(Mc2("rControl"))
   
   
    If sc2.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked And sc4.Value = ssCBUnchecked Then
          sc2.Value = ssCBChecked

       End If
    Exit Sub
    End If
  

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
        SF4.Enabled = False
        ULabel16.Caption = "轴承座号"
        sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK3    "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)



End Sub

Private Sub sc3_Click(Value As Integer)

   
   CBO_ROLL_NO.Enabled = True
   Call Gp_Ms_Cls(Mc3("rControl"))
   
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
        SF4.Enabled = False
        ULabel16.Caption = "轴承号"
        sQuery_load = "SELECT BEARING_ID FROM GP_BEARING3    "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)
        

End Sub
Private Sub sc4_Click(Value As Integer)

   
   CBO_ROLL_NO.Enabled = True
   Call Gp_Ms_Cls(Mc4("rControl"))
   
    If sc4.Value = ssCBUnchecked Then
       If sc1.Value = ssCBUnchecked And sc2.Value = ssCBUnchecked And sc3.Value = ssCBUnchecked Then
          sc4.Value = ssCBChecked

       End If
    Exit Sub
    End If
  
        sc4.ForeColor = &HFF&
        sc1.ForeColor = &H808080
        sc1.Value = ssCBUnchecked
        sc2.ForeColor = &H808080
        sc2.Value = ssCBUnchecked
        sc3.ForeColor = &H808080
        sc3.Value = ssCBUnchecked
        SF4.Enabled = True
        sf1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = False
        ULabel16.Caption = "护板号"
        sQuery_load = "SELECT PLANK_NO FROM GP_PLANK3    "
        Call Gf_ComboAdd(M_CN1, CBO_ROLL_NO, sQuery_load)

End Sub



Private Sub SSC_Click(Index As Integer, Value As Integer)
Dim i As Integer
Dim CNT As Integer

For i = 0 To 3
    If SSC(i).Value = -1 Then
       CNT = CNT + 1
       If CNT > 3 Then
          SSC(Index).Value = 0
          MsgBox "最多只能选择3种二次处理方式！", vbCritical, "系统提示信息"
          Exit Sub
       End If
    End If
Next

If SSC(Index).Value = -1 Then
   SSC(Index).ForeColor = &HFF&       'red
Else
   SSC(Index).ForeColor = &H80000012  'black
End If

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

Private Sub UTP_B_ROLL_USE_TIME_DblClick()

    UTP_B_ROLL_USE_TIME.RawData = Gf_DTSet(M_CN1, "I")

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
