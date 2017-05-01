VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CGF2020C 
   Caption         =   "轧辊、轴承座和轴承的报废、查询及修改界面_CGF2020C"
   ClientHeight    =   9720
   ClientLeft      =   465
   ClientTop       =   1470
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin Threed.SSCheck sc1 
      Height          =   315
      Left            =   450
      TabIndex        =   25
      Top             =   1080
      Width           =   780
      _ExtentX        =   1376
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
      Left            =   4155
      TabIndex        =   26
      Top             =   1080
      Width           =   915
      _ExtentX        =   1614
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
      Left            =   7860
      TabIndex        =   27
      Top             =   1080
      Width           =   780
      _ExtentX        =   1376
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
      Caption         =   "轴承"
   End
   Begin Threed.SSFrame sf3 
      Height          =   7725
      Left            =   7620
      TabIndex        =   21
      Top             =   1230
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_C_ROLL_DISUSE_RES 
         Height          =   315
         Left            =   1575
         MaxLength       =   2
         TabIndex        =   34
         Top             =   1005
         Width           =   1215
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
         Left            =   1935
         TabIndex        =   14
         Top             =   7155
         Width           =   1335
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
         Left            =   1215
         TabIndex        =   13
         Top             =   7155
         Width           =   705
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
         Left            =   495
         TabIndex        =   12
         Top             =   7155
         Width           =   705
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   255
         Top             =   405
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "报废时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   255
         Top             =   1005
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "报废原因"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   495
         Top             =   6825
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   1215
         Top             =   6825
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
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   1935
         Top             =   6825
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
      Begin CSTextLibCtl.sitxEdit TXT_UTP_C_ROLL_DISUSE_TIME 
         Height          =   315
         Left            =   1575
         TabIndex        =   23
         Top             =   405
         Width           =   1860
         _Version        =   262145
         _ExtentX        =   3281
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
   Begin Threed.SSFrame sf2 
      Height          =   7725
      Left            =   3900
      TabIndex        =   20
      Top             =   1230
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_B_ROLL_DISUSE_RES 
         Height          =   315
         Left            =   1575
         MaxLength       =   2
         TabIndex        =   33
         Top             =   1005
         Width           =   1215
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
         Left            =   1935
         TabIndex        =   11
         Top             =   7155
         Width           =   1335
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
         Left            =   1215
         TabIndex        =   10
         Top             =   7155
         Width           =   705
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
         Left            =   495
         TabIndex        =   9
         Top             =   7155
         Width           =   705
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   255
         Top             =   405
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "报废时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   255
         Top             =   1005
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "报废原因"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   495
         Top             =   6825
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   1215
         Top             =   6825
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   1935
         Top             =   6825
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
      Begin CSTextLibCtl.sitxEdit TXT_UTP_B_ROLL_DISUSE_TIME 
         Height          =   315
         Left            =   1575
         TabIndex        =   24
         Top             =   405
         Width           =   1860
         _Version        =   262145
         _ExtentX        =   3281
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
      Left            =   180
      TabIndex        =   16
      Top             =   1230
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_treat_mtd 
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
         Left            =   3090
         TabIndex        =   48
         Tag             =   "供货商"
         Top             =   6270
         Visible         =   0   'False
         Width           =   645
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
         Left            =   1410
         TabIndex        =   44
         Tag             =   "材质代码"
         Top             =   4845
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
         Left            =   1410
         TabIndex        =   43
         Tag             =   "供货商"
         Top             =   4290
         Width           =   1215
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
         Height          =   315
         Left            =   1410
         MaxLength       =   25
         TabIndex        =   42
         Tag             =   "材质代码"
         Top             =   5400
         Width           =   2295
      End
      Begin VB.TextBox TXT_ROLL_DISUSE_RES 
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
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   32
         Top             =   2070
         Width           =   1215
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
         Left            =   1935
         TabIndex        =   30
         Top             =   7155
         Width           =   1335
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
         Left            =   1215
         TabIndex        =   29
         Top             =   7155
         Width           =   705
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
         Left            =   495
         TabIndex        =   28
         Top             =   7155
         Width           =   705
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_DISUSE_DIA 
         Height          =   315
         Left            =   1410
         TabIndex        =   4
         Tag             =   "报废辊身直径"
         Top             =   960
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
         NumIntDigits    =   6
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_IN_AVE_HARD 
         Height          =   315
         Left            =   1410
         TabIndex        =   5
         Top             =   1515
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
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_ROLL_USE_NUM 
         Height          =   315
         Left            =   1410
         TabIndex        =   6
         Top             =   2625
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
         NumIntDigits    =   4
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_TOT_MILL_WGT 
         Height          =   315
         Left            =   1410
         TabIndex        =   7
         Top             =   3180
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
         NumIntDigits    =   8
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_TOT_MILL_LEN 
         Height          =   315
         Left            =   1410
         TabIndex        =   8
         Top             =   3735
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
         NumIntDigits    =   5
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   105
         Top             =   960
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "报废辊身直径"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   105
         Top             =   1515
         Width           =   1290
         _ExtentX        =   2275
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   105
         Top             =   2070
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "报废原因"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   105
         Top             =   3735
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "轧制公里数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   105
         Top             =   2625
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "使用次数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   105
         Top             =   3180
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "轧制吨位"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sitxEdit TXT_UTP_ROLL_DISUSE_TIME 
         Height          =   315
         Left            =   1410
         TabIndex        =   22
         Tag             =   "报废时间"
         Top             =   405
         Width           =   1815
         _Version        =   262145
         _ExtentX        =   3201
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
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   495
         Top             =   6825
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
         Left            =   1215
         Top             =   6825
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
         Left            =   1935
         Top             =   6825
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
      Begin InDate.ULabel ULabel24 
         Height          =   315
         Left            =   105
         Top             =   4290
         Width           =   1290
         _ExtentX        =   2275
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
         Left            =   105
         Top             =   4845
         Width           =   1290
         _ExtentX        =   2275
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
      Begin InDate.ULabel ULabel41 
         Height          =   315
         Left            =   105
         Top             =   5400
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "制造编号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel48 
         Height          =   315
         Left            =   285
         Top             =   5880
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         Caption         =   "报废处理方式"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin Threed.SSCheck SSC 
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   45
         Top             =   6330
         Width           =   810
         _ExtentX        =   1429
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
         Caption         =   "修复"
      End
      Begin Threed.SSCheck SSC 
         Height          =   315
         Index           =   2
         Left            =   2370
         TabIndex        =   47
         Top             =   6330
         Width           =   810
         _ExtentX        =   1429
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
         Caption         =   "其他"
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   105
         Top             =   405
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         Caption         =   "报废时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin Threed.SSCheck SSC 
         Height          =   315
         Index           =   1
         Left            =   1110
         TabIndex        =   46
         Top             =   6330
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
         Caption         =   "常规处理"
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "t"
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
         Left            =   2385
         TabIndex        =   19
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         TabIndex        =   18
         Top             =   1005
         Width           =   375
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "km"
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
         Left            =   2520
         TabIndex        =   17
         Top             =   3795
         Width           =   375
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   795
      Left            =   180
      TabIndex        =   15
      Top             =   150
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   1402
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_EMP_CD 
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
         Left            =   11355
         MaxLength       =   7
         TabIndex        =   31
         Tag             =   "作业人员"
         Top             =   240
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
         ItemData        =   "CGF2020C.frx":0000
         Left            =   7230
         List            =   "CGF2020C.frx":000D
         TabIndex        =   2
         Tag             =   "班次"
         Top             =   240
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
         ItemData        =   "CGF2020C.frx":001A
         Left            =   9180
         List            =   "CGF2020C.frx":002A
         TabIndex        =   3
         Tag             =   "班别"
         Top             =   240
         Width           =   735
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
         ItemData        =   "CGF2020C.frx":003A
         Left            =   4815
         List            =   "CGF2020C.frx":0041
         TabIndex        =   1
         Tag             =   "工厂代码"
         Text            =   "C3"
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox CB0_ROLL_ID 
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
         Left            =   1560
         TabIndex        =   0
         Tag             =   "轧辊号"
         Top             =   240
         Width           =   1365
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   480
         Top             =   240
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
         Left            =   3735
         Top             =   240
         Width           =   1065
         _ExtentX        =   1879
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
         Left            =   6345
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   8310
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
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
         Left            =   10275
         Top             =   240
         Width           =   1065
         _ExtentX        =   1879
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
   Begin Threed.SSCheck sc4 
      Height          =   315
      Left            =   11610
      TabIndex        =   35
      Top             =   1080
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
   Begin Threed.SSFrame sf4 
      Height          =   7725
      Left            =   11340
      TabIndex        =   36
      Top             =   1230
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   13626
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_P_ROLL_DISUSE_RES 
         Height          =   315
         Left            =   1590
         MaxLength       =   2
         TabIndex        =   41
         Top             =   990
         Width           =   1215
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
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   7155
         Width           =   705
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
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   7155
         Width           =   705
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
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   7155
         Width           =   1335
      End
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   270
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "报废时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   480
         Top             =   6825
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
         Left            =   1185
         Top             =   6825
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
         Left            =   1890
         Top             =   6825
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
      Begin CSTextLibCtl.sitxEdit TXT_UTP_P_ROLL_DISUSE_TIME 
         Height          =   315
         Left            =   1605
         TabIndex        =   40
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   270
         Top             =   990
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         Caption         =   "报废原因"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
End
Attribute VB_Name = "CGF2020C"
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
'-- Program Name      轧辊、轴承座和轴承的报废、查询及修改界面
'-- Program ID        CGF2020C
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
Public sQuery_Rt As String        'Active Form sQuery Setting

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
               Call Gp_Ms_Collection(CB0_ROLL_ID, "p", "n", "m", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   Call Gp_Ms_Collection(CBO_PLT, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                 Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(TXT_EMP_CD, " ", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(TXT_UTP_ROLL_DISUSE_TIME, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_ROLL_DISUSE_DIA, " ", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDB_ROLL_IN_AVE_HARD, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_ROLL_DISUSE_RES, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_ROLL_USE_NUM, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_TOT_MILL_WGT, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_TOT_MILL_LEN, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_R_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_R_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_R_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_ROLL_MAKER, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_ROLL_MATERIAL, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(TXT_MAKER_NO, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_treat_mtd, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
               Call Gp_Ms_Collection(CB0_ROLL_ID, "p", "n", "m", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'                  Call Gp_Ms_Collection(TXT_PLT, " ", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                 Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                 Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                Call Gp_Ms_Collection(TXT_EMP_CD, " ", "n", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
Call Gp_Ms_Collection(TXT_UTP_B_ROLL_DISUSE_TIME, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(TXT_B_ROLL_DISUSE_RES, " ", " ", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
'          Call Gp_Ms_Collection(CBO_B_ROLL_MATERIAL, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(TXT_B_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
               Call Gp_Ms_Collection(TXT_B_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
              Call Gp_Ms_Collection(TXT_B_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           
               Call Gp_Ms_Collection(CB0_ROLL_ID, "p", "n", "m", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
'                  Call Gp_Ms_Collection(TXT_PLT, " ", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
                 Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
                 Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
                Call Gp_Ms_Collection(TXT_EMP_CD, " ", "n", " ", "i", " ", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
Call Gp_Ms_Collection(TXT_UTP_C_ROLL_DISUSE_TIME, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
     Call Gp_Ms_Collection(TXT_C_ROLL_DISUSE_RES, " ", " ", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
'                  Call Gp_Ms_Collection(SDB_C_ROLL_WID, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
               Call Gp_Ms_Collection(TXT_C_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
               Call Gp_Ms_Collection(TXT_C_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              Call Gp_Ms_Collection(TXT_C_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
              
               Call Gp_Ms_Collection(CB0_ROLL_ID, "p", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
                 Call Gp_Ms_Collection(CBO_SHIFT, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
                 Call Gp_Ms_Collection(CBO_GROUP, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
                Call Gp_Ms_Collection(TXT_EMP_CD, " ", "n", " ", "i", " ", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
Call Gp_Ms_Collection(TXT_UTP_P_ROLL_DISUSE_TIME, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(TXT_P_ROLL_DISUSE_RES, " ", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
               Call Gp_Ms_Collection(TXT_P_SHIFT, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
               Call Gp_Ms_Collection(TXT_P_GROUP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
              Call Gp_Ms_Collection(TXT_P_IN_EMP, " ", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
              

    'MASTER Collection
     Mc1.Add Item:="CGF2020C.P_MODIFY1", Key:="P-M"
     Mc1.Add Item:="CGF2020C.P_REFER1", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     Mc2.Add Item:="CGF2020C.P_MODIFY2", Key:="P-M"
     Mc2.Add Item:="CGF2020C.P_REFER2", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
     
     Mc3.Add Item:="CGF2020C.P_MODIFY3", Key:="P-M"
     Mc3.Add Item:="CGF2020C.P_REFER3", Key:="P-R"
     Mc3.Add Item:=pControl3, Key:="pControl"
     Mc3.Add Item:=nControl3, Key:="nControl"
     Mc3.Add Item:=mControl3, Key:="mControl"
     Mc3.Add Item:=iControl3, Key:="iControl"
     Mc3.Add Item:=rControl3, Key:="rControl"
     Mc3.Add Item:=cControl3, Key:="cControl"
     Mc3.Add Item:=aControl3, Key:="aControl"
     Mc3.Add Item:=lControl3, Key:="lControl"
     
     Mc4.Add Item:="CGF2020C.P_MODIFY4", Key:="P-M"
     Mc4.Add Item:="CGF2020C.P_REFER4", Key:="P-R"
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

'Private Sub CB0_ROLL_ID_Change()
'
'    Dim sMesg As String
'
'    If Len(Trim(CB0_ROLL_ID.Text)) = 1 Then
'        Select Case Mid(Trim(CB0_ROLL_ID.Text), 1, 1)
'
'          Case "R"
'               sc1.Value = ssCBChecked
'               sc1.ForeColor = &HFF&
'               sc2.ForeColor = &H808080
'               sc2.Value = ssCBUnchecked
'               sc3.ForeColor = &H808080
'               sc3.Value = ssCBUnchecked
'               sf1.Enabled = True
'               sf2.Enabled = False
'               sf3.Enabled = False
'               ULabel16.Caption = "轧辊号"
'               sQuery_load = "SELECT ROLL_NO FROM GP_ROLL  "
'               Call Gf_ComboAdd(M_CN1, CB0_ROLL_ID, sQuery_load)
'
'          Case "C"
'               sc2.Value = ssCBChecked
'               sc2.ForeColor = &HFF&
'               sc1.ForeColor = &H808080
'               sc1.Value = ssCBUnchecked
'               sc3.ForeColor = &H808080
'               sc3.Value = ssCBUnchecked
'               sf2.Enabled = True
'               sf1.Enabled = False
'               sf3.Enabled = False
'               ULabel16.Caption = "轴承座号"
'               sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK  "
'               Call Gf_ComboAdd(M_CN1, CB0_ROLL_ID, sQuery_load)
'
'          Case "B"
'               sc3.Value = ssCBChecked
'               sc3.ForeColor = &HFF&
'               sc1.ForeColor = &H808080
'               sc1.Value = ssCBUnchecked
'               sc2.ForeColor = &H808080
'               sc2.Value = ssCBUnchecked
'               sf3.Enabled = True
'               sf1.Enabled = False
'               sf2.Enabled = False
'               ULabel16.Caption = "轴承号"
'               sQuery_load = "SELECT BEARING_ID FROM GP_BEARING  "
'               Call Gf_ComboAdd(M_CN1, CB0_ROLL_ID, sQuery_load)
'
'          Case Else
'               sMesg = " Must input 'R' or 'B' or 'C' "
'               Call Gp_MsgBoxDisplay(sMesg, "Q")
'        End Select
'    End If
'
'    If Len(Trim(CB0_ROLL_ID.Text)) > 7 Then
'        sMesg = sMesg + "Len Must < 7   "
'        Call Gp_MsgBoxDisplay(sMesg, "Q")
'    End If
'
'End Sub
'
Private Sub CB0_ROLL_ID_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0003"
        DD.rControl.Add Item:=CB0_ROLL_ID
        
        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

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

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    TXT_EMP_CD = sUserID ' + ":" + sUsername
    CBO_PLT.ListIndex = 0
   
    If CBO_SHIFT.Text <> "1" Or CBO_SHIFT.Text <> "2" Or CBO_SHIFT.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet3(M_CN1)
       End If
       CBO_SHIFT.Text = sShiftSet
    End If
'
    CB0_ROLL_ID.Clear
    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL3 WHERE ROLL_STATUS<>'DL' ORDER BY SUBSTR(ROLL_NO,1,1) DESC, SUBSTR(ROLL_NO,6,2) "
    Call Gf_ComboAdd(M_CN1, CB0_ROLL_ID, sQuery_load)
    
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
    sf4.Enabled = False

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
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    sc1.ForeColor = &HFF&
    sc1.Value = ssCBChecked
    sc2.ForeColor = &H808080
    sc2.Value = ssCBUnchecked
    sc3.ForeColor = &H808080
    sc3.Value = ssCBUnchecked
    sf1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
    
    SSC(0).Value = 0
    SSC(1).Value = 0
    SSC(2).Value = 0
    
    SSC(0).ForeColor = &H80000012
    SSC(1).ForeColor = &H80000012
    SSC(2).ForeColor = &H80000012

    ULabel16.Caption = "轧辊号"
    CB0_ROLL_ID.Clear
    sQuery_load = "SELECT ROLL_NO FROM GP_ROLL3 WHERE ROLL_STATUS<>'DL' ORDER BY SUBSTR(ROLL_NO,1,1) DESC, SUBSTR(ROLL_NO,6,2) "
    Call Gf_ComboAdd(M_CN1, CB0_ROLL_ID, sQuery_load)
    
    TXT_EMP_CD = sUserID ' + ":" + sUsername
    CBO_PLT.ListIndex = 0
   
    If CBO_SHIFT.Text <> "1" Or CBO_SHIFT.Text <> "2" Or CBO_SHIFT.Text <> "3" Then
       If sShiftSet = "" Or sShiftSet = "0" Then
          sShiftSet = Gf_ShiftSet3(M_CN1)
       End If
       CBO_SHIFT.Text = sShiftSet
    End If

End Sub

Public Sub Master_Cpy()
'
'    Call Gf_Ms_Copy(Mc1)
'    Call Gf_Ms_Copy(Mc2)
'    Call Gf_Ms_Copy(Mc3)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) And Gf_Ms_Paste(M_CN1, Mc2) And Gf_Ms_Paste(M_CN1, Mc3) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       ' Call Gp_Ms_ControlLock(Mc1("pControl"), False)
     End If

End Sub

Public Sub Form_Ref()
Dim i As Integer

    Select Case Mid(Trim(CB0_ROLL_ID.Text), 1, 1)
           
           Case "J"
              If Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("pControl")) Then
                
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 Call Gp_Ms_ControlLock(Mc1("pControl"), True)
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
     End Select
    
End Sub

Public Sub Form_Pro()

    Dim sMesg As String
    Dim i As Integer
    
    TXT_EMP_CD = sUserID
    
    
     If sc1.ForeColor = &HFF& Then
       
       If Mid(CB0_ROLL_ID.Text, 1, 1) <> "J" And Mid(CB0_ROLL_ID.Text, 1, 1) <> "C" Then
          sMesg = " 请输入正确的轧辊号 ！"
          Call Gp_MsgBoxDisplay(sMesg)
          Exit Sub
       End If
    End If
    
'    If sc2.ForeColor = &HFF& Then
'
'       If Mid(CB0_ROLL_ID.Text, 1, 1) <> "C" Then
'          sMesg = " 请输入正确的轴承座号 ！"
'          Call Gp_MsgBoxDisplay(sMesg)
'          Exit Sub
'       End If
'    End If
    
    
    If sc3.ForeColor = &HFF& Then
       
       If Mid(CB0_ROLL_ID.Text, 1, 1) <> "B" Then
          sMesg = " 请输入正确的轴承号 ！"
          Call Gp_MsgBoxDisplay(sMesg)
          Exit Sub
       End If
    End If
    
    
    If sc4.ForeColor = &HFF& Then
       
       If Mid(CB0_ROLL_ID.Text, 1, 1) <> "P" Then
          sMesg = " 请输入正确的护板号 ！"
          Call Gp_MsgBoxDisplay(sMesg)
          Exit Sub
       End If
    End If

    Select Case Mid(Trim(CB0_ROLL_ID.Text), 1, 1)
       
       Case "J"
       
          sMesg = "您确定要报废轧辊" + CB0_ROLL_ID.Text + "吗？"
          If Not Gf_MessConfirm(sMesg, "Q") Then Exit Sub
          
          If Not Gp_DateCheck(TXT_UTP_ROLL_DISUSE_TIME) Then
              sMesg = " 请正确输入轧辊报废时间 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          
          If SDB_ROLL_DISUSE_DIA.Value = 0 Then
              sMesg = " 请输入报废辊身直径 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          
          txt_treat_mtd.Text = ""
          For i = 0 To 2
              If SSC(i).Value = -1 And i <> 2 Then
                 txt_treat_mtd.Text = txt_treat_mtd.Text & CStr(i + 1)
              ElseIf SSC(i).Value = -1 And i = 2 Then
                 txt_treat_mtd.Text = txt_treat_mtd.Text & "9"
              End If
          Next

          If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
          
       Case "C"

          sMesg = "您确定要报废轧辊" + CB0_ROLL_ID.Text + "吗？"
          If Not Gf_MessConfirm(sMesg, "Q") Then Exit Sub


          If Not Gp_DateCheck(TXT_UTP_ROLL_DISUSE_TIME) Then
              sMesg = " 请正确输入轧辊报废时间 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          
          If SDB_ROLL_DISUSE_DIA.Value = 0 Then
              sMesg = " 请输入报废辊身直径 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          
          txt_treat_mtd.Text = ""
          For i = 0 To 2
              If SSC(i).Value = -1 And i <> 2 Then
                 txt_treat_mtd.Text = txt_treat_mtd.Text & CStr(i + 1)
              ElseIf SSC(i).Value = -1 And i = 2 Then
                 txt_treat_mtd.Text = txt_treat_mtd.Text & "9"
              End If
          Next
          
          If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'             Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
      
       Case "B"
       
          sMesg = "您确定要报废轴承座" + CB0_ROLL_ID.Text + "吗？"
          If Not Gf_MessConfirm(sMesg, "Q") Then Exit Sub
       
          If Not Gp_DateCheck(TXT_UTP_C_ROLL_DISUSE_TIME) Then
              sMesg = " 请正确输入轴承报废时间 ！"
              Call Gp_MsgBoxDisplay(sMesg)
              Exit Sub
          End If
          If Gf_Ms_Process(M_CN1, Mc3, sAuthority) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
          End If
       
       Case "P"
       
          sMesg = "您确定要报废轧辊" + CB0_ROLL_ID.Text + "吗？"
          If Not Gf_MessConfirm(sMesg, "Q") Then Exit Sub
          
          
          If Not Gp_DateCheck(TXT_UTP_P_ROLL_DISUSE_TIME) Then
              sMesg = " 请正确输入护板报废时间 ！"
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

'    Select Case Mid(Trim(CB0_ROLL_ID.Text), 1, 1)
'
'       Case "R"
'          If Not Gf_Ms_Del(M_CN1, Mc1) Then
'             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
''              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
'          End If
'       Case "B"
'          If Not Gf_Ms_Del(M_CN1, Mc3) Then
'             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
''              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
'          End If
'       Case "C"
'          If Not Gf_Ms_Del(M_CN1, Mc2) Then
'             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
''              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
'          End If
'       Case "P"
'          If Not Gf_Ms_Del(M_CN1, Mc4) Then
'             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
''              Call Gp_Ms_ControlLock(Mc1("pControl"), True)
'          End If
'    End Select

End Sub

Private Sub sc1_Click(Value As Integer)

   CB0_ROLL_ID.Enabled = True
   Call Gp_Ms_Cls(Mc1("rControl"))
    
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
        sQuery_load = "SELECT ROLL_NO FROM GP_ROLL3 WHERE ROLL_STATUS<>'DL' ORDER BY SUBSTR(ROLL_NO,1,1) DESC, SUBSTR(ROLL_NO,6,2) "
        Call Gf_ComboAdd(M_CN1, CB0_ROLL_ID, sQuery_load)
'    Else
'        sc1.Value = ssCBUnchecked
'        sc2.Value = ssCBChecked
 '   End If
   
End Sub

Private Sub sc2_Click(Value As Integer)

    CB0_ROLL_ID.Enabled = True
    Call Gp_Ms_Cls(Mc2("rControl"))
   
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
        sQuery_load = "SELECT CHOCK_ID FROM GP_CHOCK3    "
        Call Gf_ComboAdd(M_CN1, CB0_ROLL_ID, sQuery_load)

        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub

Private Sub sc3_Click(Value As Integer)
       
    CB0_ROLL_ID.Enabled = True
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
        sf4.Enabled = False
        ULabel16.Caption = "轴承号"
        sQuery_load = "SELECT BEARING_ID FROM GP_BEARING3    "
        Call Gf_ComboAdd(M_CN1, CB0_ROLL_ID, sQuery_load)
        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub
Private Sub sc4_Click(Value As Integer)

    CB0_ROLL_ID.Enabled = True
   Call Gp_Ms_Cls(Mc4("rControl"))
   
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
        sQuery_load = "SELECT PLANK_NO FROM GP_PLANK3   "
        Call Gf_ComboAdd(M_CN1, CB0_ROLL_ID, sQuery_load)
        
'    Else
'        sc2.Value = ssCBUnchecked
'        sc1.Value = ssCBChecked
  '  End If

End Sub

Private Sub SSC_Click(Index As Integer, Value As Integer)
If SSC(Index).Value = -1 Then
   SSC(Index).ForeColor = &HFF&       'red
Else
   SSC(Index).ForeColor = &H80000012  'black
End If
End Sub

Private Sub TXT_B_ROLL_DISUSE_RES_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0008"
        DD.rControl.Add Item:=TXT_B_ROLL_DISUSE_RES

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub

Private Sub TXT_C_ROLL_DISUSE_RES_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0008"
        DD.rControl.Add Item:=TXT_C_ROLL_DISUSE_RES

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub
Private Sub TXT_P_ROLL_DISUSE_RES_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0022"
        DD.rControl.Add Item:=TXT_P_ROLL_DISUSE_RES

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub

Private Sub TXT_ROLL_DISUSE_RES_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0008"
        DD.rControl.Add Item:=TXT_ROLL_DISUSE_RES

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
    
End Sub

Private Sub TXT_UTP_B_ROLL_DISUSE_TIME_LostFocus()

    With TXT_UTP_B_ROLL_DISUSE_TIME
        
        If Val(Mid(.Text, 1, 4)) < 2000 Or Val(Mid(.Text, 1, 4)) > 2050 Or Val(Mid(.Text, 6, 2)) > 12 Or Val(Mid(.Text, 9, 2)) > 31 Or _
            Val(Mid(.Text, 12, 2)) > 24 Or Val(Mid(.Text, 15, 2)) > 60 Or Val(Mid(.Text, 18, 2)) > 60 Then
            Call Gp_MsgBoxDisplay("请正确输入日期时间")
            Exit Sub
        End If
        
    End With
    
End Sub
Private Sub TXT_UTP_P_ROLL_DISUSE_TIME_LostFocus()

    With TXT_UTP_P_ROLL_DISUSE_TIME
        
        If Val(Mid(.Text, 1, 4)) < 2000 Or Val(Mid(.Text, 1, 4)) > 2050 Or Val(Mid(.Text, 6, 2)) > 12 Or Val(Mid(.Text, 9, 2)) > 31 Or _
            Val(Mid(.Text, 12, 2)) > 24 Or Val(Mid(.Text, 15, 2)) > 60 Or Val(Mid(.Text, 18, 2)) > 60 Then
            Call Gp_MsgBoxDisplay("请正确输入日期时间")
            Exit Sub
        End If
        
    End With
    
End Sub
Private Sub TXT_UTP_C_ROLL_DISUSE_TIME_LostFocus()
    
    With TXT_UTP_C_ROLL_DISUSE_TIME
        
        If Val(Mid(.Text, 1, 4)) < 2000 Or Val(Mid(.Text, 1, 4)) > 2050 Or Val(Mid(.Text, 6, 2)) > 12 Or Val(Mid(.Text, 9, 2)) > 31 Or _
            Val(Mid(.Text, 12, 2)) > 24 Or Val(Mid(.Text, 15, 2)) > 60 Or Val(Mid(.Text, 18, 2)) > 60 Then
            Call Gp_MsgBoxDisplay("请正确输入日期时间")
            Exit Sub
        End If
        
    End With
    
End Sub

Private Sub TXT_UTP_ROLL_DISUSE_TIME_DblClick()

    TXT_UTP_ROLL_DISUSE_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub

Private Sub TXT_UTP_B_ROLL_DISUSE_TIME_DblClick()

    TXT_UTP_B_ROLL_DISUSE_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub

Private Sub TXT_UTP_C_ROLL_DISUSE_TIME_DblClick()

    TXT_UTP_C_ROLL_DISUSE_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub

Private Sub TXT_UTP_P_ROLL_DISUSE_TIME_DblClick()

    TXT_UTP_P_ROLL_DISUSE_TIME.RawData = Gf_DTSet(M_CN1, "I")

End Sub

Private Sub TXT_UTP_ROLL_DISUSE_TIME_LostFocus()

    With TXT_UTP_ROLL_DISUSE_TIME
        
        If Val(Mid(.Text, 1, 4)) < 2000 Or Val(Mid(.Text, 1, 4)) > 2050 Or Val(Mid(.Text, 6, 2)) > 12 Or Val(Mid(.Text, 9, 2)) > 31 Or _
            Val(Mid(.Text, 12, 2)) > 24 Or Val(Mid(.Text, 15, 2)) > 60 Or Val(Mid(.Text, 18, 2)) > 60 Then
            Call Gp_MsgBoxDisplay("请正确输入日期时间")
            Exit Sub
        End If
        
    End With

End Sub


