VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AFC2010C 
   Caption         =   "转炉实绩修改及查询界面_AFC2010C"
   ClientHeight    =   9225
   ClientLeft      =   1380
   ClientTop       =   1110
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame sf2 
      Height          =   1275
      Left            =   7770
      TabIndex        =   12
      Top             =   2460
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   2249
      _Version        =   196609
      BackColor       =   14737632
      Begin InDate.ULabel ULabel21 
         Height          =   315
         Left            =   150
         Top             =   465
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         Caption         =   "开始时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Left            =   1080
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Caption         =   "1次"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel23 
         Height          =   315
         Left            =   4815
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Caption         =   "3次"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel24 
         Height          =   315
         Left            =   2955
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Caption         =   "2次"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   150
         Top             =   810
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         Caption         =   "结束时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_61 
         Height          =   315
         Left            =   1080
         TabIndex        =   43
         Top             =   465
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_62 
         Height          =   315
         Left            =   2955
         TabIndex        =   44
         Top             =   465
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_63 
         Height          =   315
         Left            =   4815
         TabIndex        =   45
         Top             =   465
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_71 
         Height          =   315
         Left            =   1080
         TabIndex        =   46
         Top             =   810
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_72 
         Height          =   315
         Left            =   2955
         TabIndex        =   47
         Top             =   810
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_73 
         Height          =   315
         Left            =   4815
         TabIndex        =   48
         Top             =   810
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
   End
   Begin Threed.SSFrame Frame1 
      Height          =   765
      Left            =   270
      TabIndex        =   9
      Top             =   150
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   1349
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox txt_process_no 
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
         Left            =   4455
         MaxLength       =   9
         TabIndex        =   97
         Tag             =   "处理号"
         Top             =   210
         Width           =   1185
      End
      Begin VB.ComboBox cbo_group_cd 
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
         ItemData        =   "AFC2010C.frx":0000
         Left            =   11370
         List            =   "AFC2010C.frx":0002
         TabIndex        =   21
         Tag             =   "班别"
         Top             =   210
         Width           =   720
      End
      Begin VB.ComboBox cbo_shift 
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
         ItemData        =   "AFC2010C.frx":0004
         Left            =   9285
         List            =   "AFC2010C.frx":0006
         TabIndex        =   20
         Tag             =   "班次"
         Top             =   210
         Width           =   720
      End
      Begin VB.ComboBox cbo_heat_no 
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
         ItemData        =   "AFC2010C.frx":0008
         Left            =   1245
         List            =   "AFC2010C.frx":000A
         TabIndex        =   19
         Tag             =   "炉号"
         Top             =   210
         Width           =   1350
      End
      Begin VB.ComboBox cbo_prc_line 
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
         ItemData        =   "AFC2010C.frx":000C
         Left            =   7170
         List            =   "AFC2010C.frx":000E
         TabIndex        =   18
         Tag             =   "炉座号"
         Top             =   210
         Width           =   705
      End
      Begin VB.TextBox txt_emp_cd 
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
         Left            =   13410
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   17
         Tag             =   "作业人员"
         Top             =   210
         Width           =   1035
      End
      Begin VB.CommandButton cbo_down 
         Caption         =   "▼"
         Height          =   225
         Left            =   2625
         TabIndex        =   16
         Top             =   375
         Width           =   315
      End
      Begin VB.CommandButton cbo_up 
         Caption         =   "▲"
         Height          =   225
         Left            =   2625
         TabIndex        =   15
         Top             =   120
         Width           =   315
      End
      Begin InDate.ULabel ULabel63 
         Height          =   315
         Left            =   120
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "炉号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   8130
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
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
      Begin InDate.ULabel ULabel69 
         Height          =   315
         Left            =   10245
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
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
      Begin InDate.ULabel ULabel70 
         Height          =   315
         Left            =   12300
         Top             =   210
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   6030
         Top             =   210
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "炉座号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel35 
         Height          =   315
         Left            =   3270
         Top             =   210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         Caption         =   "处理号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
   Begin VB.TextBox txt_oper 
      Height          =   270
      Left            =   12030
      TabIndex        =   8
      Text            =   "1"
      Top             =   90
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt_proc 
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
      Left            =   12795
      TabIndex        =   7
      Text            =   "BC"
      Top             =   45
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   13785
      TabIndex        =   6
      Text            =   "B1"
      Top             =   45
      Visible         =   0   'False
      Width           =   675
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   330
      Left            =   300
      TabIndex        =   0
      Top             =   2085
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1.装炉"
      Value           =   1
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   330
      Left            =   7785
      TabIndex        =   1
      Top             =   2085
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2.吹炼"
   End
   Begin Threed.SSCheck Chk_ss3 
      Height          =   330
      Left            =   300
      TabIndex        =   2
      Top             =   3915
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "3.出钢"
   End
   Begin Threed.SSCheck Chk_ss4 
      Height          =   330
      Left            =   5580
      TabIndex        =   3
      Top             =   3915
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "4.溅渣护炉"
   End
   Begin Threed.SSCheck Chk_ss5 
      Height          =   330
      Left            =   10875
      TabIndex        =   4
      Top             =   3915
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "5.倒渣"
   End
   Begin Threed.SSCheck Chk_ss6 
      Height          =   330
      Left            =   300
      TabIndex        =   5
      Top             =   5760
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "6.转炉实绩"
   End
   Begin Threed.SSFrame Frame2 
      Height          =   975
      Left            =   270
      TabIndex        =   10
      Top             =   930
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   1720
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox txt_act_steel_grd 
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
         Left            =   1260
         MaxLength       =   11
         TabIndex        =   27
         Top             =   510
         Width           =   1275
      End
      Begin VB.TextBox txt_dir_steel_grd 
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
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   120
         Width           =   1275
      End
      Begin VB.TextBox txt_mlt_prod_cd 
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
         Left            =   5970
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   510
         Width           =   1440
      End
      Begin VB.TextBox txt_stlgrd_n 
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
         Left            =   2535
         TabIndex        =   23
         Top             =   120
         Width           =   2115
      End
      Begin VB.TextBox txt_stlgrd_s 
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
         Left            =   2535
         TabIndex        =   22
         Top             =   510
         Width           =   2115
      End
      Begin CSTextLibCtl.sidbEdit txt_count_1 
         Height          =   315
         Left            =   10635
         TabIndex        =   25
         Top             =   120
         Width           =   1125
         _Version        =   262145
         _ExtentX        =   1984
         _ExtentY        =   556
         _StockProps     =   125
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
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   120
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "目标钢种号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   120
         Top             =   510
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "实际钢种号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   9495
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "钢铁料消耗"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   9495
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "铁水消耗"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   12315
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "废钢比"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_count_2 
         Height          =   315
         Left            =   10635
         TabIndex        =   28
         Top             =   480
         Width           =   1140
         _Version        =   262145
         _ExtentX        =   2011
         _ExtentY        =   556
         _StockProps     =   125
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
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_count_3 
         Height          =   315
         Left            =   13440
         TabIndex        =   29
         Top             =   120
         Width           =   870
         _Version        =   262145
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   125
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
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   4830
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "目标出钢量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.01
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_pre_heat_wgt 
         Height          =   315
         Left            =   5970
         TabIndex        =   30
         Top             =   120
         Width           =   975
         _Version        =   262145
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   125
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
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel32 
         Height          =   315
         Left            =   4830
         Top             =   510
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "工艺路线"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   14355
         TabIndex        =   33
         Top             =   180
         Width           =   180
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/t"
         Height          =   180
         Left            =   11790
         TabIndex        =   32
         Top             =   525
         Width           =   465
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg/t"
         Height          =   180
         Left            =   11790
         TabIndex        =   31
         Top             =   165
         Width           =   345
      End
   End
   Begin Threed.SSFrame sf1 
      Height          =   1275
      Left            =   270
      TabIndex        =   11
      Top             =   2460
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   2249
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox txt_ret_heat_no 
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
         Left            =   5955
         TabIndex        =   35
         Top             =   210
         Width           =   1005
      End
      Begin VB.ComboBox cbo_cld_id 
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
         ItemData        =   "AFC2010C.frx":0010
         Left            =   945
         List            =   "AFC2010C.frx":0012
         TabIndex        =   34
         Top             =   210
         Width           =   885
      End
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   120
         Top             =   210
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "铁包号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   3045
         Top             =   210
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "废钢量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   5055
         Top             =   210
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "返送炉号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   120
         Top             =   645
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   3045
         Top             =   645
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "铁水量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   5055
         Top             =   645
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "返送量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_1 
         Height          =   315
         Left            =   945
         TabIndex        =   36
         Top             =   645
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_sc_net_wgt 
         Height          =   315
         Left            =   3960
         TabIndex        =   37
         Top             =   210
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
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
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_hm_net_wgt 
         Height          =   315
         Left            =   3960
         TabIndex        =   38
         Top             =   645
         Width           =   855
         _Version        =   262145
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_ret_steel_wgt 
         Height          =   315
         Left            =   5955
         TabIndex        =   39
         Top             =   645
         Width           =   1005
         _Version        =   262145
         _ExtentX        =   1773
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   4845
         TabIndex        =   42
         Top             =   270
         Width           =   180
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   4845
         TabIndex        =   41
         Top             =   690
         Width           =   180
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   7005
         TabIndex        =   40
         Top             =   675
         Width           =   135
      End
   End
   Begin Threed.SSFrame sf5 
      Height          =   1275
      Left            =   10860
      TabIndex        =   13
      Top             =   4290
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   2249
      _Version        =   196609
      BackColor       =   14737632
      Begin CSTextLibCtl.sitxEdit txt_occr_date_10 
         Height          =   315
         Left            =   1680
         TabIndex        =   58
         Top             =   210
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_11 
         Height          =   315
         Left            =   1680
         TabIndex        =   59
         Top             =   630
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   120
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         Caption         =   "开始时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   120
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         Caption         =   "结束时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin Threed.SSFrame sf6 
      Height          =   2985
      Left            =   270
      TabIndex        =   14
      Top             =   6120
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   5265
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.ComboBox cbo_ld_id 
         Height          =   315
         ItemData        =   "AFC2010C.frx":0014
         Left            =   1605
         List            =   "AFC2010C.frx":0016
         TabIndex        =   63
         Top             =   705
         Width           =   1065
      End
      Begin VB.ComboBox cbo_o2_lance_id 
         Height          =   315
         ItemData        =   "AFC2010C.frx":0018
         Left            =   7185
         List            =   "AFC2010C.frx":0022
         TabIndex        =   62
         Top             =   705
         Width           =   1035
      End
      Begin VB.TextBox txt_wire_cd2 
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
         Left            =   1605
         TabIndex        =   61
         Top             =   2520
         Width           =   795
      End
      Begin VB.TextBox txt_wire_cd1 
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
         Left            =   1605
         TabIndex        =   60
         Top             =   2100
         Width           =   795
      End
      Begin InDate.ULabel ULabel50 
         Height          =   315
         Left            =   6075
         Top             =   705
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         Caption         =   "氧枪号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel43 
         Height          =   315
         Left            =   8580
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "氧气用量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   135
         Top             =   705
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "钢包号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel39 
         Height          =   315
         Left            =   3180
         Top             =   705
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "吹炼前温度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel40 
         Height          =   315
         Left            =   3180
         Top             =   1170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "吹炼后温度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel51 
         Height          =   315
         Left            =   6075
         Top             =   1170
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         Caption         =   "氧枪枪龄"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel45 
         Height          =   315
         Left            =   8580
         Top             =   705
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "氩气用量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel34 
         Height          =   315
         Left            =   135
         Top             =   1170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "出钢量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel41 
         Height          =   315
         Left            =   6075
         Top             =   240
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         Caption         =   "出钢后温度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel47 
         Height          =   315
         Left            =   8580
         Top             =   1170
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "氮气用量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel36 
         Height          =   315
         Left            =   135
         Top             =   1635
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "渣量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel42 
         Height          =   315
         Left            =   3180
         Top             =   1635
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "出钢口使用次数"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel49 
         Height          =   315
         Left            =   8580
         Top             =   1635
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Caption         =   "煤气回收量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel38 
         Height          =   315
         Left            =   6075
         Top             =   1635
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         Caption         =   "转炉炉龄"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel53 
         Height          =   315
         Left            =   12195
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "温度℃"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel54 
         Height          =   315
         Left            =   13335
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "吹氧量m3"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel55 
         Height          =   315
         Left            =   11220
         Top             =   705
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "1次补吹"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel56 
         Height          =   315
         Left            =   11220
         Top             =   1170
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "2次补吹"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel57 
         Height          =   315
         Left            =   11220
         Top             =   1635
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         Caption         =   "3次补吹"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit txt_steel_net_wgt 
         Height          =   315
         Left            =   1605
         TabIndex        =   64
         Top             =   1170
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_slag_net_wgt 
         Height          =   315
         Left            =   1605
         TabIndex        =   65
         Top             =   1635
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   12
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit txt_con_use_cnt 
         Height          =   315
         Left            =   7185
         TabIndex        =   66
         Top             =   1635
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_o2_lance_cnt 
         Height          =   315
         Left            =   7185
         TabIndex        =   67
         Top             =   1170
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_tap_fin_temp 
         Height          =   315
         Left            =   7185
         TabIndex        =   68
         Top             =   240
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_blow_end_temp 
         Height          =   315
         Left            =   4665
         TabIndex        =   69
         Top             =   1170
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_blow_str_temp 
         Height          =   315
         Left            =   4665
         TabIndex        =   70
         Top             =   705
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_tap_gate_cnt 
         Height          =   315
         Left            =   4665
         TabIndex        =   71
         Top             =   1635
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_total_n_usage 
         Height          =   315
         Left            =   9660
         TabIndex        =   72
         Top             =   1170
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_total_ar_usage 
         Height          =   315
         Left            =   9660
         TabIndex        =   73
         Top             =   705
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_blow_o_using 
         Height          =   315
         Left            =   9660
         TabIndex        =   74
         Top             =   240
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_ladle_degas 
         Height          =   315
         Left            =   9660
         TabIndex        =   75
         Top             =   1635
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_blow_2_oxy 
         Height          =   315
         Left            =   13335
         TabIndex        =   76
         Top             =   1170
         Width           =   1125
         _Version        =   262145
         _ExtentX        =   1984
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_blow_3_oxy 
         Height          =   315
         Left            =   13335
         TabIndex        =   77
         Top             =   1635
         Width           =   1125
         _Version        =   262145
         _ExtentX        =   1984
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_blow_1_oxy 
         Height          =   315
         Left            =   13335
         TabIndex        =   78
         Top             =   705
         Width           =   1125
         _Version        =   262145
         _ExtentX        =   1984
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_blow_2_str_temp 
         Height          =   315
         Left            =   12195
         TabIndex        =   79
         Top             =   1170
         Width           =   1125
         _Version        =   262145
         _ExtentX        =   1984
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_blow_3_str_temp 
         Height          =   315
         Left            =   12195
         TabIndex        =   80
         Top             =   1635
         Width           =   1125
         _Version        =   262145
         _ExtentX        =   1984
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin CSTextLibCtl.sidbEdit txt_blow_1_str_temp 
         Height          =   315
         Left            =   12195
         TabIndex        =   81
         Top             =   705
         Width           =   1125
         _Version        =   262145
         _ExtentX        =   1984
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
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
      Begin InDate.ULabel ULabel19 
         Height          =   315
         Left            =   135
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "实绩发生时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sitxEdit txt_OCCR_DATE 
         Height          =   315
         Left            =   1605
         TabIndex        =   82
         Top             =   240
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel27 
         Height          =   315
         Left            =   135
         Top             =   2100
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "喂丝种类 1"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   3180
         Top             =   2100
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "喂丝重量 1"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sidbEdit txt_wire_lth1 
         Height          =   315
         Left            =   4665
         TabIndex        =   83
         Top             =   2100
         Width           =   795
         _Version        =   262145
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   125
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel29 
         Height          =   315
         Left            =   135
         Top             =   2520
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "喂丝种类 2"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   3180
         Top             =   2520
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         Caption         =   "喂丝重量 2"
         Alignment       =   1
         BackColor       =   16761024
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sidbEdit txt_wire_lth2 
         Height          =   315
         Left            =   4665
         TabIndex        =   84
         Top             =   2520
         Width           =   795
         _Version        =   262145
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   125
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
         NumIntDigits    =   7
         ShowZero        =   0   'False
         MinValue        =   1
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_upd_date 
         Height          =   315
         Left            =   12195
         TabIndex        =   85
         Top             =   2520
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel44 
         Height          =   315
         Left            =   10770
         Top             =   2520
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "上次修改时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "℃"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5685
         TabIndex        =   96
         Top             =   765
         Width           =   180
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10665
         TabIndex        =   95
         Top             =   270
         Width           =   195
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   2670
         TabIndex        =   94
         Top             =   1230
         Width           =   180
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "℃"
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
         Left            =   5685
         TabIndex        =   93
         Top             =   1230
         Width           =   180
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10665
         TabIndex        =   92
         Top             =   780
         Width           =   195
      End
      Begin VB.Label Label5 
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
         Height          =   375
         Left            =   2670
         TabIndex        =   91
         Top             =   1665
         Width           =   180
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "℃"
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
         Left            =   8235
         TabIndex        =   90
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10665
         TabIndex        =   89
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "m3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10665
         TabIndex        =   88
         Top             =   1725
         Width           =   195
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5505
         TabIndex        =   87
         Top             =   2595
         Width           =   345
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Kg"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5505
         TabIndex        =   86
         Top             =   2160
         Width           =   345
      End
   End
   Begin Threed.SSFrame sf3 
      Height          =   1275
      Left            =   270
      TabIndex        =   49
      Top             =   4290
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   2249
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.ComboBox cbo_sld_id 
         Height          =   315
         Left            =   3015
         TabIndex        =   52
         Top             =   930
         Visible         =   0   'False
         Width           =   885
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_8 
         Height          =   315
         Left            =   1680
         TabIndex        =   50
         Top             =   210
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_occr_date_9 
         Height          =   315
         Left            =   1680
         TabIndex        =   51
         Top             =   630
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   120
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         Caption         =   "开始时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   120
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         Caption         =   "结束时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   90
         Top             =   870
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "出钢量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CSTextLibCtl.sidbEdit sdb_tap_wgt 
         Height          =   315
         Left            =   915
         TabIndex        =   53
         Top             =   870
         Visible         =   0   'False
         Width           =   1035
         _Version        =   262145
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   125
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   5
         ShowZero        =   0   'False
         MaxValue        =   999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel26 
         Height          =   315
         Left            =   2190
         Top             =   930
         Visible         =   0   'False
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         Caption         =   "钢包号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16711680
      End
      Begin VB.Label Label16 
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
         Height          =   375
         Left            =   1935
         TabIndex        =   54
         Top             =   570
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin Threed.SSFrame sf4 
      Height          =   1275
      Left            =   5550
      TabIndex        =   55
      Top             =   4290
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   2249
      _Version        =   196609
      BackColor       =   14737632
      Begin CSTextLibCtl.sitxEdit txt_slag_for_str_date 
         Height          =   315
         Left            =   1710
         TabIndex        =   56
         Top             =   210
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin CSTextLibCtl.sitxEdit txt_slag_for_end_date 
         Height          =   315
         Left            =   1710
         TabIndex        =   57
         Top             =   630
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   150
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         Caption         =   "开始时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   150
         Top             =   630
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         Caption         =   "结束时间"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "AFC2010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      BOF
'-- Program ID        AFC2010C
'-- Document No
'-- Designer          ZhengWen
'-- Coder             ZhengWen
'-- Date              2003.7.23
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
Public sDateTime As String           'Active Form Authority Setting
Public sYear As String               'Active Form Authority Setting
Public sMonth As String              'Active Form Authority Setting
Public sDay As String                'Active Form Authority Setting
Public sOur As String                'Active Form Authority Setting
Public sMin As String                'Active Form Authority Setting
Public sSec As String                'Active Form Authority Setting
Public sQuery_Rt As String           'Active Form Authority Setting
       
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

Dim pControl4 As New Collection      'Master Primary Key Collection
Dim nControl4 As New Collection      'Master Necessary Collection
Dim mControl4 As New Collection      'Master Maxlength check Collection
Dim iControl4 As New Collection      'Master Insert Collection
Dim rControl4 As New Collection      'Master Refer Collection
Dim cControl4 As New Collection      'Master Copy Collection
Dim aControl4 As New Collection      'Master -> Spread Collection
Dim lControl4 As New Collection      'Master Lock Collection

Dim pControl5 As New Collection      'Master Primary Key Collection
Dim nControl5 As New Collection      'Master Necessary Collection
Dim mControl5 As New Collection      'Master Maxlength check Collection
Dim iControl5 As New Collection      'Master Insert Collection
Dim rControl5 As New Collection      'Master Refer Collection
Dim cControl5 As New Collection      'Master Copy Collection
Dim aControl5 As New Collection      'Master -> Spread Collection
Dim lControl5 As New Collection      'Master Lock Collection

Dim pControl6 As New Collection      'Master Primary Key Collection
Dim nControl6 As New Collection      'Master Necessary Collection
Dim mControl6 As New Collection      'Master Maxlength check Collection
Dim iControl6 As New Collection      'Master Insert Collection
Dim rControl6 As New Collection      'Master Refer Collection
Dim cControl6 As New Collection      'Master Copy Collection
Dim aControl6 As New Collection      'Master -> Spread Collection
Dim lControl6 As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim Mc4 As New Collection           'Master Collection
Dim Mc5 As New Collection           'Master Collection
Dim Mc6 As New Collection           'Master Collection

Dim sPLC As String                  'PLC LINE

Private Sub Form_Define()
      
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Master"              'form类型
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       
          Call Gp_Ms_Collection(cbo_HEAT_NO, "p", "n", " ", "i", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_process_no, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(cbo_prc_line, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
           Call Gp_Ms_Collection(cbo_cld_id, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_occr_date_1, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_sc_net_wgt, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_hm_net_wgt, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_ret_heat_no, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    Call Gp_Ms_Collection(txt_ret_steel_wgt, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)

    'MASTER Collection
    Mc1.Add Item:="AFC2010C.P_REFER1", Key:="P-R"
    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=cControl1, Key:="cControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"

         Call Gp_Ms_Collection(cbo_HEAT_NO, "p", "n", " ", "i", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(cbo_prc_line, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_occr_date_61, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_occr_date_71, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_occr_date_62, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_occr_date_72, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_occr_date_63, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_occr_date_73, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

    'MASTER Collection
    Mc2.Add Item:="AFC2010C.P_REFER2", Key:="P-R"
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
      
        Call Gp_Ms_Collection(cbo_HEAT_NO, "p", "n", " ", "i", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
       Call Gp_Ms_Collection(cbo_prc_line, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
         'Call Gp_Ms_Collection(cbo_sld_id, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(txt_occr_date_8, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
    Call Gp_Ms_Collection(txt_occr_date_9, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
        'Call Gp_Ms_Collection(sdb_tap_wgt, " ", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)

    'MASTER Collection
    Mc3.Add Item:="AFC2010C.P_REFER3", Key:="P-R"
    Mc3.Add Item:=pControl3, Key:="pControl"
    Mc3.Add Item:=nControl3, Key:="nControl"
    Mc3.Add Item:=mControl3, Key:="mControl"
    Mc3.Add Item:=iControl3, Key:="iControl"
    Mc3.Add Item:=rControl3, Key:="rControl"
    Mc3.Add Item:=cControl3, Key:="cControl"
    Mc3.Add Item:=aControl3, Key:="aControl"
    Mc3.Add Item:=lControl3, Key:="lControl"

              Call Gp_Ms_Collection(cbo_HEAT_NO, "p", "n", " ", "i", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
             Call Gp_Ms_Collection(cbo_prc_line, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(txt_slag_for_str_date, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    Call Gp_Ms_Collection(txt_slag_for_end_date, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)

'MASTER Collection
     Mc4.Add Item:="AFC2010C.P_REFER4", Key:="P-R"
     Mc4.Add Item:=pControl4, Key:="pControl"
     Mc4.Add Item:=nControl4, Key:="nControl"
     Mc4.Add Item:=mControl4, Key:="mControl"
     Mc4.Add Item:=iControl4, Key:="iControl"
     Mc4.Add Item:=rControl4, Key:="rControl"
     Mc4.Add Item:=cControl4, Key:="cControl"
     Mc4.Add Item:=aControl4, Key:="aControl"
     Mc4.Add Item:=lControl4, Key:="lControl"

           Call Gp_Ms_Collection(cbo_HEAT_NO, "p", "n", " ", "i", "r", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
          Call Gp_Ms_Collection(cbo_prc_line, " ", " ", " ", " ", "r", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
      Call Gp_Ms_Collection(txt_occr_date_10, " ", " ", " ", " ", "r", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
      Call Gp_Ms_Collection(txt_occr_date_11, " ", " ", " ", " ", "r", " ", " ", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)

    'MASTER Collection
    Mc5.Add Item:="AFC2010C.P_REFER5", Key:="P-R"
    Mc5.Add Item:=pControl5, Key:="pControl"
    Mc5.Add Item:=nControl5, Key:="nControl"
    Mc5.Add Item:=mControl5, Key:="mControl"
    Mc5.Add Item:=iControl5, Key:="iControl"
    Mc5.Add Item:=rControl5, Key:="rControl"
    Mc5.Add Item:=cControl5, Key:="cControl"
    Mc5.Add Item:=aControl5, Key:="aControl"
    Mc5.Add Item:=lControl5, Key:="lControl"

           Call Gp_Ms_Collection(cbo_HEAT_NO, "p", "n", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
        Call Gp_Ms_Collection(txt_process_no, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
          Call Gp_Ms_Collection(cbo_prc_line, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
             Call Gp_Ms_Collection(cbo_shift, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
          Call Gp_Ms_Collection(cbo_group_cd, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
            Call Gp_Ms_Collection(txt_emp_cd, " ", " ", " ", "i", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
     Call Gp_Ms_Collection(txt_dir_steel_grd, " ", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
          Call Gp_Ms_Collection(txt_stlgrd_n, " ", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
     Call Gp_Ms_Collection(txt_act_steel_grd, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
          Call Gp_Ms_Collection(txt_stlgrd_s, " ", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_pre_heat_wgt, " ", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
       Call Gp_Ms_Collection(txt_mlt_prod_cd, " ", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
           Call Gp_Ms_Collection(txt_count_1, " ", " ", " ", " ", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
           Call Gp_Ms_Collection(txt_count_2, " ", " ", " ", " ", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
           Call Gp_Ms_Collection(txt_count_3, " ", " ", " ", " ", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
            Call Gp_Ms_Collection(cbo_cld_id, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
       Call Gp_Ms_Collection(txt_occr_date_1, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
        Call Gp_Ms_Collection(txt_sc_net_wgt, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
        Call Gp_Ms_Collection(txt_hm_net_wgt, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
       Call Gp_Ms_Collection(txt_ret_heat_no, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
     Call Gp_Ms_Collection(txt_ret_steel_wgt, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_occr_date_61, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_occr_date_71, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_occr_date_62, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_occr_date_72, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_occr_date_63, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_occr_date_73, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
            'Call Gp_Ms_Collection(cbo_sld_id, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
       Call Gp_Ms_Collection(txt_occr_date_8, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
       Call Gp_Ms_Collection(txt_occr_date_9, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
           'Call Gp_Ms_Collection(sdb_tap_wgt, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
 Call Gp_Ms_Collection(txt_slag_for_str_date, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
 Call Gp_Ms_Collection(txt_slag_for_end_date, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_occr_date_10, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_occr_date_11, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
         Call Gp_Ms_Collection(txt_OCCR_DATE, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
             Call Gp_Ms_Collection(cbo_ld_id, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
     Call Gp_Ms_Collection(txt_steel_net_wgt, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_slag_net_wgt, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
       Call Gp_Ms_Collection(txt_con_use_cnt, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
     Call Gp_Ms_Collection(txt_blow_str_temp, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
     Call Gp_Ms_Collection(txt_blow_end_temp, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_tap_fin_temp, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_tap_gate_cnt, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
       Call Gp_Ms_Collection(cbo_o2_lance_id, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_o2_lance_cnt, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
      Call Gp_Ms_Collection(txt_blow_o_using, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
    Call Gp_Ms_Collection(txt_total_ar_usage, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
     Call Gp_Ms_Collection(txt_total_n_usage, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
       Call Gp_Ms_Collection(txt_ladle_degas, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
   Call Gp_Ms_Collection(txt_blow_1_str_temp, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
        Call Gp_Ms_Collection(txt_blow_1_oxy, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
   Call Gp_Ms_Collection(txt_blow_2_str_temp, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
        Call Gp_Ms_Collection(txt_blow_2_oxy, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
   Call Gp_Ms_Collection(txt_blow_3_str_temp, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
        Call Gp_Ms_Collection(txt_blow_3_oxy, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
          Call Gp_Ms_Collection(txt_wire_cd1, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
         Call Gp_Ms_Collection(txt_wire_lth1, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
          Call Gp_Ms_Collection(txt_wire_cd2, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
         Call Gp_Ms_Collection(txt_wire_lth2, " ", " ", " ", "i", "r", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
          Call Gp_Ms_Collection(txt_upd_date, " ", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
               
               Call Gp_Ms_Collection(txt_plt, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
              Call Gp_Ms_Collection(txt_proc, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
              Call Gp_Ms_Collection(txt_oper, " ", " ", " ", "i", " ", " ", " ", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
           
    'MASTER Collection
    Mc6.Add Item:="AFC2010C.P_MODIFY", Key:="P-M"
    Mc6.Add Item:="AFC2010C.P_REFER6", Key:="P-R"
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

End Sub

Private Sub cbo_HEAT_NO_Change()

    cbo_prc_line.Text = Mid(cbo_HEAT_NO, 3, 1)

End Sub

Private Sub cbo_prc_line_Click()

'    Call Heat_ComboAdd(M_CN1, cbo_heat_no)
'
'    If cbo_heat_no.ListCount <> 0 Then
'       cbo_heat_no.ListIndex = 0
'    End If
    
End Sub

Private Sub Chk_ss1_Click(VALUE As Integer)

    If Chk_ss1.VALUE = ssCBUnchecked Then
       If Chk_ss2.VALUE = ssCBUnchecked And Chk_ss3.VALUE = ssCBUnchecked And Chk_ss4.VALUE = ssCBUnchecked And Chk_ss5.VALUE = ssCBUnchecked And Chk_ss6.VALUE = ssCBUnchecked Then
            Chk_ss1.VALUE = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Chk_ss1.VALUE = -1 Then
        Chk_ss1.ForeColor = &HFF&
        Chk_ss2.ForeColor = &H0&
        Chk_ss2.VALUE = ssCBUnchecked
        Chk_ss3.ForeColor = &H0&
        Chk_ss3.VALUE = ssCBUnchecked
        Chk_ss4.ForeColor = &H0&
        Chk_ss4.VALUE = ssCBUnchecked
        Chk_ss5.ForeColor = &H0&
        Chk_ss5.VALUE = ssCBUnchecked
        Chk_ss6.ForeColor = &H0&
        Chk_ss6.VALUE = ssCBUnchecked
        
        sf1.Enabled = True
        sf2.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = False
        sf5.Enabled = False
        sf6.Enabled = False
        sf1.ShadowStyle = ssRaisedShadow
        sf2.ShadowStyle = ssInsetShadow
        sf3.ShadowStyle = ssInsetShadow
        sf4.ShadowStyle = ssInsetShadow
        sf5.ShadowStyle = ssInsetShadow
        sf6.ShadowStyle = ssInsetShadow
        txt_oper = "1"
        cbo_cld_id.SetFocus
    Else
        Chk_ss1.VALUE = ssCBUnchecked
    End If

End Sub

Private Sub Chk_ss2_Click(VALUE As Integer)

    If Chk_ss2.VALUE = ssCBUnchecked Then
       If Chk_ss1.VALUE = ssCBUnchecked And Chk_ss3.VALUE = ssCBUnchecked And Chk_ss4.VALUE = ssCBUnchecked And Chk_ss5.VALUE = ssCBUnchecked And Chk_ss6.VALUE = ssCBUnchecked Then
            Chk_ss2.VALUE = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Chk_ss2.VALUE = -1 Then
        Chk_ss1.ForeColor = &H0&
        Chk_ss1.VALUE = ssCBUnchecked
        Chk_ss2.ForeColor = &HFF&
        Chk_ss3.ForeColor = &H0&
        Chk_ss3.VALUE = ssCBUnchecked
        Chk_ss4.ForeColor = &H0&
        Chk_ss4.VALUE = ssCBUnchecked
        Chk_ss5.ForeColor = &H0&
        Chk_ss5.VALUE = ssCBUnchecked
        Chk_ss6.ForeColor = &H0&
        Chk_ss6.VALUE = ssCBUnchecked
        
        sf1.Enabled = False
        sf2.Enabled = True
        sf3.Enabled = False
        sf4.Enabled = False
        sf5.Enabled = False
        sf6.Enabled = False
        sf1.ShadowStyle = ssInsetShadow
        sf2.ShadowStyle = ssRaisedShadow
        sf3.ShadowStyle = ssInsetShadow
        sf4.ShadowStyle = ssInsetShadow
        sf5.ShadowStyle = ssInsetShadow
        sf6.ShadowStyle = ssInsetShadow
        txt_oper = "2"
        txt_occr_date_61.SetFocus
    Else
        Chk_ss2.VALUE = ssCBUnchecked
    End If

End Sub

Private Sub Chk_ss3_Click(VALUE As Integer)

    If Chk_ss3.VALUE = ssCBUnchecked Then
       If Chk_ss1.VALUE = ssCBUnchecked And Chk_ss2.VALUE = ssCBUnchecked And Chk_ss4.VALUE = ssCBUnchecked And Chk_ss5.VALUE = ssCBUnchecked And Chk_ss6.VALUE = ssCBUnchecked Then
            Chk_ss3.VALUE = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Chk_ss3.VALUE = -1 Then
        Chk_ss1.ForeColor = &H0&
        Chk_ss1.VALUE = ssCBUnchecked
        Chk_ss2.ForeColor = &H0&
        Chk_ss2.VALUE = ssCBUnchecked
        Chk_ss3.ForeColor = &HFF&
        Chk_ss4.ForeColor = &H0&
        Chk_ss4.VALUE = ssCBUnchecked
        Chk_ss5.ForeColor = &H0&
        Chk_ss5.VALUE = ssCBUnchecked
        Chk_ss6.ForeColor = &H0&
        Chk_ss6.VALUE = ssCBUnchecked
        
        sf1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = True
        sf4.Enabled = False
        sf5.Enabled = False
        sf6.Enabled = False
        sf1.ShadowStyle = ssInsetShadow
        sf2.ShadowStyle = ssInsetShadow
        sf3.ShadowStyle = ssRaisedShadow
        sf4.ShadowStyle = ssInsetShadow
        sf5.ShadowStyle = ssInsetShadow
        sf6.ShadowStyle = ssInsetShadow
        txt_oper = "3"
        txt_occr_date_8.SetFocus
            
    Else
        Chk_ss3.VALUE = ssCBUnchecked
    End If

End Sub

Private Sub Chk_ss4_Click(VALUE As Integer)

    If Chk_ss4.VALUE = ssCBUnchecked Then
       If Chk_ss1.VALUE = ssCBUnchecked And Chk_ss2.VALUE = ssCBUnchecked And Chk_ss3.VALUE = ssCBUnchecked And Chk_ss5.VALUE = ssCBUnchecked And Chk_ss6.VALUE = ssCBUnchecked Then
            Chk_ss4.VALUE = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Chk_ss4.VALUE = -1 Then
        Chk_ss1.ForeColor = &H0&
        Chk_ss1.VALUE = ssCBUnchecked
        Chk_ss2.ForeColor = &H0&
        Chk_ss2.VALUE = ssCBUnchecked
        Chk_ss3.ForeColor = &H0&
        Chk_ss3.VALUE = ssCBUnchecked
        Chk_ss4.ForeColor = &HFF&
        Chk_ss5.ForeColor = &H0&
        Chk_ss5.VALUE = ssCBUnchecked
        Chk_ss6.ForeColor = &H0&
        Chk_ss6.VALUE = ssCBUnchecked
        
        sf1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = True
        sf5.Enabled = False
        sf6.Enabled = False
        sf1.ShadowStyle = ssInsetShadow
        sf2.ShadowStyle = ssInsetShadow
        sf3.ShadowStyle = ssInsetShadow
        sf4.ShadowStyle = ssRaisedShadow
        sf5.ShadowStyle = ssInsetShadow
        sf6.ShadowStyle = ssInsetShadow
        txt_oper = "4"
        txt_slag_for_str_date.SetFocus
    Else
        Chk_ss4.VALUE = ssCBUnchecked
    End If

End Sub

Private Sub Chk_ss5_Click(VALUE As Integer)

    If Chk_ss5.VALUE = ssCBUnchecked Then
       If Chk_ss1.VALUE = ssCBUnchecked And Chk_ss2.VALUE = ssCBUnchecked And Chk_ss3.VALUE = ssCBUnchecked And Chk_ss4.VALUE = ssCBUnchecked And Chk_ss6.VALUE = ssCBUnchecked Then
            Chk_ss5.VALUE = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Chk_ss5.VALUE = -1 Then
        Chk_ss1.ForeColor = &H0&
        Chk_ss1.VALUE = ssCBUnchecked
        Chk_ss2.ForeColor = &H0&
        Chk_ss2.VALUE = ssCBUnchecked
        Chk_ss3.ForeColor = &H0&
        Chk_ss3.VALUE = ssCBUnchecked
        Chk_ss4.ForeColor = &H0&
        Chk_ss4.VALUE = ssCBUnchecked
        Chk_ss5.ForeColor = &HFF&
        Chk_ss6.ForeColor = &H0&
        Chk_ss6.VALUE = ssCBUnchecked
        
        sf1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = False
        sf5.Enabled = True
        sf6.Enabled = False
        sf1.ShadowStyle = ssInsetShadow
        sf2.ShadowStyle = ssInsetShadow
        sf3.ShadowStyle = ssInsetShadow
        sf4.ShadowStyle = ssInsetShadow
        sf5.ShadowStyle = ssRaisedShadow
        sf6.ShadowStyle = ssInsetShadow
        txt_oper = "5"
        txt_occr_date_10.SetFocus
    Else
        Chk_ss5.VALUE = ssCBUnchecked
    End If

End Sub

Private Sub Chk_ss6_Click(VALUE As Integer)

    If Chk_ss6.VALUE = ssCBUnchecked Then
       If Chk_ss1.VALUE = ssCBUnchecked And Chk_ss2.VALUE = ssCBUnchecked And Chk_ss3.VALUE = ssCBUnchecked And Chk_ss4.VALUE = ssCBUnchecked And Chk_ss5.VALUE = ssCBUnchecked Then
            Chk_ss6.VALUE = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Chk_ss6.VALUE = -1 Then
        Chk_ss1.ForeColor = &H0&
        Chk_ss1.VALUE = ssCBUnchecked
        Chk_ss2.ForeColor = &H0&
        Chk_ss2.VALUE = ssCBUnchecked
        Chk_ss3.ForeColor = &H0&
        Chk_ss3.VALUE = ssCBUnchecked
        Chk_ss4.ForeColor = &H0&
        Chk_ss4.VALUE = ssCBUnchecked
        Chk_ss5.ForeColor = &H0&
        Chk_ss5.VALUE = ssCBUnchecked
        Chk_ss6.ForeColor = &HFF&
        
        sf1.Enabled = False
        sf2.Enabled = False
        sf3.Enabled = False
        sf4.Enabled = False
        sf5.Enabled = False
        sf6.Enabled = True
        sf1.ShadowStyle = ssInsetShadow
        sf2.ShadowStyle = ssInsetShadow
        sf3.ShadowStyle = ssInsetShadow
        sf4.ShadowStyle = ssInsetShadow
        sf5.ShadowStyle = ssInsetShadow
        sf6.ShadowStyle = ssRaisedShadow
        txt_oper = "6"
        txt_OCCR_DATE.SetFocus
    Else
        Chk_ss6.VALUE = ssCBUnchecked
    End If
    
End Sub

Private Sub cbo_up_Click()

    Dim V_HEAT_NO As String
    
    If Trim(cbo_HEAT_NO.Text) = "" Or Mid(cbo_HEAT_NO, 4, 5) = "99999" Then
       Exit Sub
    End If
    
    cbo_HEAT_NO = Mid(cbo_HEAT_NO, 1, 3) + Format(Val(Mid(cbo_HEAT_NO, 4, 5)) + 1, "00000")
    V_HEAT_NO = cbo_HEAT_NO
    
    Call Form_Cls
    cbo_HEAT_NO = V_HEAT_NO
    cbo_prc_line.Text = Mid(cbo_HEAT_NO, 3, 1)
    Call Form_Ref
  
End Sub

Private Sub cbo_down_Click()

    Dim V_HEAT_NO As String
    
    If Trim(cbo_HEAT_NO.Text) = "" Or Mid(cbo_HEAT_NO, 4, 5) = "00001" Then
      Exit Sub
    End If
    
    cbo_HEAT_NO = Mid(cbo_HEAT_NO, 1, 3) + Format(Val(Mid(cbo_HEAT_NO, 4, 5)) - 1, "00000")
    V_HEAT_NO = cbo_HEAT_NO
    
    Call Form_Cls
    cbo_HEAT_NO = V_HEAT_NO
    cbo_prc_line.Text = Mid(cbo_HEAT_NO, 3, 1)
    Call Form_Ref
  
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

    Dim sQuery  As String
    
    cbo_shift.AddItem "1"
    cbo_shift.AddItem "2"
    cbo_shift.AddItem "3"

    cbo_group_cd.AddItem "A"
    cbo_group_cd.AddItem "B"
    cbo_group_cd.AddItem "C"
    cbo_group_cd.AddItem "D"

    cbo_prc_line.AddItem "1"
    cbo_prc_line.AddItem "2"
    cbo_prc_line.AddItem "3"
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))
    Call Gp_Ms_Cls(Mc4("rControl"))
    Call Gp_Ms_Cls(Mc5("rControl"))
    Call Gp_Ms_Cls(Mc6("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_ControlLock(Mc2("lControl"), True)
    Call Gp_Ms_ControlLock(Mc3("lControl"), True)
    Call Gp_Ms_ControlLock(Mc4("lControl"), True)
    Call Gp_Ms_ControlLock(Mc5("lControl"), True)
    Call Gp_Ms_ControlLock(Mc6("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    Call Gp_Ms_NeceColor(Mc4("nControl"))
    Call Gp_Ms_NeceColor(Mc5("nControl"))
    Call Gp_Ms_NeceColor(Mc6("nControl"))
    
    cbo_prc_line.Text = "1"
    Call Heat_ComboAdd(M_CN1, cbo_HEAT_NO)
    
    If cbo_HEAT_NO.ListCount <> 0 Then
       cbo_HEAT_NO.ListIndex = 0
    End If
  
    Call Gf_ComboAdd(M_CN1, cbo_ld_id, "SELECT CD  FROM ZP_CD WHERE CD_MANA_NO='F0004' AND CD LIKE  'S%' ")
    Call Gf_ComboAdd(M_CN1, cbo_sld_id, "SELECT CD  FROM ZP_CD WHERE CD_MANA_NO='F0004' AND CD LIKE  'S%' ")
    Call Gf_ComboAdd(M_CN1, cbo_cld_id, "SELECT CD  FROM ZP_CD WHERE CD_MANA_NO='F0004' AND CD LIKE  'R%' ")
    
    Chk_ss1.VALUE = ssCBChecked
    Chk_ss1.ForeColor = &HFF&
    Chk_ss2.VALUE = ssCBUnchecked
    Chk_ss3.VALUE = ssCBUnchecked
    Chk_ss4.VALUE = ssCBUnchecked
    Chk_ss5.VALUE = ssCBUnchecked
    Chk_ss6.VALUE = ssCBUnchecked
    
    Chk_ss2.ForeColor = &H808080
    Chk_ss3.ForeColor = &H808080
    Chk_ss4.ForeColor = &H808080
    Chk_ss5.ForeColor = &H808080
    Chk_ss6.ForeColor = &H808080
    sf1.Enabled = True
    sf2.Enabled = False
    sf3.Enabled = False
    sf4.Enabled = False
    sf5.Enabled = False
    sf6.Enabled = False
    txt_act_steel_grd.Enabled = False
    
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011
    
    If cbo_HEAT_NO <> "" Then
       Call Form_Ref
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

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
    Call Gp_Ms_ControlLock(Mc2("pControl"), False)
    Call Gp_Ms_ControlLock(Mc3("pControl"), False)
    Call Gp_Ms_ControlLock(Mc4("pControl"), False)
    Call Gp_Ms_ControlLock(Mc5("pControl"), False)
    Call Gp_Ms_ControlLock(Mc6("pControl"), False)
    
    pControl1(1).SetFocus
    pControl2(1).SetFocus
    pControl3(1).SetFocus
    pControl4(1).SetFocus
    pControl5(1).SetFocus
    pControl6(1).SetFocus
     
    cbo_prc_line.Text = "1"
    Call cbo_prc_line_Click
    
'    Call Gf_ComboAdd2(M_CN1, cbo_heat_no, "SELECT C.HEAT_MANA_NO,C.STEEL_NET_WGT FROM (SELECT A.HEAT_MANA_NO,B.STEEL_NET_WGT FROM  EP_CHARGE_IDX A, FP_CONRSLT B WHERE (PRC_STS = 'A' OR PRC_STS = 'B' ) AND A.HEAT_MANA_NO = B.HEAT_NO(+) ORDER BY A.HEAT_MANA_NO) C WHERE ROWNUM <= 15  ")
    
    Chk_ss1.VALUE = ssCBChecked
    Chk_ss1.ForeColor = &HFF&
    sf2.Enabled = False
    sf3.Enabled = False
    sf4.Enabled = False
    sf5.Enabled = False
    sf6.Enabled = False
    txt_oper = "1"
    txt_count_1.Text = ""
    txt_count_2.Text = ""
    txt_count_3.Text = ""
    'txt_dir_steel_grd.Enabled = False
    'txt_act_steel_grd.Enabled = False
    
    txt_emp_cd = sUserID
    txt_emp_cd.ForeColor = &H80000011
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    txt_emp_cd.Text = sUserID
    txt_process_no.Text = ""
    
End Sub

Public Sub Form_Ref()
    
    Dim Scr_wgt, Hm_wgt, Steel_wgt As Integer
    'cbo_heat_no.Text = Mid(cbo_heat_no.Text, 1, 8)
    
    If Trim(cbo_HEAT_NO.Text) = "" Then
        Call Gp_MsgBoxDisplay("炉号必须输入", "", "错误提示")
    ElseIf Len(Trim(cbo_HEAT_NO.Text)) <> 8 And Len(Trim(cbo_HEAT_NO.Text)) <> 9 Then
        Call Gp_MsgBoxDisplay("炉号长度应为8/9位", "", "错误提示")
    Else
    
        If Gf_Ms_Refer(M_CN1, Mc1, Nothing, Nothing, False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            'Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        End If
        
        If Gf_Ms_Refer(M_CN1, Mc2, Nothing, Nothing, False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc2("pControl"), True)
        End If
        
        If Gf_Ms_Refer(M_CN1, Mc3, Nothing, Nothing, False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc3("pControl"), True)
        End If
         
        If Gf_Ms_Refer(M_CN1, Mc4, Nothing, Nothing, False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc4("pControl"), True)
        End If
        
        If Gf_Ms_Refer(M_CN1, Mc5, Nothing, Nothing, False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc5("pControl"), True)
        End If
        
        If Gf_Ms_Refer(M_CN1, Mc6, Nothing, Nothing, False) Then
        
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc6("pControl"), True)
            txt_dir_steel_grd.Enabled = True
            txt_dir_steel_grd.Locked = True
            txt_dir_steel_grd.ForeColor = &H80000011
            
            txt_act_steel_grd.Enabled = True
            
            Scr_wgt = Val(txt_sc_net_wgt.Text)
            Hm_wgt = Val(txt_hm_net_wgt.Text)
            Steel_wgt = Val(txt_steel_net_wgt.Text)
            
            If (Steel_wgt <> 0) And (Scr_wgt + Hm_wgt) <> 0 Then
               txt_count_1.Text = STR((Scr_wgt + Hm_wgt) * 1000 / Steel_wgt)
               txt_count_2.Text = STR(Hm_wgt * 1000 / Steel_wgt)
               txt_count_3.Text = STR(Scr_wgt * 100 / (Scr_wgt + Hm_wgt))
            ElseIf (Steel_wgt <> 0) And (Scr_wgt + Hm_wgt) = 0 Then
               txt_count_1.Text = STR((Scr_wgt + Hm_wgt) * 1000 / Steel_wgt)
               txt_count_2.Text = STR(Hm_wgt * 1000 / Steel_wgt)
            ElseIf (Steel_wgt = 0) And (Scr_wgt + Hm_wgt) <> 0 Then
               txt_count_3.Text = STR(Scr_wgt * 100 / (Scr_wgt + Hm_wgt))
            End If
            
        End If
        
        If cbo_prc_line = "" Then
           cbo_prc_line.Text = Mid(cbo_HEAT_NO, 3, 1)
        End If
        
    End If
             
    If txt_emp_cd = "" Then
       txt_emp_cd = sUserID
       txt_emp_cd.ForeColor = &H80000011
    End If
    
'    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT SHIFT  FROM GP_SHIFTSTD WHERE TRANCD = 'BC21' ")
'    If sQuery_Rt = "" Then
'       cbo_shift.ListIndex = 0
'    Else
'       cbo_shift.ListIndex = Val(sQuery_Rt) - 1
'    End If
    
'    sQuery_Rt = Gf_CodeFind(M_CN1, "SELECT GROUP_CD  FROM GP_SHIFTSTD WHERE TRANCD = 'BC21' ")
'
'    If sQuery_Rt = "" Then
'       cbo_GROUP_CD.ListIndex = 0
'    Else
'       If sQuery_Rt = "A" Then
'         cbo_GROUP_CD.ListIndex = 0
'       Else
'        If sQuery_Rt = "B" Then
'          cbo_GROUP_CD.ListIndex = 1
'        Else
'         If sQuery_Rt = "C" Then
'          cbo_GROUP_CD.ListIndex = 2
'         Else
'          If sQuery_Rt = "D" Then
'           cbo_GROUP_CD.ListIndex = 3
'          End If
'         End If
'        End If
'       End If
'    End If

End Sub

Public Sub Form_Pro()
    
    Dim sMesg As String
    cbo_HEAT_NO.Text = Mid(cbo_HEAT_NO.Text, 1, 8)
    
    If txt_oper = "1" Then
       If cbo_prc_line.Text <> "1" And cbo_prc_line.Text <> "2" And cbo_prc_line.Text <> "3" Then
          MsgBox "炉座号必须输入", vbCritical, "错误提示"
          Exit Sub
       End If
       If cbo_cld_id.Text = "" Then
          MsgBox "铁包号必须输入", vbCritical, "错误提示"
          Exit Sub
       End If
    ElseIf txt_oper = "3" Then
       If cbo_prc_line.Text <> "1" And cbo_prc_line.Text <> "2" And cbo_prc_line.Text <> "3" Then
          MsgBox "炉座号必须输入", vbCritical, "错误提示"
          Exit Sub
       End If
'       If cbo_sld_id.Text = "" Then
'          MsgBox "钢包号必须输入", vbCritical, "错误提示"
'          Exit Sub
'       End If
    ElseIf txt_oper = "6" Then
      
       If cbo_prc_line.Text <> "1" And cbo_prc_line.Text <> "2" And cbo_prc_line.Text <> "3" Then
          MsgBox "炉座号必须输入", vbCritical, "错误提示"
          Exit Sub
       End If
       If cbo_shift = "" Then
          MsgBox "班次必须输入", vbCritical, "错误提示"
          Exit Sub
       End If
       If cbo_group_cd = "" Then
          MsgBox "班别必须输入", vbCritical, "错误提示"
          Exit Sub
       End If
       If txt_emp_cd = "" Then
          MsgBox "作业人员必须输入", vbCritical, "错误提示"
          Exit Sub
       End If
       If cbo_ld_id.Text = "" Then
          MsgBox "钢包号必须输入", vbCritical, "错误提示"
          Exit Sub
       End If
'       If Trim(txt_steel_net_wgt) = "" Then
'          MsgBox "出钢量必须输入", vbCritical, "错误提示"
'          Exit Sub
'       End If
    End If
    
    If Len(Trim(cbo_HEAT_NO.Text)) <> 8 Then
        sMesg = sMesg + " 炉号必须是8位"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    Else
           
        If txt_process_no.Text <> "" Then
            If Len(Trim(txt_process_no.Text)) <> 9 Then
                sMesg = sMesg + " 处理号必须是9位"
                Call Gp_MsgBoxDisplay(sMesg)
                Exit Sub
            End If
        End If
        
         If Gf_Mc_Authority(sAuthority, Mc6) Then
             If Gf_Ms_Process(M_CN1, Mc6, sAuthority) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
                txt_dir_steel_grd.Enabled = True
                txt_dir_steel_grd.Locked = True
                txt_dir_steel_grd.ForeColor = &H80000011
             End If
        End If

    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
End Sub

Private Sub sdb_tap_wgt_Click()
    sdb_tap_wgt = 0
End Sub

Private Sub sdb_tap_wgt_DblClick()
    sdb_tap_wgt = 0
End Sub

Private Sub txt_act_steel_grd_Change()

    Dim sQuery As String
    
    If Len(Trim(txt_act_steel_grd.Text)) >= 10 Then
      sQuery = "SELECT STEEL_GRD_DETAIL FROM qp_nisco_chmc WHERE STLGRD = '" + txt_act_steel_grd.Text + "' and NVL(STLGRD_FL,'X') <> 'H' "
      txt_stlgrd_s.Text = Gf_CodeFind(M_CN1, sQuery)
    Else
      txt_stlgrd_s.Text = ""
    End If
    
End Sub

Private Sub txt_act_steel_grd_DblClick()

    Call txt_act_steel_grd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_act_steel_grd_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyF4 Then
           
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_act_steel_grd
        DD.rControl.Add Item:=txt_stlgrd_s
        
        Call Pf_Common_DD(M_CN1, KeyCode)
        
   Else
        Dim sQuery As String
        If Len(Trim(txt_act_steel_grd.Text)) >= 10 Then
            sQuery = "SELECT STEEL_GRD_DETAIL FROM qp_nisco_chmc WHERE STLGRD = '" + txt_act_steel_grd.Text + "' and NVL(STLGRD_FL,'X') <> 'H' "
            txt_stlgrd_s.Text = Gf_CodeFind(M_CN1, sQuery)
        Else
            txt_stlgrd_s.Text = ""
        End If
          
    End If

End Sub

Private Sub txt_occr_date_1_DblClick()
         
    txt_occr_date_1.RawData = Format(Now, "YYYYMMDDHHMM")
          
End Sub

Private Sub txt_occr_date_10_DblClick()
         
    txt_occr_date_10.RawData = Format(Now, "YYYYMMDDHHMM")
          
End Sub

Private Sub txt_occr_date_11_DblClick()
         
    txt_occr_date_11.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_occr_date_61_DblClick()
         
    txt_occr_date_61.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_occr_date_62_DblClick()
         
    txt_occr_date_62.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_occr_date_63_DblClick()
         
    txt_occr_date_63.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_occr_date_71_DblClick()
         
    txt_occr_date_71.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_occr_date_72_DblClick()
         
    txt_occr_date_72.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_occr_date_73_DblClick()
         
    txt_occr_date_73.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_occr_date_8_DblClick()
         
    txt_occr_date_8.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_occr_date_9_DblClick()
         
    txt_occr_date_9.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_OCCR_DATE_DblClick()
         
    txt_OCCR_DATE.RawData = Format(Now, "YYYYMMDDHHMM")
    
End Sub

Private Sub txt_slag_for_end_date_Click()
         
    txt_slag_for_end_date.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_slag_for_str_date_Click()
        
    txt_slag_for_str_date.RawData = Format(Now, "YYYYMMDDHHMM")

End Sub

Private Sub txt_wire_cd1_DblClick()

    Call txt_wire_cd1_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_wire_cd1_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF4 Then
     DD.sWitch = "MS"
     DD.sKey = "F0008"
     DD.rControl.Add Item:=txt_wire_cd1

     DD.nameType = "2"
    
     Call Gf_Common_DD(M_CN1, KeyCode)
  End If
  
End Sub

Private Sub txt_wire_cd2_DblClick()

    Call txt_wire_cd2_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_wire_cd2_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF4 Then
     DD.sWitch = "MS"
     DD.sKey = "F0008"
     DD.rControl.Add Item:=txt_wire_cd2

     DD.nameType = "2"
    
     Call Gf_Common_DD(M_CN1, KeyCode)
  End If
  
End Sub

Private Function Pf_Common_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "S"        'Common Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    DD.sQuery = "SELECT STLGRD ""钢种代码"", STEEL_GRD_DETAIL ""钢种名称"" FROM qp_nisco_chmc "
    
    If DD.rControl.Count > 1 Then
        DD.sWhere = " WHERE NVL(STLGRD_FL,'X') <> 'H'  "
        DD.sWhere = DD.sWhere + "   AND STLGRD           like '" & Trim(DD.rControl.Item(1).Text) & "%' "
        DD.sWhere = DD.sWhere + "   AND STEEL_GRD_DETAIL like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        
    End If
    
    Call Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False)
    
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function

Private Function Heat_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo Heat_ComboAdd_Error
    
    Dim AdoRs  As ADODB.Recordset
    Dim sQuery As String
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Heat_ComboAdd = False: Exit Function
    End If
    
    If ClsChk Then
        Cbo.Clear
    End If
     
    sQuery = "          SELECT  A.HEAT_MANA_NO                               "
    sQuery = sQuery + "   FROM (SELECT  A.HEAT_MANA_NO                       "
    sQuery = sQuery + "           FROM  EP_CHARGE_INS A                      "
    sQuery = sQuery + "          WHERE A.PRC_STS IN ('A','B')                "
    sQuery = sQuery + "            AND A.PRC_LINE = '" & cbo_prc_line.Text & "'"
    sQuery = sQuery + "          ORDER BY A.HEAT_MANA_NO ASC) A   "
    sQuery = sQuery + "  WHERE ROWNUM <= 15"
    
    Set AdoRs = New ADODB.Recordset
     
    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
     
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
               If AdoRs.Fields(0) > "0" Then
                  Cbo.AddItem AdoRs.Fields(0)
                  
               Else
                  Cbo.AddItem AdoRs.Fields(0)
              
               End If
            End If
            AdoRs.MoveNext
            
        Wend
        Heat_ComboAdd = True
    Else
         Cbo.AddItem ""
        Heat_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

Heat_ComboAdd_Error:

    Set AdoRs = Nothing
    Heat_ComboAdd = False

End Function


