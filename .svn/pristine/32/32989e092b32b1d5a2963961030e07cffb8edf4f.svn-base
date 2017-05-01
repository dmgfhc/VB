VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGF2080C 
   Caption         =   "废钢实绩查询及修改界面_AGF2080C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   3045
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   5371
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_REMARK 
         Height          =   735
         Left            =   1635
         MaxLength       =   200
         TabIndex        =   34
         Top             =   2160
         Width           =   4575
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
         Left            =   1635
         MaxLength       =   14
         ScrollBars      =   1  'Horizontal
         TabIndex        =   30
         Tag             =   "废钢号"
         Top             =   510
         Width           =   1800
      End
      Begin VB.ComboBox CBO_END_CD 
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
         ItemData        =   "AGF2080C.frx":0000
         Left            =   12090
         List            =   "AGF2080C.frx":000D
         TabIndex        =   29
         Tag             =   "去向"
         Top             =   1800
         Width           =   1630
      End
      Begin VB.ComboBox CBO_SHIFT_REF 
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
         ItemData        =   "AGF2080C.frx":0030
         Left            =   6225
         List            =   "AGF2080C.frx":0040
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txt_Flag 
         Enabled         =   0   'False
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
         Left            =   -15
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1230
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox CBO_SCRAP_CD 
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
         ItemData        =   "AGF2080C.frx":004F
         Left            =   12150
         List            =   "AGF2080C.frx":005C
         TabIndex        =   23
         Top             =   120
         Width           =   1470
      End
      Begin VB.TextBox TXT_SCRAP_CD 
         Height          =   315
         Left            =   12300
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox TXT_PRC 
         Height          =   315
         Left            =   8655
         MaxLength       =   2
         TabIndex        =   21
         Top             =   120
         Width           =   435
      End
      Begin VB.TextBox TXT_PRC_NAME 
         Enabled         =   0   'False
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
         Left            =   9090
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   1425
      End
      Begin VB.TextBox txt_UserId 
         Enabled         =   0   'False
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
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1575
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox CBO_SCRAP_INPUT 
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
         ItemData        =   "AGF2080C.frx":007B
         Left            =   4590
         List            =   "AGF2080C.frx":0088
         TabIndex        =   18
         Tag             =   "种类"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TXT_SCRAP_INPUT 
         Height          =   315
         Left            =   6105
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox TXT_PRC_INPUT 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   16
         Tag             =   "工序"
         Top             =   1080
         Width           =   465
      End
      Begin VB.TextBox TXT_PRC_INPUT_NAME 
         Enabled         =   0   'False
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
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   1245
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
         ItemData        =   "AGF2080C.frx":00A0
         Left            =   4590
         List            =   "AGF2080C.frx":00A2
         TabIndex        =   14
         Tag             =   "班次"
         Top             =   1425
         Width           =   675
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
         ItemData        =   "AGF2080C.frx":00A4
         Left            =   7425
         List            =   "AGF2080C.frx":00A6
         TabIndex        =   13
         Top             =   1425
         Width           =   675
      End
      Begin VB.TextBox TXT_SCRAP_NO 
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
         Left            =   1635
         MaxLength       =   14
         ScrollBars      =   1  'Horizontal
         TabIndex        =   7
         Tag             =   "废钢号"
         Top             =   1770
         Width           =   1800
      End
      Begin VB.TextBox txt_code 
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
         Left            =   7425
         MaxLength       =   1
         TabIndex        =   6
         Tag             =   "原因"
         Top             =   1080
         Width           =   660
      End
      Begin VB.TextBox txt_main_res_cd 
         Enabled         =   0   'False
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
         Left            =   8085
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   2400
      End
      Begin VB.ComboBox cbo_ths_d_mat_var 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "AGF2080C.frx":00A8
         Left            =   12090
         List            =   "AGF2080C.frx":00AA
         TabIndex        =   4
         Tag             =   "增减量(+,-)"
         Top             =   1425
         Width           =   630
      End
      Begin VB.ComboBox CBO_LINE 
         BackColor       =   &H00C0FFFF&
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
         ItemData        =   "AGF2080C.frx":00AC
         Left            =   14895
         List            =   "AGF2080C.frx":00B3
         TabIndex        =   3
         Text            =   "1"
         Top             =   1620
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.ComboBox CBO_PLT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         ItemData        =   "AGF2080C.frx":00BA
         Left            =   14880
         List            =   "AGF2080C.frx":00C1
         TabIndex        =   2
         Text            =   "C1"
         Top             =   1305
         Visible         =   0   'False
         Width           =   720
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   285
         Tag             =   "发生日"
         Top             =   1425
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "发生日期"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sitxEdit TXT_OCCR_TIME 
         Height          =   315
         Left            =   1635
         TabIndex        =   8
         Top             =   1425
         Width           =   1245
         _Version        =   262145
         _ExtentX        =   2196
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
         Text            =   "____-__-__"
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
         Mask            =   "____-__-__"
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   285
         Top             =   1770
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "废钢号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   10890
         Top             =   1080
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "废钢重量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sidbEdit SDB_SCRAP_WGT 
         Height          =   315
         Left            =   12090
         TabIndex        =   9
         Tag             =   "废钢重量"
         Top             =   1080
         Width           =   1605
         _Version        =   262145
         _ExtentX        =   2831
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
         MaxValue        =   9999.999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   6525
         Top             =   1080
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         Caption         =   "原因"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   10890
         Top             =   1425
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "增减量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sidbEdit sdb_ths_d_mat_var 
         Height          =   315
         Left            =   12705
         TabIndex        =   10
         Tag             =   "增减量"
         Top             =   1425
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
         NumIntDigits    =   3
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   6525
         Tag             =   "班别"
         Top             =   1425
         Width           =   870
         _ExtentX        =   1535
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
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   3690
         Top             =   1425
         Width           =   870
         _ExtentX        =   1535
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
      Begin InDate.ULabel ULabel11 
         Height          =   315
         Left            =   285
         Top             =   1080
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         Caption         =   "工序"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   3690
         Top             =   1080
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         Caption         =   "种类"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   7380
         Tag             =   "工序"
         Top             =   120
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "工序"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   10875
         Tag             =   "种类"
         Top             =   120
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "种类"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Height          =   735
         Left            =   285
         Tag             =   "备注"
         Top             =   2160
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1296
         Caption         =   "备注"
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
         Left            =   10875
         Top             =   510
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "废钢总量"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sidbEdit SDB_TOT_WGT 
         Height          =   315
         Left            =   12150
         TabIndex        =   24
         Tag             =   "废钢总量"
         Top             =   510
         Width           =   1470
         _Version        =   262145
         _ExtentX        =   2593
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
         NumIntDigits    =   5
         ShowZero        =   0   'False
         MaxValue        =   99999.999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   5340
         Top             =   120
         Width           =   870
         _ExtentX        =   1535
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
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   6495
         Tag             =   "发生日"
         Top             =   1800
         Width           =   1320
         _ExtentX        =   2328
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
         Left            =   7845
         TabIndex        =   28
         Top             =   1800
         Width           =   2115
         _Version        =   262145
         _ExtentX        =   3731
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   10890
         Top             =   1800
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "去向"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Top             =   510
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "查询废钢号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.UDate SDT_PROD_DATE_FROM 
         Height          =   315
         Left            =   1635
         TabIndex        =   31
         Tag             =   "起始日期"
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
      End
      Begin InDate.UDate SDT_PROD_DATE_TO 
         Height          =   315
         Left            =   3375
         TabIndex        =   32
         Tag             =   "起始日期"
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.74
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         BackColor       =   16777215
      End
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   240
         Top             =   120
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "废钢号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   120
         Left            =   3165
         TabIndex        =   33
         Top             =   240
         Width           =   195
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   15
         X2              =   15075
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         X1              =   15
         X2              =   15060
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ton"
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
         Left            =   13710
         TabIndex        =   25
         Top             =   555
         Width           =   360
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "ton"
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
         TabIndex        =   12
         Top             =   1125
         Width           =   360
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ton"
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
         TabIndex        =   11
         Top             =   1485
         Width           =   360
      End
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   6105
      Left            =   90
      TabIndex        =   0
      Top             =   3090
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   10769
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   21
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AGF2080C.frx":00C8
   End
End
Attribute VB_Name = "AGF2080C"
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
'-- Program Name      废钢实绩
'-- Program ID        AGF2080C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
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

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_OCCR_TIME = 1
Const SPD_SHIFT = 2
Const SPD_GROUP = 3
Const SPD_PRC_INPUT = 4
Const SPD_SCRAP_INPUT = 6
Const SPD_SCRAP_NO = 7
Const SPD_SCRAP_WGT = 8
Const SPD_CODE = 9
Const SPD_CODE_DES = 10
Const SPD_END_TIME = 11
Const SPD_END_CD = 12
Const SPD_REMARK = 20

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Hsheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(CBO_SHIFT_REF, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
           Call Gp_Ms_Collection(TXT_PRC, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(CBO_SCRAP_CD, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(TXT_SCRAP_CD, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(SDB_TOT_WGT, " ", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_Flag, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
           Call Gp_Ms_Collection(CBO_PLT, "p", "n", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(TXT_PRC_INPUT, "p", "n", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
Call Gp_Ms_Collection(TXT_PRC_INPUT_NAME, " ", " ", " ", " ", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
          Call Gp_Ms_Collection(CBO_LINE, "p", "n", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(TXT_OCCR_TIME, "p", "n", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", "n", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
         Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
   Call Gp_Ms_Collection(CBO_SCRAP_INPUT, " ", "n", " ", " ", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
   Call Gp_Ms_Collection(TXT_SCRAP_INPUT, "p", "n", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
      Call Gp_Ms_Collection(TXT_SCRAP_NO, "p", " ", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
          Call Gp_Ms_Collection(txt_code, "p", "n", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
   Call Gp_Ms_Collection(txt_main_res_cd, " ", " ", " ", " ", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
     Call Gp_Ms_Collection(SDB_SCRAP_WGT, " ", "n", " ", "i", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
 Call Gp_Ms_Collection(cbo_ths_d_mat_var, " ", " ", " ", "i", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
 Call Gp_Ms_Collection(sdb_ths_d_mat_var, " ", " ", " ", "i", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
        Call Gp_Ms_Collection(txt_UserId, " ", " ", " ", "i", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
      Call Gp_Ms_Collection(TXT_END_TIME, " ", " ", " ", "i", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
        Call Gp_Ms_Collection(CBO_END_CD, " ", " ", " ", "i", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
         Call Gp_Ms_Collection(TXT_REMARK, " ", " ", " ", "i", "", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    
    'MASTER Collection
    Mc2.Add Item:="AGF2080C.P_MODIFY", Key:="P-M"
    Mc2.Add Item:="AGF2080C.P_REFER", Key:="P-R"
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)


    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGF2080C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
'    Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 7, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub CBO_END_CD_Click()
    CBO_END_CD.Text = Mid(Trim(CBO_END_CD.Text), 1, 2)

End Sub

Private Sub CBO_SCRAP_CD_Change()
    TXT_SCRAP_CD.Text = Mid(Trim(CBO_SCRAP_CD.Text), 1, 2)
End Sub

Private Sub CBO_SCRAP_CD_Click()
    TXT_SCRAP_CD.Text = Mid(Trim(CBO_SCRAP_CD.Text), 1, 2)
End Sub

Private Sub CBO_SCRAP_INPUT_Click()
    TXT_SCRAP_INPUT.Text = Mid(Trim(CBO_SCRAP_INPUT.Text), 1, 2)

    cbo_ths_d_mat_var.Text = ""
    sdb_ths_d_mat_var.Text = ""
End Sub

Private Sub Form_Activate()

Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
Call MenuTool_ReSet

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    cbo_ths_d_mat_var.AddItem "+"
    cbo_ths_d_mat_var.AddItem "-"

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
'    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 2)
        
    SDT_PROD_DATE_FROM.RawData = Format(Date, "yyyymmdd")
    SDT_PROD_DATE_TO.RawData = Format(Date, "yyyymmdd")
    
    CBO_SHIFT.AddItem "1"
    CBO_SHIFT.AddItem "2"
    CBO_SHIFT.AddItem "3"
    
    CBO_GROUP.AddItem "A"
    CBO_GROUP.AddItem "B"
    CBO_GROUP.AddItem "C"
    CBO_GROUP.AddItem "D"
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)

    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing

    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing

    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing

    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

Public Sub Form_Cls()

        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call Gf_Sp_Cls(Proc_Sc("SC"))
'        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        Call Gp_Ms_ControlLock(Mc1("rControl"), False)
        
        txt_main_res_cd.Text = ""
        cbo_ths_d_mat_var.Text = ""
        sdb_ths_d_mat_var.Text = ""
        CBO_SCRAP_INPUT.Text = ""

End Sub

Public Sub Form_Ref()
    
    Dim iRow As Integer
    On Error Resume Next

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
 
    SDB_TOT_WGT.Value = 0
    txt_Flag.Text = "C1"
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        If ss1.MaxRows > 0 Then
           For iRow = 1 To ss1.MaxRows
               ss1.Row = iRow
               ss1.Col = 8
               SDB_TOT_WGT.Value = SDB_TOT_WGT.Value + Val(ss1.Value)
           Next iRow
        End If
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
    
    cbo_ths_d_mat_var.Text = ""
    sdb_ths_d_mat_var.Text = ""
  
End Sub

Public Sub Form_Pro()

    Dim SMESG As String

    If Not Gp_DateCheck(TXT_OCCR_TIME) Then
        SMESG = " 请正确输入发生时间 ！"
        Call Gp_MsgBoxDisplay(SMESG)
        Exit Sub
    End If
    
    If Gf_Mc_Authority(sAuthority, Mc1) Then
        txt_UserId.Text = sUserID
        CBO_PLT.Text = "C1"
        CBO_LINE.Text = "1"
        txt_Flag.Text = "C1"
       
        If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then
           Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl"))
           ss1.OperationMode = OperationModeNormal
           Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
           Call MenuTool_ReSet
        End If
    End If

End Sub

Public Sub Form_Del()

    CBO_PLT.Enabled = False
    TXT_PRC_INPUT.Enabled = False
    CBO_LINE.Enabled = False
    TXT_OCCR_TIME.Enabled = False
    CBO_SHIFT.Enabled = False
    CBO_GROUP.Enabled = False
    TXT_SCRAP_INPUT.Enabled = False
    TXT_SCRAP_NO.Enabled = False
    txt_code.Enabled = False
    
    If Not Gf_Ms_Del(M_CN1, Mc2) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
       Call MenuTool_ReSet
    End If

End Sub

Public Sub Form_Ins()

'    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

End Sub

Public Sub Spread_Cpy()

'    Call Gp_Sp_Copy(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Pst()

'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
   ' Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1

End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol

End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0

End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()

'    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    With ss1
         
            .Row = Row
            .Col = SPD_OCCR_TIME
             TXT_OCCR_TIME.RawData = Mid(.Text, 1, 4) + Mid(.Text, 6, 2) + Mid(.Text, 9, 2)
            .Col = SPD_SHIFT
             CBO_SHIFT.Text = .Text
            .Col = SPD_GROUP
             CBO_GROUP.Text = .Text
            .Col = SPD_PRC_INPUT
             TXT_PRC_INPUT.Text = .Text
            .Col = SPD_SCRAP_INPUT
             CBO_SCRAP_INPUT.Text = .Text
             TXT_SCRAP_INPUT.Text = Left(.Text, 2)
            .Col = SPD_SCRAP_NO
             TXT_SCRAP_NO.Text = .Text
            .Col = SPD_SCRAP_WGT
             SDB_SCRAP_WGT.Text = .Text
            .Col = SPD_CODE
             txt_code.Text = .Text
            .Col = SPD_CODE_DES
             txt_main_res_cd.Text = .Text
            .Col = SPD_END_TIME
             TXT_END_TIME.RawData = Mid(.Text, 1, 4) + Mid(.Text, 6, 2) + Mid(.Text, 9, 2) + Mid(.Text, 12, 2) + Mid(.Text, 15, 2) + Mid(.Text, 18, 2)
            .Col = SPD_END_CD
             CBO_END_CD.Text = .Text
            .Col = SPD_REMARK
             TXT_REMARK.Text = .Text
          
    End With
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub



Private Sub txt_code_DblClick()
    Call txt_code_KeyUp(vbKeyF4, 0)
End Sub



Private Sub TXT_END_TIME_Click()
    TXT_END_TIME.RawData = Gf_DTSet(M_CN1, "S")
End Sub

Private Sub TXT_PRC_DblClick()
    Call TXT_PRC_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_PRC_INPUT_DblClick()
    Call TXT_PRC_INPUT_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_OCCR_TIME_DblClick()
  
    TXT_OCCR_TIME.RawData = Gf_DTSet(M_CN1, "D")
         
End Sub

Private Sub txt_code_Change()
    
    If Len(Trim(txt_code)) = txt_code.MaxLength Then
        txt_main_res_cd.Text = Gf_ComnNameFind(M_CN1, "F0011", Trim(txt_code.Text), 1)
    Else
        txt_main_res_cd.Text = ""
    End If
    
End Sub

Private Sub txt_code_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "G0017"
        DD.rControl.Add Item:=txt_code
        DD.rControl.Add Item:=txt_main_res_cd

        DD.nameType = "1"

        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If

End Sub

Private Sub TXT_PRC_Change()
    
    If Len(Trim(TXT_PRC)) = TXT_PRC.MaxLength Then
        TXT_PRC_NAME.Text = Gf_ComnNameFind(M_CN1, "C0002", Trim(TXT_PRC.Text), 2)
    Else
        TXT_PRC_NAME.Text = ""
    End If
    
End Sub

Private Sub TXT_PRC_KeyUp(KeyCode As Integer, Shift As Integer)
    
     If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "C0002"
        TXT_PRC = "C"
        DD.rControl.Add Item:=TXT_PRC

        DD.nameType = "1"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If
    
End Sub

Private Sub TXT_PRC_INPUT_Change()
    
    If Len(Trim(TXT_PRC_INPUT)) = TXT_PRC_INPUT.MaxLength Then
        TXT_PRC_INPUT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0002", Trim(TXT_PRC_INPUT.Text), 2)
    Else
        TXT_PRC_INPUT_NAME.Text = ""
    End If
    
End Sub

Private Sub TXT_PRC_INPUT_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
            
        DD.sWitch = "MS"
        DD.sKey = "C0002"
        TXT_PRC_INPUT = "C"
        DD.rControl.Add Item:=TXT_PRC_INPUT

        DD.nameType = "1"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub
Private Sub SDT_PROD_DATE_FROM_GotFocus()
     If SDT_PROD_DATE_FROM.RawData = "" Then
        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_DATE_TO_GotFocus()
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub
