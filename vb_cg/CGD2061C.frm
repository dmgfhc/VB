VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGD2061C 
   Caption         =   "探伤实绩查询_CGD2061C"
   ClientHeight    =   9330
   ClientLeft      =   15
   ClientTop       =   1740
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame Single 
      Height          =   1305
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   2302
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox CBO_EMP 
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
         ItemData        =   "CGD2061C.frx":0000
         Left            =   7755
         List            =   "CGD2061C.frx":0040
         TabIndex        =   27
         Top             =   120
         Width           =   1365
      End
      Begin VB.ComboBox CBO_SURFGRD 
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
         ItemData        =   "CGD2061C.frx":00F1
         Left            =   8565
         List            =   "CGD2061C.frx":0107
         TabIndex        =   26
         Tag             =   "等级"
         Top             =   840
         Width           =   1065
      End
      Begin VB.CheckBox chk_Cond_J 
         BackColor       =   &H00E0E0E0&
         Caption         =   "技术科"
         Height          =   255
         Left            =   11490
         TabIndex        =   23
         Tag             =   "J"
         Top             =   -120
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txt_f_addr 
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
         Left            =   7755
         TabIndex        =   21
         Tag             =   "标准代码"
         Top             =   480
         Width           =   1365
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   315
         Left            =   13245
         TabIndex        =   16
         Top             =   885
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   255
         Caption         =   "探伤报告"
      End
      Begin VB.CheckBox chk_Cond_B 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中板"
         Height          =   255
         Left            =   9750
         TabIndex        =   18
         Tag             =   "B"
         Top             =   -90
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CheckBox chk_Cond_W 
         BackColor       =   &H00E0E0E0&
         Caption         =   "协力"
         Height          =   255
         Left            =   10620
         TabIndex        =   17
         Tag             =   "W"
         Top             =   -90
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox TXT_MAT_NO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5610
         MaxLength       =   14
         TabIndex        =   15
         Tag             =   "标准代码"
         Top             =   840
         Width           =   1965
      End
      Begin VB.TextBox TXT_ADDR 
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
         Index           =   2
         Left            =   10350
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2670
         Width           =   585
      End
      Begin VB.TextBox TXT_ADDR 
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
         Index           =   1
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2670
         Width           =   585
      End
      Begin VB.TextBox TXT_STDSPEC 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1305
         TabIndex        =   11
         Tag             =   "标准代码"
         Top             =   840
         Width           =   1950
      End
      Begin VB.TextBox TXT_ADDR 
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
         Index           =   0
         Left            =   8460
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2670
         Width           =   1005
      End
      Begin VB.TextBox TXT_UST_STAND_NAME 
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
         Left            =   1995
         TabIndex        =   7
         Top             =   480
         Width           =   2455
      End
      Begin VB.TextBox TXT_UST_STAND_NO 
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
         Left            =   1305
         TabIndex        =   6
         Tag             =   "检查标准"
         Top             =   480
         Width           =   690
      End
      Begin VB.ComboBox CBO_UST_DEC 
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
         ItemData        =   "CGD2061C.frx":0141
         Left            =   5610
         List            =   "CGD2061C.frx":014E
         TabIndex        =   4
         Top             =   480
         Width           =   930
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
         ItemData        =   "CGD2061C.frx":0168
         Left            =   5610
         List            =   "CGD2061C.frx":0178
         TabIndex        =   3
         Top             =   120
         Width           =   930
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   105
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "探伤日期"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   4560
         Top             =   120
         Width           =   1020
         _ExtentX        =   1799
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   4560
         Top             =   480
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         Caption         =   "探伤结果"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   9735
         Top             =   120
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         Caption         =   "厚度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sidbEdit SDB_UST_THK 
         Height          =   315
         Left            =   10695
         TabIndex        =   5
         Top             =   120
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   105
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "UST标准"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Index           =   3
         Left            =   7350
         Top             =   2670
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         Caption         =   "垛位号"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   1
         Left            =   105
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "标准号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin VB.TextBox TXT_STDSPEC_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -150
         MaxLength       =   2
         TabIndex        =   13
         Tag             =   "标准代码"
         Top             =   840
         Visible         =   0   'False
         Width           =   465
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   9735
         Top             =   480
         Width           =   930
         _ExtentX        =   1640
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
      Begin CSTextLibCtl.sidbEdit SDB_UST_WID 
         Height          =   315
         Left            =   10695
         TabIndex        =   14
         Top             =   480
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   0
         Left            =   4560
         Top             =   840
         Width           =   1020
         _ExtentX        =   1799
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
         ForeColor       =   16711680
      End
      Begin VB.TextBox TXT_CO_CD 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12510
         MaxLength       =   2
         TabIndex        =   19
         Tag             =   "标准代码"
         Top             =   -150
         Visible         =   0   'False
         Width           =   465
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   315
         Left            =   13245
         TabIndex        =   20
         Top             =   525
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   255
         Caption         =   "码堆报告"
      End
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   2
         Left            =   6555
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "码堆垛位"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   3
         Left            =   6555
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "探伤人员"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   9555
         Top             =   930
         Visible         =   0   'False
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "生产日期"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sitxEdit SDT_PROD_DATE 
         Height          =   315
         Left            =   10755
         TabIndex        =   22
         Tag             =   "探伤日期"
         Top             =   930
         Visible         =   0   'False
         Width           =   1200
         _Version        =   262145
         _ExtentX        =   2117
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin CSTextLibCtl.sitxEdit SDT_PROD_DATETO 
         Height          =   315
         Left            =   12300
         TabIndex        =   24
         Tag             =   "探伤日期"
         Top             =   930
         Visible         =   0   'False
         Width           =   1230
         _Version        =   262145
         _ExtentX        =   2170
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
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   7605
         Top             =   840
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         Caption         =   "表面"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sidbEdit SDB_WGT 
         Height          =   225
         Left            =   13860
         TabIndex        =   28
         Top             =   210
         Width           =   990
         _Version        =   262145
         _ExtentX        =   1746
         _ExtentY        =   397
         _StockProps     =   125
         Text            =   " 0"
         ForeColor       =   255
         BackColor       =   14804173
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.000"
         Text            =   ""
         StartText.x     =   2
         StartText.y     =   0
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
         NumIntDigits    =   10
         ShowZero        =   0   'False
         MaxValue        =   9999.99
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   13185
         Top             =   150
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Caption         =   " 重量（         ）"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin InDate.UDate SDT_PROD_DATE_FROM 
         Height          =   315
         Left            =   1305
         TabIndex        =   29
         Tag             =   "起始日期"
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
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
         Left            =   3015
         TabIndex        =   30
         Tag             =   "起始日期"
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
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
      Begin CSTextLibCtl.sidbEdit SDB_UST_THK_TO 
         Height          =   315
         Left            =   11910
         TabIndex        =   32
         Top             =   120
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
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
      Begin CSTextLibCtl.sidbEdit SDB_UST_WID_TO 
         Height          =   315
         Left            =   11910
         TabIndex        =   34
         Top             =   480
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   9735
         Top             =   840
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         Caption         =   "长度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin CSTextLibCtl.sidbEdit SDB_UST_LEN 
         Height          =   315
         Left            =   10695
         TabIndex        =   35
         Top             =   840
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
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
         FmtControl      =   1
         NumDecDigits    =   1
         NumIntDigits    =   6
         ShowZero        =   0   'False
         MaxValue        =   999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit SDB_UST_LEN_TO 
         Height          =   315
         Left            =   11910
         TabIndex        =   37
         Top             =   840
         Width           =   1050
         _Version        =   262145
         _ExtentX        =   1852
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
         FmtControl      =   1
         NumDecDigits    =   1
         NumIntDigits    =   6
         ShowZero        =   0   'False
         MaxValue        =   999999.9
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin Threed.SSPanel SSP4 
         Height          =   315
         Left            =   3270
         TabIndex        =   38
         Top             =   840
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "重点订单"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   180
         Index           =   5
         Left            =   11715
         TabIndex        =   36
         Top             =   975
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   180
         Index           =   4
         Left            =   11715
         TabIndex        =   33
         Top             =   615
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   180
         Index           =   3
         Left            =   11715
         TabIndex        =   31
         Top             =   255
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   180
         Index           =   2
         Left            =   2640
         TabIndex        =   25
         Top             =   1140
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   13080
         X2              =   13080
         Y1              =   30
         Y2              =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   240
         Index           =   1
         Left            =   10050
         TabIndex        =   12
         Top             =   2790
         Width           =   300
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   180
         Index           =   0
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   240
      End
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   5985
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   10557
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   57
      MaxRows         =   50
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      ScrollBarExtMode=   -1  'True
      SpreadDesigner  =   "CGD2061C.frx":0187
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   39
      Top             =   1320
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   3625
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_UST_GRADE 
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
         Left            =   10080
         TabIndex        =   61
         Top             =   1200
         Width           =   2115
      End
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10080
         MaxLength       =   99
         TabIndex        =   60
         Tag             =   "标准代码"
         Text            =   "100%"
         Top             =   840
         Width           =   2085
      End
      Begin VB.TextBox TXT_UST_STAND_REPORT 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7200
         MaxLength       =   99
         TabIndex        =   59
         Top             =   1200
         Width           =   1485
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7200
         MaxLength       =   99
         TabIndex        =   58
         Tag             =   "标准代码"
         Text            =   "纵波直接接触法"
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox Text17 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         MaxLength       =   99
         TabIndex        =   57
         Tag             =   "标准代码"
         Text            =   "夏成胜"
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         MaxLength       =   99
         TabIndex        =   56
         Tag             =   "标准代码"
         Text            =   "0dB"
         Top             =   1200
         Width           =   1485
      End
      Begin VB.TextBox Text14 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   99
         TabIndex        =   55
         Tag             =   "标准代码"
         Text            =   "3级"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox Text13 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   99
         TabIndex        =   54
         Tag             =   "标准代码"
         Text            =   "3mm  FBH"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13440
         MaxLength       =   99
         TabIndex        =   53
         Tag             =   "标准代码"
         Text            =   "杨德蓉"
         Top             =   1200
         Width           =   1485
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13440
         MaxLength       =   99
         TabIndex        =   52
         Tag             =   "标准代码"
         Text            =   "平行于轧制线"
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13440
         MaxLength       =   99
         TabIndex        =   51
         Tag             =   "标准代码"
         Text            =   "30*2000*73500"
         Top             =   480
         Width           =   1485
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7200
         MaxLength       =   99
         TabIndex        =   50
         Tag             =   "标准代码"
         Text            =   "5MHz"
         Top             =   480
         Width           =   1485
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         MaxLength       =   99
         TabIndex        =   49
         Tag             =   "标准代码"
         Text            =   "60×20"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   99
         TabIndex        =   48
         Tag             =   "标准代码"
         Text            =   "3STSE 18.3/8PB5"
         Top             =   480
         Width           =   1725
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   13440
         MaxLength       =   99
         TabIndex        =   47
         Tag             =   "标准代码"
         Text            =   "HPT/B/LT-TR"
         Top             =   120
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7200
         MaxLength       =   99
         TabIndex        =   46
         Tag             =   "标准代码"
         Text            =   "热轧后"
         Top             =   120
         Width           =   1485
      End
      Begin VB.ComboBox cbx_flag 
         Height          =   300
         Left            =   4320
         TabIndex        =   45
         Top             =   120
         Width           =   1455
      End
      Begin VB.ComboBox cbx_ord 
         Height          =   300
         ItemData        =   "CGD2061C.frx":1867
         Left            =   1320
         List            =   "CGD2061C.frx":1869
         TabIndex        =   44
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   99
         TabIndex        =   43
         Tag             =   "标准代码"
         Text            =   "水"
         Top             =   840
         Width           =   1725
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         MaxLength       =   99
         TabIndex        =   42
         Tag             =   "标准代码"
         Text            =   "间隙式水膜法"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10080
         MaxLength       =   99
         TabIndex        =   41
         Tag             =   "标准代码"
         Text            =   "轧制表面/热处理表面"
         Top             =   120
         Width           =   2085
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10080
         MaxLength       =   99
         TabIndex        =   40
         Tag             =   "标准代码"
         Text            =   "NISCO-2800-1"
         Top             =   480
         Width           =   2085
      End
      Begin InDate.ULabel ULabel24 
         Height          =   315
         Left            =   12240
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "试块尺寸"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   120
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "订单号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   3120
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "牌号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   8880
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "表面状态"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   3120
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "晶片尺寸"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   8880
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "试块型号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   120
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "耦合剂"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   12240
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "仪器型号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   120
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "探头型号"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   6000
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "探头频率"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   3120
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "耦合方式"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel25 
         Height          =   315
         Left            =   6000
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "检测方式"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   8880
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "检测比例"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   12240
         Top             =   840
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "扫查方向"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   120
         Top             =   1200
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "检测灵敏度"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   3120
         Top             =   1200
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "表面补偿"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   6000
         Top             =   1200
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "检测标准"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   8880
         Top             =   1200
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "验收级别"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin InDate.ULabel ULabel33 
         Height          =   315
         Left            =   12240
         Top             =   1200
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "审核员"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   3120
         Top             =   1560
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "批准员"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   120
         Top             =   1560
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "级别"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   6000
         Top             =   120
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         Caption         =   "检测时机"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
Attribute VB_Name = "CGD2061C"
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
'-- Program Name      探伤实绩查询界面
'-- Program ID        AGC2041C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2005.7.22
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

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_STDSPEC = 1
Const SS1_LOC = 2
Const SS1_PLATE_NO = 3
Const SS1_OUT_SHEET_NO = 5
Const SS1_PROD_SIZE = 6
Const SS1_CNT = 7
Const SS1_BEF_APLY_STDSPEC = 8
Const SS1_STDSPEC_UPD_FL1 = 9
Const SS1_STDSPEC_UPD_FL2 = 10
Const SS1_UST_LEN = 11
Const SS1_PLATE_CUT_REASON = 12
Const SS1_UST_DEC = 13
Const SS1_PROD_GRD = 14
Const SS1_SURF_GRD = 15
Const SS1_OLD_WGT = 16
Const SS1_WGT = 17
Const SS1_UST_MACHINE_NO = 18
Const SS1_UST_HEAD_KIND = 19
Const SS1_UST_METHOD = 20
Const SS1_UST_STATESCOPE = 21
Const SS1_UST_FL = 22
Const SS1_UST_END_DATE = 23
Const SS1_EMP_CD = 24
Const SS1_UST_MAN = 25
Const SS1_PROD_DATE = 26
Const SS1_REMARK = 27
Const SS1_SHIFT = 28
Const SS1_GROUP_CD = 29
Const SS1_BED_PILE_DATE = 30
Const SS1_SEQ_PLACE = 31
Const SS1_ORD = 32
Const SS1_THK = 34
Const SS1_PRC_LINE = 35
Const SS1_STLGRD_CD = 36
Const SS1_STLGRD = 37
Const SS1_PROC_CD = 38
Const SS1_CUST_CD = 39
Const SS1_SIZE = 40
Const SS1_URGNT_FL = 45     '紧急订单绿色标记 2012-08-16  by  LiQian
Const SS1_IMP_CONT = 54
Const SS1_ORD_REPORT = 56
Const SS1_STD_FLAG = 57 '牌号

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_UST_STAND_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CBO_UST_DEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_UST_THK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_UST_THK_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_UST_WID, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_UST_WID_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(SDB_UST_LEN, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(SDB_UST_LEN_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
          Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_CO_CD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_f_addr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(TXT_ADDR(0), "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(TXT_ADDR(1), "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'          Call Gp_Ms_Collection(TXT_ADDR(2), "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(SDB_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CBO_EMP, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'        Call Gp_Ms_Collection(SDT_PROD_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'      Call Gp_Ms_Collection(SDT_PROD_DATETO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CBO_SURFGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' 紧急订单绿色标记 2012-08-16  by  LiQian
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2012-08-24 炉座号
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2012-08-24 在炉时间
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2012-08-24 出炉温度
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2012-08-30 板坯切割时间
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2013-04-16 精轧结束时间
    Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2013-04-16 轧制班别
    Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2013-04-27 加热1段驻段时间
    Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) ' Add by LiQian at 2013-04-27 均热段驻段时间
    Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
  
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGD2061C.P_SREFER", Key:="P-R"
    sc1.Add Item:="CGD2061C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0

End Sub

Private Sub CBO_UST_DEC_Click()
   Select Case CBO_UST_DEC.ListIndex
          Case 1
               CBO_UST_DEC.Text = "Y"
          Case 2
               CBO_UST_DEC.Text = "N"
   End Select
End Sub

Private Sub chk_Cond_B_Click()

    If chk_Cond_B Then
        TXT_CO_CD.Text = chk_Cond_B.Tag
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        chk_Cond_W = False
        chk_Cond_J = False
    End If
    
    If chk_Cond_B = False And chk_Cond_W = False And chk_Cond_J = False Then
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        TXT_CO_CD.Text = ""
    End If
    
End Sub

Private Sub chk_Cond_W_Click()

    If chk_Cond_W Then
        TXT_CO_CD.Text = chk_Cond_W.Tag
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        chk_Cond_B = False
        chk_Cond_J = False
    End If
    
    If chk_Cond_B = False And chk_Cond_W = False And chk_Cond_J = False Then
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        TXT_CO_CD.Text = ""
    End If
    
End Sub
Private Sub chk_Cond_J_Click()

    If chk_Cond_J Then
        TXT_CO_CD.Text = chk_Cond_J.Tag
        SSCommand1.Enabled = False
        SSCommand2.Enabled = False
        chk_Cond_B = False
        chk_Cond_W = False
    End If
    
    If chk_Cond_B = False And chk_Cond_W = False And chk_Cond_J = False Then
        SSCommand1.Enabled = True
        SSCommand2.Enabled = True
        TXT_CO_CD.Text = ""
    End If
    
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Len(TXT_MAT_NO.Text) >= 8 Then
           Call Form_Ref
        End If
'        KeyAscii = 0
'        SendKeys "{TAB}"
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
    SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If

End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

     If Gf_Ms_Paste(M_CN1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
     End If

End Sub

Public Sub Form_Ref()
    
    Dim iCount          As Integer
    Dim bCount          As Integer
    Dim dMillCal_Wgt    As Double
    Dim simpcont As String
    Dim ord_no As String
    Dim std_flag As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    dMillCal_Wgt = 0
    With ss1
        If .MaxRows = 0 Then
            SDB_WGT.Value = 0
            Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .ROW = iCount
            .Col = SS1_WGT
             If .Value = 0 Then
                .Col = SS1_OLD_WGT
                 dMillCal_Wgt = dMillCal_Wgt + .Value
             Else
                 dMillCal_Wgt = dMillCal_Wgt + .Value
             End If
             
              '紧急订单绿色标记 2012-08-16  by  LiQian
            .Col = SS1_URGNT_FL
            If .Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SS1_STDSPEC, .MaxRows, .ROW, .ROW, &HC000&)
            End If
            
            .ROW = iCount:
            .Col = SS1_IMP_CONT:    simpcont = Trim(.Text)
            If simpcont = "Y" Then
              Call Gp_Sp_BlockColor(ss1, SS1_PLATE_NO, SS1_PLATE_NO, iCount, iCount, SSP4.BackColor)
              Call Gp_Sp_BlockColor(ss1, SS1_IMP_CONT, SS1_IMP_CONT, iCount, iCount, SSP4.BackColor)
            End If
            
            'EDIT HANCHAO 20171027
            '下面这段代码逻辑是循环遍历SPREAD列表，取出订单数据，同时每一个订单数据在COMBOBOX选项框中进行遍历，当该选项框中没有这个订单号时，将这个订单号添加进去
            '定义一个控制值
            Dim flag1 As Boolean
            Dim flag2 As Boolean
            '默认为真
            flag1 = True
            flag2 = True
            '获取订单号
            .Col = SS1_ORD_REPORT: ord_no = .Text
            '循环遍历选项框中的数据
            For bCount = 0 To cbx_ord.ListCount - 1
            '当选项框中存在和列表中相同的订单数据时
            If ord_no = cbx_ord.List(bCount) Then
            '将控制器设置为假
            flag1 = False
            '退出循环
            Exit For
            End If
            Next bCount
            
            .Col = SS1_STD_FLAG: std_flag = .Text
             '循环遍历选项框中的数据
            For bCount = 0 To cbx_flag.ListCount - 1
            '当选项框中存在和列表中相同的牌号数据时
            If std_flag = cbx_flag.List(bCount) Then
            '将控制器设置为假
            flag2 = False
            '退出循环
            Exit For
            End If
            Next bCount
            
            '只有控制器为真的时候才会添加数据
            If flag1 Then
            cbx_ord.AddItem (ord_no)
            End If
            If flag2 Then
            cbx_flag.AddItem (std_flag)
            End If
        Next iCount
    End With
    SDB_WGT.Value = dMillCal_Wgt
               
End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    If ss1.MaxRows < 1 Then Exit Sub
    
    If ROW = 0 Then 'And (Col = 1 Or Col = 2 Or Col = 3 Or Col = 4) Then
        Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0
    End If

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If

End Sub

Private Sub SSCommand1_Click()

    If Trim(TXT_UST_STAND_NO.Text) = "" Then
        Call Gp_MsgBoxDisplay(TXT_UST_STAND_NO.Tag & "必须输入", "", "错误提示")
        Exit Sub
    End If
    
    Call ExcelPrn
    
End Sub

Private Sub SSCommand2_Click()

    If Trim(TXT_UST_STAND_NO.Text) = "" Then
        Call Gp_MsgBoxDisplay(TXT_UST_STAND_NO.Tag & "必须输入", "", "错误提示")
        Exit Sub
    End If
    
    Call ExcelPrn_Pile

End Sub




Private Sub TXT_STDSPEC_CD_Change()
    If Len(Trim(TXT_STDSPEC_CD)) = TXT_STDSPEC_CD.MaxLength Then
       TXT_STDSPEC.Text = Gf_ComnNameFind(M_CN1, "G0018", Trim(TXT_STDSPEC_CD.Text), 1)
    End If
End Sub

Private Sub TXT_STDSPEC_Change()
    If Len(TXT_STDSPEC.Text) = 0 Then
       TXT_STDSPEC_CD.Text = ""
    End If
End Sub

Private Sub txt_stdspec_DblClick()
    DD.sWitch = "MS"
    DD.sKey = "G0018"
    DD.rControl.Add Item:=TXT_STDSPEC_CD

    DD.nameType = "2"
    
    Call Gf_Common_DD(M_CN1, vbKeyF4)
End Sub

Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STDSPEC
        Call Gf_StdSPEC_DD2(M_CN1, vbKeyF4)
        Exit Sub
    End If
End Sub

Private Sub TXT_UST_STAND_NO_Change()
    If Len(TXT_UST_STAND_NO.Text) = 4 Then
       TXT_UST_STAND_NAME.Text = Gf_ComnNameFind(M_CN1, "Q0046", TXT_UST_STAND_NO.Text, 1)
    End If
End Sub

Private Sub TXT_UST_STAND_NO_dblClick()

    DD.sWitch = "MS"
    DD.sKey = "Q0046"
    DD.rControl.Add Item:=TXT_UST_STAND_NO

    DD.nameType = "2"
    
    Call Gf_Mill_Common_DD(M_CN1, vbKeyF4)
    
End Sub


Private Sub ExcelPrn()
    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sDate           As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\CGD2061C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = SDT_PROD_DATE_FROM.Text
    
    If SDT_PROD_DATE_FROM.Text <> SDT_PROD_DATE_TO.Text Then
        xlApp.Range("D2").Value = "日期: " & Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日 - " + Mid(SDT_PROD_DATE_TO.Text, 9, 2) + "日"
    Else
        xlApp.Range("D2").Value = "日期: " & Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日"
    End If
 
    ss1.ROW = 1
    ss1.Col = SS1_UST_MACHINE_NO:     xlApp.Range("A4").Value = ss1.Text
    ss1.Col = SS1_UST_HEAD_KIND:      xlApp.Range("B4").Value = ss1.Text
    ss1.Col = SS1_UST_METHOD:         xlApp.Range("C4").Value = ss1.Text
    ss1.Col = SS1_UST_STATESCOPE:     xlApp.Range("D4").Value = ss1.Text
    ss1.Col = SS1_UST_FL:             xlApp.Range("G4").Value = ss1.Text
    
    Clipboard.Clear
    ss1.SetSelection SS1_STDSPEC, 1, SS1_PLATE_NO, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("A7").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection SS1_PROD_SIZE, 1, SS1_STDSPEC_UPD_FL1, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("D7").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss1.SetSelection SS1_UST_DEC, 1, SS1_UST_DEC, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("H7").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear

    For i = 7 To ss1.MaxRows + 6
        If xlApp.Range("C" & i).Value <> "" Then
           xlApp.Range("E" & i).Value = "1"
        End If
    Next i
    
    xlApp.Range("I2").Select
    xlApp.ActiveSheet.Paste
    
'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
'     xlApp.Application.Visible = False
'     xlSheet.Close False
'     xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub ExcelPrn_Pile()

    Dim i               As Integer
    Dim j               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim sDate           As String
    Dim sShift          As String
    
    Dim sPage_Num       As Integer
    Dim sPage_X         As Integer
    Dim sPage           As Double
    Dim sLastPage       As Double
    Dim sRow1           As Integer
    Dim sRow2           As Integer
    
    Dim xl_A            As String
    Dim xl_B            As String
    Dim xl_C            As String
    Dim xl_D            As String
    Dim xl_E            As String
    Dim xl_F            As String
    Dim xl_G            As String
    Dim xl_H            As String
    Dim xl_I            As String
    Dim xl_J            As String
    
    Dim xl_clr_body     As String
    Dim xl_clr_sum      As String
    Dim xl_clr_spc      As String
    
    Dim Xl_Cnt          As String
    Dim Xl_Wgt          As String
    Dim Xl_Wgt_Val      As String
    Dim Xl_Ust          As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\CGD2063C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = SDT_PROD_DATE_FROM.Text
    
    If SDT_PROD_DATE_FROM.Text <> SDT_PROD_DATE_TO.Text Then
        xlApp.Range("A2").Value = "日期: " & Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日 - " + Mid(SDT_PROD_DATE_TO.Text, 9, 2) + "日"
    Else
        xlApp.Range("A2").Value = "日期: " & Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日"
    End If
    
    If CBO_SHIFT.Text = "1" Then
       sShift = "大夜班"
    ElseIf CBO_SHIFT.Text = "2" Then
       sShift = "白班"
    ElseIf CBO_SHIFT.Text = "3" Then
       sShift = "小夜班"
    Else
       sShift = ""
    End If
    
    xlApp.Range("C2").Value = Mid(xlApp.Range("C2").Value, 1, 3) & sShift
    '''ADDED BY GUOLI AT 20080419123200
    '  COMMENT BY YANGMENG AT 20080420123200
'    ss1.Col = 20: ss1.Row = 1
'    xlApp.Range("D2").Value = "码堆员：" & ss1.Text
    ''''''''''''''''''''''''''''''''''''''
    'ADDED BY YANGMENG AT 20080420123200
    xlApp.Range("E2").Value = "码堆员：" & sUserName
        
    sPage_Num = 30
    sPage_X = 32
    
    sPage = Int(ss1.MaxRows / sPage_Num) + 1
    sLastPage = ss1.MaxRows - Int(ss1.MaxRows / sPage_Num) * sPage_Num
    
    For i = 0 To 9
        xl_clr_body = "A" + CStr(4 + i * sPage_X) + ":" + "I" + CStr(33 + i * sPage_X)
        xl_clr_sum = "C" + CStr(34 + i * sPage_X) + ":" + "C" + CStr(35 + i * sPage_X)
        xl_clr_spc = "E" + CStr(34 + i * sPage_X)
        xlApp.Range(xl_clr_body).Value = Null
        xlApp.Range(xl_clr_sum).Value = Null
        xlApp.Range(xl_clr_spc).Value = Mid(xlApp.Range(xl_clr_spc).Value, 1, 5)
    Next i
    
    For i = 0 To sPage - 1
       
        sRow1 = 1 + sPage_Num * i
        sRow2 = sPage_Num * (i + 1)

        If i = sPage - 1 Then
           sRow2 = sPage_Num * i + sLastPage
        End If

        xl_A = "A" + CStr(4 + i * sPage_X)
        xl_B = "B" + CStr(4 + i * sPage_X)
        xl_C = "C" + CStr(4 + i * sPage_X)
        xl_D = "D" + CStr(4 + i * sPage_X)
        xl_E = "E" + CStr(4 + i * sPage_X)
        xl_F = "F" + CStr(4 + i * sPage_X)
        xl_G = "G" + CStr(4 + i * sPage_X)
        xl_H = "H" + CStr(4 + i * sPage_X)
        xl_I = "I" + CStr(4 + i * sPage_X)
        xl_J = "J" + CStr(4 + i * sPage_X)
        
        Xl_Cnt = "C" + CStr(3 + (i + 1) * sPage_X - 1)
        Xl_Wgt = "C" + CStr(3 + (i + 1) * sPage_X)
        Xl_Ust = "E" + CStr(3 + (i + 1) * sPage_X - 1)
        
        Clipboard.Clear
        ss1.SetSelection 2, sRow1, 2, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_A).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_STDSPEC, sRow1, SS1_STDSPEC, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_B).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_PROC_CD, sRow1, SS1_PROC_CD, sRow2  ''ss1.SetSelection 3, sRow1, 3, sRow2  PLATE_NO-->LOT_NO MODIFIED BY GUOLI AT 20080505193800
        ss1.ClipboardCopy
        xlApp.Range(xl_C).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_PLATE_NO, sRow1, SS1_PLATE_NO, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_D).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_PROD_SIZE, sRow1, SS1_PROD_SIZE, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_E).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_WGT, sRow1, SS1_WGT, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_F).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_UST_DEC, sRow1, SS1_UST_DEC, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_G).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_CUST_CD, sRow1, SS1_CUST_CD, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_H).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_ORD, sRow1, SS1_ORD, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_I).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_SIZE, sRow1, SS1_SIZE, sRow2
        ss1.ClipboardCopy
        xlApp.Range(xl_J).Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        If i = sPage - 1 Then
           xlApp.Range(Xl_Cnt).Value = sLastPage
        Else
           xlApp.Range(Xl_Cnt).Value = sPage_Num
        End If
        
        For j = 1 To sPage_Num
            Xl_Wgt_Val = "F" & CStr((Val(Mid(xl_F, 2)) + j - 1))
            xlApp.Range(Xl_Wgt).Value = xlApp.Range(Xl_Wgt).Value + xlApp.Range(Xl_Wgt_Val).Value
        Next j
        
        ss1.ROW = 1
        ss1.Col = SS1_UST_FL
        xlApp.Range(Xl_Ust).Value = xlApp.Range(Xl_Ust).Value + ss1.Text
              
    Next i
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub



Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub
Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1

End Sub
Private Sub txt_f_addr_DblClick()
     Call txt_f_addr_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_f_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "F0009"
        txt_f_addr.Text = "P"
        DD.rControl.Add Item:=txt_f_addr
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

End Sub
Private Sub SDT_PROD_DATE_TO_GotFocus()
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()

     If SDT_PROD_DATE_FROM.RawData = "" Then
        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     End If
     
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
     
End Sub


Private Sub TXT_UST_STAND_REPORT_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0046"
        DD.rControl.Add Item:=TXT_UST_STAND_REPORT
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

End Sub

Private Sub TXT_UST_STAND_REPORT_DblClick()
     Call TXT_UST_STAND_REPORT_KeyUp(vbKeyF4, 0)
End Sub

Private Sub TXT_UST_GRADE_DblClick()

    Call TXT_UST_GRADE_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub TXT_UST_GRADE_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0053"
        DD.rControl.Add Item:=TXT_UST_GRADE
    
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, vbKeyF4)
    
    End If

End Sub


