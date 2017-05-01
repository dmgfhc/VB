VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0060C 
   Caption         =   "标准交付条件查询_AQA0060C"
   ClientHeight    =   9090
   ClientLeft      =   210
   ClientTop       =   1860
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_PROD_NAME 
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
      Left            =   7710
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   90
      Width           =   1725
   End
   Begin VB.TextBox txt_PROD_CD 
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
      Left            =   7110
      TabIndex        =   33
      Top             =   90
      Width           =   540
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4440
      Left            =   135
      ScaleHeight     =   4440
      ScaleWidth      =   15150
      TabIndex        =   1
      Top             =   4740
      Width           =   15150
      Begin VB.TextBox txt_LEN_TOL_UNIT 
         Height          =   300
         Left            =   8685
         TabIndex        =   13
         Top             =   3630
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txt_WID_TOL_UNIT 
         Height          =   300
         Left            =   8205
         TabIndex        =   12
         Top             =   3630
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txt_THK_TOL_UNIT 
         Height          =   300
         Left            =   7755
         TabIndex        =   11
         Top             =   3630
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txt_RECT_NAME 
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3540
         Width           =   975
      End
      Begin VB.TextBox txt_FLT_NAME 
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2685
         Width           =   975
      End
      Begin VB.TextBox txt_STRT_NAME 
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2265
         Width           =   975
      End
      Begin VB.ComboBox cbo_LEN_TOL_UNIT 
         Height          =   300
         ItemData        =   "AQA0060C.frx":0000
         Left            =   6224
         List            =   "AQA0060C.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1410
         Width           =   1215
      End
      Begin VB.ComboBox cbo_WID_TOL_UNIT 
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
         ItemData        =   "AQA0060C.frx":0016
         Left            =   6224
         List            =   "AQA0060C.frx":0020
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   990
         Width           =   1215
      End
      Begin VB.ComboBox cbo_THK_TOL_UNIT 
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
         ItemData        =   "AQA0060C.frx":002C
         Left            =   6224
         List            =   "AQA0060C.frx":0036
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   555
         Width           =   1215
      End
      Begin VB.TextBox txt_RECT_TYP 
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
         Left            =   1646
         MaxLength       =   1
         TabIndex        =   7
         Top             =   3540
         Width           =   345
      End
      Begin VB.TextBox txt_FLT_TYP 
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
         Left            =   1646
         MaxLength       =   1
         TabIndex        =   6
         Top             =   2685
         Width           =   345
      End
      Begin VB.TextBox txt_STRT_KND 
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
         Left            =   1646
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2265
         Width           =   345
      End
      Begin InDate.ULabel ul_INS_DATE 
         Height          =   300
         Left            =   1380
         Top             =   4080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   0
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
         Index           =   1
         Left            =   120
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "项目"
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
         Index           =   2
         Left            =   120
         Top             =   555
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "公差"
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
         Index           =   3
         Left            =   1646
         Top             =   990
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "宽度"
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
         Index           =   4
         Left            =   1646
         Top             =   1410
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "长度"
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
         Index           =   5
         Left            =   1646
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "类型"
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
         Index           =   6
         Left            =   3172
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "下限值"
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
         Index           =   7
         Left            =   4698
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "上限值"
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
         Index           =   8
         Left            =   6224
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "单位"
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
         Index           =   10
         Left            =   120
         Top             =   2685
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "不平度"
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
      Begin InDate.ULabel ULabel1 
         Height          =   300
         Index           =   11
         Left            =   120
         Top             =   3120
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "塔形度"
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
         Index           =   12
         Left            =   120
         Top             =   3540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "直角度"
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
      Begin InDate.ULabel ULabel1 
         Height          =   300
         Index           =   13
         Left            =   7750
         Top             =   555
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "同板差"
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
         Index           =   14
         Left            =   7750
         Top             =   1050
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "楔形度"
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
         Index           =   15
         Left            =   7750
         Top             =   1545
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "凸度"
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
         Index           =   16
         Left            =   7750
         Top             =   2055
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "波浪度"
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
         Index           =   18
         Left            =   9276
         Top             =   3045
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "宽度方向"
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
         Index           =   19
         Left            =   9276
         Top             =   3540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "长度方向"
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
         Index           =   20
         Left            =   7750
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "项目"
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
         Index           =   21
         Left            =   9276
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "类型"
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
         Index           =   22
         Left            =   10802
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "下限值"
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
         Index           =   23
         Left            =   12328
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "上限值"
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
         Index           =   24
         Left            =   13860
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "单位"
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
         Index           =   25
         Left            =   9276
         Top             =   2550
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "厚度方向"
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
         Index           =   9
         Left            =   120
         Top             =   2265
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "镰刀弯"
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
      Begin InDate.ULabel ULabel1 
         Height          =   300
         Index           =   17
         Left            =   7750
         Top             =   2550
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "切斜"
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
         Index           =   26
         Left            =   3172
         Top             =   1830
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "总上限值"
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
         Index           =   27
         Left            =   4698
         Top             =   1830
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "每单位上限值"
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
         Index           =   28
         Left            =   6224
         Top             =   1830
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "单位"
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
         Index           =   29
         Left            =   120
         Top             =   4080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "编制日期"
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
         Index           =   30
         Left            =   2820
         Top             =   4080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "编制人"
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
         Index           =   31
         Left            =   6210
         Top             =   4080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "修改日期"
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
         Index           =   32
         Left            =   8910
         Top             =   4080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "修改人"
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
         Index           =   33
         Left            =   1646
         Top             =   555
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   "厚度"
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
      Begin InDate.ULabel ul_INS_EMP 
         Height          =   300
         Left            =   4080
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   0
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
      Begin InDate.ULabel ul_UPD_DATE 
         Height          =   300
         Left            =   7470
         Top             =   4080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   0
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
      Begin InDate.ULabel ul_UPD_EMP 
         Height          =   300
         Left            =   10170
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         Caption         =   ""
         Alignment       =   0
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
         Index           =   34
         Left            =   6674
         Top             =   3540
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   35
         Left            =   14310
         Top             =   1054
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   36
         Left            =   14310
         Top             =   1551
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   37
         Left            =   14310
         Top             =   2048
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   38
         Left            =   14310
         Top             =   2545
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   39
         Left            =   14310
         Top             =   557
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   40
         Left            =   6674
         Top             =   3120
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   41
         Left            =   14310
         Top             =   3042
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   42
         Left            =   14310
         Top             =   3540
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   43
         Left            =   6674
         Top             =   2685
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
         Index           =   44
         Left            =   6674
         Top             =   2265
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         Caption         =   "mm"
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
      Begin CSTextLibCtl.sidbEdit sdb_WID_TOL_MIN 
         Height          =   300
         Left            =   3180
         TabIndex        =   15
         Top             =   990
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_LEN_TOL_MIN 
         Height          =   300
         Left            =   3180
         TabIndex        =   16
         Top             =   1380
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_THK_TOL_MAX 
         Height          =   300
         Left            =   4710
         TabIndex        =   17
         Top             =   570
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_WID_TOL_MAX 
         Height          =   300
         Left            =   4710
         TabIndex        =   18
         Top             =   990
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_LEN_TOL_MAX 
         Height          =   300
         Left            =   4710
         TabIndex        =   19
         Top             =   1380
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_RECT_MAX 
         Height          =   300
         Left            =   3180
         TabIndex        =   20
         Top             =   3540
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_STRT_MAX 
         Height          =   300
         Left            =   3180
         TabIndex        =   21
         Top             =   2280
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_FLT_MAX 
         Height          =   300
         Left            =   3180
         TabIndex        =   22
         Top             =   2700
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         AutoScroll      =   0   'False
         BorderEffect    =   2
         DataProperty    =   2
         FocusSelect     =   -1  'True
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.0"
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
         NumDecDigits    =   1
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TELSCOPE_MAX 
         Height          =   300
         Left            =   3180
         TabIndex        =   23
         Top             =   3120
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_STRT_UNIT_MAX 
         Height          =   300
         Left            =   4710
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_FLT_UNIT_MAX 
         Height          =   300
         Left            =   4710
         TabIndex        =   25
         Top             =   2700
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_TELSCOPE_UNIT_MAX 
         Height          =   300
         Left            =   4710
         TabIndex        =   26
         Top             =   3120
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_SAPE_WAVE 
         Height          =   300
         Left            =   12330
         TabIndex        =   27
         Top             =   2040
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_THK_SHR_DRAG_MAX 
         Height          =   300
         Left            =   12330
         TabIndex        =   28
         Top             =   2580
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_WID_SHR_DRAG_MAX 
         Height          =   300
         Left            =   12330
         TabIndex        =   29
         Top             =   3060
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_THK_DVT_MAX 
         Height          =   300
         Left            =   12330
         TabIndex        =   30
         Top             =   540
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_VRT_MAX 
         Height          =   300
         Left            =   12330
         TabIndex        =   31
         Top             =   1020
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_CROWN_MAX 
         Height          =   300
         Left            =   12330
         TabIndex        =   32
         Top             =   1530
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   1
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_THK_TOL_MIN 
         Height          =   315
         Left            =   3180
         TabIndex        =   35
         Top             =   570
         Width           =   1200
         _Version        =   262145
         _ExtentX        =   2117
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sdb_LEN_SHR_DRAG_MAX 
         Height          =   300
         Left            =   12330
         TabIndex        =   36
         Top             =   3540
         Width           =   1215
         _Version        =   262145
         _ExtentX        =   2143
         _ExtentY        =   529
         _StockProps     =   125
         Text            =   " 0.00"
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
         NumIntDigits    =   2
         ShowZero        =   0   'False
         Undo            =   0
         Data            =   0
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   7560
         X2              =   7560
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Line Line2 
         Index           =   8
         X1              =   13710
         X2              =   13710
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Line Line2 
         Index           =   7
         X1              =   10650
         X2              =   10650
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Line Line2 
         Index           =   6
         X1              =   9120
         X2              =   9120
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Line Line2 
         Index           =   5
         X1              =   12210
         X2              =   12210
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   0
         X2              =   15195
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line2 
         Index           =   4
         X1              =   6060
         X2              =   6060
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   3030
         X2              =   3030
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   1500
         X2              =   1500
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   0
         X2              =   7550
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   0
         X2              =   7550
         Y1              =   1770
         Y2              =   1770
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   4560
         X2              =   4560
         Y1              =   0
         Y2              =   3960
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         Height          =   3975
         Left            =   0
         Top             =   -30
         Width           =   15195
      End
   End
   Begin VB.TextBox txt_DEV_STD_CD_P 
      Height          =   300
      Left            =   2460
      MaxLength       =   18
      TabIndex        =   0
      Top             =   90
      Width           =   2295
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   0
      Left            =   150
      Top             =   90
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   529
      Caption         =   "代表性交付条件标准编号"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   4215
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   15195
      _Version        =   393216
      _ExtentX        =   26802
      _ExtentY        =   7435
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   44
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0060C.frx":0042
   End
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   14
      Left            =   5190
      Top             =   90
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   529
      Caption         =   "产品"
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
End
Attribute VB_Name = "AQA0060C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量标准管理
'-- Program Name      标准交付条件输入
'-- Program ID        AQA0060C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       标准交付条件输入
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

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim oRd_cnt As Integer              'Select Order Count
Dim bCheck As Boolean
Dim bCheck1 As Boolean
Dim bCheck2 As Boolean
Dim bCheck3 As Boolean
Dim lCopyRow As Long                'Copy Row

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_DEV_STD_CD_P, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_PROD_CD, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0060C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AQA0060C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:="AQA0060C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
         
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        
        Case "txt_PROD_CD"
            sCode = "B0005"
            Set oCodeName = txt_PROD_NAME
        
        Case "txt_STRT_KND"
            sCode = "Q0030"
            Set oCodeName = txt_STRT_NAME
            
        Case "txt_FLT_TYP"
            sCode = "Q0029"
            Set oCodeName = txt_FLT_NAME
            
        Case "txt_RECT_TYP"
            sCode = "Q0041"
            Set oCodeName = txt_RECT_NAME
            
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub



Private Sub cbo_THK_TOL_UNIT_Click()

    bCheck1 = True
        
    Select Case cbo_THK_TOL_UNIT.ListIndex
        Case 0
            txt_THK_TOL_UNIT.Text = "A"
        Case 1
            txt_THK_TOL_UNIT.Text = "B"
        Case Else
            txt_THK_TOL_UNIT.Text = "A"
    End Select
End Sub

Private Sub cbo_WID_TOL_UNIT_Click()

    bCheck2 = True
    
    Select Case cbo_WID_TOL_UNIT.ListIndex
        Case 0
            txt_WID_TOL_UNIT.Text = "A"
        Case 1
            txt_WID_TOL_UNIT.Text = "B"
        Case Else
            txt_WID_TOL_UNIT.Text = "A"
    End Select

End Sub

Private Sub cbo_LEN_TOL_UNIT_Click()

    bCheck3 = True
        
    Select Case cbo_LEN_TOL_UNIT.ListIndex
        Case 0
            txt_LEN_TOL_UNIT.Text = "A"
        Case 1
            txt_LEN_TOL_UNIT.Text = "B"
        Case Else
            txt_LEN_TOL_UNIT.Text = "A"
    End Select
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
       
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 1)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 2)
    
'    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call GP_ROW_BACKCOLOR(ss1)
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 1)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 2)
    
    'MDIMain.MenuTool.Buttons
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing

    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Ins()
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 39)
    'Call Spread_to_Master(ss1, ss1.ActiveRow)
End Sub

Public Sub Form_Pro()

    Dim iRow As Long
    Dim i As Integer
    
    'iRow = ss1.ActiveRow
    iRow = ss1.Row
    
    If Value_check() Then
        If Gf_Mc_Authority(sAuthority, Mc1) Then
            
             If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
                Call GP_SELECT_ROW(ss1, iRow)
             End If
             
             ss1.Row = iRow
    
        End If
    End If
    
End Sub

Public Sub Form_Del()

    If Not Gf_Ms_AllDel(M_CN1, Proc_Sc("Sc"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub
Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call MS_Cls
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        pControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        bCheck = False
        Call Spread_to_Master(ss1, 1)
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call GP_SELECT_ROW(ss1, 1)
        ss1.Row = 1
        Exit Sub
    End If
                
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub





Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
    If Gf_Sc_Authority(sAuthority, "U") Then

        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), 0)

        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 39)

    End If
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim i As Integer

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    bCheck = False
    
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub



Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 39)
    End If
End Sub



Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sTemp_Code As String
    Dim iCol As Integer

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

   Select Case ss1.ActiveCol

        Case 1

            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1

                DD.sWitch = "SP"
                DD.rControl.Add Item:=1
                DD.nameType = "2"
                    
                Call Gf_STD_DELV_DD(M_CN1, KeyCode)
        
            End If
            
       Case 2

            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1

                DD.sWitch = "SP"
                DD.sKey = "B0005"
                DD.rControl.Add Item:=2
                DD.rControl.Add Item:=3
                DD.nameType = "2"
                    
                Call Gf_Common_DD(M_CN1, KeyCode)
        
            End If

    End Select

   '  ss1.SetFocus
End Sub


Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    bCheck = False
    Call Spread_to_Master(ss1, NewRow)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub Spread_to_Master(ByVal sp As vaSpread, ByVal iRow As Long)
    Dim RowLabel As String
    
        bCheck = False
        bCheck1 = False
        bCheck2 = False
        bCheck3 = False

        With sp
        
            If iRow > 0 Then
                .Row = iRow
               
                .Col = 0: RowLabel = .Text
                .Col = 10: sdb_THK_TOL_MIN.Text = .Text
                .Col = 11: sdb_THK_TOL_MAX.Text = .Text
                .Col = 12: txt_THK_TOL_UNIT.Text = .Text
                .Col = 13: sdb_WID_TOL_MIN.Text = .Text
                .Col = 14: sdb_WID_TOL_MAX.Text = .Text
                .Col = 15: txt_WID_TOL_UNIT.Text = .Text
                .Col = 16: sdb_LEN_TOL_MIN.Text = .Text
                .Col = 17: sdb_LEN_TOL_MAX.Text = .Text
                .Col = 18: txt_LEN_TOL_UNIT.Text = .Text
                .Col = 19: txt_STRT_KND.Text = .Text
                .Col = 20: txt_STRT_NAME.Text = .Text
                .Col = 21: sdb_STRT_MAX.Text = .Text
                .Col = 22: sdb_STRT_UNIT_MAX.Text = .Text
                .Col = 23: txt_FLT_TYP.Text = .Text
                .Col = 24: txt_FLT_NAME.Text = .Text
                .Col = 25: sdb_FLT_MAX.Text = .Text
                .Col = 26: sdb_FLT_UNIT_MAX.Text = .Text
                .Col = 27: sdb_TELSCOPE_MAX.Text = .Text
                .Col = 28: sdb_TELSCOPE_UNIT_MAX.Text = .Text
                .Col = 29: txt_RECT_TYP.Text = .Text
                .Col = 30: sdb_RECT_MAX.Text = .Text
                .Col = 31: sdb_THK_DVT_MAX.Text = .Text
                .Col = 32: sdb_VRT_MAX.Text = .Text
                .Col = 33: sdb_CROWN_MAX.Text = .Text
                .Col = 34: sdb_SAPE_WAVE.Text = .Text
                .Col = 35: sdb_THK_SHR_DRAG_MAX.Text = .Text
                .Col = 36: sdb_WID_SHR_DRAG_MAX.Text = .Text
                .Col = 37: sdb_LEN_SHR_DRAG_MAX.Text = .Text
                .Col = 38: ul_INS_DATE.Caption = .Text
                .Col = 40: ul_INS_EMP.Caption = .Text
                .Col = 41: ul_UPD_DATE.Caption = .Text
                .Col = 43: ul_UPD_EMP.Caption = .Text
                .Col = 44: txt_RECT_NAME.Text = .Text
                If RowLabel = "Input" Then
'                    txt_MLT_STD_NO.Locked = False
'                    txt_APP_DATE.Enabled = True
                Else
'                    txt_MLT_STD_NO.Locked = True
'                    txt_APP_DATE.Enabled = False
                End If
            Else
                Exit Sub
            End If
        
        End With

End Sub

Public Sub Spread_Can()
    
    Call GP_SELECT_ROW(ss1, ss1.Row)
    Call GP_ROW_CANCEL(Proc_Sc("SC"))
    'Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
      
End Sub

Public Sub Spread_Del()
    
    Call GP_SET_CELL_VALUE(ss1, ss1.Row, 0, "Delete")
    'Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Public Sub Spread_Cpy()

    lCopyRow = ss1.ActiveRow
    'Call Gp_Sp_Copy(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Pst()
    Call GP_ROW_PASTE(Proc_Sc("Sc"), lCopyRow)
    'Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 39)
End Sub

Public Sub Ms_To_SP(ByVal sp As vaSpread, ByVal iRow As Long, ByVal iCol As Long, vName As String)
    Dim old_Value As Variant
    Dim iValue As Variant
    
    If (vName <> "0") And (vName <> "1") Then
        If TypeName(Me.Controls(vName)) = "TextBox" Then
            iValue = Me.Controls(vName).Text
        End If
        
        If TypeName(Me.Controls(vName)) = "sidbEdit" Then
            iValue = Me.Controls(vName).Value
        End If
        If TypeName(Me.Controls(vName)) = "UDate" Then
            iValue = Me.Controls(vName).Text
        End If
        If TypeName(Me.Controls(vName)) = "sidtEdit" Then
            iValue = Format(Me.Controls(vName).Text, "YYYYMMDD")
        End If
    Else
        iValue = vName
    End If
    
    With sp
        If iCol = 1 Or iCol = 2 Then
            .Row = iRow
            .Col = 0
            If (.Text = "Input") Then
                .Col = iCol
                .Value = iValue
                .Text = iValue
            Else
                Exit Sub
            End If
        Else
            .Row = iRow
            .Col = iCol
            old_Value = .Value
            .Value = iValue
            .Text = iValue
            If old_Value <> .Value Then
                .Col = 0
                    If (.Text = "Input") Or (.Text = "Update") Then
                        .Text = .Text
                    Else
                        .Text = "Update"
                    End If
                    .Col = iCol
            Else
                Exit Sub
            End If
        End If
    End With
End Sub


Private Sub MS_Cls()
    Dim i As Integer
    For i = 0 To Me.COUNT - 1
        If TypeName(Me.Controls(i)) = "TextBox" Or TypeName(Me.Controls(i)) = "sidbEdit" Then
            Me.Controls(i).Text = ""
        End If
        
    Next i
End Sub







Private Sub txt_DEV_STD_CD_P_KeyUp(KeyCode As Integer, Shift As Integer)
            If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_DEV_STD_CD_P
        
        Call Gf_STD_DELV_DD(M_CN1, KeyCode)
        Exit Sub
    
    End If
End Sub


'Private Sub sdb_STRT_UNIT_MAX_KeyPress(KeyAscii As Integer)
'    KeyAscii = txt_KeyPress(KeyAscii, sdb_STRT_UNIT_MAX.Name)
'End Sub


'公差-厚度-下限值 -------------------------------------------------------------------------------------
Private Sub sdb_THK_TOL_MIN_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_THK_TOL_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 10, sdb_THK_TOL_MIN.Name)
    End If
End Sub

Private Sub sdb_THK_TOL_MIN_LostFocus()
    bCheck = False
End Sub

'公差-厚度-上限值 -------------------------------------------------------------------------------------
Private Sub sdb_THK_TOL_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_THK_TOL_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 11, sdb_THK_TOL_MAX.Name)
    End If
End Sub

Private Sub sdb_THK_TOL_MAX_LostFocus()
    bCheck = False
End Sub


Private Sub txt_RECT_NAME_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 44, txt_RECT_NAME.Name)
    End If
End Sub



Private Sub txt_STRT_NAME_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call GP_SET_CELL_VALUE(ss1, ss1.ActiveRow, 20, txt_STRT_NAME.Text)
    End If
End Sub


'公差-厚度-单位 -------------------------------------------------------------------------------------

Private Sub txt_THK_TOL_UNIT_GotFocus()
    bCheck1 = True
End Sub

Private Sub txt_THK_TOL_UNIT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck1 = True Then
        Call Ms_To_SP(ss1, ss1.Row, 12, txt_THK_TOL_UNIT.Name)
    End If
    
    If txt_THK_TOL_UNIT.Text = "B" Then
        cbo_THK_TOL_UNIT.ListIndex = 1
    Else
        cbo_THK_TOL_UNIT.ListIndex = 0
    End If
End Sub

Private Sub txt_THK_TOL_UNIT_LostFocus()
    bCheck1 = False
End Sub


'公差-宽度-下限值 -------------------------------------------------------------------------------------
Private Sub sdb_WID_TOL_MIN_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_WID_TOL_MIN_Change()
    
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 13, sdb_WID_TOL_MIN.Name)
    End If
End Sub

Private Sub sdb_WID_TOL_MIN_LostFocus()
    bCheck = False
End Sub

'公差-宽度-上限值 -------------------------------------------------------------------------------------
Private Sub sdb_WID_TOL_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_WID_TOL_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 14, sdb_WID_TOL_MAX.Name)
    End If
End Sub

Private Sub sdb_WID_TOL_MAX_LostFocus()
    bCheck = False
End Sub

'公差-宽度-单位 -------------------------------------------------------------------------------------
Private Sub txt_WID_TOL_UNIT_GotFocus()
    bCheck2 = True
End Sub

Private Sub txt_WID_TOL_UNIT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck2 = True Then
        Call Ms_To_SP(ss1, ss1.Row, 15, txt_WID_TOL_UNIT.Name)
    End If
    If txt_WID_TOL_UNIT.Text = "B" Then
        cbo_WID_TOL_UNIT.ListIndex = 1
    Else
        cbo_WID_TOL_UNIT.ListIndex = 0
    End If
End Sub

Private Sub txt_WID_TOL_UNIT_LostFocus()
    bCheck2 = False
End Sub

'公差-长度-下限值 -------------------------------------------------------------------------------------

Private Sub sdb_LEN_TOL_MIN_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_LEN_TOL_MIN_LostFocus()
    bCheck = False
End Sub

Private Sub sdb_LEN_TOL_MIN_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 16, sdb_LEN_TOL_MIN.Name)
    End If
End Sub

'公差-长度-上限值 -------------------------------------------------------------------------------------
Private Sub sdb_LEN_TOL_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_LEN_TOL_MAX_LostFocus()
    bCheck = False
End Sub

Private Sub sdb_LEN_TOL_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 17, sdb_LEN_TOL_MAX.Name)
    End If
End Sub

'公差-长度-单位 -------------------------------------------------------------------------------------
Private Sub txt_LEN_TOL_UNIT_GotFocus()
    bCheck3 = True
End Sub

Private Sub txt_LEN_TOL_UNIT_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck3 = True Then
        Call Ms_To_SP(ss1, ss1.Row, 18, txt_LEN_TOL_UNIT.Name)
    End If
    If txt_LEN_TOL_UNIT.Text = "B" Then
        cbo_LEN_TOL_UNIT.ListIndex = 1
    Else
        cbo_LEN_TOL_UNIT.ListIndex = 0
    End If
End Sub

Private Sub txt_LEN_TOL_UNIT_LostFocus()
    bCheck3 = False
End Sub



Private Sub txt_STRT_KND_GotFocus()
    bCheck = True
End Sub

'Private Sub txt_STRT_KND_LostFocus()
'    bCheck = False
'End Sub

Private Sub txt_STRT_KND_Change()
    
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 19, txt_STRT_KND.Name)
        Call GP_SET_CELL_VALUE(ss1, ss1.ActiveRow, 20, txt_STRT_NAME.Text)
    End If
End Sub


'镰刀弯 - 总上限值 -------------------------------------------------------------------------------------
Private Sub sdb_STRT_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_STRT_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 21, sdb_STRT_MAX.Name)
    End If
End Sub

Private Sub sdb_STRT_MAX_LostFocus()
    bCheck = False
End Sub

'镰刀弯 - 每单位上限值 -------------------------------------------------------------------------------------
Private Sub sdb_STRT_UNIT_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_STRT_UNIT_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 22, sdb_STRT_UNIT_MAX.Name)
    End If
End Sub

Private Sub sdb_STRT_UNIT_MAX_LostFocus()
    bCheck = False
End Sub



Private Sub txt_FLT_TYP_GotFocus()
    bCheck = True
End Sub

Private Sub txt_FLT_TYP_Change()
        
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 23, txt_FLT_TYP.Name)
    End If
End Sub

Private Sub txt_FLT_NAME_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 24, txt_FLT_NAME.Name)
    End If
End Sub
'Private Sub txt_FLT_TYP_LostFocus()
'    bCheck = False
'End Sub

'不平度 - 总上限值 -------------------------------------------------------------------------------------
Private Sub sdb_FLT_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_FLT_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 25, sdb_FLT_MAX.Name)
    End If
End Sub

Private Sub sdb_FLT_MAX_LostFocus()
    bCheck = False
End Sub


'不平度 - 每单位上限值 -------------------------------------------------------------------------------------
Private Sub sdb_FLT_UNIT_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_FLT_UNIT_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 26, sdb_FLT_UNIT_MAX.Name)
    End If
End Sub

Private Sub sdb_FLT_UNIT_MAX_LostFocus()
    bCheck = False
End Sub

'塔形度 - 总上限值 -------------------------------------------------------------------------------------
Private Sub sdb_TELSCOPE_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_TELSCOPE_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 27, sdb_TELSCOPE_MAX.Name)
    End If
End Sub

Private Sub sdb_TELSCOPE_MAX_LostFocus()
    bCheck = False
End Sub


'塔形度 - 每单位上限值 -------------------------------------------------------------------------------------
Private Sub sdb_TELSCOPE_UNIT_MAX_GotFocus()
    bCheck = False
End Sub

Private Sub sdb_TELSCOPE_UNIT_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 28, sdb_TELSCOPE_UNIT_MAX.Name)
    End If
End Sub

Private Sub sdb_TELSCOPE_UNIT_MAX_LostFocus()
    bCheck = False
End Sub



Private Sub txt_RECT_TYP_GotFocus()
    bCheck = True
End Sub

Private Sub txt_RECT_TYP_Change()
    
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 29, txt_RECT_TYP.Name)
    End If
End Sub


'Private Sub txt_RECT_TYP_LostFocus()
'    bCheck = False
'End Sub


'直角度 - 总上限值 -------------------------------------------------------------------------------------
Private Sub sdb_RECT_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_RECT_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 30, sdb_RECT_MAX.Name)
    End If
End Sub

Private Sub sdb_RECT_MAX_LostFocus()
    bCheck = False
End Sub

'上限值 - 1 -------------------------------------------------------------------------------------
Private Sub sdb_THK_DVT_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_THK_DVT_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 31, sdb_THK_DVT_MAX.Name)
    End If
End Sub

Private Sub sdb_THK_DVT_MAX_LostFocus()
    bCheck = False
End Sub

'上限值 - 2 -------------------------------------------------------------------------------------
Private Sub sdb_VRT_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_VRT_MAX_LostFocus()
    bCheck = False
End Sub

Private Sub sdb_VRT_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 32, sdb_VRT_MAX.Name)
    End If
End Sub

'上限值 - 3 -------------------------------------------------------------------------------------
Private Sub sdb_CROWN_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_CROWN_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 33, sdb_CROWN_MAX.Name)
    End If
End Sub

Private Sub sdb_CROWN_MAX_LostFocus()
    bCheck = False
End Sub

'上限值 - 4 -------------------------------------------------------------------------------------
Private Sub sdb_SAPE_WAVE_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_SAPE_WAVE_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 34, sdb_SAPE_WAVE.Name)
    End If
End Sub

Private Sub sdb_SAPE_WAVE_LostFocus()
    bCheck = False
End Sub

'上限值 - 5 -------------------------------------------------------------------------------------
Private Sub sdb_THK_SHR_DRAG_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_THK_SHR_DRAG_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 35, sdb_THK_SHR_DRAG_MAX.Name)
    End If
End Sub

Private Sub sdb_THK_SHR_DRAG_MAX_LostFocus()
    bCheck = False
End Sub

'上限值 - 6 -------------------------------------------------------------------------------------

Private Sub sdb_WID_SHR_DRAG_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_WID_SHR_DRAG_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 36, sdb_WID_SHR_DRAG_MAX.Name)
    End If
End Sub

Private Sub sdb_WID_SHR_DRAG_MAX_LostFocus()
    bCheck = False
End Sub

'上限值 - 7 -------------------------------------------------------------------------------------
Private Sub sdb_LEN_SHR_DRAG_MAX_GotFocus()
    bCheck = True
End Sub

Private Sub sdb_LEN_SHR_DRAG_MAX_Change()
    If (ss1.ActiveRow > 0) And (ss1.Row <> 0) And bCheck = True Then
        Call Ms_To_SP(ss1, ss1.Row, 37, sdb_LEN_SHR_DRAG_MAX.Name)
    End If
End Sub

Private Sub sdb_LEN_SHR_DRAG_MAX_LostFocus()
    bCheck = False
End Sub

Private Function Value_check() As Boolean
 
 
    If Not Gf_subValueCheck(sdb_THK_TOL_MIN, sdb_THK_TOL_MAX) Then
       Value_check = False
       Exit Function
    End If
    If Not Gf_subValueCheck(sdb_WID_TOL_MIN, sdb_WID_TOL_MAX) Then
       Value_check = False
       Exit Function
    End If
    If Not Gf_subValueCheck(sdb_LEN_TOL_MIN, sdb_LEN_TOL_MAX) Then
       Value_check = False
       Exit Function
    End If
     
    Value_check = True

End Function
