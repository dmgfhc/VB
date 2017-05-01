VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F3F3F48B-0749-465F-9B1C-1914C440CB19}#1.0#0"; "indate.ocx"
Begin VB.Form ACE1100C 
   BackColor       =   &H80000000&
   Caption         =   "替代结果查询及修改"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   15180
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   1335
      Left            =   285
      TabIndex        =   9
      Top             =   120
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   2355
      _Version        =   196609
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
         Left            =   1320
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox CBO_PROD_CD 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6480
         TabIndex        =   2
         Text            =   "未选择"
         Top             =   120
         Width           =   1095
      End
      Begin VB.ComboBox CBO_dome_FL 
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
         Left            =   3960
         TabIndex        =   1
         Text            =   "未选择"
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txt_prod_len_to 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "INS_EMP"
         Top             =   720
         Width           =   705
      End
      Begin VB.TextBox txt_prod_len_from 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   3
         Tag             =   "INS_EMP"
         Top             =   720
         Width           =   705
      End
      Begin VB.TextBox txt_prod_wid_to 
         BackColor       =   &H00FFFFFF&
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
         Left            =   5760
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "INS_EMP"
         Top             =   720
         Width           =   705
      End
      Begin VB.TextBox txt_prod_wid_from 
         BackColor       =   &H00FFFFFF&
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
         Left            =   4560
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "INS_EMP"
         Top             =   720
         Width           =   705
      End
      Begin VB.TextBox Txt_proc_thk_to 
         BackColor       =   &H00FFFFFF&
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
         Left            =   9000
         MaxLength       =   7
         TabIndex        =   8
         Tag             =   "INS_EMP"
         Top             =   720
         Width           =   705
      End
      Begin VB.TextBox Txt_prod_thk_from 
         BackColor       =   &H00FFFFFF&
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
         Left            =   7800
         MaxLength       =   7
         TabIndex        =   7
         Tag             =   "INS_EMP"
         Top             =   720
         Width           =   705
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   5280
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "产品分类"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   120
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   3360
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
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
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   6600
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   120
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "工厂"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
         Left            =   2760
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "接受订单"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
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
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   2400
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line2 
         X1              =   5400
         X2              =   5640
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         X1              =   8640
         X2              =   8880
         Y1              =   840
         Y2              =   840
      End
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   7170
      Left            =   270
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Width           =   8205
      _Version        =   393216
      _ExtentX        =   14473
      _ExtentY        =   12647
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   15
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACE1100C.frx":0000
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   7170
      Left            =   8685
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1680
      Width           =   6330
      _Version        =   393216
      _ExtentX        =   11165
      _ExtentY        =   12647
      _StockProps     =   64
      BackColorStyle  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      MaxRows         =   9
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACE1100C.frx":1B26
   End
End
Attribute VB_Name = "ACE1100C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
