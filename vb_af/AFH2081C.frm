VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AFH2081C 
   Caption         =   "板坯低倍取样委托界面_AFH2081C"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   FillColor       =   &H00FF0000&
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_TRUSTDEED_DATE 
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
      Height          =   310
      Left            =   1800
      MaxLength       =   18
      TabIndex        =   13
      Top             =   9240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_TRUSTDEED_EMP 
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
      Height          =   310
      Left            =   120
      MaxLength       =   18
      TabIndex        =   12
      Top             =   9240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15901
      _Version        =   196609
      Locked          =   -1  'True
      PaneTree        =   "AFH2081C.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   1080
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   1905
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox txt_PRC_LINE 
            Height          =   300
            ItemData        =   "AFH2081C.frx":0052
            Left            =   1560
            List            =   "AFH2081C.frx":005F
            TabIndex        =   15
            Top             =   600
            Width           =   615
         End
         Begin VB.ComboBox txt_TEST_TYPE 
            Height          =   300
            ItemData        =   "AFH2081C.frx":006C
            Left            =   9480
            List            =   "AFH2081C.frx":007C
            TabIndex        =   14
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_SMP_NO 
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
            Height          =   310
            Left            =   6360
            MaxLength       =   18
            TabIndex        =   8
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox txt_SMP_PROCESS 
            Height          =   180
            Left            =   11520
            TabIndex        =   7
            Top             =   120
            Width           =   180
         End
         Begin VB.TextBox txt_TRUSTDEED_NO 
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
            Height          =   310
            Left            =   9480
            MaxLength       =   18
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
         Begin VB.TextBox txt_STLGRD 
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
            Height          =   310
            Left            =   6360
            MaxLength       =   18
            TabIndex        =   5
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox txt_CAST_NO 
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
            Height          =   310
            Left            =   3480
            MaxLength       =   18
            TabIndex        =   4
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmd_print 
            Caption         =   "打印"
            Height          =   375
            Left            =   11400
            TabIndex        =   3
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txt_CAST_CNT 
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
            Height          =   310
            Left            =   4680
            MaxLength       =   18
            TabIndex        =   2
            Top             =   600
            Width           =   375
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   0
            Left            =   5280
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            Caption         =   "试样号"
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Index           =   3
            Left            =   11760
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "已委托"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   2
            Left            =   240
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "录入日期"
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
         Begin InDate.UDate from_date 
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Tag             =   "发放日期"
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            Text            =   "2016-12-23"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.74
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            BackColor       =   16777215
            RawData         =   "20161223"
         End
         Begin InDate.UDate to_date 
            Height          =   315
            Left            =   3480
            TabIndex        =   10
            Tag             =   "发放日期"
            Top             =   120
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.74
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483630
            BackColor       =   16777215
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   3120
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            Caption         =   "-"
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
            ForeColor       =   -2147483647
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   3
            Left            =   8280
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "委托单号"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   4
            Left            =   240
            Top             =   600
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "铸机号"
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Index           =   4
            Left            =   5280
            Top             =   120
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            Caption         =   "钢种"
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Index           =   2
            Left            =   2400
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            Caption         =   "浇次号"
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Index           =   0
            Left            =   4390
            Top             =   600
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   556
            Caption         =   "-"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   1
            Left            =   8280
            Top             =   600
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "试验种类"
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
      Begin FPSpread.vaSpread ss1 
         Height          =   7785
         Left            =   30
         TabIndex        =   11
         Top             =   1200
         Width           =   15195
         _Version        =   393216
         _ExtentX        =   26802
         _ExtentY        =   13732
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
         MaxCols         =   37
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFH2081C.frx":00A8
      End
   End
End
Attribute VB_Name = "AFH2081C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sAuthority As String
Public FormType As String
Public Toolbar_St As String

Dim Mc1 As New Collection
Dim sc1 As New Collection
Dim Proc_Sc As New Collection

Dim pControl As New Collection
Dim nControl As New Collection
Dim mControl As New Collection
Dim iControl As New Collection
Dim rControl As New Collection
Dim cControl As New Collection
Dim aControl As New Collection
Dim lControl As New Collection

Dim pColumn1 As New Collection
Dim nColumn1 As New Collection
Dim mColumn1 As New Collection
Dim iColumn1 As New Collection
Dim aColumn1 As New Collection
Dim lColumn1 As New Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private xlApp       As Object   'Execel object
Private xlSheet     As Object   'Execel Sheet object





Private Sub Form_Load()
  sAuthority = Gf_Pgm_Authority(Me.Name)
  
  Call Form_Define
  
  Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
  
  Call Gp_Ms_Cls(Mc1("rControl"))
  Call Gp_Ms_NeceColor(Mc1("nControl"))
  
  Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
  Call Gf_Sp_Cls(Proc_Sc("Sc"))
  Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
  
End Sub

Private Sub Form_Define()
  FormType = "Msheet"
  
  'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
  Call Gp_Ms_Collection(from_date, "p", "n", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(to_date, "p", "n", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_TRUSTDEED_NO, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_PRC_LINE, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_CAST_NO, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_CAST_CNT, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_TEST_TYPE, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_STLGRD, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_SMP_PROCESS, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_TRUSTDEED_EMP, "", " ", "", "", "", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_TRUSTDEED_DATE, "", " ", "", "", "", " ", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
  Mc1.Add Item:=pControl, Key:="pControl"
  Mc1.Add Item:=nControl, Key:="nControl"
  Mc1.Add Item:=mControl, Key:="mControl"
  Mc1.Add Item:=iControl, Key:="iControl"
  Mc1.Add Item:=rControl, Key:="rControl"
  Mc1.Add Item:=cControl, Key:="cControl"
  Mc1.Add Item:=aControl, Key:="aControl"
  Mc1.Add Item:=lControl, Key:="lControl"
  
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
  Call Gp_Sp_Collection(ss1, 1, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 2, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 3, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 4, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 5, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 6, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 7, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 8, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 9, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 10, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 11, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 12, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 13, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 14, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 15, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 16, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 17, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 18, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 19, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 20, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 21, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 22, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 23, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 24, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 25, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 26, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 27, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 28, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 29, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 30, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 31, "", "", "", "i", "", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 32, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 33, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 34, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 35, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 36, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 37, "", "", "", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  
  sc1.Add Item:=ss1, Key:="Spread"
  sc1.Add Item:="AFH2081C.P_MODIFY", Key:="P-M"  '-----
  sc1.Add Item:="AFH2081C.P_REFER", Key:="P-R"
  
  sc1.Add Item:=pColumn1, Key:="pColumn"
  sc1.Add Item:=nColumn1, Key:="nColumn"
  sc1.Add Item:=aColumn1, Key:="aColumn"
  sc1.Add Item:=mColumn1, Key:="mColumn"
  sc1.Add Item:=iColumn1, Key:="iColumn"
  sc1.Add Item:=lColumn1, Key:="lColumn"
   sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
  Proc_Sc.Add Item:=sc1, Key:="Sc"    '-----
  
  sc1.Item("Spread").Col = 0  '给
  sc1.Item("Spread").Row = 0
  sc1.Item("Spread").Text = "◎"
  
  Me.KeyPreview = True
  Me.BackColor = &HE0E0E0
  ss1.ColsFrozen = 7

End Sub

Private Sub Form_Activate()

  Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
  MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(10).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(13).Enabled = False

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
  Set pControl = Nothing
  Set nControl = Nothing
  Set iControl = Nothing
  Set rControl = Nothing
  Set cControl = Nothing
  Set aControl = Nothing
  Set lControl = Nothing
  Set mControl = Nothing
  
  Set iControl = Nothing
  Set pControl = Nothing
  Set lControl = Nothing
  Set nControl = Nothing
  Set mControl = Nothing
  Set aControl = Nothing
  
  Set Mc1 = Nothing
  Set sc1 = Nothing
  Set Proc_Sc = Nothing
  
  Call MDIMain.FormMenuSetting(Me, "start", Toolbar_St, "")
    
  

End Sub

Public Sub Form_Ref()  '查询
  On Error GoTo Refer_Err
  If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    ss1.OperationMode = OperationModeNormal
    
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(10).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    Exit Sub
  End If
Refer_Err:
End Sub

Public Sub Form_Pro()  '保存
  Dim arrRecords1 As Variant
    Dim arrRecords2 As Variant
  Dim NowDate       As String
  Dim sQuery        As String
  Dim sQuery1        As String
  Dim StrNum           As Integer
  Dim AdoRs As ADODB.Recordset
  Dim AdoRs1 As ADODB.Recordset
  Dim STR       As String
    
  If txt_SMP_PROCESS = "0" Then
  
      Set AdoRs1 = New ADODB.Recordset
      
      sQuery1 = "SELECT GF_SYSDATE FROM DUAL"
       AdoRs1.Open sQuery1, M_CN1, adOpenKeyset
        
        arrRecords2 = AdoRs1.GetRows
        AdoRs1.Close
        
      NowDate = arrRecords2(0, 0)
      'Call Gp_MsgBoxDisplay(NowDate, "", "错误提示")
       
      Set AdoRs = New ADODB.Recordset
    
        sQuery = "SELECT MAX(SUBSTR(A.TRUSTDEED_NO,9,2)) FROM QP_SLAB_LOWPOWER_SMP A WHERE  1 = 1 "
        sQuery = sQuery + "AND A.TRUSTDEED_DATE = '" + NowDate + "'"
        sQuery = sQuery + "AND A.SMP_PROCESS =  'B' "
    
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        arrRecords1 = AdoRs.GetRows
        AdoRs.Close
    
        If IsEmpty(arrRecords1(0, 0)) Or IsNull(arrRecords1(0, 0)) Then
            With ss1
             .Col = 31:
                    .Row = ss1.ActiveRow
                    .Text = NowDate + "01"
            .Col = 34:
                         .Row = ss1.ActiveRow
                         .Text = sUserID
            txt_TRUSTDEED_NO.Text = NowDate + "01"
            End With
        Else
          StrNum = Val(arrRecords1(0, 0))
          StrNum = StrNum + 1
          If StrNum < 10 Then
            STR = "0" & StrNum
          Else
            STR = StrNum
          End If
          With ss1
              .Col = 31:
                  .Row = ss1.ActiveRow
                  .Text = NowDate + STR
              .Col = 34:
                  .Row = ss1.ActiveRow
                  .Text = sUserID
              txt_TRUSTDEED_NO.Text = NowDate + STR
         End With
        End If
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
          Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
          txt_SMP_PROCESS.VALUE = 1
          Form_Ref
        Else
          txt_TRUSTDEED_NO.Text = ""
        End If
  Else
    Call Gp_MsgBoxDisplay("已委托不可修改，只能打印", "I")
  End If
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(10).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(13).Enabled = False
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
   
    If txt_SMP_PROCESS = "1" Then
        With ss1
          .Row = .ActiveRow
                .Col = 31
                txt_TRUSTDEED_NO.Text = .Text
                .Col = 34
                txt_TRUSTDEED_EMP.Text = .Text
                .Col = 32
                txt_TRUSTDEED_DATE.Text = .Text
         End With
         Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
   
   End If
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub
Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If
    

End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)  '修改
  If Gf_Sc_Authority(sAuthority, "U") Then
    Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
  End If
  
End Sub

Public Sub Form_Exc()  '导出数据到Excel
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Public Sub Form_Exit()  '关闭界面
    Unload Me
End Sub

Private Sub cmd_print_Click()
Dim sMsg As String
Dim arrRecords1 As Variant
Dim sTRUSTDEED_NO    As String


    sTRUSTDEED_NO = Trim(txt_TRUSTDEED_NO.Text)
    If IsNull(sTRUSTDEED_NO) Or sTRUSTDEED_NO = "" Then
      sMsg = "没有填写委托单号"
      Call Gp_MsgBoxDisplay(sMsg, "I")
    Else
       Call funslabcardQuery1(sTRUSTDEED_NO)
    End If


End Sub

Private Function funslabcardQuery1(ByVal sTRUSTDEED_NO As String) As String

Dim arrRecords1 As Variant
Dim sQuery        As String
Dim AdoRs As ADODB.Recordset
           
  Set AdoRs = New ADODB.Recordset
    sQuery = "SELECT   smp_no,"
sQuery = sQuery & "          ord_no,"
sQuery = sQuery & "          ord_item,"
sQuery = sQuery & "          test_knd,"
sQuery = sQuery & "          Gf_comnnamefind('Q0100',test_type),"
sQuery = sQuery & "          Gf_comnnamefind('Q0101',test_standard),"
sQuery = sQuery & "          (SELECT sys_name"
sQuery = sQuery & "           FROM   qp_steel_process_cd"
sQuery = sQuery & "           WHERE  process_id = 'BF14'"
sQuery = sQuery & "                  AND sys_id = a.smp_req),"
sQuery = sQuery & "          smp_width,"
sQuery = sQuery & "          smp_date,"
sQuery = sQuery & "          smp_time,"
sQuery = sQuery & "          smp_emp,"
sQuery = sQuery & "          smp_length,"
sQuery = sQuery & "          center_segregation,"
sQuery = sQuery & "          center_porosity,"
sQuery = sQuery & "          middle_cracks,"
sQuery = sQuery & "          corner_cracks,"
sQuery = sQuery & "          traingular_area_cracks,"
sQuery = sQuery & "          al2o3_inclusion,"
sQuery = sQuery & "          pinhole_bubble,"
sQuery = sQuery & "          honeycomb_bubble,"
sQuery = sQuery & "          center_cracks,"
sQuery = sQuery & "          silicate_inclusion,"
sQuery = sQuery & "          foreign_meal,"
sQuery = sQuery & "          slag_inclusion,"
sQuery = sQuery & "          shrinkage_cavity,"
sQuery = sQuery & "          smp_slab_no,"
sQuery = sQuery & "          smp_process,"
sQuery = sQuery & "          trustdeed_no,"
sQuery = sQuery & "          trustdeed_date,"
sQuery = sQuery & "          trustdeed_time,"
sQuery = sQuery & "          trustdeed_emp,"
sQuery = sQuery & "          Gf_stlgrd_detail(stlgrd),"
sQuery = sQuery & "          cast_no,"
sQuery = sQuery & "          cast_cnt"
sQuery = sQuery & " FROM     qp_slab_lowpower_smp a"
sQuery = sQuery & " WHERE    1 = 1 AND trustdeed_no = '" & sTRUSTDEED_NO & "'"
sQuery = sQuery & " ORDER BY a.smp_no,"
sQuery = sQuery & "          a.test_type,"
sQuery = sQuery & "          a.stlgrd"
 
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
        Set AdoRs = Nothing
        Call Gp_MsgBoxDisplay("未查询到该委托单号", "I")
        funslabcardQuery1 = "Err Database"
        Exit Function
    End If
    arrRecords1 = AdoRs.GetRows
    AdoRs.Close
    
   Set AdoRs = Nothing
    
    If MillSheetPrint_D1(arrRecords1) = "" Then
        funslabcardQuery1 = ""
    Else
        funslabcardQuery1 = "Err Database"
    End If

End Function


Private Function MillSheetPrint_D1(ByVal arrRecords1 As Variant) As String
    Dim PrintStr          As Boolean
    Dim MaxRow         As Long
    Dim MaxCol         As Long
    Dim PrtCnt          As Long
    Dim LneCnt          As Long
    Dim top             As String
    Dim pAry11()        As String                   'CHEM
    Dim pAry12()        As String                   'CHEM
    Dim pAry13()        As String                   'CHEM
    Dim pAry14()        As String                   'CHEM
    
    Dim pAry15()        As String                   'CHEM
   
    Dim i               As Integer
    Dim j               As Integer


    If IsEmpty(arrRecords1) Then
       MillSheetPrint_D1 = "Err Data"
       Exit Function
    End If
    
    MaxRow = UBound(arrRecords1, 1)
    MaxCol = UBound(arrRecords1, 2)
    PrtCnt = -1
    LneCnt = 0

    ReDim pAry11(1 To MaxCol + 1, 1 To 1) '1
    ReDim pAry12(1 To MaxCol + 1, 1 To 3) '5,6,7
    ReDim pAry13(1 To MaxCol + 1, 1 To 13) '13-25
    ReDim pAry14(1 To MaxCol + 1, 1 To 5) '28-32
    
   Set xlApp = GetObject("", "Excel.Application")
        xlApp.Workbooks.Open (App.Path & "\AFH2081C.xls")
        Set xlSheet = xlApp.Worksheets("Sheet1")
                      
      top = "板坯低倍检验委托单:" & txt_TRUSTDEED_NO.Text
      With xlApp.ActiveSheet.PageSetup
             .LeftHeader = ""
             .CenterHeader = "&""宋体,斜体""&20" & top
             .RightHeader = ""
             .LeftFooter = "&""宋体,斜体""&16委托人：" & txt_TRUSTDEED_EMP.Text
             .CenterFooter = ""
             .RightFooter = "&""宋体,斜体""&16委托日期：" & txt_TRUSTDEED_DATE.Text
      End With
    For j = 0 To MaxCol
        For i = 0 To MaxRow
           If i = 1 Then
                pAry11(j + 1, i) = arrRecords1(i - 1, j) & ""
           End If
            If i > 4 And i < 8 Then
                pAry12(j + 1, i - 4) = arrRecords1(i - 1, j) & ""
                
            End If
            If i > 12 And i < 26 Then
                pAry13(j + 1, i - 12) = arrRecords1(i - 1, j) & ""
            End If
            If i > 27 And i < 33 Then
                pAry14(j + 1, i - 27) = arrRecords1(i - 1, j) & ""
            End If
           
        Next i
           
   Next j
         For j = 0 To MaxCol
           For i = 0 To 23
           If i = 1 Then
             xlSheet.Range("B" & (6 + j) & ":B" & (6 + j)).VALUE = pAry11(j + 1, i)
           End If
           If i = 2 Then
             xlSheet.Range("D" & (6 + j) & ":D" & (6 + j)).VALUE = pAry12(j + 1, i - 1)
           End If
           If i = 3 Then
             xlSheet.Range("E" & (6 + j) & ":E" & (6 + j)).VALUE = pAry12(j + 1, i - 1)
           End If
           If i = 4 Then
             xlSheet.Range("F" & (6 + j) & ":F" & (6 + j)).VALUE = pAry12(j + 1, i - 1)
           End If
           If i = 5 Then
             xlSheet.Range("G" & (6 + j) & ":G" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 6 Then
             xlSheet.Range("H" & (6 + j) & ":H" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 7 Then
             xlSheet.Range("I" & (6 + j) & ":I" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 8 Then
             xlSheet.Range("J" & (6 + j) & ":J" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 9 Then
             xlSheet.Range("K" & (6 + j) & ":K" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 10 Then
             xlSheet.Range("L" & (6 + j) & ":L" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 11 Then
             xlSheet.Range("M" & (6 + j) & ":M" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 12 Then
             xlSheet.Range("N" & (6 + j) & ":N" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 13 Then
             xlSheet.Range("O" & (6 + j) & ":O" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 14 Then
             xlSheet.Range("P" & (6 + j) & ":P" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 15 Then
             xlSheet.Range("Q" & (6 + j) & ":Q" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 16 Then
             xlSheet.Range("R" & (6 + j) & ":R" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
           If i = 17 Then
             xlSheet.Range("S" & (6 + j) & ":S" & (6 + j)).VALUE = pAry13(j + 1, i - 4)
           End If
'           If i = 18 Then
'             xlSheet.Range("C" & (6 + j) & ":C" & (6 + j)).Value = pAry14(j + 1, i - 17)
'           End If
'           If i = 19 Then
'             xlSheet.Range("D" & (6 + j) & ":D" & (6 + j)).Value = pAry14(j + 1, i - 17)
'           End If
'           If i = 20 Then
'             xlSheet.Range("E" & (6 + j) & ":E" & (6 + j)).Value = pAry14(j + 1, i - 17)
'           End If
'           If i = 21 Then
'             xlSheet.Range("F" & (6 + j) & ":F" & (6 + j)).Value = pAry14(j + 1, i - 17)
'           End If
           If i = 22 Then
             xlSheet.Range("C" & (6 + j) & ":C" & (6 + j)).VALUE = pAry14(j + 1, i - 17)
           End If

           Next i
         Next j
             xlSheet.PageSetup.PrintArea = "$B$1:$S$" & (MaxCol + 6)
             xlSheet.PageSetup.Orientation = 2
             xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
             Set xlSheet = Nothing
             xlApp.ActiveWorkbook.Close False

    
             xlApp.Quit
    Set xlApp = Nothing
    Call Gp_MsgBoxDisplay("打印完毕", "I")
    Exit Function

End Function
