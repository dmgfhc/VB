VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0094C 
   Caption         =   "低倍检验结果查询_AQC0094C"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   16325
      _Version        =   196609
      PaneTree        =   "AQC0094C.frx":0000
      Begin Threed.SSPanel SSPanel1 
         Height          =   1050
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15075
         _ExtentX        =   26591
         _ExtentY        =   1852
         _Version        =   196609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.ComboBox txt_PRC_LINE 
            Height          =   300
            ItemData        =   "AQC0094C.frx":0052
            Left            =   6600
            List            =   "AQC0094C.frx":005F
            TabIndex        =   15
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txt_ORD_NO 
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
            TabIndex        =   7
            Top             =   600
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
            Left            =   6600
            MaxLength       =   18
            TabIndex        =   6
            Top             =   600
            Width           =   1335
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
            Left            =   9480
            MaxLength       =   18
            TabIndex        =   5
            Top             =   120
            Width           =   1335
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
            Left            =   11160
            MaxLength       =   18
            TabIndex        =   4
            Top             =   120
            Width           =   615
         End
         Begin VB.TextBox txt_ORD_ITEM 
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
            Left            =   11160
            MaxLength       =   18
            TabIndex        =   3
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txt_HEAT_NO 
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
            Left            =   13320
            MaxLength       =   18
            TabIndex        =   2
            Top             =   120
            Width           =   1335
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   2
            Left            =   120
            Top             =   600
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "检验日期"
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
         Begin InDate.UDate from_date 
            Height          =   315
            Left            =   1440
            TabIndex        =   8
            Tag             =   "发放日期"
            Top             =   600
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
            Left            =   3360
            TabIndex        =   9
            Tag             =   "发放日期"
            Top             =   600
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
            Left            =   3000
            Top             =   600
            Width           =   345
            _ExtentX        =   609
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
            Top             =   600
            Width           =   1185
            _ExtentX        =   2090
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Index           =   4
            Left            =   5400
            Top             =   600
            Width           =   1185
            _ExtentX        =   2090
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
            Left            =   8280
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   1
            Left            =   12120
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "炉号/板坯号"
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
            Index           =   0
            Left            =   10800
            Top             =   600
            Width           =   345
            _ExtentX        =   609
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
            Index           =   7
            Left            =   120
            Top             =   90
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "铸坯生产日期"
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
         Begin InDate.UDate dtp_from_date 
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Tag             =   "发放日期"
            Top             =   90
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
         Begin InDate.UDate dtp_to_date 
            Height          =   315
            Left            =   3360
            TabIndex        =   14
            Tag             =   "发放日期"
            Top             =   90
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   3000
            Top             =   90
            Width           =   345
            _ExtentX        =   609
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Index           =   0
            Left            =   10800
            Top             =   120
            Width           =   345
            _ExtentX        =   609
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
            Index           =   4
            Left            =   5400
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
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
         Begin VB.Label Label1 
            Caption         =   "炉号/板坯号输入之后，查询与其他条件无关"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   12120
            TabIndex        =   16
            Top             =   480
            Width           =   2535
         End
      End
      Begin FPSpread.vaSpread SS1 
         Height          =   8055
         Left            =   30
         TabIndex        =   10
         Top             =   1170
         Width           =   15075
         _Version        =   393216
         _ExtentX        =   26591
         _ExtentY        =   14208
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
         MaxCols         =   38
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0094C.frx":006C
      End
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   6
      Left            =   240
      Top             =   720
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "检验日期"
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
   Begin InDate.UDate UDate1 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Tag             =   "发放日期"
      Top             =   720
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
   Begin InDate.UDate UDate2 
      Height          =   315
      Left            =   3480
      TabIndex        =   12
      Tag             =   "发放日期"
      Top             =   720
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   3120
      Top             =   720
      Width           =   345
      _ExtentX        =   609
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
End
Attribute VB_Name = "AQC0094C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sAuthority As String
Public FormType As String
Public Toolbar_St As String

Dim Mc1 As New Collection
Dim Sc1 As New Collection
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
  Call Gp_Ms_Collection(from_date, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(to_date, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_CAST_NO, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_CAST_CNT, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_STLGRD, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_ORD_NO, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_ORD_ITEM, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_HEAT_NO, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_PRC_LINE, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(dtp_from_date, "p", "n", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(dtp_to_date, "p", "n", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
  Mc1.Add Item:=pControl, Key:="pControl"
  Mc1.Add Item:=nControl, Key:="nControl"
  Mc1.Add Item:=mControl, Key:="mControl"
  Mc1.Add Item:=iControl, Key:="iControl"
  Mc1.Add Item:=rControl, Key:="rControl"
  Mc1.Add Item:=cControl, Key:="cControl"
  Mc1.Add Item:=aControl, Key:="aControl"
  Mc1.Add Item:=lControl, Key:="lControl"
  
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
  Call Gp_Sp_Collection(ss1, 1, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 2, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
  Call Gp_Sp_Collection(ss1, 31, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 32, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 33, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 34, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 35, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 36, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 37, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 38, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  
  Sc1.Add Item:=ss1, Key:="Spread"
  Sc1.Add Item:="AQC0094C.P_REFER", Key:="P-R"
  
  Sc1.Add Item:=pColumn1, Key:="pColumn"
  Sc1.Add Item:=nColumn1, Key:="nColumn"
  Sc1.Add Item:=aColumn1, Key:="aColumn"
  Sc1.Add Item:=mColumn1, Key:="mColumn"
  Sc1.Add Item:=iColumn1, Key:="iColumn"
  Sc1.Add Item:=lColumn1, Key:="lColumn"
   Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
  Proc_Sc.Add Item:=Sc1, Key:="Sc"    '-----
  
  Sc1.Item("Spread").Col = 0  '给
  Sc1.Item("Spread").Row = 0
  Sc1.Item("Spread").Text = "◎"
  
  Me.KeyPreview = True
  Me.BackColor = &HE0E0E0
  ss1.ColsFrozen = 5

End Sub

Private Sub Form_Activate()

  Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.MenuTool.Buttons(4).Enabled = False
  MDIMain.MenuTool.Buttons(5).Enabled = False
  MDIMain.MenuTool.Buttons(6).Enabled = False
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
  Set Sc1 = Nothing
  Set Proc_Sc = Nothing
  
  Call MDIMain.FormMenuSetting(Me, "start", Toolbar_St, "")
    
  

End Sub

Public Sub Form_Ref()  '查询
  If Len(Trim(txt_HEAT_NO)) = 0 Or Len(Trim(txt_HEAT_NO)) > 7 Then
    On Error GoTo Refer_Err
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
      Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       ss1.OperationMode = OperationModeNormal
    
      Call GP_SELECT_ROW(ss1, 1)
      MDIMain.MenuTool.Buttons(4).Enabled = False
      MDIMain.MenuTool.Buttons(5).Enabled = False
    MDIMain.MenuTool.Buttons(6).Enabled = False
      MDIMain.MenuTool.Buttons(7).Enabled = False
      MDIMain.MenuTool.Buttons(8).Enabled = False
      MDIMain.MenuTool.Buttons(9).Enabled = False
      MDIMain.MenuTool.Buttons(10).Enabled = False
      MDIMain.MenuTool.Buttons(11).Enabled = False
      MDIMain.MenuTool.Buttons(12).Enabled = False
      Exit Sub
    End If
Refer_Err:
Else
    Call GeneralCommon.Gp_MsgBoxDisplay("炉号/板坯号长度为0位/8位/10位：当前为" & Len(Trim(txt_HEAT_NO)) & "位，请重新输入", "", "")
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

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If
    

End Sub

