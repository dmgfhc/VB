VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AFH2080C 
   Caption         =   "板坯低倍取样实绩录入与修改_AFH2080C"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txt_PRC_LINE 
      Height          =   300
      ItemData        =   "AFH2080C.frx":0000
      Left            =   1560
      List            =   "AFH2080C.frx":000D
      TabIndex        =   11
      Top             =   120
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
      Left            =   8880
      MaxLength       =   18
      TabIndex        =   10
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
      Left            =   6100
      MaxLength       =   18
      TabIndex        =   9
      Top             =   600
      Width           =   495
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
      Left            =   8880
      MaxLength       =   18
      TabIndex        =   8
      Top             =   120
      Width           =   1575
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
      Left            =   1320
      MaxLength       =   18
      TabIndex        =   7
      Top             =   600
      Width           =   1215
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
      Left            =   2880
      MaxLength       =   18
      TabIndex        =   6
      Top             =   600
      Width           =   375
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
      Left            =   4680
      MaxLength       =   18
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   15255
      Begin FPSpread.vaSpread ss1 
         Height          =   6735
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   15255
         _Version        =   393216
         _ExtentX        =   26908
         _ExtentY        =   11880
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
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
         MaxCols         =   44
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFH2080C.frx":001A
      End
   End
   Begin VB.CheckBox txt_CHECK 
      Height          =   180
      Left            =   10800
      TabIndex        =   2
      Top             =   675
      Width           =   180
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   0
      Left            =   240
      Top             =   120
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   1
      Left            =   3000
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
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
      ForeColor       =   16711680
   End
   Begin InDate.UDate from_date 
      Height          =   315
      Left            =   4320
      TabIndex        =   0
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
      Left            =   6100
      TabIndex        =   1
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Index           =   2
      Left            =   3600
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
      Index           =   3
      Left            =   11040
      Top             =   600
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "已取样"
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
      Left            =   240
      Top             =   600
      Width           =   1065
      _ExtentX        =   1879
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   5830
      Top             =   120
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
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Index           =   4
      Left            =   7800
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
      Index           =   5
      Left            =   5830
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Index           =   6
      Left            =   7800
      Top             =   600
      Width           =   1065
      _ExtentX        =   1879
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Index           =   1
      Left            =   2520
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "有订单特殊要求"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10800
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "AFH2080C"
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


Private Sub Form_Load()
  sAuthority = Gf_Pgm_Authority(Me.Name)
  
  Call Form_Define
  
  Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
  
  Call Gp_Ms_Cls(Mc1("rControl"))
  Call Gp_Ms_NeceColor(Mc1("nControl"))
  
  Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
  Call Gf_Sp_Cls(Proc_Sc("Sc"))
  Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
  txt_prc_line.ItemData(0) = "1"
End Sub

Private Sub Form_Define()
  FormType = "Msheet"
  
  'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
  Call Gp_Ms_Collection(txt_prc_line, "p", "n", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(from_date, "p", "n", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(to_date, "p", "n", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_stlgrd, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_ORD_NO, "p", "", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_ORD_ITEM, "p", "", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_CAST_NO, "p", "", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_CAST_CNT, "p", "", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_CHECK, "p", "", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_heat_no, "p", "", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
  Mc1.Add Item:=pControl, Key:="pControl"
  Mc1.Add Item:=nControl, Key:="nControl"
  Mc1.Add Item:=mControl, Key:="mControl"
  Mc1.Add Item:=iControl, Key:="iControl"
  Mc1.Add Item:=rControl, Key:="rControl"
  Mc1.Add Item:=cControl, Key:="cControl"
  Mc1.Add Item:=aControl, Key:="aControl"
  Mc1.Add Item:=lControl, Key:="lControl"
  
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
  Call Gp_Sp_Collection(ss1, 1, "", "", "", "i", "", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 2, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 3, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 4, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 5, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 6, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 7, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 8, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 9, "", "", "", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 10, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 11, "", "", "", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 12, "", "", "", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 13, "", "", "", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 14, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 15, "", "", "", " ", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 16, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 17, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 18, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 19, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 20, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 21, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 22, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 23, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 24, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 25, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 26, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 27, "", "", "", " ", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
  Call Gp_Sp_Collection(ss1, 39, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 40, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 41, "", "", "", "i", "", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 42, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 43, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 44, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  sc1.Add Item:=ss1, Key:="Spread"
  sc1.Add Item:="AFH2080C.P_MODIFY", Key:="P-M"  '-----
  sc1.Add Item:="AFH2080C.P_REFER", Key:="P-R"
  'Sc1.Add Item:="WZL0010C.P_SONEROW", Key:="P-O"
  
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
  ss1.ColsFrozen = 8

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
    MDIMain.MenuTool.Buttons(14).Enabled = False
    MDIMain.MenuTool.Buttons(15).Enabled = False

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
  Dim j As Integer
  
  If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    ss1.OperationMode = OperationModeNormal
    
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(10).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(13).Enabled = False
    MDIMain.MenuTool.Buttons(14).Enabled = False
    MDIMain.MenuTool.Buttons(15).Enabled = False
    If txt_CHECK.VALUE = 0 Then
    For j = 1 To ss1.MaxRows
        With ss1
         .Col = 13
         .Row = j
         If .Text <> "N" Then
           Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, j, j, , &HFF80FF)
         End If
         End With
        Next j
    
    End If
  End If
Exit Sub
Refer_Err:
End Sub


Public Sub Form_Pro()  '保存
    Dim j As Integer
  If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    If txt_CHECK.VALUE = 0 Then
    For j = 1 To ss1.MaxRows
        With ss1
         .Col = 13
         .Row = j
         If .Text <> "N" Then
           Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, j, j, , &HFF80FF)
         End If
         End With
        Next j
    
    End If
  End If
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(10).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(13).Enabled = False
    MDIMain.MenuTool.Buttons(14).Enabled = False
    MDIMain.MenuTool.Buttons(15).Enabled = False
End Sub
Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If
    

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
  With ss1
     .Col = 41:
            .Row = ss1.ActiveRow
            .Text = sUserID
    End With
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)  '修改
  If Gf_Sc_Authority(sAuthority, "U") Then
    Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
  End If
  
End Sub

Public Sub Form_Exit()  '关闭界面
    Unload Me
End Sub

