VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0093C 
   Caption         =   "板坯低倍检验结果录入_AQC0093C"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6.29672e6
   ScaleMode       =   0  'User
   ScaleWidth      =   3.01255e6
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14420
      _Version        =   196609
      Locked          =   -1  'True
      PaneTree        =   "AQC0093C.frx":0000
      Begin VB.Frame Frame2 
         Height          =   6855
         Left            =   30
         TabIndex        =   10
         Top             =   1290
         Width           =   14955
         Begin FPSpread.vaSpread ss1 
            Height          =   6420
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   14655
            _Version        =   393216
            _ExtentX        =   25850
            _ExtentY        =   11324
            _StockProps     =   64
            AllowDragDrop   =   -1  'True
            AllowMultiBlocks=   -1  'True
            AllowUserFormulas=   -1  'True
            ButtonDrawMode  =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   45
            MaxRows         =   1
            ProcessTab      =   -1  'True
            Protect         =   0   'False
            SpreadDesigner  =   "AQC0093C.frx":0052
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1170
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   14955
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
            Left            =   1440
            MaxLength       =   18
            TabIndex        =   7
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox txt_TEST_TYPE 
            Height          =   300
            ItemData        =   "AQC0093C.frx":13A5
            Left            =   6480
            List            =   "AQC0093C.frx":13B5
            TabIndex        =   6
            Top             =   720
            Width           =   1575
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
            Left            =   6480
            MaxLength       =   18
            TabIndex        =   5
            Top             =   240
            Width           =   1575
         End
         Begin VB.ComboBox txt_PRC_LINE 
            Height          =   300
            ItemData        =   "AQC0093C.frx":13E1
            Left            =   9720
            List            =   "AQC0093C.frx":13EE
            TabIndex        =   4
            Top             =   240
            Width           =   615
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
            Left            =   9720
            MaxLength       =   18
            TabIndex        =   3
            Top             =   720
            Width           =   1455
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
            Left            =   12240
            MaxLength       =   18
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   0
            Left            =   120
            Top             =   720
            Width           =   1305
            _ExtentX        =   2302
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   1
            Left            =   5160
            Top             =   720
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "试验种类"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   2
            Left            =   120
            Top             =   240
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
            Left            =   1440
            TabIndex        =   8
            Tag             =   "发放日期"
            Top             =   240
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
            Top             =   240
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
            Top             =   240
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
            Left            =   5160
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "委托单号"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Index           =   4
            Left            =   8400
            Top             =   240
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
               Size            =   9.75
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
            Left            =   8400
            Top             =   720
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "钢种"
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
         Begin InDate.ULabel ULabel4 
            Height          =   315
            Index           =   2
            Left            =   11160
            Top             =   240
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
               Size            =   9.75
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
End
Attribute VB_Name = "AQC0093C"
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
  Call Gp_Ms_Collection(from_date, "p", "n", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(to_date, "p", "n", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_TRUSTDEED_NO, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_PRC_LINE, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_CAST_NO, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_TEST_TYPE, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_STLGRD, "p", " ", "", "", "r", "", "", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
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
  Call Gp_Sp_Collection(ss1, 2, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 3, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 4, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 5, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 6, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 7, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 8, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 9, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 10, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 11, "", "", "", "i", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 12, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 13, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 14, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 15, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 16, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 17, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 18, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 19, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 20, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 21, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 22, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 23, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 24, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 25, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 26, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 27, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 28, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 29, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 30, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 31, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 32, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 33, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 34, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 35, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 36, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 37, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 38, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 39, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 40, "", "", "", " ", "", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 41, "", "", "", "i", "", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 42, "", "", "", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 43, "", "", "", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 44, "", "", "", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  Call Gp_Sp_Collection(ss1, 45, "", "", "", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
  
  Sc1.Add Item:=ss1, Key:="Spread"
  Sc1.Add Item:="AQC0093C.P_MODIFY", Key:="P-M"
  Sc1.Add Item:="AQC0093C.P_REFER", Key:="P-R"
  
  Sc1.Add Item:=pColumn1, Key:="pColumn"
  Sc1.Add Item:=nColumn1, Key:="nColumn"
  Sc1.Add Item:=aColumn1, Key:="aColumn"
  Sc1.Add Item:=mColumn1, Key:="mColumn"
  Sc1.Add Item:=iColumn1, Key:="iColumn"
  Sc1.Add Item:=lColumn1, Key:="lColumn"
   Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
  Proc_Sc.Add Item:=Sc1, Key:="Sc"    '-----
  
  Sc1.Item("Spread").Col = 0  '
  Sc1.Item("Spread").Row = 0
  Sc1.Item("Spread").Text = "◎"
  
  Me.KeyPreview = True
  Me.BackColor = &HE0E0E0
  ss1.ColsFrozen = 5


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
  Set Sc1 = Nothing
  Set Proc_Sc = Nothing
  
  Call MDIMain.FormMenuSetting(Me, "start", Toolbar_St, "")
    
  

End Sub

Public Sub Form_Ref()  '查询
  On Error GoTo Refer_Err
  If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
     ss1.OperationMode = OperationModeNormal
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
 '   Call GP_SELECT_ROW(ss1, 1)
    
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
    
  If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(10).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(13).Enabled = False
    
  End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub
Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    

End Sub
Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)  '修改
   
  If Gf_Sc_Authority(sAuthority, "U") Then
  Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
  With ss1
     .Col = 41:
            .Row = ss1.ActiveRow
            .Text = sUserID
    End With
    
  End If
  
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

