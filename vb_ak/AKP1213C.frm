VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKP1213C 
   Caption         =   "中厚板卷厂生产简报_V1010_AKP1213C"
   ClientHeight    =   9435
   ClientLeft      =   525
   ClientTop       =   1815
   ClientWidth     =   17235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9435
   ScaleWidth      =   17235
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text_UPD 
      Height          =   270
      Left            =   4050
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   930
   End
   Begin Threed.SSCommand SSCommand_CREAT 
      Height          =   375
      Left            =   13830
      TabIndex        =   1
      Top             =   60
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "生成报表"
      BevelWidth      =   3
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8910
      Left            =   90
      TabIndex        =   2
      Top             =   465
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   15716
      _Version        =   196609
      SplitterBarWidth=   3
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AKP1213C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   2595
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   14970
         _Version        =   393216
         _ExtentX        =   26405
         _ExtentY        =   4577
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   28
         MaxRows         =   10
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP1213C.frx":0092
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   3435
         Left            =   9360
         TabIndex        =   4
         Top             =   2640
         Width           =   5610
         _Version        =   393216
         _ExtentX        =   9895
         _ExtentY        =   6059
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   12
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP1213C.frx":1609
         UnitType        =   0
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   2790
         Left            =   0
         TabIndex        =   5
         Top             =   6120
         Width           =   14970
         _Version        =   393216
         _ExtentX        =   26405
         _ExtentY        =   4921
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   41
         MaxRows         =   11
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP1213C.frx":3D86
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3435
         Left            =   0
         TabIndex        =   7
         Top             =   2640
         Width           =   9315
         _Version        =   393216
         _ExtentX        =   16431
         _ExtentY        =   6059
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   32
         MaxRows         =   10
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP1213C.frx":6F34
      End
   End
   Begin InDate.UDate txt_DATE 
      Height          =   315
      Left            =   1215
      TabIndex        =   6
      Tag             =   "起始日期"
      Top             =   105
      Width           =   1485
      _ExtentX        =   2619
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   90
      Top             =   105
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "日期"
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
Attribute VB_Name = "AKP1213C"
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
'-- Program Name      PROD REPORT V1010
'-- Program ID        AKP1213C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2010.10.8
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Sc4 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection
Dim iSumCol As New Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Sheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_DATE, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AKP1213C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AKP1213C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
 
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AKP1213C.P_MODIFY", Key:="P-M"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss1.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    
    ss3.BlockMode = True
        ss3.Col = 1: ss3.Col2 = ss3.MaxCols
        ss3.Row = 1: ss3.Row2 = ss3.MaxRows
        ss3.Lock = True
        
        Call Sp_ColLock(ss3, 2, 8, False)
        Call Sp_ColLock(ss3, 4, 8, False)
        Call Sp_ColLock(ss3, 5, 6, False)
        Call Sp_ColLock(ss3, 6, 7, False)
        Call Sp_ColLock(ss3, 6, 8, False)
        Call Sp_ColLock(ss3, 6, 9, False)
        Call Sp_ColLock(ss3, 6, 10, False)
        Call Sp_ColLock(ss3, 6, 11, False)
        Call Sp_ColLock(ss3, 6, 12, False)
    
    ss3.BlockMode = False


    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 5, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 6, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 7, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 8, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 9, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 10, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 11, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 12, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 13, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 14, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 15, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 16, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 17, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 18, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 19, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 20, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
       
    'Spread_Collection
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:=pColumn4, Key:="pColumn"
    Sc4.Add Item:=nColumn4, Key:="nColumn"
    Sc4.Add Item:=aColumn4, Key:="aColumn"
    Sc4.Add Item:=mColumn4, Key:="mColumn"
    Sc4.Add Item:=iColumn4, Key:="iColumn"
    Sc4.Add Item:=lColumn4, Key:="lColumn"
    Sc4.Add Item:=1, Key:="First"
    Sc4.Add Item:=ss4.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=Sc4, Key:="sc4"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Sp_ColLock(ss4, 1, 11, False)
    
    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    Call Form_Define
        
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc4")("Spread"))

    Call Gp_Spl_SizeGet(SSSplitter1, "K-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc4")("Spread"), "K-System.INI", Me.Name)

    If Gf_Sc_Authority(sAuthority, "U") Then
       SSCommand_CREAT.Enabled = True
    End If

    txt_DATE.RawData = Format(Date - 1, "yyyymmdd")
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Dim sMessg As String
    
    If Trim(Text_UPD.Text) = "Update" Then
        sMessg = "表格中还有数据未处理，" + vbCrLf
        sMessg = sMessg + "放弃并继续吗？"
        
        If Not Gf_MessConfirm(sMessg, "Q") Then Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "K-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc4")("Spread"), "K-System.INI", Me.Name)
    
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
   
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
   
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
   
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing
   
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Sc4 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()

    ss1.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, True
    ss2.ClearRange 1, 1, ss2.MaxCols, ss2.MaxRows, True
    ss4.ClearRange 1, 1, ss4.MaxCols, ss4.MaxRows, True
    Call ss3_clear
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
'    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    
End Sub

Public Sub Form_Ref()
    
    On Error Resume Next
    
    If txt_DATE.RawData < "20101010" Then
        MsgBox "2010年10月10日以前的报表请到 '中厚板卷厂生产简报_AKP1211C' 画面查询", vbCritical, "错误提示"
       Exit Sub
    End If

    Call Form_Cls
    
    ss1.ReDraw = False
    ss2.ReDraw = False
    ss3.ReDraw = False
    ss4.ReDraw = False
   
'    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl")) Then
    If Sp_Display(M_CN1, Proc_Sc("Sc1")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc1").Item("P-R"), "R", Mc1("pControl"))) Then
        Call Sp_Display(M_CN1, Proc_Sc("Sc2")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
'        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc1, Mc1("nControl"))
        Call Ss3_Data_Refer
        Call Ss4_Data_Refer
        Call Zero_Cls
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True
    ss4.ReDraw = True
'    ss1.ReDraw = False
'    ss2.ReDraw = False
'    ss3.ReDraw = False
'    ss4.ReDraw = False
'   MDIMain.MenuTool.Buttons(4).Enabled = True                 'Save
End Sub

Public Sub Zero_Cls()
    Dim iRow  As Long
    Dim iCol  As Long
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        For iCol = 1 To ss1.MaxCols
            ss1.Col = iCol
            If Val(ss1.Text & "") = 0 Then
                ss1.Text = ""
            End If
        Next iCol
    Next iRow
    
    For iRow = 1 To ss2.MaxRows
        ss2.Row = iRow
        For iCol = 1 To ss2.MaxCols
            ss2.Col = iCol
            If Val(ss2.Text & "") = 0 Then
                ss2.Text = ""
            End If
        Next iCol
    Next iRow
    
    For iRow = 2 To ss3.MaxRows
        ss3.Row = iRow
        For iCol = 2 To ss3.MaxCols
            ss3.Col = iCol
            If ss3.CellType = SS_CELL_TYPE_NUMBER And Val(ss3.Text & "") = 0 Then
               ss3.Text = ""
            End If
        Next iCol
    Next iRow
    
    For iRow = 1 To ss4.MaxRows - 1
        ss4.Row = iRow
        For iCol = 1 To ss4.MaxCols
            ss4.Col = iCol
            If Val(ss4.Text & "") = 0 Then
                ss4.Text = ""
            End If
        Next iCol
    Next iRow
End Sub

Public Sub Form_Pro()
    Dim sDate       As String
    Dim sComtDate   As String
    
    ss3.Col = 5: ss3.Row = 6: sComtDate = Left(ss3.Text, 8)
    sDate = Format(Left(ss3.Text, 8), "####-##-##")
    
    If Not IsDate(sDate) Or Left(txt_DATE.RawData, 6) <> Left(ss3.Text, 6) Then
       MsgBox "日期必须输入", vbCritical, "错误提示"
       Exit Sub
    End If
    
    If Sp_Process(M_CN1, Proc_Sc("Sc3")) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
End Sub

Public Sub Form_Exc()

'    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call ExcelPrn
    
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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc1")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
End Sub

Public Sub Sp_Setting(ByVal sPname As Variant)

    Dim iRow As Integer

    With sPname
        .RowHeight(-1) = 13
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 13
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 13
        Else
            .RowHeight(0) = 24
        End If
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
        
        
        .OperationMode = OperationModeNormal
        .RetainSelBlock = True
        .UserResize = UserResizeColumns
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .Row = 0
        .FontBold = True
    End With
    
End Sub

Public Sub Ss3_Data_Refer()

On Error GoTo Ss3_Display_Error

    Dim sTdate      As String
    Dim sBfdate     As String
    Dim sQuery      As String
    Dim IDc         As Integer

    Dim dNewDate    As Date
    Dim dEndDate    As Date
    Dim lDiff       As Long
    Dim dPlanWgt    As Double
    Dim dActWgt     As Double

    Dim AdoRs As ADODB.Recordset

    Set AdoRs = New ADODB.Recordset
  
    sQuery = "SELECT            PROD_DATE,          MON_S_PLN_PROD,     " & vbCrLf
    sQuery = sQuery & "         MON_S_ACT_PROD,     MON_PROG,           " & vbCrLf
    sQuery = sQuery & "         MON_SLAB_WGT_PROG,  YARD_SLAB_WGT,      " & vbCrLf
    sQuery = sQuery & "         SCRAP_REM_WGT,      YEAR_S_PLN_PROD,    " & vbCrLf
    sQuery = sQuery & "         YEAR_S_ACT_PROD,    YEAR_PROG,          " & vbCrLf
    sQuery = sQuery & "         YEAR_SLAB_WGT_PROG, AVE_DAY_SLAB_WGT,   " & vbCrLf
    sQuery = sQuery & "         DAY_ADD_SLAB_WGT,   CON_LIFE_1,         " & vbCrLf
    sQuery = sQuery & "         TAP_TOT_WGT_1,      TAP_AVE_WGT_1,      " & vbCrLf
    sQuery = sQuery & "         CON_LIFE_2,         TAP_TOT_WGT_2,      " & vbCrLf
    sQuery = sQuery & "         TAP_AVE_WGT_2,      AVE_YEAR_SLAB_WGT,  " & vbCrLf
    sQuery = sQuery & "         YEAR_ADD_SLAB_WGT,  DAY_SCRAP_WGT,      " & vbCrLf
    sQuery = sQuery & "         SCRAP_WGT,          DAY_SCRAP_WGT1,     " & vbCrLf
    sQuery = sQuery & "         DAY_SCRAP_WGT2,     DAY_SCRAP_WGT3,     " & vbCrLf
    sQuery = sQuery & "         DAY_SCRAP_WGT4,     DAY_SCRAP_WGT5,     " & vbCrLf
    sQuery = sQuery & "         DAY_SCRAP_WGT6,     DAY_SCRAP_WGT7,     " & vbCrLf
    sQuery = sQuery & "         DAY_SCRAP_WGT8,     CON_LIFE_3 ,        " & vbCrLf
    sQuery = sQuery & "         TAP_TOT_WGT_3,      TAP_AVE_WGT_3       " & vbCrLf
    sQuery = sQuery & "   FROM  FP_PROD_DAY_REPORT_2                    " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE  =  '" & txt_DATE.RawData & "'" & vbCrLf
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Do Until AdoRs.EOF
       
        With ss3
    
            dNewDate = CDate(txt_DATE.Text)
            dEndDate = DateAdd("m", 1, dNewDate)
            lDiff = DateDiff("d", dNewDate, dEndDate) - Val(Mid(txt_DATE.RawData, 7, 2))
            .Col = 2:   .Row = 2:    .Text = Val(AdoRs.Fields(1) & "")   '月计划钢产量
                                     dPlanWgt = Val(AdoRs.Fields(1) & "")
                        .Row = 3:    .Text = Val(AdoRs.Fields(2) & "")   '已完成钢产量
                                     dActWgt = Val(AdoRs.Fields(2) & "")
                        '.ROW = 3:    .Text = Val(AdoRs.Fields(3) & "")
                        .Row = 4:    .Text = lDiff                       '月剩余天数
                        .Row = 5:    If lDiff <> 0 Then .Text = Round((dPlanWgt - dActWgt) / lDiff, 3)  '日需均产
                        .Row = 6:    .Text = Val(AdoRs.Fields(3) & "")   '月日历进度
                        .Row = 7:    .Text = Val(AdoRs.Fields(4) & "")   '月产量进度
                        .Row = 8:    .Text = Val(AdoRs.Fields(5) & "")   '坯料库存
                        .Row = 10:   .Text = Val(AdoRs.Fields(13) & "")  '1#炉龄
                        .Row = 11:   .Text = Val(AdoRs.Fields(14) & "")  '1#炉累计产量
                        .Row = 12:   .Text = Val(AdoRs.Fields(15) & "")  '1#炉平均炉产
                                  
            .Col = 3:
                        .Row = 10:   .Text = Val(AdoRs.Fields(16) & "")      '2#炉龄
                        .Row = 11:   .Text = Val(AdoRs.Fields(17) & "")      '2#炉累计产量
                        .Row = 12:   .Text = Val(AdoRs.Fields(18) & "")      '2#炉平均炉产
                        
            dEndDate = CDate(Left(txt_DATE.Text, 5) & "12-31")
            lDiff = DateDiff("d", dNewDate, dEndDate)
            .Col = 4:   .Row = 2:    .Text = Val(AdoRs.Fields(7) & "")       '年计划钢产量
                                       dPlanWgt = Val(AdoRs.Fields(7) & "")
                        .Row = 3:    .Text = Val(AdoRs.Fields(8) & "")       '已完成钢产量
                                       dActWgt = Val(AdoRs.Fields(8) & "")
                        .Row = 4:     .Text = lDiff                           '年剩余天数
                        .Row = 5:     If lDiff <> 0 Then .Text = Round((dPlanWgt - dActWgt) / lDiff, 3)  '日需均产
                        .Row = 6:     .Text = Val(AdoRs.Fields(9) & "")     '年日历进度
                        .Row = 7:     .Text = Val(AdoRs.Fields(10) & "")     '年产量进度
                        .Row = 8:     .Text = Val(AdoRs.Fields(6) & "")      '废钢库存
                        
                        .Row = 10:   .Text = Val(AdoRs.Fields(31) & "")      '3#炉龄
                        .Row = 11:   .Text = Val(AdoRs.Fields(32) & "")      '3#炉累计产量
                        .Row = 12:   .Text = Val(AdoRs.Fields(33) & "")      '3#炉平均炉产
                        
            .Col = 6:   .Row = 2:     .Text = Val(AdoRs.Fields(11) & "")     '月平均日产
                        .Row = 3:     .Text = Val(AdoRs.Fields(12) & "")     '月预计产量
                        .Row = 4:     .Text = Val(AdoRs.Fields(19) & "")     '年平均日产
                        .Row = 5:     .Text = Val(AdoRs.Fields(20) & "")     '年预计产量
            
            .Col = 8:   .Row = 3:     .Text = Val(AdoRs.Fields(23) & "")     '当日自产渣铁
                        .Row = 4:     .Text = Val(AdoRs.Fields(24) & "")     '当日自产渣钢
                        .Row = 5:     .Text = Val(AdoRs.Fields(25) & "")     '当日自产落地废钢
                        .Row = 6:     .Text = Val(AdoRs.Fields(26) & "")     '当日自产中包余钢
                        .Row = 7:     .Text = Val(AdoRs.Fields(27) & "")     '当日自产事故槽钢
                        .Row = 8:     .Text = Val(AdoRs.Fields(28) & "")     '当日自产坯头坯尾
                        .Row = 9:     .Text = Val(AdoRs.Fields(29) & "")     '当日自产毛刺
                        .Row = 10:    .Text = Val(AdoRs.Fields(30) & "")     '当日自产板坯库废钢
                        
                        .Row = 11:    .Text = Val(AdoRs.Fields(21) & "")     '当日自产废钢
                        .Row = 12:    .Text = Val(AdoRs.Fields(22) & "")     '累计自产废钢

        End With
    
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
    
    sTdate = txt_DATE.RawData
    sBfdate = ""
    
    sQuery = "SELECT            MAX(PROD_DATE)                             " & vbCrLf
    sQuery = sQuery & "   FROM  FP_PROD_DAY_REPORT_2                       " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE   <=  '" & sTdate & "'           " & vbCrLf
    sQuery = sQuery & "    AND  PROD_DATE   >=  '" & Left(sTdate, 6) & "01'" & vbCrLf
    sQuery = sQuery & "    AND ( BF_WGT          >   0                     " & vbCrLf
    sQuery = sQuery & "      OR  BOF_SCRAP_WGT   >   0                     " & vbCrLf
    sQuery = sQuery & "      OR  HM_CONSP_WGT    >   0                     " & vbCrLf
    sQuery = sQuery & "      OR  SCRAP_RTN_WGT   >   0                     " & vbCrLf
    sQuery = sQuery & "      OR  MAT_CONSP       >   0 )                   " & vbCrLf

    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Do Until AdoRs.EOF
        sBfdate = AdoRs.Fields(0) & ""
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
    
    If Trim(sBfdate) = "" Then Exit Sub
    
    sQuery = "SELECT            BF_WGT,             BOF_SCRAP_WGT,      " & vbCrLf
    sQuery = sQuery & "         PH_CONSP_WGT,       HM_CONSP_WGT,       " & vbCrLf
    sQuery = sQuery & "         SCRAP_RTN_WGT,      MAT_CONSP           " & vbCrLf
    sQuery = sQuery & "   FROM  FP_PROD_DAY_REPORT_2                    " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE  =  '" & sBfdate & "'         " & vbCrLf

    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Do Until AdoRs.EOF
       
        With ss3

'            .Col = 2:   .ROW = 7:    .Text = "至  " & Right(sBfdate, 2) & " 日止"
            .Col = 5:   .Row = 6:    .Text = sBfdate & " 日止实际钢铁料消耗"
            
            .Col = 6:   .Row = 7:    .Text = Val(AdoRs.Fields(0) & "")    '实际进铁量
                        .Row = 8:    .Text = Val(AdoRs.Fields(1) & "")    '实际废钢消耗
                        .Row = 9:    .Text = Val(AdoRs.Fields(2) & "")    '实际生铁消耗
                        .Row = 10:   .Text = Val(AdoRs.Fields(3) & "")    '切边/对接坯
                        .Row = 11:   .Text = Val(AdoRs.Fields(4) & "")    '退废量
                        .Row = 12:   .Text = Val(AdoRs.Fields(5) & "")    '实际钢铁料消耗
        End With
    
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
    
    Exit Sub

Ss3_Display_Error:
    
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Ss3_Display_Error : " & Error)
    
End Sub

Public Sub Ss4_Data_Refer()

On Error GoTo Ss4_Display_Error

    Dim sQuery      As String
    Dim sStlgrd     As String
    Dim sStlgrdName As String
    Dim sgroup      As String
    Dim lWgt        As Double
    Dim IDc         As Integer
    Dim iCol        As Integer
    Dim strTemp     As String
    'Dim TEMP_CNT    As Integer  'ADD BY GUOLI 200706041107
    
    Dim AdoRs       As ADODB.Recordset
    Dim AdoRs1      As ADODB.Recordset

    Set AdoRs = New ADODB.Recordset
    Set AdoRs1 = New ADODB.Recordset
    
    For iCol = 1 To ss4.MaxCols
        ss4.Row = 0
        ss4.Col = iCol
        ss4.Text = " "
        Call Gp_Sp_ColHidden(Sc4.Item("Spread"), iCol, False)
    Next iCol
    
    sQuery = "SELECT            DECODE(PROD_GROUP,'A',2,'B',4,'C',6,8),    PROD_GROUP,          " & vbCrLf
    sQuery = sQuery & "         STLGRD, Gf_Stlgrd_Detail(trim(STLGRD)), SUM(PROD_WGT)           " & vbCrLf
    sQuery = sQuery & "   FROM  FP_PROD_STLGRD_2                                                " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE             <=  '" & txt_DATE.RawData & "'            " & vbCrLf
    sQuery = sQuery & "    AND  SUBSTR(PROD_DATE,1,6)  =  SUBSTR('" & txt_DATE.RawData & "',1,6)" & vbCrLf
    sQuery = sQuery & "  GROUP  BY STLGRD, PROD_GROUP                                           " & vbCrLf
    sQuery = sQuery & "  ORDER  BY 4, PROD_GROUP                                                " & vbCrLf

    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    IDc = 1
    iCol = 1
    Do Until AdoRs.EOF
        
        With ss4
            sgroup = Trim(AdoRs.Fields(1) & "")
            sStlgrd = Trim(AdoRs.Fields(2) & "")
            sStlgrdName = Trim(AdoRs.Fields(3) & "")
            lWgt = Val(AdoRs.Fields(4) & "")
            
            If iCol = 42 Then
               Exit Do
            End If
            
            For iCol = 1 To ss4.MaxCols
                ss4.Row = 0
                ss4.Col = iCol
                If Trim(ss4.Text) = "" Then
                    .Col = IDc
                    IDc = IDc + 1
                    Exit For
                ElseIf Trim(ss4.Text) = Trim(sStlgrdName) Then
                    .Col = iCol
                    Exit For
                End If
            Next iCol
            
            .Row = 0
            If Trim(.Text) = "" Then
               .Text = sStlgrdName
            End If
            
            
            .Row = Trim(AdoRs.Fields(0) & "")
            If Trim(.Text) = "" Then
                .Text = Val(AdoRs.Fields(4) & "")
            End If
            
            sQuery = "SELECT            DECODE(PROD_GROUP,'A',1,'B',3,'C',5,7) , " & vbCrLf
            sQuery = sQuery & "         PROD_WGT                                 " & vbCrLf
            sQuery = sQuery & "   FROM  FP_PROD_STLGRD_2                         " & vbCrLf
            sQuery = sQuery & "  WHERE  PROD_DATE   =  '" & txt_DATE.RawData & "'" & vbCrLf
            sQuery = sQuery & "    AND  STLGRD      =  '" & sStlgrd & "'         " & vbCrLf
            sQuery = sQuery & "    AND  PROD_GROUP  =  '" & sgroup & "'          " & vbCrLf
        
            AdoRs1.Open sQuery, M_CN1, adOpenKeyset
            
            Do Until AdoRs1.EOF
                .Row = Trim(AdoRs1.Fields(0) & "")
                If Trim(.Text) = "" Then
                .Text = Val(AdoRs1.Fields(1) & "")
                End If
                AdoRs1.MoveNext
            Loop
            
            AdoRs1.Close
        End With
        
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
    
'    For iCol = 1 To ss4.MaxCols
'        ss4.ROW = 0
'        ss4.Col = iCol
'        If Trim(ss4.Text) = "" Then
'           Call Gp_Sp_ColHidden(sc4.Item("Spread"), iCol, True)
'        Else
'           TEMP_CNT = TEMP_CNT + 1
'        End If
'    Next iCol
    
    sQuery = "SELECT            DESCRIPTION, D_ML, M_ML                                         " & vbCrLf
    sQuery = sQuery & "   FROM  FP_PROD_DAY_REPORT_2                                            " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE             =  '" & txt_DATE.RawData & "'             " & vbCrLf
                                           
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.EOF Then
          ss4.Col = 1: ss4.Row = 11: ss4.Text = Trim(AdoRs.Fields(0) & "")
          ss4.Col = 54: ss4.Row = 11: ss4.Text = Trim(AdoRs.Fields(1) & "")
          ss4.Col = 57: ss4.Row = 11: ss4.Text = Trim(AdoRs.Fields(2) & "")
    End If
    AdoRs.Close
    
    Screen.MousePointer = vbDefault
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Sub

Ss4_Display_Error:
    
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Ss4_Display_Error : " & Error)
    
End Sub

Public Function ss3_clear() As Boolean
    Dim i As Integer
    
    With ss3
        For i = 2 To 12
            .Row = i
            If .Row <> 9 Then
                .Col = 2
                .Text = ""
                .Col = 4
                .Text = ""
            End If
            
            If .Row <> 6 Then
                .Col = 6
                .Text = ""
            End If
            
            If .Row >= 3 Then
               .Col = 8
               .Text = ""
            End If
            
            If .Row = 6 Then
               .Col = 5
               .Text = "日止实际钢铁料消耗"
            End If
        Next i
    End With
    
End Function

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc2")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    

End Sub

Public Function Sp_Display(Conn As ADODB.Connection, sPname As Variant, sQuery As String) As Boolean

    On Error Resume Next

    Dim iCount          As Integer
    Dim iRowCount       As Long
    Dim iColcount       As Long
    Dim AdoRs           As ADODB.Recordset
    Dim ArrayRecords    As Variant

    Sp_Display = True

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If

    Set AdoRs = New ADODB.Recordset

    With sPname

        .ReDraw = False
        iCount = 0

'        .ClearRange 1, 1, .MaxCols, .MaxRows, True

        Screen.MousePointer = vbHourglass

        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset

        If AdoRs.BOF Or AdoRs.EOF Then

            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Sp_Display = False
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Call Form_Cls
            Screen.MousePointer = vbDefault
            Exit Function

        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) <> 0 Then

            For iRowCount = 0 To .MaxRows - 1
                Select Case Trim(ArrayRecords(0, iRowCount))
                    Case "A0"
                        .Row = 1
                    Case "A1"
                        .Row = 2
                    Case "B0"
                        .Row = 3
                    Case "B1"
                        .Row = 4
                    Case "C0"
                        .Row = 5
                    Case "C1"
                        .Row = 6
                    Case "D0"
                        .Row = 7
                    Case "D1"
                        .Row = 8
                    Case "T0"
                        .Row = 9
                    Case "T1"
                        .Row = 10
                End Select
            
'            .ROW = iRowCount + 1

                For iColcount = 1 To .MaxCols
    
                    .Col = iColcount
    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iColcount, iRowCount))
                    End If

                Next iColcount

            Next iRowCount

        End If

        .ReDraw = True
        Screen.MousePointer = vbDefault

    End With

End Function

Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
       Text_UPD.Text = "Update"
    End If
End Sub

Private Sub ss3_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc3")("Spread").MaxRows < 1 Then Exit Sub
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
End Sub

Public Sub Sp_ColLock(sPname As Variant, ColNum As Variant, RowNum As Variant, LockType As Boolean)

    With sPname
        .Protect = True
        .Col = ColNum: .Col2 = ColNum
        .Row = RowNum: .Row2 = RowNum
        
        .BlockMode = True
        .Lock = LockType
        .BlockMode = False
    End With
    
End Sub

Public Function Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

    On Error GoTo SpreadPro_Error

    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    
    Dim sMesg As String
    Dim sTemp As String
    
    Dim adoCmd As ADODB.Command

    If txt_DATE.RawData = "" Then
       Sp_Process = False
       Call Form_Ref
       Exit Function
    End If
    If Text_UPD.Text = "" Then
        Sp_Process = False
        Call Form_Ref
        Exit Function
    End If
    
    Sp_Process = True
     
    With ss3
    
        'MaxRow = 0 is Exit Function Or iCount = 0
        If .MaxRows < 1 Then
            Sp_Process = False
            Exit Function
        End If
        
        Screen.MousePointer = vbHourglass
        
        .ReDraw = False
        
        'Db Connection Check
        If Conn Is Nothing Then
            If GF_DbConnect = False Then Sp_Process = False: Exit Function
        End If
        
        'Ado Setting
        Conn.CursorLocation = adUseServer
        Set adoCmd = New ADODB.Command
        
        Set adoCmd.ActiveConnection = Conn
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = Sc.Item("P-M")
        
        Conn.BeginTrans
        
        'Ceate Parameter (Input) iType + iColumn
        For iCount = 0 To 12
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        If Text_UPD.Text <> "" Then
          adoCmd.Parameters(0).Value = "U"
        End If
        
        adoCmd.Parameters(1).Value = txt_DATE.RawData
        .Col = 2:    .Row = 8:     adoCmd.Parameters(2).Value = Val(.Text & "")     '坯料库存
        .Col = 4:    .Row = 8:     adoCmd.Parameters(3).Value = Val(.Text & "")     '废钢库存
        .Col = 6:    .Row = 7:     adoCmd.Parameters(4).Value = Val(.Text & "")     '实际进铁量
                     .Row = 8:     adoCmd.Parameters(5).Value = Val(.Text & "")     '实际废钢消耗量
                     .Row = 9:     adoCmd.Parameters(6).Value = Val(.Text & "")     '实际生铁消耗量
                     .Row = 10:    adoCmd.Parameters(7).Value = Val(.Text & "")     '切边/对接坯
                     .Row = 11:    adoCmd.Parameters(8).Value = Val(.Text & "")     '退废量
                     .Row = 12:    adoCmd.Parameters(9).Value = Val(.Text & "")     '实绩钢铁料消耗
        ss4.Col = 1: ss4.Row = 11: adoCmd.Parameters(10).Value = Trim(ss4.Text)      'DESCRIPTION
        .Col = 5:    .Row = 6:     adoCmd.Parameters(11).Value = Trim(Left(.Text, 8))
        adoCmd.Execute
                                       
                    
        'Error Check
        If adoCmd("Error") <> "0" Then
        
            ret_Result_ErrCode = adoCmd("Error")
            ret_Result_ErrMsg = adoCmd("Messg")
            sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
            
            Call Gp_MsgBoxDisplay(sErrMessg)
            Screen.MousePointer = vbDefault
            Set adoCmd = Nothing
            Call ss3_clear
            Conn.RollbackTrans
            Sp_Process = False
            Exit Function
        
         End If
                        
        Conn.CommitTrans
        MDIMain.StatusBar1.Panels(1) = "提示信息: 数据处理完成"
        Text_UPD = ""
        Screen.MousePointer = vbDefault
        Set adoCmd = Nothing
        Call Form_Ref
    
    End With
    
    Exit Function

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("SpreadPro_Error : " & Error)

End Function

Private Sub ss4_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
       Text_UPD.Text = "Update"
    End If
End Sub

'Private Sub ss4_Advance(ByVal AdvanceNext As Boolean)
'    If Gf_Sc_Authority(sAuthority, "U") Then
'       Text_UPD.Text = "Update"
'    End If
'End Sub

Private Sub SSCommand_CREAT_Click()

    Dim OutParam(1, 4)      As Variant
    Dim adoCmd              As ADODB.Command
    Dim Response            As Variant

    On Error GoTo Process_Exec_ERROR
    
    Response = MsgBox("生成" + Mid(txt_DATE.RawData, 1, 4) + "年" + Mid(txt_DATE.RawData, 5, 2) + "月" + Mid(txt_DATE.RawData, 7, 2) + "日  " + "新报表吗?", vbYesNo, "系统提示信息")
    If Response = vbNo Then
        Exit Sub
    End If
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
             
    Screen.MousePointer = vbHourglass
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = "{call AKP1111P ('" + txt_DATE.RawData + "')}"

    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    Call MsgBox("生成报表！", vbInformation, "系统提示信息")
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Call Form_Ref
    Exit Sub
    
Process_Exec_ERROR:
    
    Set adoCmd = Nothing
    Call Gp_MsgBoxDisplay(Err.Description & "{call AKP1111P ('" + txt_DATE.RawData + "')}")
    
End Sub

Private Sub ExcelPrn()
    Dim i               As Integer
    Dim IDc             As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDate           As String
    Dim sExlRange       As String
    Dim sTwoStlgrd      As Integer
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AKP1213C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
       
    sDate = txt_DATE.Text
    xlApp.Range("A3").Value = "报表日期：" + Left(sDate, 4) + "年" + Mid(sDate, 6, 2) + "月" + Mid(sDate, 9, 2) + "日"
    xlApp.Range("D54").Value = Format(Now, "YYYY-MM-DD HH:MM:SS")
    xlApp.Range("Y54").Value = "制表人： " & sUserID
    
    Clipboard.Clear
    ss1.SetSelection 1, 1, ss1.MaxCols, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("C6").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss2.SetSelection 1, 1, ss2.MaxCols - 5, ss2.MaxRows
    ss2.ClipboardCopy
    xlApp.Range("C18").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss2.SetSelection ss2.MaxCols - 4, 1, ss2.MaxCols, ss2.MaxRows
    ss2.ClipboardCopy
    xlApp.Range("Q42").Select
    xlApp.ActiveSheet.Paste
    
    
   '-----------------Peocess ss3-----------------------------------------
    With ss3
        ''
        Clipboard.Clear
        ss3.SetSelection 2, 2, 2, 12
        ss3.ClipboardCopy
        xlApp.Range("W41").Select
        xlApp.ActiveSheet.Paste
        
        Clipboard.Clear
        ss3.SetSelection 3, 9, 3, 12
        ss3.ClipboardCopy
        xlApp.Range("X48").Select
        xlApp.ActiveSheet.Paste
        
        Clipboard.Clear
        ss3.SetSelection 4, 2, 4, 12
        ss3.ClipboardCopy
        xlApp.Range("Y41").Select
        xlApp.ActiveSheet.Paste
         
        Clipboard.Clear
        ss3.SetSelection 6, 2, 6, 5
        ss3.ClipboardCopy
        xlApp.Range("AA41").Select
        xlApp.ActiveSheet.Paste
        
        .Col = 5
            .Row = 6:    xlApp.Range("Z45").Value = "至  " & .Text
       
        Clipboard.Clear
        ss3.SetSelection 6, 7, 6, 12
        ss3.ClipboardCopy
        xlApp.Range("AA46").Select
        xlApp.ActiveSheet.Paste

        Clipboard.Clear
        ss3.SetSelection 8, 3, 8, 12
        ss3.ClipboardCopy
        xlApp.Range("AC42").Select
        xlApp.ActiveSheet.Paste

    End With

    '-----------------Peocess ss4-----------------------------------------
    xlApp.Application.Visible = True
    If ss4.MaxCols > 27 Then
        IDc = 27
        sTwoStlgrd = IIf(ss4.MaxCols > 41, 41, ss4.MaxCols)
    Else
        IDc = ss4.MaxCols
    End If
    
    For i = 1 To IDc
        ss4.Col = i
        ss4.Row = 0
        If i < 25 Then
            sExlRange = Chr(i + 66) & 28
        Else
            sExlRange = "A" & Chr(i + 40) & 28
        End If
        xlApp.Range(sExlRange).Value = ss4.Text
    Next i

    Clipboard.Clear
    ss4.SetSelection 1, 1, IDc, 10
    ss4.ClipboardCopy
    xlApp.Range("C30").Select
    xlApp.ActiveSheet.Paste
        
    IDc = 0
    If ss4.MaxCols > 27 Then
        For i = 28 To sTwoStlgrd
            IDc = IDc + 1
            
            ss4.Col = i
            ss4.Row = 0
            If IDc < 25 Then
               sExlRange = Chr(IDc + 66) & 40
            Else
               sExlRange = "A" & Chr(IDc + 40) & 40
            End If
            'sExlRange = Chr(IDc + 66) & 40
            xlApp.Range(sExlRange).Value = ss4.Text
        Next i
    
        Clipboard.Clear
        ss4.SetSelection 28, 1, sTwoStlgrd, 10
        ss4.ClipboardCopy
        xlApp.Range("C42").Select
        xlApp.ActiveSheet.Paste
    End If
    ss4.Col = 1: ss4.Row = 11: xlApp.Range("C52").Value = ss4.Text
'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
    ss1.ClearSelection
    ss2.ClearSelection
    ss3.ClearSelection
    ss4.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    
'    xlApp.Application.Visible = True
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


