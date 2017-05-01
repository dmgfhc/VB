VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CKP1010C 
   Caption         =   "中板厂生产简报_CKP1010C"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter AW 
      Height          =   8625
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   15214
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "CKP1010C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   1710
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   3016
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
         MaxCols         =   23
         MaxRows         =   10
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKP1010C.frx":0072
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   2535
         Left            =   0
         TabIndex        =   2
         Top             =   1800
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   4471
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
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
         MaxCols         =   34
         MaxRows         =   6
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKP1010C.frx":14BD
         UnitType        =   0
         TextTip         =   1
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   4200
         Left            =   0
         TabIndex        =   6
         Top             =   4425
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   7408
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
         MaxCols         =   30
         MaxRows         =   49
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKP1010C.frx":2B63
      End
   End
   Begin Threed.SSFrame Single 
      Height          =   555
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   979
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
      Begin Threed.SSCommand Cmd_Edit 
         Height          =   360
         Left            =   10335
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   635
         _Version        =   196609
         Font3D          =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "更新数据"
      End
      Begin InDate.UDate txt_DATE 
         Height          =   315
         Left            =   2595
         TabIndex        =   5
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   1410
         Top             =   120
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         Caption         =   "日期"
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
   End
End
Attribute VB_Name = "CKP1010C"
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
'-- Program Name      PROD REPORT
'-- Program ID        AGC2600C
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2005.08.21
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
Public QueryYN As Boolean

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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

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
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="CKP1010C.P_SREFER1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc1"

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
    
    'Spread_Collection
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="CKP1010C.P_SREFER2", Key:="P-R"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
    
        'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 16, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 17, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 18, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 19, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 20, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 21, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 22, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 23, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 24, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 25, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 26, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 27, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 28, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 29, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 30, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss1.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"

    ss3.Col = 2: ss3.Col2 = 2
    ss3.ROW = 49: ss3.Row2 = 49

    ss3.Lock = False
    ss3.BlockMode = False
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
        
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Sp_Setting(Sc1.Item("Spread"))
    Call Sp_Setting(Sc2.Item("Spread"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "G-System.INI", Me.Name)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
       Cmd_Edit.Enabled = True
    End If

    txt_DATE.RawData = Format(Date - 1, "yyyymmdd")
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "G-System.INI", Me.Name)
    
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
   
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
End Sub

Public Sub Form_Cls()
    Dim iRow  As Long
    Dim iCol  As Long
    
    Call Form_SP_Cls
    
'    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)

    txt_DATE.RawData = Format(Date - 1, "yyyymmdd")
End Sub

Public Sub Form_SP_Cls()
    Dim iRow  As Long
    Dim iCol  As Long
    
    For iRow = 1 To ss1.MaxRows
        ss1.ROW = iRow
        For iCol = 1 To ss1.MaxCols
           ss1.Col = iCol
           ss1.Text = ""
        Next iCol
    Next iRow
    
    For iRow = 1 To ss2.MaxRows
        ss2.ROW = iRow
        For iCol = 1 To ss2.MaxCols
           If iCol <> 21 And iCol <> 22 Then
              ss2.Col = iCol
              ss2.Text = ""
           End If
        Next iCol
    Next iRow
    
    For iRow = 1 To ss3.MaxRows
         ss3.ROW = iRow
         For iCol = 3 To ss3.MaxCols
             ss3.Col = iCol
             If ss3.CellType = SS_CELL_TYPE_NUMBER Then
                ss3.Text = ""
             End If
         Next iCol
    Next iRow
End Sub

Public Sub Form_Ref()
    
    If Trim(txt_DATE.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_DATE.Tag + "必须输入")
        Exit Sub
    End If
    
    Call Form_SP_Cls
    Screen.MousePointer = vbHourglass
    
    If Sp_Display(M_CN1, Proc_Sc("Sc1")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc1").Item("P-R"), "R", Mc1("pControl"))) Then
       Call Sp_Display2(M_CN1, Proc_Sc("Sc2")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
        Call SearchStlGrdData
        Call SearchCommentsData
    End If
    
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub Form_Exc()

'    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call ExcelPrn
    
End Sub

Public Sub Form_Pro()
    Dim sQuery      As String
    Dim sComments   As String
    Dim sDate       As String
    
    On Error GoTo UPDATE_ERROR

    Screen.MousePointer = vbHourglass
    
    M_CN1.BeginTrans
 
    ss3.ROW = 49
    ss3.Col = 2
    sComments = Trim(ss3.Text)
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    
    sQuery = ""
    sQuery = "         UPDATE  gp_zbrpt_mon                                 " & vbCrLf
    sQuery = sQuery & "   SET  COMMENT1         = '" & sComments & "'       " & vbCrLf
    sQuery = sQuery & " WHERE  PLT              = 'C3'                      " & vbCrLf
    sQuery = sQuery & "   AND  PROD_DATE        = '" & sDate & "'           " & vbCrLf

    M_CN1.Execute sQuery
        
    M_CN1.CommitTrans

    Screen.MousePointer = vbDefault
    Exit Sub

UPDATE_ERROR:

    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay(Err.Description & sQuery)
    
    M_CN1.RollbackTrans
    
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

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
Sub SearchStlGrdData()

    Dim AdoRs As New ADODB.Recordset
    Dim sql               As String
    Dim sDate             As String
    Dim istlknd           As Integer
    Dim iRow              As Integer
    Dim iCol              As Integer
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    
    sql = "       SELECT  PROD_GROUP || '0' PROD_GROUP,                                             " & vbCrLf
    sql = sql & "         TO_NUMBER(STLGRD_KND) STLGRD_KND,                                         " & vbCrLf
    sql = sql & "         SUM(PROD_WGT)         PROD_WGT  ,                                         " & vbCrLf
    sql = sql & "         DECODE(SUM(SLAB_WGT),0,0,ROUND(SUM(PROD_WGT) * 100/SUM(SLAB_WGT),2)) PROD_RATE,       " & vbCrLf
    sql = sql & "         DECODE((SUM(PLATE_WGT)-SUM(INSP_READY_WGT)),0,0,ROUND(SUM(PROD_OK_WGT) * 100/(SUM(PLATE_WGT)-SUM(INSP_READY_WGT)),2)) MILL_OK_RATE, " & vbCrLf
    sql = sql & "         DECODE(SUM(WORK_TIME),0,0,ROUND(SUM(PROD_WGT)*3600/SUM(WORK_TIME),3)) TON_HOUR       " & vbCrLf
    sql = sql & "   From  GP_ZBRPT_DAILY_STDSPEC                                                    " & vbCrLf
    sql = sql & "  Where  PROD_DATE = '" & sDate & "'                                               " & vbCrLf
    sql = sql & "  GROUP  BY PROD_GROUP, STLGRD_KND                                                 " & vbCrLf
    sql = sql & "  Union  All                                                                       " & vbCrLf
    sql = sql & " SELECT  PROD_GROUP || '1' PROD_GROUP,                                             " & vbCrLf
    sql = sql & "         TO_NUMBER(STLGRD_KND)  STLGRD_KND ,                                       " & vbCrLf
    sql = sql & "         SUM(PROD_WGT)          PROD_WGT   ,                                       " & vbCrLf
    sql = sql & "         DECODE(SUM(PROD_WGT),0,0,ROUND(SUM(PROD_WGT) * 100/SUM(SLAB_WGT),2)) PROD_RATE,      " & vbCrLf
    sql = sql & "         DECODE((SUM(PLATE_WGT)-SUM(INSP_READY_WGT)),0,0,ROUND(SUM(PROD_OK_WGT) * 100/(SUM(PLATE_WGT)-SUM(INSP_READY_WGT)),2)) MILL_OK_RATE," & vbCrLf
    sql = sql & "         DECODE(SUM(WORK_TIME),0,0,ROUND(SUM(PROD_WGT)*3600/SUM(WORK_TIME),3)) TON_HOUR      " & vbCrLf
    sql = sql & "   From  GP_ZBRPT_DAILY_STDSPEC                                                    " & vbCrLf
    sql = sql & "  Where  PROD_DATE <= '" & sDate & "'                                              " & vbCrLf
    sql = sql & "    AND  SUBSTR(PROD_DATE, 1, 6) = '" & Left(sDate, 6) & "'                        " & vbCrLf
    sql = sql & "  GROUP  BY PROD_GROUP, STLGRD_KND                                                 " & vbCrLf
    sql = sql & "  Union  All                                                                       " & vbCrLf
    sql = sql & " SELECT  'T0'    PROD_GROUP,                                                       " & vbCrLf
    sql = sql & "         TO_NUMBER(STLGRD_KND)  STLGRD_KND ,                                       " & vbCrLf
    sql = sql & "         SUM(PROD_WGT)          PROD_WGT   ,                                       " & vbCrLf
    sql = sql & "         DECODE(SUM(PROD_WGT),0,0,ROUND(SUM(PROD_WGT) * 100/SUM(SLAB_WGT),2)) PROD_RATE,      " & vbCrLf
    sql = sql & "         DECODE((SUM(PLATE_WGT)-SUM(INSP_READY_WGT)),0,0,ROUND(SUM(PROD_OK_WGT) * 100/(SUM(PLATE_WGT)-SUM(INSP_READY_WGT)),2)) MILL_OK_RATE," & vbCrLf
    sql = sql & "         DECODE(SUM(WORK_TIME),0,0,ROUND(SUM(PROD_WGT)*3600/SUM(WORK_TIME),3)) TON_HOUR      " & vbCrLf
    sql = sql & "   From  GP_ZBRPT_DAILY_STDSPEC                                                    " & vbCrLf
    sql = sql & "  Where  PROD_DATE = '" & sDate & "'                                               " & vbCrLf
    sql = sql & "  GROUP  BY STLGRD_KND                                                             " & vbCrLf
    sql = sql & "  Union  All                                                                       " & vbCrLf
    sql = sql & " SELECT  'T1'    PROD_GROUP,                                                       " & vbCrLf
    sql = sql & "         TO_NUMBER(STLGRD_KND)  STLGRD_KND ,                                       " & vbCrLf
    sql = sql & "         SUM(PROD_WGT)          PROD_WGT   ,                                       " & vbCrLf
    sql = sql & "         DECODE(SUM(PROD_WGT),0,0,ROUND(SUM(PROD_WGT) * 100/SUM(SLAB_WGT),2)) PROD_RATE,      " & vbCrLf
    sql = sql & "         DECODE((SUM(PLATE_WGT)-SUM(INSP_READY_WGT)),0,0,ROUND(SUM(PROD_OK_WGT) * 100/(SUM(PLATE_WGT)-SUM(INSP_READY_WGT)),2)) MILL_OK_RATE," & vbCrLf
    sql = sql & "         DECODE(SUM(WORK_TIME),0,0,ROUND(SUM(PROD_WGT)*3600/SUM(WORK_TIME),3)) TON_HOUR      " & vbCrLf
    sql = sql & "   From  GP_ZBRPT_DAILY_STDSPEC                                                    " & vbCrLf
    sql = sql & "  Where  PROD_DATE <= '" & sDate & "'                                              " & vbCrLf
    sql = sql & "    AND  SUBSTR(PROD_DATE, 1, 6) = '" & Left(sDate, 6) & "'                       " & vbCrLf
    sql = sql & "  GROUP  BY STLGRD_KND                                                             " & vbCrLf
        
    AdoRs.Open sql, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    iRow = 0
    
    Do Until AdoRs.EOF
        
        Select Case Left(AdoRs.Fields("PROD_GROUP"), 1)
            Case "A"
                If AdoRs.Fields("STLGRD_KND") <= 7 Then
                   iRow = 3
                ElseIf AdoRs.Fields("STLGRD_KND") > 7 And AdoRs.Fields("STLGRD_KND") <= 14 Then
                   iRow = 15
                ElseIf AdoRs.Fields("STLGRD_KND") > 14 And AdoRs.Fields("STLGRD_KND") <= 18 Then
                   iRow = 27
                End If
            Case "B"
                If AdoRs.Fields("STLGRD_KND") <= 7 Then
                   iRow = 5
                ElseIf AdoRs.Fields("STLGRD_KND") > 7 And AdoRs.Fields("STLGRD_KND") <= 14 Then
                   iRow = 17
                ElseIf AdoRs.Fields("STLGRD_KND") > 14 And AdoRs.Fields("STLGRD_KND") <= 18 Then
                   iRow = 29
                End If
            Case "C"
                If AdoRs.Fields("STLGRD_KND") <= 7 Then
                   iRow = 7
                ElseIf AdoRs.Fields("STLGRD_KND") > 7 And AdoRs.Fields("STLGRD_KND") <= 14 Then
                   iRow = 19
                ElseIf AdoRs.Fields("STLGRD_KND") > 14 And AdoRs.Fields("STLGRD_KND") <= 18 Then
                   iRow = 31
                End If
            Case "D"
                If AdoRs.Fields("STLGRD_KND") <= 7 Then
                   iRow = 9
                ElseIf AdoRs.Fields("STLGRD_KND") > 7 And AdoRs.Fields("STLGRD_KND") <= 14 Then
                   iRow = 21
                ElseIf AdoRs.Fields("STLGRD_KND") > 14 And AdoRs.Fields("STLGRD_KND") <= 18 Then
                   iRow = 33
                End If
            Case "T"
                If AdoRs.Fields("STLGRD_KND") <= 7 Then
                   iRow = 11
                ElseIf AdoRs.Fields("STLGRD_KND") > 7 And AdoRs.Fields("STLGRD_KND") <= 14 Then
                   iRow = 23
                ElseIf AdoRs.Fields("STLGRD_KND") > 14 And AdoRs.Fields("STLGRD_KND") <= 18 Then
                   iRow = 35
                End If
        End Select
        
        iRow = iRow + Mid(AdoRs.Fields("PROD_GROUP"), 2, 1)
        
        If AdoRs.Fields("STLGRD_KND") <= 7 Then
            iCol = AdoRs.Fields("STLGRD_KND") * 4 - 1
        ElseIf AdoRs.Fields("STLGRD_KND") > 7 And AdoRs.Fields("STLGRD_KND") <= 14 Then
            iCol = (AdoRs.Fields("STLGRD_KND") - 7) * 4 - 1
        ElseIf AdoRs.Fields("STLGRD_KND") > 14 And AdoRs.Fields("STLGRD_KND") <= 18 Then
            iCol = (AdoRs.Fields("STLGRD_KND") - 14) * 4 - 1
        Else
            iCol = 0
        End If
        
        ss3.ROW = iRow
        ss3.Col = iCol
        If Not (VarType(AdoRs.Fields("PROD_WGT")) = vbNull Or AdoRs.Fields("PROD_WGT").Value = 0) Then
            ss3.Text = Val(AdoRs.Fields("PROD_WGT"))
        End If
        
        ss3.Col = iCol + 1
        If Not (VarType(AdoRs.Fields("PROD_RATE")) = vbNull Or AdoRs.Fields("PROD_RATE").Value = 0) Then
            ss3.Text = Val(AdoRs.Fields("PROD_RATE"))
        End If
             
        ss3.Col = iCol + 2
        If Not (VarType(AdoRs.Fields("MILL_OK_RATE")) = vbNull Or AdoRs.Fields("MILL_OK_RATE").Value = 0) Then
            ss3.Text = Val(AdoRs.Fields("MILL_OK_RATE"))
        End If
        
        ss3.Col = iCol + 3
        If Not (VarType(AdoRs.Fields("TON_HOUR")) = vbNull Or AdoRs.Fields("TON_HOUR").Value = 0) Then
            ss3.Text = Val(AdoRs.Fields("TON_HOUR"))
        End If
        
        iRow = 0
        iCol = 0
                
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
      
End Sub

Sub SearchCommentsData()

    Dim AdoRs As New ADODB.Recordset
    Dim sql               As String
    Dim sDate             As String
    Dim i                 As Integer
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    
    sql = "      SELECT  MONTH_PLAN_WGT  ,   MONTH_FIN_WGT   , MONTH_AVE_WGT    ,  " & vbCrLf
    sql = sql + "        MONTH_FOR_WGT   ,   MONTH_LEFT_DAY  , MONTH_AVE_NEED   ,  " & vbCrLf
    sql = sql + "        MONTH_PROG      ,   MONTH_PROD_PROG , MONTH_DAY_BED_WGT,  " & vbCrLf
    sql = sql + "        MONTH_BED_WGT   ,   YEAR_PLAN_WGT   , YEAR_FIN_WGT     ,  " & vbCrLf
    sql = sql + "        YEAR_AVE_WGT    ,   YEAR_FOR_WGT    , YEAR_LEFT_DAY    ,  " & vbCrLf
    sql = sql + "        YEAR_AVE_NEED   ,   YEAR_PROG       , YEAR_PROD_PROG   ,  " & vbCrLf
    sql = sql + "        COMMENT1                                                  " & vbCrLf
    sql = sql & "  FROM  gp_zbrpt_mon                                              " & vbCrLf
    sql = sql & " WHERE  PLT                     = 'C3'                            " & vbCrLf
    sql = sql & "   AND  PROD_DATE               = '" & sDate & "'                 " & vbCrLf
    
    AdoRs.Open sql, M_CN1, adOpenForwardOnly, adLockReadOnly
    If Not AdoRs.EOF Then
       With ss3
                .Col = 25
                For i = 39 To 48
                    .ROW = i
                    If Not (VarType(AdoRs.Fields(i - 39)) = vbNull Or AdoRs.Fields(i - 39).Value = 0) Then
                      .Text = Val(AdoRs.Fields(i - 39))
                    End If
                Next
                .Col = 29
                For i = 39 To 46
                    .ROW = i
                    If Not (VarType(AdoRs.Fields(i - 29)) = vbNull Or AdoRs.Fields(i - 29).Value = 0) Then
                      .Text = Val(AdoRs.Fields(i - 29))
                    End If
                Next
                .Col = 2
                .ROW = 49
                If Not VarType(AdoRs.Fields(18)) = vbNull Or AdoRs.Fields(18).Value = 0 Then
                    .Text = AdoRs.Fields(18).Value
                End If
        End With
    End If
    
    AdoRs.Close
    
    
    
End Sub
Private Sub Cmd_Edit_Click()
    'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    Dim sQuery As String
          
    If Trim(txt_DATE.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_DATE.Tag + "必须输入")
        Exit Sub
    End If

    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call CKP1010P ('" + Trim(Format(txt_DATE.Text, "YYYYMMDD")) + "',?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
            
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        strRet_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & strRet_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        
        Call Gp_MsgBoxDisplay("更新成功..!!", "I")
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("更新失败！！")

End Sub


Private Sub ExcelPrn()
    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDate           As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\CKP1010C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    xlApp.Range("A2").Value = "报表日期：" + Left(sDate, 4) + "年" + Mid(sDate, 5, 2) + "月" + Mid(sDate, 7, 2) + "日"
    xlApp.Range("B72").Value = "制表日期：" + Format(Now, "YYYY-MM-DD HH:MM:SS")
    xlApp.Range("W72").Value = "制表人：" + sUserID

    Clipboard.Clear
    ss1.SetSelection 1, 1, ss1.MaxCols, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("C5").Select
    xlApp.ActiveSheet.Paste

    Clipboard.Clear
    ss2.SetSelection 1, 1, ss2.MaxCols, ss2.MaxRows
    ss2.ClipboardCopy
    xlApp.Range("B17").Select
    xlApp.ActiveSheet.Paste

    Clipboard.Clear
    ss3.SetSelection 3, 3, ss3.MaxCols, 12
    ss3.ClipboardCopy
    xlApp.Range("C25").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss3.SetSelection 3, 15, ss3.MaxCols, 24
    ss3.ClipboardCopy
    xlApp.Range("C37").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss3.SetSelection 3, 27, ss3.MaxCols, 36
    ss3.ClipboardCopy
    xlApp.Range("C49").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss3.SetSelection 25, 39, 25, ss3.MaxRows - 1
    ss3.ClipboardCopy
    xlApp.Range("Y61").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss3.SetSelection 29, 39, 29, ss3.MaxRows - 1
    ss3.ClipboardCopy
    xlApp.Range("AC61").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss3.SetSelection 2, ss3.MaxRows, 2, ss3.MaxRows
    ss3.ClipboardCopy
    xlApp.Range("B71").Select
    xlApp.ActiveSheet.Paste
    
    ss1.ClearSelection
    ss2.ClearSelection
    ss3.ClearSelection
       
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

Public Sub Sp_Setting(ByVal sPname As Variant, Optional MsgChk As Boolean = True)
    With sPname
    
        .RowHeight(-1) = 12.54
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 12
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 12
        Else
            .RowHeight(0) = 24
        End If
        
        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        .RetainSelBlock = True

        .UserResize = UserResizeColumns
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .ROW = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .ROW = 0
        .FontBold = True
        
        
        If MsgChk Then
            .LockBackColor = RGB(255, 255, 255)
        End If

    End With
    
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
                        .ROW = 1
                    Case "A1"
                        .ROW = 2
                    Case "B0"
                        .ROW = 3
                    Case "B1"
                        .ROW = 4
                    Case "C0"
                        .ROW = 5
                    Case "C1"
                        .ROW = 6
                    Case "D0"
                        .ROW = 7
                    Case "D1"
                        .ROW = 8
                    Case "T0"
                        .ROW = 9
                    Case "T1"
                        .ROW = 10
                End Select
            
'            .ROW = iRowCount + 1

                For iColcount = 1 To .MaxCols
    
                    .Col = iColcount
    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or ArrayRecords(iColcount, iRowCount) = 0 Then
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

Public Function Sp_Display2(Conn As ADODB.Connection, sPname As Variant, sQuery As String) As Boolean

    On Error Resume Next

    Dim iCount          As Integer
    Dim iRowCount       As Long
    Dim iColcount       As Long
    Dim AdoRs           As ADODB.Recordset
    Dim ArrayRecords    As Variant

    Sp_Display2 = True

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display2 = False: Exit Function
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
            Sp_Display2 = False
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Call Form_Cls
            Screen.MousePointer = vbDefault
            Exit Function

        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) <> 0 Then

            For iRowCount = 0 To UBound(ArrayRecords, 2)
                Select Case Left(ArrayRecords(0, iRowCount), 1)
                    Case "A"
                        .ROW = 1
                    Case "B"
                        .ROW = 2
                    Case "C"
                        .ROW = 3
                    Case "D"
                        .ROW = 4
                    Case "T"
                        .ROW = 5
                End Select
            
'            .ROW = iRowCount + 1

                For iColcount = 1 To 10  '10 --> 9 by guhf 2011.5.12 删除压力空气
    
                    If Mid(ArrayRecords(0, iRowCount), 2, 1) = "0" Then
                        .Col = iColcount * 2 - 1  '奇数列
                    ElseIf Mid(ArrayRecords(0, iRowCount), 2, 1) = "1" Then
                        .Col = iColcount * 2  '偶数列
                    End If
    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or ArrayRecords(iColcount, iRowCount) = 0 Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iColcount, iRowCount))
                    End If

                Next iColcount
                
                If ArrayRecords(0, iRowCount) = "A0" Then
                   ss3.ROW = 28
                   ss3.Col = 20
                   If VarType(ArrayRecords(47, iRowCount)) = vbNull Or ArrayRecords(47, iRowCount) = 0 Then  '29 --> 46 by guhf 2011.5.12
                      ss3.Text = ""
                   Else
                      ss3.Text = Trim(ArrayRecords(47, iRowCount))  '29 --> 46 by guhf 2011.5.12
                   End If
                ElseIf ArrayRecords(0, iRowCount) = "B0" Then
                   ss3.ROW = 28
                   ss3.Col = 21
                   If VarType(ArrayRecords(47, iRowCount)) = vbNull Or ArrayRecords(47, iRowCount) = 0 Then  '29 --> 46 by guhf 2011.5.12
                      ss3.Text = ""
                   Else
                      ss3.Text = Trim(ArrayRecords(47, iRowCount))  '29 --> 46 by guhf 2011.5.12
                   End If
                ElseIf ArrayRecords(0, iRowCount) = "C0" Then
                   ss3.ROW = 28
                   ss3.Col = 22
                   If VarType(ArrayRecords(47, iRowCount)) = vbNull Or ArrayRecords(47, iRowCount) = 0 Then  '29 --> 46 by guhf 2011.5.12
                      ss3.Text = ""
                   Else
                      ss3.Text = Trim(ArrayRecords(47, iRowCount))  '29 --> 46 by guhf 2011.5.12
                   End If
                ElseIf ArrayRecords(0, iRowCount) = "D0" Then
                   ss3.ROW = 28
                   ss3.Col = 23
                   If VarType(ArrayRecords(47, iRowCount)) = vbNull Or ArrayRecords(47, iRowCount) = 0 Then  '29 --> 46 by guhf 2011.5.12
                      ss3.Text = ""
                   Else
                      ss3.Text = Trim(ArrayRecords(47, iRowCount)) '29 --> 46 by guhf 2011.5.12
                   End If
                End If
                
                If ArrayRecords(0, iRowCount) = "T0" Then
                    For iCount = 48 To 53  ' 30 --> 47 ,35-->52 by guhf 2011.5.12
                       ss3.ROW = 28
                       ss3.Col = iCount - 23 '6 --> 23 by guhf 2011.5.12
                       If VarType(ArrayRecords(iCount, iRowCount)) = vbNull Or ArrayRecords(iCount, iRowCount) = 0 Then
                          ss3.Text = ""
                       Else
                          ss3.Text = Trim(ArrayRecords(iCount, iRowCount))
                       End If
                    Next
                End If
                
                '由于后台SQL不是按照查询顺序写的，所以这段代码的主要含义是将不同的SQL值按照顺序排列到表单之中，需要到后台一列列的去读取顺序并在前台做相应的顺序处理。20161222 ADD HAN
                If Left(ArrayRecords(0, iRowCount), 1) = "T" Then
                    For iCount = 11 To 46   '11-->10 ,28-->45 by guhf 2011.5.12
                        If ArrayRecords(0, iRowCount) = "T0" Then
                           If iCount <= 18 Then
                              .ROW = 1
                              .Col = iCount + 12    '第一行23列开始共8列
                           ElseIf iCount >= 19 And iCount < 23 Then    '总累计待切割量
                              .ROW = 2
                              .Col = iCount + 12
                           ElseIf iCount >= 23 And iCount < 31 Then   '17-->22 ,23-->30 by guhf
                              .ROW = 3
                              .Col = iCount    '+6 --> -1 by guhf 2011.5.12
                           ElseIf iCount >= 31 And iCount < 35 Then    'add by guhf 2011.5.12 增加总累计待探伤量
                              .ROW = 4                                 'add by guhf 2011.5.12
                              .Col = iCount                       'add by guhf 2011.5.12
                           ElseIf iCount >= 35 And iCount <= 42 Then  ''MODIFIED BY GUOLI AT 20100318180800 FOR 避免SS2最后一列不显示数据,原来没有=
                                            '23-->34, 28-->41 by guhf
                              .ROW = 5
                              .Col = iCount - 12 ' 又减了12
                           ElseIf iCount >= 43 And iCount < 47 Then    'add by guhf 2011.5.12 增加总累计待切割两
                              .ROW = 6                                 'add by guhf 2011.5.12
                              .Col = iCount - 12                       'add by guhf 2011.5.12
                           End If
                    '
                        ElseIf ArrayRecords(0, iRowCount) = "T1" Then
                           If iCount <= 18 Then      '17-->18  by guhf 2011.5.12
                              .ROW = 2
                              .Col = iCount + 12    '12--> 11 by guhf 2011.5.12
                           ElseIf iCount >= 23 And iCount < 31 Then   '17-->22 ,23-->30 by guhf
                              .ROW = 4
                              .Col = iCount    '+6 --> -1 by guhf 2011.5.12
                           ElseIf iCount >= 35 And iCount <= 42 Then  ''MODIFIED BY GUOLI AT 20100318180800 FOR 避免SS2最后一列不显示数据,原来没有=
                                            '23-->34, 28-->41 by guhf
                              .ROW = 6
                              .Col = iCount - 12 ' add - 13 by guhf
'                           If iCount < 17 Then
'                              .ROW = 2
'                              .Col = iCount + 12
'                           ElseIf iCount >= 17 And iCount < 23 Then
'                              .ROW = 4
'                              .Col = iCount + 6
'                           ElseIf iCount >= 23 And iCount <= 28 Then  ''MODIFIED BY GUOLI AT 20100318180800 FOR 避免SS2最后一列不显示数据,原来没有=
'                              .ROW = 6
'                              .Col = iCount
                           End If
                        End If
                        If VarType(ArrayRecords(iCount, iRowCount)) = vbNull Or ArrayRecords(iCount, iRowCount) = 0 Then
                          .Text = ""
                        Else
                          .Text = Trim(ArrayRecords(iCount, iRowCount))
                        End If
                    Next
                End If
            
            Next iRowCount
            
        End If

        .ReDraw = True
        Screen.MousePointer = vbDefault

    End With

End Function


