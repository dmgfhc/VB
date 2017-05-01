VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKP3052C 
   Caption         =   "中厚板卷厂生产日报_AKP3052C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter AW 
      Height          =   8670
      Left            =   30
      TabIndex        =   0
      Top             =   615
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   15293
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      PaneTree        =   "AKP3052C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   2415
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   4260
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
         MaxCols         =   26
         MaxRows         =   10
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP3052C.frx":0092
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3855
         Left            =   0
         TabIndex        =   2
         Top             =   2475
         Width           =   5040
         _Version        =   393216
         _ExtentX        =   8890
         _ExtentY        =   6800
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
         MaxCols         =   27
         MaxRows         =   10
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP3052C.frx":162E
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   2280
         Left            =   0
         TabIndex        =   3
         Top             =   6390
         Width           =   15120
         _Version        =   393216
         _ExtentX        =   26670
         _ExtentY        =   4022
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
         MaxCols         =   65
         MaxRows         =   11
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP3052C.frx":25C9
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   3855
         Left            =   5100
         TabIndex        =   4
         Top             =   2475
         Width           =   10020
         _Version        =   393216
         _ExtentX        =   17674
         _ExtentY        =   6800
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
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
         MaxCols         =   16
         MaxRows         =   12
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP3052C.frx":6D81
      End
   End
   Begin Threed.SSFrame Single 
      Height          =   555
      Left            =   30
      TabIndex        =   5
      Top             =   30
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
         ItemData        =   "AKP3052C.frx":9E06
         Left            =   5850
         List            =   "AKP3052C.frx":9E10
         TabIndex        =   6
         Tag             =   "工厂代码"
         Top             =   120
         Width           =   735
      End
      Begin Threed.SSCommand Cmd_Edit 
         Height          =   360
         Left            =   10335
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   90
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   635
         _Version        =   196609
         Font3D          =   1
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
         TabIndex        =   8
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
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   4635
         Top             =   120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "工厂代码"
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
Attribute VB_Name = "AKP3052C"
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
'-- Program ID        AKP3052C
'-- Designer          YANGMENG
'-- Coder             YANGMENG
'-- Date              2007.01.25
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

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
    Dim I As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Sheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_DATE, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(CBO_PLT, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

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
    sc1.Add Item:=ss1, Key:="Spread"
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

     
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AKP3052C.P_SREFER2", Key:="P-R"
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
 
    'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AKP1111C.P_MODIFY", Key:="P-M"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss1.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=sc3, Key:="Sc3"
    
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
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc4, Key:="sc4"
    
    ss4.BlockMode = True
    
        ss4.Col = 1: ss4.Col2 = ss4.MaxCols
        ss4.Row = 1: ss4.Row2 = ss4.MaxRows
        ss4.Lock = True
        
        Call Sp_ColLock(ss4, 1, 11, False)
        
    ss4.BlockMode = False
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
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
    
    Call Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc4")("Spread"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "Z-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Z-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc4")("Spread"), "Z-System.INI", Me.Name)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
       Cmd_Edit.Enabled = True
    End If

    txt_DATE.RawData = Format(Date - 1, "yyyymmdd")
    CBO_PLT.ListIndex = 0
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

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
    Set sc3 = Nothing
    Set sc4 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
End Sub

Public Sub Form_Cls()

    Dim iRow  As Long
    Dim iCol  As Long

    ss1.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, True
    ss2.ClearRange 1, 1, ss2.MaxCols, ss2.MaxRows, True
    Call ss3_clear
    ss4.ClearRange 1, 1, ss4.MaxCols, ss4.MaxRows, True
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
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
    
'
'    txt_DATE.RawData = Format(Date - 1, "yyyymmdd")
'    CBO_PLT.ListIndex = 0
    
End Sub

Public Sub Form_SP_Cls()
    Dim iRow  As Long
    Dim iCol  As Long
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        For iCol = 1 To ss1.MaxCols
           ss1.Col = iCol
           ss1.Text = ""
        Next iCol
    Next iRow
    
    For iRow = 1 To ss2.MaxRows
        ss2.Row = iRow
        For iCol = 1 To ss2.MaxCols
           ss2.Col = iCol
           ss2.Text = ""
        Next iCol
    Next iRow
    
    For iRow = 1 To ss3.MaxRows
         ss3.Row = iRow
         ss3.Col = 2
         If ss3.CellType = SS_CELL_TYPE_NUMBER Then
            For iCol = 2 To ss3.MaxCols
               ss3.Col = iCol
               ss3.Text = ""
            Next iCol
        End If
    Next iRow
End Sub

Public Sub Form_Ref()
    
    If Trim(txt_DATE.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_DATE.Tag + "必须输入")
        Exit Sub
    End If
    
    If Trim(CBO_PLT.Text) = "" Then
        Call Gp_MsgBoxDisplay(CBO_PLT.Tag + "必须输入")
        Exit Sub
    End If
    
    ss1.ReDraw = False
    ss2.ReDraw = False
    ss3.ReDraw = False
    ss4.ReDraw = False
    
    Call Form_Cls
    Screen.MousePointer = vbHourglass
        
    Call SearchProductResultData
    If ss1.MaxRows > 0 Then
        Call Mill_Sp_Display(M_CN1, Proc_Sc("Sc2")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
        Call Ss3_Data_Refer
        Call Ss4_Data_Refer
        Call Ss4_Data_Refer_HC
        Call Zero_Cls
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True
    ss4.ReDraw = True
    
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
    
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub Form_Exc()

    Call ExcelPrn
    
End Sub

Public Sub Form_Pro()
    Dim sQuery      As String
    Dim sComments   As String
    Dim sDate       As String
    Dim lSeq        As Long
    Dim iRow        As Integer
    
    On Error GoTo UPDATE_ERROR

    Screen.MousePointer = vbHourglass
    
    M_CN1.BeginTrans
 
    ss4.Row = 11
    ss4.Col = 1
    sComments = Trim(ss4.Text)
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    
    sQuery = "         UPDATE  GP_RPT_DAILY_SUM                                 " & vbCrLf
    sQuery = sQuery & "   SET  COMMENTS1        = '" & sComments & "'           " & vbCrLf
    sQuery = sQuery & " WHERE  PLT              = '" & Trim(CBO_PLT.Text) & "'  " & vbCrLf
    sQuery = sQuery & "   AND  PROD_DATE        = '" & sDate & "'               " & vbCrLf
    sQuery = sQuery & "   AND  PROD_GROUP       = 'A'                           " & vbCrLf

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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Sub SearchProductResultData()

    Dim AdoRs As New ADODB.Recordset
    Dim sql               As String
    Dim sDate             As String
    Dim sGROUP_CD         As String
    Dim iRow              As Integer
    Dim iCol              As Integer
    
    Dim iCnt              As Integer
    Dim iINSP_READY_WGT(3) As Double
    Dim iINSP_READY_WGT_ALL As Double
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    
    sql = ""
    
    
    sql = " SELECT    DECODE(PROD_GROUP, 'A', 1, 'B', 3, 'C', 5, 7)       DISP_LINE  "
    sql = sql & "                 , DAY_PLAN_WGT        " & vbCrLf
    sql = sql & "                 , MILL_PLATE_WGT      " & vbCrLf
    sql = sql & "                 , COIL_WGT       " & vbCrLf
    sql = sql & "                 , MILL_PLATE_WGT + COIL_WGT      " & vbCrLf
    sql = sql & "                 , MILL_CNT            " & vbCrLf
    sql = sql & "                 , PLATE_WGT           " & vbCrLf
    sql = sql & "                 , PLATE_CNT           " & vbCrLf
    sql = sql & "                 , PLATE_WGT2           " & vbCrLf
    sql = sql & "                 , PLATE_CNT2           " & vbCrLf
    sql = sql & "                 , SLAB_WGT            " & vbCrLf
    sql = sql & "                 , SCRAP_1_WGT         " & vbCrLf
    sql = sql & "                 , SCRAP_2_WGT + SCRAP_3_WGT " & vbCrLf
    sql = sql & "                 , DECODE(SLAB_WGT,0,0,(ORD_PP_WGT + ORD_HC_WGT) * 100/ SLAB_WGT)   " & vbCrLf
    sql = sql & "                 , PRODUCT_S_RATE " & vbCrLf
    sql = sql & "                 , PRODUCT_D_RATE " & vbCrLf
    sql = sql & "                 , PRODUCT_RATE   " & vbCrLf
    sql = sql & "                 , DECODE(PROD_PLAN_SLAB_WGT,0,0,PROD_PLAN_WGT * 100 / PROD_PLAN_SLAB_WGT)    " & vbCrLf
    sql = sql & "                 , TON_HOUR       " & vbCrLf
    sql = sql & "                 , NON_PLAN_TB_WGT  " & vbCrLf
    sql = sql & "                 , SURF_GRD_3 - NON_PLAN_TB_WGT  " & vbCrLf
    sql = sql & "                 , SURF_GRD_2  " & vbCrLf
    sql = sql & "                 , NON_PLAN_WGT  " & vbCrLf
    sql = sql & "                 , NON_PLAN_ERP_WGT  " & vbCrLf
    sql = sql & "                 , WCR_WGT + HCR_WGT      " & vbCrLf
    sql = sql & "                 , COILMILL_WGT       " & vbCrLf
    sql = sql & "                 , DECODE(INSP_THK_CNT,0,0,ROUND(INSP_THK_GC / INSP_THK_CNT,3))   " & vbCrLf
    sql = sql & "             FROM  GP_RPT_DAILY_SUM "
    sql = sql & "            WHERE  PLT        =  '" & Trim(CBO_PLT.Text) & "'" & vbCrLf
    sql = sql & "              AND  PROD_DATE  =  '" & sDate & "'" & vbCrLf
    sql = sql & "       UNION ALL"
    sql = sql & "       SELECT      DECODE(PROD_GROUP, 'A', 2, 'B', 4, 'C', 6, 8)   DISP_LINE"
    sql = sql & "                 , SUM(DAY_PLAN_WGT)        " & vbCrLf
    sql = sql & "                 , SUM(MILL_PLATE_WGT)      " & vbCrLf
    sql = sql & "                 , SUM(COIL_WGT)       " & vbCrLf
    sql = sql & "                 , SUM(MILL_PLATE_WGT) + SUM(COIL_WGT)           " & vbCrLf
    sql = sql & "                 , SUM(MILL_CNT)            " & vbCrLf
    sql = sql & "                 , SUM(PLATE_WGT)           " & vbCrLf
    sql = sql & "                 , SUM(PLATE_CNT)           " & vbCrLf
    sql = sql & "                 , SUM(PLATE_WGT2)           " & vbCrLf
    sql = sql & "                 , SUM(PLATE_CNT2)           " & vbCrLf
    sql = sql & "                 , SUM(SLAB_WGT)            " & vbCrLf
    sql = sql & "                 , SUM(SCRAP_1_WGT)         " & vbCrLf
    sql = sql & "                 , SUM(SCRAP_2_WGT + SCRAP_3_WGT)    " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_WGT),0,0,SUM(ORD_PP_WGT + ORD_HC_WGT) * 100 / SUM(SLAB_WGT))               " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_S_WGT),0,0,SUM(MILL_PLATE_WGT)*100 /SUM(SLAB_S_WGT))                       " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_D_WGT),0,0,SUM(COIL_WGT)*100 /SUM(SLAB_D_WGT))                             " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_WGT),0,0,                                                                  " & vbCrLf
    sql = sql & "                  (SUM(MILL_PLATE_WGT)+SUM(COIL_WGT)) *100 /SUM(SLAB_WGT))                                    " & vbCrLf
    sql = sql & "                 , DECODE(SUM(PROD_PLAN_SLAB_WGT),0,0,SUM(PROD_PLAN_WGT) * 100 / SUM(PROD_PLAN_SLAB_WGT))     " & vbCrLf
    sql = sql & "                 , ROUND(DECODE(SUM(WORK_TIME),0,0,SUM(MILL_PLATE_WGT+COIL_WGT) /(SUM(WORK_TIME)/60)),3)      " & vbCrLf
    sql = sql & "                 , SUM(NON_PLAN_TB_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(SURF_GRD_3 - NON_PLAN_TB_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(SURF_GRD_2)    " & vbCrLf
    sql = sql & "                 , SUM(NON_PLAN_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(NON_PLAN_ERP_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(WCR_WGT+HCR_WGT)         " & vbCrLf
    sql = sql & "                 , SUM(COILMILL_WGT)         " & vbCrLf
    sql = sql & "                 , DECODE(SUM(INSP_THK_CNT),0,0,ROUND(SUM(INSP_THK_GC) / SUM(INSP_THK_CNT),3))                " & vbCrLf
    sql = sql & "             FROM  GP_RPT_DAILY_SUM                                                                           " & vbCrLf
    sql = sql & "            WHERE  PLT        = '" & Trim(CBO_PLT.Text) & "'                                                  " & vbCrLf
    sql = sql & "              AND  PROD_DATE <= '" & sDate & "'                                                               " & vbCrLf
    sql = sql & "              AND  SUBSTR(PROD_DATE,1,6) = '" & Left(sDate, 6) & "'                                           " & vbCrLf
    sql = sql & "            GROUP  BY PROD_GROUP                                                                              " & vbCrLf
    sql = sql & "       UNION ALL                                                                                              " & vbCrLf
    sql = sql & "       SELECT      9   DISP_LINE                                                                              " & vbCrLf
    sql = sql & "                 , SUM(DAY_PLAN_WGT)        " & vbCrLf
    sql = sql & "                 , SUM(MILL_PLATE_WGT)      " & vbCrLf
    sql = sql & "                 , SUM(COIL_WGT)       " & vbCrLf
    sql = sql & "                 , SUM(MILL_PLATE_WGT) + SUM(COIL_WGT)           " & vbCrLf
    sql = sql & "                 , SUM(MILL_CNT)            " & vbCrLf
    sql = sql & "                 , SUM(PLATE_WGT)           " & vbCrLf
    sql = sql & "                 , SUM(PLATE_CNT)           " & vbCrLf
    sql = sql & "                 , SUM(PLATE_WGT2)           " & vbCrLf
    sql = sql & "                 , SUM(PLATE_CNT2)           " & vbCrLf
    sql = sql & "                 , SUM(SLAB_WGT)            " & vbCrLf
    sql = sql & "                 , SUM(SCRAP_1_WGT)         " & vbCrLf
    sql = sql & "                 , SUM(SCRAP_2_WGT + SCRAP_3_WGT)    " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_WGT),0,0,SUM(ORD_PP_WGT + ORD_HC_WGT) * 100 / SUM(SLAB_WGT))               " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_S_WGT),0,0,SUM(MILL_PLATE_WGT)*100 /SUM(SLAB_S_WGT))                       " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_D_WGT),0,0,SUM(COIL_WGT)*100 /SUM(SLAB_D_WGT))                             " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_WGT),0,0,                                                                  " & vbCrLf
    sql = sql & "                  (SUM(MILL_PLATE_WGT)+SUM(COIL_WGT)) *100 /SUM(SLAB_WGT))                                    " & vbCrLf
    sql = sql & "                 , DECODE(SUM(PROD_PLAN_SLAB_WGT),0,0,SUM(PROD_PLAN_WGT) * 100 / SUM(PROD_PLAN_SLAB_WGT))     " & vbCrLf
    sql = sql & "                 , ROUND(DECODE(SUM(WORK_TIME),0,0,SUM(MILL_PLATE_WGT+COIL_WGT) /(SUM(WORK_TIME)/60)),3)      " & vbCrLf
    sql = sql & "                 , SUM(NON_PLAN_TB_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(SURF_GRD_3 - NON_PLAN_TB_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(SURF_GRD_2)    " & vbCrLf
    sql = sql & "                 , SUM(NON_PLAN_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(NON_PLAN_ERP_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(WCR_WGT+HCR_WGT)         " & vbCrLf
    sql = sql & "                 , SUM(COILMILL_WGT)         " & vbCrLf
    sql = sql & "                 , DECODE(SUM(INSP_THK_CNT),0,0,ROUND(SUM(INSP_THK_GC) / SUM(INSP_THK_CNT),3))                " & vbCrLf
    sql = sql & "             FROM  GP_RPT_DAILY_SUM                                                                           " & vbCrLf
    sql = sql & "            WHERE  PLT        = '" & Trim(CBO_PLT.Text) & "'                                                  " & vbCrLf
    sql = sql & "              AND  PROD_DATE  = '" & sDate & "'                                                               " & vbCrLf
    sql = sql & "       UNION ALL                                                                                              " & vbCrLf
    sql = sql & "       SELECT      10       DISP_LINE  " & vbCrLf
    sql = sql & "                 , SUM(DAY_PLAN_WGT)        " & vbCrLf
    sql = sql & "                 , SUM(MILL_PLATE_WGT)      " & vbCrLf
    sql = sql & "                 , SUM(COIL_WGT)       " & vbCrLf
    sql = sql & "                 , SUM(MILL_PLATE_WGT) + SUM(COIL_WGT)           " & vbCrLf
    sql = sql & "                 , SUM(MILL_CNT)            " & vbCrLf
    sql = sql & "                 , SUM(PLATE_WGT)           " & vbCrLf
    sql = sql & "                 , SUM(PLATE_CNT)           " & vbCrLf
    sql = sql & "                 , SUM(PLATE_WGT2)           " & vbCrLf
    sql = sql & "                 , SUM(PLATE_CNT2)           " & vbCrLf
    sql = sql & "                 , SUM(SLAB_WGT)            " & vbCrLf
    sql = sql & "                 , SUM(SCRAP_1_WGT)         " & vbCrLf
    sql = sql & "                 , SUM(SCRAP_2_WGT + SCRAP_3_WGT)    " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_WGT),0,0,SUM(ORD_PP_WGT + ORD_HC_WGT) * 100 / SUM(SLAB_WGT))               " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_S_WGT),0,0,SUM(MILL_PLATE_WGT)*100 /SUM(SLAB_S_WGT))                       " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_D_WGT),0,0,SUM(COIL_WGT)*100 /SUM(SLAB_D_WGT))                             " & vbCrLf
    sql = sql & "                 , DECODE(SUM(SLAB_WGT),0,0,                                                                  " & vbCrLf
    sql = sql & "                  (SUM(MILL_PLATE_WGT)+SUM(COIL_WGT)) *100 /SUM(SLAB_WGT))                                    " & vbCrLf
    sql = sql & "                 , DECODE(SUM(PROD_PLAN_SLAB_WGT),0,0,SUM(PROD_PLAN_WGT) * 100 / SUM(PROD_PLAN_SLAB_WGT))     " & vbCrLf
    sql = sql & "                 , ROUND(DECODE(SUM(WORK_TIME),0,0,SUM(MILL_PLATE_WGT+COIL_WGT) /(SUM(WORK_TIME)/60)),3)      " & vbCrLf
    sql = sql & "                 , SUM(NON_PLAN_TB_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(SURF_GRD_3 - NON_PLAN_TB_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(SURF_GRD_2)    " & vbCrLf
    sql = sql & "                 , SUM(NON_PLAN_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(NON_PLAN_ERP_WGT)    " & vbCrLf
    sql = sql & "                 , SUM(WCR_WGT+HCR_WGT)         " & vbCrLf
    sql = sql & "                 , SUM(COILMILL_WGT)         " & vbCrLf
    sql = sql & "                 , DECODE(SUM(INSP_THK_CNT),0,0,ROUND(SUM(INSP_THK_GC) / SUM(INSP_THK_CNT),3))                " & vbCrLf
    sql = sql & "             FROM  GP_RPT_DAILY_SUM                                                                           " & vbCrLf
    sql = sql & "            WHERE  PLT        = '" & Trim(CBO_PLT.Text) & "'                                                  " & vbCrLf
    sql = sql & "              AND  PROD_DATE <= '" & sDate & "'                                                               " & vbCrLf
    sql = sql & "              AND  SUBSTR(PROD_DATE,1,6) = '" & Left(sDate, 6) & "'                                           " & vbCrLf
    
    AdoRs.Open sql, M_CN1, adOpenForwardOnly, adLockReadOnly

    ss1.ReDraw = False
    Do Until AdoRs.EOF
        ss1.Row = AdoRs.Fields(0)
        For iCol = 1 To ss1.MaxCols
        ss1.Col = iCol
            ss1.Text = Val(AdoRs.Fields(iCol) & "")
        Next iCol
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
    
End Sub
Public Sub Ss4_Data_Refer()

On Error GoTo Ss4_Display_Error

    Dim sQuery      As String
    Dim sStlgrd     As String
    Dim sStlgrdName As String
    Dim sStlgrdHeadName As String
    Dim sgroup      As String
    Dim lWgt        As Double
    Dim IDc         As Integer
    Dim iCol        As Integer
    Dim strTemp     As String
    
    Dim iPPCol      As Integer
    Dim iHCCol      As Integer
    
    Dim AdoRs       As ADODB.Recordset
    Dim AdoRs1      As ADODB.Recordset

    Set AdoRs = New ADODB.Recordset
    Set AdoRs1 = New ADODB.Recordset
    
    iPPCol = 46
    iHCCol = 69
    
    For iCol = 1 To iPPCol
        ss4.Row = 0
        ss4.Col = iCol
        ss4.Text = " "
        Call Gp_Sp_ColHidden(sc4.Item("Spread"), iCol, False)
    Next iCol
    
    sQuery = "SELECT            *   FROM(                                                       " & vbCrLf
    sQuery = sQuery & "SELECT   DECODE(PROD_GROUP,'A',2,'B',4,'C',6,8),    PROD_GROUP,          " & vbCrLf
    sQuery = sQuery & "         APLY_STDSPEC_GROUP, SUM(PROD_WGT)                               " & vbCrLf
    sQuery = sQuery & "   FROM  GP_RPT_DAILY_STDSPEC                                            " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE  <=  '" & txt_DATE.RawData & "'                       " & vbCrLf
    sQuery = sQuery & "    AND  PROD_DATE  >=  SUBSTR('" & txt_DATE.RawData & "',1,6)" & "||'00'" & vbCrLf
    sQuery = sQuery & "    AND  PROD_CD     =  'PP'                                             " & vbCrLf
    sQuery = sQuery & "    AND  APLY_STDSPEC_GROUP     =  'X42~X80'                             " & vbCrLf
    sQuery = sQuery & "    AND  PROD_WGT    >   0                                               " & vbCrLf
    sQuery = sQuery & "  GROUP  BY APLY_STDSPEC_GROUP,PROD_GROUP                                " & vbCrLf
    sQuery = sQuery & "  ORDER  BY PROD_GROUP                                                 ) " & vbCrLf
    
    sQuery = sQuery & "UNION ALL SELECT   *   FROM(                                             " & vbCrLf
    sQuery = sQuery & "SELECT    DECODE(PROD_GROUP,'A',2,'B',4,'C',6,8),    PROD_GROUP,         " & vbCrLf
    sQuery = sQuery & "         APLY_STDSPEC, SUM(PROD_WGT)                                     " & vbCrLf
    sQuery = sQuery & "   FROM  GP_RPT_DAILY_STDSPEC                                            " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE  <=  '" & txt_DATE.RawData & "'                       " & vbCrLf
    sQuery = sQuery & "    AND  PROD_DATE  >=  SUBSTR('" & txt_DATE.RawData & "',1,6)" & "||'00'" & vbCrLf
    sQuery = sQuery & "    AND  PROD_CD     =  'PP'                                             " & vbCrLf
    sQuery = sQuery & "    AND  PROD_WGT    >   0                                               " & vbCrLf
    sQuery = sQuery & "  GROUP  BY APLY_STDSPEC, PROD_GROUP                                     " & vbCrLf
    sQuery = sQuery & "  ORDER  BY 4 DESC, PROD_GROUP                                         ) " & vbCrLf

    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    IDc = 1
    iCol = 1
    Do Until AdoRs.EOF
        
        With ss4
            sgroup = Trim(AdoRs.Fields(1) & "")
            sStlgrd = Trim(AdoRs.Fields(2) & "")
            sStlgrdName = Trim(AdoRs.Fields(2) & "")
            lWgt = Val(AdoRs.Fields(3) & "")
                        
            sStlgrdHeadName = sStlgrdName
            If sStlgrdName = "X42~X80" Then
               sStlgrdHeadName = "管线钢总量"
            End If
            
            If iCol = iPPCol + 1 Then
               Exit Do
            End If
            
            For iCol = 1 To iPPCol
                ss4.Row = 0
                ss4.Col = iCol
                If Trim(ss4.Text) = "" Then
                    .Col = IDc
                    IDc = IDc + 1
                    Exit For
                ElseIf Trim(ss4.Text) = Trim(sStlgrdHeadName) Then
                    .Col = iCol
                    Exit For
                End If
            Next iCol
            
'            For iCol = 1 To iPPCol
'                ss4.ROW = 0
'                ss4.Col = iCol
'                If Trim(ss4.Text) = "" Or Trim(ss4.Text) = Trim(sStlgrdHeadName) Then
'                    .Col = iCol
'                    Exit For
'                End If
'            Next iCol
'
            .Row = 0
            If Trim(.Text) = "" Then
               .Text = sStlgrdHeadName
            End If
            
            
            .Row = Trim(AdoRs.Fields(0) & "")
            If Trim(.Text) = "" Then
                .Text = Val(AdoRs.Fields(3) & "")
            End If
            
            sQuery = "SELECT            DECODE(MAX(PROD_GROUP),'A',1,'B',3,'C',5,7) , " & vbCrLf
            sQuery = sQuery & "         SUM(PROD_WGT)                                 " & vbCrLf
            sQuery = sQuery & "   FROM  GP_RPT_DAILY_STDSPEC                     " & vbCrLf
            sQuery = sQuery & "  WHERE  PROD_DATE   =  '" & txt_DATE.RawData & "'" & vbCrLf
            sQuery = sQuery & "    AND  (APLY_STDSPEC        =  '" & sStlgrdName & "'   " & vbCrLf
            sQuery = sQuery & "     OR  APLY_STDSPEC_GROUP  =  '" & sStlgrdName & "'   )" & vbCrLf
            sQuery = sQuery & "    AND  PROD_GROUP  =  '" & sgroup & "'          " & vbCrLf
            sQuery = sQuery & "    AND  PROD_WGT    >   0                        " & vbCrLf
        
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
    
'    sQuery = "SELECT            COMMENTS1                                                       " & vbCrLf
'    sQuery = sQuery & "   FROM  GP_RPT_DAILY_SUM                                                " & vbCrLf
'    sQuery = sQuery & "  WHERE  PROD_DATE             =  '" & txt_DATE.RawData & "'             " & vbCrLf
'    sQuery = sQuery & "    AND  PROD_GROUP              = 'A'                                   " & vbCrLf
'
'    AdoRs.Open sQuery, M_CN1, adOpenKeyset
'    If Not AdoRs.EOF Then
'          ss4.Col = 1: ss4.ROW = 11: ss4.Text = Trim(AdoRs.Fields(0) & "")
'    End If
'    AdoRs.Close
   
'    For iCol = 1 To iPPCol
'        ss4.ROW = 0
'        ss4.Col = iCol
'        If Trim(ss4.Text) = "" Then Call Gp_Sp_ColHidden(sc4.Item("Spread"), iCol, True)
'    Next iCol
    
    Screen.MousePointer = vbDefault
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Sub

Ss4_Display_Error:
    
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Ss4_Display_Error : " & Error)
    
End Sub
Public Sub Ss4_Data_Refer_HC()

On Error GoTo Ss4_Display_Error

    Dim sQuery      As String
    Dim sStlgrd     As String
    Dim sStlgrdName As String
    Dim sgroup      As String
    Dim lWgt        As Double
    Dim IDc         As Integer
    Dim iCol        As Integer
    Dim strTemp     As String
    
    Dim iPPCol      As Integer
    Dim iHCCol      As Integer
    
    Dim AdoRs       As ADODB.Recordset
    Dim AdoRs1      As ADODB.Recordset

    Set AdoRs = New ADODB.Recordset
    Set AdoRs1 = New ADODB.Recordset
    
    iPPCol = 46
    iHCCol = 69
    
    For iCol = iPPCol + 1 To iHCCol
        ss4.Row = 0
        ss4.Col = iCol
        ss4.Text = " "
        Call Gp_Sp_ColHidden(sc4.Item("Spread"), iCol, False)
    Next iCol
    
    sQuery = "SELECT            DECODE(PROD_GROUP,'A',2,'B',4,'C',6,8),    PROD_GROUP,          " & vbCrLf
    sQuery = sQuery & "         APLY_STDSPEC, SUM(PROD_WGT)                                     " & vbCrLf
    sQuery = sQuery & "   FROM  GP_RPT_DAILY_STDSPEC                                            " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE  <=  '" & txt_DATE.RawData & "'                       " & vbCrLf
    sQuery = sQuery & "    AND  PROD_DATE  >=  SUBSTR('" & txt_DATE.RawData & "',1,6)" & "||'00'" & vbCrLf
    sQuery = sQuery & "    AND  PROD_CD     =  'HC'                                             " & vbCrLf
    sQuery = sQuery & "    AND  PROD_WGT    >   0                                               " & vbCrLf
    sQuery = sQuery & "  GROUP  BY APLY_STDSPEC, PROD_GROUP                                     " & vbCrLf
    sQuery = sQuery & "  ORDER  BY 4 DESC, PROD_GROUP                                           " & vbCrLf

    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    IDc = iPPCol + 1
    iCol = 1
    Do Until AdoRs.EOF
        
        With ss4
            sgroup = Trim(AdoRs.Fields(1) & "")
            sStlgrd = Trim(AdoRs.Fields(2) & "")
            sStlgrdName = Trim(AdoRs.Fields(2) & "")
            lWgt = Val(AdoRs.Fields(3) & "")
            
            If iCol = iHCCol Then
               Exit Do
            End If
            
            For iCol = iPPCol + 1 To iHCCol
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
                .Text = Val(AdoRs.Fields(3) & "")
            End If
            
            sQuery = "SELECT            DECODE(PROD_GROUP,'A',1,'B',3,'C',5,7) , " & vbCrLf
            sQuery = sQuery & "         PROD_WGT                                 " & vbCrLf
            sQuery = sQuery & "   FROM  GP_RPT_DAILY_STDSPEC                     " & vbCrLf
            sQuery = sQuery & "  WHERE  PROD_DATE   =  '" & txt_DATE.RawData & "'" & vbCrLf
            sQuery = sQuery & "    AND  APLY_STDSPEC      =  '" & sStlgrdName & "'     " & vbCrLf
            sQuery = sQuery & "    AND  PROD_GROUP  =  '" & sgroup & "'          " & vbCrLf
            sQuery = sQuery & "    AND  PROD_CD     =  'HC'                      " & vbCrLf
            sQuery = sQuery & "    AND  PROD_WGT    >   0                        " & vbCrLf
        
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
    
    sQuery = "SELECT            COMMENTS1                                                       " & vbCrLf
    sQuery = sQuery & "   FROM  GP_RPT_DAILY_SUM                                                " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE             =  '" & txt_DATE.RawData & "'             " & vbCrLf
    sQuery = sQuery & "    AND  PROD_GROUP              = 'A'                                   " & vbCrLf
                                           
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.EOF Then
          ss4.Col = 1: ss4.Row = 11: ss4.Text = Trim(AdoRs.Fields(0) & "")
    End If
    AdoRs.Close
   
'    For iCol = iPPCol + 1 To iHCCol
'        ss4.ROW = 0
'        ss4.Col = iCol
'        If Trim(ss4.Text) = "" Then Call Gp_Sp_ColHidden(sc4.Item("Spread"), iCol, True)
'    Next iCol
    
    Screen.MousePointer = vbDefault
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Sub

Ss4_Display_Error:
    
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Ss4_Display_Error : " & Error)
    
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
    
    For iRow = 3 To ss3.MaxRows
        ss3.Row = iRow
        For iCol = 2 To ss3.MaxCols
            ss3.Col = iCol
            If ss3.CellType = SS_CELL_TYPE_NUMBER Then
                If Val(ss3.Text & "") = 0 Then
                    ss3.Text = ""
                End If
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
    
    sQuery = "{call AGC2640P ('" + Trim(Format(txt_DATE.Text, "YYYYMMDD")) + "','" + Trim(CBO_PLT.Text) + "',?)}"

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
    Dim I               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDate           As String
    
    Dim sExlRange       As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AKP3052C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    xlApp.Range("A2").Value = "报表日期：" + Left(sDate, 4) + "年" + Mid(sDate, 5, 2) + "月" + Mid(sDate, 7, 2) + "日"
    xlApp.Range("B64").Value = "制表日期：" + Format(Now, "YYYY-MM-DD HH:MM:SS")
    xlApp.Range("J64").Value = "制表人：" + sUserID

    Clipboard.Clear
    ss1.SetSelection 1, 1, ss1.MaxCols, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("C5").Select
    xlApp.ActiveSheet.Paste

    Clipboard.Clear
'    ss2.SetSelection 1, 1, ss2.MaxCols, ss2.MaxRows
'    ss2.ClipboardCopy
'    xlApp.Range("C17").Select
'    xlApp.ActiveSheet.Paste

''20140331
    ss2.Col = 1
    ss2.Row = 1: xlApp.Range("C17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("C18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("C19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("C20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("C21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("C22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("C23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("C24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("C25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("C26").Value = ss2.Text
    
    ss2.Col = 2
    ss2.Row = 1: xlApp.Range("D17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("D18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("D19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("D20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("D21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("D22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("D23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("D24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("D25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("D26").Value = ss2.Text
    
    ss2.Col = 3
    ss2.Row = 1: xlApp.Range("E17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("E18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("E19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("E20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("E21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("E22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("E23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("E24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("E25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("E26").Value = ss2.Text
    
    ss2.Col = 4
    ss2.Row = 1: xlApp.Range("F17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("F18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("F19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("F20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("F21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("F22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("F23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("F24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("F25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("F26").Value = ss2.Text
    
    ss2.Col = 5
    ss2.Row = 1: xlApp.Range("G17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("G18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("G19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("G20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("G21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("G22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("G23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("G24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("G25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("G26").Value = ss2.Text
    
    
'    ---------------------------------------------------热处理车间生产情况

    
    
    ss2.Col = 12
    ss2.Row = 1: xlApp.Range("H17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("H18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("H19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("H20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("H21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("H22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("H23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("H24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("H25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("H26").Value = ss2.Text
    
    ss2.Col = 13
    ss2.Row = 1: xlApp.Range("I17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("I18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("I19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("I20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("I21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("I22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("I23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("I24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("I25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("I26").Value = ss2.Text
    
    ss2.Col = 14
    ss2.Row = 1: xlApp.Range("J17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("J18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("J19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("J20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("J21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("J22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("J23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("J24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("J25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("J26").Value = ss2.Text
    
    ss2.Col = 15
    ss2.Row = 1: xlApp.Range("K17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("K18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("K19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("K20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("K21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("K22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("K23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("K24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("K25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("K26").Value = ss2.Text
    
    ss2.Col = 16
    ss2.Row = 1: xlApp.Range("L17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("L18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("L19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("L20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("L21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("L22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("L23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("L24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("L25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("L26").Value = ss2.Text
    
    ss2.Col = 17
    ss2.Row = 1: xlApp.Range("M17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("M18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("M19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("M20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("M21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("M22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("M23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("M24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("M25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("M26").Value = ss2.Text
    
    ss2.Col = 18
    ss2.Row = 1: xlApp.Range("N17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("N18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("N19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("N20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("N21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("N22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("N23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("N24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("N25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("N26").Value = ss2.Text
    
    ss2.Col = 19
    ss2.Row = 1: xlApp.Range("O17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("O18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("O19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("O20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("O21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("O22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("O23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("O24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("O25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("O26").Value = ss2.Text
    
'    ---------------------------------------------------热处理车间生产情况


    ss2.Col = 20
    ss2.Row = 1: xlApp.Range("P17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("P18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("P19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("P20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("P21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("P22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("P23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("P24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("P25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("P26").Value = ss2.Text
    
    ss2.Col = 21
    ss2.Row = 1: xlApp.Range("Q17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("Q18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("Q19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("Q20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("Q21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("Q22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("Q23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("Q24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("Q25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("Q26").Value = ss2.Text

'    ---------------------------------------------------轧钢故障停时（分钟）开始

   ss2.Col = 22
    ss2.Row = 1: xlApp.Range("R17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("R18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("R19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("R20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("R21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("R22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("R23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("R24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("R25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("R26").Value = ss2.Text
    
    ss2.Col = 23
    ss2.Row = 1: xlApp.Range("S17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("S18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("S19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("S20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("S21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("S22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("S23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("S24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("S25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("S26").Value = ss2.Text

   ss2.Col = 24
    ss2.Row = 1: xlApp.Range("T17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("T18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("T19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("T20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("T21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("T22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("T23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("T24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("T25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("T26").Value = ss2.Text
    
    ss2.Col = 25
    ss2.Row = 1: xlApp.Range("U17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("U18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("U19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("U20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("U21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("U22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("U23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("U24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("U25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("U26").Value = ss2.Text

   ss2.Col = 26
    ss2.Row = 1: xlApp.Range("V17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("V18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("V19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("V20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("V21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("V22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("V23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("V24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("V25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("V26").Value = ss2.Text
    
    ss2.Col = 27
    ss2.Row = 1: xlApp.Range("W17").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("W18").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("W19").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("W20").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("W21").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("W22").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("W23").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("W24").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("W25").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("W26").Value = ss2.Text

'   ss2.Col = 28
'    ss2.Row = 1: xlApp.Range("X17").Value = ss2.Text
'    ss2.Row = 2: xlApp.Range("X18").Value = ss2.Text
'    ss2.Row = 3: xlApp.Range("X19").Value = ss2.Text
'    ss2.Row = 4: xlApp.Range("X20").Value = ss2.Text
'    ss2.Row = 5: xlApp.Range("X21").Value = ss2.Text
'    ss2.Row = 6: xlApp.Range("X22").Value = ss2.Text
'    ss2.Row = 7: xlApp.Range("X23").Value = ss2.Text
'    ss2.Row = 8: xlApp.Range("X24").Value = ss2.Text
'    ss2.Row = 9: xlApp.Range("X25").Value = ss2.Text
'    ss2.Row = 10: xlApp.Range("X26").Value = ss2.Text
'
'    ss2.Col = 29
'    ss2.Row = 1: xlApp.Range("Y17").Value = ss2.Text
'    ss2.Row = 2: xlApp.Range("Y18").Value = ss2.Text
'    ss2.Row = 3: xlApp.Range("Y19").Value = ss2.Text
'    ss2.Row = 4: xlApp.Range("Y20").Value = ss2.Text
'    ss2.Row = 5: xlApp.Range("Y21").Value = ss2.Text
'    ss2.Row = 6: xlApp.Range("Y22").Value = ss2.Text
'    ss2.Row = 7: xlApp.Range("Y23").Value = ss2.Text
'    ss2.Row = 8: xlApp.Range("Y24").Value = ss2.Text
'    ss2.Row = 9: xlApp.Range("Y25").Value = ss2.Text
'    ss2.Row = 10: xlApp.Range("Y26").Value = ss2.Text
'
'   ss2.Col = 30
'    ss2.Row = 1: xlApp.Range("Z17").Value = ss2.Text
'    ss2.Row = 2: xlApp.Range("Z18").Value = ss2.Text
'    ss2.Row = 3: xlApp.Range("Z19").Value = ss2.Text
'    ss2.Row = 4: xlApp.Range("Z20").Value = ss2.Text
'    ss2.Row = 5: xlApp.Range("Z21").Value = ss2.Text
'    ss2.Row = 6: xlApp.Range("Z22").Value = ss2.Text
'    ss2.Row = 7: xlApp.Range("Z23").Value = ss2.Text
'    ss2.Row = 8: xlApp.Range("Z24").Value = ss2.Text
'    ss2.Row = 9: xlApp.Range("Z25").Value = ss2.Text
'    ss2.Row = 10: xlApp.Range("Z26").Value = ss2.Text
    
'    ---------------------------------------------------轧钢故障停时（分钟）开始
 
    ss2.Col = 6
    ss2.Row = 1: xlApp.Range("G53").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("G54").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("G55").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("G56").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("G57").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("G58").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("G59").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("G60").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("G61").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("G62").Value = ss2.Text

  ss2.Col = 7
    ss2.Row = 1: xlApp.Range("H53").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("H54").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("H55").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("H56").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("H57").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("H58").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("H59").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("H60").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("H61").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("H62").Value = ss2.Text
    
  ss2.Col = 8
    ss2.Row = 1: xlApp.Range("I53").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("I54").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("I55").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("I56").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("I57").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("I58").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("I59").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("I60").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("I61").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("I62").Value = ss2.Text
    
   ss2.Col = 9
    ss2.Row = 1: xlApp.Range("J53").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("J54").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("J55").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("J56").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("J57").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("J58").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("J59").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("J60").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("J61").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("J62").Value = ss2.Text
    
   ss2.Col = 10
    ss2.Row = 1: xlApp.Range("K53").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("K54").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("K55").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("K56").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("K57").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("K58").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("K59").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("K60").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("K61").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("K62").Value = ss2.Text
    
   ss2.Col = 11
    ss2.Row = 1: xlApp.Range("L53").Value = ss2.Text
    ss2.Row = 2: xlApp.Range("L54").Value = ss2.Text
    ss2.Row = 3: xlApp.Range("L55").Value = ss2.Text
    ss2.Row = 4: xlApp.Range("L56").Value = ss2.Text
    ss2.Row = 5: xlApp.Range("L57").Value = ss2.Text
    ss2.Row = 6: xlApp.Range("L58").Value = ss2.Text
    ss2.Row = 7: xlApp.Range("L59").Value = ss2.Text
    ss2.Row = 8: xlApp.Range("L60").Value = ss2.Text
    ss2.Row = 9: xlApp.Range("L61").Value = ss2.Text
    ss2.Row = 10: xlApp.Range("L62").Value = ss2.Text

''20140331

    ss3.Col = 2
    ss3.Row = 3:  xlApp.Range("Z53").Value = ss3.Text
    ss3.Row = 4:  xlApp.Range("Z54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("Z55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("Z56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("Z57").Value = ss3.Text
    ss3.Row = 8:  xlApp.Range("Z58").Value = ss3.Text
    ss3.Row = 9:  xlApp.Range("Z59").Value = ss3.Text
    ss3.Row = 10:  xlApp.Range("Z60").Value = ss3.Text
    ss3.Row = 11:  xlApp.Range("Z61").Value = ss3.Text
    ss3.Row = 12:  xlApp.Range("Z62").Value = ss3.Text
    
    ss3.Col = 4
    ss3.Row = 3:  xlApp.Range("AB53").Value = ss3.Text
    ss3.Row = 4:  xlApp.Range("AB54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("AB55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("AB56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("AB57").Value = ss3.Text
    ss3.Row = 8:  xlApp.Range("AB58").Value = ss3.Text
    ss3.Row = 9:  xlApp.Range("AB59").Value = ss3.Text
    ss3.Row = 10:  xlApp.Range("AB60").Value = ss3.Text
    ss3.Row = 11:  xlApp.Range("AB61").Value = ss3.Text
    ss3.Row = 12:  xlApp.Range("AB62").Value = ss3.Text
    
    ss3.Col = 6
    ss3.Row = 4:  xlApp.Range("W54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("W55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("W56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("W57").Value = ss3.Text
    
    ss3.Col = 7
    ss3.Row = 4:  xlApp.Range("X54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("X55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("X56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("X57").Value = ss3.Text
    ss3.Row = 8:  xlApp.Range("X58").Value = ss3.Text
    ss3.Row = 9:  xlApp.Range("X59").Value = ss3.Text
    ss3.Row = 10:  xlApp.Range("X60").Value = ss3.Text
    ss3.Row = 11:  xlApp.Range("X61").Value = ss3.Text
    ss3.Row = 12:  xlApp.Range("X62").Value = ss3.Text
    
    ss3.Col = 9
    ss3.Row = 4:  xlApp.Range("Q54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("Q55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("Q56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("Q57").Value = ss3.Text
    ss3.Row = 8:  xlApp.Range("Q58").Value = ss3.Text
    ss3.Row = 9:  xlApp.Range("Q59").Value = ss3.Text
    ss3.Row = 10:  xlApp.Range("Q60").Value = ss3.Text
    ss3.Row = 11:  xlApp.Range("Q61").Value = ss3.Text
    ss3.Row = 12:  xlApp.Range("Q62").Value = ss3.Text
    
    ss3.Col = 10
    ss3.Row = 4:  xlApp.Range("R54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("R55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("R56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("R57").Value = ss3.Text
    ss3.Row = 8:  xlApp.Range("R58").Value = ss3.Text
    ss3.Row = 9:  xlApp.Range("R59").Value = ss3.Text
    ss3.Row = 10:  xlApp.Range("R60").Value = ss3.Text
    ss3.Row = 11:  xlApp.Range("R61").Value = ss3.Text
    ss3.Row = 12:  xlApp.Range("R62").Value = ss3.Text
    
    ss3.Col = 12
    ss3.Row = 4:  xlApp.Range("T54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("T55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("T56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("T57").Value = ss3.Text
    ss3.Row = 8:  xlApp.Range("T58").Value = ss3.Text
    ss3.Row = 9:  xlApp.Range("T59").Value = ss3.Text
    ss3.Row = 10:  xlApp.Range("T60").Value = ss3.Text
    ss3.Row = 11:  xlApp.Range("T61").Value = ss3.Text
    ss3.Row = 12:  xlApp.Range("T62").Value = ss3.Text
    
    ss3.Col = 13
    ss3.Row = 4:  xlApp.Range("U54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("U55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("U56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("U57").Value = ss3.Text
    ss3.Row = 8:  xlApp.Range("U58").Value = ss3.Text
    ss3.Row = 9:  xlApp.Range("U59").Value = ss3.Text
    ss3.Row = 10:  xlApp.Range("U60").Value = ss3.Text
    ss3.Row = 11:  xlApp.Range("U61").Value = ss3.Text
    ss3.Row = 12:  xlApp.Range("U62").Value = ss3.Text
    
'    ....................................20140416
    ss3.Col = 15
    ss3.Row = 4:  xlApp.Range("N54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("N55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("N56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("N57").Value = ss3.Text
    ss3.Row = 8:  xlApp.Range("N58").Value = ss3.Text
    ss3.Row = 9:  xlApp.Range("N59").Value = ss3.Text
    ss3.Row = 10:  xlApp.Range("N60").Value = ss3.Text

    
    ss3.Col = 16
    ss3.Row = 4:  xlApp.Range("O54").Value = ss3.Text
    ss3.Row = 5:  xlApp.Range("O55").Value = ss3.Text
    ss3.Row = 6:  xlApp.Range("O56").Value = ss3.Text
    ss3.Row = 7:  xlApp.Range("O57").Value = ss3.Text
    ss3.Row = 8:  xlApp.Range("O58").Value = ss3.Text
    ss3.Row = 9:  xlApp.Range("O59").Value = ss3.Text
    ss3.Row = 10:  xlApp.Range("O60").Value = ss3.Text
'    ....................................20140416
    
    
    
    For I = 1 To 26
        ss4.Col = I
        ss4.Row = 0
        sExlRange = Chr(I + 66) & 27
        If sExlRange = "[27" Then sExlRange = "AA27"
        If sExlRange = "\27" Then sExlRange = "AB27"
        xlApp.Range(sExlRange).Value = ss4.Text
    Next I
    
    For I = 27 To 52
        ss4.Col = I
        ss4.Row = 0
        sExlRange = Chr(I - 26 + 66) & 39
        If sExlRange = "[39" Then sExlRange = "AA39"
        If sExlRange = "\39" Then sExlRange = "AB39"
        xlApp.Range(sExlRange).Value = ss4.Text
    Next I
    
'    For I = 53 To 65
'        ss4.Col = I
'        ss4.Row = 0
'        sExlRange = Chr(I - 52 + 66) & 51
'        If sExlRange = "[51" Then sExlRange = "AA51"
'        If sExlRange = "\51" Then sExlRange = "AB51"
'        xlApp.Range(sExlRange).Value = ss4.Text
'    Next I

'20140416
     For I = 53 To 56
        ss4.Col = I
        ss4.Row = 0
        sExlRange = Chr(I - 52 + 66) & 51
        If sExlRange = "[51" Then sExlRange = "AA51"
        If sExlRange = "\51" Then sExlRange = "AB51"
        xlApp.Range(sExlRange).Value = ss4.Text
    Next I
'20140416
    Clipboard.Clear
    ss4.SetSelection 1, 1, 26, 10
    ss4.ClipboardCopy
    xlApp.Range("C29").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss4.SetSelection 27, 1, 52, 10
    ss4.ClipboardCopy
    xlApp.Range("C41").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss4.SetSelection 53, 1, 65, 10
    ss4.ClipboardCopy
    xlApp.Range("C53").Select
    xlApp.ActiveSheet.Paste

    Clipboard.Clear
    ss4.SetSelection 1, 11, 65, 11
    ss4.ClipboardCopy
    xlApp.Range("C63").Select
    xlApp.ActiveSheet.Paste

    ss1.ClearSelection
    ss2.ClearSelection
    ss3.ClearSelection
    ss4.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
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
' 115,80,195
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
        
        
        If MsgChk Then
            .LockBackColor = RGB(255, 255, 255)
        End If

    End With
    
End Sub
Public Function Mill_Sp_Display(Conn As ADODB.Connection, sPname As Variant, sQuery As String) As Boolean

    On Error Resume Next

    Dim iCount          As Integer
    Dim iRowCount       As Long
    Dim iColcount       As Long
    Dim AdoRs           As ADODB.Recordset
    Dim ArrayRecords    As Variant

    Mill_Sp_Display = True

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Mill_Sp_Display = False: Exit Function
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
            Mill_Sp_Display = False
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
            
                .Row = iRowCount + 1

                For iColcount = 1 To .MaxCols
    
                    .Col = iColcount
    
                    If VarType(ArrayRecords(iColcount - 1, iRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iColcount - 1, iRowCount))
                    End If

                Next iColcount

            Next iRowCount

        End If

        .ReDraw = True
        Screen.MousePointer = vbDefault

    End With

End Function
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
    
    Dim dNoPlanWgt_d   As Double
    Dim dNoPlanWgt_m   As Double

    Dim AdoRs As ADODB.Recordset

    Set AdoRs = New ADODB.Recordset
  
    sQuery = "SELECT            *                                       " & vbCrLf
    sQuery = sQuery & "   FROM  GP_RPT_COMMENTS                         " & vbCrLf
    sQuery = sQuery & "  WHERE  PROD_DATE  =  '" & txt_DATE.RawData & "'" & vbCrLf
    sQuery = sQuery & "    AND  PLT        =  '" & CBO_PLT.Text & "'" & vbCrLf

    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Do Until AdoRs.EOF
       
        With ss3
    
            .Col = 2:   .Row = 3:    .Text = Val(AdoRs.Fields(8) & "")
                        .Row = 4:    .Text = Val(AdoRs.Fields(9) & "")
                        .Row = 5:    .Text = Val(AdoRs.Fields(20) & "")
                        .Row = 6:    .Text = Val(AdoRs.Fields(21) & "")
                        .Row = 7:    .Text = Val(AdoRs.Fields(10) & "")
                        .Row = 8:    .Text = Val(AdoRs.Fields(11) & "")
                        .Row = 9:    .Text = Val(AdoRs.Fields(12) & "")
                        .Row = 10:   .Text = Val(AdoRs.Fields(13) & "")
                        .Row = 11:   .Text = Val(AdoRs.Fields(22) & "")
                        .Row = 12:   .Text = Val(AdoRs.Fields(23) & "")
            .Col = 4:   .Row = 3:    .Text = Val(AdoRs.Fields(2) & "")
                        .Row = 4:    .Text = Val(AdoRs.Fields(3) & "")
                        .Row = 5:    .Text = Val(AdoRs.Fields(18) & "")
                        .Row = 6:    .Text = Val(AdoRs.Fields(19) & "")
                        .Row = 7:    .Text = Val(AdoRs.Fields(4) & "")
                        .Row = 8:    .Text = Val(AdoRs.Fields(5) & "")
                        .Row = 9:    .Text = Val(AdoRs.Fields(6) & "")
                        .Row = 10:   .Text = Val(AdoRs.Fields(7) & "")
                        .Row = 11:   .Text = Val(AdoRs.Fields(14) & "")
                        .Row = 12:   .Text = Val(AdoRs.Fields(15) & "")
            .Col = 6:   .Row = 4:    .Text = Val(AdoRs.Fields(56) & ""):    dNoPlanWgt_d = .Value
                        .Row = 5:    .Text = Val(AdoRs.Fields(58) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 6:    .Text = Val(AdoRs.Fields(54) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 7:    .Text = dNoPlanWgt_d
            .Col = 7:   .Row = 4:    .Text = Val(AdoRs.Fields(57) & ""):    dNoPlanWgt_m = .Value
                        .Row = 5:    .Text = Val(AdoRs.Fields(59) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 6:    .Text = Val(AdoRs.Fields(55) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 7:    .Text = dNoPlanWgt_m
                        .Row = 8:    .Text = Val(AdoRs.Fields(73) & ""):
                        .Row = 9:    .Text = Val(AdoRs.Fields(74) & ""):
                        .Row = 10:   .Text = Val(AdoRs.Fields(75) & ""):
                        .Row = 11:   .Text = Val(AdoRs.Fields(77) & ""):
                        .Row = 12:   .Text = Val(AdoRs.Fields(76) & ""):
            .Col = 9:   .Row = 4:    .Text = Val(AdoRs.Fields(24) & ""):    dNoPlanWgt_d = .Value
                        .Row = 5:    .Text = Val(AdoRs.Fields(26) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 6:    .Text = Val(AdoRs.Fields(28) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 7:    .Text = Val(AdoRs.Fields(30) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 8:    .Text = Val(AdoRs.Fields(32) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 9:    .Text = Val(AdoRs.Fields(34) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 10:   .Text = Val(AdoRs.Fields(36) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 11:   .Text = Val(AdoRs.Fields(38) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 12:   .Text = Val(AdoRs.Fields(40) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
            .Col = 10:  .Row = 4:    .Text = Val(AdoRs.Fields(25) & ""):    dNoPlanWgt_m = .Value
                        .Row = 5:    .Text = Val(AdoRs.Fields(27) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 6:    .Text = Val(AdoRs.Fields(29) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 7:    .Text = Val(AdoRs.Fields(31) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 8:    .Text = Val(AdoRs.Fields(33) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 9:    .Text = Val(AdoRs.Fields(35) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 10:   .Text = Val(AdoRs.Fields(37) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 11:   .Text = Val(AdoRs.Fields(39) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 12:   .Text = Val(AdoRs.Fields(41) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
            .Col = 12:  .Row = 4:    .Text = Val(AdoRs.Fields(42) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 5:    .Text = Val(AdoRs.Fields(44) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 6:    .Text = Val(AdoRs.Fields(46) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 7:    .Text = Val(AdoRs.Fields(48) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 8:    .Text = Val(AdoRs.Fields(50) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 9:    .Text = Val(AdoRs.Fields(52) & ""):    dNoPlanWgt_d = dNoPlanWgt_d + .Value
                        .Row = 10:   .Text = dNoPlanWgt_d
                        .Row = 11:    .Text = Val(AdoRs.Fields(70) & ""):
            .Col = 13:  .Row = 4:    .Text = Val(AdoRs.Fields(43) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 5:    .Text = Val(AdoRs.Fields(45) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 6:    .Text = Val(AdoRs.Fields(47) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 7:    .Text = Val(AdoRs.Fields(49) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 8:    .Text = Val(AdoRs.Fields(51) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 9:    .Text = Val(AdoRs.Fields(53) & ""):    dNoPlanWgt_m = dNoPlanWgt_m + .Value
                        .Row = 10:   .Text = dNoPlanWgt_m
                        .Row = 11:   .Text = Val(AdoRs.Fields(71) & ""):
                        .Row = 12:   .Text = Val(AdoRs.Fields(72) & ""):
                        
                        
'                        -----------------------------------------20140418
                        
            .Col = 15:  .Row = 4:    .Text = Val(AdoRs.Fields(78) & ""):
                        .Row = 5:    .Text = Val(AdoRs.Fields(80) & ""):
                        .Row = 6:    .Text = Val(AdoRs.Fields(82) & ""):
                        .Row = 7:    .Text = Val(AdoRs.Fields(84) & ""):
                        .Row = 8:    .Text = Val(AdoRs.Fields(86) & ""):
                        .Row = 9:    .Text = Val(AdoRs.Fields(88) & ""):
                        .Row = 10:   .Text = Val(AdoRs.Fields(90) & ""):
                      
            .Col = 15:  .Row = 4:    .Text = Val(AdoRs.Fields(79) & ""):
                        .Row = 5:    .Text = Val(AdoRs.Fields(81) & ""):
                        .Row = 6:    .Text = Val(AdoRs.Fields(83) & ""):
                        .Row = 7:    .Text = Val(AdoRs.Fields(85) & ""):
                        .Row = 8:    .Text = Val(AdoRs.Fields(87) & ""):
                        .Row = 9:    .Text = Val(AdoRs.Fields(89) & ""):
                        .Row = 10:   .Text = Val(AdoRs.Fields(91) & ""):
                
                        
'                  -----------------------------------------20140418
                        
                        


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
Public Function ss3_clear() As Boolean
    Dim I As Integer
    
    With ss3
        .Col = 2
        For I = 3 To 12
            .Row = I
            .Text = ""
        Next I
        .Col = 4
        For I = 3 To 12
            .Row = I: .Text = ""
        Next I

    End With
    
End Function




