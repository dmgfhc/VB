VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2050C 
   Caption         =   "火切实绩查询及修改界面_AGC2050C"
   ClientHeight    =   9450
   ClientLeft      =   690
   ClientTop       =   2325
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SP1 
      Height          =   8190
      Left            =   30
      TabIndex        =   0
      Top             =   1020
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   14446
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "AGC2050C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   3600
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   15030
         _Version        =   393216
         _ExtentX        =   26511
         _ExtentY        =   6350
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   31
         MaxRows         =   20
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGC2050C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4500
         Left            =   0
         TabIndex        =   4
         Top             =   3690
         Width           =   15030
         _Version        =   393216
         _ExtentX        =   26511
         _ExtentY        =   7938
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGC2050C.frx":4619
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   450
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   1058
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GulimChe"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox CBO_AREA 
         Height          =   315
         ItemData        =   "AGC2050C.frx":647E
         Left            =   11490
         List            =   "AGC2050C.frx":6480
         TabIndex        =   17
         Top             =   150
         Width           =   825
      End
      Begin VB.TextBox txt_Loc 
         CausesValidation=   0   'False
         Height          =   310
         Left            =   7620
         MaxLength       =   7
         TabIndex        =   15
         Tag             =   "生产工厂"
         Top             =   150
         Width           =   1770
      End
      Begin VB.TextBox TXT_INQNO 
         Height          =   330
         Left            =   4680
         MaxLength       =   14
         TabIndex        =   12
         Tag             =   "材料号"
         Top             =   150
         Width           =   1890
      End
      Begin VB.TextBox txt_plt_name 
         CausesValidation=   0   'False
         Height          =   310
         Left            =   1890
         TabIndex        =   6
         Tag             =   "机号"
         Top             =   150
         Width           =   1500
      End
      Begin VB.TextBox txt_plt 
         CausesValidation=   0   'False
         Height          =   310
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "生产工厂"
         Top             =   150
         Width           =   450
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   3450
         Top             =   150
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "材料号"
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
         Left            =   6690
         Top             =   150
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   556
         Caption         =   "货位"
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
         Height          =   300
         Left            =   240
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Caption         =   "生产工厂"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   10380
         Top             =   150
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         Caption         =   "切割区域"
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
      Begin VB.Label lbl_moplate_wgt 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   13740
         TabIndex        =   2
         Top             =   225
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   480
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   847
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GulimChe"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TXT_MILL_LOT_NO 
         Height          =   330
         Left            =   7875
         MaxLength       =   14
         TabIndex        =   16
         Top             =   90
         Width           =   1830
      End
      Begin VB.TextBox TXT_MPLATE_NO 
         Height          =   330
         Left            =   13125
         MaxLength       =   14
         TabIndex        =   14
         Tag             =   "材料号"
         Top             =   90
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.TextBox TXT_PRODCD 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   12900
         TabIndex        =   13
         Tag             =   "产品代码"
         Top             =   90
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txt_PrcLine 
         CausesValidation=   0   'False
         Height          =   310
         Left            =   5160
         TabIndex        =   11
         Tag             =   "产线分类"
         Text            =   "3"
         Top             =   90
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txt_WkPlt 
         CausesValidation=   0   'False
         Height          =   310
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Tag             =   "作业工厂"
         Text            =   "C1"
         Top             =   90
         Width           =   420
      End
      Begin VB.ComboBox cbo_PrcLine 
         Height          =   315
         ItemData        =   "AGC2050C.frx":6482
         Left            =   3180
         List            =   "AGC2050C.frx":6484
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   90
         Width           =   1935
      End
      Begin InDate.ULabel ULabel10 
         Height          =   300
         Left            =   240
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         Caption         =   "作业工厂"
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
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   2010
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         Caption         =   "产线分类"
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
         Left            =   6660
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         Caption         =   "轧批号"
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
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   13740
         TabIndex        =   8
         Top             =   225
         Visible         =   0   'False
         Width           =   885
      End
   End
End
Attribute VB_Name = "AGC2050C"
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
'-- Program Name      钢板剪切实绩查询及修改界面
'-- Program ID        CGD2035C
'-- Document No       Q-00-0010(Specification)
'-- Designer          SHIN.C.S
'-- Coder             SHIN.C.S
'-- Date              2007.8.13
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
Public sQuery_Rt As String          'Active Form sQuery Setting

Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

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



Dim Mc1 As New Collection           'Master Collectionn
Dim Mc2 As New Collection           'Master Collectionn

Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_PLATE_NO = 1
Const SPD_DS_CUT_END_DATE = 12
Const SPD_THK = 14
Const SPD_WID = 15
Const SPD_LEN = 16
Const SPD_WGT = 17
Const SPD_DS_LAST_YN = 18
Const SPD_SURF_GRD = 16
Const SPD_TRIM_FL = 17
Const SPD_DS_KNIFE_GAP = 18
Const SPD_EMP_CD = 21


Const SPD_PROC_CD = 23
Const SPD_END_USE = 24
Const SPD_STLGRD = 25

Const SS2_URGNT_FL = 19      'Add by LiQian at 2012-11-09 紧急订单标记

Dim sQuery   As String
Dim sLoopFl  As String

Dim Screen_Fl As Boolean

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"
       
    'MASTER Collection
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_WkPlt, "p", " ", " ", " ", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(txt_PrcLine, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(TXT_PLT, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(TXT_INQNO, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(TXT_MILL_LOT_NO, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(TXT_LOC, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          
    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=cControl1, Key:="cControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"
    
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
   
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGC2050C.P_REFER1", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_mplate_no, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(TXT_PRODCD, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
        
    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", "n", "m", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2050C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AGC2050C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 26, True)
    Call Gp_Sp_ColHidden(ss1, 27, True)
        
    Screen_Fl = False
     
End Sub


Private Sub cbo_PrcLine_Change()
       If cbo_PrcLine.Text = "一号线" Then
          txt_PrcLine = "3"
       ElseIf cbo_PrcLine.Text = "二号线" Then
          txt_PrcLine = "4"
       End If
End Sub

Private Sub cbo_PrcLine_Click()
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    txt_mplate_no = ""
       
    If cbo_PrcLine.Text = "一号线" Then
       txt_PrcLine = "3"
    ElseIf cbo_PrcLine.Text = "二号线" Then
       txt_PrcLine = "4"
    End If
    
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

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))

    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    ss1.RowHeight(-1) = 13.5

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "CG-System.INI", Me.Name)
    
    If App.Title = "BG" Then
        txt_WkPlt = "C1"
        cbo_PrcLine.AddItem "一号线"
        cbo_PrcLine.ListIndex = 0
    ElseIf App.Title = "DG" Then
        txt_WkPlt = "C1"
        cbo_PrcLine.AddItem "一号线"
        cbo_PrcLine.ListIndex = 0
    ElseIf App.Title = "CG" Then
        txt_WkPlt = "C3"
        cbo_PrcLine.AddItem "一号线"
        cbo_PrcLine.ListIndex = 0
    End If
    
    CBO_AREA.AddItem "L"
    CBO_AREA.AddItem "T"
    CBO_AREA.AddItem "B"
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "CG-System.INI", Me.Name)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If


    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing

    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
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
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    
    Set sc1 = Nothing
    Set sc2 = Nothing
    
    Set Proc_Sc = Nothing
    

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
        If Gf_Sp_Cls(sc2) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            Call Gp_Ms_ControlLock(Mc2("lControl"), False)
            
            lbl_moplate_wgt.Caption = ""
            
        End If
    End If
End Sub

Public Sub Form_Ref()
    
    On Error GoTo Refer_Err
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sUrgnt_Fl As String
    
    Dim iCount As Integer
    
    Call Gf_Sp_Cls(sc1)
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
    ss2.OperationMode = OperationModeNormal
    
     '紧急订单绿色显示 add by liqian 2012-11-09
     With ss2
          For iRow = 1 To .MaxRows
             .Row = iRow:
              .Col = SS2_URGNT_FL:    sUrgnt_Fl = Trim(.Text)
            
              If sUrgnt_Fl = "Y" Then
                 Call Gp_Sp_BlockColor(ss2, 1, .MaxCols, iRow, iRow, &HC000&)
              End If
          Next iRow
    End With
    
    Call ss2_DblClick(1, 1)
           
Refer_Err:
       
End Sub

Public Sub Form_Pro()

Dim iCount      As Integer
Dim START_FOR   As Integer
Dim sDateFrom   As String
Dim sDateTo     As String
Dim sPlateNo    As String
    
Dim inum As Integer
Dim lRow As Integer
    
    For iCount = 1 To ss1.MaxRows
        ss1.Row = iCount
'        ss1.Col = 0
'        ss1.Text = "Update"
        ss1.Col = 18
        If ss1.Value = 1 Then
            START_FOR = iCount
            Exit For
        End If
    Next

    If START_FOR < ss1.MaxRows Then
        START_FOR = START_FOR + 1
        For iCount = START_FOR To ss1.MaxRows
            ss1.Row = iCount
            ss1.Col = 0
            ss1.Text = ""
        Next
    End If
   
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    Call Form_Ref
    
End Sub

Public Sub Form_Ins()
Dim dThk        As Double
Dim dWid        As Double
Dim dLen        As Double
Dim dWgt        As Double
Dim lRow        As Long
Dim sPlateNo    As String
Dim sClipText   As String
Dim cCUTDATE As String
Dim cTRIM As String
Dim cHEAD As String
Dim cTAIL As String
Dim cSAMP1 As String
Dim cSAMP2 As String

Dim iIdr        As Integer
Dim iCount As Integer

    sPlateNo = ""

    With ss1
        If .MaxRows = 0 Then
           If Len(txt_mplate_no.Text) = 12 Then
               Call Gp_Sp_Ins(Proc_Sc("Sc"))
              .Row = 1
              .Col = 1
              .Text = txt_mplate_no.Text & "01"
              .Col = 27
              .Text = txt_WkPlt
              .Col = 28
              .Text = txt_PrcLine
              .Col = 26
              .Text = sUserID
           Else
               Call Gp_MsgBoxDisplay("请正确输入母板号 ！")
           End If
           Exit Sub
        End If
        For iCount = .ActiveRow To .MaxRows
            .Row = iCount
            .Col = 1
            If Left(.Text, 12) = Left(sPlateNo, 12) Or sPlateNo = "" Then
               sPlateNo = .Text
               lRow = iCount
            Else
               Exit For
            End If
        Next iCount
    End With

    sPlateNo = ""

    Call ss1.SetActiveCell(1, lRow)
    Call Gp_Sp_Ins(Proc_Sc("Sc"))

    With ss1
        .ReDraw = False
        .Row = .ActiveRow
        .Col = 27
        .Text = txt_WkPlt
        .Col = 28
        .Text = txt_PrcLine
        .Col = 26
        .Text = sUserID

        If lRow > 0 Then
            .Row = lRow
            .Col = SPD_PLATE_NO:      sPlateNo = .Text
            .Col = SPD_THK:           dThk = Val(.Text & "")
            .Col = SPD_WID:           dWid = Val(.Text & "")
            .Col = SPD_LEN:           dLen = Val(.Text & "")
            .Col = SPD_WGT:           dWgt = Val(.Text & "")
        Else
            sPlateNo = txt_mplate_no.Text & "00"
        End If

        .Row = lRow + 1
        .Col = SPD_PLATE_NO:      .Text = sPlateNo
        .Col = SPD_THK:           .Text = dThk
        .Col = SPD_WID:           .Text = dWid
        .Col = SPD_LEN:           .Text = dLen
        .Col = SPD_WGT:           .Text = dWgt
        .Col = 0: .Text = "Input"
        .Col = SPD_PLATE_NO: .Text = Format(Val(.Text & "") + 1, "00000000000000")

         Call .SetActiveCell(1, .Row)
        .ReDraw = True
    End With
    
    ss1.Row = 1
    ss1.Col = 13
    cCUTDATE = ss1.Text
    ss1.Col = 19
    cTRIM = ss1.Value
    ss1.Col = 20
    cHEAD = ss1.Value
    ss1.Col = 21
    cTAIL = ss1.Value
    ss1.Col = 22
    cSAMP1 = ss1.Value
    ss1.Col = 23
    cSAMP2 = ss1.Value

    For iIdr = 1 To ss1.MaxRows
        ss1.Row = iIdr
        
        ss1.Col = 13
        ss1.Text = cCUTDATE
        ss1.Col = 18
        ss1.Value = 0
        
        ss1.Col = 19
        ss1.Value = cTRIM
        ss1.Col = 20
        ss1.Value = cHEAD
        ss1.Col = 21
        ss1.Value = cTAIL
        ss1.Col = 22
        ss1.Value = cSAMP1
        ss1.Col = 23
        ss1.Value = cSAMP2
        
        ss1.Col = 24
        ss1.Text = Gf_ShiftSet3(M_CN1)
        ss1.Col = 25
        ss1.Text = Gf_GroupSet(M_CN1, Gf_ShiftSet3(M_CN1), Gf_DTSet(M_CN1, , "X"))
        ss1.Col = 26
        ss1.Text = sUserID
        ss1.Col = 27
        ss1.Text = txt_WkPlt
        ss1.Col = 28
        ss1.Text = txt_PrcLine
        If iIdr = ss1.MaxRows Then
           ss1.Row = iIdr
           ss1.Col = 18
           ss1.Value = 1
        End If

    Next iIdr


End Sub

Private Sub PlateWgtEdit()
    Dim dThk        As Double
    Dim dWid        As Double
    Dim dLen        As Double
    Dim dWgt        As Double
    Dim sProcCode   As Double
    Dim sEndUseCd   As String
    Dim sStlgrd     As String
    Dim iCount      As Integer
         
    lbl_moplate_wgt.Caption = ""
    With ss1
        For iCount = 1 To ss1.MaxRows
            .Row = iCount
            
            .Col = SPD_THK:  dThk = Val(.Text & "")
            .Col = SPD_WID:  dWid = Val(.Text & "")
            .Col = SPD_LEN:  dLen = Val(.Text & "")
            .Col = SPD_WGT:  dWgt = Val(.Text & "")
            lbl_moplate_wgt.Caption = Val(lbl_moplate_wgt.Caption & "") + Val(.Text & "")
            If dWgt = 0 And dThk > 0 And dWid > 0 And dLen > 0 Then
                .Col = SPD_WGT
                .Text = Cal_Plate_Wgt("WGT", sEndUseCd, sStlgrd, dThk, dWid, dLen)
            End If
        Next iCount
    End With
End Sub

Private Function Cal_Plate_Wgt(sMode As String, sEndUseCd As String, sStlgrd As String, _
                                dThk As Double, dWid As Double, dLen As Double) As Double

    Dim RS  As New ADODB.Recordset
    
    Cal_Plate_Wgt = 0
    
    sQuery = "SELECT  Gf_Cal_Plate_Wgt('" & sMode & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & sEndUseCd & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & sStlgrd & "'" & vbCrLf
    sQuery = sQuery & "             ," & dThk & vbCrLf
    sQuery = sQuery & "             ," & dWid & vbCrLf
    sQuery = sQuery & "             ," & dLen & vbCrLf
    sQuery = sQuery & "             ,0 )" & vbCrLf
    sQuery = sQuery & "       FROM  DUAL " & vbCrLf
    RS.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        Cal_Plate_Wgt = Val(RS(0).Value & "")
    End If
    
    RS.Close
    Set RS = Nothing
     
End Function

Public Sub Spread_Can()
    ss1.Col = 0
    ss1.Row = ss1.ActiveRow
    Select Case Trim(ss1.Text)
        Case "Input"
              Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
        Case Else
             ss1.Text = ""
    End Select
End Sub
Public Sub Spread_Del()

    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Public Sub Spread_Cpy()
    Call Gp_Sp_Copy(Proc_Sc("Sc"))
End Sub

Public Sub Spread_Pst()
    Call Gp_Sp_Paste(Proc_Sc("Sc"))
End Sub


Public Sub Spread_Forzens_Setting()
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
End Sub

Public Sub Spread_Forzens_Cancel()
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
End Sub




Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub


Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim FOR_CNT
Dim START_FOR As Integer
    If Col <> 18 Then Exit Sub
    If ButtonDown = 0 Then Exit Sub
    For FOR_CNT = 1 To ss1.MaxRows
        If FOR_CNT <> Row Then
            ss1.Col = 18
            ss1.Row = FOR_CNT
            ss1.Value = 0
        End If
    Next
       
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim sDate     As String
Dim sDateTo   As String
Dim FOR_CNT   As Long
Dim tmpYYMMDD As String

    If Row < 1 Then Exit Sub
    If Col < 11 Then Exit Sub
    
    tmpYYMMDD = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL")
    
    ss1.Row = Row: ss1.Col = Col
    
    'For FOR_CNT = 1 To ss1.MaxRows
    
        ss1.Row = Row
        If ss1.Col = 13 Then
           ss1.Text = tmpYYMMDD
        End If
        ss1.Col = 28
        ss1.Text = txt_PrcLine.Text
        ss1.Col = 27
        ss1.Text = txt_WkPlt
        ss1.Col = 26
        ss1.Text = sUserID
        
        ss1.Col = 31
        ss1.Text = CBO_AREA
        Call ss1_Row_Edit(Row)
    'Next
    
End Sub

Private Sub ss1_LostFocus()
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ss2_LostFocus()
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)
    
    Dim dThk        As Double
    Dim dWid        As Double
    Dim dLen        As Double
    Dim sEndUseCd   As String
    Dim sStlgrd     As String
    
    If Row < 1 Then Exit Sub
    
    ss1.Row = Row
            
    If Col = SPD_THK Or Col = SPD_WID Or Col = SPD_LEN Then
        ss1.Col = SPD_THK:  dThk = Val(ss1.Text & "")
        ss1.Col = SPD_WID:  dWid = Val(ss1.Text & "")
        ss1.Col = SPD_LEN:  dLen = Val(ss1.Text & "")
        ss1.Col = SPD_END_USE:   sEndUseCd = Trim(ss1.Text)
        ss1.Col = SPD_STLGRD:    sStlgrd = Trim(ss1.Text)
        If dThk > 0 And dWid > 0 And dLen > 0 Then
            ss1.Col = SPD_WGT
            ss1.Text = Cal_Plate_Wgt("WGT", sEndUseCd, sStlgrd, dThk, dWid, dLen)
        End If
    End If
    
    Call ss1_Row_Edit(Row)
End Sub

Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
    If Row < 1 Then Exit Sub
       
    Call ss1_Row_Edit(Row)

End Sub
Private Sub ss1_Data_Edit()
    Dim iIdr        As Integer
    Dim iThk        As Long
    Dim iWid        As Long
    Dim iLen        As Long
    Dim iWGT        As Double
    Dim Row         As Long
    Dim sDate       As String
    Dim sDateTo     As String
    
    For iIdr = 1 To ss1.MaxRows
        ss1.Row = iIdr
        ss1.Col = 2
        If ss1.Text = "" Then
           iThk = 0
        Else
           iThk = ss1.Value
        End If
        ss1.Col = 3
        If ss1.Text = "" Then
           iWid = 0
        Else
           iWid = ss1.Value
        End If
        ss1.Col = 4
        If ss1.Text = "" Then
           iLen = 0
        Else
           iLen = ss1.Value
        End If
        
        ss1.Col = 24
        ss1.Text = Gf_ShiftSet3(M_CN1)
        ss1.Col = 25
        ss1.Text = Gf_GroupSet(M_CN1, Gf_ShiftSet3(M_CN1), Gf_DTSet(M_CN1, , "X"))
        ss1.Col = 26
        ss1.Text = sUserID
        ss1.Col = 27
        ss1.Text = txt_WkPlt
        ss1.Col = 28
        ss1.Text = txt_PrcLine
        
    Next iIdr
    
    
End Sub

Private Sub ss1_Row_Edit(ByVal Row As Long)
    Dim iIdr        As Integer
    Dim sLastFlag   As String
    
    ss1.Row = Row
    
    ss1.Col = 0
    ss1.Row = Row
    Select Case Trim(ss1.Text)
          Case "Input", "Update", "Delete"
          Case Else
               ss1.Text = "Update"
    End Select
    
    sLastFlag = ""
    lbl_moplate_wgt.Caption = ""
    For iIdr = 1 To ss1.MaxRows
        ss1.Col = 17
        lbl_moplate_wgt.Caption = Val(lbl_moplate_wgt.Caption & "") + Val(ss1.Text & "")
    Next iIdr
    
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    Dim iCount As Integer

    If Row < 1 Then Exit Sub
    
    ss2.Row = Row
    ss2.Col = 1
    txt_mplate_no.Text = ss2.Text
    ss2.Col = 3
    TXT_PRODCD.Text = ss2.Text
    
    If Trim(txt_mplate_no.Text) <> "" Then
        Call Gf_Sp_Refer(M_CN1, sc1, Mc2, Mc2("nControl"), Mc2("mControl"), False)
        ss1.OperationMode = OperationModeNormal
        Call ss1_Data_Edit
        Call PlateWgtEdit
        
        ss1.Col = 18
        ss1.Row = ss1.MaxRows
        ss1.Value = 1
        
    End If
    
    ss2.Row = Row
    ss2.Col = 3
    If ss2.Text = "MP" Then
        MDIMain.MenuTool.Buttons(7).Enabled = True
        MDIMain.MenuTool.Buttons(8).Enabled = True
        MDIMain.MenuTool.Buttons(9).Enabled = True
        MDIMain.MenuTool.Buttons(11).Enabled = True
        MDIMain.MenuTool.Buttons(12).Enabled = True
    ElseIf ss2.Text = "PP" Then
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(9).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    End If

End Sub

Private Sub txt_plt_DblClick()
    Call txt_plt_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=TXT_PLT
        DD.rControl.Add Item:=TXT_PLT_NAME

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(TXT_PLT)) = TXT_PLT.MaxLength Then
        TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(TXT_PLT.Text), 2)
    Else
        TXT_PLT_NAME.Text = ""
    End If
End Sub

Private Sub txt_WkPlt_Change()
    cbo_PrcLine.Clear
    
    If txt_WkPlt = "C1" Then
       cbo_PrcLine.AddItem "一号线"
       cbo_PrcLine.AddItem "二号线"
    Else
       cbo_PrcLine.AddItem "一号线"
    End If
    cbo_PrcLine.ListIndex = 0
End Sub
