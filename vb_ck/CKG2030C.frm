VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CKG2030C 
   Caption         =   "精整作业指示查询界面_CKG2030C"
   ClientHeight    =   8055
   ClientLeft      =   1110
   ClientTop       =   2025
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10650
      TabIndex        =   13
      Top             =   0
      Width           =   4695
      Begin VB.OptionButton OPT_FP_SLAB_DES 
         BackColor       =   &H00E0E0E0&
         Caption         =   "轧制实绩"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   170
         Value           =   -1  'True
         Width           =   1110
      End
      Begin VB.OptionButton OPT_EP_SLAB_DES 
         BackColor       =   &H00E0E0E0&
         Caption         =   "炼钢计划"
         Height          =   195
         Left            =   3300
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   170
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.OptionButton OPT_FP_SLAB_DES1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "订单材"
         Height          =   195
         Left            =   1860
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   170
         Width           =   1110
      End
   End
   Begin VB.ComboBox CBO_SHIFT 
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
      ItemData        =   "CKG2030C.frx":0000
      Left            =   6195
      List            =   "CKG2030C.frx":000D
      TabIndex        =   1
      Top             =   90
      Width           =   735
   End
   Begin VB.TextBox TXT_SLAB_NO 
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
      Left            =   8745
      TabIndex        =   0
      Top             =   90
      Width           =   1665
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   7440
      Top             =   90
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "板坯号"
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
      Left            =   120
      Top             =   90
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "轧制时间"
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
      Left            =   5190
      Top             =   90
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "班次"
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
   Begin InDate.UDate SDT_PROD_DATE_FROM 
      Height          =   315
      Left            =   1425
      TabIndex        =   2
      Tag             =   "起始日期"
      Top             =   90
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
   Begin InDate.UDate SDT_PROD_DATE_TO 
      Height          =   315
      Left            =   3180
      TabIndex        =   3
      Tag             =   "起始日期"
      Top             =   90
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
   Begin SSSplitter.SSSplitter SP1 
      Height          =   8760
      Left            =   60
      TabIndex        =   5
      Top             =   495
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   15452
      _Version        =   196609
      SplitterBarWidth=   3
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      PaneTree        =   "CKG2030C.frx":001A
      Begin FPSpread.vaSpread ss1 
         Height          =   2790
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   15270
         _Version        =   393216
         _ExtentX        =   26935
         _ExtentY        =   4921
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColsFrozen      =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   16
         MaxRows         =   20
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKG2030C.frx":006C
      End
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   5925
         Left            =   0
         TabIndex        =   7
         Top             =   2835
         Width           =   15270
         _ExtentX        =   26935
         _ExtentY        =   10451
         _Version        =   196609
         SplitterBarWidth=   3
         PaneTree        =   "CKG2030C.frx":0BBB
         Begin Threed.SSPanel SSPanel2 
            Height          =   345
            Left            =   30
            TabIndex        =   8
            Top             =   30
            Width           =   15210
            _ExtentX        =   26829
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   14737632
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSPanel SSP1 
               Height          =   315
               Left            =   10890
               TabIndex        =   9
               Top             =   15
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   196609
               ForeColor       =   16711680
               BackColor       =   16761087
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "轧件"
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel SSP2 
               Height          =   315
               Left            =   12330
               TabIndex        =   10
               Top             =   15
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   196609
               ForeColor       =   16711680
               BackColor       =   16777152
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "母板"
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel SSP3 
               Height          =   315
               Left            =   13770
               TabIndex        =   11
               Top             =   15
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   196609
               ForeColor       =   16711680
               BackColor       =   12648384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "钢板"
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSCheck chk_detail 
               Height          =   285
               Left            =   510
               TabIndex        =   17
               Top             =   30
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   " 明细导出"
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   5460
            Left            =   30
            TabIndex        =   12
            Top             =   435
            Width           =   15210
            _Version        =   393216
            _ExtentX        =   26829
            _ExtentY        =   9631
            _StockProps     =   64
            ColsFrozen      =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   18
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CKG2030C.frx":0C0D
         End
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   2985
      TabIndex        =   4
      Top             =   210
      Width           =   195
   End
End
Attribute VB_Name = "CKG2030C"
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
'-- Program Name      精整作业指示查询界面
'-- Program ID        CKG2030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          YANGMENG
'-- Coder             YANGMENG
'-- Date              2003.7.23
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

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
Dim sc5 As New Collection           'Spread Collection
Dim sc6 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim SE, MPLATE_NO, plate_no As String

Const SS2_BLOCK_SEQ = 2
Const SS2_SEQ = 3
Const SS2_PROD_CD = 4
Const SS2_ORD = 9
Const SS2_SIZE_KND = 10
Const SS2_TRIM_FL = 11
Const SS2_UST_FL = 12
Const SS2_HTM = 13
Const SS2_STDSPEC_YY = 14
Const SS2_STLGRD = 15
Const SS2_VESSEL_NO = 16
Const SS2_COLOR_STROKE = 17

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)

         
    'MASTER Collection
    Mc1.Add Item:="CKG2030C.P_SREFER1", Key:="P-R"
    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"

          Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

    'MASTER Collection
    Mc2.Add Item:="CKG2030C.P_REFER", Key:="P-R"
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"

    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="CKG2030C.P_SREFER", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", "a", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", "a", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
  
    'Spread_Collection
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="CKG2030C.P_SREFER2", Key:="P-R"
    Sc2.Add Item:=pColumn1, Key:="pColumn"
    Sc2.Add Item:=nColumn1, Key:="nColumn"
    Sc2.Add Item:=aColumn1, Key:="aColumn"
    Sc2.Add Item:=mColumn1, Key:="mColumn"
    Sc2.Add Item:=iColumn1, Key:="iColumn"
    Sc2.Add Item:=lColumn1, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    'Spread_Collection
    Sc3.Add Item:=ss2, Key:="Spread"
    Sc3.Add Item:="CKG2030C.P_SREFER3", Key:="P-R"
    Sc3.Add Item:=pColumn2, Key:="pColumn"
    Sc3.Add Item:=nColumn2, Key:="nColumn"
    Sc3.Add Item:=aColumn2, Key:="aColumn"
    Sc3.Add Item:=mColumn2, Key:="mColumn"
    Sc3.Add Item:=iColumn2, Key:="iColumn"
    Sc3.Add Item:=lColumn2, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss2.MaxCols, Key:="Last"
    
    'Spread_Collection
    sc4.Add Item:=ss1, Key:="Spread"
    sc4.Add Item:="CKG2030C.P_SREFER1", Key:="P-R"
    sc4.Add Item:=pColumn1, Key:="pColumn"
    sc4.Add Item:=nColumn1, Key:="nColumn"
    sc4.Add Item:=aColumn1, Key:="aColumn"
    sc4.Add Item:=mColumn1, Key:="mColumn"
    sc4.Add Item:=iColumn1, Key:="iColumn"
    sc4.Add Item:=lColumn1, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Spread_Collection
    sc5.Add Item:=ss1, Key:="Spread"
    sc5.Add Item:="CKG2030C.P_SREFER5", Key:="P-R"
    sc5.Add Item:=pColumn1, Key:="pColumn"
    sc5.Add Item:=nColumn1, Key:="nColumn"
    sc5.Add Item:=aColumn1, Key:="aColumn"
    sc5.Add Item:=mColumn1, Key:="mColumn"
    sc5.Add Item:=iColumn1, Key:="iColumn"
    sc5.Add Item:=lColumn1, Key:="lColumn"
    sc5.Add Item:=1, Key:="First"
    sc5.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Spread_Collection
    sc6.Add Item:=ss2, Key:="Spread"
    sc6.Add Item:="CKG2030C.P_SREFER6", Key:="P-R"
    sc6.Add Item:=pColumn2, Key:="pColumn"
    sc6.Add Item:=nColumn2, Key:="nColumn"
    sc6.Add Item:=aColumn2, Key:="aColumn"
    sc6.Add Item:=mColumn2, Key:="mColumn"
    sc6.Add Item:=iColumn2, Key:="iColumn"
    sc6.Add Item:=lColumn2, Key:="lColumn"
    sc6.Add Item:=1, Key:="First"
    sc6.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

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
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"))
    Call Gp_Sp_Setting(Sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(Sc2)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc2.Item("Spread"), "G-System.INI", Me.Name)
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc2.Item("Spread"), "G-System.INI", Me.Name)

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
    
    Set pControl3 = Nothing
    Set nControl3 = Nothing
    Set iControl3 = Nothing
    Set rControl3 = Nothing
    Set cControl3 = Nothing
    Set aControl3 = Nothing
    Set lControl3 = Nothing
    Set mControl3 = Nothing
    
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
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set Sc3 = Nothing
    Set sc4 = Nothing
    Set sc5 = Nothing
    Set sc6 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Sc2) Then
        If Gf_Sp_Cls(Sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)

        End If
    End If
    
    MPLATE_NO = ""
    plate_no = ""
    
End Sub

Public Sub Form_Exc()

    If chk_detail.Value = -1 Then
        Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Else
        Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
    If OPT_FP_SLAB_DES.Value = True Then
        If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    ElseIf OPT_FP_SLAB_DES1.Value = True Then
        If Gf_Sp_Refer(M_CN1, sc4, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    Else
        If Gf_Sp_Refer(M_CN1, sc5, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    End If
    
    Call Gf_Sp_Cls(Sc2)
            
    Exit Sub

Refer_Err:

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

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()
     SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     If SDT_PROD_DATE_TO.RawData = "" Then
     SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub
Private Sub SDT_PROD_DATE_TO_GotFocus()
     SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    Dim lRow As Long
    Dim sBlockSeq As String
    Dim sSeq As String
    
'    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ROW <= 0 Then Exit Sub
    If Col > 1 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 1
    TXT_SLAB_NO.Text = ss1.Text
    
    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows)
    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW, , SSP1.BackColor)
    If OPT_FP_SLAB_DES.Value = True Then
        Call Gf_Sp_Refer(M_CN1, Sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    ElseIf OPT_FP_SLAB_DES1.Value = True Then
        Call Gf_Sp_Refer(M_CN1, Sc3, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    Else
        Call Gf_Sp_Refer(M_CN1, sc6, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    End If
    ss2.OperationMode = OperationModeNormal
    TXT_SLAB_NO.Text = ""
    
    For lRow = 1 To ss2.MaxRows
    
        ss2.ROW = lRow
        ss2.Col = SS2_BLOCK_SEQ: sBlockSeq = ss2.Text
        ss2.Col = SS2_SEQ:       sSeq = ss2.Text
        
        If sBlockSeq & sSeq = "0000" Then
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ss2.ROW, ss2.ROW, , SSP1.BackColor)
            ss2.Col = SS2_PROD_CD:       ss2.Text = "轧件"
        ElseIf sSeq = "00" Then
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ss2.ROW, ss2.ROW, , SSP2.BackColor)
            ss2.Col = SS2_PROD_CD:       ss2.Text = "母板" & sBlockSeq
        Else
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, ss2.ROW, ss2.ROW, , SSP3.BackColor)
            ss2.Col = SS2_PROD_CD: ss2.Text = "钢板"
        End If
    
    Next lRow

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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)
    ss2.ROW = ss2.ActiveRow
    ss2.Col = 1
    MPLATE_NO = ss2.Text
    ss2.Col = 2
    plate_no = ss2.Text
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub



