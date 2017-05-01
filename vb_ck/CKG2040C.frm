VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CKG2040C 
   Caption         =   "批号查询及修改_CKG2040C"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   1440
   ClientWidth     =   15240
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8505
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   18420
      _ExtentX        =   32491
      _ExtentY        =   15002
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      Locked          =   -1  'True
      PaneTree        =   "CKG2040C.frx":0000
      Begin Threed.SSFrame SSFrame3 
         Height          =   585
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   18420
         _ExtentX        =   32491
         _ExtentY        =   1032
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.TextBox txt_all_ord 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10800
            MaxLength       =   11
            TabIndex        =   17
            Text            =   "N"
            Top             =   90
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.TextBox txt_HeatNo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1670
            MaxLength       =   10
            TabIndex        =   4
            Tag             =   "工厂"
            Top             =   120
            Width           =   2220
         End
         Begin VB.TextBox txt_plt_dec 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6240
            MaxLength       =   11
            TabIndex        =   3
            Top             =   120
            Width           =   1260
         End
         Begin VB.TextBox txt_plt 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5590
            MaxLength       =   2
            TabIndex        =   2
            Top             =   120
            Width           =   630
         End
         Begin InDate.ULabel ULabel63 
            Height          =   315
            Left            =   120
            Top             =   120
            Width           =   1515
            _ExtentX        =   2672
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
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   4260
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "板坯来源"
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
         Begin InDate.UDate txt_DATE 
            Height          =   315
            Left            =   9195
            TabIndex        =   5
            Tag             =   "起始日期"
            Top             =   120
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
            Left            =   7890
            Top             =   120
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            Caption         =   "批号日期"
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
         Begin Threed.SSCheck chk_all_ord 
            Height          =   285
            Left            =   11430
            TabIndex        =   16
            Top             =   150
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
            Caption         =   "多订单查询"
         End
         Begin Threed.SSPanel SSPpdt 
            Height          =   345
            Left            =   13260
            TabIndex        =   18
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   609
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
            Caption         =   "一坯多订单"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   900
         Left            =   0
         TabIndex        =   6
         Top             =   645
         Width           =   18420
         _ExtentX        =   32491
         _ExtentY        =   1588
         _Version        =   196609
         BackColor       =   12632319
         Begin VB.TextBox TXT_WGT 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2330
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   21
            Top             =   480
            Width           =   1560
         End
         Begin VB.TextBox TXT_CNT 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1670
            MaxLength       =   5
            TabIndex        =   20
            Top             =   480
            Width           =   630
         End
         Begin VB.TextBox Temp 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   13320
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   19
            Tag             =   "工厂"
            Top             =   120
            Width           =   1500
         End
         Begin VB.TextBox txt_LotNo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1670
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   11
            Tag             =   "工厂"
            Top             =   135
            Width           =   2220
         End
         Begin VB.TextBox txt_InPlt 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5590
            MaxLength       =   2
            TabIndex        =   10
            Top             =   135
            Width           =   630
         End
         Begin VB.TextBox txt_INplt_dec 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   9
            Top             =   135
            Width           =   1260
         End
         Begin VB.TextBox txt_InSeq 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9195
            MaxLength       =   4
            TabIndex        =   8
            Tag             =   "工厂"
            Top             =   135
            Width           =   720
         End
         Begin VB.TextBox txt_OutSeq 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   9930
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   7
            Tag             =   "工厂"
            Top             =   135
            Width           =   720
         End
         Begin InDate.ULabel ULabel23 
            Height          =   315
            Left            =   120
            Top             =   135
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            Caption         =   "轧制批号"
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
            Left            =   4260
            Top             =   135
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            Caption         =   "板坯来源"
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   7890
            Top             =   135
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            Caption         =   "当月顺序"
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
         Begin Threed.SSOption opt_bz1 
            Height          =   330
            Left            =   7860
            TabIndex        =   12
            Top             =   825
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   2
            ForeColor       =   255
            BackColor       =   14737632
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "外购坯"
         End
         Begin Threed.SSOption opt_bz2 
            Height          =   330
            Left            =   8940
            TabIndex        =   13
            Top             =   825
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   582
            _Version        =   196609
            Font3D          =   2
            ForeColor       =   8421504
            BackColor       =   14737632
            BackStyle       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "来料加工"
         End
         Begin Threed.SSCommand cmd_mill_exc 
            Height          =   405
            Left            =   11400
            TabIndex        =   15
            Top             =   90
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   714
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   11.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "轧制单导出"
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   120
            Top             =   480
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            Caption         =   "件数/重量（吨）"
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
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   6900
         Left            =   0
         TabIndex        =   14
         Top             =   1605
         Width           =   18420
         _Version        =   393216
         _ExtentX        =   32491
         _ExtentY        =   12171
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
         MaxCols         =   30
         MaxRows         =   50
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKG2040C.frx":0072
      End
   End
End
Attribute VB_Name = "CKG2040C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   HTM System
'-- Program Name      热处理出炉作业实绩查询及修改
'-- Program ID        CKG2040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2007.11.20
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String            'Form Type
Public Toolbar_St As String          'Active Form ToolBar Setting
Public sAuthority As String          'Active Form Authority Setting
Public sDateTime As String           'Active Form Time Setting
Public sQuery_Rt As String

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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_LOT_NO = 1
Const SS1_LOT_NO_TEMP = 2
Const SS1_PLT = 3
Const SS1_SLAB_NO = 4
Const SS1_STDSPEC = 5
Const SS1_STLGRD = 6
Const SS1_THK = 7
Const SS1_WID = 8
Const SS1_LEN = 9
Const SS1_LEN_L = 10
Const SS1_THK_L = 11
Const SS1_TRIM_FL = 12
Const SS1_SIZE_KND = 13
Const SS1_ASROLL_LEN = 22
Const SS1_ORD_CNT = 28
Const SS1_PLATE_CNT = 29


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_HeatNo, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
           Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_all_ord, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
           
    'MASTER Collection
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
          
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
   'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="CKG2040C.P_REFER1", Key:="P-R"
    Sc1.Add Item:="CKG2040C.P_MODIFY", Key:="P-M"
    
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 2, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
End Sub
Public Sub Form_Ref()

Dim sL2_Send As String
Dim sSlab_No As String
Dim sPrc_Sts As String
Dim iRow As Integer
Dim iCol As Integer
On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1)
    ss1.OperationMode = OperationModeNormal
            
    MDIMain.MenuTool.Buttons(7).Enabled = False                 'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False                 'Row Cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False                'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False                'Paste
    MDIMain.MenuTool.Buttons(14).Enabled = True                 'Excel
    
    With ss1
          For iRow = 1 To .MaxRows
             .ROW = iRow
             .Col = SS1_ORD_CNT
              If .Value > 1 Then
                For iCol = 1 To .MaxCols
                   .Col = iCol
                   .BackColor = SSPpdt.BackColor
                Next iCol
              End If
          Next iRow
    End With
      
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    Dim sQuery As String
    Dim sCurDate As String
    Dim sLotDate As String
    Dim sInsDate As String
    
    Dim sOutSeq  As String

    
    If txt_all_ord.Text <> "N" Then
        Call Gp_MsgBoxDisplay("请先取消多订单查询选项", "I")
        Exit Sub
    End If

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    sCurDate = Gf_DTSet(M_CN1, "D", "X")
    sLotDate = "____" & Mid(sCurDate, 3, 4) & "%"
    sInsDate = Gf_DTSet_D(M_CN1, "D", "-10")
    
'    sQuery = "SELECT NVL(MAX(SUBSTR(MILL_LOT_NO,11,4)),0) FROM EP_MILL_INS3 WHERE MILL_LOT_NO LIKE '" & sLotDate & "' AND INS_DATE >  '" & sInsDate & "'"
    sQuery = "SELECT NVL(MAX(SUBSTR(MILL_LOT_NO,9,6)),'000000') FROM EP_MILL_INS3 WHERE MILL_LOT_NO LIKE '" & sLotDate & "' AND INS_DATE >  '" & sInsDate & "'"
    sOutSeq = Gf_CodeFind(M_CN1, sQuery)
    txt_OutSeq = Mid(sOutSeq, 3, 4)
    sOutSeq = Val(sOutSeq) + 1
    sOutSeq = Format(sOutSeq, "000000")
    txt_InSeq = Mid(sOutSeq, 3, 4)
    TXT_CNT.Text = ""
    TXT_WGT.Text = ""
    
    Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
    
End Sub

Private Sub chk_all_ord_Click(Value As Integer)
    If chk_all_ord.Value = ssCBUnchecked Then
        txt_all_ord.Text = "N"
    Else
        txt_all_ord.Text = "Y"
    End If
    ss1.MaxRows = 0
End Sub

Private Sub cmd_mill_exc_Click()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

   Call Gp_CKG2040C_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
   
End Sub

Private Sub opt_bz1_Click(Value As Integer)
If opt_bz1.Value = True Then
   txt_LotNo = "74" & "10" & Mid(txt_DATE.RawData, 3, 6) & txt_InSeq
   opt_bz1.ForeColor = &HFF&
   opt_bz2.ForeColor = &H808080
End If
End Sub

Private Sub opt_bz2_Click(Value As Integer)
If opt_bz2.Value = True Then
   txt_LotNo = "74" & "30" & Mid(txt_DATE.RawData, 3, 6) & txt_InSeq
   opt_bz1.ForeColor = &H808080
   opt_bz2.ForeColor = &HFF&
End If
End Sub




Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    Dim iRow As Integer
    Dim tmpLOTNO As String
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
    
    If BlockRow < 0 Then Exit Sub
    
    If Trim(txt_LotNo) = "" Or Len(txt_LotNo) <> 14 Then
        MsgBox "请先确认批号......!"
        Exit Sub
    End If
    
    For iRow = BlockRow To BlockRow2
    
        ss1.ROW = iRow
        ss1.Col = 0
        If ss1.Text = "Update" Then
            ss1.Text = ""
            ss1.Col = SS1_LOT_NO
            ss1.Text = ""
            ss1.Col = SS1_LOT_NO_TEMP
            tmpLOTNO = ss1.Text
            ss1.Col = SS1_LOT_NO
            ss1.Text = tmpLOTNO
            
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow)
            Call Gp_Sp_BlockColor(ss1, SS1_LOT_NO, SS1_LOT_NO, iRow, iRow, , 12648447)
            Call Gp_Sp_BlockColor(ss1, SS1_SLAB_NO, SS1_SLAB_NO, iRow, iRow, , 12648447)
            
        Else
            ss1.Col = SS1_PLT
            If ss1.Text = txt_InPlt Then
                ss1.Col = 0
                ss1.Text = "Update"
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)
                ss1.Col = SS1_LOT_NO
                ss1.Text = txt_LotNo
            End If
        End If
        
    Next iRow

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
        MDIMain.MenuTool.Buttons(7).Enabled = False                 'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False                 'Row Delete
        MDIMain.MenuTool.Buttons(9).Enabled = False                 'Row Cancel
        MDIMain.MenuTool.Buttons(11).Enabled = False                'Copy
        MDIMain.MenuTool.Buttons(12).Enabled = False                'Paste
        MDIMain.MenuTool.Buttons(14).Enabled = True                 'Excel
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Dim sQuery As String
    Dim sCurDate As String
    Dim sLotDate As String
    Dim sInsDate As String
    
    Dim sOutSeq  As String

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"), False)
   
    Call Gf_Sp_Cls(Sc1)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "CG-System.INI", Me.Name)
    
    txt_InPlt = "B1"
    txt_INplt_dec = "#1 炼钢"
    
    sCurDate = Gf_DTSet(M_CN1, "D", "X")
    txt_DATE.RawData = sCurDate
    sLotDate = "____" & Mid(sCurDate, 3, 4) & "%"
    sInsDate = Gf_DTSet_D(M_CN1, "D", "-10")
    
'    sQuery = "SELECT NVL(MAX(SUBSTR(MILL_LOT_NO,11,4)),0) FROM EP_MILL_INS3 WHERE MILL_LOT_NO LIKE '" & sLotDate & "' AND INS_DATE >  '" & sInsDate & "'"
    sQuery = "SELECT NVL(MAX(SUBSTR(MILL_LOT_NO,9,6)),'000000') FROM EP_MILL_INS3 WHERE MILL_LOT_NO LIKE '" & sLotDate & "' AND INS_DATE >  '" & sInsDate & "'"
    sOutSeq = Gf_CodeFind(M_CN1, sQuery)
    txt_OutSeq = Mid(sOutSeq, 3, 4)
    sOutSeq = Val(sOutSeq) + 1
    sOutSeq = Format(sOutSeq, "000000")
    txt_InSeq = Mid(sOutSeq, 3, 4)
    
    ss1.OperationMode = OperationModeNormal
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "CG-System.INI", Me.Name)

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
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()

    Dim sMesg As String

    Call Gp_Ms_Cls(Mc1("rControl"))
'    Call Gp_Ms_Cls(Mc2("rControl"))

    Call Gf_Sp_Cls(Sc1)

    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = False                'Excel
    End With
    
    TXT_CNT.Text = ""
    TXT_WGT.Text = ""
    
    pControl1(1).SetFocus
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    Dim iDR As Long
    Dim sStlgrd As String
    Dim dPthk As String
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ROW = 0 And (Col = 1 Or Col = 4) Then
        Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    End If
    
    If ROW < 1 Or Col > 0 Then Exit Sub
    If ss1.MaxRows < 1 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = SS1_STDSPEC
    sStlgrd = Trim(ss1.Text)
    ss1.Col = SS1_THK
    dPthk = Trim(ss1.Text)
    
    For iDR = 1 To ss1.MaxRows
        ss1.ROW = iDR
        ss1.Col = 0
        If ss1.Text = "Update" Then
            ss1.Col = SS1_STDSPEC
            If sStlgrd <> Trim(ss1.Text) Then
                Call Gp_MsgBoxDisplay("不一样标准")
                Exit Sub
            End If
'            ss1.Col = SS1_THK
'            If dPthk <> Trim(ss1.Text) Then
'                Call Gp_MsgBoxDisplay("不一样厚度")
'                Exit Sub
'            End If
        End If
    Next iDR
    
     Dim plate_no As String
    Dim iCnt As Integer
    Dim iPlate_cnt As Integer
    Dim iPlate_wgt As Double

    Dim tRow  As Integer

    iPlate_cnt = 0
    iPlate_wgt = 0

    If ROW <= 0 Then Exit Sub

    ss1.ROW = ROW
    ss1.Col = 0
  
    If ss1.Text = "" Then
        ss1.Col = 3
        plate_no = ss1.Text
'        With ss2
'
'            For iCnt = .MaxRows To 1 Step -1
'               .Col = 0
'               .ROW = iCnt
'                If Trim(.Text) = "Input" Then
'                   .Col = 2
'                    If .Text = plate_no Then
'                       .Text = ""
'                       .BackColor = &H80000005
'                       .Col = 0
'                       .Text = ""
'                        Exit For
'                    End If
'                End If
'            Next iCnt
'
'        End With
        ss1.Col = 0
        ss1.Text = "Update"
        With ss1
               For iCnt = 1 To .MaxCols Step 1
                    .Col = iCnt
                    .BackColor = &H80000005
               Next iCnt
        End With

        With ss1

               For iCnt = 1 To .MaxRows Step 1
                    .Col = 0
                    .ROW = iCnt
                     If Trim(.Text) <> "" Then
                         iPlate_cnt = iPlate_cnt + 1
                         .Col = 25
                         iPlate_wgt = iPlate_wgt + .Value
                     End If
               Next iCnt

        End With
        TXT_CNT.Text = Str(iPlate_cnt)
        TXT_WGT.Text = Str(iPlate_wgt)
        ss1.ROW = ROW
        ss1.Col = 0
        ss1.Text = ""
        Exit Sub
    End If
If ss1.Text <> "" Then


    ss1.Col = 3
    plate_no = Trim(ss1.Text)
'
''    If ss2.MaxRows = 0 Then
''       Exit Sub
''    End If
'
    ss1.ROW = ROW
    ss1.Col = 0
    ss1.Text = ""
'
    With ss1
           For iCnt = 1 To .MaxCols Step 1
                .Col = iCnt
                .BackColor = &HFFC0FF
           Next iCnt
    End With

    With ss1

           For iCnt = 1 To .MaxRows Step 1
                .Col = 0
                 .ROW = iCnt
                 If Trim(.Text) <> "" Then
                     iPlate_cnt = iPlate_cnt + 1
                     .Col = 25
                     iPlate_wgt = iPlate_wgt + .Value
                 End If
           Next iCnt

    End With
     ss1.ROW = ROW
    ss1.Col = 0
    ss1.Text = "Update"

    TXT_CNT.Text = Str(iPlate_cnt)
    TXT_WGT.Text = Str(iPlate_wgt)
    
    End If

'    With ss2
'
'        tRow = .ActiveRow
'        .ROW = tRow
'        .Col = 2
'
'    If Len(.Text) = 14 Then
'
'         For iCnt = .MaxRows To 1 Step -1
'            .Col = 2
'            .ROW = iCnt
'             If Trim(.Text) = "" Then
'                .Text = plate_no
'                .Col = 0
'                .Text = "Input"
'                .Col = 12
'                .Text = sUserID
'                 Exit Sub
'             End If
'         Next iCnt
'
'    Else
'
'        .Col = 2
'        .ROW = tRow
'         If Trim(.Text) = "" Then
'            .Text = plate_no
'            .Col = 0
'            .Text = "Input"
'            .Col = 12
'            .Text = sUserID
'             If tRow > 1 Then
'             Call .SetActiveCell(1, tRow - 1)
'             End If
'             Exit Sub
'         End If
'
'    End If
'
'    End With


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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub txt_InPlt_Change()

Dim lot_no As String

If Len(txt_InPlt) = 2 Then
    If txt_InPlt = "B1" Then
       txt_LotNo = "74" & "50" & Mid(txt_DATE.RawData, 3, 6) & txt_InSeq
       opt_bz1.Visible = False
       opt_bz2.Visible = False
       opt_bz1.Value = False
       opt_bz2.Value = False
    ElseIf txt_InPlt = "B3" Then
       txt_LotNo = "74" & "01" & Mid(txt_DATE.RawData, 3, 6) & txt_InSeq
       opt_bz1.Visible = False
       opt_bz2.Visible = False
       opt_bz1.Value = False
       opt_bz2.Value = False
    ElseIf txt_InPlt = "BZ" Then
       txt_LotNo = "74" & "10" & Mid(txt_DATE.RawData, 3, 6) & txt_InSeq
       opt_bz1.Visible = True
       opt_bz2.Visible = True
       opt_bz1.Value = True
    Else
       MsgBox "请确认板坏生产工厂...!"
       Exit Sub
    End If
End If
    
End Sub

Private Sub txt_InPlt_DblClick()
    Call txt_InPlt_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_InPlt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_InPlt
        DD.rControl.Add Item:=txt_INplt_dec
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_InPlt.Text)) = txt_InPlt.MaxLength Then
        txt_INplt_dec.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_InPlt.Text), 2)
    Else
        txt_INplt_dec.Text = ""
    End If
End Sub

Private Sub txt_InSeq_Change()
    If txt_InPlt = "B1" Then
       txt_LotNo = "74" & "50" & Mid(txt_DATE.RawData, 3, 6) & txt_InSeq
    ElseIf txt_InPlt = "B3" Then
       txt_LotNo = "74" & "01" & Mid(txt_DATE.RawData, 3, 6) & txt_InSeq
    ElseIf txt_InPlt = "BZ" Then
       If opt_bz1.Value = True Then
          txt_LotNo = "74" & "10" & Mid(txt_DATE.RawData, 3, 6) & txt_InSeq
       ElseIf opt_bz2.Value = True Then
          txt_LotNo = "74" & "30" & Mid(txt_DATE.RawData, 3, 6) & txt_InSeq
       End If
    Else
       MsgBox "请确认板坏生产工厂...!"
       Exit Sub
    End If
End Sub

Private Sub txt_plt_DblClick()
    Call txt_plt_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_dec
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        txt_plt_dec.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_dec.Text = ""
    End If
End Sub
'---------------------------------------------------------------------------------------
'   1.ID           : Gf_DTSet_D
'   2.Name         : Get System/Vb Date,Time
'   3.Input  Value : Conn Connection, {DTCheck,Date_Num String}
'   4.Return Value : Variant
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .24
'   7.Modify Date  :
'   8.Comment      : Get System/Vb Date,Time
'---------------------------------------------------------------------------------------
Private Function Gf_DTSet_D(Conn As ADODB.Connection, Optional DTCheck As String = "S", Optional Date_Num As String = 0) As Variant

On Error GoTo DTSet_Error

    Dim sQuery As String
    Dim sQuery_Len As Long
    
    Select Case DTCheck
           Case "S"
           sQuery = "SELECT TO_CHAR(SYSDATE" & Date_Num & ",'YYYYMMDDHH24MISS') FROM DUAL"
           sQuery_Len = 14
           Case "I"
           sQuery = "SELECT TO_CHAR(SYSDATE" & Date_Num & ",'YYYYMMDDHH24MI') FROM DUAL"
           sQuery_Len = 12
           Case "H"
           sQuery = "SELECT TO_CHAR(SYSDATE" & Date_Num & ",'YYYYMMDDHH24') FROM DUAL"
           sQuery_Len = 10
           Case "D"
           sQuery = "SELECT TO_CHAR(SYSDATE" & Date_Num & ",'YYYYMMDD') FROM DUAL"
           sQuery_Len = 8
           Case "T"
           sQuery = "SELECT TO_CHAR(SYSDATE,'HH24MISS') FROM DUAL"
           sQuery_Len = 6
           Case "M"
           sQuery = "SELECT TO_CHAR(SYSDATE" & Date_Num & ",'YYYYMM') FROM DUAL"
           sQuery_Len = 6
           Case "Y"
           sQuery = "SELECT TO_CHAR(SYSDATE" & Date_Num & ",'YYYY') FROM DUAL"
           sQuery_Len = 4
    End Select
       
    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_DTSet_D = "00000000000000": Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            If VarType(AdoRs.Fields(0)) = vbNull Then
                Gf_DTSet_D = ""
            Else
                Gf_DTSet_D = AdoRs.Fields(0)
            End If
        End If
        
    Else
        Gf_DTSet_D = "00000000000000"
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

DTSet_Error:

    Set AdoRs = Nothing
    Gf_DTSet_D = "00000000000000"

End Function

'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Excel
'   2.Name         : Spread --> Excel
'   3.Input  Value : Fm Form, sPname Variant, bLkcol1 Long, bLkcol2 Long, bLkrow1 Long, bLkrow2 Long
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread --> Excel
'---------------------------------------------------------------------------------------
Private Sub Gp_CKG2040C_Excel(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
        If bLkcol1 = 0 Then
           bLkcol1 = 1
        End If
        
        If bLkcol2 = 0 Then
            bLkcol2 = -1
        End If
        
        If bLkrow2 = 0 Then
            bLkrow2 = -1
        End If
        
'        Clipboard.Clear
'
'        .Col = bLkcol1: .Col2 = bLkcol2
'        .ROW = bLkrow1: .Row2 = bLkrow2
'        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "@"
        
        ss1.ROW = SpreadHeader + 1 'ss1.ColHeaderRow + 1
        ss1.Col = SS1_SLAB_NO:      xlSheet.Range("B2").Value = ss1.Text:    xlSheet.Range("C2").Value = "分段号"
        ss1.Col = SS1_LOT_NO:       xlSheet.Range("D2").Value = ss1.Text
        ss1.Col = SS1_STDSPEC:      xlSheet.Range("E2").Value = ss1.Text
        ss1.Col = SS1_THK:          xlSheet.Range("F2").Value = ss1.Text
        ss1.Col = SS1_WID:          xlSheet.Range("G2").Value = ss1.Text
        ss1.Col = SS1_LEN:          xlSheet.Range("H2").Value = ss1.Text
        ss1.Col = SS1_TRIM_FL:      xlSheet.Range("I2").Value = ss1.Text
        ss1.Col = SS1_SIZE_KND:     xlSheet.Range("J2").Value = ss1.Text
        ss1.Col = SS1_LEN_L:        xlSheet.Range("K2").Value = ss1.Text
        ss1.Col = SS1_THK_L:        xlSheet.Range("L2").Value = ss1.Text
        ss1.Col = SS1_ASROLL_LEN:   xlSheet.Range("M2").Value = ss1.Text:
        ss1.Col = SS1_ORD_CNT:      xlSheet.Range("N2").Value = ss1.Text:
        ss1.Col = SS1_PLATE_CNT:    xlSheet.Range("O2").Value = ss1.Text:    xlSheet.Range("P2").Value = "备注"
        
        Clipboard.Clear
        ss1.SetSelection SS1_SLAB_NO, 1, SS1_SLAB_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("B3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_LOT_NO, 1, SS1_LOT_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("D3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_STDSPEC, 1, SS1_STDSPEC, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("E3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_THK, 1, SS1_LEN, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("F3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_TRIM_FL, 1, SS1_SIZE_KND, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("I3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_LEN_L, 1, SS1_THK_L, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("K3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SS1_ASROLL_LEN, 1, SS1_PLATE_CNT, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("M3").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        ss1.ClearSelection
        Screen.MousePointer = vbDefault
    
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:

    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel" & Error, "W")

End Sub





