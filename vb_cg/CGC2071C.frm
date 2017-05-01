VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGC2071C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "精整线在线查询_CGC2071C"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14595
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_SEQ 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6360
      MaxLength       =   12
      TabIndex        =   15
      Top             =   540
      Width           =   870
   End
   Begin VB.TextBox txt_line 
      Alignment       =   2  'Center
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
      Left            =   9570
      MaxLength       =   1
      TabIndex        =   10
      Tag             =   "CD_MANA_NO"
      Text            =   "1"
      Top             =   120
      Width           =   480
   End
   Begin VB.TextBox TXT_MAT_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      MaxLength       =   14
      TabIndex        =   2
      Top             =   540
      Width           =   2160
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
      ItemData        =   "CGC2071C.frx":0000
      Left            =   6360
      List            =   "CGC2071C.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox CBO_GROUP 
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
      ItemData        =   "CGC2071C.frx":0004
      Left            =   7215
      List            =   "CGC2071C.frx":0006
      TabIndex        =   0
      Tag             =   "班别"
      Top             =   120
      Width           =   855
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4485
      Left            =   60
      TabIndex        =   3
      Top             =   960
      Width           =   15060
      _Version        =   393216
      _ExtentX        =   26564
      _ExtentY        =   7911
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   36
      MaxRows         =   10
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGC2071C.frx":0008
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   120
      Top             =   540
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "查询号"
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
      Left            =   5040
      Top             =   540
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "分段号"
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
      Left            =   5040
      Top             =   120
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "班次/别"
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
      Left            =   1440
      TabIndex        =   4
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
   Begin InDate.UDate SDT_PROD_DATE_TO 
      Height          =   315
      Left            =   3195
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
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1290
      _ExtentX        =   2275
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
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   735
      Left            =   10080
      TabIndex        =   7
      Top             =   120
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   1296
      _Version        =   196609
      BackColor       =   12632319
      Begin Threed.SSOption opt_line1 
         Height          =   255
         Left            =   270
         TabIndex        =   8
         Top             =   60
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "# 1"
         Value           =   -1
      End
      Begin Threed.SSOption opt_line2 
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   60
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "# 2"
      End
      Begin Threed.SSOption opt_line3 
         Height          =   255
         Left            =   270
         TabIndex        =   13
         Top             =   420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "# 3"
      End
      Begin Threed.SSOption opt_line4 
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   12632319
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "# 4"
      End
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   8520
      Top             =   120
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "剪切线    "
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
      ForeColor       =   255
   End
   Begin Threed.SSPanel SSP3 
      Height          =   315
      Left            =   13800
      TabIndex        =   11
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   255
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "计划取样"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP1 
      Height          =   315
      Left            =   13800
      TabIndex        =   12
      Top             =   540
      Width           =   1320
      _ExtentX        =   2328
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
      Caption         =   "一坯多订单"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSCommand CMD_EXCEL 
      Height          =   375
      Left            =   8490
      TabIndex        =   16
      Top             =   540
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   196609
      ForeColor       =   255
      Caption         =   "Excel导出"
   End
   Begin Threed.SSPanel SSP2 
      Height          =   315
      Left            =   12360
      TabIndex        =   17
      Top             =   540
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   0
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "热处理指示"
      FloodColor      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   3015
      Left            =   0
      TabIndex        =   18
      Top             =   5460
      Width           =   15120
      _Version        =   393216
      _ExtentX        =   26670
      _ExtentY        =   5318
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      MaxRows         =   10
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGC2071C.frx":1137
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   3015
      TabIndex        =   6
      Top             =   240
      Width           =   195
   End
End
Attribute VB_Name = "CGC2071C"
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
'-- Program Name      钢板指示查询界面
'-- Program ID        CGC2071C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Li   Qian
'-- Date              2011.2.15
'-- Description       中板精整线上/下线工位，钢板相关剪切信息查询
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
Public sQuery_load As String        'Active Form sQuery Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

'Dim pControl1 As New Collection      'Master Primary Key Collection
'Dim nControl1 As New Collection      'Master Necessary Collection
'Dim mControl1 As New Collection      'Master Maxlength check Collection
'Dim iControl1 As New Collection      'Master Insert Collection
'Dim rControl1 As New Collection      'Master Refer Collection
'Dim cControl1 As New Collection      'Master Copy Collection
'Dim aControl1 As New Collection      'Master -> Spread Collection
'Dim lControl1 As New Collection      'Master Lock Collection

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
'Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim mOplate_No As String

Const SS1_MOTHER_NO = 1
Const SS1_MV_DATE = 2
Const SS1_SHIFT = 3
Const SS1_TRNS_CMPY_CD = 4
Const SS1_OUT_SHEET_NO = 5
Const SS1_TRIM_FL = 6
Const SS1_PLAN_SMP = 7
Const SS1_PLATE_NO = 8
Const SS1_ORD_THK = 11
Const SS1_ORD_WID = 12
Const SS1_ORD_LEN = 13
Const SS1_SIZE_KND = 14
Const SS1_LEN_LIM = 15
Const SS1_THK_LIM = 16
Const SS1_APLY_STDSPEC = 18
Const SS1_LEN = 22
Const SS1_UST_STATUS = 23
Const SS1_GAS_STATUS = 24
Const SS1_CL_STATUS = 25
Const SS1_HTM_METH = 26
Const SS1_QT = 27
Const SS1_ORD_NO = 28
Const SS1_CUST_CD = 30
Const SS1_CUST_NAME = 31
Const SS1_ORD_CNT = 35
Const SS1_ORD_REMARK = 32
Const SS1_STDSPEC_ORG_KND = 33
Const SS1_STDSPEC_STLGRD = 34
Const SS1_CD_MANA_NO = 36


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_SEQ, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
    Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

    'MASTER Collection
    'Mc2.Add Item:="CGC2071C.P_SREFER2", Key:="P-R"
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
     
        Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
'
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGC2071C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    
     'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
  
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGC2071C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
     
End Sub

Private Sub CMD_EXCEL_Click()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

   Call Gp_CGC2071C_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
   
End Sub

Private Sub opt_line1_Click(Value As Integer)
   txt_line.Text = "1"
   Call Form_Ref
End Sub

Private Sub opt_line2_Click(Value As Integer)
   txt_line.Text = "2"
   Call Form_Ref
End Sub

Private Sub opt_line3_Click(Value As Integer)
   txt_line.Text = "3"
   Call Form_Ref
End Sub

Private Sub opt_line4_Click(Value As Integer)
   txt_line.Text = "4"
   Call Form_Ref
End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()
     If SDT_PROD_DATE_FROM.RawData = "" Then
        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_DATE_TO_GotFocus()
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    CBO_SHIFT.AddItem "1"
    CBO_SHIFT.AddItem "2"
    CBO_SHIFT.AddItem "3"
    
    CBO_GROUP.AddItem "A"
    CBO_GROUP.AddItem "B"
    CBO_GROUP.AddItem "C"
    CBO_GROUP.AddItem "D"
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColHidden(ss1, SS1_ORD_CNT, True)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing

    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
       Call Gp_Ms_Cls(Mc1("pControl"))
       SDT_PROD_DATE_FROM.RawData = ""
       SDT_PROD_DATE_TO.RawData = ""
    End If

End Sub

Public Sub Form_Ref()
    
    Dim sMesg As String
    Dim lRow As Long
    '定义一个变量标记，用来控制颜色显示
    Dim iColor As Integer
    Dim sord_cnt As Integer
    Dim sHtm_Meth As String
        
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Nothing) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    For lRow = 1 To ss1.MaxRows

        ss1.ROW = lRow:       ss1.Col = SS1_MOTHER_NO
        '取母板号，初始值为空，把颜色标记置为1
        If mOplate_No = "" Then
            iColor = 1
        Else
        '母板号不为空时，检查与上一母板号是否为相同母板号
            If ss1.Text <> mOplate_No Then
            '如果是不同母板号，而且颜色标记为1，那么颜色标记改为2，表示改变颜色
                If iColor = 1 Then
                   iColor = 2
                   '如果母板号相同，那么颜色标记还为1，表示颜色不变
                Else
                   iColor = 1
                End If
            End If
       End If
       '用1表示颜色置为浅灰色，用2表示颜色置为白色
       '每次循环结束，如果iColor为1，则颜色为浅灰色，否则颜色为白色
       If iColor = 1 Then
         '取样颜色改变，如果为Y表示取样，则该行字体颜色变红色
          ss1.Col = SS1_PLAN_SMP
          If ss1.Text = "Y" Then
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, &HFF&, &HE0E0E0) '浅灰色
          Else
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, , &HE0E0E0) '浅灰色
          End If
       Else
          '取样颜色改变，如果为Y表示取样，则该行字体颜色变红色
          ss1.Col = SS1_PLAN_SMP
          If ss1.Text = "Y" Then
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, &HFF&, &HFFFFFF) '白
          Else
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, , &HFFFFFF) '白
          End If
       End If
       '把值还原为for循环中母板号的取值
       ss1.Col = SS1_MOTHER_NO
       mOplate_No = ss1.Text
       
       ss1.ROW = lRow:          ss1.Col = SS1_ORD_CNT:          sord_cnt = Val(ss1.Text)
            If sord_cnt > 1 Then
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRow, lRow, , SSP1.BackColor)
            End If
        
        '热处理指示蓝色显示
        ss1.ROW = lRow:          ss1.Col = SS1_HTM_METH:          sHtm_Meth = Val(ss1.Text)
        If Mid(sHtm_Meth, 1, 1) = "N" And Mid(sHtm_Meth, 1, 1) <> "/ / /" Then
'        If sHtm_Meth <> "/ / /" Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRow, lRow, , SSP2.BackColor)
        End If
       
    Next lRow

End Sub

Private Sub Gp_CGC2071C_Excel(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

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
    
    Const xlCenter = -4108
    Const xlNone = -4142
    Const xlAutomatic = -4105
    Const xlDiagonalDown = 5
    Const xlDiagonalUp = 6
    Const xlEdgeLeft = 7
    Const xlEdgeTop = 8
    Const xlEdgeBottom = 9
    Const xlEdgeRight = 10
    Const xlInsideVertical = 11
    Const xlInsideHorizontal = 12
    Const xlContinuous = 1
    Const xlMedium = -4138
    Const xlThick = 4
    Const xlthin = 2
    
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
        

        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "@"
        
        ss1.ROW = ss1.ActiveRow
        
        xlSheet.Range("A1").Value = "日期"
        xlSheet.Range("A2").Value = "客户"
        xlSheet.Range("A3").Value = "产品号"
        xlSheet.Range("A4").Value = "厚度"
        xlSheet.Range("A5").Value = "宽度"
        xlSheet.Range("A6").Value = "长度"
        xlSheet.Range("A7").Value = "切边"
        xlSheet.Range("A8").Value = "厚度公差"
        xlSheet.Range("A9").Value = "订单备注"
        xlSheet.Range("A10").Value = "探伤"
        xlSheet.Range("A11").Value = "切割"
        xlSheet.Range("A12").Value = "矫直"
        xlSheet.Range("A13").Value = "热处理"
        xlSheet.Range("A14").Value = "其它"
        xlSheet.Range("A15").Value = "标识标准"
        xlSheet.Range("A16").Value = "钢种"
     
        xlSheet.Range("C1").Value = "班次"
        xlSheet.Range("C2").Value = "分断号"
        xlSheet.Range("C3").Value = "轧批号"
        xlSheet.Range("C5").Value = "母板长"
        xlSheet.Range("C7").Value = "定尺"
        xlSheet.Range("C8").Value = "长度公差"
        xlSheet.Range("C15").Value = "子公司代码"
        xlSheet.Range("C16").Value = "客户代码"
 
        ss1.Col = SS1_MV_DATE:           xlSheet.Range("B1").Value = ss1.Text
        ss1.Col = SS1_CUST_NAME:           xlSheet.Range("B2").Value = ss1.Text
        ss1.Col = SS1_PLATE_NO:           xlSheet.Range("B3").Value = ss1.Text
        ss1.Col = SS1_ORD_THK:           xlSheet.Range("B4").Value = ss1.Text
        ss1.Col = SS1_ORD_WID:           xlSheet.Range("B5").Value = ss1.Text
        ss1.Col = SS1_ORD_LEN:           xlSheet.Range("B6").Value = ss1.Text
        ss1.Col = SS1_TRIM_FL:           xlSheet.Range("B7").Value = ss1.Text
        ss1.Col = SS1_THK_LIM:           xlSheet.Range("B8").Value = ss1.Text
        ss1.Col = SS1_ORD_REMARK:        xlSheet.Range("B9").Value = ss1.Text
        ss1.Col = SS1_UST_STATUS:        xlSheet.Range("B10").Value = ss1.Text
        ss1.Col = SS1_GAS_STATUS:        xlSheet.Range("B11").Value = ss1.Text
        ss1.Col = SS1_CL_STATUS:         xlSheet.Range("B12").Value = ss1.Text
        ss1.Col = SS1_HTM_METH:          xlSheet.Range("B13").Value = ss1.Text
        ss1.Col = SS1_QT:                xlSheet.Range("B14").Value = ss1.Text
        ss1.Col = SS1_STDSPEC_STLGRD:   xlSheet.Range("B15").Value = ss1.Text
        ss1.Col = SS1_STDSPEC_ORG_KND:    xlSheet.Range("B16").Value = ss1.Text
        ss1.Col = SS1_SHIFT:             xlSheet.Range("D1").Value = ss1.Text
        ss1.Col = SS1_TRNS_CMPY_CD:      xlSheet.Range("D2").Value = ss1.Text
        ss1.Col = SS1_OUT_SHEET_NO:      xlSheet.Range("D3").Value = ss1.Text
        ss1.Col = SS1_LEN:               xlSheet.Range("D5").Value = ss1.Text
        ss1.Col = SS1_SIZE_KND:          xlSheet.Range("D7").Value = ss1.Text
        ss1.Col = SS1_LEN_LIM:           xlSheet.Range("D8").Value = ss1.Text
        ss1.Col = SS1_CUST_CD:         xlSheet.Range("D16").Value = ss1.Text
        ss1.Col = SS1_CD_MANA_NO:        xlSheet.Range("D15").Value = ss1.Text
        
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        xlApp.Range("A1:D16").Select
        xlApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        xlApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With xlApp.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
        With xlApp.Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlthin
        End With
            
            
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

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
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
    TXT_MAT_NO.Text = ss1.Text
    

    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    ss2.OperationMode = OperationModeNormal
    TXT_MAT_NO.Text = ""
    

End Sub


