VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form DGA1210C 
   Caption         =   "钢板剩磁检查实绩查询及修改_DGA1210C"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame2 
      Height          =   1290
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   2275
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_STDSPEC_CHG 
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
         Left            =   1380
         TabIndex        =   13
         Top             =   480
         Width           =   3765
      End
      Begin VB.TextBox TXT_ORD_NO 
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
         Left            =   6720
         TabIndex        =   12
         Top             =   870
         Width           =   1695
      End
      Begin VB.ComboBox cbo_chg_no 
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
         ItemData        =   "DGA1210C.frx":0000
         Left            =   6720
         List            =   "DGA1210C.frx":0002
         TabIndex        =   11
         Tag             =   "炉座号"
         Top             =   480
         Width           =   1365
      End
      Begin VB.TextBox txt_f_addr 
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
         Left            =   10035
         TabIndex        =   9
         Tag             =   "标准代码"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox CBO_PRODGRD 
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
         ItemData        =   "DGA1210C.frx":0004
         Left            =   11400
         List            =   "DGA1210C.frx":001A
         TabIndex        =   5
         Tag             =   "等级"
         Top             =   90
         Width           =   1365
      End
      Begin VB.TextBox TXT_PLATE_NO 
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
         Left            =   1380
         TabIndex        =   4
         Top             =   870
         Width           =   1695
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
         ItemData        =   "DGA1210C.frx":0054
         Left            =   7620
         List            =   "DGA1210C.frx":0064
         TabIndex        =   3
         Tag             =   "班别"
         Top             =   90
         Width           =   885
      End
      Begin VB.ComboBox CBO_SURFGRD 
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
         ItemData        =   "DGA1210C.frx":0074
         Left            =   10035
         List            =   "DGA1210C.frx":008D
         TabIndex        =   2
         Tag             =   "等级"
         Top             =   90
         Width           =   1365
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
         ItemData        =   "DGA1210C.frx":00D0
         Left            =   6720
         List            =   "DGA1210C.frx":00DD
         TabIndex        =   1
         Top             =   90
         Width           =   885
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   5475
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   1
         Left            =   120
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "标准号"
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   120
         Top             =   870
         Width           =   1215
         _ExtentX        =   2143
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   8790
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "表面/综合"
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
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   120
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "热处理时间"
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
      Begin CSTextLibCtl.sitxEdit SDT_PROD_DATE_FROM 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Tag             =   "探伤日期"
         Top             =   90
         Width           =   1800
         _Version        =   262145
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __:__"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   "____-__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin CSTextLibCtl.sitxEdit SDT_PROD_DATE_TO 
         Height          =   315
         Left            =   3345
         TabIndex        =   8
         Tag             =   "探伤日期"
         Top             =   90
         Width           =   1800
         _Version        =   262145
         _ExtentX        =   3175
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderEffect    =   2
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__ __:__"
         StartText.x     =   3
         StartText.y     =   3
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   15
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   "____-__-__ __:__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   2
         Left            =   8790
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "垛位号"
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
         Left            =   5475
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "炉座号"
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
      Begin InDate.ULabel ULabel22 
         Height          =   315
         Index           =   0
         Left            =   5475
         Top             =   870
         Width           =   1215
         _ExtentX        =   2143
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
      End
      Begin Threed.SSCommand cmd_Exc 
         Height          =   345
         Left            =   8790
         TabIndex        =   14
         Top             =   870
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
         _Version        =   196609
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "记录导出"
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   180
         Index           =   0
         Left            =   3150
         TabIndex        =   10
         Top             =   210
         Width           =   240
      End
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8145
      Left            =   120
      TabIndex        =   6
      Top             =   1350
      Width           =   15255
      _Version        =   393216
      _ExtentX        =   26908
      _ExtentY        =   14367
      _StockProps     =   64
      ColsFrozen      =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   65
      MaxRows         =   5
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "DGA1210C.frx":00ED
   End
End
Attribute VB_Name = "DGA1210C"
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
'-- Program Name      钢板剩磁检查实绩查询及修改
'-- Program ID        DGA1210C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2010.11.24
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER     DATE              EDITOR      DESCRIPTION
'-- 1.01    2010.11.24        Yang Meng
'-- 1.02    2012.04.25        Li Qian     画面新增录入项目
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

Dim Mc1 As New Collection           'Master Collection
'Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_PLATE_NO = 1
Const SPD_CHARGE_FUR_LINE = 2
Const SPD_DIS_CHARGE_DATE = 3
Const SPD_PROC_CD = 4
Const SPD_APLY_STDSPEC = 5
Const SPD_THK = 6
Const SPD_WID = 7
Const SPD_LEN = 8
Const SPD_WGT = 9
Const SPD_INSP_AVE_WID = 10
Const SPD_INSP_AVE_LEN = 11
Const SPD_THICKNESS1 = 12
Const SPD_THICKNESS2 = 13
Const SPD_THICKNESS3 = 14
Const SPD_THICKNESS4 = 15
Const SPD_THICKNESS5 = 16
Const SPD_THICKNESS6 = 17
Const SPD_THICKNESS7 = 18
Const SPD_THICKNESS8 = 19
Const SPD_PROD_GRD = 22
Const SPD_SURF_GRD = 23
Const SPD_MAGNET_MIN = 24
Const SPD_MAGNET_MAX = 25
Const SPD_MAGNET_GRD = 26
Const SPD_MAGNET1 = 27
Const SPD_MAGNET2 = 28
Const SPD_MAGNET3 = 29
Const SPD_MAGNET4 = 30
Const SPD_MAGNET5 = 31
Const SPD_MAGNET6 = 32
Const SPD_MAGNET7 = 33
Const SPD_MAGNET8 = 34
Const SPD_INSP_WAVE = 37
Const SPD_INSP_T_FLAW = 38
Const SPD_INSP_B_FLAW = 39
Const SPD_AXIS_X = 40
Const SPD_AXIS_Y = 41
Const SPD_REMAIN_THK = 42
Const SPD_ROUGHNESS = 43    '表面粗糙度
Const SPD_ROUGHNESS_NAME = 44
Const SPD_ISREVERSE = 45
Const SPD_ISREVERSE1 = 46
Const SPD_DISPOSE = 47      '处置方式
Const SPD_DISPOSE_NAME = 48
Const SPD_CUR_INV = 49
Const SPD_LOC = 50
Const SPD_BED_PILE_DATE = 51
Const SPD_PROD_DATE = 52
Const SPD_GROUP = 53
Const SPD_SHIFT = 54
Const SPD_INSP_DATE = 55
Const SPD_INSP_MAN = 56
Const SPD_PROD_REMARK = 57
Const SPD_ORD = 58
Const SPD_UST_FL = 59
Const SPD_GAS_FL = 60
Const SPD_CL_FL = 61
Const SPD_HTM_METH = 62
Const SPD_MAGNET_DATE = 63
Const SPD_USERID = 64


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(CBO_SURFGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(CBO_PRODGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_STDSPEC_CHG, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_f_addr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(cbo_chg_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
        Mc1.Add Item:=pControl, Key:="pControl"
        Mc1.Add Item:=nControl, Key:="nControl"
        Mc1.Add Item:=mControl, Key:="mControl"
        Mc1.Add Item:=iControl, Key:="iControl"
        Mc1.Add Item:=rControl, Key:="rControl"
        Mc1.Add Item:=cControl, Key:="cControl"
        Mc1.Add Item:=aControl, Key:="aControl"
        Mc1.Add Item:=lControl, Key:="lControl"
        
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '实测宽度
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '实测长度
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度1
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度2
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度3
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度4
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度5
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度6
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度7
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度8
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度最大值
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '厚度最小值
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '不平度
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '上表缺陷
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '下表缺陷
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) 'X-轴
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) 'Y-轴
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '剩余厚度
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '表面粗糙度
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '是否翻板
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '是否翻板
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '处置方式
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '生产班别
    Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '生产班次
    Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '检验时间
    Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '检验人员
    Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '备注
    Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)

    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="DGA1210C.P_REFER", Key:="P-R"
    sc1.Add Item:="DGA1210C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="DGA1210C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
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

Private Sub cmd_Exc_Click()
    Call Gp_Sp_Excel_Re(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        ss1.Col = SPD_USERID:     ss1.Text = sUserID
    End If
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

  If KeyCode = vbKeyF4 Then
        With ss1
            .Col = .ActiveCol
            .Row = .ActiveRow
            If .ActiveCol = SPD_ROUGHNESS Then
                Set DD.sPname = Me.ss1
                DD.sWitch = "SP"
                DD.sKey = "BG003"
                DD.rControl.Add Item:=SPD_ROUGHNESS
                DD.nameType = "1"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            End If
            If .ActiveCol = SPD_DISPOSE Then
                Set DD.sPname = Me.ss1
                DD.sWitch = "SP"
                DD.sKey = "BG004"
                DD.rControl.Add Item:=SPD_DISPOSE
                DD.nameType = "1"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            End If
        End With
   End If
    
End Sub


Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
'        If Len(TXT_PLATE_NO.Text) >= 8 Then
'           Call Form_Ref
'        End If
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

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_ColHidden(ss1, SPD_ISREVERSE, True)
    
'    SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
'    SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
   
'    SDT_PROD_DATE_FROM.RawData = Format(Now, "yyyymmdd") & "0000"
'    SDT_PROD_DATE_TO.RawData = Format(Now, "yyyymmdd") & "2400"

  '  Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 39)

    cbo_chg_no.AddItem "1"
    cbo_chg_no.AddItem "2"
    
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
    Set sc1 = Nothing
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
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If

End Sub

Public Sub Master_Cpy()

'    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

'     If Gf_Ms_Paste(M_CN1, Mc1) Then
'        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
'     End If

End Sub

Public Sub Form_Ref()
    
    Dim sMsg As String
    Dim I As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If SDT_PROD_DATE_FROM.RawData = "" Then
       SDT_PROD_DATE_FROM.RawData = Format(Now, "yyyymmdd") & "0000"
    End If

    If SDT_PROD_DATE_TO.RawData = "" Then
       SDT_PROD_DATE_TO.RawData = Format(Now, "yyyymmdd") & "2400"
    End If

    If Val(SDT_PROD_DATE_FROM.RawData) - Val(SDT_PROD_DATE_TO.RawData) > 0 Then
         sMsg = " 时间范围输入错误，请重新输入时间信息 ！！！"
         Call Gp_MsgBoxDisplay(sMsg)
         Exit Sub
    End If
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        For iRow = 1 To ss1.MaxRows
    
               ss1.Row = iRow
               ss1.Col = 65
                If ss1.Text = "Y" Then
                  For I = 1 To ss1.MaxCols
                       ss1.Col = I
                       ss1.ForeColor = &HC000&
                  Next
                End If
        
         Next iRow
        
      
    End If

End Sub
Public Sub Form_Pro()

    Dim iCount      As Integer
    Dim sPlateNo    As String
    Dim I As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    
    Dim inum As Integer
    Dim lRow As Integer
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    For iRow = 1 To ss1.MaxRows
    
               ss1.Row = iRow
               ss1.Col = 65
                If ss1.Text = "Y" Then
                  For I = 1 To ss1.MaxCols
                       ss1.Col = I
                       ss1.ForeColor = &HC000&
                  Next
                End If
    Next iRow
        
    
End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

'Private Sub SDT_PROD_DATE_FROM_GotFocus()
'     If SDT_PROD_DATE_FROM.RawData = "" Then
'        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
'     End If
'     If SDT_PROD_DATE_TO.RawData = "" Then
'        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
'     End If
'End Sub
'
'Private Sub SDT_PROD_DATE_TO_GotFocus()
'     If SDT_PROD_DATE_TO.RawData = "" Then
'        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
'     End If
'End Sub
Private Sub SDT_PROD_DATE_FROM_DblClick()
     SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D") + "0000"
     SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "I")
End Sub

Private Sub SDT_PROD_DATE_TO_DblClick()
     SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "I")
End Sub

Private Sub txt_stdspec_chg_DblClick()
    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STDSPEC_CHG

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
    
End Sub
'Private Sub txt_f_addr_DblClick()
'     Call txt_f_addr_KeyUp(vbKeyF4, 0)
'End Sub
'
'Private Sub txt_f_addr_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "F0009"
'        txt_f_addr.Text = "P%R"
'        DD.rControl.Add Item:=txt_f_addr
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'
'End Sub
'---------------------------------------------------------------------------------------
'   1.ID           : Gp_Sp_Excel
'   2.Name         : Spread --> Excel
'   3.Input  Value : Fm Form, sPname Variant, bLkcol1 Long, bLkcol2 Long, bLkrow1 Long, bLkrow2 Long
'   4.Return Value :
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  : 2012.04.26   Li Qian
'   8.Comment      : Spread --> Excel
'---------------------------------------------------------------------------------------
Public Sub Gp_Sp_Excel_Re(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

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
        
        Clipboard.Clear
        
        .Col = bLkcol1: .Col2 = bLkcol2
        .Row = bLkrow1: .Row2 = bLkrow2
        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
                        
        xlSheet.Cells.NumberFormatLocal = "G/通用格式"
        xlApp.Range("S1").Value = "  中厚板卷厂热处理线钢板检验记录  "
        xlApp.Range("AA2").Value = "质量记录编号： JL320217/B"
        
        sExlRange1 = ""
        For ColIndex = 1 To .MaxCols
            .Col = ColIndex
            .Row = 1

            iExlCol = ColIndex
'            If IsNumeric(.Text) And (Left(.Text, 1) = "0" Or Left(.Text, 1) = "1" Or Left(.Text, 1) = "7") And _
'               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14) Then
            If .CellType = SS_CELL_TYPE_EDIT Then
                If ColIndex > 104 Then
                    sExlRange1 = "D"
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C"
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B"
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If

                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14 Then
                     xlSheet.Range(sExlRange).NumberFormat = "@"
                End If
            End If
        Next
        
        ss1.Row = 0
        ss1.Col = SPD_PLATE_NO:            xlApp.Range("A3").Value = ss1.Text
        ss1.Col = SPD_APLY_STDSPEC:        xlApp.Range("B3").Value = ss1.Text
        ss1.Col = SPD_THK:                 xlApp.Range("C3").Value = ss1.Text
        ss1.Col = SPD_WID:                 xlApp.Range("D3").Value = ss1.Text
        ss1.Col = SPD_LEN:                 xlApp.Range("E3").Value = ss1.Text
        ss1.Col = SPD_WGT:                 xlApp.Range("F3").Value = ss1.Text
        ss1.Col = SPD_INSP_AVE_WID:        xlApp.Range("G3").Value = ss1.Text
                                           xlApp.Range("G4").Value = "测量宽度"
        ss1.Col = SPD_INSP_AVE_LEN:        xlApp.Range("H3").Value = ss1.Text
                                           xlApp.Range("H4").Value = "测量长度"
        ss1.Col = SPD_THICKNESS1:          xlApp.Range("I3").Value = ss1.Text
                                           xlApp.Range("I4").Value = "厚度1"
        ss1.Col = SPD_THICKNESS2:          xlApp.Range("J3").Value = ss1.Text
                                           xlApp.Range("J4").Value = "厚度2"
        ss1.Col = SPD_THICKNESS3:          xlApp.Range("K3").Value = ss1.Text
                                           xlApp.Range("K4").Value = "厚度3"
        ss1.Col = SPD_THICKNESS4:          xlApp.Range("L3").Value = ss1.Text
                                           xlApp.Range("L4").Value = "厚度4"
        ss1.Col = SPD_THICKNESS5:          xlApp.Range("M3").Value = ss1.Text
                                           xlApp.Range("M4").Value = "厚度5"
        ss1.Col = SPD_THICKNESS6:          xlApp.Range("N3").Value = ss1.Text
                                           xlApp.Range("N4").Value = "厚度6"
        ss1.Col = SPD_THICKNESS7:          xlApp.Range("O3").Value = ss1.Text
                                           xlApp.Range("O4").Value = "厚度7"
        ss1.Col = SPD_THICKNESS8:          xlApp.Range("P3").Value = ss1.Text
                                           xlApp.Range("P4").Value = "厚度8"
        ss1.Col = SPD_MAGNET1:             xlApp.Range("Q3").Value = ss1.Text
                                           xlApp.Range("Q4").Value = "测点1"
        ss1.Col = SPD_MAGNET2:             xlApp.Range("R3").Value = ss1.Text
                                           xlApp.Range("R4").Value = "测点2"
        ss1.Col = SPD_MAGNET3:             xlApp.Range("S3").Value = ss1.Text
                                           xlApp.Range("S4").Value = "测点3"
        ss1.Col = SPD_MAGNET4:             xlApp.Range("T3").Value = ss1.Text
                                           xlApp.Range("T4").Value = "测点4"
        ss1.Col = SPD_MAGNET5:             xlApp.Range("U3").Value = ss1.Text
                                           xlApp.Range("U4").Value = "测点5"
        ss1.Col = SPD_MAGNET6:             xlApp.Range("V3").Value = ss1.Text
                                           xlApp.Range("V4").Value = "测点6"
        ss1.Col = SPD_MAGNET7:             xlApp.Range("W3").Value = ss1.Text
                                           xlApp.Range("W4").Value = "测点7"
        ss1.Col = SPD_MAGNET8:             xlApp.Range("X3").Value = ss1.Text
                                           xlApp.Range("X4").Value = "测点8"
        ss1.Col = SPD_INSP_WAVE:           xlApp.Range("Y3").Value = ss1.Text
                                           xlApp.Range("Y4").Value = "不平度(mm/2m)"
        ss1.Col = SPD_INSP_T_FLAW:         xlApp.Range("Z3").Value = ss1.Text
                                           xlApp.Range("Z4").Value = "上表缺陷"
        ss1.Col = SPD_INSP_B_FLAW:         xlApp.Range("AA3").Value = ss1.Text
                                           xlApp.Range("AA4").Value = "下表缺陷"
        ss1.Col = SPD_AXIS_X:              xlApp.Range("AB3").Value = ss1.Text
                                           xlApp.Range("AB4").Value = "X-轴"
        ss1.Col = SPD_AXIS_Y:              xlApp.Range("AC3").Value = ss1.Text
                                           xlApp.Range("AC4").Value = "Y-轴"
        ss1.Col = SPD_REMAIN_THK:          xlApp.Range("AD3").Value = ss1.Text
                                           xlApp.Range("AD4").Value = "剩余厚度"
        ss1.Col = SPD_ROUGHNESS_NAME:           xlApp.Range("AE3").Value = ss1.Text
                                           xlApp.Range("AE4").Value = "表面粗糙度"
        ss1.Col = SPD_ISREVERSE:           xlApp.Range("AF3").Value = ss1.Text
                                           xlApp.Range("AF4").Value = "是否翻板"
                                           xlApp.Range("AG4").Value = "表面等级"
        ss1.Col = SPD_DISPOSE_NAME:             xlApp.Range("AH3").Value = ss1.Text
                                           xlApp.Range("AH4").Value = "处置方式"
        ss1.Col = SPD_GROUP:               xlApp.Range("AI3").Value = ss1.Text
        ss1.Col = SPD_SHIFT:               xlApp.Range("AJ3").Value = ss1.Text
        ss1.Col = SPD_INSP_DATE:           xlApp.Range("AK3").Value = ss1.Text
        ss1.Col = SPD_INSP_MAN:            xlApp.Range("AL3").Value = ss1.Text
        ss1.Col = SPD_PROD_REMARK:         xlApp.Range("AM3").Value = ss1.Text
        

        'xlSheet.Range("A1").Select
        'xlSheet.Paste
        Clipboard.Clear
        ss1.SetSelection SPD_PLATE_NO, 1, SPD_PLATE_NO, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("A5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_APLY_STDSPEC, 1, SPD_APLY_STDSPEC, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("B5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_THK, 1, SPD_THK, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("C5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_WID, 1, SPD_WID, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("D5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_LEN, 1, SPD_LEN, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("E5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_WGT, 1, SPD_WGT, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("F5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_INSP_AVE_WID, 1, SPD_INSP_AVE_WID, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("G5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_INSP_AVE_LEN, 1, SPD_INSP_AVE_LEN, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("H5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_THICKNESS1, 1, SPD_THICKNESS1, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("I5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_THICKNESS2, 1, SPD_THICKNESS2, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("J5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_THICKNESS3, 1, SPD_THICKNESS3, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("K5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_THICKNESS4, 1, SPD_THICKNESS4, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("L5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_THICKNESS5, 1, SPD_THICKNESS5, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("M5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_THICKNESS6, 1, SPD_THICKNESS6, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("N5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_THICKNESS7, 1, SPD_THICKNESS7, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("O5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_THICKNESS8, 1, SPD_THICKNESS8, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("P5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_MAGNET1, 1, SPD_MAGNET1, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("Q5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_MAGNET2, 1, SPD_MAGNET2, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("R5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_MAGNET3, 1, SPD_MAGNET3, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("S5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_MAGNET4, 1, SPD_MAGNET4, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("T5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_MAGNET5, 1, SPD_MAGNET5, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("U5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_MAGNET6, 1, SPD_MAGNET6, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("V5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_MAGNET7, 1, SPD_MAGNET7, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("W5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_MAGNET8, 1, SPD_MAGNET8, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("X5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_INSP_WAVE, 1, SPD_INSP_WAVE, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("Y5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_INSP_T_FLAW, 1, SPD_INSP_T_FLAW, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("Z5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_INSP_B_FLAW, 1, SPD_INSP_B_FLAW, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AA5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_AXIS_X, 1, SPD_AXIS_X, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AB5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_AXIS_Y, 1, SPD_AXIS_Y, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AC5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_REMAIN_THK, 1, SPD_REMAIN_THK, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AD5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_ROUGHNESS_NAME, 1, SPD_ROUGHNESS_NAME, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AE5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_ISREVERSE, 1, SPD_ISREVERSE, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AF5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

         Clipboard.Clear
        ss1.SetSelection SPD_SURF_GRD, 1, SPD_SURF_GRD, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AG5").Select
         xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_DISPOSE_NAME, 1, SPD_DISPOSE_NAME, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AH5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_GROUP, 1, SPD_GROUP, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AI5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_SHIFT, 1, SPD_SHIFT, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AJ5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear
        
        Clipboard.Clear
        ss1.SetSelection SPD_INSP_DATE, 1, SPD_INSP_DATE, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AK5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_INSP_MAN, 1, SPD_INSP_MAN, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AL5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        Clipboard.Clear
        ss1.SetSelection SPD_PROD_REMARK, 1, SPD_PROD_REMARK, ss1.MaxRows
        ss1.ClipboardCopy
        xlApp.Range("AM5").Select
        xlApp.ActiveSheet.Paste
        Clipboard.Clear

        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:
    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel", "W")

End Sub

