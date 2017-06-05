VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGC2070C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "母板分产线处理作业_CGC2070C"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_GAS_FL 
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
      ItemData        =   "CGC2070C.frx":0000
      Left            =   13155
      List            =   "CGC2070C.frx":0002
      TabIndex        =   14
      Top             =   120
      Width           =   1965
   End
   Begin VB.ComboBox CBO_NUM 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "CGC2070C.frx":0004
      Left            =   7860
      List            =   "CGC2070C.frx":0006
      TabIndex        =   10
      Top             =   450
      Width           =   1125
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
      ItemData        =   "CGC2070C.frx":0008
      Left            =   6945
      List            =   "CGC2070C.frx":000A
      TabIndex        =   6
      Tag             =   "班别"
      Top             =   120
      Width           =   855
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
      ItemData        =   "CGC2070C.frx":000C
      Left            =   6090
      List            =   "CGC2070C.frx":000E
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txt_tmpseq 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   15390
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   150
      Visible         =   0   'False
      Width           =   1425
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
      TabIndex        =   0
      Top             =   540
      Width           =   2160
   End
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
      Left            =   6090
      MaxLength       =   12
      TabIndex        =   1
      Top             =   540
      Width           =   870
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8145
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   14367
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
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
      MaxCols         =   28
      MaxRows         =   10
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGC2070C.frx":0010
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
      Left            =   4770
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
   Begin Threed.SSCheck SSCHK_GAS_FL 
      Height          =   360
      Left            =   15360
      TabIndex        =   3
      Top             =   870
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   "不显示已有火切指示的母板"
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   4770
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
      TabIndex        =   7
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
      TabIndex        =   8
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
   Begin InDate.ULabel ULabel42 
      Height          =   345
      Left            =   9000
      Top             =   450
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      Caption         =   "块内"
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
   Begin InDate.ULabel ULabel43 
      Height          =   315
      Left            =   7860
      Top             =   120
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      Caption         =   "已分线母板"
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
   Begin Threed.SSPanel SSP1 
      Height          =   315
      Left            =   11160
      TabIndex        =   11
      Top             =   540
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "已选择"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP2 
      Height          =   315
      Left            =   12480
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
      Caption         =   "已分线"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP3 
      Height          =   315
      Left            =   13800
      TabIndex        =   13
      Top             =   540
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   8454016
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   11160
      Top             =   120
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      Caption         =   "是否有火切指示"
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
   Begin Threed.SSPanel SSP4 
      Height          =   315
      Left            =   10110
      TabIndex        =   15
      Top             =   540
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "重点订单"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   3015
      TabIndex        =   9
      Top             =   240
      Width           =   195
   End
End
Attribute VB_Name = "CGC2070C"
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
'-- Program Name      钢板实绩查询界面
'-- Program ID        AGC2200C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
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

Const SS1_SMP_YN = 1
Const SS1_MPLATE_NO = 2
Const SS1_TRNS_CMPY_CD = 3
Const SS1_LINE1 = 14
Const SS1_LINE2 = 15
Const SS1_LINE3 = 16
Const SS1_LINE4 = 17
Const SS1_OFFLINE_DATE = 18
Const SS1_USERID = 23
Const SS1_PLAN_SMP = 24
Const SS1_PRC_LINE = 25
Const SS1_ORD_CNT = 26        '一坯多订单  2011-08-18  by  LiQian
Const SS1_URGNT_FL = 27       '紧急订单绿色标记 2012-08-16  by  LiQian
Const SS1_IMP_CONT = 28




Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_SEQ, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_GAS_FL, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(CBO_NUM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
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
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    ' 成品宽度  2011-08-18  by  LiQian
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    ' 成品长度  2011-08-18  by  LiQian
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    ' 一坯多订单  2011-08-18  by  LiQian
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)  '紧急订单绿色标记 2012-08-16  by  LiQian
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)  '计划切边  2012-09-14  by  LiQian
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGC2070C.P_SREFER", Key:="P-R"
    sc1.Add Item:="CGC2070C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:="CGC2070C.P_SONEROW", Key:="P-O"
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

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

   If Col = SS1_SMP_YN Or Col = SS1_LINE1 Or Col = SS1_LINE2 Or Col = SS1_LINE3 Or Col = SS1_LINE4 Then
        If Gf_Sc_Authority(sAuthority, "U") Then
             Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), 0)
             ss1.ROW = ss1.ActiveRow:        ss1.Col = SS1_USERID:        ss1.Text = sUserID
             ss1.Col = SS1_LINE1
             If ss1.Value = 0 Then
                ss1.Col = SS1_LINE2
                 If ss1.Value = 0 Then
                     ss1.Col = SS1_LINE3
                     If ss1.Value = 0 Then
                        ss1.Col = SS1_LINE4
                         If ss1.Value = 0 Then
                            ss1.Col = SS1_SMP_YN
                             If ss1.Value = 0 Then
                                ss1.Col = 0
                                ss1.Text = ""
                             End If
                         End If
                     End If
                 End If
             End If

        End If
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
    
    CBO_NUM.AddItem "2"
    CBO_NUM.AddItem "4"
    CBO_NUM.AddItem "6"
    CBO_NUM.AddItem "8"
    
    CBO_GAS_FL.AddItem ""
    CBO_GAS_FL.AddItem "Y 有火切指示"
    CBO_GAS_FL.AddItem "N 无火切指示"
    
    
    
'    Call Gp_Sp_ColHidden(ss1, SS1_LINE3, True)
    Call Gp_Sp_ColHidden(ss1, SS1_SMP_YN, True)
    Call Gp_Sp_ColHidden(ss1, SS1_OFFLINE_DATE, True)
    
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
    Dim sSmp_color As Variant
    Dim sCnt_color As Variant
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Nothing) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    For lRow = 1 To ss1.MaxRows
    
        ' 一坯多订单,字体显示蓝色  2011-08-18  by  LiQian
        ss1.ROW = lRow:       ss1.Col = SS1_ORD_CNT
        If ss1.Text <> "" Then
            If ss1.Text > "1" Then
               sCnt_color = &HFF0000
            Else
               sCnt_color = vbBlack
            End If
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, sCnt_color)
        End If
    
        ss1.ROW = lRow:       ss1.Col = SS1_PRC_LINE
        If ss1.Text <> "X" Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, , SSP2.BackColor)
        End If
        
        ss1.ROW = lRow:       ss1.Col = SS1_PLAN_SMP
        If ss1.Text <> "" Then
            If ss1.Text = "Y" Then
               sSmp_color = &HFF&
            Else
               sSmp_color = vbBlack
            End If
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, sSmp_color, SSP3.BackColor)
        End If
         '紧急订单绿色标记 2012-08-16  by  LiQian
        ss1.ROW = lRow:       ss1.Col = SS1_URGNT_FL
        If ss1.Text = "Y" Then
             Call Gp_Sp_BlockColor(ss1, SS1_MPLATE_NO, SS1_MPLATE_NO, lRow, lRow, &HC000&)
             Call Gp_Sp_BlockColor(ss1, SS1_TRNS_CMPY_CD, SS1_TRNS_CMPY_CD, lRow, lRow, &HC000&)
             Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, lRow, lRow, &HC000&)
        End If
        
        ss1.ROW = lRow:       ss1.Col = SS1_IMP_CONT
        
        If ss1.Text = "Y" Then
            Call Gp_Sp_BlockColor(ss1, SS1_MPLATE_NO, SS1_MPLATE_NO, lRow, lRow, SSP4.BackColor)
            Call Gp_Sp_BlockColor(ss1, SS1_IMP_CONT, SS1_IMP_CONT, lRow, lRow, SSP4.BackColor)
        End If
        
    Next lRow

End Sub
Public Sub Form_Pro()

    Dim iCount      As Integer
    Dim sPlateNo    As String
    
    Dim inum As Integer
    Dim lRow As Double
    Dim sSmp_color As Variant
    Dim sCnt_color As Variant
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.ROW = IIf(Val(txt_tmpseq.Text) = 0, 1, Val(txt_tmpseq.Text))
        ss1.SetActiveCell ss1.ActiveCol, ss1.ROW
    End If
    
    For lRow = 1 To ss1.MaxRows
    
        ss1.ROW = lRow:       ss1.Col = SS1_PRC_LINE
        If ss1.Text <> "X" Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, , SSP2.BackColor)
        End If
        
        ss1.ROW = lRow:       ss1.Col = SS1_PLAN_SMP
        If ss1.Text <> "" Then
            If ss1.Text = "Y" Then
               sSmp_color = &HFF&
            Else
               sSmp_color = vbBlack
            End If
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, sSmp_color, SSP3.BackColor)
        End If
        
         ' 一坯多订单,字体显示蓝色  2011-08-18  by  LiQian
        ss1.ROW = lRow:       ss1.Col = SS1_ORD_CNT
        If ss1.Text <> "" Then
            If ss1.Text > "1" Then
               sCnt_color = &HFF0000
            Else
               sCnt_color = vbBlack
            End If
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, sCnt_color)
        End If
        
         '紧急订单绿色标记 2012-08-16  by  LiQian
        ss1.ROW = lRow:       ss1.Col = SS1_URGNT_FL
        If ss1.Text = "Y" Then
             Call Gp_Sp_BlockColor(ss1, SS1_MPLATE_NO, SS1_MPLATE_NO, lRow, lRow, &HC000&)
             Call Gp_Sp_BlockColor(ss1, SS1_TRNS_CMPY_CD, SS1_TRNS_CMPY_CD, lRow, lRow, &HC000&)
             Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, lRow, lRow, &HC000&)
        End If
        
    Next lRow
    
End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub
Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    If ss1.MaxRows < 1 Then Exit Sub

    If ROW = 0 Then
    
        Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)

        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0

    ElseIf (Col = SS1_SMP_YN Or Col = SS1_LINE1 Or Col = SS1_LINE2 Or Col = SS1_LINE3 Or Col = SS1_LINE4) Then

        ss1.ROW = ss1.ActiveRow
        ss1.Col = Col
        ss1.EditMode = True
        If ss1.Value = 0 Then
           ss1.Value = 1
        Else
           ss1.Value = 0
        End If
        
        ss1.Col = Col
        If ss1.Value = 1 Then
           If Col = SS1_LINE1 Then
              ss1.Col = SS1_LINE2
              ss1.Value = 0
              ss1.Col = SS1_LINE3
              ss1.Value = 0
              ss1.Col = SS1_LINE4
              ss1.Value = 0
              ss1.Col = SS1_OFFLINE_DATE
              ss1.Text = ""
           ElseIf Col = SS1_LINE2 Then
              ss1.Col = SS1_LINE1
              ss1.Value = 0
              ss1.Col = SS1_LINE3
              ss1.Value = 0
              ss1.Col = SS1_LINE4
              ss1.Value = 0
              ss1.Col = SS1_OFFLINE_DATE
              ss1.Text = ""
           ElseIf Col = SS1_LINE3 Then
              ss1.Col = SS1_LINE1
              ss1.Value = 0
              ss1.Col = SS1_LINE2
              ss1.Value = 0
              ss1.Col = SS1_LINE4
              ss1.Value = 0
           ElseIf Col = SS1_LINE4 Then
              ss1.Col = SS1_LINE1
              ss1.Value = 0
              ss1.Col = SS1_LINE2
              ss1.Value = 0
              ss1.Col = SS1_LINE3
              ss1.Value = 0
           End If
        Else
            ss1.Col = 0
            ss1.Text = ""
        End If
        
        ss1.Col = SS1_USERID
        ss1.Text = sUserID
        
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, , SSP1.BackColor)
        
        txt_tmpseq.Text = ss1.ActiveRow

    End If

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub SSCHK_GAS_FL_Click(Value As Integer)
    If SSCHK_GAS_FL.Value = ssCBUnchecked Then
       SSCHK_GAS_FL.ForeColor = &H808080
    Else
       SSCHK_GAS_FL.ForeColor = &HFF&
    End If
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Sp_Setting
'   2.Name         : Spread initialize Setting
'   3.Input  Value : sPname Variant, {MsgChk Boolean}
'   4.Return Value :
'   5.Writer       : 杨猛
'   6.Create Date  : 2010. 12 .21
'   7.Modify Date  :
'   8.Comment      : Spread initialize Setting
'---------------------------------------------------------------------------------------
Public Sub Sp_Setting(ByVal sPname As Variant, Optional MsgChk As Boolean = True)

    With sPname
    
        .RowHeight(-1) = 15 '12.54
        
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
     
        .OperationMode = OperationModeRow
        .RetainSelBlock = True

        .UserResize = UserResizeColumns
        .AllowDragDrop = False
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .ROW = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 12
        .BlockMode = False
        
        .Col = -1
        .ROW = 0
        .FontBold = True
        
        'If .ColHeaderRows > 1 Then
        '    .Row = SpreadHeader + 1
        '    .FontBold = True
        'End If
        
        If MsgChk Then
            .LockBackColor = RGB(255, 255, 255)
        End If
        
        .MaxRows = 0
                
    End With
    
End Sub


