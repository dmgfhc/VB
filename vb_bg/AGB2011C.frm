VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGB2011C 
   Caption         =   "轧钢实绩查询_AGB2011C"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
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
      ItemData        =   "AGB2011C.frx":0000
      Left            =   11715
      List            =   "AGB2011C.frx":0010
      TabIndex        =   19
      Tag             =   "班别"
      Top             =   75
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
      Left            =   8595
      TabIndex        =   18
      Top             =   75
      Width           =   1665
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
      ItemData        =   "AGB2011C.frx":0020
      Left            =   6105
      List            =   "AGB2011C.frx":002D
      TabIndex        =   0
      Top             =   75
      Width           =   735
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   7290
      Top             =   75
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   90
      Top             =   75
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "生产时间"
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
      Left            =   5100
      Top             =   75
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8565
      Left            =   90
      TabIndex        =   14
      Top             =   720
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   15108
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   52
      MaxRows         =   20
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AGB2011C.frx":003A
   End
   Begin InDate.UDate SDT_PROD_DATE_FROM 
      Height          =   315
      Left            =   1395
      TabIndex        =   15
      Tag             =   "起始日期"
      Top             =   75
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
      Left            =   3150
      TabIndex        =   16
      Tag             =   "起始日期"
      Top             =   75
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
   Begin Threed.SSFrame SSFrame3 
      Height          =   870
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Visible         =   0   'False
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   1535
      _Version        =   196609
      Font3D          =   2
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TXT_NO_WGT 
         Alignment       =   1  'Right Justify
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
         Left            =   11085
         TabIndex        =   13
         Top             =   450
         Width           =   1215
      End
      Begin VB.TextBox TXT_NO_NUM 
         Alignment       =   1  'Right Justify
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
         Left            =   10185
         TabIndex        =   12
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox TXT_GP_WGT 
         Alignment       =   1  'Right Justify
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
         Left            =   6300
         TabIndex        =   11
         Top             =   450
         Width           =   1230
      End
      Begin VB.TextBox TXT_GP_NUM 
         Alignment       =   1  'Right Justify
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
         Left            =   5400
         TabIndex        =   10
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox TXT_ZP_WGT 
         Alignment       =   1  'Right Justify
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
         Left            =   3885
         TabIndex        =   9
         Top             =   450
         Width           =   1230
      End
      Begin VB.TextBox TXT_ZP_NUM 
         Alignment       =   1  'Right Justify
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
         Left            =   2970
         TabIndex        =   8
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox TXT_PLATE_WGT 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   7
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox TXT_PLATE_NUM 
         Alignment       =   1  'Right Justify
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
         Left            =   360
         TabIndex        =   6
         Top             =   450
         Width           =   990
      End
      Begin VB.TextBox TXT_XY_NUM 
         Alignment       =   1  'Right Justify
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
         Left            =   7800
         TabIndex        =   5
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox TXT_XY_WGT 
         Alignment       =   1  'Right Justify
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
         Left            =   8700
         TabIndex        =   4
         Top             =   450
         Width           =   1230
      End
      Begin VB.TextBox TXT_FG_NUM 
         Alignment       =   1  'Right Justify
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
         Left            =   12570
         TabIndex        =   3
         Top             =   450
         Width           =   900
      End
      Begin VB.TextBox TXT_FG_WGT 
         Alignment       =   1  'Right Justify
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
         Left            =   13470
         TabIndex        =   2
         Top             =   450
         Width           =   1215
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   360
         Top             =   120
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   556
         Caption         =   "钢板数量 | 总重量(Ton)"
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
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   2985
         Top             =   120
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         Caption         =   "正品量 | 重量(Ton)"
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
      Begin InDate.ULabel ULabel20 
         Height          =   315
         Left            =   5400
         Top             =   120
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         Caption         =   "改判量 | 重量(Ton)"
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
      Begin InDate.ULabel TXT_DP_NUM 
         Height          =   315
         Left            =   10185
         Top             =   120
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         Caption         =   "待判量 | 重量(Ton)"
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
         Left            =   7800
         Top             =   120
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         Caption         =   "协议量 | 重量(Ton)"
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
         Left            =   12570
         Top             =   120
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         Caption         =   "判废量 | 重量(Ton)"
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
   Begin InDate.ULabel ULabel39 
      Height          =   315
      Left            =   10710
      Top             =   75
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "班别"
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
   Begin Threed.SSPanel SSP1 
      Height          =   285
      Left            =   90
      TabIndex        =   20
      Top             =   450
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   503
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "无 母 板 实 绩"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP2 
      Height          =   285
      Left            =   7665
      TabIndex        =   21
      Top             =   450
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   503
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
      Caption         =   "轧 制 厚 度 超 公 差"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   2955
      TabIndex        =   17
      Top             =   195
      Width           =   195
   End
End
Attribute VB_Name = "AGB2011C"
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
'-- Program ID        AGB2011C
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

Const SPD_SLAB_NO = 1
Const SPD_GROUP = 4
Const SPD_THK_AVE = 16
Const SPD_SLAB_WGT = 22
Const SPD_COOL_RATE = 26
Const SPD_PLAN_PROD_WGT = 28
Const SPD_PLAN_PROD_RAT = 29
Const SPD_RST_PROD_WGT = 30
Const SPD_RST_PROD_RAT = 33
Const SPD_DISCHARGE_DATE = 35
Const SPD_RHFMILL_DATE = 36
Const SPD_MILLSTR_DATE = 37
Const SPD_MILL_DATE = 38
Const SPD_MILLEND_DATE = 39
Const SPD_DUR_DATE = 40
Const SPD_MP_CNT = 43
Const SPD_THK_MIN = 46
Const SPD_THK_MAX = 47
Const SPD_URGNT_FL = 48

Const SS1_FIRST_REMARK = 2


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_PLATE_NUM, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_PLATE_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_ZP_NUM, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_ZP_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_GP_NUM, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_GP_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_XY_NUM, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_XY_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_NO_NUM, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_NO_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_FG_NUM, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_FG_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

        Mc1.Add Item:=pControl, Key:="pControl"
        Mc1.Add Item:=nControl, Key:="nControl"
        Mc1.Add Item:=mControl, Key:="mControl"
        Mc1.Add Item:=iControl, Key:="iControl"
        Mc1.Add Item:=rControl, Key:="rControl"
        Mc1.Add Item:=cControl, Key:="cControl"
        Mc1.Add Item:=aControl, Key:="aControl"
        Mc1.Add Item:=lControl, Key:="lControl"
               
'            Call Gp_Ms_Collection(TXT_PLATE, "p", "n", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'       Call Gp_Ms_Collection(TXT_MARKING_YN, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'         Call Gp_Ms_Collection(TXT_STAMP_YN, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'           Call Gp_Ms_Collection(TXT_BAR_YN, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'           Call Gp_Ms_Collection(TXT_BND_YN, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'           Call Gp_Ms_Collection(TXT_ORD_NO, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'              Call Gp_Ms_Collection(SDB_THK, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'              Call Gp_Ms_Collection(SDB_WGT, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'           Call Gp_Ms_Collection(TXT_SMP_FL, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'          Call Gp_Ms_Collection(TXT_SMP_LOC, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'           Call Gp_Ms_Collection(TXT_SHP_DT, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'           Call Gp_Ms_Collection(TXT_STLGRD, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'              Call Gp_Ms_Collection(SDB_WID, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'    Call Gp_Ms_Collection(TXT_INSP_MAIN_GRD, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'           Call Gp_Ms_Collection(TXT_SMP_NO, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'          Call Gp_Ms_Collection(SDB_SMP_LEN, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'        Call Gp_Ms_Collection(TXT_YARD_ADDR, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'              Call Gp_Ms_Collection(SDB_LEN, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'         Call Gp_Ms_Collection(SDT_DEC_DATE, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'
'    'MASTER Collection
'
'     Mc2.Add Item:="AGC2200C.P_REFER1", Key:="P-R"
'     Mc2.Add Item:=pControl1, Key:="pControl"
'     Mc2.Add Item:=nControl1, Key:="nControl"
'     Mc2.Add Item:=mControl1, Key:="mControl"
'     Mc2.Add Item:=iControl1, Key:="iControl"
'     Mc2.Add Item:=rControl1, Key:="rControl"
'     Mc2.Add Item:=cControl1, Key:="cControl"
'     Mc2.Add Item:=aControl1, Key:="aControl"
'     Mc2.Add Item:=lControl1, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '首件标识
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '热卡量厚度
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '同板差
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)  'LICHAO 紧急订单 20121109
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    
   
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGB2011C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AGB2011C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AGB2011C.P_ONEROW", Key:="P-O"
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
     
     'sQuery_load
'    sQuery_load = "SELECT GOODS_ID FROM (SELECT B.GOODS_ID FROM FP_TRACKIDX A, FP_TRACKDATA B WHERE B.SEQ_NO <= A.LAST_SEQ AND A.FACT_CD = 'C1' " _
'    & "AND A.PRC = 'CC' AND A.PRC_LINE='1' AND A.FACT_CD=B.FACT_CD  AND A.PRC=B.PRC AND A.PRC_LINE=B.PRC_LINE ORDER BY B.SEQ_NO DESC) WHERE ROWNUM<=5"

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

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
'    Call Gf_ComboAdd(M_CN1, CBO_COIL_NO, sQuery_load)
    Call Gp_Sp_ColHidden(ss1, SPD_COOL_RATE, True)
    
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
'    Set Mc2 = Nothing
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
'       Call Gp_Ms_Cls(Mc2("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If

End Sub

Public Sub Master_Cpy()

'    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()
'
'     If Gf_Ms_Paste(M_CN1, Mc1) Then
'        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
'       ' Call Gp_Ms_ControlLock(Mc1("pControl"), False)
'     End If

End Sub

Public Sub Form_Ref()

    Dim iCount              As Integer
    Dim dMillCal_Wgt        As Double
    Dim dMillSlab_Wgt       As Double
    Dim dMillPlan_Wgt       As Double
    Dim dMillRst_Wgt        As Double
    
    Dim dMillSlab_Cnt       As Double
    Dim dMillSlab_SumWgt    As Double
    Dim dMillRst_SumWgt     As Double
    Dim dMillRstSlab_SumWgt As Double
    
    Dim dDischarge_date     As String
    Dim dMillstr_date       As String
    Dim dMill_date          As String
    Dim dMill_dur           As String    '轧制间隔
    Dim dMill_dtmin         As String    '纯轧时间
    Dim dRhfMill_dtmin      As String    '出炉到开轧时间
    Dim dMoplate_cnt        As Long      '出炉到开轧时间
    
    Dim dThk_Ave            As Double
    Dim dThk_Min            As Double
    Dim dThk_Max            As Double
    
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
    End If
    
    dMillCal_Wgt = 0
    With ss1
        If .MaxRows < 1 Then
           Exit Sub
        End If
        .MaxRows = .MaxRows + 1
        For iCount = 1 To .MaxRows - 1
            .Row = iCount
            '轧制块数
             dMillSlab_Cnt = dMillSlab_Cnt + 1
            '计划产品重量（累计）
            .Col = SPD_PLAN_PROD_WGT:             dMillPlan_Wgt = Val(.Text):                           dMillCal_Wgt = dMillCal_Wgt + Val(.Text)
            '实绩钢板重量（累计）
            .Col = SPD_RST_PROD_WGT:              dMillRst_Wgt = Val(.Text):                            dMillRst_SumWgt = dMillRst_SumWgt + dMillRst_Wgt
            '板坯重量（累计）
            .Col = SPD_SLAB_WGT:                  dMillSlab_Wgt = Val(.Text):                           dMillSlab_SumWgt = dMillSlab_SumWgt + dMillSlab_Wgt
            '实绩钢板重量对应板坯量（累计）
            If dMillRst_Wgt > 0 Then dMillRstSlab_SumWgt = dMillRstSlab_SumWgt + dMillSlab_Wgt
            '计划成材率
            .Col = SPD_PLAN_PROD_RAT:            .Text = dMillPlan_Wgt * 100 / dMillSlab_Wgt
            '实绩成材率
            .Col = SPD_RST_PROD_RAT:             .Text = dMillRst_Wgt * 100 / dMillSlab_Wgt
            '开轧时间
            .Col = SPD_MILLSTR_DATE:              dMillstr_date = .Text

             If dMillstr_date = "" Or dMill_date = "" Then
                dMill_dur = ""
             Else
'                dMill_dur = CStr(Round(((CDate(dMillstr_date) - CDate(dMill_date)) * 24 * 60 * 60), 1))
'                 fix(DateDiff("s", CDate(dMill_date), CDate(dMillstr_date)) / 60)    '取整
'                 DateDiff("s", "2010-07-24 05:54:16", "2010-07-24 05:55:19") mod 60  '取余
                 dMill_dur = DateDiff("s", CDate(dMill_date), CDate(dMillstr_date))

             End If
            '终轧时间
            .Col = SPD_MILLEND_DATE:             dMill_date = .Text
            '轧制间隔
            .Col = SPD_DUR_DATE:                .Text = dMill_dur
            
             If dMillstr_date = "" Or dMill_date = "" Then
                dMill_dtmin = ""
             Else
                dMill_dtmin = CStr(Round(((CDate(dMill_date) - CDate(dMillstr_date)) * 24 * 60), 1))
             End If
            '纯轧时间
            .Col = SPD_MILL_DATE:                .Text = dMill_dtmin
            
            '出炉时间
            .Col = SPD_DISCHARGE_DATE:             dDischarge_date = .Text
             If dMillstr_date = "" Or dDischarge_date = "" Then
                dRhfMill_dtmin = ""
             Else
                dRhfMill_dtmin = CStr(Round(((CDate(dMillstr_date) - CDate(dDischarge_date)) * 24 * 60), 1))
             End If
            '出炉到开轧时间
            .Col = SPD_RHFMILL_DATE:             .Text = dRhfMill_dtmin
            
            '实绩剪切母板块数
            .Col = SPD_MP_CNT:                   dMoplate_cnt = .Value
            If dMoplate_cnt < 1 Then
               Call Gp_Sp_BlockColor(ss1, 2, 2, ss1.Row, ss1.Row, , SSP1.BackColor)
            End If
            
            '轧制平均厚度
            .Col = SPD_THK_AVE:                   dThk_Ave = Val(.Text)
            '厚度公差下限
            .Col = SPD_THK_MIN:                   dThk_Min = Val(.Text)
            '厚度公差上限
            .Col = SPD_THK_MAX:                   dThk_Max = Val(.Text)
            
            If dThk_Ave < dThk_Min Or dThk_Ave > dThk_Max Then
               Call Gp_Sp_BlockColor(ss1, SPD_THK_AVE, SPD_THK_AVE, ss1.Row, ss1.Row, , SSP2.BackColor)
            End If
            
             '紧急订单绿色标记 2012-11-09  BY  LICHAO
            ss1.Row = .Row:       ss1.Col = SPD_URGNT_FL
           If ss1.Text = "Y" Then
                Call Gp_Sp_BlockColor(ss1, SPD_SLAB_NO, SPD_SLAB_NO, .Row, .Row, &HC000&)
                Call Gp_Sp_BlockColor(ss1, SPD_URGNT_FL, SPD_URGNT_FL, .Row, .Row, &HC000&)
           End If
            
        Next iCount
            .Row = .MaxRows
            .Col = SPD_SLAB_NO:                  .Text = "轧制块数"
            '累计轧制块数
            .Col = SPD_GROUP:                    .Text = dMillSlab_Cnt
            '累计轧制重量
            .Col = SPD_SLAB_WGT:                 .Text = dMillSlab_SumWgt
            '累计计划产品重量
            .Col = SPD_PLAN_PROD_WGT:            .Text = Str(Round(dMillCal_Wgt, 3)):             dMillPlan_Wgt = Val(.Text)
            '累计计划成材率
            .Col = SPD_PLAN_PROD_RAT:            .Text = dMillPlan_Wgt * 100 / dMillSlab_SumWgt
            '累计实绩产品重量
            .Col = SPD_RST_PROD_WGT:             .Text = Str(dMillRst_SumWgt)
            '累计实际成材率
            .Col = SPD_RST_PROD_RAT:             .Text = dMillRst_SumWgt * 100 / dMillSlab_SumWgt
            
            
    End With
    
    Call ss1.SetActiveCell(1, ss1.MaxRows)
               
End Sub
Public Sub Form_Pro()

    If Gf_Mc_Authority(sAuthority, Mc1) Then
        If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
               Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        End If
    End If

End Sub

Public Sub Form_Del()

'    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

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

Public Sub Spread_Del()

    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

Public Sub Spread_Cpy()

'    Call Gp_Sp_Copy(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Pst()

'    Call Gp_Sp_Paste(Proc_Sc("Sc"))

End Sub
Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol

End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If (Col = SS1_FIRST_REMARK) Then

        ss1.Row = ss1.ActiveRow
        ss1.Col = 0
        ss1.Text = "Update"
        
    End If
        
End Sub
