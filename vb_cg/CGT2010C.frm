VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGT2010C 
   Caption         =   "轧钢实绩查询_CGT2010C"
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10815
   ScaleWidth      =   15720
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
      ItemData        =   "CGT2010C.frx":0000
      Left            =   6720
      List            =   "CGT2010C.frx":0002
      TabIndex        =   16
      Tag             =   "班别"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txt_stlgrd1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   4620
      MaxLength       =   12
      TabIndex        =   15
      Tag             =   "钢种"
      Top             =   540
      Width           =   1785
   End
   Begin VB.TextBox txt_stdspec 
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
      Left            =   11565
      TabIndex        =   7
      Top             =   540
      Width           =   1620
   End
   Begin VB.TextBox txt_mill_stlgrd 
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
      Left            =   8760
      TabIndex        =   6
      Top             =   540
      Width           =   1500
   End
   Begin VB.TextBox TXT_MILL_LOT_NO 
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
      Left            =   11565
      MaxLength       =   14
      TabIndex        =   4
      Top             =   120
      Width           =   1620
   End
   Begin VB.TextBox TXT_CH_CD 
      Height          =   270
      Left            =   8610
      TabIndex        =   12
      Top             =   9510
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txt_stlgrd 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   540
      Width           =   1680
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
      Left            =   8775
      MaxLength       =   10
      TabIndex        =   3
      Top             =   120
      Width           =   1500
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
      ItemData        =   "CGT2010C.frx":0004
      Left            =   5985
      List            =   "CGT2010C.frx":0011
      TabIndex        =   2
      Top             =   120
      Width           =   780
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   7635
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
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
      Left            =   120
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "粗轧时间"
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
      Left            =   1245
      TabIndex        =   0
      Tag             =   "起始日期"
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
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
      Left            =   2985
      TabIndex        =   1
      Tag             =   "起始日期"
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
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
   Begin Threed.SSOption OPT_Cm 
      Height          =   330
      Index           =   2
      Left            =   14370
      TabIndex        =   9
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   8421504
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
      Caption         =   "精轧"
   End
   Begin Threed.SSOption OPT_Cm 
      Height          =   330
      Index           =   1
      Left            =   13350
      TabIndex        =   8
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   8421504
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
      Caption         =   "粗轧"
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   120
      Top             =   540
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "板坯钢种"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8295
      Left            =   120
      TabIndex        =   10
      Top             =   930
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   14631
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColsFrozen      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   78
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGT2010C.frx":001E
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   10440
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
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
   Begin Threed.SSOption OPT_Cm 
      Height          =   330
      Index           =   3
      Left            =   13350
      TabIndex        =   13
      Top             =   540
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   8421504
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
      Caption         =   "轧废"
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   10440
      Top             =   540
      Width           =   1080
      _ExtentX        =   1905
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
   Begin Threed.SSOption OPT_Cm 
      Height          =   330
      Index           =   4
      Left            =   14370
      TabIndex        =   14
      Top             =   540
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   8421504
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
      Caption         =   "出炉"
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   4860
      Top             =   120
      Width           =   1080
      _ExtentX        =   1905
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   7650
      Top             =   540
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "轧制钢种"
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   3150
      Top             =   540
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      Caption         =   "原始坯料钢种"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin Threed.SSPanel SSP4 
      Height          =   315
      Left            =   6450
      TabIndex        =   17
      Top             =   540
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   0
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
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   120
      Left            =   2775
      TabIndex        =   11
      Top             =   180
      Width           =   195
   End
End
Attribute VB_Name = "CGT2010C"
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
'-- Program Name      轧制实绩查询界面
'-- Program ID        CGT2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2007.11.5
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

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SS1_SLAB_NO = 1
Const SS1_FIRST_REMARK = 2
Const SS1_ORD_NO = 50
Const SS1_URGNT_FL = 52  '紧急订单绿色标记 2012-08-16  by  LiQian
Const SS1_IMP_CONT = 53

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_MILL_LOT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) ''ADDED BY LICHAO AT 2011-11-15
       Call Gp_Ms_Collection(txt_stlgrd1, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_mill_stlgrd, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl) ''ADDED BY GUOLI AT 20100315152000
       Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_CH_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
        Mc1.Add Item:=pControl, Key:="pControl"
        Mc1.Add Item:=nControl, Key:="nControl"
        Mc1.Add Item:=mControl, Key:="mControl"
        Mc1.Add Item:=iControl, Key:="iControl"
        Mc1.Add Item:=rControl, Key:="rControl"
        Mc1.Add Item:=cControl, Key:="cControl"
        Mc1.Add Item:=aControl, Key:="aControl"
        Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '首件标识
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
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
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '同板差
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
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)  '实绩成品重量
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn) '紧急订单绿色标记 2012-08-16  by  LiQian
    Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 67, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 68, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 69, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 70, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 71, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 72, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 73, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 74, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 75, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 76, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 77, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 78, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    
  
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGT2010C.P_SREFER", Key:="P-R"
    sc1.Add Item:="CGT2010C.P_MODIFY", Key:="P-M"
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
    
    OPT_Cm(1).Value = True
    
    CBO_GROUP.AddItem "A"
    CBO_GROUP.AddItem "B"
    CBO_GROUP.AddItem "C"
    CBO_GROUP.AddItem "D"
    
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

Public Sub Form_Ref()

    Dim iCount          As Integer
    Dim dMillCal_Wgt    As Double
    Dim dMillSlab_Wgt   As Double
    Dim dMillPlan_Wgt   As Double
    Dim simpcont As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    If Trim(TXT_MILL_LOT_NO.Text) <> "" Then
       If Len(Trim(TXT_MILL_LOT_NO)) < 12 Then
          MsgBox "轧批号必须输满12位，才可按轧批号查询！", vbCritical, "系统提示信息"
          Exit Sub
       End If
    End If
    
    If Trim(txt_slab_no.Text) <> "" Then
       If Len(Trim(txt_slab_no)) < 8 Then
          MsgBox "板坯号必须输满8位，才可按板坯号查询！", vbCritical, "系统提示信息"
          Exit Sub
       End If
    End If
    
    ss1.MaxRows = 0
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    dMillCal_Wgt = 0
    With ss1
        If .MaxRows <= 1 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .ROW = iCount

            .Col = 35 '31
             dMillCal_Wgt = dMillCal_Wgt + Val(.Text)
             dMillPlan_Wgt = Val(.Text)

            .Col = 18 '16
             dMillSlab_Wgt = Val(.Text)
            If dMillSlab_Wgt <> 0 Then
               .Col = 36 '32
               .Text = Round(dMillPlan_Wgt * 100 / dMillSlab_Wgt, 2)
            End If
            
             '紧急订单绿色标记 2012-08-16  by  LiQian
            .Col = SS1_URGNT_FL
            If .Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SS1_SLAB_NO, .MaxRows, .ROW, .ROW, &HC000&)
            End If
            
            .ROW = iCount:
            .Col = SS1_IMP_CONT:    simpcont = Trim(.Text)
            If simpcont = "Y" Then
              Call Gp_Sp_BlockColor(ss1, SS1_SLAB_NO, SS1_SLAB_NO, iCount, iCount, SSP4.BackColor)
              Call Gp_Sp_BlockColor(ss1, SS1_IMP_CONT, SS1_IMP_CONT, iCount, iCount, SSP4.BackColor)
            End If

        Next iCount

    End With
    
    Call ss1.SetActiveCell(1, ss1.MaxRows)
               
End Sub

Private Sub OPT_Cm_Click(Index As Integer, Value As Integer)
    Dim i As Integer
    
    For i = 1 To 4
        If i = Index Then
           OPT_Cm(i).ForeColor = &HFF&
        Else
           OPT_Cm(i).ForeColor = &H808080
        End If
    Next i
    
    Select Case Index
           Case 1
                TXT_CH_CD = "C"
                ULabel5.Caption = "粗轧时间"
           Case 2
                TXT_CH_CD = "J"
                ULabel5.Caption = "精轧时间"
           Case 3
                TXT_CH_CD = "S"
                ULabel5.Caption = "轧废时间"
           Case 4
                TXT_CH_CD = "H"
                ULabel5.Caption = "出炉时间"
    End Select

End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()
     SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
End Sub


Private Sub SDT_PROD_DATE_TO_GotFocus()
     SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
End Sub



Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    If ss1.MaxRows < 1 Then Exit Sub
    
        
      If (Col = SS1_FIRST_REMARK) Then

        ss1.ROW = ss1.ActiveRow
        ss1.Col = 0
        ss1.Text = "Update"
        
    End If
End Sub

Private Sub txt_mill_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

         DD.sWitch = "MS"
         DD.rControl.Add Item:=txt_mill_stlgrd

         DD.nameType = "1"
         Call Gf_Stlgrd_DD_AC(M_CN1, KeyCode)
    End If
End Sub

Private Sub txt_stdspec_DblClick()
    Call txt_STDSPEC_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STDSPEC

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub txt_stlgrd1_DblClick()

    Call txt_stlgrd1_KeyUp(vbKeyF4, 0)

End Sub
Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_mill_stlgrd_DblClick()

    Call txt_mill_stlgrd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_stlgrd1_KeyUp(KeyCode As Integer, Shift As Integer)
   
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd1
        DD.nameType = "1"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    End If
        
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyF4 Then

         DD.sWitch = "MS"
         DD.rControl.Add Item:=txt_stlgrd

         DD.nameType = "1"
         Call Gf_Stlgrd_DD_AC(M_CN1, KeyCode)
    End If
End Sub

Public Function Gf_Stlgrd_DD_AC(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "S"        'Stlgrd Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT STLGRD ""钢种"", STEEL_GRD_DETAIL ""目标说明"" FROM  NISCO.QP_NISCO_CHMC "
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.rControl.Item(1).Text) & "%' AND STLGRD_FL <> 'H'  "
            
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(STEEL_GRD_DETAIL,'%')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  STLGRD  ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "SELECT STLGRD ""钢种"", STEEL_GRD_DETAIL ""目标说明"" FROM  NISCO.QP_NISCO_CHMC "
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.sPname.Text) & "%' AND STLGRD_FL <> 'H' "
            
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            DD.sWhere = DD.sWhere + " AND NVL(STEEL_GRD_DETAIL,'%')   like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  STLGRD  ASC "
   
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
            If DD.rControl.Count > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                sNew_Name = DD.sPname.Text
            End If
            
            DD.sPname.TabStop = True
            DD.sPname.SetFocus
            DD.sPname.SetActiveCell DD.rControl.Item(1), DD.sPname.ActiveRow
            DD.sPname.Action = SS_ACTION_ACTIVE_CELL
            DD.sPname.EditMode = True
            DD.sPname.TabStop = False
            
            If DD.sSelect Then
                If sOld_Code <> sNew_Code Then Call Gp_Sp_UpdateMake(DD.sPname, False)
            End If
            
        End If
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function

Public Sub Form_Pro()
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
End Sub

Public Sub Spread_Del()
    
    Dim i As Long
    
    With Proc_Sc("Sc").Item("Spread")
        
        If .MaxRows < 1 Then Exit Sub
        If .SelBlockRow < 1 Then Exit Sub
        
        For i = .SelBlockRow To .SelBlockRow2
                .Col = 0
                If Trim(.Text) = "" Then
                    .Text = "Delete"
                End If
        Next i
        
    End With
    
End Sub


Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
   If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
   End If
End Sub


