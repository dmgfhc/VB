VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQD1000C 
   Caption         =   "原料成分实绩输入_AQD1000C"
   ClientHeight    =   9090
   ClientLeft      =   315
   ClientTop       =   1470
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_RAW_ITEM_CD 
      Height          =   300
      Index           =   1
      Left            =   1740
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cbo_RAW_ITEM_CD 
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
      Index           =   0
      Left            =   1740
      TabIndex        =   16
      Top             =   390
      Width           =   1215
   End
   Begin VB.TextBox txt_test_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1710
      MaxLength       =   7
      TabIndex        =   15
      Tag             =   "EMP_ID"
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox txt_RAW_AREA_NAME 
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
      Height          =   310
      Left            =   2970
      MaxLength       =   7
      TabIndex        =   14
      Tag             =   "EMP_ID"
      Top             =   2310
      Width           =   1965
   End
   Begin VB.TextBox txt_RAW_AREA_CD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   13
      Tag             =   "供应商"
      Top             =   2310
      Width           =   1245
   End
   Begin VB.TextBox txt_RAW_SUP_NAME 
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
      Height          =   310
      Left            =   2970
      MaxLength       =   7
      TabIndex        =   11
      Tag             =   "EMP_ID"
      Top             =   1650
      Width           =   1965
   End
   Begin VB.TextBox txt_RAW_ITEM_NAME 
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
      Height          =   310
      Left            =   4230
      TabIndex        =   10
      Top             =   395
      Width           =   1995
   End
   Begin VB.TextBox txt_RAW_ITEM_CD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   3030
      MaxLength       =   6
      TabIndex        =   8
      Tag             =   "原料名称"
      Top             =   395
      Width           =   1185
   End
   Begin VB.TextBox txt_TEST_EMP 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1710
      MaxLength       =   7
      TabIndex        =   7
      Tag             =   "EMP_ID"
      Top             =   6720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txt_SMP_CUT_EMP 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1710
      MaxLength       =   7
      TabIndex        =   6
      Tag             =   "EMP_ID"
      Top             =   6090
      Width           =   2655
   End
   Begin VB.TextBox txt_RAW_TEST_SPEC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1710
      MaxLength       =   30
      TabIndex        =   5
      Tag             =   "EMP_ID"
      Top             =   2940
      Width           =   2655
   End
   Begin VB.TextBox txt_SMP_CUT_TIME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "EMP_ID"
      Top             =   5460
      Width           =   1485
   End
   Begin VB.TextBox txt_SMP_STD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1710
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "EMP_ID"
      Top             =   3570
      Width           =   2655
   End
   Begin VB.TextBox txt_SMP_CUT_LOC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1710
      MaxLength       =   15
      TabIndex        =   2
      Tag             =   "EMP_ID"
      Top             =   4170
      Width           =   2655
   End
   Begin VB.TextBox txt_RAW_SUP_CD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1710
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "供应商"
      Top             =   1650
      Width           =   1245
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8745
      Left            =   6570
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   8655
      _Version        =   393216
      _ExtentX        =   15266
      _ExtentY        =   15425
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
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
      MaxCols         =   9
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQD1000C.frx":0000
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   270
      Top             =   390
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "原料名称"
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
      Left            =   270
      Top             =   1023
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "检验日期"
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
      Index           =   2
      Left            =   270
      Top             =   1656
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "供应商"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   3
      Left            =   270
      Top             =   4188
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "取样地点"
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
      Index           =   4
      Left            =   270
      Top             =   3585
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "取样标准"
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
      Index           =   5
      Left            =   270
      Top             =   4821
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "取样日期"
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
      Index           =   6
      Left            =   270
      Top             =   2955
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "试验标准"
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
      Index           =   7
      Left            =   270
      Top             =   6087
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "取样人员"
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
      Index           =   9
      Left            =   270
      Top             =   6720
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "试验人员"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   270
      Top             =   2289
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "产地"
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
   Begin InDate.UDate dtp_TEST_ISU_DATE 
      Height          =   310
      Left            =   1710
      TabIndex        =   9
      Tag             =   "检验日期"
      Top             =   1020
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
   Begin InDate.UDate dtp_SMP_CUT_DATE 
      Height          =   310
      Left            =   1710
      TabIndex        =   12
      Tag             =   "INS_DATE"
      Top             =   4800
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   270
      Top             =   5454
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "取样时间"
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
Attribute VB_Name = "AQD1000C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   原料成分实绩管理
'-- Program Name      原料成分实绩输入
'-- Program ID        AQD1000C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.09.05
'-- Description       原料成分实绩输入
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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

'Dim ArrayRecords As Variant

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Hsheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_RAW_ITEM_CD, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_RAW_ITEM_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(dtp_TEST_ISU_DATE, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_RAW_SUP_CD, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_RAW_SUP_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_RAW_AREA_CD, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_RAW_AREA_NAME, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_RAW_TEST_SPEC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_SMP_STD, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_SMP_CUT_LOC, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dtp_SMP_CUT_DATE, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_SMP_CUT_TIME, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_SMP_CUT_EMP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_TEST_EMP, " ", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_test_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:="AQD1000C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="AQD1000C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQD1000C.P_SMODIFY", Key:="P-M"
    Sc1.Add Item:="AQD1000C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AQD1000C.P_SONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=2, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub


Private Sub cbo_RAW_ITEM_CD_Click(Index As Integer)
    txt_RAW_ITEM_CD.Text = ""
    txt_RAW_ITEM_NAME.Text = ""
    cbo_RAW_ITEM_CD(1).ListIndex = cbo_RAW_ITEM_CD(0).ListIndex
    txt_RAW_ITEM_CD.Text = cbo_RAW_ITEM_CD(1).Text
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
        Case "txt_RAW_ITEM_CD"              '品种
            sCode = "F0001"
            Set oCodeName = txt_RAW_ITEM_NAME
            
        Case "txt_RAW_SUP_CD"               '订单用途
            sCode = "Q0044"
            Set oCodeName = txt_RAW_SUP_NAME
            
        Case "txt_RAW_AREA_CD"              '标准编号
            sCode = "Q0043"
            Set oCodeName = txt_RAW_AREA_NAME
            
'        Case "txt_RAW_ITEM_GRP"              '品种
'            sCode = "Q0054"
' '           Set oCodeName = txt_RAW_ITEM_NAME
'            txt_RAW_ITEM_CD = txt_RAW_ITEM_GRP
                
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub


Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.StatusBar1.Panels(1) = "Message : "
    Call GP_MENU_SHOW_HIDE("11F12F")

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
'    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 6)
    
    Call subComboAdd
    
    Call GP_MENU_SHOW_HIDE("11F12F")
    
    dtp_TEST_ISU_DATE.Text = ""
    dtp_SMP_CUT_DATE.Text = ""
    
    Screen.MousePointer = vbDefault
           
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
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
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
      
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


Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        Call GP_MENU_SHOW_HIDE("11F12F")
        pControl(1).SetFocus
    End If
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    If Len(Trim(txt_RAW_ITEM_CD.Text)) < 2 Then
    
        sMesg = "原料名称 is Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    
    End If

    If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        
        Call Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call GP_MENU_SHOW_HIDE("11F12F")
    End If


    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
    
    If Gf_Mc_Authority(sAuthority, Mc1, Proc_Sc("Sc")) Then
        txt_TEST_EMP.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1)
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call GP_MENU_SHOW_HIDE("11F12F")
        End If
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 8)
    
    Call GP_SET_CELL_VALUE(ss1, ss1.ActiveRow, 5, ss1.ActiveRow)

End Sub

Public Sub Form_Del()

    If Not Gf_Ms_AllDel(M_CN1, Proc_Sc("Sc"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub

Public Sub form_Cpy()
    
    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub form_Pst()
    
    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10, "P")
    End If
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_Paste(M_CN1, Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 8)
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("Sc"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 8)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 8)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sTemp_Code As String
    Dim iCol As Integer
    Dim sText As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    sText = ss1.ActiveRow

    Select Case ss1.ActiveCol

        Case 6

            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1

                DD.sWitch = "SP"
                DD.sKey = "Q0047"
                DD.rControl.Add Item:=ss1.ActiveCol
              '  DD.rControl.Add Item:=5
                DD.nameType = "2"

              Call Gf_Common_DD(M_CN1, KeyCode)
                    
            End If

    End Select

   '  ss1.SetFocus
End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
End Sub

Private Sub ss1_LostFocus()

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



'---------------------------------------------------------------------------------------
'   1.ID           : subComboAdd
'   2.Name         : 原料 Group ComboBox
'   3.Input  Value :
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 10. 14
'   7.Modify Date  :
'   8.Comment      : 原料 Group ComboBox
'---------------------------------------------------------------------------------------
Private Sub subComboAdd()

On Error GoTo Error_Rtn

    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant
    Dim sQuery As String
    Dim i As Integer
    
    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT CD_NAME , CD FROM ZP_CD WHERE CD_MANA_NO = 'Q0054'"
 
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    cbo_RAW_ITEM_CD(0).Clear
    cbo_RAW_ITEM_CD(1).Clear
    
    If Not AdoRs.EOF Then

        ArrayRecords = AdoRs.GetRows
        
        For i = 0 To UBound(ArrayRecords, 2)
            cbo_RAW_ITEM_CD(0).AddItem ArrayRecords(0, i)
            cbo_RAW_ITEM_CD(1).AddItem ArrayRecords(1, i)
        Next i
                                    
    End If
                
    AdoRs.Close
    Set AdoRs = Nothing
    
Error_Rtn:

End Sub

