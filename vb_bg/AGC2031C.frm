VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2031C 
   Caption         =   "钢板剪切实绩查询及修改界面_AGC2031C"
   ClientHeight    =   8835
   ClientLeft      =   1815
   ClientTop       =   3240
   ClientWidth     =   15195
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   15195
   WindowState     =   2  'Maximized
   Begin VB.Frame FRM4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "订单事项"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   120
      TabIndex        =   12
      Top             =   8310
      Width           =   15060
      Begin VB.TextBox TXT_SIZE 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   " "
         Top             =   270
         Width           =   2865
      End
      Begin VB.TextBox TXT_MARKING 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12180
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   " "
         Top             =   270
         Width           =   1440
      End
      Begin VB.TextBox TXT_UST 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   12180
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   " "
         Top             =   660
         Width           =   1440
      End
      Begin VB.TextBox TXT_SPEC 
         Height          =   315
         Left            =   8805
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   " "
         Top             =   660
         Width           =   1965
      End
      Begin VB.TextBox TXT_CUST 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   " "
         Top             =   645
         Width           =   1605
      End
      Begin VB.TextBox TXT_DEL_TO 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6135
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   " "
         Top             =   660
         Width           =   1245
      End
      Begin VB.TextBox TXT_DEL_FROM 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   " "
         Top             =   660
         Width           =   1245
      End
      Begin VB.TextBox TXT_ORD_NO 
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   " "
         Top             =   270
         Width           =   1605
      End
      Begin VB.TextBox TXT_WGT 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8805
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   " "
         Top             =   270
         Width           =   1965
      End
      Begin InDate.ULabel ULabel31 
         Height          =   315
         Left            =   3285
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "规格"
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
      Begin InDate.ULabel ULabel38 
         Height          =   315
         Left            =   7575
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "重量"
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
      Begin InDate.ULabel ULabel27 
         Height          =   315
         Left            =   240
         Top             =   270
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   3285
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "交货期"
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
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   240
         Top             =   645
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "客户"
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
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   7575
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "标准"
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
      End
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   10950
         Top             =   660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "是否UST"
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
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   10950
         Top             =   270
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "标识方法"
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
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         Height          =   75
         Left            =   5880
         TabIndex        =   22
         Top             =   780
         Width           =   150
      End
   End
   Begin VB.TextBox TXT_WGT_MIN 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   6675
      TabIndex        =   11
      Text            =   " "
      Top             =   10320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox TXT_WGT_MAX 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   5640
      TabIndex        =   10
      Text            =   " "
      Top             =   10320
      Visible         =   0   'False
      Width           =   1020
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   570
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   1005
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
      Begin VB.TextBox TXT_SEARCH_FL 
         Height          =   330
         Left            =   11430
         MaxLength       =   1
         TabIndex        =   6
         Text            =   "1"
         Top             =   30
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TXT_MPLATE_NO 
         Height          =   330
         Left            =   1605
         MaxLength       =   12
         TabIndex        =   3
         Top             =   120
         Width           =   1560
      End
      Begin VB.ComboBox CBO_LINE 
         Height          =   315
         ItemData        =   "AGC2031C.frx":0000
         Left            =   6585
         List            =   "AGC2031C.frx":000D
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
      Begin VB.ComboBox CBO_PLT 
         Height          =   315
         ItemData        =   "AGC2031C.frx":001D
         Left            =   4575
         List            =   "AGC2031C.frx":0024
         TabIndex        =   0
         Text            =   " "
         Top             =   120
         Width           =   735
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   360
         Top             =   120
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   3465
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "工厂代码"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
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
      Begin InDate.ULabel ULabel43 
         Height          =   315
         Left            =   5490
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         Caption         =   "机号"
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
      Begin Threed.SSOption opt_wait_product 
         Height          =   285
         Left            =   11520
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
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
         Caption         =   "等待剪切作业"
      End
      Begin Threed.SSOption opt_wait_inspect 
         Height          =   285
         Left            =   13065
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
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
         Caption         =   "等待表面检查"
      End
      Begin Threed.SSOption opt_all 
         Height          =   285
         Left            =   8310
         TabIndex        =   7
         Top             =   150
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   255
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
         Caption         =   "全部(等待剪切作业/表面检查)"
         Value           =   -1
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
         TabIndex        =   27
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "母板重量:"
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
         Left            =   12720
         TabIndex        =   26
         Top             =   225
         Width           =   990
      End
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7515
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   13256
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "AGC2031C.frx":002C
      Begin FPSpread.vaSpread ss2 
         Height          =   3855
         Left            =   0
         TabIndex        =   9
         Top             =   3660
         Width           =   15045
         _Version        =   393216
         _ExtentX        =   26538
         _ExtentY        =   6800
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
         MaxCols         =   13
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGC2031C.frx":007E
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3570
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   15045
         _Version        =   393216
         _ExtentX        =   26538
         _ExtentY        =   6297
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
         MaxCols         =   32
         MaxRows         =   20
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGC2031C.frx":1D8B
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "上限"
      Height          =   225
      Left            =   5970
      TabIndex        =   24
      Top             =   10095
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "下限"
      Height          =   225
      Left            =   6960
      TabIndex        =   23
      Top             =   10095
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "AGC2031C"
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
'-- Program ID        AGC2031C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2005.6.13
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
 
Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

Dim pControl1 As New Collection     'Master Primary Key Collection
Dim nControl1 As New Collection     'Master Necessary Collection
Dim mControl1 As New Collection     'Master Maxlength check Collection
Dim iControl1 As New Collection     'Master Insert Collection
Dim rControl1 As New Collection     'Master Refer Collection
Dim cControl1 As New Collection     'Master Copy Collection
Dim aControl1 As New Collection     'Master -> Spread Collection
Dim lControl1 As New Collection     'Master Lock Collection

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

Dim Mc1      As New Collection      'Master Collection
Dim Mc2      As New Collection      'Master Collection

Dim sc1      As New Collection      'Spread Collection
Dim sc2      As New Collection      'Spread Collection
Dim Proc_Sc  As New Collection      'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_PLATE_NO = 1
Const SPD_DS_CUT_STA_DATE = 9
Const SPD_DS_CUT_END_DATE = 10
Const SPD_THK = 11
Const SPD_WID = 12
Const SPD_LEN = 13
Const SPD_WGT = 14
Const SPD_DS_LAST_YN = 15
Const SPD_SURF_GRD = 16
Const SPD_TRIM_FL = 17
Const SPD_SMP_LEN = 18
Const SPD_DS_KNIFE_GAP = 19
Const SPD_DS_H_CROP_YN = 20
Const SPD_DS_T_CROP_YN = 21
Const SPD_MARK_YN = 22
Const SPD_STAMP_YN = 23
Const SPD_BAR_YN = 24
Const SPD_LEN_FLAG = 25
Const SPD_SF_ORNOT = 26
Const SPD_PLT = 27
Const SPD_PRC_LINE = 28
Const SPD_EMP_CD = 29
Const SPD_PROC_CD = 30
Const SPD_END_USE = 31
Const SPD_STLGRD = 32

Dim sQuery   As String
Dim sLoopFl  As String

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"
       
    'MASTER Collection
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_mplate_no, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_LINE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_SEARCH_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    Call Gp_Ms_Collection(txt_mplate_no, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(TXT_ORD_NO, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_SIZE, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'         Call Gp_Ms_Collection(TXT_WTH, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'         Call Gp_Ms_Collection(TXT_LTH, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_WGT, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'     Call Gp_Ms_Collection(TXT_WGT_MAX, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
'     Call Gp_Ms_Collection(TXT_WGT_MIN, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(TXT_DEL_FROM, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(TXT_DEL_TO, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_CUST, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
         Call Gp_Ms_Collection(TXT_SPEC, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_UST, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
      Call Gp_Ms_Collection(TXT_MARKING, " ", " ", " ", " ", "r", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    Mc2.Add Item:="AGC2031C.P_REFER", Key:="P-R"
    Mc2.Add Item:=pControl1, Key:="pControl"
    Mc2.Add Item:=nControl1, Key:="nControl"
    Mc2.Add Item:=mControl1, Key:="mControl"
    Mc2.Add Item:=iControl1, Key:="iControl"
    Mc2.Add Item:=rControl1, Key:="rControl"
    Mc2.Add Item:=cControl1, Key:="cControl"
    Mc2.Add Item:=aControl1, Key:="aControl"
    Mc2.Add Item:=lControl1, Key:="lControl"
        
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
     Call Gp_Sp_Collection(ss1, 9, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 27, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2031C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AGC2031C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
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
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGC2031C.P_SREFER2", Key:="P-R"
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

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
'        If Len(TXT_MPLATE_NO.Text) >= 8 Then
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
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))

    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    ss1.RowHeight(-1) = 13.5
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColHidden(ss1, 18, True)
    Call Gp_Sp_ColHidden(ss1, 19, True)
    Call Gp_Sp_ColHidden(ss1, 27, True)
    Call Gp_Sp_ColHidden(ss1, 28, True)
    Call Gp_Sp_ColHidden(ss1, 29, True)
    Call Gp_Sp_ColHidden(ss1, 30, True)
    Call Gp_Sp_ColHidden(ss1, 31, True)
    Call Gp_Sp_ColHidden(ss1, 32, True)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)
            
    CBO_PLT.ListIndex = 0
    CBO_LINE.ListIndex = 0
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing

    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing

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
            
            CBO_PLT.ListIndex = 0
            CBO_LINE.ListIndex = 0
            lbl_moplate_wgt.Caption = ""
            
            pControl(1).SetFocus
        End If
    End If
End Sub

Public Sub Form_Ref()
    
    On Error GoTo Refer_Err
    
    Dim iCount As Integer
    
    If TXT_SEARCH_FL.Text = "" Then TXT_SEARCH_FL.Text = "1"
        
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Len(txt_mplate_no.Text) > 9 Then
        
'        sLoopFl = "**"
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call PlateWgtEdit
            Call Gf_Ms_Refer(M_CN1, Mc2, , , False)
        End If
'        sLoopFl = ""
    Else
        Call Gf_Sp_Refer(M_CN1, sc2, Mc1)
        Call ss2_DblClick(1, 1)
    End If
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    With ss1
        .ReDraw = False
         For iCount = 1 To .MaxRows
            .Row = iCount
            .Col = SPD_PROC_CD
             If .Text = "CGB" Then
                .Col = SPD_MARK_YN:       .Value = 1
                .Col = SPD_STAMP_YN:      .Value = 1
                .Col = SPD_BAR_YN:        .Value = 1
            End If
'         Call .SetActiveCell(1, .ROW)
         Next iCount
        .ReDraw = True
    End With
           
Refer_Err:
       
End Sub

Public Sub Form_Pro()

    Dim iCount      As Integer
    Dim sDateFrom   As String
    Dim sDateTo     As String
    Dim sPlateNo    As String
    
    Dim inum As Integer
    Dim lRow As Integer
    
    If TXT_SEARCH_FL.Text = "" Then TXT_SEARCH_FL.Text = "1"
    
    With ss1
        For iCount = 1 To .MaxRows
            .Row = iCount
            .Col = SPD_PLATE_NO
            If Left(.Text, 12) = Left(sPlateNo, 12) Or sPlateNo = "" Then
               sPlateNo = .Text
               .Col = SPD_DS_LAST_YN
               If .Value = 1 Then
                   lRow = iCount
                   inum = inum + 1
                   If inum > 1 Then
                       Call Gp_MsgBoxDisplay("一块母板只能有一块尾板.." & Left(sPlateNo, 12))
                       Exit Sub
                   End If
               End If
               If inum = 1 Then
                    If iCount > lRow Then
                       .Col = 0
                       If .Text <> "Delete" Then
                          .Text = ""
                       End If
                    End If
               End If
            Else
               inum = 0
               sPlateNo = .Text
               .Col = SPD_DS_LAST_YN
               If .Value = 1 Then
                   lRow = iCount
                   inum = inum + 1
                   If inum > 1 Then
                       Call Gp_MsgBoxDisplay("一块母板只能有一块尾板.." & Left(sPlateNo, 12))
                       Exit Sub
                   End If
               End If
               If inum = 1 Then
                    If iCount > lRow Then
                       .Col = 0
                       .Text = ""
                    End If
               End If
            End If
        Next iCount
    End With
    
    For iCount = 1 To ss1.MaxRows
        With ss1
            Select Case Trim(Gf_Sp_RcvData(ss1, 0, iCount))
                
                Case "Input", "Update"
                        .Col = SPD_PLATE_NO: sPlateNo = .Text
                        .Col = SPD_DS_CUT_STA_DATE
                        sDateFrom = .Text
                        If Not Gp_DateCheck(.Text, "S") Then
                           Call Gp_MsgBoxDisplay("请正确输入开始时间.." & sPlateNo)
                           Exit Sub
                        End If
                        
                        .Col = SPD_DS_CUT_END_DATE
                        sDateTo = .Text
                        If Not Gp_DateCheck(.Text, "S") Then
                           Call Gp_MsgBoxDisplay("请正确输入结束时间.." & sPlateNo)
                           Exit Sub
                        End If
                        
                        If sDateFrom > sDateTo Then
                           Call Gp_MsgBoxDisplay("请正确输入开始时间还是结束时间.." & sPlateNo)
                           Exit Sub
                        End If
                        
'                        .Col = SPD_DS_LAST_YN
'                        If .Value = 1 And iCount <> .MaxRows Then
'                           If Gp_MsgBox("确定此钢板(" & sPlateNo & ")为尾板？", "C") = 7 Then
'                                Exit Sub
''                           Else
''                                .BlockMode = True
''                                .Row = iCount + 1:  .Row2 = .MaxRows
''                                .Col = 0:        .Col2 = 0:         .Text = ""
''                                .Col = SPD_PLT:  .Col2 = SPD_PLT:   .Text = ""
''                                .BlockMode = False
'                           End If
'                        End If
            End Select
        End With
    Next iCount
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
End Sub

Public Sub Form_Ins()
    Dim dThk        As Double
    Dim dWid        As Double
    Dim dLen        As Double
    Dim dWgt        As Double
    Dim lRow        As Long
    Dim sPlateNo    As String
    Dim sClipText   As String
    
    Dim iCount As Integer
    
    sPlateNo = ""
    
    With ss1
        If .MaxRows = 0 Then
           If Len(txt_mplate_no.Text) = 12 Then
               Call Gp_Sp_Ins(Proc_Sc("Sc"))
              .Row = 1
              .Col = 1
              .Text = txt_mplate_no.Text & "01"
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
        .Col = SPD_PLATE_NO: .Text = Mid(.Text, 1, 12) & Format(Val(Mid(.Text, 13, 2) & "") + 1, "00")
        .Col = SPD_SURF_GRD:      .Value = 1
        .Col = SPD_MARK_YN:       .Value = 1
        .Col = SPD_STAMP_YN:      .Value = 1
        .Col = SPD_BAR_YN:        .Value = 1
        
         Call .SetActiveCell(1, .Row)
        .ReDraw = True
    End With

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
            .Col = SPD_END_USE:   sEndUseCd = Trim(ss1.Text)
            .Col = SPD_STLGRD:    sStlgrd = Trim(ss1.Text)
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

    ss1.Row = ss1.ActiveRow:    ss1.Col = SPD_EMP_CD:      ss1.Text = sUserID

    Call Gp_Sp_Del(Proc_Sc("sc"))

End Sub

Public Sub Spread_Cpy()
    Call Gp_Sp_Copy(Proc_Sc("Sc"))
End Sub

Public Sub Spread_Pst()
    Call Gp_Sp_Paste(Proc_Sc("Sc"))
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

Private Sub opt_wait_product_Click(Value As Integer)
    opt_wait_product.ForeColor = &HFF&
    opt_wait_inspect.ForeColor = &H808080
    Opt_all.ForeColor = &H808080
    TXT_SEARCH_FL.Text = "1"
    txt_mplate_no.Text = ""
End Sub

Private Sub opt_wait_inspect_Click(Value As Integer)
    opt_wait_inspect.ForeColor = &HFF&
    opt_wait_product.ForeColor = &H808080
    Opt_all.ForeColor = &H808080
    TXT_SEARCH_FL.Text = "2"
    txt_mplate_no.Text = ""
End Sub

Private Sub opt_all_Click(Value As Integer)
    Opt_all.ForeColor = &HFF&
    opt_wait_inspect.ForeColor = &H808080
    opt_wait_product.ForeColor = &H808080
    TXT_SEARCH_FL.Text = "3"
    txt_mplate_no.Text = ""
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub


'Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
'    Dim iRow      As Long
'    Dim iTrimFl   As Long
'
'    If Col <> SPD_TRIM_FL Or Row < 1 Or sLoopFl = "**" Then Exit Sub
'
'    sLoopFl = "**"
'    ss1.Row = ss1.ActiveRow:   ss1.Col = SPD_TRIM_FL
'    If ss1.Value = 1 Then
'        iTrimFl = 1
'    Else
'        iTrimFl = 0
'    End If
'
'    For iRow = 1 To ss1.MaxRows
'        ss1.Row = iRow:   ss1.Col = SPD_TRIM_FL:  ss1.Value = iTrimFl
'    Next iRow
'    sLoopFl = ""
'
'End Sub

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
    
    If Row < 1 Then Exit Sub
    
    ss1.Row = Row: ss1.Col = Col
    If Col <> SPD_DS_CUT_STA_DATE And Col <> SPD_DS_CUT_END_DATE Then Exit Sub
    
    With ss1
        .Row = Row
        .Col = SPD_PLT:         .Text = Trim(CBO_PLT.Text)
        .Col = SPD_PRC_LINE:    .Text = Trim(CBO_LINE.Text)
        .Col = SPD_EMP_CD:      .Text = sUserID

        If Row > 1 Then
            .Row = Row - 1:   .Col = SPD_DS_CUT_END_DATE:     sDate = .Text

            .Row = Row: .Col = SPD_DS_CUT_STA_DATE
            If IsDate(sDate) Then
                .Text = sDate
            Else
                .Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
            End If
        End If

        .Row = Row:   .Col = SPD_DS_CUT_STA_DATE
        If IsDate(.Text) Then
            sDateTo = Format(DateAdd("n", 1, CDate(.Text)), "YYYY-MM-DD HH:MM:SS")
        Else
            .Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
        End If

        .Row = Row:   .Col = SPD_DS_CUT_END_DATE

        .Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
        
        Call ss1_Row_Edit(Row)
    End With
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
'    If Row > 0 Then
'        Set Active_Spread = Me.ss1
'        PopupMenu MDIMain.PopUp_Spread
'    End If
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


Private Sub ss1_Row_Edit(ByVal Row As Long)
    Dim iIdr        As Integer
    Dim sLastFlag   As String
    
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
        ss1.Row = iIdr
        ss1.Col = SPD_DS_LAST_YN
        If ss1.Value = 1 Then sLastFlag = "Y"
        
        ss1.Col = SPD_WGT
        lbl_moplate_wgt.Caption = Val(lbl_moplate_wgt.Caption & "") + Val(ss1.Text & "")
    Next iIdr
    
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    Dim iCount As Integer

    If Row < 1 Then Exit Sub
    
    ss2.Row = Row
    ss2.Col = 1
    txt_mplate_no.Text = ss2.Text
    
    If Len(txt_mplate_no.Text) = 12 Then
        
        Call Gf_Sp_Cls(sc1)
        Call Gp_Ms_Cls(Mc2("rControl"))
'        sLoopFl = "**"
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call PlateWgtEdit
            Call Gf_Ms_Refer(M_CN1, Mc2, , , False)
        End If
'        sLoopFl = ""
    End If
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    With ss1
        .ReDraw = False
         For iCount = 1 To .MaxRows
            .Row = iCount
            .Col = SPD_PROC_CD
             If .Text = "CGB" Then
                .Col = SPD_MARK_YN:       .Value = 1
                .Col = SPD_STAMP_YN:      .Value = 1
                .Col = SPD_BAR_YN:        .Value = 1
            End If
'         Call .SetActiveCell(1, .ROW)
         Next iCount
        .ReDraw = True
    End With
End Sub

