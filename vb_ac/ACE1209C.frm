VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACE1209C 
   Caption         =   "物料替代履历查询_ACE1209C"
   ClientHeight    =   9225
   ClientLeft      =   195
   ClientTop       =   2055
   ClientWidth     =   14370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14370
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9165
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   16166
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      Locked          =   -1  'True
      PaneTree        =   "ACE1209C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   7740
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1425
         Width           =   15210
         _Version        =   393216
         _ExtentX        =   26829
         _ExtentY        =   13653
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         ColsFrozen      =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   81
         MaxRows         =   2
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "ACE1209C.frx":0052
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   1395
         Left            =   0
         TabIndex        =   2
         Tag             =   "产品分类"
         Top             =   0
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   2461
         _Version        =   196609
         BackColor       =   14737632
         ShadowStyle     =   1
         Begin VB.ComboBox cbo_fp 
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
            Left            =   4695
            TabIndex        =   24
            Top             =   960
            Width           =   750
         End
         Begin VB.ComboBox Combo_ORD_ITEM 
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
            Left            =   2760
            TabIndex        =   22
            Top             =   960
            Width           =   660
         End
         Begin VB.TextBox Text_BB_ORD_NO 
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
            Left            =   1410
            MaxLength       =   11
            TabIndex        =   21
            Top             =   960
            Width           =   1350
         End
         Begin VB.TextBox txt_upd_cur_inv_code 
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
            Left            =   4890
            MaxLength       =   2
            TabIndex        =   20
            Top             =   540
            Width           =   495
         End
         Begin VB.TextBox txt_upd_cur_inv 
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
            Left            =   5430
            TabIndex        =   19
            Top             =   540
            Width           =   1230
         End
         Begin VB.TextBox txt_rep_kind 
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
            Left            =   14610
            MaxLength       =   40
            TabIndex        =   16
            Top             =   150
            Visible         =   0   'False
            Width           =   555
         End
         Begin Threed.SSOption opt1 
            Height          =   345
            Left            =   12810
            TabIndex        =   15
            Top             =   540
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   14737632
            Caption         =   "在制品"
         End
         Begin VB.TextBox txt_plt 
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
            Height          =   310
            Left            =   1470
            MaxLength       =   2
            TabIndex        =   12
            Tag             =   "生产厂"
            Top             =   120
            Width           =   435
         End
         Begin VB.TextBox txt_plt_nm 
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
            Left            =   1905
            TabIndex        =   11
            Tag             =   "生产厂"
            Top             =   120
            Width           =   1290
         End
         Begin VB.TextBox text_cur_inv 
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
            TabIndex        =   9
            Top             =   120
            Width           =   1470
         End
         Begin VB.TextBox text_cur_inv_code 
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
            Left            =   4890
            MaxLength       =   2
            TabIndex        =   8
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txt_prod_cd_name 
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
            Left            =   8745
            MaxLength       =   40
            TabIndex        =   7
            Tag             =   "产品"
            Top             =   120
            Width           =   1260
         End
         Begin VB.TextBox txt_prod_cd 
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
            Height          =   310
            Left            =   8265
            MaxLength       =   2
            TabIndex        =   6
            Tag             =   "产品"
            Top             =   120
            Width           =   465
         End
         Begin VB.TextBox txt_mat_no 
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
            Left            =   1470
            MaxLength       =   15
            TabIndex        =   5
            Tag             =   "物料编号"
            Top             =   540
            Width           =   1725
         End
         Begin VB.TextBox Txt_rep_typ_name 
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
            Left            =   12120
            MaxLength       =   40
            TabIndex        =   4
            Tag             =   "产品分类"
            Top             =   120
            Width           =   1605
         End
         Begin VB.TextBox Txt_rep_typ 
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
            Height          =   310
            Left            =   11655
            MaxLength       =   1
            TabIndex        =   3
            Tag             =   "产品分类"
            Top             =   120
            Width           =   465
         End
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   7245
            Top             =   120
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   556
            Caption         =   "产品"
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
            Left            =   150
            Top             =   540
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "物料编号"
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
            Left            =   10410
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "替代分类"
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
         Begin InDate.UDate dte_ins_date 
            Height          =   315
            Left            =   7770
            TabIndex        =   10
            Tag             =   "录入日期"
            Top             =   540
            Width           =   1470
            _ExtentX        =   2593
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
            Index           =   0
            Left            =   3570
            Top             =   120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "当前仓库"
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
         Begin InDate.ULabel ULabel17 
            Height          =   315
            Left            =   150
            Top             =   120
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "生产厂"
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
         Begin InDate.UDate dte_ins_date_to 
            Height          =   315
            Left            =   9300
            TabIndex        =   13
            Tag             =   "录入日期"
            Top             =   540
            Width           =   1470
            _ExtentX        =   2593
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   6810
            Top             =   540
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            Caption         =   "录入日期"
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
         Begin Threed.SSOption opt3 
            Height          =   345
            Left            =   12150
            TabIndex        =   17
            Top             =   540
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   609
            _Version        =   196609
            ForeColor       =   255
            BackColor       =   14737632
            Caption         =   "全部"
            Value           =   -1
         End
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   10890
            Top             =   540
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            Caption         =   "替代类型"
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
         Begin Threed.SSOption opt2 
            Height          =   345
            Left            =   13740
            TabIndex        =   18
            Top             =   540
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   14737632
            Caption         =   "产品"
         End
         Begin InDate.ULabel ULabel7 
            Height          =   315
            Left            =   3570
            Top             =   540
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            Caption         =   "替代时仓库"
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
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Index           =   1
            Left            =   120
            Top             =   960
            Width           =   1260
            _ExtentX        =   2223
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
         Begin Threed.SSCommand SSCommand2 
            Height          =   315
            Left            =   12240
            TabIndex        =   23
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   255
            Caption         =   "Excel导出"
         End
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   3600
            Top             =   960
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   556
            Caption         =   "火切实绩"
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
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "~"
            Height          =   120
            Left            =   9210
            TabIndex        =   14
            Top             =   630
            Width           =   90
         End
      End
   End
End
Attribute VB_Name = "ACE1209C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACE1209C
'-- Document No       Q-00-0010(Specification)
'-- Designer          JIANING
'-- Coder             JIANING
'-- Date              2003.9.29
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE        EDITOR       DESCRIPTION
'-- 1.01  2003.9.29   JIANING
'-- 1.02  2010.12.01  LiQian       加在制品、产品查询条件区分MES/ERP替代
'-- 1.03  2010.12.21  LiQian       添加原始订单厚度，宽度，长度和重量，替代前列表靠后显示
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
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iCount As Integer
Const C_LOST_WGT = 44   '29->30->31->33->35->36->37->42->43
Const C_ORG_WGT = 42    '27->28->29->31->33->34->35->40->41
Const C_WGT = 31        '21->22->23->25->27->28->29->30

Const SS1_PLT = 3
Const SS1_MV_DATE = 5
Const SS1_PLATE_NO = 2
Const SS1_CUST_CD = 15
Const SS1_OUT_SHEET_NO = 52
Const SS1_Shift = 53
Const SS1_TRNS_CMPY_CD = 54
Const SS1_CUST_CD1 = 55
Const SS1_LEN = 56
Const SS1_ORD_THK = 57
Const SS1_ORD_WID = 58
Const SS1_ORD_LEN = 59
Const SS1_TRIM_FL = 60  '切边
Const SS1_SIZE_KND = 61
Const SS1_ORD_REMARK = 62
Const SS1_PROD_REMARK = 63
Const SS1_UST_STATUS = 64
Const SS1_GAS_STATUS = 65 '切割
Const SS1_CL_STATUS = 66
Const SS1_HTM_METH = 67
Const SS1_QT = 68
Const SS1_STDSPEC_ORG_KND = 69
Const SS1_STDSPEC_STLGRD = 70
Const SS1_PLATE_CON = 71
Const SS1_PLATE_SIZE = 72
Const SS1_APLY_STDSPEC = 73
Const SS1_RM_CR_STAGE3_TIME = 74
Const SS1_SURFACE_REQUESTS = 75
Const SS1_VESSEL_NO = 76
Const SS1_SIDEMARK = 77

Const SS1_PAINTNUM = 78
Const SS1_GANGYIN = 79
Const SS1_PUNCH = 80
Const SS1_CUST = 81



Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(txt_plt, "p", " ", "m", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_prod_cd, "p", "n", "m", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(dte_ins_date, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(dte_ins_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(Txt_rep_typ, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(Txt_rep_typ_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(text_cur_inv, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_upd_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_upd_cur_inv, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     '替代类型：在制品或产品(在ERP中做替代还是MES中做替代)
     '20101130  015725
     Call Gp_Ms_Collection(txt_rep_kind, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Text_BB_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Combo_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(cbo_fp, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                                                                     
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)
   Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '轧制号
   Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '班次
   Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '分段号
   Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '客户代码
   Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '母板长
   Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '订单厚度
   Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '订单宽度
   Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '订单长度
   Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '切边
   Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '定尺
   Call Gp_Sp_Collection(ss1, 62, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '订单备注
   Call Gp_Sp_Collection(ss1, 63, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '产品备注
   Call Gp_Sp_Collection(ss1, 64, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '探伤
   Call Gp_Sp_Collection(ss1, 65, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '切割
   Call Gp_Sp_Collection(ss1, 66, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '矫直
   Call Gp_Sp_Collection(ss1, 67, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '热处理
   Call Gp_Sp_Collection(ss1, 68, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '其他
   Call Gp_Sp_Collection(ss1, 69, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '标识标准
   Call Gp_Sp_Collection(ss1, 70, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '标识钢种
   Call Gp_Sp_Collection(ss1, 71, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '子板数
   Call Gp_Sp_Collection(ss1, 72, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '子板尺寸
   Call Gp_Sp_Collection(ss1, 73, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '子板标准号
   Call Gp_Sp_Collection(ss1, 74, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '订单数量
   Call Gp_Sp_Collection(ss1, 75, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '客户表面要求
   Call Gp_Sp_Collection(ss1, 76, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '加喷内容
   Call Gp_Sp_Collection(ss1, 77, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '侧喷加喷
   Call Gp_Sp_Collection(ss1, 78, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '表喷次数
   Call Gp_Sp_Collection(ss1, 79, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '是否钢印
   Call Gp_Sp_Collection(ss1, 80, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '钢印加冲
   Call Gp_Sp_Collection(ss1, 81, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, True)   '用户交货期
    
    
  
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACE1209C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    
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
    cbo_fp.AddItem " "
    cbo_fp.AddItem "Y"
 
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    'Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
     Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    If App.Title = "CE" Then
        txt_plt.Text = "C3"
        text_cur_inv_code.Text = "ZB"
        Call text_cur_inv_code_KeyUp(0, 0)
    Else
        txt_plt.Text = "C1"
        text_cur_inv_code.Text = "00"
    End If
    
    txt_prod_cd.Text = "PP"
    
    opt3.Value = True
    
    txt_rep_kind.Text = "3"

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
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
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        txt_prod_cd.Text = ""
        TXT_MAT_NO.Text = ""
        Txt_rep_typ.Text = ""
        Txt_rep_typ_name.Text = ""
            
        If App.Title = "CE" Then
            txt_plt.Text = "C3"
            text_cur_inv_code.Text = "ZB"
            Call text_cur_inv_code_KeyUp(0, 0)
        Else
            txt_plt.Text = "C1"
            text_cur_inv_code.Text = "00"
        End If
        
        txt_prod_cd.Text = "PP"
        
        opt3.Value = True
        
        txt_rep_kind.Text = "3"

    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim sQuery As String
    Dim sMesg As String
    Dim ORG_WGT As Double
    Dim WGTS As Double
    
    
    If Len(Trim(dte_ins_date.RawData)) < 4 Then
        Call Gp_MsgBoxDisplay("录入日期(年) Must input necessarily")
        Exit Sub
    End If
    
    If (txt_upd_cur_inv_code.Text = "") And ((text_cur_inv_code.Text = "" And txt_plt.Text = "")) Then
       'If (text_cur_inv_code.Text = "" And txt_plt.Text = "") Then
          Call Gp_MsgBoxDisplay("当前仓库和生产厂有一个不能为空，或者替代时仓库不能为空!", "I", "错误提示")
          'Exit Sub
       'End If
    End If
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    With ss1
   
       If .MaxRows < 1 Then
           Exit Sub
       End If
       If txt_prod_cd.Text = "SL" Then Exit Sub
       For iCount = 1 To .MaxRows

   
            .Row = iCount:            .Col = C_ORG_WGT
             ORG_WGT = Val(.Text)
            .Row = iCount:            .Col = C_WGT
             WGTS = Val(.Text)
            .Row = iCount:            .Col = C_LOST_WGT
            .Text = ORG_WGT - WGTS

        Next iCount
      
       
   End With

  
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

Private Sub opt1_Click(Value As Integer)
    If opt1.Value = True Then
        opt3.ForeColor = &H80000012
        opt2.ForeColor = &H80000012
        opt1.ForeColor = &HFF&
        txt_rep_kind.Text = "1"
        End If
End Sub

Private Sub opt2_Click(Value As Integer)
    If opt2.Value = True Then
        opt3.ForeColor = &H80000012
        opt2.ForeColor = &HFF&
        opt1.ForeColor = &H80000012
        txt_rep_kind.Text = "2"
        End If
End Sub

Private Sub opt3_Click(Value As Integer)
    If opt3.Value = True Then
        opt3.ForeColor = &HFF&
        opt2.ForeColor = &H80000012
        opt1.ForeColor = &H80000012
        txt_rep_kind.Text = "3"
        End If
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

Private Sub SSCommand2_Click()

   
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

   
   Call Gp_ACE1209C_Excel1(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
   
  
End Sub
Private Sub Gp_ACE1209C_Excel1(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)

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
    Dim i           As Integer
    
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
        
        For i = 1 To ss1.MaxRows
        Set xlSheet = xlBook.Worksheets(i)
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "@"
        
        ss1.Row = i
        
        xlSheet.Range("A1").Value = "生产厂"
        xlSheet.Range("A2").Value = "日期"
        xlSheet.Range("A3").Value = "产品号"
        xlSheet.Range("A4").Value = "客户"
        xlSheet.Range("A5").Value = "母板长"
        xlSheet.Range("A6").Value = "厚度"
        xlSheet.Range("A7").Value = "宽度"
        xlSheet.Range("A8").Value = "长度"
        xlSheet.Range("A9").Value = "切边"
        xlSheet.Range("A10").Value = "订单备注"
        xlSheet.Range("A11").Value = "产品备注"
        
        xlSheet.Range("A12").Value = "探伤"
        xlSheet.Range("A13").Value = "切割"
        xlSheet.Range("A14").Value = "矫直"
        xlSheet.Range("A15").Value = "热处理"
        xlSheet.Range("A16").Value = "标识标准"
        xlSheet.Range("A17").Value = "标识钢种"
        xlSheet.Range("A18").Value = "其它"
        xlSheet.Range("A19").Value = "子板数"
        xlSheet.Range("A20").Value = "子板尺寸"
        xlSheet.Range("A22").Value = "子板标准号"     '
'

        
        
        xlSheet.Range("C1").Value = "轧批号"
        xlSheet.Range("C2").Value = "班次"
        xlSheet.Range("C3").Value = "分断号"
        xlSheet.Range("C4").Value = "客户代码"
        
        xlSheet.Range("C6").Value = "厚度公差"
        xlSheet.Range("C7").Value = "宽度公差"
        xlSheet.Range("C8").Value = "长度公差"
        xlSheet.Range("C9").Value = "定尺"
        xlSheet.Range("C10").Value = "订单数量"
        xlSheet.Range("C11").Value = "客户表面要求"
        xlSheet.Range("C12").Value = "是否加喷CE"
        xlSheet.Range("C13").Value = "重量"
        xlSheet.Range("C14").Value = "加喷内容"
        xlSheet.Range("C15").Value = "侧喷加喷"
        xlSheet.Range("C16").Value = "表喷次数"
        xlSheet.Range("C17").Value = "是否钢印"
        xlSheet.Range("C18").Value = "钢印加冲"
        
        xlSheet.Range("C23").Value = "交货期"
     
        
  
        xlSheet.Range("A20", "A21").Merge
        xlSheet.Range("B20", "D21").Merge
        
        
        ss1.Col = SS1_PLT:               xlSheet.Range("B1").Value = ss1.Text
        ss1.Col = SS1_MV_DATE:           xlSheet.Range("B2").Value = ss1.Text
        ss1.Col = SS1_PLATE_NO:          xlSheet.Range("B3").Value = ss1.Text
        ss1.Col = SS1_CUST_CD:           xlSheet.Range("B4").Value = ss1.Text
        ss1.Col = SS1_LEN:               xlSheet.Range("B5").Value = ss1.Text
        ss1.Col = SS1_ORD_THK:           xlSheet.Range("B6").Value = ss1.Text
        ss1.Col = SS1_ORD_WID:           xlSheet.Range("B7").Value = ss1.Text
        ss1.Col = SS1_ORD_LEN:           xlSheet.Range("B8").Value = ss1.Text
        ss1.Col = SS1_TRIM_FL:           xlSheet.Range("B9").Value = ss1.Text
        ss1.Col = SS1_ORD_REMARK:        xlSheet.Range("B10").Value = ss1.Text
        ss1.Col = SS1_PROD_REMARK:       xlSheet.Range("B11").Value = ss1.Text
        ss1.Col = SS1_UST_STATUS:        xlSheet.Range("B12").Value = ss1.Text
        ss1.Col = SS1_GAS_STATUS:        xlSheet.Range("B13").Value = ss1.Text
        ss1.Col = SS1_CL_STATUS:         xlSheet.Range("B14").Value = ss1.Text
        ss1.Col = SS1_HTM_METH:          xlSheet.Range("B15").Value = ss1.Text
        ss1.Col = SS1_STDSPEC_ORG_KND:   xlSheet.Range("B16").Value = ss1.Text
        ss1.Col = SS1_STDSPEC_STLGRD:    xlSheet.Range("B17").Value = ss1.Text
        ss1.Col = SS1_QT:                xlSheet.Range("B18").Value = ss1.Text
'        ss1.Col = SS1_PLATE_CON:         xlSheet.Range("B19").Value = ss1.MaxRows
        ss1.Col = SS1_PLATE_SIZE:        xlSheet.Range("B20").Value = ss1.Text
        ss1.Col = SS1_APLY_STDSPEC:      xlSheet.Range("B22").Value = ss1.Text



         ss1.Col = SS1_OUT_SHEET_NO:      xlSheet.Range("D1").Value = ss1.Text
         ss1.Col = SS1_Shift:             xlSheet.Range("D2").Value = ss1.Text
         ss1.Col = SS1_TRNS_CMPY_CD:      xlSheet.Range("D3").Value = ss1.Text
         ss1.Col = SS1_CUST_CD1:          xlSheet.Range("D4").Value = ss1.Text
'
'        ss1.Col = SS1_THK_AVG:           xlSheet.Range("D6").Value = ss1.Text
'        ss1.Col = SS1_WID_AVG:           xlSheet.Range("D7").Value = ss1.Text
'        ss1.Col = SS1_LEN_AVG:           xlSheet.Range("D8").Value = ss1.Text
         ss1.Col = SS1_SIZE_KND:          xlSheet.Range("D9").Value = ss1.Text
        ss1.Col = SS1_RM_CR_STAGE3_TIME: xlSheet.Range("D10").Value = ss1.Text
        ss1.Col = SS1_SURFACE_REQUESTS:  xlSheet.Range("D11").Value = ss1.Text
'        ss1.Col = SS1_CE_APPR_FL:        xlSheet.Range("D12").Value = ss1.Text
''
'        ss1.Col = SS1_WGT:               xlSheet.Range("D13").Value = ss1.Text
        ss1.Col = SS1_VESSEL_NO:         xlSheet.Range("D14").Value = ss1.Text
        ss1.Col = SS1_SIDEMARK:          xlSheet.Range("D15").Value = ss1.Text
        ss1.Col = SS1_PAINTNUM:          xlSheet.Range("D16").Value = ss1.Text
        ss1.Col = SS1_GANGYIN:           xlSheet.Range("D17").Value = ss1.Text
        ss1.Col = SS1_PUNCH:             xlSheet.Range("D18").Value = ss1.Text

       ss1.Col = SS1_CUST:              xlSheet.Range("D23").Value = ss1.Text
        
       
'
        
'
'        xlSheet.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'
'        xlSheet.Application.Visible = True
        
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        
        
        xlApp.Range("A1:D23").Select
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
            
        Next i
        
            Set xlSheet = Nothing
            Set xlBook = Nothing
            Set xlApp = Nothing
            
        End With
        
        Exit Sub
    
Excel_Error:

'    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel" & Error, "W")

End Sub

Private Sub text_cur_inv_code_Change()
        If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
            Exit Sub
        Else
            text_cur_inv.Text = ""
        End If
End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
    
        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        
    
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
            Exit Sub
        Else
            text_cur_inv.Text = ""
        End If
    
    End If
    
End Sub

Private Sub txt_upd_cur_inv_code_Change()
        If Len(Trim(txt_upd_cur_inv_code.Text)) = txt_upd_cur_inv_code.MaxLength Then
            txt_upd_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_upd_cur_inv_code.Text, 2)
            Exit Sub
        Else
            txt_upd_cur_inv.Text = ""
        End If
End Sub

Private Sub txt_upd_cur_inv_code_DblClick()

    Call txt_upd_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_upd_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
    
        DD.rControl.Add Item:=txt_upd_cur_inv_code
        DD.rControl.Add Item:=txt_upd_cur_inv
        
    
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_upd_cur_inv_code.Text)) = txt_upd_cur_inv_code.MaxLength Then
            txt_upd_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_upd_cur_inv_code.Text, 2)
            Exit Sub
        Else
            txt_upd_cur_inv.Text = ""
        End If
    
    End If
    
End Sub

Private Sub txt_PLT_Change()
    If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
        txt_plt_nm.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_nm.Text = ""
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
        DD.rControl.Add Item:=txt_plt_nm

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

End Sub

Private Sub txt_prod_cd_Change()
    If Len(Trim(txt_prod_cd)) = txt_prod_cd.MaxLength Then
        txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
    Else
        txt_prod_cd_name.Text = ""
    End If
End Sub

Private Sub txt_prod_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_cd_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

'    If Len(Trim(txt_prod_cd)) = txt_prod_cd.MaxLength Then
'        txt_prod_cd_name.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
'    Else
'        txt_prod_cd_name.Text = ""
'    End If

End Sub

Private Sub Txt_rep_typ_DblClick()

    Call Txt_rep_typ_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Txt_rep_typ_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0010"
        DD.rControl.Add Item:=Txt_rep_typ
        DD.rControl.Add Item:=Txt_rep_typ_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    If Len(Trim(Txt_rep_typ)) = Txt_rep_typ.MaxLength Then
        Txt_rep_typ_name.Text = Gf_ComnNameFind(M_CN1, "C0010", Trim(Txt_rep_typ.Text), 2)
    Else
       Txt_rep_typ_name.Text = ""
    End If
    
End Sub

