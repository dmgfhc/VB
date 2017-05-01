VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AFL2101C 
   Caption         =   "板坯统计情况综合查询_AFL2101C"
   ClientHeight    =   9225
   ClientLeft      =   870
   ClientTop       =   2520
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_cond 
      Height          =   315
      Left            =   9630
      TabIndex        =   29
      Top             =   660
      Visible         =   0   'False
      Width           =   420
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   855
      Left            =   105
      TabIndex        =   0
      Top             =   1140
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   1508
      _Version        =   196609
      Font3D          =   1
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "汇总字段"
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "连铸机号"
         Height          =   255
         Index           =   14
         Left            =   9420
         TabIndex        =   38
         Tag             =   ",A.PRC_LINE"
         Top             =   240
         Width           =   1290
      End
      Begin VB.TextBox txt_Disp 
         Height          =   315
         Left            =   9015
         TabIndex        =   19
         Top             =   495
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "切割日期"
         Height          =   255
         Index           =   5
         Left            =   6315
         TabIndex        =   16
         Tag             =   ",A.PROD_DATE"
         Top             =   240
         Width           =   1020
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "班次"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   9
         Left            =   2790
         TabIndex        =   15
         Tag             =   ",SHIFT"
         Top             =   525
         Width           =   780
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "班别"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   10
         Left            =   3945
         TabIndex        =   14
         Tag             =   ",GROUP_CD"
         Top             =   525
         Width           =   1020
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "钢种"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Tag             =   ",Gf_Stlgrd_Detail(A.STLGRD)"
         Top             =   240
         Width           =   1020
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "厚度"
         Height          =   255
         Index           =   2
         Left            =   2790
         TabIndex        =   12
         Tag             =   ",A.THK"
         Top             =   240
         Width           =   1020
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "宽度"
         Height          =   255
         Index           =   3
         Left            =   3945
         TabIndex        =   11
         Tag             =   ",A.WID"
         Top             =   240
         Width           =   1020
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "长度"
         Height          =   255
         Index           =   4
         Left            =   5100
         TabIndex        =   10
         Tag             =   ",A.LEN"
         Top             =   240
         Width           =   1020
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "炉号"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   9
         Tag             =   ",SUBSTR(A.SLAB_NO,1,8)"
         Top             =   240
         Width           =   1020
      End
      Begin VB.TextBox txt_Order 
         Height          =   450
         Left            =   11985
         TabIndex        =   8
         Top             =   90
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.TextBox txt_Disp_Order 
         Enabled         =   0   'False
         Height          =   570
         Left            =   10800
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   195
         Width           =   4080
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "去向"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   420
         TabIndex        =   6
         Tag             =   ",A.OUT_PLT"
         Top             =   525
         Width           =   705
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "堆放仓库"
         Height          =   255
         Index           =   6
         Left            =   7995
         TabIndex        =   5
         Tag             =   ",A.CUR_INV"
         Top             =   240
         Width           =   1020
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "转库日"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   11
         Left            =   5100
         TabIndex        =   4
         Tag             =   ",Gf_AFL2100C_DATE(A.SLAB_NO,'MOVE')"
         Top             =   525
         Width           =   870
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "转炉出钢日"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   12
         Left            =   6315
         TabIndex        =   3
         Tag             =   ",Gf_AFL2100C_DATE(A.SLAB_NO,'BOF')"
         Top             =   525
         Width           =   1290
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "轧制日期"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   2
         Tag             =   ",A.OUT_PLT_DATE"
         Top             =   525
         Width           =   1065
      End
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "发货日"
         Height          =   255
         Index           =   13
         Left            =   7995
         TabIndex        =   1
         Tag             =   ",A.SHP_DATE"
         Top             =   525
         Width           =   1290
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7125
      Left            =   60
      TabIndex        =   17
      Top             =   2055
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   12568
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "汇总信息"
      TabPicture(0)   =   "AFL2101C.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ss1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "详细信息"
      TabPicture(1)   =   "AFL2101C.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "ss2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin FPSpread.vaSpread ss1 
         Height          =   6690
         Left            =   -74940
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Width           =   14955
         _Version        =   393216
         _ExtentX        =   26379
         _ExtentY        =   11800
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         MaxRows         =   2
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AFL2101C.frx":0038
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   6690
         Left            =   60
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   360
         Width           =   14955
         _Version        =   393216
         _ExtentX        =   26379
         _ExtentY        =   11800
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   57
         MaxRows         =   1
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AFL2101C.frx":07B6
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1020
      Left            =   90
      TabIndex        =   20
      Top             =   60
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   1799
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_ccm_line 
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
         ItemData        =   "AFL2101C.frx":2205
         Left            =   3900
         List            =   "AFL2101C.frx":2207
         TabIndex        =   26
         Tag             =   "连铸机号"
         Top             =   525
         Width           =   615
      End
      Begin VB.TextBox txt_slab_no_to 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   7620
         MaxLength       =   10
         TabIndex        =   25
         Top             =   150
         Width           =   1200
      End
      Begin VB.ComboBox cbo_prc_line 
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
         ItemData        =   "AFL2101C.frx":2209
         Left            =   1635
         List            =   "AFL2101C.frx":220B
         TabIndex        =   24
         Tag             =   "炉座号"
         Top             =   525
         Width           =   615
      End
      Begin VB.TextBox txt_slab_no 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   6390
         MaxLength       =   10
         TabIndex        =   23
         Top             =   150
         Width           =   1200
      End
      Begin VB.TextBox txt_cur_inv_code 
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
         Left            =   6390
         MaxLength       =   2
         TabIndex        =   22
         Top             =   525
         Width           =   375
      End
      Begin VB.TextBox txt_cur_inv 
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
         Left            =   6765
         TabIndex        =   21
         Top             =   525
         Width           =   2625
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   210
         Top             =   150
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "入库日(生产)"
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
      Begin InDate.UDate txt_DateFrom 
         Height          =   315
         Left            =   1635
         TabIndex        =   27
         Top             =   150
         Width           =   1425
         _ExtentX        =   2514
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
      Begin InDate.UDate txt_DateTo 
         Height          =   315
         Left            =   3105
         TabIndex        =   28
         Top             =   150
         Width           =   1425
         _ExtentX        =   2514
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   4965
         Top             =   525
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "堆放仓库"
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
         Left            =   4965
         Top             =   150
         Width           =   1395
         _ExtentX        =   2461
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
      Begin InDate.ULabel ULabel6 
         Height          =   315
         Left            =   210
         Top             =   525
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "炉座号"
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
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   2460
         Top             =   525
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         Caption         =   "连铸机号"
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
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   1020
      Left            =   9780
      TabIndex        =   30
      Top             =   60
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   1799
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.OptionButton opt_Search 
         BackColor       =   &H00E0E0E0&
         Caption         =   "发货日"
         Height          =   195
         Index           =   4
         Left            =   2310
         TabIndex        =   36
         Top             =   555
         Width           =   945
      End
      Begin VB.OptionButton opt_Search 
         BackColor       =   &H00E0E0E0&
         Caption         =   "转炉出钢日"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   35
         Top             =   555
         Width           =   1395
      End
      Begin VB.OptionButton opt_Search 
         BackColor       =   &H00E0E0E0&
         Caption         =   "转库日"
         Height          =   195
         Index           =   2
         Left            =   3810
         TabIndex        =   34
         Top             =   210
         Width           =   1185
      End
      Begin VB.OptionButton opt_Search 
         BackColor       =   &H00E0E0E0&
         Caption         =   "连铸开浇日"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   33
         Top             =   210
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton opt_Search 
         BackColor       =   &H00E0E0E0&
         Caption         =   "轧制日"
         Height          =   240
         Index           =   1
         Left            =   2310
         TabIndex        =   32
         Top             =   210
         Width           =   960
      End
      Begin VB.OptionButton opt_Search 
         BackColor       =   &H00E0E0E0&
         Caption         =   "当前库存"
         Height          =   195
         Index           =   5
         Left            =   3810
         TabIndex        =   31
         Top             =   555
         Width           =   1185
      End
   End
End
Attribute VB_Name = "AFL2101C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       SLAB DATAS QUERY
'-- Sub_System Name
'-- Program Name
'-- Program ID        AFL2101C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder             GUOLI
'-- Date              2009.11.11
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-- yidujun    2010-10-18  汇总字段添加按“连铸机号”选框查询
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
Dim nColumn1 As New Collection      'Spread Necessary Column Collection
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
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iSumCol As New Collection       'Sum Column

Dim iSumCol1  As New Collection       'Sum Column
Dim iSumCol2  As New Collection       'Sum Column


Const iss1MaxCols = 10
Const SS2_INGOT_FL = 2                 'INGOT_FL
Const SS2_URGNT_FL = 5                 'URGNT_FL
Const SS2_PLAN_STLGRD = 7              'PLAN_STLGRD
Const SS2_PLAN_STLGRD_DET = 8          'PLAN_STLGRD_DET
Const SS2_CC_STLGRD = 9                'CC_STLGRD
Const SS2_CC_STLGRD_DET = 10            'CC_STLGRD_DET
Const SS2_REASON_CD = 13               'REASON_CD
Const SS2_EST_CD = 14                  'EST_CD
Const SS2_WGT = 22                     'WGT
Const SS2_SIZE_WGT = 23                'SIZE_WGT
Const SS2_SIZE_WGT1 = 24               'SIZE_WGT
Const SS2_HEAD_SLAB_WID = 25           'HEAD_SLAB_WID
Const SS2_MIXED_FL = 27                'MIXED_FL
Const SS2_STLGRD_UPD_FL = 28           'STLGRD_UPD_FL
Const SS2_OVER_FL = 29                 'OVER_FL
Const SS2_RESNM = 33                   'RESNM






Private Sub Form_Define()

    Dim iIndex As Integer
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_DateFrom, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_DateTo, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_slab_no_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(cbo_prc_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(cbo_ccm_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(txt_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_cond, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_Order, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
      
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
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 33, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 34, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 35, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 36, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 37, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 38, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 39, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 40, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 41, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 42, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 43, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 44, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 45, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 46, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 47, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 48, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 49, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 50, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 51, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 52, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 53, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 54, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 55, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 56, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 57, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

'
'
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFL2101C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFL2101C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call opt_Search_Click(0)
    
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

    Dim i As Integer

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "F-System.INI", Me.Name)
        
    opt_Search(0).VALUE = True
    txt_cond.Text = 0
    
    cbo_prc_line.AddItem "1"
    cbo_prc_line.AddItem "2"
    cbo_prc_line.AddItem "3"
    
    cbo_ccm_line.AddItem "1"
    cbo_ccm_line.AddItem "2"
    cbo_ccm_line.AddItem "3"

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    
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
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()
    Dim iCol As Integer
        
    If Gf_Sp_Cls(sc1) Then
        For iCol = 0 To 14
            chk_Cond(iCol).VALUE = ssCBUnchecked
        Next iCol
        txt_Disp_Order = ""
        txt_Order = ""
        txt_Disp = ""
        Call Gf_Sp_Cls(sc2)
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
End Sub

Public Sub Form_Exc()
    If SSTab1.Tab = 0 Then
        Call Gp_Sp_Excel(Me, sc1.Item("Spread"), 0, 0, 0, 0)
    ElseIf SSTab1.Tab = 1 Then
        Call Gp_Sp_Excel(Me, sc2.Item("Spread"), 0, 0, 0, 0)
    End If
End Sub

Public Sub Form_Ref()

    Dim sQuery      As String
    Dim dSlabwgt    As Double
    Dim dProdwgt    As Double
    Dim dOkwgt      As Double
    Dim iIdx        As Integer
    Dim iCol        As Integer
    
    Dim iRow        As Integer
    Dim i        As Integer
        
On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    Select Case SSTab1.Tab
    
           Case 0
            
                Call Display_ss1_Set
           
                sQuery = Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", pControl)
                If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, 0, iSumCnt, iSumCol) Then
                        For iIdx = 1 To ss1.MaxRows
                            ss1.Row = iIdx
                            For iCol = ss1.MaxCols - iSumCnt + 1 To ss1.MaxCols
                                ss1.Col = iCol
                                If Val(ss1.Text & "") = 0 Then
                                    ss1.Text = ""
                                End If
                            Next iCol
                        Next iIdx
                        
                        If ss1.MaxCols = iSumCnt Then
                           ss1.Col = 0:   ss1.Row = ss1.MaxRows:    ss1.Text = "合计"
                        End If
                        
                        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                        
                        ss1.OperationMode = OperationModeNormal
                End If
           
           Case 1
                
                'Sum Column Count
                iSumCnt = 8
                Set iSumCol = Nothing
                'Sum Column Setting
                iSumCol.Add Item:=SS2_WGT
                iSumCol.Add Item:=SS2_SIZE_WGT
                iSumCol.Add Item:=SS2_SIZE_WGT1
                iSumCol.Add Item:=SS2_HEAD_SLAB_WID

                iSumCol.Add Item:=SS2_MIXED_FL
                iSumCol.Add Item:=SS2_STLGRD_UPD_FL
                iSumCol.Add Item:=SS2_OVER_FL
                iSumCol.Add Item:=SS2_RESNM
                
                sQuery = Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", pControl)
                If Gf_Total_Display(M_CN1, Proc_Sc("Sc2"), sQuery, 0, iSumCnt, iSumCol) Then
                    
                    For iIdx = 1 To ss2.MaxRows
                        ss2.Row = iIdx
                        For iCol = SS2_SIZE_WGT To SS2_HEAD_SLAB_WID
                            ss2.Col = iCol
                            If Val(ss2.Text & "") = 0 Then
                                ss2.Text = ""
                            End If
                        Next iCol

                        For iCol = SS2_MIXED_FL To SS2_OVER_FL
                            ss2.Col = iCol
                            If Val(ss2.Text & "") = 0 Then
                                ss2.Text = ""
                            End If
                        Next iCol
                    Next iIdx
                    
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                    
                    ss2.OperationMode = OperationModeNormal
                End If
            
    End Select
    
    For iRow = 1 To ss2.MaxRows
    
               ss2.Row = iRow
               ss2.Col = SS2_URGNT_FL
                If ss2.Text = "Y" Then
                  For i = 1 To ss2.MaxCols
                       ss2.Col = i
                       ss2.ForeColor = &HC000&
                  Next
                End If

      
     Next iRow


    Exit Sub

Refer_Err:
    
End Sub

Private Sub Display_ss1_Set()
    Dim sSelCol     As String
    Dim iCol        As Integer
    Dim iIdx        As Integer
    Dim iInsCnt     As Integer
       
    ss1.DeleteCols 1, ss1.MaxCols - iss1MaxCols
    ss1.MaxCols = iss1MaxCols
    ss1.MaxRows = 0
    
    sSelCol = Trim(txt_Disp.Text)
    
    If sSelCol <> "" Then
        For iCol = 1 To Len(sSelCol) Step 2
            iInsCnt = iInsCnt + 1
            iIdx = Mid(sSelCol, iCol, 2)
            
            ss1.MaxCols = ss1.MaxCols + 1
            ss1.InsertCols ss1.MaxCols - iss1MaxCols, 1
            ss1.Col = ss1.MaxCols - iss1MaxCols
            ss1.Row = 0
            ss1.Text = chk_Cond(iIdx).Caption
        Next iCol
    End If
    
    'Sum Column Count
    iSumCnt = 10
    Set iSumCol = Nothing
    'Sum Column Setting
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 1
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 2
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 3
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 4
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 5
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 6
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 7
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 8
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 9
    iSumCol.Add Item:=ss1.MaxCols - iSumCnt + 10
    
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

Private Sub opt_Search_Click(Index As Integer)
    txt_cond.Text = Index
    If Index = 5 Then
       txt_DateFrom.Enabled = False
       txt_DateTo.Enabled = False
       txt_cur_inv_code = "00"
       Call Gp_Sp_ColHidden(ss2, SS2_INGOT_FL, False)
       Call Gp_Sp_ColHidden(ss2, SS2_PLAN_STLGRD, False)
       Call Gp_Sp_ColHidden(ss2, SS2_PLAN_STLGRD_DET, False)
       Call Gp_Sp_ColHidden(ss2, SS2_CC_STLGRD, False)
       Call Gp_Sp_ColHidden(ss2, SS2_CC_STLGRD_DET, False)
       Call Gp_Sp_ColHidden(ss2, SS2_REASON_CD, False)
       Call Gp_Sp_ColHidden(ss2, SS2_EST_CD, False)
    Else
       txt_DateFrom.Enabled = True
       txt_DateTo.Enabled = True
       Call Gp_Sp_ColHidden(ss2, SS2_INGOT_FL, True)
       Call Gp_Sp_ColHidden(ss2, SS2_PLAN_STLGRD, True)
       Call Gp_Sp_ColHidden(ss2, SS2_PLAN_STLGRD_DET, True)
       Call Gp_Sp_ColHidden(ss2, SS2_CC_STLGRD, True)
       Call Gp_Sp_ColHidden(ss2, SS2_CC_STLGRD_DET, True)
       Call Gp_Sp_ColHidden(ss2, SS2_REASON_CD, True)
       Call Gp_Sp_ColHidden(ss2, SS2_EST_CD, True)
    End If
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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)

    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

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

Private Sub ss2_LostFocus()

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

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub chk_Cond_Click(Index As Integer)

    Dim Ord_Index As Integer

    If chk_Cond(Index) Then
        txt_Disp_Order = Trim(txt_Disp_Order & " " & chk_Cond(Index).Caption)
        txt_Order = Trim(txt_Order & chk_Cond(Index).Tag)
        txt_Disp = Trim(txt_Disp & Format(Index, "0#"))
    Else
        txt_Disp_Order = Trim(Replace(txt_Disp_Order, chk_Cond(Index).Caption, ""))
        txt_Order = Trim(Replace(txt_Order, chk_Cond(Index).Tag, ""))
        txt_Disp = Trim(Replace(txt_Disp, Format(Index, "0#"), ""))
    End If
    
End Sub

Private Sub txt_cur_inv_code_Change()
    If Len(Trim(txt_cur_inv_code.Text)) = txt_cur_inv_code.MaxLength Then
          txt_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv_code.Text, 2)
          Exit Sub
    Else
          txt_cur_inv.Text = ""
    End If

End Sub

Private Sub txt_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=txt_cur_inv_code
        DD.rControl.Add Item:=txt_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
       
        If Len(Trim(txt_cur_inv_code.Text)) = txt_cur_inv_code.MaxLength Then
            txt_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv_code.Text, 2)
            Exit Sub
        Else
            txt_cur_inv.Text = ""
        End If
    End If
End Sub

