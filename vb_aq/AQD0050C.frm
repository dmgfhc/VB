VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQD0050C 
   Caption         =   "船板质量证明书编制_AQD0050C"
   ClientHeight    =   9255
   ClientLeft      =   -180
   ClientTop       =   2055
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15345
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txt_OutOrder 
      Height          =   270
      Left            =   8520
      TabIndex        =   22
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.CheckBox Chk_OutOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "出口"
      Height          =   315
      Left            =   8160
      TabIndex        =   21
      Top             =   540
      Width           =   735
   End
   Begin VB.ComboBox cbo_STDSPEC 
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
      ItemData        =   "AQD0050C.frx":0000
      Left            =   6975
      List            =   "AQD0050C.frx":0002
      TabIndex        =   20
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "录入控制号"
      Enabled         =   0   'False
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
      Left            =   9285
      TabIndex        =   19
      Top             =   90
      Width           =   1545
   End
   Begin VB.CommandButton cmd_AllCheck 
      Caption         =   "全部确认"
      Height          =   300
      Left            =   1680
      TabIndex        =   16
      Top             =   1800
      Width           =   1275
   End
   Begin VB.TextBox txt_plt 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1125
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "plt"
      Top             =   90
      Width           =   600
   End
   Begin VB.CommandButton CMDCANCEL 
      Caption         =   "取消质量证明书"
      Enabled         =   0   'False
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
      Left            =   12540
      TabIndex        =   6
      Top             =   90
      Width           =   1545
   End
   Begin VB.TextBox TXT_SMP_LIST 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9270
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   505
      Width           =   5865
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "编制质量证明书"
      Enabled         =   0   'False
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
      Left            =   10920
      TabIndex        =   5
      Top             =   90
      Width           =   1545
   End
   Begin VB.TextBox txt_SMP_NO 
      Height          =   345
      Left            =   13230
      TabIndex        =   8
      Top             =   -105
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox txt_TEST_STS 
      Height          =   315
      Left            =   12705
      TabIndex        =   7
      Text            =   "A"
      Top             =   -90
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt_INSP_CD 
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
      Left            =   3270
      TabIndex        =   1
      Top             =   90
      Width           =   645
   End
   Begin VB.TextBox txt_CONTROL_NO 
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
      Left            =   1125
      TabIndex        =   2
      Top             =   540
      Width           =   2265
   End
   Begin VB.TextBox txt_STD_ORGAN_NAME 
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
      Left            =   3960
      MaxLength       =   14
      TabIndex        =   15
      Top             =   90
      Width           =   1575
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   5925
      Top             =   90
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "标准编号"
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
      Index           =   1
      Left            =   2220
      Top             =   90
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "检查机关"
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
      Left            =   120
      Top             =   540
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "控制号"
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
      Left            =   120
      Top             =   1800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "船检取样选择"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   8925
      Top             =   1800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      Caption         =   "              号包含产品"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   7425
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   8790
      _Version        =   393216
      _ExtentX        =   15505
      _ExtentY        =   13097
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
      MaxCols         =   13
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQD0050C.frx":0004
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   7545
      Left            =   8955
      TabIndex        =   10
      Top             =   2160
      Width           =   6180
      _Version        =   393216
      _ExtentX        =   10901
      _ExtentY        =   13309
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
      MaxCols         =   6
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQD0050C.frx":06D6
   End
   Begin Threed.SSOption opt_TEST_STS 
      Height          =   315
      Index           =   0
      Left            =   3600
      TabIndex        =   12
      Top             =   540
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   196609
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
      Caption         =   "未生成质保书"
      Value           =   -1
   End
   Begin Threed.SSOption opt_TEST_STS 
      Height          =   315
      Index           =   1
      Left            =   5280
      TabIndex        =   13
      Top             =   540
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      _Version        =   196609
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
      Caption         =   "已生成质保书"
   End
   Begin Threed.SSOption opt_TEST_STS 
      Height          =   315
      Index           =   2
      Left            =   7080
      TabIndex        =   14
      Top             =   540
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      _Version        =   196609
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
      Caption         =   "全部"
   End
   Begin InDate.UDate dtp_fr_date 
      Height          =   315
      Left            =   1125
      TabIndex        =   3
      Tag             =   "发放日期"
      Top             =   900
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
   Begin InDate.UDate dtp_to_date 
      Height          =   315
      Left            =   2550
      TabIndex        =   4
      Tag             =   "发放日期"
      Top             =   900
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   120
      Top             =   900
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "剪切日期"
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
   Begin InDate.ULabel ULabel7 
      Height          =   300
      Left            =   120
      Top             =   90
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   529
      Caption         =   "工厂"
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
      ForeColor       =   16711680
   End
   Begin InDate.UDate dtp_DSC_DATE_FR 
      Height          =   315
      Left            =   5490
      TabIndex        =   17
      Tag             =   "发放日期"
      Top             =   900
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
   Begin InDate.UDate dtp_DSC_DATE_TO 
      Height          =   315
      Left            =   6915
      TabIndex        =   18
      Tag             =   "发放日期"
      Top             =   900
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   4440
      Top             =   900
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "判定日期"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   12720
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      Caption         =   "钢板重量(吨)："
      Alignment       =   0
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   120
      Top             =   1280
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "实验日期"
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
   Begin InDate.UDate ins_fr_date 
      Height          =   315
      Left            =   1125
      TabIndex        =   23
      Tag             =   "指示日期"
      Top             =   1280
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
   Begin InDate.UDate ins_to_date 
      Height          =   315
      Left            =   2550
      TabIndex        =   24
      Tag             =   "指示日期"
      Top             =   1280
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   15120
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "AQD0050C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      质量证明书(船板)
'-- Program ID        AQD0050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Chu Kyo Su
'-- Coder             Chu Kyo Su
'-- Date              2003.07. 25
'-- Description       质量证明书(船板)
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
Public sPLT_Authority As String     'Active User Plant Authority Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim COMB As New Collection          'ComboBox Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection
Dim bPrintCheck As Boolean

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

'---------------------------------------------------------------------------------------------
'------------------------------ Report Variable ----------------------------------------------
'---------------------------------------------------------------------------------------------
'Dim crxApplication As New CRAXDRT.Application
'
'Public WithEvents Report As CRAXDRT.Report
'
'Dim crxDatabaseTable As CRAXDRT.DatabaseTable
'Dim crxSubreport As CRAXDRT.Report
''Dim CPProperties As CRAXDRT.ConnectionProperties
Dim cVal As New Collection
Dim sVal As New Collection
Dim sQueryHeadC As String        'QP_CERT_HEAD   -C  QUERY
Dim sQueryDetailC As String      'QP_CERT_DETAIL - C QUERY
Dim sQueryHeadS As String        'QP_CERT_HEAD   -S  QUERY
Dim sQueryDetailS As String      'QP_CERT_DETAIL - S QUERY
Const SS1_CONTROL_NO = 6 '  控制号
Const SS1_USERID = 11  ' 用户名

'---------------------------------------------------------------------------------------------

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_INSP_CD, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_STD_ORGAN_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(cbo_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_CONTROL_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_TEST_STS, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(dtp_fr_date, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(dtp_to_date, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(dtp_DSC_DATE_FR, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(dtp_DSC_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(Txt_OutOrder, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)  '添加出口查询20130328
         Call Gp_Ms_Collection(ins_fr_date, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(ins_to_date, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)

   
    'MASTER Collection
    Mc2.Add Item:=pControl1, Key:="pControl"
    Mc2.Add Item:=nControl1, Key:="nControl"
    Mc2.Add Item:=mControl1, Key:="mControl"
    Mc2.Add Item:=iControl1, Key:="iControl"
    Mc2.Add Item:=rControl1, Key:="rControl"
    Mc2.Add Item:=cControl1, Key:="cControl"
    Mc2.Add Item:=aControl1, Key:="aControl"
    Mc2.Add Item:=lControl1, Key:="lControl"
    
        'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_INSP_CD, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(txt_STD_ORGAN_NAME, " ", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(cbo_STDSPEC, "p", " ", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_CONTROL_NO, "p", " ", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(txt_TEST_STS, "p", " ", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(TXT_SMP_LIST, "p", " ", " ", "i", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc3.Add Item:="AQD0050C.P_MODIFY_MC", Key:="P-M"
    Mc3.Add Item:=pControl2, Key:="pControl"
    Mc3.Add Item:=nControl2, Key:="nControl"
    Mc3.Add Item:=mControl2, Key:="mControl"
    Mc3.Add Item:=iControl2, Key:="iControl"
    Mc3.Add Item:=rControl2, Key:="rControl"
    Mc3.Add Item:=cControl2, Key:="cControl"
    Mc3.Add Item:=aControl2, Key:="aControl"
    Mc3.Add Item:=lControl2, Key:="lControl"

    
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQD0050C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQD0050C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQD0050C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
        
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
     Call Gp_Sp_Collection(SS2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(SS2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(SS2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(SS2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(SS2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(SS2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=SS2, Key:="Spread"
    sc2.Add Item:="AQD0050C.P_REFER_D", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=SS2.MaxCols, Key:="Last"
    
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Chk_OutOrder_Click()
    If Chk_OutOrder.Value = ssCBChecked Then
        Txt_OutOrder = "1"
    Else
        Txt_OutOrder = "0"
    End If
End Sub

Private Sub cmd_AllCheck_Click()
    Dim i       As Integer
    Dim sAllChk As String
    
    If ss1.MaxRows < 1 Or ss1.Row = 0 Then Exit Sub
    
    If cmd_AllCheck.Caption = "全部确认" Then
        sAllChk = "ALL"
    Else
        sAllChk = ""
    End If
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        
        For i = 1 To ss1.MaxRows
            ss1.Row = i
            If sAllChk = "ALL" Then
                ss1.Col = 1
                ss1.Text = 1
                ss1.Col = 0
                ss1.Text = "Update"
                cmd_AllCheck.Caption = "全部取消"
            Else
                ss1.Col = 1
                ss1.Text = 0
                ss1.Col = 0
                ss1.Text = ""
                cmd_AllCheck.Caption = "全部确认"
            End If
        Next i
              
    End If

End Sub

Private Sub CMDCANCEL_Click()

    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If

    Call Form_Pro
    If txt_TEST_STS = "B" And txt_CONTROL_NO <> "" Then
         If Gf_Ms_Process1(M_CN1, Mc3, sAuthority) = False Then Exit Sub
    Else
       MsgBox ("录入控制号请选择已生成质保书选项，再输入控制号！")
    End If

End Sub

Private Sub CmdInput_Click()

    Dim iRow As Integer
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 9
        ss1.BackColor = &H808080
        'Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H808080)                        ' 红色
    Next iRow

    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If
    
    Call Form_Pro
    If txt_TEST_STS = "A" And txt_CONTROL_NO <> "" Then
         If Gf_Sp_Process1(M_CN1, Sc1, Mc1) = False Then Exit Sub
    Else
       MsgBox ("录入控制号请选择未生成质保书选项，再输入控制号！")
    End If
    

End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    Call subButtonHide
    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
           
        Case "txt_INSP_CD"
            sCode = "Q0052"
            Set oCodeName = txt_STD_ORGAN_NAME
    
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
        
'    sAuthority = Gf_Pgm_Authority(Me.Name, True)
'
'    If sAuthority = "1000" Then
'
'       CmdInput.Visible = False
'       cmdReport.Visible = False
'       CMDCANCEL.Visible = False
'
'    End If
    
    sPLT_Authority = Gf_PLT_Authority(Me.Name)
    
    sPLT_Authority = "**"
    
    If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
       txt_plt.Text = sPLT_Authority
    Else
       txt_plt.Text = ""
    End If
    
    If Gf_Sc_Authority(sAuthority, "U") Then
       cmdReport.Enabled = True
       CmdInput.Enabled = True
       CMDCANCEL.Enabled = True
       
    End If

    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"))
    
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(Sc1)
    
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "Q-System.INI", Me.Name)
    
    txt_TEST_STS.Text = "A"
'    opt_TEST_STS(0).Value = True
    Txt_OutOrder.Visible = False

    Screen.MousePointer = vbDefault
    
    Call subButtonHide
    
    

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "Q-System.INI", Me.Name)
    
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
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
    Call subButtonHide
    
End Sub



Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gf_Sp_Cls(sc2)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
     '  rControl(1).SetFocus
        If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
           txt_plt.Text = sPLT_Authority
        Else
           txt_plt.Text = ""
        End If
    End If
    txt_INSP_CD = ""
    cbo_STDSPEC.Clear
    txt_CONTROL_NO = ""
    txt_TEST_STS = "A"
    txt_SMP_NO = ""
    TXT_SMP_LIST = ""
    dtp_fr_date.RawData = ""
    dtp_to_date.RawData = ""
'    dtp_fr_date.Visible = False
'    dtp_to_date.Visible = False
'    ULabel5.Visible = False
    ULabel4.Caption = "       号包含产品"
'    opt_TEST_STS(0).Value = True
    Call subMasterClear

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc1").Item("Spread")) Then Exit Sub
                
'    If txt_TEST_STS <> "A" Then
       If dtp_fr_date.RawData = "" Then
          dtp_fr_date.RawData = Format(Now, "yyyymm") + "01"
       End If
       If dtp_to_date.RawData = "" Then
          dtp_to_date.RawData = Format(Now, "yyyymmdd")
       End If
       
'    End If
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call subButtonHide
        Exit Sub
    End If

    Call subButtonHide
    
    bPrintCheck = False
    
'    Dim iRow As Integer
'    For iRow = 1 To ss1.MaxRows
'        ss1.Row = iRow
'        ss1.Col = 9
'        Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H808080)                        ' 红色
'    Next iRow

    
    Exit Sub
   

Refer_Err:

End Sub


Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub


Private Sub opt_TEST_STS_Click(Index As Integer, Value As Integer)
        Select Case Index
        Case 0
            txt_TEST_STS.Text = "A"
            ULabel5.Visible = True
            ULabel5.Caption = "剪切日期"
            dtp_fr_date.Visible = True
            dtp_to_date.Visible = True
'            dtp_fr_date.RawData = ""
'            dtp_to_date.RawData = ""

        Case 1
            txt_TEST_STS.Text = "B"
            ULabel5.Visible = True
            ULabel5.Caption = "生成日期"
            dtp_fr_date.Visible = True
            dtp_to_date.Visible = True
        Case 2
            txt_TEST_STS.Text = "C"
            ULabel5.Visible = True
            ULabel5.Caption = "生成日期"
            dtp_fr_date.Visible = True
            dtp_to_date.Visible = True

    End Select
    Call Form_Ref

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    
    If Row <= 0 Then Call Gp_Sp_Sort(Proc_Sc("Sc1")("Spread"), Col, Row)
    
    If Row >= 1 Then
        With ss1
            .Row = Row
            .Col = 3
                If Trim(txt_SMP_NO.Text) <> Trim(.Text) Then
                    txt_SMP_NO.Text = .Text
                    ULabel4.Caption = .Text + "号包含产品"
                     Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"))
                End If
        End With
    End If
    
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    '计算该试样号下钢板总重量------------------------
    Dim WGT As String
    Dim sQuery As String
    Dim arrRecords As Variant
    Dim AdoRs As adodb.Recordset
    
    Set AdoRs = New adodb.Recordset
    
    'WHERE 条件和AQD0050C.PREFER_D 条件保持一致（SS2的查询条件）
    sQuery = "SELECT SUM(A.WGT) FROM GP_PLATE A WHERE A.SMP_NO = '" & txt_SMP_NO.Text & "' AND  A.PROD_CD  =  'PP' "
    sQuery = sQuery + "AND (A.REC_STS  =  '1' OR A.REC_STS  =  '2' OR (A.REC_STS  =  '3' AND A.PROC_CD = 'XAF')) "
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If AdoRs.EOF Then
        AdoRs.Close
    Else
        arrRecords = AdoRs.GetRows
        WGT = arrRecords(0, 0) & ""
        ULabel8.Caption = "钢板重量(吨)：" + WGT
        AdoRs.Close
    End If
    Set AdoRs = Nothing
    '------------------------------------------------

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
'    If Gf_Sc_Authority(sAuthority, "U") Then
'        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 12)
'        With ss1
'            .Row = Row
'            .Col = 1
'            If .Value = 0 Then
'                .Col = 0
'                .Text = ""
'            Else
'                .Col
'            End If
'        End With
'    End If
   
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


Private Sub subButtonHide()

    MDIMain.MenuTool.Buttons(5).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = True     'Row cancel
    
    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
    

End Sub
Private Sub CHECK_ERR()
Dim i As Integer

    With ss1
        For i = 1 To .MaxRows
            .Row = i
            .Col = 0
            If .Text = "Update" Then
               .Col = SS1_CONTROL_NO
                If Trim(.Text) = "" Then
                   .Col = 1
                   .Text = 0
                   .Col = 0
                   .Text = i
                End If
            End If
        Next i
    End With
    

End Sub



'-----------------------------------------------------------------------
'---------------------------- Report Main ------------------------------
'-----------------------------------------------------------------------
Private Sub cmdReport_Click()
    
  Dim sSQL As String
  Dim i As Integer
  Dim j As Integer
  Dim count1, count2, count As Integer
    count1 = 0
    count2 = 0
    count = 0
    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If
    
    Call Form_Pro
    Call CHECK_ERR
    Screen.MousePointer = vbHourglass
     count1 = ss1.MaxRows
 
    If txt_TEST_STS = "A" And txt_CONTROL_NO <> "" Then
        Call ship_proc
        
    Else
        MsgBox ("编制质保书请选择未生成质保书选项，再输入控制号！")
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Call Form_Ref
    count2 = ss1.MaxRows
    count = count1 - count2
     MDIMain.StatusBar1.Panels(1).Text = "提示信息：已处理" & count & "条信息"
    TXT_SMP_LIST = ""
    Screen.MousePointer = vbDefault
   

        
End Sub
Public Function Gf_Ms_Process1(Conn As adodb.Connection, MC As Collection, sAuthority As String) As Boolean

On Error GoTo MasterPro_Error

    Dim II As Integer
    Dim sQuery As String
    Dim sWhere As String
    Dim sMessg As String
    Dim OutParam(2, 4) As Variant
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adInteger
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256

    'Necessarily Check
    sMessg = Gf_Ms_NeceCheck(MC.Item("nControl"))
    
    If Trim(sMessg) <> "OK" Then
        Call Gp_MsgBoxDisplay(Trim(sMessg) + "必须输入", "I")
        Gf_Ms_Process1 = False
        Exit Function
    End If

    'Maxlength Check
    sMessg = Gf_Ms_NeceCheck2(MC.Item("mControl"))
    
    If Trim(sMessg) <> "OK" Then
        Call Gp_MsgBoxDisplay(Trim(sMessg) + "长度不正确", "I")
        Gf_Ms_Process1 = False
        Exit Function
    End If
    
    If MC!pControl.count > 0 And MC.Item("pControl")(1).Enabled = True Then
    
        'Insert Make Query
         sQuery = Gf_Ms_MakeQuery(MC.Item("P-M"), "I", MC.Item("iControl"))
        
        If sQuery = "FAIL" Then
            Call Gp_MsgBoxDisplay("Insert Query Error : " & sErrMessg)
            Gf_Ms_Process1 = False
            Exit Function
        End If
        
        If Gf_Ms_ExecQuery(OutParam, Conn, sQuery) Then

            'sQuery = Gf_Ms_MakeQuery(MC.Item("P-R"), "R", MC.Item("pControl"))

            'If sQuery = "FAIL" Then
             If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl"), Mc1("mControl")) = False Then
'                Call Gp_MsgBoxDisplay("Refer Query Error : " & sErrMessg)
                Gf_Ms_Process1 = False
                Exit Function
            End If

           ' Call Gf_Ms_Display(Conn, sQuery, MC!rControl, MC!lControl)
          '  Call Gp_Ms_ControlLock(MC!pControl, True)
            Gf_Ms_Process1 = True
            MDIMain.StatusBar1.Panels(1) = "提示信息：新增数据成功"
        Else
            Gf_Ms_Process1 = False
            Call Gp_MsgBoxDisplay(sErrMessg)
        End If
        
    Else
    
        If Mid(sAuthority, 3, 1) = "0" Then Gf_Ms_Process1 = True: Exit Function
        
        'Update Make Query
        sQuery = Gf_Ms_MakeQuery(MC.Item("P-M"), "U", MC.Item("iControl"))
        
        If sQuery = "FAIL" Then
            Call Gp_MsgBoxDisplay("Modify Query Error : " & sErrMessg)
            Gf_Ms_Process1 = False
            Exit Function
        End If
        
        If Gf_Ms_ExecQuery(OutParam, Conn, sQuery) Then
        
'            sQuery = Gf_Ms_MakeQuery(MC.Item("P-R"), "R", MC.Item("pControl"))
'
'            If sQuery = "FAIL" Then
'                Call Gp_MsgBoxDisplay("Refer Query Error : " & sErrMessg)
'                Gf_Ms_Process = False
'                Exit Function
'            End If
'
'            Call Gf_Ms_Display(Conn, sQuery, MC!rControl, MC!lControl)
            Gf_Ms_Process1 = True
            MDIMain.StatusBar1.Panels(1) = "提示信息：数据更新成功"
        Else
            Gf_Ms_Process1 = False
            Call Gp_MsgBoxDisplay(sErrMessg)
        End If
        
    End If
    
    Exit Function
    
MasterPro_Error:

    Gf_Ms_Process1 = False
    Call Gp_MsgBoxDisplay("Failed in data processing")

End Function

Private Function ship_proc() As Boolean

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim sMesg As String
    
    
    Dim adoCmd As adodb.Command
    Screen.MousePointer = vbHourglass
    

    OutParam(1, 1) = "arg_CD"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    

    sQuery = "{call AQD0520P ('PP', '" + Trim(txt_INSP_CD.Text) + "','" + Trim(cbo_STDSPEC.Text) + "','" + txt_CONTROL_NO + "','" + TXT_SMP_LIST + "',?,?)}"
'AQD0520P(P_PROD_CD,P_INSP_CD,P_STDSPEC,P_CON_NO)
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Function
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Function



'########################################################################################################################
'################################################### REPORT END #########################################################
'########################################################################################################################


'--------------------------------------------------------------------------------------------------------
'------------------------------------------- Local Procedure --------------------------------------------
'--------------------------------------------------------------------------------------------------------

Private Sub subMasterClear()
'    txt_CERT_NO.Text = ""
'    dtp_fr_date.Text = ""
'    dtp_to_date.Text = ""
'    txt_CUST_CD.Text = ""
'    txt_CUST_NAME.Text = ""
'    txt_ORD_NO.Text = ""
'    txt_PROD_CD.Text = ""
End Sub


Private Sub txt_INSP_CD_Change()

Dim sSQL As String

    
   ' COMB.Add Item:=cbo_STDSPEC
   '2010,10,25 楼燕南 START
    If txt_INSP_CD.Text = "GB" Then
   
   
      sSQL = "SELECT STDSPEC FROM QP_STD_HEAD WHERE STDSPEC LIKE 'GB712%'"
   
    Else
   
   
      sSQL = "SELECT STDSPEC FROM QP_STD_HEAD WHERE STDSPEC LIKE '" + Trim(txt_INSP_CD.Text) + "%'"
    
    End If
    '2010,10,25 楼燕南 END
    
    Call Gf_ComboAdd(M_CN1, cbo_STDSPEC, sSQL)
    
    If Len(txt_INSP_CD.Text) = 0 Then
        cbo_STDSPEC.Clear
    End If

End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc1"))
      
End Sub
Public Function Gf_Sp_Process1(Conn As adodb.Connection, Sc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean = False) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, icount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim dTempFloat As Double
    
    Dim sMesg As String
    Dim sTemp As String
    Dim ProcessChk As String
    Dim DelYN As Boolean
    Dim Msg_Count As Integer
    Dim Msg_Yes As String
    
    Dim adoCmd As adodb.Command

    Gf_Sp_Process1 = True
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function Or iCount = 0
    If Sc.Item("Spread").MaxRows < 1 Or Sc.Item("iColumn").count = 0 Then
        Gf_Sp_Process1 = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    Sc.Item("Spread").ReDraw = False
    
    'NeceCheck
    For icount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, icount))
            
            Case "Input", "Update"
            
                If Not MC Is Nothing Then
                    Call Gp_Sp_Move(icount, Sc, MC)
                End If
                
                'Maxlength Check
                sMesg = Gf_Sp_NeceCheck2(Sc.Item("Spread"), Sc.Item("mColumn"), icount, Sc.Item("nColumn"))
                        
                If Trim(sMesg) = "OK" Then
                    
                ElseIf Mid(sMesg, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), icount, , vbYellow)
                    sMesg = Mid(sMesg, 6, Len(sMesg))
                    sMesg = sMesg + "长度不正确"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process1 = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), icount, , vbYellow)
                    sMesg = sMesg + "必须输入"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process1 = False
                    Exit Function
                End If
        
        End Select
    
    Next icount
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Sp_Process1 = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For icount = 0 To Sc.Item("iColumn").count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next icount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    Msg_Count = 1
    For icount = 1 To Sc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        DelYN = False
        
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, icount))
        
            Case "Input"
                adoCmd.Parameters(0).Value = "I"
                ProcessChk = "YES"
                
            Case "Update"
                adoCmd.Parameters(0).Value = "U"
                ProcessChk = "YES"
                
            Case "Delete"
                adoCmd.Parameters(0).Value = "D"
                If Msg_Count = 1 Then
                   DelYN = Gf_MessConfirm("您确定要删除状态为[Delete]的数据吗？", "Q")
                   If DelYN Then Msg_Yes = "yes"
                   Msg_Count = Msg_Count + 1
                End If
                If Msg_Yes = "yes" Then DelYN = True
        End Select
          
        If ProcessChk = "YES" Or DelYN Then
            
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").count
            
                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                
                Select Case Sc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempFloat = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Trim(str(dTempFloat))
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempInt = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Trim(str(dTempInt))
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Sc.Item("Spread").Value = "1" Then
                            adoCmd.Parameters(iCol).Value = "1"
                        Else
                            adoCmd.Parameters(iCol).Value = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = "0"
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc.Item("Spread").Value) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).Value = Trim(sTemp)
                        
                End Select
           
            Next iCol
                           
            iProcessCount = iProcessCount + 1
            adoCmd.Execute
            
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), icount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Gf_Sp_Process1 = False
                Exit Function
        
             End If
        
        End If
        
    Next icount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For icount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, icount))
        
            Case "Input", "Update"
                Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, icount)
                
            Case "Delete"
                If DelYN Then
                   Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, icount)
                   Call Gp_Sp_DeleteRow(Sc.Item("Spread"), icount)
                   icount = icount - 1
                End If
        End Select
        
    Next icount
    
    Sc.Item("Spread").ReDraw = True
    
    If iProcessCount > 0 Then
'        If Not MC Is Nothing Then
'            If RefChek = False Then Call Gf_Sp_Display(Conn, Sc.Item("Spread"), _
'                                                    Gf_Ms_MakeQuery(Sc.Item("P-R"), "R", MC("pControl")), Sc.Item("pColumn"), False)
'
'        Else
'            If RefChek = False Then Call Gf_Sp_Display(Conn, Sc.Item("Spread"), _
'                           Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-R"), "R", Sc.Item("aColumn"), 1), Sc.Item("pColumn"), False)
'        End If
        
        MDIMain.StatusBar1.Panels(1) = "提示信息：成功处理了" & iProcessCount & "条记录"
        'Call Gp_MsgBoxDisplay("Data that handle is " & iProcessCount & " items", "I")
        
    End If
            
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            Call Gp_Ms_ControlLock(MC.Item("lControl"), True)
        End If
    Else
        Gf_Sp_Process1 = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Gf_Sp_Process1 = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Process Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

Public Sub Form_Pro()

Dim iCyc, i, count As Integer
Dim SMP_NO, OLD_CONTROL As String

    TXT_SMP_LIST.Text = ""
    
    If Len(Trim(txt_CONTROL_NO.Text)) > 0 Then
        With ss1
            .Row = 1
            .Col = 1
            For iCyc = 1 To .MaxRows
                .Row = iCyc
                .Col = 1
                If .Value = 0 Then
                    .Col = 0
                    If txt_TEST_STS = "B" Then
                       .Text = "Delete"
                    Else
                       .Text = ""
                    End If
                    .Col = SS1_USERID
                    .Text = ""
                Else
                    .Col = 0
                    .Text = "Update"
                    .Col = SS1_CONTROL_NO
                    OLD_CONTROL = .Text
                    .Text = Trim(txt_CONTROL_NO.Text)
                    .Col = SS1_USERID
                    .Text = sUserID
                     .Col = 3
                    TXT_SMP_LIST.Text = TXT_SMP_LIST.Text + .Text + ","
                    If Trim(txt_SMP_NO.Text) <> Trim(.Text) Then
                        txt_SMP_NO.Text = .Text
                        ULabel4.Caption = .Text + "号包含产品"
                        SMP_NO = .Text
                        Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"))
                    End If
                    count = 0
'                    With ss2 --录入控制号时，因技术质量部要求 自动清空表判超过10天 DZB状态的船板
'                         For i = 1 To .MaxRows
'                             .Row = i
'                             .Col = 3
'                             If Mid(.Text, 1, 1) <> "X" Then
'                                MsgBox ("该取样号对应的产品还有未综判的，不能录入控制号！")
'                                COUNT = COUNT + 1
'                             End If
'                         Next i
'                    End With
                    If count > 0 Then
                       .Col = 0
                       .Text = iCyc
                       .Col = 1
                       .Text = 0
                       .Col = SS1_CONTROL_NO
                       .Text = OLD_CONTROL
                       TXT_SMP_LIST.Text = Mid(TXT_SMP_LIST.Text, 1, InStr(TXT_SMP_LIST.Text, SMP_NO))
                    End If
                    
                End If
            Next iCyc
        End With
    Else
        MsgBox ("请输入控制号")
        txt_CONTROL_NO.SetFocus
        Exit Sub
    End If

End Sub

Private Sub txt_PLT_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub


