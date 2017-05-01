VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGG2040C 
   Caption         =   "轧钢计划查询界面_CGG2040C"
   ClientHeight    =   9450
   ClientLeft      =   1110
   ClientTop       =   2025
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
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
      ItemData        =   "CGG2040C.frx":0000
      Left            =   3885
      List            =   "CGG2040C.frx":000D
      TabIndex        =   0
      Top             =   9165
      Visible         =   0   'False
      Width           =   735
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   3000
      Top             =   9165
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "班次"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9075
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   16007
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "轧钢作业计划"
      TabPicture(0)   =   "CGG2040C.frx":001A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSSplitter1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "轧钢大计划"
      TabPicture(1)   =   "CGG2040C.frx":0036
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSSplitter2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "余坯切割计划"
      TabPicture(2)   =   "CGG2040C.frx":0052
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "SSSplitter3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   8775
         Left            =   -75000
         TabIndex        =   2
         Top             =   300
         Width           =   15345
         _ExtentX        =   27067
         _ExtentY        =   15478
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   1
         PaneTree        =   "CGG2040C.frx":006E
         Begin Threed.SSFrame SSFrame2 
            Height          =   555
            Left            =   15
            TabIndex        =   3
            Top             =   15
            Width           =   15315
            _ExtentX        =   27014
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737632
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
               Left            =   2070
               TabIndex        =   10
               Top             =   120
               Width           =   1515
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Left            =   480
               Top             =   120
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   556
               Caption         =   "起始板坯号"
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
            Begin Threed.SSCommand Cmd_Edit 
               Height          =   360
               Left            =   7620
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   120
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   635
               _Version        =   196609
               Font3D          =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "更新数据"
            End
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   8130
            Left            =   15
            TabIndex        =   12
            Top             =   630
            Width           =   15315
            _Version        =   393216
            _ExtentX        =   27014
            _ExtentY        =   14340
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
            MaxCols         =   25
            MaxRows         =   20
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGG2040C.frx":00C0
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   8775
         Left            =   -75000
         TabIndex        =   4
         Top             =   300
         Width           =   15345
         _ExtentX        =   27067
         _ExtentY        =   15478
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   1
         PaneTree        =   "CGG2040C.frx":0BE5
         Begin Threed.SSFrame SSFrame3 
            Height          =   555
            Left            =   15
            TabIndex        =   5
            Top             =   15
            Width           =   15315
            _ExtentX        =   27014
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737632
            Begin InDate.UDate udFmDate 
               Height          =   315
               Left            =   2070
               TabIndex        =   6
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
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
            Begin InDate.ULabel labDate 
               Height          =   315
               Left            =   480
               Tag             =   "查询日FROM"
               Top             =   120
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   556
               Caption         =   "查询日期"
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
            Begin InDate.UDate udToDate 
               Height          =   315
               Left            =   3840
               TabIndex        =   7
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
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
            Begin Threed.SSCommand SCmd2 
               Height          =   360
               Left            =   8500
               TabIndex        =   8
               Top             =   90
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   635
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "上传计划"
            End
            Begin InDate.ULabel ULabel16 
               Height          =   315
               Left            =   11970
               Top             =   120
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               Caption         =   "计划工作日"
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
            Begin InDate.UDate TXT_DATE 
               Height          =   315
               Left            =   13320
               TabIndex        =   25
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
            Begin VB.Label Label1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "~"
               Height          =   120
               Left            =   3630
               TabIndex        =   9
               Top             =   240
               Width           =   195
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   8130
            Left            =   15
            TabIndex        =   13
            Top             =   630
            Width           =   15315
            _Version        =   393216
            _ExtentX        =   27014
            _ExtentY        =   14340
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
            MaxCols         =   22
            MaxRows         =   20
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGG2040C.frx":0C37
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter3 
         Height          =   8775
         Left            =   0
         TabIndex        =   14
         Top             =   300
         Width           =   15345
         _ExtentX        =   27067
         _ExtentY        =   15478
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   1
         PaneTree        =   "CGG2040C.frx":1956
         Begin Threed.SSFrame SSFrame1 
            Height          =   555
            Left            =   15
            TabIndex        =   15
            Top             =   15
            Width           =   15315
            _ExtentX        =   27014
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737632
            Begin VB.ComboBox CBO_CUTYN 
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
               ItemData        =   "CGG2040C.frx":19A8
               Left            =   7305
               List            =   "CGG2040C.frx":19B5
               TabIndex        =   24
               Tag             =   "轧辊号"
               Top             =   120
               Width           =   780
            End
            Begin InDate.UDate udFmDate2 
               Height          =   315
               Left            =   2070
               TabIndex        =   16
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
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
               Left            =   480
               Tag             =   "查询日FROM"
               Top             =   120
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   556
               Caption         =   "查询日期"
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
            Begin InDate.UDate udToDate2 
               Height          =   315
               Left            =   3840
               TabIndex        =   17
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
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
            Begin Threed.SSCommand SCmd3 
               Height          =   360
               Left            =   8500
               TabIndex        =   18
               Top             =   90
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   635
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "上传计划"
            End
            Begin Threed.SSPanel SSP1 
               Height          =   315
               Left            =   10740
               TabIndex        =   21
               Top             =   120
               Width           =   1440
               _ExtentX        =   2540
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
               Caption         =   "待切割"
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel SSP3 
               Height          =   315
               Left            =   12180
               TabIndex        =   22
               Top             =   120
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   196609
               ForeColor       =   16711680
               BackColor       =   12648384
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "切割计划"
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin Threed.SSPanel SSP4 
               Height          =   315
               Left            =   13620
               TabIndex        =   23
               Top             =   120
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   556
               _Version        =   196609
               ForeColor       =   16711680
               BackColor       =   65535
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "可修改项"
               RoundedCorners  =   0   'False
               FloodShowPct    =   -1  'True
            End
            Begin InDate.ULabel ULabel3 
               Height          =   315
               Left            =   5700
               Top             =   120
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   556
               Caption         =   "是否切割完成"
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
            Begin VB.Label Label2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "~"
               Height          =   120
               Left            =   3630
               TabIndex        =   19
               Top             =   240
               Width           =   195
            End
         End
         Begin FPSpread.vaSpread ss3 
            Height          =   8130
            Left            =   15
            TabIndex        =   20
            Top             =   630
            Width           =   15315
            _Version        =   393216
            _ExtentX        =   27014
            _ExtentY        =   14340
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
            MaxCols         =   20
            MaxRows         =   20
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGG2040C.frx":19C5
         End
      End
   End
End
Attribute VB_Name = "CGG2040C"
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
'-- Program Name      精整作业计划查询界面
'-- Program ID        CKG2040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2007.12.19
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE         EDITOR       DESCRIPTION
'-- 1.01  2007.12.19   GUOLI
'-- 1.02  2012.03.02   LiQian       余坯切割计划需记录切割人员和时间,增加字段记录相应信息
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

Dim pControl2 As New Collection     'Master Primary Key Collection
Dim nControl2 As New Collection     'Master Necessary Collection
Dim mControl2 As New Collection     'Master Maxlength check Collection
Dim iControl2 As New Collection     'Master Insert Collection
Dim rControl2 As New Collection     'Master Refer Collection
Dim cControl2 As New Collection     'Master Copy Collection
Dim aControl2 As New Collection     'Master -> Spread Collection
Dim lControl2 As New Collection     'Master Lock Collection

Dim pControl3 As New Collection     'Master Primary Key Collection
Dim nControl3 As New Collection     'Master Necessary Collection
Dim mControl3 As New Collection     'Master Maxlength check Collection
Dim iControl3 As New Collection     'Master Insert Collection
Dim rControl3 As New Collection     'Master Refer Collection
Dim cControl3 As New Collection     'Master Copy Collection
Dim aControl3 As New Collection     'Master -> Spread Collection
Dim lControl3 As New Collection     'Master Lock Collection

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

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim SE, MPLATE_NO, plate_no As String

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim lRow        As Long
Dim lRowRange   As Long

Const SS2_EMP_NO = 21
Const SS2_INS_DATE = 22

Const SS3_RST_CUT = 10
Const SS3_CUT_SEQ = 15
Const SS3_REMARK = 16  '导出ExcelPrn2中用到
Const SS3_UPD_EMP = 19 ' Add by liqian at 2012-03-02

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                   
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
   Call Gp_Ms_Collection(udFmDate, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
   Call Gp_Ms_Collection(udToDate, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
   Call Gp_Ms_Collection(udFmDate2, "p", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(udToDate2, "p", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
   Call Gp_Ms_Collection(CBO_CUTYN, "p", " ", " ", " ", "r", " ", " ", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
      
    Mc3.Add Item:=pControl3, Key:="pControl"
    Mc3.Add Item:=nControl3, Key:="nControl"
    Mc3.Add Item:=mControl3, Key:="mControl"
    Mc3.Add Item:=iControl3, Key:="iControl"
    Mc3.Add Item:=rControl3, Key:="rControl"
    Mc3.Add Item:=cControl3, Key:="cControl"
    Mc3.Add Item:=aControl3, Key:="aControl"
    Mc3.Add Item:=lControl3, Key:="lControl"
    
    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    'Spread_Collection
    
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGG2040C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)


    'Spread_Collection
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGG2040C.P_SREFER2", Key:="P-R"
    sc2.Add Item:="CGG2040C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, "p", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 16, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 17, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 18, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    '2012-03-02  add by liqian 切割人员/切割时间添加,系统记录切割人员和时间
    Call Gp_Sp_Collection(ss3, 19, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 20, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)


    'Spread_Collection
    
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="CGG2040C.P_SREFER3", Key:="P-R"
    sc3.Add Item:="CGG2040C.P_MODIFY3", Key:="P-M"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc3, Key:="Sc3"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    MDIMain.MenuTool.Buttons(7).Enabled = False
'    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

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
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_Cls(Mc3("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    Call Gp_Ms_NeceColor(Mc3("nControl"))

    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    Call Gp_Sp_Setting(sc3.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc3.Item("Spread"), "G-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc3.Item("Spread"), "G-System.INI", Me.Name)

    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set pControl3 = Nothing
    Set nControl3 = Nothing
    Set iControl3 = Nothing
    Set rControl3 = Nothing
    Set cControl3 = Nothing
    Set aControl3 = Nothing
    Set lControl3 = Nothing
    Set mControl3 = Nothing
    
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
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If SSTab1.Tab = 0 Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        End If
    ElseIf SSTab1.Tab = 1 Then
        Call Gf_Sp_Cls(sc2)
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc2("lControl"), False)
    Else
        Call Gf_Sp_Cls(sc3)
        Call Gp_Ms_Cls(Mc3("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc3("lControl"), False)
    End If

End Sub

Public Sub Form_Exc()

    Select Case SSTab1.Tab
           
           Case 0
     
                Call ExcelPrn
            
           Case 1
           
                Call ExcelPrn1
                
           Case 2
           
                Call ExcelPrn2
    
    End Select
    
End Sub

Public Sub Form_Ref()

    Dim iRow As Long
    Dim iCUT_SEQ As Double

On Error GoTo Refer_Err

    If SSTab1.Tab = 0 Then
        If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
        If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
'            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        End If
    ElseIf SSTab1.Tab = 1 Then
        
        If Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl")) Then
            ss2.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
'            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        End If
    Else
    
        If Gf_Sp_Refer(M_CN1, sc3, Mc3) Then
            ss3.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
'            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        End If
        
        Call SS3_Format
        
    End If
    
    Exit Sub

Refer_Err:

End Sub


Public Sub Form_Pro()

    Dim iCount      As Integer

    Select Case SSTab1.Tab
           
           Case 0
     
                If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
            
           Case 1
           
                For iCount = 1 To ss2.MaxRows
                    ss2.Row = iCount
                    ss2.Col = 0
                    If ss2.Text = "Input" Then
                        ss2.Col = SS2_INS_DATE
                        ss2.Text = TXT_DATE.Text
                    End If
                Next
           
                If Gf_Sp_Process(M_CN1, sc2, Mc2) Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
                
           Case 2
           
                If Gf_Sp_Process(M_CN1, sc3, Mc3) Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
                
                Call SS3_Format
    
    End Select
    
End Sub

    
Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1

End Sub
Public Sub Spread_Del()

    Select Case SSTab1.Tab
            
           Case 1
           
                Call Gp_Sp_Del(sc2)
                
           Case 2
           
                Call Gp_Sp_Del(sc3)
                
    End Select

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

Private Sub SCmd2_Click()
   Load LoadExcel
   LoadExcel.TXT_SS.Text = "ss2"
   LoadExcel.Show 1
End Sub

Private Sub SCmd3_Click()
   Load LoadExcel
   LoadExcel.TXT_SS.Text = "ss3"
   LoadExcel.Show 1
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

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

Private Sub ss1_SelChange(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long, ByVal CurCol As Long, ByVal curRow As Long)
  
    lRowRange = BlockRow2 - BlockRow + 1
    lRow = BlockRow
    
End Sub

Private Sub ExcelPrn()
    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDateText       As String
    Dim sDate           As String
    Dim sDateNext       As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\CGG2040C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
       
    sDate = Format(Now, "YYYY-MM-DD HH:MM")
    sDateText = Left(sDate, 4) & "年"
    sDateText = sDateText & Mid(sDate, 6, 2) & "月"
    sDateText = sDateText & Mid(sDate, 9, 2) & "日 "
    sDateText = sDateText & Mid(sDate, 12, 2) & "点"
    sDateText = sDateText & Mid(sDate, 15, 2) & "分"
    
    xlApp.Range("A2").Value = sDateText

    If lBlkrow1 = 0 Then
        lBlkrow1 = 1
    End If
    If lBlkrow2 = 0 Then
        lBlkrow2 = ss1.MaxRows
    End If
    
    For i = 2 To lBlkrow2 - lBlkrow1 + 1
          xlApp.Rows("4:4").Select
          xlApp.Selection.Copy
          xlApp.Selection.Insert Shift:=1
    Next i
    
    Clipboard.Clear
    
    ss1.Col = 1:        ss1.Col2 = 9
    ss1.Row = lBlkrow1: ss1.Row2 = lBlkrow2
    
    Clipboard.SetText ss1.Clip
        
    For i = 1 To lBlkrow2 - lBlkrow1 + 1
          xlApp.Range("A" & i + 3).Value = i
    Next i
    
    xlApp.Range("B4").Select
    xlApp.ActiveSheet.Paste

    
    Clipboard.Clear
    
    ss1.Col = 12:       ss1.Col2 = ss1.MaxCols - 1
    ss1.Row = lBlkrow1: ss1.Row2 = lBlkrow2
    
    Clipboard.SetText ss1.Clip
        
    For i = 1 To lBlkrow2 - lBlkrow1 + 1
          xlApp.Range("A" & i + 3).Value = i
    Next i

    xlApp.Range("K4").Select
    xlApp.ActiveSheet.Paste
    
'    Clipboard.Clear
'    ss1.SetSelection 1, 1, ss1.MaxCols, ss1.MaxRows
'    ss1.ClipboardCopy
    
'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
'     xlApp.Application.Visible = False
'     xlSheet.Close False
'     xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub ExcelPrn1()

    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDateText       As String
    Dim sDate           As String
    Dim sDateNext       As String
    
    If ss2.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\CGG2041C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    Clipboard.Clear
    ss2.SetSelection 1, 1, 1, ss2.MaxRows
    ss2.ClipboardCopy
    xlApp.Range("A4").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss2.SetSelection 3, 1, ss2.MaxCols - 2, ss2.MaxRows
    ss2.ClipboardCopy
    xlApp.Range("B4").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    ss2.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub ExcelPrn2()

    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDateText       As String
    Dim sDate           As String
    Dim sDateNext       As String
    
    If ss3.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\CGG2042C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    Clipboard.Clear
    ss3.SetSelection 1, 1, 1, ss3.MaxRows
    ss3.ClipboardCopy
    xlApp.Range("A4").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    'modify by liqian at 2012-03-05
    ss3.SetSelection 3, 1, SS3_REMARK, ss3.MaxRows
    ss3.ClipboardCopy
    xlApp.Range("B4").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    ss3.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub Cmd_Edit_Click()
    'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    'Dim strEdit_date As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    'strEdit_date = CStr(TXT_CHK_TIME.RawData)
    
    sQuery = "{call CGG2040P ('" + "C3" + "',?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
            
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        strRet_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & strRet_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        
        Call Gp_MsgBoxDisplay("更新成功..!!", "I")
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("更新失败！！")

End Sub

Private Sub Cmd_Seq_Update_Click()

    Dim AdoRs As ADODB.Recordset
    Dim i             As Integer
    Dim sSlabNo       As String
    Dim SqlCmd        As String
    Dim ProcessCnt    As Integer

    On Error GoTo ErrHandle
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    M_CN1.BeginTrans

    For i = 1 To ss1.MaxRows
        ss1.Row = i
        ss1.Col = 1
        sSlabNo = Trim(ss1.Text)
                                  
        SqlCmd = ""
        SqlCmd = " UPDATE  GP_MILL_PLAN   " & vbCrLf
        SqlCmd = SqlCmd & "   SET  CHG_SEQ_NO  = " & i & vbCrLf
        SqlCmd = SqlCmd & "  Where SLAB_NO     = '" & sSlabNo & "' " & vbCrLf
        
        M_CN1.Execute SqlCmd
    Next i
    
    M_CN1.CommitTrans
    Call Gp_MsgBoxDisplay("更新成功..!!", "I")
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandle:

    M_CN1.RollbackTrans
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("更新失败！！")

End Sub
Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(ss2, Mode)
    End If

End Sub

Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    Dim iRow As Long
    iRow = Row
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(ss3, Mode)
        ss3.Row = iRow
        ss3.Col = SS3_UPD_EMP
        ss3.Text = sUserID
    End If

End Sub

Public Sub SS3_Format()

    Dim iRow  As Long
    Dim iCol  As Long
    Dim iCUT_SEQ As Double
    
    If ss3.MaxRows > 0 Then
    
        For iRow = 1 To ss3.MaxRows
        
            ss3.Row = iRow
            
            For iCol = 6 To 15
            
                ss3.Col = iCol
                
                If iCol = SS3_CUT_SEQ Then
                    
                    iCUT_SEQ = ss3.Value
                    
                    If iCUT_SEQ = 0 Then
                        Call Gp_Sp_BlockColor(ss3, 3, ss3.MaxCols, ss3.Row, ss3.Row, , SSP1.BackColor)
                    Else
                        Call Gp_Sp_BlockColor(ss3, 3, ss3.MaxCols, ss3.Row, ss3.Row, , SSP3.BackColor)
                    End If
                    
                End If
                
                If ss3.CellType = SS_CELL_TYPE_NUMBER Then
                    If Val(ss3.Text & "") = 0 Then
                        ss3.Text = ""
                    End If
                End If
                
            Next iCol
        
        Next iRow
        
    End If
    
    Call Gp_Sp_BlockColor(ss3, SS3_RST_CUT, SS3_RST_CUT, 1, ss3.MaxRows, , SSP4.BackColor)
    Call Gp_Sp_BlockColor(ss3, SS3_REMARK, SS3_REMARK, 1, ss3.MaxRows, , SSP4.BackColor)
  
End Sub

