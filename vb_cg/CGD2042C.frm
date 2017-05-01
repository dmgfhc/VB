VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGD2042C 
   Caption         =   "火切钢板取样信息查询及修改界面_CGD2042C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9060
      Left            =   30
      TabIndex        =   0
      Top             =   -15
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   15981
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "CGD2042C.frx":0000
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   1455
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   15090
         Begin VB.TextBox txt_charge_no 
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
            Left            =   1425
            MaxLength       =   14
            TabIndex        =   13
            Top             =   600
            Width           =   1710
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
            ItemData        =   "CGD2042C.frx":0072
            Left            =   4320
            List            =   "CGD2042C.frx":0082
            TabIndex        =   12
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox TXT_PROD_CD 
            Height          =   270
            Left            =   2880
            TabIndex        =   11
            Top             =   210
            Visible         =   0   'False
            Width           =   255
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
            ItemData        =   "CGD2042C.frx":0092
            Left            =   4320
            List            =   "CGD2042C.frx":009F
            TabIndex        =   10
            Top             =   210
            Width           =   855
         End
         Begin VB.TextBox txt_SMP_LOC_NAME 
            Height          =   315
            Left            =   11835
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   195
            Width           =   585
         End
         Begin VB.TextBox txt_SMP_LOC 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
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
            Left            =   11535
            MaxLength       =   1
            TabIndex        =   8
            Tag             =   "取样位置"
            Top             =   195
            Width           =   270
         End
         Begin VB.TextBox txt_SMP_LEN 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
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
            Left            =   9825
            MaxLength       =   8
            TabIndex        =   7
            Tag             =   "试样长度"
            Top             =   195
            Width           =   630
         End
         Begin VB.TextBox txt_SMP_NO 
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
            Left            =   6675
            MaxLength       =   14
            TabIndex        =   6
            Tag             =   "试样号"
            Top             =   195
            Width           =   1605
         End
         Begin VB.TextBox txt_CHG_SMP_NO 
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
            Left            =   6675
            MaxLength       =   14
            TabIndex        =   5
            Tag             =   "改判时试样号"
            Top             =   585
            Width           =   1605
         End
         Begin VB.TextBox txt_CHG_STDSPEC 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
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
            Left            =   9825
            TabIndex        =   4
            Tag             =   "改判时适用标准"
            Top             =   585
            Width           =   2610
         End
         Begin VB.TextBox txt_SLAB_NO 
            Height          =   270
            Left            =   9480
            MaxLength       =   10
            TabIndex        =   3
            Top             =   1080
            Visible         =   0   'False
            Width           =   2040
         End
         Begin InDate.ULabel ULabel9 
            Height          =   315
            Left            =   240
            Top             =   600
            Width           =   1155
            _ExtentX        =   2037
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   240
            Top             =   210
            Width           =   1155
            _ExtentX        =   2037
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
            Left            =   3450
            Top             =   210
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
               Size            =   9.76
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin InDate.UDate SDT_PROD_DATE 
            Height          =   315
            Left            =   1425
            TabIndex        =   14
            Top             =   210
            Width           =   1440
            _ExtentX        =   2540
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
            MaxLength       =   10
         End
         Begin InDate.ULabel ULabel8 
            Height          =   315
            Left            =   3450
            Top             =   600
            Width           =   855
            _ExtentX        =   1508
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
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   5460
            Top             =   195
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "试样号"
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   8595
            Top             =   195
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "试样长度"
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   10530
            Top             =   195
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   556
            Caption         =   "取样位置"
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
         Begin InDate.ULabel ULabel6 
            Height          =   315
            Left            =   5460
            Top             =   585
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "改判试样号"
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
            Height          =   315
            Left            =   8595
            Top             =   585
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            Caption         =   "改判时标准"
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
         Begin Threed.SSCommand Cmd_Set_Save 
            Height          =   330
            Left            =   12720
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   180
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
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
            Caption         =   "多行设定"
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   420
            Left            =   240
            TabIndex        =   18
            Top             =   990
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   741
            _Version        =   196609
            BackColor       =   14737632
            Begin VB.TextBox TXT_PRCLINE 
               Height          =   285
               Left            =   2550
               TabIndex        =   19
               Text            =   " "
               Top             =   30
               Visible         =   0   'False
               Width           =   225
            End
            Begin Threed.SSOption opt_LineFlag 
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   20
               Top             =   90
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   2
               ForeColor       =   0
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
               Caption         =   "# 3"
            End
            Begin Threed.SSOption opt_LineFlag 
               Height          =   255
               Index           =   3
               Left            =   1740
               TabIndex        =   21
               Top             =   90
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               _Version        =   196609
               Font3D          =   2
               ForeColor       =   0
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
               Caption         =   "# 4"
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   420
            Left            =   3450
            TabIndex        =   22
            Top             =   990
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   741
            _Version        =   196609
            BackColor       =   14737632
            Begin Threed.SSOption opt_Product 
               Height          =   330
               Index           =   1
               Left            =   330
               TabIndex        =   23
               Top             =   60
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   582
               _Version        =   196609
               Font3D          =   2
               ForeColor       =   8421504
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "板坯号"
            End
            Begin Threed.SSOption opt_Product 
               Height          =   330
               Index           =   2
               Left            =   1680
               TabIndex        =   24
               Top             =   60
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   582
               _Version        =   196609
               Font3D          =   2
               ForeColor       =   8421504
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "轧批号"
            End
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "出口板,锅炉板, 压力容器板必须输入"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   390
            Left            =   12735
            TabIndex        =   16
            Top             =   570
            Width           =   1815
         End
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3600
         Left            =   0
         TabIndex        =   1
         Top             =   5460
         Width           =   15090
         _Version        =   393216
         _ExtentX        =   26617
         _ExtentY        =   6350
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   31
         MaxRows         =   2
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGD2042C.frx":00AC
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3825
         Left            =   0
         TabIndex        =   17
         Top             =   1545
         Width           =   15090
         _Version        =   393216
         _ExtentX        =   26617
         _ExtentY        =   6747
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   16
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGD2042C.frx":1024
      End
   End
End
Attribute VB_Name = "CGD2042C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name   OLD MILL PLATE
'-- Program Name      SAMPLE CUT
'-- Program ID        CGD2041C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2007.8.30
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

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

Dim pContro12 As New Collection      'Master Primary Key Collection
Dim nContro12 As New Collection      'Master Necessary Collection
Dim mContro12 As New Collection      'Master Maxlength check Collection
Dim iContro12 As New Collection      'Master Insert Collection
Dim rContro12 As New Collection      'Master Refer Collection
Dim cContro12 As New Collection      'Master Copy Collection
Dim aContro12 As New Collection      'Master -> Spread Collection
Dim lContro12 As New Collection      'Master Lock Collection


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
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sWgtLenFlag As String
Dim sQuery      As String
Dim bCheck      As Boolean
Dim sLoopChk    As String

Const SS1_SLAB_NO = 1
Const SS1_ORD_NO = 10
Const SS1_ORD_ITEM = 11
Const SS1_URGNT_FL = 16

Const SS2_PLATE_NO = 2                  'PLATE NO
Const SS2_PROC_CD = 6                   '进程代码
Const SS2_PROD_CD = 3                   'PRODUCT STATUS
Const SS2_SMP_FLAG = 16                 '实绩标记
Const SS2_SMP_LOC = 17                  '位置
Const SS2_SMP_LEN = 18                  '长度
Const SS2_SMP_NO = 19                   '试样号
Const SS2_STDSPEC = 20                  '标准号
Const SS2_BEF_STDSPEC = 21              'BEFORE 标准号
Const SS2_USER_ID = 24                  'USER ID
Const SS2_BEF_SMP_FLAG = 25             'BEFORE 实绩标记
Const SS2_BEF_SMP_LOC = 26              'BEFORE 位置
Const SS2_BEF_SMP_LEN = 27              'BEFORE 长度
Const SS2_BEF_SMP_NO = 28               'BEFORE 试样号
Const SS2_CHG_SMP_NO = 29               '改判时试样号
Const SS2_CHG_STDSPEC = 30              '改判时适用标准


Private Sub Form_Define()
      
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_charge_no, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(SDT_PROD_DATE, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(TXT_PROD_CD, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(txt_PrcLine, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
        'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_SLAB_NO, "p", "n", "m", " ", " ", " ", " ", pContro12, nContro12, mContro12, iContro12, rContro12, aContro12, lContro12)
    
    'MASTER Collection
     Mc2.Add Item:=pContro12, Key:="pControl"
     Mc2.Add Item:=nContro12, Key:="nControl"
     Mc2.Add Item:=mContro12, Key:="mControl"
     Mc2.Add Item:=iContro12, Key:="iControl"
     Mc2.Add Item:=rContro12, Key:="rControl"
     Mc2.Add Item:=cContro12, Key:="cControl"
     Mc2.Add Item:=aContro12, Key:="aControl"
     Mc2.Add Item:=lContro12, Key:="lControl"


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

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    'sc1.Add Item:="CGD2041C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="CGD2042C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
      
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "◎"


    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, "p", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGD2042C.P_MODIFY", Key:="P-M"
    sc2.Add Item:="CGD2042C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").ROW = 0
    sc2.Item("Spread").Text = "◎"
    
    Call Gp_Sp_ColHidden(ss2, SS2_USER_ID, True)
    Call Gp_Sp_ColHidden(ss2, SS2_BEF_SMP_FLAG, True)
    Call Gp_Sp_ColHidden(ss2, SS2_BEF_SMP_LOC, True)
    Call Gp_Sp_ColHidden(ss2, SS2_BEF_SMP_LEN, True)
    Call Gp_Sp_ColHidden(ss2, SS2_BEF_SMP_NO, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

    
End Sub

Private Sub Cmd_Set_Save_Click()
    Dim intRow As Integer
    Dim intCount As Integer
    Dim strQuery As String
    Dim strQuery_H As String
    Dim strQuery_L As String
    
    ss2.MaxRows = 0
    If ss1.MaxRows < 1 Then Exit Sub
    
    strQuery = "  SELECT CHK                                                                    " & vbCrLf
    strQuery = strQuery & "         ,MATA_NO                                                    " & vbCrLf
    strQuery = strQuery & "         ,PROD_CD                                                    " & vbCrLf
    strQuery = strQuery & "         ,LOT_NO                                                     " & vbCrLf
    strQuery = strQuery & "         ,MOCUT_NO                                                   " & vbCrLf
    strQuery = strQuery & "         ,PROC_CD                                                    " & vbCrLf
    strQuery = strQuery & "         ,THK                                                        " & vbCrLf
    strQuery = strQuery & "         ,WID                                                        " & vbCrLf
    strQuery = strQuery & "         ,LEN                                                        " & vbCrLf
    strQuery = strQuery & "         ,WGT                                                        " & vbCrLf
    strQuery = strQuery & "         ,ORD_FL                                                     " & vbCrLf
    strQuery = strQuery & "         ,ORD_NO                                                     " & vbCrLf
    strQuery = strQuery & "         ,ORD_ITEM                                                   " & vbCrLf
    strQuery = strQuery & "         ,SMP_FL                                                     " & vbCrLf
    strQuery = strQuery & "         ,SMP_LEN                                                    " & vbCrLf
    strQuery = strQuery & "         ,ACT_SMP_FL                                                 " & vbCrLf
    strQuery = strQuery & "         ,SMP_LOC                                                    " & vbCrLf
    strQuery = strQuery & "         ,ACT_SMP_LEN                                                " & vbCrLf
    strQuery = strQuery & "         ,SMP_NO                                                     " & vbCrLf
    strQuery = strQuery & "         ,STD_SPEC                                                   " & vbCrLf
    strQuery = strQuery & "         ,BEF_STDSPEC                                                " & vbCrLf
    strQuery = strQuery & "         ,CUST_NAME                                                  " & vbCrLf
    strQuery = strQuery & "         ,PROD_DATE                                                  " & vbCrLf
    strQuery = strQuery & "         ,UPD_EMP_CD                                                 " & vbCrLf
    strQuery = strQuery & "         ,BEF_ACT_SMP_FL                                             " & vbCrLf
    strQuery = strQuery & "         ,BEF_SMP_LOC                                                " & vbCrLf
    strQuery = strQuery & "         ,BEF_ACT_SMP_LEN                                            " & vbCrLf
    strQuery = strQuery & "         ,BEF_SMP_NO                                                 " & vbCrLf
    strQuery = strQuery & "         ,''                                                         " & vbCrLf
    strQuery = strQuery & "         ,''                                                         " & vbCrLf
    strQuery = strQuery & "         ,SPECIAL_OPR_REQ                                            " & vbCrLf
    strQuery = strQuery & " FROM (                                                              " & vbCrLf
    
    strQuery = strQuery & "  SELECT 0                                          CHK              " & vbCrLf
    strQuery = strQuery & "         ,A.PLATE_NO                                MATA_NO          " & vbCrLf
    strQuery = strQuery & "         ,A.PROD_CD                                 PROD_CD          " & vbCrLf
    
    strQuery = strQuery & "         ,A.OUT_SHEET_NO LOT_NO                                       " & vbCrLf
    strQuery = strQuery & "         ,(SELECT MAX(TRNS_CMPY_CD) FROM NISCO.GP_PLATE WHERE PLATE_NO = SUBSTR(A.PLATE_NO,1,12))    MOCUT_NO        " & vbCrLf
    
    strQuery = strQuery & "         ,A.PROC_CD                                 PROC_CD          " & vbCrLf
    strQuery = strQuery & "         ,A.THK                                     THK              " & vbCrLf
    strQuery = strQuery & "         ,A.WID                                     WID              " & vbCrLf
    strQuery = strQuery & "         ,A.LEN                                     LEN              " & vbCrLf
    strQuery = strQuery & "         ,A.WGT                                     WGT              " & vbCrLf
    strQuery = strQuery & "         ,A.ORD_FL                                  ORD_FL           " & vbCrLf
    strQuery = strQuery & "         ,A.ORD_NO                                  ORD_NO           " & vbCrLf
    strQuery = strQuery & "         ,A.ORD_ITEM                                ORD_ITEM         " & vbCrLf
    strQuery = strQuery & "         ,A.SMP_FL                                  SMP_FL           " & vbCrLf
    strQuery = strQuery & "         ,DECODE(A.SMP_FL,'N',0,A.SMP_LEN)          SMP_LEN          " & vbCrLf
    strQuery = strQuery & "         ,A.ACT_SMP_FL                              ACT_SMP_FL       " & vbCrLf
    strQuery = strQuery & "         ,A.SMP_LOC                                 SMP_LOC          " & vbCrLf
    strQuery = strQuery & "         ,A.ACT_SMP_LEN                             ACT_SMP_LEN      " & vbCrLf
    strQuery = strQuery & "         ,A.SMP_NO                                  SMP_NO           " & vbCrLf
    strQuery = strQuery & "         ,A.APLY_STDSPEC                            STD_SPEC         " & vbCrLf
    strQuery = strQuery & "         ,A.BEF_APLY_STDSPEC                        BEF_STDSPEC      " & vbCrLf
    strQuery = strQuery & "         ,GF_CUSTNAMEFIND(A.CUST_CD)                CUST_NAME        " & vbCrLf
    strQuery = strQuery & "         ,A.PROD_DATE                               PROD_DATE        " & vbCrLf
    strQuery = strQuery & "         ,A.UPD_EMP_CD                              UPD_EMP_CD       " & vbCrLf
    strQuery = strQuery & "         ,A.ACT_SMP_FL                              BEF_ACT_SMP_FL   " & vbCrLf
    strQuery = strQuery & "         ,A.SMP_LOC                                 BEF_SMP_LOC      " & vbCrLf
    strQuery = strQuery & "         ,A.ACT_SMP_LEN                             BEF_ACT_SMP_LEN  " & vbCrLf
    strQuery = strQuery & "         ,A.SMP_NO                                  BEF_SMP_NO       " & vbCrLf
    strQuery = strQuery & "         ,B.SPECIAL_OPR_REQ                         SPECIAL_OPR_REQ  " & vbCrLf
    strQuery = strQuery & "   FROM  GP_PLATE A,BP_ORDER_ITEM  B                                 " & vbCrLf
    strQuery = strQuery & "   WHERE (A.REC_STS          =     '2')     AND                      " & vbCrLf
    strQuery = strQuery & "         (A.PRC_LINE         IN ( '3' ,'4'))     AND                      " & vbCrLf
    strQuery = strQuery & "         (A.PROD_CD          =     'PP')    AND                      " & vbCrLf
    strQuery = strQuery & "         (A.PLT              =     'C3')    AND                      " & vbCrLf
    strQuery = strQuery & "         (NVL(A.HTM_METH1,'H')   = 'H' )    AND                      " & vbCrLf
    strQuery = strQuery & "         A.ORD_NO                =   B.ORD_NO    AND                 " & vbCrLf
    strQuery = strQuery & "         A.ORD_ITEM              =   B.ORD_ITEM  AND                 " & vbCrLf
    
    strQuery_H = strQuery
    
    intCount = 0
    
    With ss1
        For intRow = 1 To .MaxRows
            .Col = 0: .ROW = intRow
            If .Text = "Selected" Then
                .Col = 1
                If intCount = 0 Then

                    If TXT_PROD_CD = "SL" Then
                      .Col = 1
                    Else
                      .Col = 2
                    End If
                    txt_charge_no = .Text
                    .Col = 1
                    strQuery_L = strQuery_H & " ( A.PLATE_NO  LIKE  '" & Trim(.Text) & "'|| '%'  )) " & vbCrLf
                Else
                    strQuery_L = strQuery_L & "  UNION ALL " & strQuery_H & vbCrLf
                    strQuery_L = strQuery_L & " ( A.PLATE_NO  LIKE  '" & Trim(.Text) & "'|| '%'  )) " & vbCrLf
                End If
                intCount = intCount + 1
                
            End If
        Next intRow
        
    End With
    
    strQuery_L = strQuery_L & "    ORDER BY MATA_NO                                              " & vbCrLf
    
    sLoopChk = "**"
    If intCount > 0 Then
        If Gf_Sp_Display(M_CN1, ss2, strQuery_L) Then
            Call Sample_No_Edit
            sLoopChk = ""
            Call ss2_set_check
        End If
    End If
    sLoopChk = ""
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuToolSet

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Len(txt_charge_no.Text) >= 8 Then
           Call Form_Ref
        End If
'        KeyAscii = 0
'        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuToolSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    
    With ss2
        .ROW = .ColHeaderRows + 1: .Row2 = .ColHeaderRows + 1
        .Col = SS2_SMP_LOC: .Col2 = SS2_SMP_LOC
        
        .BlockMode = True
        
        .CellType = SS_CELL_TYPE_STATIC_TEXT
        .TypeHAlign = SS_CELL_H_ALIGN_CENTER
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .TypeTextWordWrap = True
        
        .BackColor = &HE1E4CD
        .ForeColor = BLUE
        
        .BlockMode = False

    End With
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)
        
    SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
    
    opt_Product(1).Value = True
    opt_LineFlag(2).Value = True
    ULabel9.Caption = opt_Product(1).Caption
    
    bCheck = False
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc2")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
    
    Set pContro12 = Nothing
    Set nContro12 = Nothing
    Set iContro12 = Nothing
    Set rContro12 = Nothing
    Set cContro12 = Nothing
    Set aContro12 = Nothing
    Set lContro12 = Nothing
    Set mContro12 = Nothing
      
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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC2"))
      
End Sub

Public Sub Form_Cls()
    
    Dim sProd_cd As String
    
    sProd_cd = TXT_PROD_CD
    
    If Gf_Sp_Cls(sc2) Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("pControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call MenuToolSet
            Call TextClear
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            Call pContro1(1).SetFocus
            TXT_PROD_CD = sProd_cd
        End If
    End If
    bCheck = False
End Sub

Public Sub MenuToolSet()

    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(14).Enabled = False

End Sub

Public Sub TextClear()

    txt_smp_no.Text = ""
    txt_SMP_LEN.Text = ""
    txt_SMP_LOC.Text = ""
    txt_SMP_LOC_NAME.Text = ""
    txt_CHG_SMP_NO.Text = ""
    txt_CHG_STDSPEC.Text = ""

End Sub

Public Sub Form_Ref()
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sUrgnt_Fl As String
    
    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    If Trim(txt_charge_no.Text) = "" And SDT_PROD_DATE.RawData = "" Then
        Call Gp_MsgBoxDisplay("请输入查询号还是生产开始日期！！！")
        Exit Sub
    End If
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuToolSet
        Call TextClear
        bCheck = False
    End If
    
    '紧急订单绿色显示 add by liqian 2012-10-08
     With ss1
          For iRow = 1 To .MaxRows
             .ROW = iRow:
              .Col = SS1_URGNT_FL:    sUrgnt_Fl = Trim(.Text)

              If sUrgnt_Fl = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SS1_SLAB_NO, SS1_SLAB_NO, .ROW, .ROW, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_NO, SS1_ORD_NO, .ROW, .ROW, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_ITEM, SS1_ORD_ITEM, .ROW, .ROW, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, .ROW, .ROW, &HC000&)
              End If
          Next iRow
    End With
    
    
    ss2.MaxRows = 0
    
     
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub


Public Sub Form_Pro()
    Dim intRow      As Integer
    Dim iDR         As Long
    Dim iChgCnt     As Integer
    Dim iCnt        As Integer
    Dim sSpec       As String
    Dim sStdspec    As String
    Dim sBefStdspec As String
    Dim sCharge_no As String
    
    iCnt = 0
    iChgCnt = 0
    
    With ss2
        For iDR = 1 To .MaxRows
            If .Text <> "XAC" And .Text <> "XAF" Then
            
                .Col = SS2_STDSPEC:         sStdspec = Trim(.Text)
                .Col = SS2_BEF_STDSPEC:     sBefStdspec = Trim(.Text)
                
                If sBefStdspec <> "" And sStdspec <> sBefStdspec And (ExpoCheck(sBefStdspec) And Not ExpoCheck(sStdspec)) Then
                    iChgCnt = iChgCnt + 1
                    Exit For
                End If
                
            End If
        Next iDR
        
        For iDR = 1 To .MaxRows
            .ROW = iDR
            .Col = 0
            If .Text = "Update" Then
                .Col = SS2_STDSPEC
                sSpec = Trim(.Text)
                If Trim(txt_CHG_SMP_NO) = "" Or Trim(txt_CHG_STDSPEC) = "" Then
                    If ExpoCheck(sSpec) Then
                        Call Gp_MsgBoxDisplay("必须输入(" & txt_CHG_SMP_NO.Tag & "与" & txt_CHG_STDSPEC.Tag & ")")
                        Exit Sub
                    End If
                End If
                
                If iChgCnt = 0 Then
                    .Col = SS2_CHG_SMP_NO:      .Text = txt_CHG_SMP_NO.Text
                    .Col = SS2_CHG_STDSPEC:     .Text = txt_CHG_STDSPEC.Text
                     iChgCnt = 1
                End If
            End If
        Next iDR
    End With
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC2"), Mc2, False) Then
        sCharge_no = txt_charge_no.Text
        txt_charge_no.Text = ""
        Call Form_Ref
        txt_charge_no.Text = sCharge_no
        Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        Call MenuToolSet
    End If
    
End Sub

Private Sub SDT_PROD_DATE_Click()
     SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub


Private Sub opt_LineFlag_Click(Index As Integer, Value As Integer)
    If opt_LineFlag(2).Value = True Then
       txt_PrcLine = "3"
       opt_LineFlag(2).ForeColor = &HFF&       'red
       opt_LineFlag(3).ForeColor = &H80000012  'black
    ElseIf opt_LineFlag(3).Value = True Then
       txt_PrcLine = "4"
       opt_LineFlag(2).ForeColor = &H80000012  'black
       opt_LineFlag(3).ForeColor = &HFF&       'red
    End If
    ss1.MaxRows = 0
    ss2.MaxRows = 0
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
    Dim intRow      As Integer
    Dim sLot_no     As String
    Dim sSelect     As String
    
    If ROW < 1 Then Exit Sub
    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    With ss1
        For intRow = 1 To .MaxRows
            .Col = 0: .ROW = intRow: .Text = ""
            Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, .ROW, .ROW)
        Next intRow
        
        .Col = 0: .ROW = .ActiveRow: .Text = "Selected"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW, , &HFFFF80)
        
        ss2.MaxRows = 0
        .Col = 1
        txt_SLAB_NO = .Text
        .Col = 2
        sLot_no = .Text
        
        sLoopChk = "**"
        If Gf_Sp_Refer(M_CN1, sc2, Mc2) Then
            Call TextClear
            Call Sample_No_Edit
            ss2.OperationMode = OperationModeNormal
            If TXT_PROD_CD = "SL" Then
              .Col = 1
            Else
              .Col = 2
            End If
            txt_charge_no = .Text
            bCheck = True
        End If
        sLoopChk = ""
        For intRow = 1 To .MaxRows
            .Col = 0
            .ROW = intRow
             sSelect = .Text
            .Col = 2
            .ROW = intRow
            If sLot_no = .Text And Len(sLot_no) = 14 And sSelect = "" Then
               Call ss1_Click(0, intRow)
            End If
        Next intRow
    End With
    
End Sub

Private Sub Sample_No_Edit()
    
    Dim intRow      As Integer
    Dim sPlateNo    As String
    Dim sStdspec    As String
    Dim sBefStdspec As String
    Dim sSmpFl      As String
    Dim sSmpNo      As String
    Dim sProdCd     As String
    Dim sSmp_No     As String
    
    If ss2.MaxRows < 1 Then Exit Sub
    
    With ss2
        For intRow = 1 To .MaxRows
            .ROW = intRow
            .Col = SS2_PROC_CD
            If .Text Like "X??" Then
                .Protect = True
                .Col = SS2_SMP_FLAG: .Col2 = SS2_SMP_NO
                .ROW = intRow:  .Row2 = intRow
                
                .BlockMode = True
                .Lock = True
                .ForeColor = vbBlack
                .BackColor = vbWhite
                .BlockMode = False
                
                .Col = 1: .Col2 = 1
                .ROW = intRow:  .Row2 = intRow
                
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            Else
                .Col = SS2_STDSPEC:         sStdspec = Trim(.Text)
                .Col = SS2_BEF_STDSPEC:     sBefStdspec = Trim(.Text)
                .Col = SS2_PLATE_NO:        sPlateNo = Trim(.Text)
                .Col = SS2_SMP_FLAG:        sSmpFl = Trim(.Text)
                .Col = SS2_SMP_NO:          sSmpNo = Trim(.Text)
                .Col = SS2_PROD_CD:         sProdCd = Trim(.Text)
                
                If sProdCd = "PP" Then
                   sSmp_No = "00"
                Else
                   sSmp_No = "99"
                End If
                
                If txt_smp_no.Text = "" Then
                    If sSmpFl <> "" And Right(sSmpNo, 2) <> sSmp_No Then
                        txt_smp_no.Text = sSmpNo
                    Else
                        If (Len(Trim(txt_smp_no.Text)) <> 14 And sProdCd = "PP") Or (Len(Trim(txt_smp_no.Text)) <> 12 And sProdCd = "HC") Then
                            txt_smp_no.Text = sPlateNo
                        End If
                    End If
                End If
                
                If ExpoCheck(sBefStdspec) Or ExpoCheck(sStdspec) Then
                    If sSmpFl <> "" And sSmpNo <> txt_smp_no.Text Then
'                        txt_CHG_SMP_NO.Text = sSmpNo
                        txt_CHG_STDSPEC.Text = sStdspec
                    Else
                        If sBefStdspec <> "" And sStdspec <> sBefStdspec Then
                            txt_CHG_STDSPEC.Text = sStdspec
                        End If
                    End If
                    
                    If sProdCd = "PP" Then
                       txt_CHG_SMP_NO.Text = Left(txt_smp_no.Text, 12) & sSmp_No
                    Else
                       txt_CHG_SMP_NO.Text = Left(txt_smp_no.Text, 10) & sSmp_No
                    End If
                End If
            End If
            
        Next intRow
    End With
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    Dim PRE As Long
    Dim iDR As Long
    Dim sSpec As String
    Dim iSelCnt As Long
    Dim sCharNo As String
    
    If ROW < 1 Or Col > 0 Then Exit Sub
    If ss1.MaxRows < 1 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 2
    sSpec = Trim(ss1.Text)
    ss1.Col = 1
    sCharNo = Left(Trim(ss1.Text), 8)
    
    iSelCnt = 0
    For iDR = 1 To ss1.MaxRows
        ss1.ROW = iDR
        ss1.Col = 0
        If ss1.Text = "Selected" Then
'            ss1.Col = 1
'            If sCharNo <> Left(Trim(ss1.Text), 8) Then
'                Call Gp_MsgBoxDisplay("不一样炉号")
'                Exit Sub
'            End If
            ss1.Col = 2
            If sSpec <> ss1.Text Then
                Call Gp_MsgBoxDisplay("不一样轧批号")
                Exit Sub
            End If
            iSelCnt = iSelCnt + 1
        End If
    Next iDR
    
    sLoopChk = "**"
    ss1.ROW = ROW
    ss1.Col = 0

    If ss1.Text <> "Selected" Then
        ss1.Col = 0
        ss1.Text = "Selected"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW, , &HFFFF80)
    Else
        If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub

        ss1.Col = 0
        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)

        If iSelCnt < 2 Then
            ss2.MaxRows = 0
        Else
            Call Cmd_Set_Save_Click
        End If

    End If

    iSelCnt = 0
    For iDR = 1 To ss1.MaxRows
        ss1.ROW = iDR
        ss1.Col = 0
        If ss1.Text = "Selected" Then
            iSelCnt = iSelCnt + 1
        End If
    Next iDR

    If iSelCnt = 0 Then Call TextClear
    sLoopChk = ""
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub


Private Sub ss2_ButtonClicked(ByVal Col As Long, ByVal ROW As Long, ByVal ButtonDown As Integer)
    Dim strSmpNO    As String
    Dim sSmpFlag    As String
    Dim sSmpLoc     As String
    Dim lSmpLen     As Long
    Dim sSmpNo      As String
    Dim sStdspec    As String
    Dim sBefStdspec As String
    Dim sSmpFl      As String
    Dim sProdCd     As String
    Dim sSmp_No     As String
    Dim iDR         As Long

    If ROW < 1 Or Trim(sLoopChk) <> "" Then Exit Sub

    sSmpFl = ""

    With ss2
        .ROW = ROW
        .Col = SS2_PROC_CD
        If .Text Like "X??" Then
            Exit Sub
        End If

        sLoopChk = "**"

        .ROW = ROW
        .Col = Col
        If .Value = 1 Then

'            For iDr = 1 To .MaxRows
'                .Row = iDr

                .Col = SS2_PROC_CD
                If Left(.Text, 1) <> "X" Then
                    .Col = 1
                    .Value = 1

                    If Len(txt_SMP_LOC.Text) = 1 Then .Col = SS2_SMP_LOC: .Text = txt_SMP_LOC.Text
                    If Len(txt_SMP_LEN.Text) > 0 Then .Col = SS2_SMP_LEN: .Text = txt_SMP_LEN.Text

                    .Col = SS2_STDSPEC:         sStdspec = Trim(.Text)
                    .Col = SS2_BEF_STDSPEC:     sBefStdspec = Trim(.Text)
                    .Col = SS2_PROD_CD:         sProdCd = Trim(.Text)
                    
                    If sProdCd = "PP" Then
                       sSmp_No = "00"
                    Else
                       sSmp_No = "99"
                    End If

                    If sBefStdspec <> "" And sStdspec <> sBefStdspec And (ExpoCheck(sBefStdspec) And Not ExpoCheck(sStdspec)) Then
                        If (Len(Trim(txt_CHG_SMP_NO.Text)) = 14 And sProdCd = "PP") Or (Len(Trim(txt_CHG_SMP_NO.Text)) = 12 And sProdCd = "HC") Then
                           .Col = SS2_SMP_NO: .Text = txt_CHG_SMP_NO.Text
                        End If
                    Else
                        If (Len(Trim(txt_smp_no.Text)) = 14 And sProdCd = "PP") Or (Len(Trim(txt_smp_no.Text)) = 12 And sProdCd = "HC") Then
                           .Col = SS2_SMP_NO: .Text = txt_smp_no.Text
                        End If
                    End If

'                    .Row = iDr
                    .Col = 0:               .Text = "Update"
                    .Col = SS2_USER_ID:     .Text = sUserID
                    .Col = SS2_PLATE_NO:    strSmpNO = .Text

                    .Col = SS2_SMP_NO
                    If strSmpNO = .Text Then
                        .Col = SS2_SMP_FLAG:    .Text = "Y"
                        .ForeColor = RED
                    Else
                        If sSmpFl = "P" And strSmpNO <> txt_smp_no.Text And (Right(Trim(.Text), 2) = "00" Or Right(Trim(.Text), 2) = "99") Then
                            .Col = SS2_SMP_FLAG:    .Text = sSmpFl
                            .Col = SS2_SMP_LEN:     .Text = "0"

                            .Col = SS2_SMP_FLAG:    .Col2 = SS2_SMP_NO
                            .ROW = ROW:             .Row2 = ROW
'                            .Row = iDr:             .Row2 = iDr

                            .BlockMode = True
                            .Lock = True
                            .BlockMode = False
                            sSmpFl = ""
                        End If

                        .Col = SS2_SMP_FLAG
                        If Trim(.Text) <> "P" Then
                            .Text = "N"
                            .ForeColor = BLACK
                            .Col = SS2_SMP_LEN:  .Text = "0"
                        Else
                            .ForeColor = RED
                        End If
                    End If

                    .Col = SS2_SMP_NO
                    If strSmpNO = txt_smp_no.Text And (Right(Trim(.Text), 2) = "00" Or Right(Trim(.Text), 2) = "99") Then
                        sSmpFl = "P"
                    End If
                End If
'            Next iDr
        Else
            For iDR = 1 To .MaxRows
                .ROW = iDR
                .Col = 1
                .Value = 0

                .Col = 0: .Text = ""
                .Col = SS2_BEF_SMP_FLAG:    sSmpFlag = .Text
                .Col = SS2_BEF_SMP_LOC:     sSmpLoc = .Text
                .Col = SS2_BEF_SMP_LEN:     lSmpLen = Val(.Text & "")
                .Col = SS2_BEF_SMP_NO:      sSmpNo = .Text

                .Col = SS2_SMP_FLAG:        .Text = sSmpFlag
                .Col = SS2_SMP_LOC:         .Text = sSmpLoc
                .Col = SS2_SMP_LEN:         .Text = lSmpLen
                .Col = SS2_SMP_NO:          .Text = sSmpNo
            Next iDR
        End If
    End With
    sLoopChk = ""
End Sub

Private Sub ss2_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim str_orgin As String
    If KeyCode = vbKeyF4 Then
    
        With ss2
            .Col = .ActiveCol
            .ROW = .ActiveRow
            If .ActiveCol = SS2_SMP_LOC Then
                
                str_orgin = .Text
                .Text = ""
                DD.sWitch = "MS"
                DD.sKey = "Q0021"
                DD.rControl.Add Item:=ss2
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
                If str_orgin <> .Text And .Text <> "" Then
                    Call Gp_Sp_UpdateMake(ss2, 2)
                Else
                    .Text = str_orgin
                End If
            End If
        End With
        
    End If
    
End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If ROW > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
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



Private Sub ss2_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    Dim intCheck As Integer
    Dim strSmpNO As String
    Dim strOrginSmpNO As String
    Dim sSmpNo  As String
    
    If Col = 1 Or Col = 0 Then Exit Sub
        
    If ROW > 0 Then
        ss2.ROW = ROW
        ss2.Col = 1

        If ss2.Value <> 1 Then ss2.Value = 1:    Exit Sub
            
        ss2.ROW = ROW
        ss2.Col = Col
        If Col = SS2_SMP_FLAG Then
            ss2.Text = UCase(ss2.Text)
            ss2.Col = SS2_PLATE_NO
            strSmpNO = ss2.Text
            ss2.Col = SS2_SMP_NO
            If strSmpNO = ss2.Text Then
                ss2.Col = SS2_SMP_FLAG: 'ss2.Text = "Y" Modified By YangMeng At 2006.06.01
            Else
                ss2.Col = SS2_SMP_FLAG
                If ss2.Text <> "P" Then
                    ss2.Text = "N"
                    ss2.Col = SS2_SMP_LEN
                    ss2.Text = "0"
                End If
            End If
         ElseIf Col = SS2_SMP_LOC Then
            ss2.Text = UCase(ss2.Text)
            Select Case Trim(ss2.Text)
                   Case "M", "B", "T"
                   Case Else
                        ss2.Text = "T"
            End Select
         ElseIf Col = SS2_SMP_NO Then
            ss2.Col = SS2_PLATE_NO
            strSmpNO = ss2.Text
            ss2.Col = Col
            If Len(ss2.Text) <> Len(strSmpNO) Then 'Or Left(ss2.Text, 8) <> Left(strSmpNO, 8) Then 'Modified By YangMeng At 2007.03.29
                Call Gp_MsgBoxDisplay("试样号错误")
                ss2.Col = SS2_BEF_SMP_NO:      sSmpNo = ss2.Text
                ss2.Col = Col:                 ss2.Text = sSmpNo
            End If

            ss2.Col = Col
            If strSmpNO = ss2.Text Then
                ss2.Col = SS2_SMP_FLAG: ss2.Text = "Y"
                ss2.ForeColor = RED
            Else
                ss2.Col = SS2_SMP_FLAG
                If ss2.Text <> "P" Then
                    ss2.Text = "N"
                    ss2.ForeColor = BLACK
                    ss2.Col = SS2_SMP_LEN
                    ss2.Text = "0"
                Else
                    ss2.ForeColor = RED
                End If
            End If
         End If
    End If
End Sub


Public Function ExpoCheck(sSpec As String) As Boolean

    Dim RS          As New ADODB.Recordset
    Dim iCnt        As Integer
    
    iCnt = 0
    ExpoCheck = False
    
    sQuery = "SELECT  Gf_Expo_Smp_Check('" & sSpec & "')" & vbCrLf
    sQuery = sQuery & "       FROM  DUAL " & vbCrLf
    RS.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        iCnt = Val(RS(0).Value & "")
    End If
    
    RS.Close
    Set RS = Nothing
    
'    iCnt = InStr(1, sSpec, "SM490A") + InStr(1, sSpec, "A709-50T-2") + InStr(1, sSpec, "A709-50F-2") + _
'           InStr(1, sSpec, "PLCA003") + InStr(1, sSpec, "300W") + InStr(1, sSpec, "A572Gr50") + _
'           InStr(1, sSpec, "SN400B") + InStr(1, sSpec, "S355JR") + InStr(1, sSpec, "S355J2G3")
'
   
    If iCnt > 0 Then
        ExpoCheck = True
        Exit Function
    End If
    
End Function


Private Sub txt_charge_no_Change()
'   Dim sMesg As String
'   If Len(txt_charge_no.Text) > 14 Then
'      sMesg = "查询号长度不能超过10位，请确认查询号 ！！！"
'      Call Gp_MsgBoxDisplay(sMesg)
'   End If
End Sub

Private Sub txt_SMP_LOC_Change()
    txt_SMP_LOC.Text = UCase(txt_SMP_LOC.Text)
   Select Case Trim(txt_SMP_LOC.Text)
          Case "M"
                txt_SMP_LOC_NAME.Text = "中部"
          Case "T"
                txt_SMP_LOC_NAME.Text = "头部"
          Case "B"
                txt_SMP_LOC_NAME.Text = "尾部"
          Case Else
                txt_SMP_LOC_NAME.Text = ""
                txt_SMP_LOC.Text = ""
   End Select
End Sub

Private Sub txt_SMP_LOC_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
        txt_SMP_LOC.Text = ""
        DD.sWitch = "MS"
        DD.sKey = "Q0021"
        DD.rControl.Add Item:=txt_SMP_LOC
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
            If Len(Trim(txt_SMP_LOC.Text)) = 1 Then
            txt_SMP_LOC_NAME.Text = Gf_ComnNameFind(M_CN1, "Q0021", Trim(txt_SMP_LOC.Text), 1)
        Else
            txt_SMP_LOC_NAME.Text = ""
        End If
    End If
End Sub

Private Sub txt_CHG_STDSPEC_DblClick()
    Call txt_CHG_STDSPEC_KeyUp(vbKeyF4, 0)
End Sub
Private Sub txt_CHG_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_CHG_STDSPEC

        Call Pf_Common_DD(M_CN1, KeyCode)
    End If
    
'  If KeyCode = vbKeyF4 Then
'         DD.sWitch = "MS"
'         DD.DataDicType = "C"
'         DD.rControl.Add Item:=txt_stdspec_chg
'         DD.rControl.Add Item:=txt_stdspec_name_chg
'
'         Call Pf_Common_DD(M_CN1, KeyCode)
'
'         Exit Sub
'  End If
    
End Sub

Private Sub ss2_set_check()
    Dim intRow As Integer
    Dim strSmpNO As String
    
    If ss2.MaxRows < 1 Then Exit Sub
    With ss2
        For intRow = 1 To .MaxRows
            .Col = SS2_PROC_CD: .ROW = intRow
            If Left(.Text, 1) <> "X" Then
                .Col = 1:           .Value = 1
            End If
        Next intRow
    End With

End Sub

Private Sub txt_SMP_NO_Change()
    Dim iDR As Long
    If ss2.MaxRows < 1 Then Exit Sub
    
    If Len(Trim(txt_smp_no.Text)) = 14 Then
        If Trim(txt_CHG_SMP_NO.Text) <> "" Then
            txt_CHG_SMP_NO.Text = Left(txt_smp_no.Text, 12) & "00"
        End If
    End If
    
    For iDR = 1 To ss2.MaxRows
        ss2.Col = 1: ss2.ROW = iDR:  ss2.Value = 0
    Next iDR
End Sub
Private Sub opt_Product_Click(Index As Integer, Value As Integer)
    If Index = 1 Then
       TXT_PROD_CD = "SL"
       ULabel9.Caption = opt_Product(1).Caption
       opt_Product(1).ForeColor = &HFF&
       opt_Product(2).ForeColor = &H808080
    Else
       TXT_PROD_CD = "LO"
       ULabel9.Caption = opt_Product(2).Caption
       opt_Product(2).ForeColor = &HFF&
       opt_Product(1).ForeColor = &H808080
    End If
End Sub
Private Function Pf_Common_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
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
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "HC"        'Common Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    DD.sQuery = "SELECT CD_SHORT_NAME ""标准代号"", CD_NAME ""标准中文名"" FROM ZP_CD WHERE CD_MANA_NO = 'G0030'"
    
    Call Gf_DD_Display(Conn, DD.sQuery, False)
    
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function






