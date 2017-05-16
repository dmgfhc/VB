VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKN2090C 
   Caption         =   "炼钢作业指示对接下达及取消界面_AKN2090C"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_heat_mana_no 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   15360
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   22
      Tag             =   "起始炉号"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   15405
      MaxLength       =   2
      TabIndex        =   20
      Tag             =   "工厂"
      Top             =   1440
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt_plt_name 
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
      Left            =   15390
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   19
      Tag             =   "工厂"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.ComboBox cbo_prc_line 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AKN2090C.frx":0000
      Left            =   1365
      List            =   "AKN2090C.frx":0002
      TabIndex        =   5
      Tag             =   "炉座号"
      Top             =   105
      Width           =   600
   End
   Begin VB.ComboBox cbo_prc_line1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AKN2090C.frx":0004
      Left            =   1950
      List            =   "AKN2090C.frx":0006
      TabIndex        =   1
      Tag             =   "炉座号"
      Top             =   105
      Width           =   600
   End
   Begin VB.ComboBox cbo_prc_line2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AKN2090C.frx":0008
      Left            =   2535
      List            =   "AKN2090C.frx":000A
      TabIndex        =   0
      Tag             =   "炉座号"
      Top             =   105
      Width           =   600
   End
   Begin Threed.SSFrame Frame2 
      Height          =   465
      Left            =   3270
      TabIndex        =   2
      Top             =   30
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin Threed.SSOption Opt_InqBof 
         Height          =   285
         Left            =   225
         TabIndex        =   3
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
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
         Caption         =   "转炉标准"
         Value           =   -1
      End
      Begin Threed.SSOption Opt_InqCcm 
         Height          =   285
         Left            =   1530
         TabIndex        =   4
         Top             =   90
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
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
         Caption         =   "连铸标准"
      End
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8235
      Left            =   120
      TabIndex        =   6
      Top             =   570
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   14526
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AKN2090C.frx":000C
      Begin FPSpread.vaSpread ss1 
         Height          =   4080
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   5835
         _Version        =   393216
         _ExtentX        =   10292
         _ExtentY        =   7197
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   25
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2090C.frx":00DE
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   4080
         Left            =   5895
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   5430
         _Version        =   393216
         _ExtentX        =   9578
         _ExtentY        =   7197
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   25
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2090C.frx":0DDC
      End
      Begin FPSpread.vaSpread ss5 
         Height          =   4080
         Left            =   11385
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   3825
         _Version        =   393216
         _ExtentX        =   6747
         _ExtentY        =   7197
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   25
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2090C.frx":1AAE
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4095
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4140
         Width           =   5835
         _Version        =   393216
         _ExtentX        =   10292
         _ExtentY        =   7223
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   20
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2090C.frx":2766
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   4095
         Left            =   5895
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4140
         Width           =   5430
         _Version        =   393216
         _ExtentX        =   9578
         _ExtentY        =   7223
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   20
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2090C.frx":31FF
      End
      Begin FPSpread.vaSpread ss6 
         Height          =   4095
         Left            =   11385
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4140
         Width           =   3825
         _Version        =   393216
         _ExtentX        =   6747
         _ExtentY        =   7223
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   20
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2090C.frx":3C7B
      End
   End
   Begin CSTextLibCtl.sidbEdit sdb_from 
      Height          =   285
      Left            =   8205
      TabIndex        =   13
      Top             =   8805
      Visible         =   0   'False
      Width           =   465
      _Version        =   262145
      _ExtentX        =   820
      _ExtentY        =   503
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
      StartText.x     =   2
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      NumDecDigits    =   0
      NumIntDigits    =   8
      MaxValue        =   9999
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_to 
      Height          =   285
      Left            =   8685
      TabIndex        =   14
      Top             =   8805
      Visible         =   0   'False
      Width           =   465
      _Version        =   262145
      _ExtentX        =   820
      _ExtentY        =   503
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   " 0.00"
      StartText.x     =   2
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   13
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      NumDecDigits    =   0
      NumIntDigits    =   8
      MaxValue        =   9999
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   150
      Top             =   105
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "炉座号"
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   465
      Left            =   6075
      TabIndex        =   15
      Top             =   30
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox txt_heat_mana_no1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         Tag             =   "起始炉号"
         Top             =   75
         Width           =   1230
      End
      Begin VB.ComboBox cbo_dj 
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
         ItemData        =   "AKN2090C.frx":46F7
         Left            =   3810
         List            =   "AKN2090C.frx":4704
         TabIndex        =   21
         Tag             =   "对接指示"
         Top             =   75
         Width           =   1710
      End
      Begin Threed.SSPanel SSPrtn 
         Height          =   420
         Left            =   8040
         TabIndex        =   16
         Top             =   15
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "返送"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPpdt 
         Height          =   420
         Left            =   6870
         TabIndex        =   17
         Top             =   15
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   16761087
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "生产中"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPsend 
         Height          =   420
         Left            =   5580
         TabIndex        =   18
         Top             =   15
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   741
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "已下达"
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   2640
         Top             =   75
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "对接指示"
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
         Left            =   120
         Top             =   75
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         Caption         =   "炉座号"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   15360
      Top             =   1080
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      Caption         =   "工厂"
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
Attribute VB_Name = "AKN2090C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       System
'-- Sub_System Name
'-- Program Name
'-- Program ID        AKN2030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.6.23
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
Public Complete As Boolean
Public P_MODE As String             'CHANGE PRC_LINE = 'U', MOVE = 'M', DELETE = 'D', CANCEL = 'C', SEND = 'L',   TIME = 'T'
Public p_cur_prd As Integer
Public iProd As String
Public iRet  As String
Public p_up_down As String
Public Chg_Lf    As Boolean         'LF Change Check
Public Chg_VD    As Boolean         'VD Change Check
Public Chg_RH    As Boolean         'RH Change Check
Public Ref_FL    As Boolean
Public sAut As String
'Public CS As String

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

Dim pContro3 As New Collection      'Master Primary Key Collection
Dim nContro3 As New Collection      'Master Necessary Collection
Dim mContro3 As New Collection      'Master Maxlength check Collection
Dim iContro3 As New Collection      'Master Insert Collection
Dim rContro3 As New Collection      'Master Refer Collection
Dim cContro3 As New Collection      'Master Copy Collection
Dim aContro3 As New Collection      'Master -> Spread Collection
Dim lContro3 As New Collection      'Master Lock Collection

Dim pContro4 As New Collection      'Master Primary Key Collection
Dim nContro4 As New Collection      'Master Necessary Collection
Dim mContro4 As New Collection      'Master Maxlength check Collection
Dim iContro4 As New Collection      'Master Insert Collection
Dim rContro4 As New Collection      'Master Refer Collection
Dim cContro4 As New Collection      'Master Copy Collection
Dim aContro4 As New Collection      'Master -> Spread Collection
Dim lContro4 As New Collection      'Master Lock Collection

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

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection

Dim pColumn5 As New Collection      'Spread Primary Key Collection
Dim nColumn5 As New Collection      'Spread necessary Column Collection
Dim mColumn5 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn5 As New Collection      'Spread Insert Column Collection
Dim aColumn5 As New Collection      'Master -> Spread Column Collection
Dim lColumn5 As New Collection      'Spread Lock Column Collection

Dim pColumn6 As New Collection      'Spread Primary Key Collection
Dim nColumn6 As New Collection      'Spread necessary Column Collection
Dim mColumn6 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn6 As New Collection      'Spread Insert Column Collection
Dim aColumn6 As New Collection      'Master -> Spread Column Collection
Dim lColumn6 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection
Dim Mc4 As New Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Sc4 As New Collection           'Spread Collection
Dim Sc5 As New Collection           'Spread Collection
Dim Sc6 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim errMsg   As String
Dim iSelRow  As Integer
Dim txt_AFT_Prc_line As String
Dim txt_AFT_SS_Col As Integer
Dim txt_AFT_SS_Row As Integer

Dim lHeat_Edt_Seq_Fr As Long
Dim lHeat_Edt_Seq_To As Long
Dim lHeat_Edt_Seq_Ta As Long

Private Sub Form_Define()

    Dim iCol As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(cbo_prc_line, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(Opt_InqBof, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        
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
    Call Gp_Ms_Collection(txt_heat_mana_no, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
          Call Gp_Ms_Collection(Opt_InqBof, "p", " ", " ", " ", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
       
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    Call Gp_Ms_Collection(cbo_prc_line1, "p", "n", " ", " ", "r", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
       Call Gp_Ms_Collection(Opt_InqBof, "p", " ", " ", " ", " ", " ", " ", pContro3, nContro3, mContro3, iContro3, rContro3, aContro3, lContro3)
    
    'MASTER Collection
    Mc3.Add Item:=pContro3, Key:="pControl"
    Mc3.Add Item:=nContro3, Key:="nControl"
    Mc3.Add Item:=mContro3, Key:="mControl"
    Mc3.Add Item:=iContro3, Key:="iControl"
    Mc3.Add Item:=rContro3, Key:="rControl"
    Mc3.Add Item:=cContro3, Key:="cControl"
    Mc3.Add Item:=aContro3, Key:="aControl"
    Mc3.Add Item:=lContro3, Key:="lControl"
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
    Call Gp_Ms_Collection(cbo_prc_line2, "p", "n", " ", " ", "r", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
       Call Gp_Ms_Collection(Opt_InqBof, "p", " ", " ", " ", " ", " ", " ", pContro4, nContro4, mContro4, iContro4, rContro4, aContro4, lContro4)
    
    'MASTER Collection
    Mc4.Add Item:=pContro4, Key:="pControl"
    Mc4.Add Item:=nContro4, Key:="nControl"
    Mc4.Add Item:=mContro4, Key:="mControl"
    Mc4.Add Item:=iContro4, Key:="iControl"
    Mc4.Add Item:=rContro4, Key:="rControl"
    Mc4.Add Item:=cContro4, Key:="cControl"
    Mc4.Add Item:=aContro4, Key:="aControl"
    Mc4.Add Item:=lContro4, Key:="lControl"
     
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AKN2090C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AKN2090C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
   
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss3.MaxCols
        Call Gp_Sp_Collection(ss3, iCol, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Next iCol
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AKN2090C.P_REFER1", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=3, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss4.MaxCols
        Call Gp_Sp_Collection(ss4, iCol, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Next iCol
    
    'Spread_Collection
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:="AKN2090C.P_REFER2", Key:="P-R"
    Sc4.Add Item:=pColumn4, Key:="pColumn"
    Sc4.Add Item:=nColumn4, Key:="nColumn"
    Sc4.Add Item:=aColumn4, Key:="aColumn"
    Sc4.Add Item:=mColumn4, Key:="mColumn"
    Sc4.Add Item:=iColumn4, Key:="iColumn"
    Sc4.Add Item:=lColumn4, Key:="lColumn"
    Sc4.Add Item:=1, Key:="First"
    Sc4.Add Item:=ss4.MaxCols, Key:="Last"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss5.MaxCols
        Call Gp_Sp_Collection(ss5, iCol, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Next iCol
    
    'Spread_Collection
    Sc5.Add Item:=ss5, Key:="Spread"
    Sc5.Add Item:="AKN2090C.P_REFER1", Key:="P-R"
    Sc5.Add Item:=pColumn5, Key:="pColumn"
    Sc5.Add Item:=nColumn5, Key:="nColumn"
    Sc5.Add Item:=aColumn5, Key:="aColumn"
    Sc5.Add Item:=mColumn5, Key:="mColumn"
    Sc5.Add Item:=iColumn5, Key:="iColumn"
    Sc5.Add Item:=lColumn5, Key:="lColumn"
    Sc5.Add Item:=3, Key:="First"
    Sc5.Add Item:=ss5.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss6.MaxCols
        Call Gp_Sp_Collection(ss6, iCol, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Next iCol
    
    'Spread_Collection
    Sc6.Add Item:=ss6, Key:="Spread"
    Sc6.Add Item:="AKN2090C.P_REFER2", Key:="P-R"
    Sc6.Add Item:=pColumn6, Key:="pColumn"
    Sc6.Add Item:=nColumn6, Key:="nColumn"
    Sc6.Add Item:=aColumn6, Key:="aColumn"
    Sc6.Add Item:=mColumn6, Key:="mColumn"
    Sc6.Add Item:=iColumn6, Key:="iColumn"
    Sc6.Add Item:=lColumn6, Key:="lColumn"
    Sc6.Add Item:=1, Key:="First"
    Sc6.Add Item:=ss6.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    Proc_Sc.Add Item:=Sc4, Key:="Sc4"
    Proc_Sc.Add Item:=Sc5, Key:="Sc5"
    Proc_Sc.Add Item:=Sc6, Key:="Sc6"

    Me.KeyPreview = True
    Me.Opt_InqBof.BackColor = &HE0E0E0
    Me.Opt_InqCcm.BackColor = &HE0E0E0
    Me.BackColor = &HE0E0E0
    

    Call Gp_Sp_ColHidden(ss1, 22, True)    'Heat_edt_seq
    
    Call Gp_Sp_ColHidden(ss2, 1, True)     'SEQ_NO
    Call Gp_Sp_ColHidden(ss2, 18, True)    'STATUS
    
    'Call Gp_Sp_ColHidden(ss3, 16, True)   'l2_send y/n
    Call Gp_Sp_ColHidden(ss3, 22, True)    'Heat_edt_seq
    
    Call Gp_Sp_ColHidden(ss4, 1, True)     'SEQ_NO
    Call Gp_Sp_ColHidden(ss4, 18, True)    'STATUS
    
    Call Gp_Sp_ColHidden(ss5, 22, True)    'Heat_edt_seq
    
    Call Gp_Sp_ColHidden(ss6, 1, True)     'SEQ_NO
    Call Gp_Sp_ColHidden(ss6, 18, True)    'STATUS
    


End Sub


Private Sub cbo_prc_line_Change()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub cbo_prc_line_Click()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub cbo_prc_line1_Change()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub cbo_prc_line1_Click()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub cbo_prc_line2_Change()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub cbo_prc_line2_Click()

    If Ref_FL = False Then Exit Sub
    Call Form_Ref
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Dim sQuery As String
    Dim bDyanmic_start As Boolean
    Dim Dynamic_Slab As String

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
  
    Dynamic_Slab = "SC1"
    sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
    
    If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
    
        Dynamic_Slab = "SC2"
        sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
    
        If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
        
            Dynamic_Slab = "SC3"
            sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
    
            If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
                bDyanmic_start = True
            Else
                bDyanmic_start = False
            End If
        
        Else
            bDyanmic_start = False
        End If
    
    Else
        bDyanmic_start = False
    End If
    
'    Call Dynamic_Slab_ScreenSet(bDyanmic_start)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc3("nControl"))
    Call Gp_Ms_NeceColor(Mc4("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc4.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc5.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc6.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc4.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc5.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc6.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(Sc4)
    Call Gf_Sp_Cls(Sc5)
    Call Gf_Sp_Cls(Sc6)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc4.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc5.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc6.Item("Spread"), "K-System.INI", Me.Name)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "K-System.INI", Me.Name)
    
    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    Ref_FL = False
    
    cbo_prc_line.Clear
    cbo_prc_line.AddItem "1"
    cbo_prc_line.AddItem "2"
    cbo_prc_line.AddItem "3"
    cbo_prc_line.ListIndex = 0
    
    cbo_prc_line1.Clear
    cbo_prc_line1.AddItem "1"
    cbo_prc_line1.AddItem "2"
    cbo_prc_line1.AddItem "3"
    cbo_prc_line1.ListIndex = 1
    
    cbo_prc_line2.Clear
    cbo_prc_line2.AddItem "1"
    cbo_prc_line2.AddItem "2"
    cbo_prc_line2.AddItem "3"
    cbo_prc_line2.ListIndex = 2

    txt_heat_mana_no.Text = ""

    Ref_FL = True

    Call Form_Ref
   
  
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "K-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc4.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc5.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc6.Item("Spread"), "K-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
    
    Set pContro3 = Nothing
    Set nContro3 = Nothing
    Set iContro3 = Nothing
    Set rContro3 = Nothing
    Set cContro3 = Nothing
    Set aContro3 = Nothing
    Set lContro3 = Nothing
    Set mContro3 = Nothing
    
    Set pContro4 = Nothing
    Set nContro4 = Nothing
    Set iContro4 = Nothing
    Set rContro4 = Nothing
    Set cContro4 = Nothing
    Set aContro4 = Nothing
    Set lContro4 = Nothing
    Set mContro4 = Nothing
    
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
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing
    
    Set iColumn5 = Nothing
    Set pColumn5 = Nothing
    Set lColumn5 = Nothing
    Set nColumn5 = Nothing
    Set mColumn5 = Nothing
    Set aColumn5 = Nothing
    
    Set iColumn6 = Nothing
    Set pColumn6 = Nothing
    Set lColumn6 = Nothing
    Set nColumn6 = Nothing
    Set mColumn6 = Nothing
    Set aColumn6 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Mc4 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Sc4 = Nothing
    Set Sc5 = Nothing
    Set Sc6 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
    
        If Gf_Sp_Cls(sc1) Then
        
            Ref_FL = False
            
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gf_Sp_Cls(Sc3)
            Call Gf_Sp_Cls(Sc4)
            Call Gf_Sp_Cls(Sc5)
            Call Gf_Sp_Cls(Sc6)
            
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call MenuTool_ReSet
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            txt_heat_mana_no.Enabled = True

            txt_plt.Text = "B1"
            Call txt_plt_KeyUp(0, 0)
            
            txt_heat_mana_no.Text = ""
       
            
            Ref_FL = True
            cbo_prc_line.ListIndex = 0
            cbo_prc_line1.ListIndex = 1
            cbo_prc_line2.ListIndex = 2
        End If
        
    End If
    
End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    Dim iRow As Integer
    Dim iCol As Integer
    Dim sQuery As String
    Dim PGM_ID As String
    Dim Ref_FL As String
    
  
 
    txt_heat_mana_no.Text = ""
 
    Ref_FL = "0"
    

 
    
'    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc3, Mc3, Mc3("nControl"), Mc3("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc5, Mc4, Mc4("nControl"), Mc4("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Ref_FL = "1" Then
    
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Gp_Sp_EvenRowBackcolor(sc1.Item("Spread"))
        Call Gp_Sp_EvenRowBackcolor(Sc3.Item("Spread"))
        Call Gp_Sp_EvenRowBackcolor(Sc5.Item("Spread"))
        
        Call Gf_Sp_Cls(sc2)
        Call Gf_Sp_Cls(Sc4)
        Call Gf_Sp_Cls(Sc6)
        
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(9).Enabled = False
        
        ss1.OperationMode = OperationModeNormal
        ss3.OperationMode = OperationModeNormal
        ss5.OperationMode = OperationModeNormal
        
        SSPsend.Visible = True
        SSPpdt.Visible = True
        SSPrtn.Visible = True

        Call Spread_Color_Setting(ss1)
        Call Spread_Color_Setting(ss3)
        Call Spread_Color_Setting(ss5)
       
       
    
    Else
        Call Gf_Sp_Cls(sc2)
        Call Gf_Sp_Cls(Sc3)
        Call Gf_Sp_Cls(Sc4)
        Call Gf_Sp_Cls(Sc5)
        Call Gf_Sp_Cls(Sc6)
    End If
            
End Sub

Public Sub Spread_Color_Setting(oSpr As vaSpread)

    Dim iRow As Integer
    Dim iCol As Integer
    
    With oSpr
    
        If oSpr.Name = "ss2" Or oSpr.Name = "ss4" Or oSpr.Name = "ss6" Then
        
            For iRow = 1 To .MaxRows
            
                .Row = iRow:  .Col = 18
                
                If .Text = "B" Then
                
                    For iCol = 1 To .MaxCols
                        .Col = iCol
                        .BackColor = SSPpdt.BackColor
                    Next iCol
                    
                ElseIf .Text = "Y" Then
                
                    For iCol = 1 To .MaxCols
                        .Col = iCol
                        .BackColor = SSPsend.BackColor
                    Next iCol
                    
                End If
                
            Next iRow
            
            Exit Sub
        
        End If
        
        .Col = 1
        
        For iRow = 1 To .MaxRows
            .Row = iRow:  .Col = 19
            
            If .Text = "Y" Then
                For iCol = 1 To .MaxCols
                    .Col = iCol
                    .BackColor = SSPsend.BackColor
                Next iCol
            End If
            
        Next iRow
        
        For iRow = 1 To .MaxRows
            .Row = iRow:  .Col = 24
            
            If .Text = "Y" Then
                For iCol = 1 To .MaxCols
                    .Col = iCol
                    .ForeColor = BLUE
                Next iCol
            End If
            
        Next iRow
          
        For iRow = 1 To .MaxRows
            .Col = 16:  .Row = iRow
           
            If Trim(.Text) <> "" Then
               p_cur_prd = iRow
               For iCol = 1 To .MaxCols
                   .Col = iCol
                   .BackColor = SSPpdt.BackColor
               Next
            End If
          
            .Col = 17
            .Row = iRow
          
            If .Text > "0" Then
               For iCol = 1 To .MaxCols
                   .Col = iCol
                   .BackColor = SSPrtn.BackColor
               Next
            End If
            
        Next
        

       For iRow = 1 To .MaxRows
            .Row = iRow:  .Col = 16

            If .Text = "B" Then
            
               .Row = iRow:  .Col = 25
               
                If .Text = "" Then
                    For iCol = 1 To .MaxCols
                        .Col = iCol
                        .ForeColor = &HFF
    
                    Next iCol
                    
                End If
                
                
           End If

        Next iRow
        
    End With
    
End Sub

Public Sub Spread_Forzens_Setting()
    
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Spread_Can()

    Dim sPrcLine As String
    Dim sMltProc As String
    Dim iRow     As Integer
    
    For iRow = lBlkrow1 To lBlkrow2
        With ss1
            .Row = iRow
            .Col = 0:     .Text = ""
            .Col = 19:    sPrcLine = .Text
            .Col = 20:    sMltProc = .Text
            
            .Col = 2:    .Text = sPrcLine
            .Col = 8:    .Text = sMltProc
        End With
    Next iRow
    
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Active_Spread, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Pro()

On Error GoTo Process_Error

    Dim OutParam(1, 4) As Variant
    Dim errMsg As String
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    Dim sFrom_No As String
    Dim sTo_No As String
    Dim sTarget_No As String
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
     
    sQuery = "{call AKN2090C.P_MODIFY ('" & txt_plt.Text & "',  '" & txt_heat_mana_no1.Text & "', '" & cbo_dj.Text & "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    M_CN1.BeginTrans
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        errMsg = sErrMessg
        M_CN1.RollbackTrans
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Exit Sub
    End If
    
    M_CN1.CommitTrans
    Set adoCmd = Nothing
    
    Call Form_Ref
    
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    sErrMessg = Err.Description & sQuery
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub
Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub Search_Last_Line()

    Dim iRow        As Integer
    
    iSelRow = 1
    
    With ss1
    
        For iRow = .MaxRows To 1 Step -1
            .Row = iRow
            .Col = 4
            If .BackColor = SSPpdt.BackColor Then
                iSelRow = iRow
                Exit Sub
            End If
        Next
        
    End With

End Sub

Public Function sf_Sp_ProceExist() As Integer

    Dim iRow        As Integer
    Dim sColor      As String
    
    sf_Sp_ProceExist = 0
    
    With ss1
    
        For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 0
            If Trim(.Text) = "Input" Or Trim(.Text) = "Update" Or Trim(.Text) = "Delete" Then
                sf_Sp_ProceExist = 1
            End If
            
            .Col = 4
             sColor = .BackColor
             
             .Col = 2:   .Col2 = 2
             .BackColor = sColor
             
             .Col = 9:   .Col2 = 9
             .BackColor = sColor
        Next
        
    End With
    
    MDIMain.MenuTool.Buttons(9).Enabled = False
    
End Function



Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor, sHeat, sTemp As String
    Dim sChgPrcLine          As String
    Dim sL2SendFL            As String
    
    Dim i As Integer
    Dim iRow As Integer


    Set Active_Spread = Me.ss1
    If Row <= 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 1
    txt_heat_mana_no.Text = ss1.Text
    txt_heat_mana_no1.Text = ss1.Text
    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)
    
   


    
   
    lBlkrow1 = Row
    lBlkrow2 = Row
    
    If ss1.MaxRows < 1 Then
    
        Call Gf_Sp_Cls(sc2)
        Exit Sub
    
    End If
    
    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub

    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    ss2.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss2)
    Call Spread_Color_Setting(ss2)
    
     For iRow = 1 To ss2.MaxRows
    
               ss2.Row = iRow
               ss2.Col = 3
                If ss2.Text = "Y" Then
                  For i = 1 To ss2.MaxCols
                       ss2.Col = i
                       ss2.ForeColor = &HC000&
                  Next
                End If
                
                
               ss2.Row = iRow
               ss2.Col = 20
                If ss2.Text = "Y" Then
                  For i = 1 To ss2.MaxCols
                       ss2.Col = i
                       ss2.ForeColor = &HFF&
                  Next
                End If

      
     Next iRow
    

    With ss1
        For iRow1 = .ActiveRow To .MaxRows
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            With ss2
              .Col = 4
              For iRow2 = 1 To .MaxRows
                  .Row = iRow2
                  'If stemp <> "" And stemp <> Left(.Text, 8) Then
                        If Left(.Text, 8) = sHeat Then
                           For iCol = 1 To .MaxCols
                               .Col = iCol
                               .BackColor = sColor

                           Next iCol
                           sTemp = .Text
                        End If

                        If sTemp <> "" And sTemp <> Left(.Text, 8) Then
                           sTemp = ""
                           Exit For
                        End If
                  'End If
              .Col = 4
              Next iRow2
            End With

        Next iRow1
    
    End With
    
End Sub

Private Sub ss1_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss1.MaxCols
        ss3.ColWidth(iCol) = ss1.ColWidth(iCol)
        ss5.ColWidth(iCol) = ss1.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim iRow, iCnt As Integer
    Dim sColor, M_TEMP As String
    
   
            
    ss1.Row = Row
    ss1.Col = 1
    txt_heat_mana_no.Text = ss1.Text
    
    txt_heat_mana_no1.Text = ss1.Text
   
    
    Call Search_Last_Line

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
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

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss2
    
End Sub

Private Sub ss2_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss2.MaxCols
        ss4.ColWidth(iCol) = ss2.ColWidth(iCol)
        ss6.ColWidth(iCol) = ss2.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If
    
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If

End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor, sHeat, sTemp As String
    Dim sChgPrcLine          As String
    
    Dim i As Integer
    Dim iRow As Integer


    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss3
    If Row <= 0 Then Exit Sub
    
    If ss3.MaxRows < 1 Then
        Call Gf_Sp_Cls(Sc4)
        Exit Sub
    End If
    
    ss3.Row = Row
    ss3.Col = 1
    txt_heat_mana_no.Text = ss3.Text
    txt_heat_mana_no1.Text = ""
    
    Call Gf_Sp_Refer(M_CN1, Sc4, Mc2, , , False)
    ss4.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss4)
    Call Spread_Color_Setting(ss4)
    
    
    For iRow = 1 To ss4.MaxRows
    
               ss4.Row = iRow
               ss4.Col = 3
                If ss4.Text = "Y" Then
                  For i = 1 To ss4.MaxCols
                       ss4.Col = i
                       ss4.ForeColor = &HC000&
                  Next
                End If
                
               ss4.Row = iRow
               ss4.Col = 20
                If ss4.Text = "Y" Then
                  For i = 1 To ss4.MaxCols
                       ss4.Col = i
                       ss4.ForeColor = &HFF&
                  Next
                End If
                
        
     Next iRow

    
    
   
    With ss3
    
        For iRow1 = .ActiveRow To .MaxRows
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            
            With ss4
            
                .Col = 4
                For iRow2 = 1 To .MaxRows
                    .Row = iRow2
                    If Left(.Text, 8) = sHeat Then
                       For iCol = 1 To .MaxCols
                           .Col = iCol
                           .BackColor = sColor
                       Next iCol
                       sTemp = .Text
                    End If
                    
                    If sTemp <> "" And sTemp <> Left(.Text, 8) Then
                       sTemp = ""
                       Exit For
                    End If
                    .Col = 4
                Next iRow2
                
            End With

        Next iRow1
        
    End With
    
    
     

    
End Sub

Private Sub ss3_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss3.MaxCols
        ss1.ColWidth(iCol) = ss3.ColWidth(iCol)
        ss5.ColWidth(iCol) = ss3.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss4
    
End Sub

Private Sub ss4_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss4.MaxCols
        ss2.ColWidth(iCol) = ss4.ColWidth(iCol)
        ss6.ColWidth(iCol) = ss4.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss5_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss5.MaxCols
        ss1.ColWidth(iCol) = ss5.ColWidth(iCol)
        ss3.ColWidth(iCol) = ss5.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss6_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss6
    
End Sub

Private Sub ss5_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor, sHeat, sTemp As String
    Dim sChgPrcLine          As String
    
    Dim i As Integer
    Dim iRow As Integer

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss5
    If Row <= 0 Then Exit Sub
    
    If ss5.MaxRows < 1 Then
        Call Gf_Sp_Cls(Sc6)
        Exit Sub
    End If
    
    ss5.Row = Row
    ss5.Col = 1
    txt_heat_mana_no.Text = ss5.Text
    txt_heat_mana_no1.Text = ""
    
    Call Gf_Sp_Refer(M_CN1, Sc6, Mc2, , , False)
    ss6.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss6)
    Call Spread_Color_Setting(ss6)
    
    For iRow = 1 To ss6.MaxRows
    
               ss6.Row = iRow
               ss6.Col = 3
                If ss6.Text = "Y" Then
                  For i = 1 To ss6.MaxCols
                       ss6.Col = i
                       ss6.ForeColor = &HC000&
                  Next
                End If
                
               ss6.Row = iRow
               ss6.Col = 20
                If ss6.Text = "Y" Then
                  For i = 1 To ss6.MaxCols
                       ss6.Col = i
                       ss6.ForeColor = &HFF&
                  Next
                End If
     Next iRow
    

    With ss5
    
        For iRow1 = .ActiveRow To .MaxRows
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            
            With ss6
            
                .Col = 4
                For iRow2 = 1 To .MaxRows
                    .Row = iRow2
                    If Left(.Text, 8) = sHeat Then
                       For iCol = 1 To .MaxCols
                           .Col = iCol
                           .BackColor = sColor
                       Next iCol
                       sTemp = .Text
                    End If
                    
                    If sTemp <> "" And sTemp <> Left(.Text, 8) Then
                       sTemp = ""
                       Exit For
                    End If
                    .Col = 4
                Next iRow2
                
            End With

        Next iRow1
        
    End With

End Sub

Private Sub ss6_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss6.MaxCols
        ss2.ColWidth(iCol) = ss6.ColWidth(iCol)
        ss4.ColWidth(iCol) = ss6.ColWidth(iCol)
    Next iCol

End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_name
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
        
        If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_name.Text = ""
        End If
        
    End If

End Sub




Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(9).Enabled = False                  'Row Cancel
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

