VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGE2020C 
   Caption         =   "在线钢板入库界面_CGE2020C"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   1455
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_WGT_MAX 
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
      Height          =   330
      Left            =   13800
      TabIndex        =   22
      Top             =   870
      Width           =   1065
   End
   Begin VB.TextBox txt_f_addr 
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
      Left            =   3165
      MaxLength       =   7
      TabIndex        =   21
      Top             =   480
      Width           =   1935
   End
   Begin VB.ComboBox CBO_CUR_INV 
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
      ItemData        =   "CGE2020C.frx":0000
      Left            =   3165
      List            =   "CGE2020C.frx":000D
      TabIndex        =   18
      Top             =   90
      Width           =   705
   End
   Begin VB.TextBox TXT_CNT 
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
      Height          =   330
      Left            =   11025
      TabIndex        =   17
      Top             =   870
      Width           =   525
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
      ItemData        =   "CGE2020C.frx":001D
      Left            =   11250
      List            =   "CGE2020C.frx":002A
      TabIndex        =   5
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txt_t_addr 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8115
      MaxLength       =   7
      TabIndex        =   4
      Top             =   870
      Width           =   1170
   End
   Begin VB.TextBox TXT_PLATE_NO 
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
      Left            =   3165
      TabIndex        =   3
      Top             =   870
      Width           =   1935
   End
   Begin VB.TextBox TXT_WGT 
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
      Height          =   330
      Left            =   11565
      TabIndex        =   2
      Top             =   870
      Width           =   1065
   End
   Begin VB.TextBox txt_stdspec_chg 
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
      Left            =   6735
      MaxLength       =   18
      TabIndex        =   1
      Tag             =   "标准号"
      Top             =   480
      Width           =   3125
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
      Left            =   3870
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   1230
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   5430
      Top             =   90
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
      Left            =   10185
      Top             =   90
      Width           =   1035
      _ExtentX        =   1826
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8130
      Left            =   60
      TabIndex        =   6
      Top             =   1265
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14340
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   1
      PaneTree        =   "CGE2020C.frx":0037
      Begin FPSpread.vaSpread ss1 
         Height          =   8100
         Left            =   15
         TabIndex        =   7
         Top             =   15
         Width           =   9255
         _Version        =   393216
         _ExtentX        =   16325
         _ExtentY        =   14287
         _StockProps     =   64
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
         MaxCols         =   17
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGE2020C.frx":0089
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   8100
         Left            =   9330
         TabIndex        =   8
         Top             =   15
         Width           =   5820
         _Version        =   393216
         _ExtentX        =   10266
         _ExtentY        =   14287
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
         MaxCols         =   12
         MaxRows         =   50
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGE2020C.frx":1EFC
      End
   End
   Begin InDate.ULabel ULabel6 
      Height          =   330
      Left            =   6900
      Top             =   870
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Caption         =   "目标垛位"
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   120
      Top             =   1380
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "后道工序"
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   1860
      Top             =   870
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "物料号"
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
      Height          =   330
      Left            =   9330
      Top             =   870
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      Caption         =   "件数/重量(吨)"
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
   Begin InDate.ULabel ULabel22 
      Height          =   315
      Index           =   1
      Left            =   5430
      Top             =   480
      Width           =   1275
      _ExtentX        =   2249
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   10185
      Top             =   480
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "厚度"
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
   Begin CSTextLibCtl.sidbEdit SDB_THK 
      Height          =   315
      Left            =   11250
      TabIndex        =   10
      Top             =   480
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      ShowZero        =   0   'False
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel25 
      Height          =   315
      Left            =   1860
      Top             =   90
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "当前库"
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
   Begin VB.ComboBox CBO_UST_USE 
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
      ItemData        =   "CGE2020C.frx":282A
      Left            =   1620
      List            =   "CGE2020C.frx":2837
      TabIndex        =   9
      Top             =   1380
      Width           =   1455
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1095
      Left            =   90
      TabIndex        =   12
      Top             =   90
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1931
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox txt_PrcLine 
         Height          =   285
         Left            =   1350
         TabIndex        =   13
         Text            =   " "
         Top             =   60
         Visible         =   0   'False
         Width           =   225
      End
      Begin Threed.SSOption opt_LineFlag 
         Height          =   255
         Index           =   1
         Left            =   330
         TabIndex        =   14
         Top             =   420
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "二号线"
      End
      Begin Threed.SSOption opt_LineFlag 
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   15
         Top             =   90
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "一号线"
      End
      Begin Threed.SSOption opt_LineFlag 
         Height          =   255
         Index           =   2
         Left            =   330
         TabIndex        =   16
         Top             =   750
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "全线"
      End
   End
   Begin InDate.UDate SDT_PROD_DATE_FROM 
      Height          =   315
      Left            =   6735
      TabIndex        =   19
      Tag             =   "起始日期"
      Top             =   90
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
   Begin InDate.UDate SDT_PROD_DATE_TO 
      Height          =   315
      Left            =   8415
      TabIndex        =   20
      Tag             =   "起始日期"
      Top             =   90
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
      Left            =   1860
      Top             =   480
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "起始垛位"
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
   Begin InDate.ULabel ULabel3 
      Height          =   330
      Left            =   12840
      Top             =   870
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   582
      Caption         =   "最大量"
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
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   8250
      TabIndex        =   11
      Top             =   225
      Width           =   255
   End
End
Attribute VB_Name = "CGE2020C"
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
'-- Program Name      在线钢板
'-- Program ID        CGE2020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2007.7.23
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

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Sheet"
    
       Call Gp_Ms_Collection(CBO_CUR_INV, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_stdspec_chg, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_f_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_t_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDB_THK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_PrcLine, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_CNT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_WGT, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
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
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    Call Gp_Sp_Collection(ss2, 1, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGE2020C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGE2020C.P_MODIFY", Key:="P-M"
    sc2.Add Item:="CGE2020C.P_ONEROW", Key:="P-O"
    sc2.Add Item:="CGE2020C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "◎"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
'    Call Gp_Sp_ColHidden(ss1, 7, True)

End Sub

Private Sub CBO_CUR_INV_Click()
    text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", CBO_CUR_INV.Text, 2)
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_ColHidden(ss2, 3, True)
    
    TXT_CNT.Text = ""
    TXT_WGT.Text = ""
    
    If App.Title = "BG" Then
        CBO_CUR_INV.Text = "00"
    ElseIf App.Title = "DG" Then
        CBO_CUR_INV.Text = "WD"
    ElseIf App.Title = "CG" Then
        CBO_CUR_INV.Text = "ZB"
    End If
    
    SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
    SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")

    opt_LineFlag(0).Value = True
    
'    AGE2020C.Height = 9165
'    AGE2020C.Width = 8820
'    AGE2020C.Left = Screen.Width - Me.Width
'    AGE2020C.Top = (Screen.Height - Me.Height) / 2
    
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
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
    Set iSumCol = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    If Col = 0 Then
       Call ss1_DblClick(Col, ROW)
    End If

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub


Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)

    Dim plate_no As String
    Dim iCnt As Integer
    Dim iPlate_cnt As Integer
    Dim iPlate_wgt As Double
    
    Dim tRow  As Integer
    
    iPlate_cnt = 0
    iPlate_wgt = 0

    If ROW <= 0 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 0
    If ss1.Text = "Input" Then
        ss1.Col = 1
        plate_no = ss1.Text
        With ss2
            
            For iCnt = .MaxRows To 1 Step -1
               .Col = 0
               .ROW = iCnt
                If Trim(.Text) = "Input" Then
                   .Col = 2
                    If .Text = plate_no Then
                       .Text = ""
                       .BackColor = &H80000005
                       .Col = 0
                       .Text = ""
                        Exit For
                    End If
                End If
            Next iCnt
             
        End With
        ss1.Col = 0
        ss1.Text = ""
        With ss1
               For iCnt = 1 To .MaxCols Step 1
                    .Col = iCnt
                    .BackColor = &H80000005
               Next iCnt
        End With
        
        With ss1
    
               For iCnt = 1 To .MaxRows Step 1
                    .Col = 0
                    .ROW = iCnt
                     If Trim(.Text) <> "" Then
                         iPlate_cnt = iPlate_cnt + 1
                         .Col = 8
                         iPlate_wgt = iPlate_wgt + .Value
                     End If
               Next iCnt
    
        End With
        TXT_CNT.Text = Str(iPlate_cnt)
        TXT_WGT.Text = Str(iPlate_wgt)
        Exit Sub
    End If
    
    ss1.Col = 1
    plate_no = Trim(ss1.Text)

    If ss2.MaxRows = 0 Then
       Exit Sub
    End If

    ss1.ROW = ROW
    ss1.Col = 0
    ss1.Text = "Input"
    
    With ss1
           For iCnt = 1 To .MaxCols Step 1
                .Col = iCnt
                .BackColor = &HFFC0FF
           Next iCnt
    End With
    
    With ss1

           For iCnt = 1 To .MaxRows Step 1
                .Col = 0
                .ROW = iCnt
                 If Trim(.Text) <> "" Then
                     iPlate_cnt = iPlate_cnt + 1
                     .Col = 8
                     iPlate_wgt = iPlate_wgt + .Value
                 End If
           Next iCnt

    End With
    
    TXT_CNT.Text = Str(iPlate_cnt)
    TXT_WGT.Text = Str(iPlate_wgt)

        
    With ss2
        
        tRow = .ActiveRow
        .ROW = tRow
        .Col = 2
    
    If Len(.Text) = 14 Then
    
         For iCnt = .MaxRows To 1 Step -1
            .Col = 2
            .ROW = iCnt
             If Trim(.Text) = "" Then
                .Text = plate_no
                .Col = 0
                .Text = "Input"
                .Col = 12
                .Text = sUserID
                 Exit Sub
             End If
         Next iCnt
         
    Else
    
        .Col = 2
        .ROW = tRow
         If Trim(.Text) = "" Then
            .Text = plate_no
            .Col = 0
            .Text = "Input"
            .Col = 12
            .Text = sUserID
             If tRow > 1 Then
             Call .SetActiveCell(1, tRow - 1)
             End If
             Exit Sub
         End If
         
    End If
         
    End With
    
End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ROW > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If
    
End Sub



Public Sub Form_Exit()
    Unload Me
End Sub
Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
    TXT_CNT.Text = ""
    TXT_WGT.Text = ""
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim iRow  As Integer
    Dim sRow  As Integer
    Dim tRow  As Integer

    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub

    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    If Len(txt_t_addr) = 7 Then
       If Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       End If
    End If
    
    If ss2.MaxRows = 0 Then Exit Sub
    With ss2
         For iRow = 1 To .MaxRows
            .ROW = iRow
            .Col = 2
             If Trim(.Text) <> "" Then
                sRow = iRow
                Exit For
             End If
             sRow = .MaxRows
         Next iRow
         
         tRow = sRow + 15
         If tRow > .MaxRows Then
            tRow = .MaxRows
         End If
         
         Call .SetActiveCell(1, tRow)
    End With
    
    TXT_CNT.Text = ""
    TXT_WGT.Text = ""
    
    Exit Sub

Refer_Err:

End Sub
Public Sub Form_Pro()

    Dim iRow  As Integer
    Dim sRow  As Integer
    Dim tRow  As Integer
    
    If Gf_Mill_Process(M_CN1, sc2, Mc1, , "P") Then
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    End If
    
    If ss2.MaxRows = 0 Then Exit Sub
    With ss2
         For iRow = 1 To .MaxRows
            .ROW = iRow
            .Col = 2
             If Trim(.Text) <> "" Then
                sRow = iRow
                Exit For
             End If
             sRow = .MaxRows
         Next iRow
         
         tRow = sRow + 15
         If tRow > .MaxRows Then
            tRow = .MaxRows
         End If
         
         Call .SetActiveCell(1, tRow)
    End With
    
    TXT_CNT.Text = ""
    TXT_WGT.Text = ""
       
End Sub
Public Sub Spread_Del()
    
    Dim i As Long
    
    With sc2.Item("Spread")
        
        If .MaxRows < 1 Then Exit Sub
        If .SelBlockRow < 1 Then Exit Sub
        
        For i = .SelBlockRow To .SelBlockRow2
            .ROW = i
            .Col = 2
            If Len(Trim(.Text)) = 14 Then
                .Col = 0
                If Trim(.Text) = "" Then
                    .Text = "Delete"
                End If
            End If
        Next i
        
    End With
    
End Sub

Private Sub CBO_CUR_INV_Change()
    If Len(Trim(CBO_CUR_INV.Text)) = 2 Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", CBO_CUR_INV.Text, 2)
    Else
      text_cur_inv.Text = ""
    End If
End Sub

Private Sub CBO_CUR_INV_DblClick()
    Call CBO_CUR_INV_KeyUp(vbKeyF4, 0)
End Sub
Private Sub CBO_CUR_INV_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=CBO_CUR_INV
        DD.rControl.Add Item:=text_cur_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
     
        If Len(Trim(CBO_CUR_INV.Text)) = 2 Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", CBO_CUR_INV.Text, 2)
        Else
          text_cur_inv.Text = ""
        End If
        
    End If
End Sub

Private Sub txt_f_addr_DblClick()
     Call txt_f_addr_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_f_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        If CBO_CUR_INV.Text = "ZB" Then
           DD.sKey = "F0037"
        ElseIf CBO_CUR_INV.Text = "WG" Then
           DD.sKey = "F0036"
        ElseIf CBO_CUR_INV.Text = "52" Then
           DD.sKey = "F0038"
        Else
           DD.sKey = "X"
        End If
        txt_f_addr.Text = "P"
        DD.rControl.Add Item:=txt_f_addr
'        DD.rControl.Add Item:=txt_o_f_addr_nm
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

End Sub

Private Sub txt_t_addr_DblClick()
     Call txt_t_addr_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_t_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        If CBO_CUR_INV.Text = "ZB" Then
           DD.sKey = "F0037"
        ElseIf CBO_CUR_INV.Text = "WG" Then
           DD.sKey = "F0036"
        ElseIf CBO_CUR_INV.Text = "52" Then
           DD.sKey = "F0038"
        Else
           DD.sKey = "X"
        End If
        txt_t_addr.Text = "P"
        DD.rControl.Add Item:=txt_t_addr
'        DD.rControl.Add Item:=txt_o_f_addr_nm
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
'        txt_o_f_addr.Text = txt_t_addr.Text
        
        Exit Sub
        
    End If

End Sub

Private Sub txt_t_addr_Change()
    If Len(Trim(txt_t_addr.Text)) = 7 Then
        TXT_WGT_MAX.Text = Gf_CarInfFind(M_CN1, txt_t_addr.Text, "R", 1)
    Else
      TXT_WGT_MAX.Text = ""
    End If
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal ROW As Long)

    If ss2.MaxRows < 1 Then Exit Sub

    With ss2
         If Col = 2 Then
            .ROW = ROW + 1
            .Col = 2
            If Trim(.Text) = "" And .ROW <> .MaxRows + 1 Then Exit Sub
            .ROW = ROW
            If Trim(.Text) = "" Then
               .Col = 0
               .Text = "Input"
            Else
               .Col = 0
               .Text = "Update"
            End If
         End If
    End With

End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1

End Sub

Public Sub Spread_Can()

'    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    Dim ss1Row  As Integer
    Dim ss2Row  As Integer
    Dim iCnt  As Integer
    Dim iCnt1  As Integer
    Dim iPlate_no As String
    
    With ss2
         .Col = 0
         .ROW = .ActiveRow
          If .Text = "Input" Then
                For ss2Row = .ROW To 1 Step -1
                   .Col = 2
                   .ROW = ss2Row
                If Len(.Text) = 14 Then
                    iPlate_no = .Text
                   .Text = ""
                   .Col = 0
                   .Text = ""
                    For ss1Row = 1 To ss1.MaxRows
                        ss1.ROW = ss1Row
                        ss1.Col = 0
                        If ss1.Text = "Input" Then
                           ss1.Col = 1
                            If ss1.Text = iPlate_no Then
                                ss1.Col = 0
                                ss1.Text = ""

                                For iCnt1 = 1 To ss1.MaxCols Step 1
                                     ss1.Col = iCnt1
                                     ss1.BackColor = &HFFFFFF
                                Next iCnt1

                                ss1.Col = 7
                                TXT_CNT.Text = Str(Val(TXT_CNT.Text) - 1)
                                TXT_WGT.Text = Str(Val(TXT_WGT.Text) - ss1.Value)
                                If TXT_CNT.Text = "0" Then
                                    TXT_CNT.Text = ""
                                    TXT_WGT.Text = ""
                                End If
                            End If
                        End If
                    Next ss1Row
                End If
                Next ss2Row
          End If
    End With

End Sub
Private Sub txt_stdspec_chg_DblClick()
    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_chg

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub ULabel6_DblClick()
    Dim sMsg As String
    Dim mResult As String
    
    If Gf_Sp_ProceExist(sc2.Item("Spread"), True) Then Exit Sub
    
    If text_cur_inv.Text = "" Then
       sMsg = "请正确选择当前库"
       mResult = MsgBox(sMsg, vbYesNo, "重要提示")
       Exit Sub
    End If
    
    If txt_t_addr.Text <> "" Then
       sMsg = "确定对垛位（" + txt_t_addr.Text + "）进行调整吗？"
       mResult = MsgBox(sMsg, vbYesNo, "重要提示")
       If mResult = vbYes Then
           If Gp_LOC_Exec(CBO_CUR_INV.Text, txt_t_addr.Text) = "" Then
              MsgBox ("垛位调整完毕 ！")
              Call Form_Ref
           Else
              MsgBox (" 垛位调整失败！")
           End If
       End If
       Exit Sub
    End If
End Sub

Public Function Gp_LOC_Exec(Cur_Inv As String, Loc As String) As String

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iRow As Integer

    Dim adoCmd As ADODB.Command

    Screen.MousePointer = vbHourglass

    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256

    sQuery = "{call CGE2020C.P_MODIFY1 ('" + Cur_Inv + "','" + Loc + "',?)}"

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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")

        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg

        Screen.MousePointer = vbDefault
        Gp_LOC_Exec = sErrMessg
        Set adoCmd = Nothing
        Exit Function

    End If

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Gp_LOC_Exec = ""
    Exit Function

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Gp_LOC_Exec = "Process_Exec_ERROR"
    Err.Raise Err.Number, Err.Description & sQuery

End Function
Private Sub opt_LineFlag_Click(Index As Integer, Value As Integer)
    If opt_LineFlag(0).Value = True Then
       txt_PrcLine = "1"
       opt_LineFlag(0).ForeColor = &HFF&       'red
       opt_LineFlag(1).ForeColor = &H80000012  'black
       opt_LineFlag(2).ForeColor = &H80000012  'black
    ElseIf opt_LineFlag(1).Value = True Then
       txt_PrcLine = "2"
       opt_LineFlag(0).ForeColor = &H80000012       'black
       opt_LineFlag(1).ForeColor = &HFF&  'red
       opt_LineFlag(2).ForeColor = &H80000012       'black
    ElseIf opt_LineFlag(2).Value = True Then
       txt_PrcLine = "3"
       opt_LineFlag(0).ForeColor = &H80000012       'black
       opt_LineFlag(1).ForeColor = &H80000012       'black
       opt_LineFlag(2).ForeColor = &HFF&  'red
    End If
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




