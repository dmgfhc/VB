VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGB2030C 
   Caption         =   "再设计申请界面_CGB2030C"
   ClientHeight    =   9495
   ClientLeft      =   1605
   ClientTop       =   1410
   ClientWidth     =   15120
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9345
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   16484
      _Version        =   196609
      BorderStyle     =   1
      BackColor       =   14737632
      PaneTree        =   "CGB2030C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter3 
         Height          =   5520
         Left            =   15
         TabIndex        =   17
         Top             =   15
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   9737
         _Version        =   196609
         PaneTree        =   "CGB2030C.frx":0052
         Begin Threed.SSPanel SSPanel2 
            Height          =   525
            Left            =   30
            TabIndex        =   19
            Top             =   30
            Width           =   14895
            _ExtentX        =   26273
            _ExtentY        =   926
            _Version        =   196609
            BackColor       =   14737632
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.OptionButton opt_rhf 
               BackColor       =   &H00E0E0E0&
               Caption         =   "四号炉"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   3
               Left            =   5310
               TabIndex        =   24
               Top             =   150
               Width           =   1035
            End
            Begin VB.OptionButton opt_rhf 
               BackColor       =   &H00E0E0E0&
               Caption         =   "三号炉"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   2
               Left            =   4040
               TabIndex        =   23
               Top             =   150
               Width           =   1035
            End
            Begin VB.OptionButton opt_rhf 
               BackColor       =   &H00E0E0E0&
               Caption         =   "二号炉"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   1
               Left            =   2770
               TabIndex        =   22
               Top             =   150
               Width           =   1035
            End
            Begin VB.OptionButton opt_rhf 
               BackColor       =   &H00E0E0E0&
               Caption         =   "一号炉"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   0
               Left            =   1500
               MaskColor       =   &H8000000F&
               TabIndex        =   21
               Top             =   150
               Width           =   1035
            End
            Begin VB.TextBox txt_PrcLine 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7245
               Locked          =   -1  'True
               TabIndex        =   20
               Text            =   " "
               Top             =   120
               Visible         =   0   'False
               Width           =   345
            End
            Begin InDate.ULabel ULabel10 
               Height          =   315
               Left            =   120
               Top             =   120
               Width           =   1350
               _ExtentX        =   2381
               _ExtentY        =   556
               Caption         =   "加热炉"
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
         Begin FPSpread.vaSpread ss1 
            Height          =   4845
            Left            =   30
            TabIndex        =   18
            Top             =   645
            Width           =   14895
            _Version        =   393216
            _ExtentX        =   26273
            _ExtentY        =   8546
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
            MaxCols         =   12
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGB2030C.frx":00A4
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   3705
         Left            =   15
         TabIndex        =   1
         Top             =   5625
         Width           =   14955
         _ExtentX        =   26379
         _ExtentY        =   6535
         _Version        =   196609
         BackColor       =   14737632
         PaneTree        =   "CGB2030C.frx":1D15
         Begin Threed.SSPanel SSPanel1 
            Height          =   840
            Left            =   30
            TabIndex        =   3
            Top             =   30
            Width           =   14895
            _ExtentX        =   26273
            _ExtentY        =   1482
            _Version        =   196609
            BackColor       =   14737632
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_StlgrdDesc 
               Height          =   315
               Left            =   5310
               TabIndex        =   16
               Top             =   450
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.TextBox tmpSTLGRD 
               Height          =   315
               Left            =   3270
               TabIndex        =   13
               Top             =   450
               Visible         =   0   'False
               Width           =   1845
            End
            Begin InDate.ULabel ULabel8 
               Height          =   315
               Left            =   90
               Top             =   90
               Width           =   825
               _ExtentX        =   1455
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
               Left            =   945
               TabIndex        =   4
               Top             =   90
               Width           =   975
               _Version        =   262145
               _ExtentX        =   1720
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
            Begin CSTextLibCtl.sidbEdit SDB_THK_TO 
               Height          =   315
               Left            =   2130
               TabIndex        =   5
               Top             =   90
               Width           =   975
               _Version        =   262145
               _ExtentX        =   1720
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
            Begin InDate.ULabel ULabel9 
               Height          =   315
               Left            =   3270
               Top             =   90
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   556
               Caption         =   "宽度"
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
            Begin CSTextLibCtl.sidbEdit SDB_WID 
               Height          =   315
               Left            =   4125
               TabIndex        =   7
               Top             =   90
               Width           =   975
               _Version        =   262145
               _ExtentX        =   1720
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
            Begin CSTextLibCtl.sidbEdit SDB_WID_TO 
               Height          =   315
               Left            =   5310
               TabIndex        =   8
               Top             =   90
               Width           =   975
               _Version        =   262145
               _ExtentX        =   1720
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
            Begin Threed.SSCommand ssc_cmd 
               Height          =   345
               Left            =   90
               TabIndex        =   10
               Top             =   450
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   609
               _Version        =   196609
               Caption         =   "替代查询"
            End
            Begin CSTextLibCtl.sidbEdit tmpThk 
               Height          =   315
               Left            =   2130
               TabIndex        =   11
               Top             =   450
               Visible         =   0   'False
               Width           =   975
               _Version        =   262145
               _ExtentX        =   1720
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
               RawData         =   ""
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
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               MaxValue        =   9999.99
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit tmpWid 
               Height          =   315
               Left            =   7560
               TabIndex        =   12
               Top             =   450
               Visible         =   0   'False
               Width           =   975
               _Version        =   262145
               _ExtentX        =   1720
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
               RawData         =   ""
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
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               MaxValue        =   9999.99
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin Threed.SSCommand cmd_Change 
               Height          =   675
               Left            =   12180
               TabIndex        =   14
               Top             =   90
               Width           =   1830
               _ExtentX        =   3228
               _ExtentY        =   1191
               _Version        =   196609
               Font3D          =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "替代"
            End
            Begin CSTextLibCtl.sidbEdit tmpLen 
               Height          =   315
               Left            =   8730
               TabIndex        =   15
               Top             =   450
               Visible         =   0   'False
               Width           =   975
               _Version        =   262145
               _ExtentX        =   1720
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
               RawData         =   ""
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
               NumDecDigits    =   0
               NumIntDigits    =   4
               ShowZero        =   0   'False
               MaxValue        =   9999.99
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin VB.Label Label3 
               BackColor       =   &H00E0E0E0&
               Caption         =   "~"
               Height          =   120
               Left            =   5160
               TabIndex        =   9
               Top             =   210
               Width           =   195
            End
            Begin VB.Label Label2 
               BackColor       =   &H00E0E0E0&
               Caption         =   "~"
               Height          =   120
               Left            =   1980
               TabIndex        =   6
               Top             =   210
               Width           =   195
            End
         End
         Begin FPSpread.vaSpread ss4 
            Height          =   2715
            Left            =   30
            TabIndex        =   2
            Top             =   960
            Width           =   14895
            _Version        =   393216
            _ExtentX        =   26273
            _ExtentY        =   4789
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
            MaxCols         =   17
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGB2030C.frx":1D67
         End
      End
   End
End
Attribute VB_Name = "CGB2030C"
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
'-- Program Name      加热炉作业实绩查询及修改界面
'-- Program ID        CGB2020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          SHIN.C.S
'-- Coder             SHIN.C.S
'-- Date              2007.7.23
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String            'Form Type
Public Toolbar_St As String          'Active Form ToolBar Setting
Public sAuthority As String          'Active Form Authority Setting
Public sDateTime As String           'Active Form Time Setting
Public sQuery_Rt As String

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

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

Dim pControl4 As New Collection      'Master Primary Key Collection
Dim nControl4 As New Collection      'Master Necessary Collection
Dim mControl4 As New Collection      'Master Maxlength check Collection
Dim iControl4 As New Collection      'Master Insert Collection
Dim rControl4 As New Collection      'Master Refer Collection
Dim cControl4 As New Collection      'Master Copy Collection
Dim aControl4 As New Collection      'Master -> Spread Collection
Dim lControl4 As New Collection      'Master Lock Collection

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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim Mc4 As New Collection           'Master Collection

Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim inqFL As String                 'INQUERY FLAG

Private Sub Form_Define()
     
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_PrcLine, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    'MASTER Collection
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"
     
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
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

   
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGB2030C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(tmpSTLGRD, "p", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
 Call Gp_Ms_Collection(txt_StlgrdDesc, " ", " ", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
         Call Gp_Ms_Collection(tmpThk, "p", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
         Call Gp_Ms_Collection(tmpWid, "p", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
         Call Gp_Ms_Collection(tmpLen, "p", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    
        Call Gp_Ms_Collection(SDB_THK, "p", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(SDB_THK_TO, "p", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
        Call Gp_Ms_Collection(SDB_WID, "p", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
     Call Gp_Ms_Collection(SDB_WID_TO, "p", "n", " ", " ", "r", " ", " ", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
    
    'MASTER Collection
     Mc4.Add Item:=pControl4, Key:="pControl"
     Mc4.Add Item:=nControl4, Key:="nControl"
     Mc4.Add Item:=mControl4, Key:="mControl"
     Mc4.Add Item:=iControl4, Key:="iControl"
     Mc4.Add Item:=rControl4, Key:="rControl"
     Mc4.Add Item:=cControl4, Key:="cControl"
     Mc4.Add Item:=aControl4, Key:="aControl"
     Mc4.Add Item:=lControl4, Key:="lControl"
     
   'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss4, 1, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 5, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 6, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 7, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 8, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Call Gp_Sp_Collection(ss4, 9, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 10, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 11, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 12, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 13, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 14, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 15, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 16, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
   Call Gp_Sp_Collection(ss4, 17, " ", " ", " ", " ", " ", " ", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)

   'Spread_Collection
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="CGB2030C.P_REFER4", Key:="P-R"
    sc4.Add Item:=pColumn4, Key:="pColumn"
    sc4.Add Item:=nColumn4, Key:="nColumn"
    sc4.Add Item:=aColumn4, Key:="aColumn"
    sc4.Add Item:=mColumn4, Key:="mColumn"
    sc4.Add Item:=iColumn4, Key:="iColumn"
    sc4.Add Item:=lColumn4, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 1, True)
    Call Gp_Sp_ColHidden(ss4, 1, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
     
End Sub

Private Sub cmd_Change_Click()
Dim SlabNo As String
Dim ForCnt As Long
Dim OutParam(2, 4) As Variant
Dim ret_Result_ErrMsg As String
Dim sQuery As String
Dim iType As String
Dim OrdNo As String
Dim OrdItem As String
Dim pThk As Double
Dim pWid As Double
Dim pLen As Double
Dim sndDATA As Long
Dim ProdCnt As Long

    For ForCnt = 1 To ss1.MaxRows
        ss1.ROW = ForCnt
        ss1.Col = 1
        If ss1.Text = "Y" Then
            ss1.Col = 2
            SlabNo = SlabNo & Trim(ss1.Text)
            ss1.Col = 5
            pThk = ss1.Text
            ss1.Col = 6
            pWid = ss1.Text
            ss1.Col = 7
            pLen = ss1.Text
        End If
    Next


    For ForCnt = 1 To ss4.MaxRows
        ss4.ROW = ForCnt
        ss4.Col = 1
        If ss4.Text = "Y" Then
            ss4.Col = 2
            OrdNo = ss4.Text
            ss4.Col = 3
            OrdItem = ss4.Text
            ss4.Col = 16
            ProdCnt = ss4.Text
        End If
    Next
    
   
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    
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
    
    iType = "U"
    
    sndDATA = Len(Trim(SlabNo))
    SlabNo = SlabNo & Space(200 - sndDATA)
    sQuery = "{call CGB2030C.P_MODIFY1 ('" & iType & "','" & SlabNo & "','" & OrdNo & "','" & OrdItem & "',"
    sQuery = sQuery & pThk & "," & pWid & "," & pLen & "," & ProdCnt & ",'" & sUserID & "',?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
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
        Exit Sub
        
    End If
    
    Set adoCmd = Nothing
    
    Call Form_Ref
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub opt_rhf_Click(Index As Integer)
    If Index = 0 Then
        txt_PrcLine.Text = "1"
        opt_rhf(0).ForeColor = &HFF&
        opt_rhf(1).ForeColor = &H80000011
        opt_rhf(2).ForeColor = &H80000011
        opt_rhf(3).ForeColor = &H80000011
        Call Form_Ref
    ElseIf Index = 1 Then
        txt_PrcLine.Text = "2"
        opt_rhf(1).ForeColor = &HFF&
        opt_rhf(0).ForeColor = &H80000011
        opt_rhf(2).ForeColor = &H80000011
        opt_rhf(3).ForeColor = &H80000011
        Call Form_Ref
    ElseIf Index = 2 Then
        txt_PrcLine.Text = "3"
        opt_rhf(2).ForeColor = &HFF&
        opt_rhf(0).ForeColor = &H80000011
        opt_rhf(1).ForeColor = &H80000011
        opt_rhf(3).ForeColor = &H80000011
        Call Form_Ref
    ElseIf Index = 3 Then
        txt_PrcLine.Text = "4"
        opt_rhf(3).ForeColor = &HFF&
        opt_rhf(0).ForeColor = &H80000011
        opt_rhf(1).ForeColor = &H80000011
        opt_rhf(2).ForeColor = &H80000011
        Call Form_Ref
    End If
    
    Call Gp_Ms_Cls(Mc4("rControl"))
    Call Gf_Sp_Cls(sc4)

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
Dim ForCnt As Long
Dim chkFL As String
Dim chkOrd As String

    If ROW < 1 Then Exit Sub
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    ss1.ROW = ROW
    ss1.Col = 1
    
    If Trim(ss1.Text) = "" Then
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW, "&H00000000", "&HFFFF80")
        ss1.Text = "Y"
    Else
        ss1.Text = ""
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)
    End If
    
    
    If Trim(tmpSTLGRD) = "" Then
        ss1.ROW = ROW
        ss1.Col = 3
        tmpSTLGRD = ss1.Value
        ss1.Col = 4
        txt_StlgrdDesc = ss1.Value
        ss1.Col = 5
        tmpThk = ss1.Value
        ss1.Col = 6
        tmpWid = ss1.Value
        ss1.Col = 7
        tmpLen = ss1.Value
    Else
        ss1.ROW = ROW
        ss1.Col = 3
        If ss1.Value <> Trim(tmpSTLGRD) Then
            ss1.Col = 1
            ss1.Text = ""
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)
        End If
        
        ss1.Col = 5
        If ss1.Text <> Trim(tmpThk) Then
            ss1.Col = 1
            ss1.Text = ""
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)
        End If
        
        ss1.Col = 6
        If ss1.Text <> Trim(tmpWid) Then
            ss1.Col = 1
            ss1.Text = ""
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ROW, ROW)
        End If
    
    End If

    chkFL = ""
    For ForCnt = 1 To ss1.MaxRows
        ss1.ROW = ForCnt
        ss1.Col = 1
        If Trim(ss1.Text) <> "" Then
            chkFL = "Y"
        End If
    Next
    
    If chkFL = "" Then
        tmpSTLGRD = ""
        tmpThk = ""
        tmpWid = ""
        tmpLen = ""
    End If
    
    For ForCnt = 1 To ss1.MaxRows
        ss1.ROW = ForCnt
        ss1.Col = 1
        If ss1.Text = "Y" Then
            Exit Sub
        End If
    Next
        
    Call Gp_Ms_Cls(Mc4("rControl"))
    Call Gf_Sp_Cls(sc4)
    
End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal ROW As Long)
Dim ForCnt As Long
Dim chkFL As String

    If ROW < 1 Then Exit Sub
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    chkFL = ""
    For ForCnt = 1 To ss4.MaxRows
        ss4.ROW = ForCnt
        ss4.Col = 1
        If Trim(ss4.Text) <> "" Then
            chkFL = "Y"
        End If
    Next
    
    ss4.ROW = ROW
    ss4.Col = 1
    If Trim(ss4.Text) = "" Then
        If chkFL <> "Y" Then
            Call Gp_Sp_BlockColor(ss4, 1, ss4.MaxCols, ROW, ROW, "&H00000000", "&HFFFF80")
            ss4.Text = "Y"
        End If
    Else
        ss4.Text = ""
        Call Gp_Sp_BlockColor(ss4, 1, ss4.MaxCols, ROW, ROW)
    End If
    
End Sub

Private Sub ssc_cmd_Click()
    Call Gf_Sp_Refer(M_CN1, sc4, Mc4, Mc4("nControl"), Mc4("mControl"))
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc4.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc4)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc4.Item("Spread"), "CG-System.INI", Me.Name)

    txt_PrcLine = "1"
    opt_rhf(0).Value = True
    
    If Mid(sAuthority, 3, 2) = "11" Then
       cmd_Change.Enabled = True
    Else
       cmd_Change.Enabled = False
    End If
    
    Call Form_Ref
          
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "CG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc4.Item("Spread"), "CG-System.INI", Me.Name)
    
    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
   
    Set pControl4 = Nothing
    Set nControl4 = Nothing
    Set iControl4 = Nothing
    Set rControl4 = Nothing
    Set cControl4 = Nothing
    Set aControl4 = Nothing
    Set lControl4 = Nothing
    Set mControl4 = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing

    Set Mc1 = Nothing
    Set Mc4 = Nothing
    
    Set sc1 = Nothing
    Set sc4 = Nothing

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

    Dim sMesg As String

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc4("rControl"))

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc4)

    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    
    Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    
    opt_rhf(0).Value = True
    txt_PrcLine = "1"

End Sub

Public Sub Form_Ref()
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"))
    inqFL = ""
    ss1.OperationMode = OperationModeNormal
    
     
End Sub

Public Sub Form_Pro()


'    Dim sMesg As String
'    Dim sLoc As String
'    Dim Temp_no As String
'
'    TXT_UPD_EMP.Text = sUserID
'
'    If sc1 = -1 Then
'       If Not Gp_DateCheck(TXT_RHF_CH_TIME) Then
'            sMesg = " 请正确输入装炉时间 ！"
'            Call Gp_MsgBoxDisplay(sMesg)
'            Exit Sub
'       End If
'       If TXT_RHF_CH_TIME_UPD.RawData <> "" Then
'            If Not Gp_DateCheck(TXT_RHF_CH_TIME_UPD) Then
'                 sMesg = " 请正确输入装炉时间修正 ！"
'                 Call Gp_MsgBoxDisplay(sMesg)
'                 Exit Sub
'            End If
'       End If
'       If Gf_Mc_Authority(sAuthority, Mc1) Then
'            If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
'               Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'            End If
'       End If
'    ElseIf sc2 = -1 Then
'        If Not Gp_DateCheck(TXT_DISCHARGE_TIME) Then
'            sMesg = " 请正确输入出炉时间 ！"
'            Call Gp_MsgBoxDisplay(sMesg)
'            Exit Sub
'        End If
''        If Len(TXT_RHF_CH_TIME.RawData) = 14 Then
''           If Val(TXT_RHF_CH_TIME.RawData) - Val(TXT_DISCHARGE_TIME.RawData) > 0 Then
''                sMesg = " 出炉时间应大于装炉时间，请正确输入时间信息 ！"
''                Call Gp_MsgBoxDisplay(sMesg)
''                Exit Sub
''           End If
''        Else
''            sMesg = " 请先进行装炉操作或装炉时间错误 ！"
''            Call Gp_MsgBoxDisplay(sMesg)
''            Exit Sub
''        End If
'
'        If Gf_Mc_Authority(sAuthority, Mc2) Then
'           If Gf_Ms_Process(M_CN1, Mc2, sAuthority) Then
''              Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'              Call MDIMain.FormMenuSetting(Me, FormType, "SE", "1111")
'
'           End If
'        End If
'    ElseIf sc3 = -1 Then
'
'        If Not Gp_DateCheck(TXT_REJ_OCCR_TIME) Then
'            sMesg = " 请正确输入缺号时间 ！"
'            Call Gp_MsgBoxDisplay(sMesg)
'            Exit Sub
'        End If
'
'        If Trim(TXT_REJ_LOC.Text) = "1" Then
'           sLoc = "入口"
'        Else
'           sLoc = "出口"
'        End If
'
'        sMesg = " 确定此板坯在加热炉 （ " + sLoc + " ）处缺号 ？ "
'
'        If Gp_MsgBox(sMesg, "C") = 6 Then
'            If Gf_Mc_Authority(sAuthority, Mc3) Then
'               If Gf_Ms_Process(M_CN1, Mc3, sAuthority) Then
'                  'Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
'                  Call MDIMain.FormMenuSetting(Me, FormType, "SE", "1111")
'               End If
'            End If
'        End If
'    End If
'
'    TXT_RHF_CH_TIME_UPD.RawData = ""
'    TXT_RHF_CH_NUM_REF = ""

End Sub

Public Sub Form_Del()

    If Not Gf_Ms_Del(M_CN1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub
