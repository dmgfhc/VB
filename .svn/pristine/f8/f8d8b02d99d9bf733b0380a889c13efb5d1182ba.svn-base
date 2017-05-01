VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB4200C 
   Caption         =   "外销板坯判定处理_ACB4200C"
   ClientHeight    =   9225
   ClientLeft      =   165
   ClientTop       =   2595
   ClientWidth     =   15330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15330
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9165
      Left            =   60
      TabIndex        =   14
      Top             =   30
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   16166
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "ACB4200C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter3 
         Height          =   7485
         Left            =   0
         TabIndex        =   15
         Top             =   1680
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   13203
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   12632319
         PaneTree        =   "ACB4200C.frx":0052
         Begin FPSpread.vaSpread ss1 
            Height          =   7485
            Left            =   0
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   15225
            _Version        =   393216
            _ExtentX        =   26855
            _ExtentY        =   13203
            _StockProps     =   64
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
            MaxCols         =   31
            MaxRows         =   2
            SpreadDesigner  =   "ACB4200C.frx":0084
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   1650
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   2910
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "ACB4200C.frx":0FE0
         Begin Threed.SSPanel SSPanel1 
            Height          =   975
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   1720
            _Version        =   196609
            BackColor       =   14737632
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_stlgrd 
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
               Left            =   6225
               MaxLength       =   11
               TabIndex        =   1
               Tag             =   "钢种"
               Top             =   120
               Width           =   1260
            End
            Begin VB.TextBox txt_stlgrd_name 
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
               Left            =   7500
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   2
               Tag             =   "钢种"
               Top             =   120
               Width           =   2385
            End
            Begin VB.ComboBox cbo_ord_item 
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
               Left            =   13605
               TabIndex        =   4
               Tag             =   "订单号"
               Top             =   150
               Width           =   750
            End
            Begin VB.TextBox txt_ord_no 
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
               Left            =   12195
               MaxLength       =   11
               TabIndex        =   3
               Tag             =   "订单号"
               Top             =   150
               Width           =   1380
            End
            Begin VB.TextBox txt_change_no 
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
               Left            =   1590
               MaxLength       =   8
               TabIndex        =   0
               Tag             =   "炉号"
               Top             =   120
               Width           =   1260
            End
            Begin InDate.ULabel ULabel3 
               Height          =   315
               Left            =   180
               Top             =   120
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "炉号"
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
               Left            =   180
               Top             =   510
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "生产日期"
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
            Begin InDate.UDate dpt_prod_date_fr 
               Height          =   315
               Left            =   1590
               TabIndex        =   5
               Tag             =   "生产日期"
               Top             =   510
               Width           =   1410
               _ExtentX        =   2487
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
            Begin InDate.UDate dpt_prod_date_to 
               Height          =   315
               Left            =   3180
               TabIndex        =   6
               Tag             =   "生产日期"
               Top             =   510
               Width           =   1410
               _ExtentX        =   2487
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
            Begin InDate.ULabel lbl_slab_size 
               Height          =   315
               Left            =   4830
               Top             =   510
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "板坯尺寸"
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
            Begin CSTextLibCtl.sidbEdit sdb_thk 
               Height          =   315
               Left            =   6225
               TabIndex        =   7
               Top             =   510
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   -2147483640
               BackColor       =   16777215
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
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   "0.00"
               Text            =   " 0.00"
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
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit sdb_wid 
               Height          =   315
               Left            =   7785
               TabIndex        =   8
               Top             =   510
               Width           =   1185
               _Version        =   262145
               _ExtentX        =   2090
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   -2147483640
               BackColor       =   16777215
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
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
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
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin CSTextLibCtl.sidbEdit sdb_len 
               Height          =   315
               Left            =   9345
               TabIndex        =   9
               Top             =   510
               Width           =   1260
               _Version        =   262145
               _ExtentX        =   2222
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
               ForeColor       =   -2147483640
               BackColor       =   16777215
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
               Modified        =   0   'False
               HideSelection   =   -1  'True
               RawData         =   ""
               Text            =   " 0"
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
               NumIntDigits    =   7
               MinValue        =   0
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Left            =   7425
               Top             =   510
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   "X"
               Alignment       =   1
               BackColor       =   14737632
               BackgroundStyle =   1
               BorderEffect    =   0
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
            Begin InDate.ULabel ULabel2 
               Height          =   315
               Left            =   8985
               Top             =   510
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   556
               Caption         =   "X"
               Alignment       =   1
               BackColor       =   14737632
               BackgroundStyle =   1
               BorderEffect    =   0
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
            Begin InDate.ULabel ULabel6 
               Height          =   315
               Left            =   10800
               Top             =   150
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "订单号"
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
            Begin InDate.ULabel ULabel5 
               Height          =   315
               Left            =   4830
               Top             =   120
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "钢种"
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "~"
               Height          =   120
               Left            =   3045
               TabIndex        =   20
               Top             =   570
               Width           =   90
            End
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   645
            Left            =   0
            TabIndex        =   18
            Top             =   1005
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   1138
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_woo_rsn_name 
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
               Left            =   4890
               MaxLength       =   60
               TabIndex        =   13
               Tag             =   "余材原因"
               Top             =   150
               Width           =   3135
            End
            Begin VB.TextBox txt_woo_rsn 
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
               Left            =   4320
               MaxLength       =   2
               TabIndex        =   12
               Tag             =   "余材原因"
               Top             =   150
               Width           =   555
            End
            Begin InDate.ULabel ULabel77 
               Height          =   315
               Left            =   2910
               Top             =   150
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "余材原因"
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
            Begin Threed.SSOption opt_ok 
               Height          =   285
               Left            =   330
               TabIndex        =   10
               Top             =   180
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "外销板坯"
               Value           =   -1
            End
            Begin Threed.SSOption opt_notok 
               Height          =   285
               Left            =   1590
               TabIndex        =   11
               Top             =   180
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "内需板坯"
            End
            Begin Threed.SSCommand cmd_all 
               Height          =   405
               Left            =   13860
               TabIndex        =   21
               Top             =   120
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   714
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   16744576
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "批次适用"
               BevelWidth      =   3
            End
         End
      End
   End
End
Attribute VB_Name = "ACB4200C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   PROCESS MANAGEMENT
'-- Program Name      SLAB DECISION PROCESS
'-- Program ID        ACB4200C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             kim.sung.ho
'-- Date              2009.8.28
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
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_change_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_STLGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STLGRD_Name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dpt_prod_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dpt_prod_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_thk, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_wid, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(sdb_len, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(opt_ok, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(opt_notok, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_woo_rsn, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_woo_rsn_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
     For i = 4 To ss1.MaxCols - 2
     
        Call Gp_Sp_Collection(ss1, i, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
     Next i
     
     Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB4200C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="ACB4200C.P_REFER", Key:="P-R"
    sc1.Add Item:="ACB4200C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 30, True)
    Call Gp_Sp_ColHidden(ss1, 31, True)

End Sub

Private Sub cmd_all_Click()

    Dim i As Integer
    
    If opt_notok.Value = True Then
    
        If txt_woo_rsn.Text = "" Or txt_woo_rsn_name.Text = "" Then
            Call MsgBox("处理等级必须输入!!", vbInformation, "操作提示")
            Exit Sub
        End If
        
    End If
    
    For i = 1 To ss1.MaxRows
    
        ss1.Row = i
        ss1.Col = 0
        
        If ss1.Text = "" Then
        
            ss1.Text = "Update"
            
            If opt_ok.Value = True Then
            
                ss1.Col = 2
                ss1.Value = True
                
            ElseIf opt_notok.Value = True Then
            
                ss1.Col = 2
                ss1.Value = False
                ss1.Col = 3
                ss1.Text = txt_woo_rsn.Text
                ss1.Col = 4
                ss1.Text = txt_woo_rsn_name.Text
            
            End If
            
        
        End If
        
    Next i
    
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet
    
End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gf_Sp_Cls(sc1)
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(sc1.Item("Spread"), 3)
    
    dpt_prod_date_fr.RawData = ""
    dpt_prod_date_to.RawData = ""
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
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

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        dpt_prod_date_fr.RawData = ""
        dpt_prod_date_to.RawData = ""
        opt_ok.ForeColor = &HFF&
        opt_notok.ForeColor = &H80000012
    End If
    
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ins()
    
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

Public Sub Spread_Can()

     Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
     
End Sub

Public Sub Form_Ref()

    If txt_change_no.Text = "" And TXT_STLGRD.Text = "" And sdb_thk.Text = 0 And sdb_wid.Text = 0 And sdb_len.Text = 0 And dpt_prod_date_fr.RawData = "" And dpt_prod_date_to.RawData = "" Then
        Call MsgBox("炉号、钢种或板坯规格不能全空，请输入要查询板坯的条件之一！", vbInformation, "操作提示")
        Exit Sub
    End If

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
        
End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
    End If

End Sub

Public Sub Form_Del()

End Sub

Private Sub opt_notok_Click(Value As Integer)

    If opt_notok.Value Then
        opt_notok.ForeColor = &HFF&
        opt_ok.ForeColor = &H80000012
    End If

End Sub

Private Sub opt_ok_Click(Value As Integer)

    If opt_ok.Value Then
        opt_ok.ForeColor = &HFF&
        opt_notok.ForeColor = &H80000012
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

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 30)
        
        ss1.Row = Row
        ss1.Col = 2
        
        If ss1.Value = True Then
        
            ss1.Col = 3
            
            If ss1.Text <> "" Then
            
                ss1.Text = ""
                ss1.Col = 4
                ss1.Text = ""
            
            End If
            
        End If
        
    End If
    
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol
    
        Case 3
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "C0008"
                DD.rControl.Add Item:=3
                DD.rControl.Add Item:=4
                
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            Else
            
                ss1.Col = ss1.ActiveCol
                
                If Len(Trim(ss1.Text)) = ss1.TypeMaxEditLen Then
                    sTemp_Code = ss1.Text
                    ss1.Col = 4
                    ss1.Text = Gf_ComnNameFind(M_CN1, "C0008", Trim(sTemp_Code), 2)
                Else
                    ss1.Col = 4
                    ss1.Text = ""
                End If
            
            End If
            
    End Select
    
    ss1.Col = 3
    
    If ss1.Text <> "" Then
    
        ss1.Col = 2
        
        If ss1.Value = True Then
        
            ss1.Col = 2
            ss1.Value = False
            
        End If
        
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 30)
        
    End If
    
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(TXT_ORD_NO.Text)) = TXT_ORD_NO.MaxLength Then
    
        If CBO_ORD_ITEM.Text <> "" Then Exit Sub
        
        TXT_ORD_NO.Text = StrConv(TXT_ORD_NO.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(TXT_ORD_NO.Text) & "'"
        Call Gf_ComboAdd(M_CN1, CBO_ORD_ITEM, sQuery)

    Else
        CBO_ORD_ITEM.Clear
    End If

End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STLGRD
        DD.rControl.Add Item:=txt_STLGRD_Name
        
        DD.nameType = "2"
        
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    Else
        
        If Len(Trim(TXT_STLGRD)) >= 10 Then
            txt_STLGRD_Name.Text = Gf_CodeFind(M_CN1, "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(TXT_STLGRD.Text) + "'")
        Else
            txt_STLGRD_Name.Text = ""
        End If
    
    End If
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False
        .Buttons(8).Enabled = False
        .Buttons(11).Enabled = False
        .Buttons(12).Enabled = False
    End With
    
End Sub

Private Sub txt_woo_rsn_DblClick()

    Call txt_woo_rsn_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_woo_rsn_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0008"
        DD.rControl.Add Item:=txt_woo_rsn
        DD.rControl.Add Item:=txt_woo_rsn_name
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_woo_rsn)) = txt_woo_rsn.MaxLength Then
        txt_woo_rsn_name.Text = Gf_ComnNameFind(M_CN1, "C0008", Trim(txt_woo_rsn.Text), 2)
    Else
        txt_woo_rsn_name.Text = ""
    End If

End Sub
