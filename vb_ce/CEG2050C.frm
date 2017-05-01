VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form CEG2050C 
   Caption         =   "炼钢厂均衡查询/调整_CEG2050C"
   ClientHeight    =   9225
   ClientLeft      =   285
   ClientTop       =   1965
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15300
   WindowState     =   2  'Maximized
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
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   90
      Width           =   2835
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
      Left            =   6495
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   90
      Width           =   465
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   5190
      Top             =   90
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt1 
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Top             =   480
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   5160
      Top             =   480
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "板卷炼钢厂"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_cnt1 
      Height          =   315
      Left            =   6480
      TabIndex        =   2
      Top             =   480
      Width           =   720
      _Version        =   262145
      _ExtentX        =   1270
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   16711680
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_wgt2 
      Height          =   315
      Left            =   10950
      TabIndex        =   4
      Top             =   480
      Width           =   1410
      _Version        =   262145
      _ExtentX        =   2487
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   8910
      Top             =   480
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      Caption         =   "老炼钢厂"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_cnt2 
      Height          =   315
      Left            =   10230
      TabIndex        =   5
      Top             =   480
      Width           =   720
      _Version        =   262145
      _ExtentX        =   1270
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   16711680
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
      ReadOnly        =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   12
      Undo            =   0
      Data            =   0
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8310
      Left            =   60
      TabIndex        =   6
      Top             =   870
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   14658
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "CEG2050C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   4410
         Left            =   0
         TabIndex        =   7
         Top             =   3900
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   7779
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "CEG2050C.frx":0052
         Begin Threed.SSPanel SSPanel1 
            Height          =   555
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   15210
            _ExtentX        =   26829
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_sms_plt 
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
               Left            =   10230
               MaxLength       =   2
               TabIndex        =   15
               Tag             =   "工厂"
               Top             =   120
               Visible         =   0   'False
               Width           =   465
            End
            Begin Threed.SSOption opt_plt_no1 
               Height          =   285
               Left            =   1830
               TabIndex        =   12
               Top             =   150
               Width           =   1245
               _ExtentX        =   2196
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
               Caption         =   "板卷炼钢厂"
               Value           =   -1
            End
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
               Height          =   310
               Left            =   13500
               MaxLength       =   11
               TabIndex        =   10
               Top             =   120
               Visible         =   0   'False
               Width           =   1275
            End
            Begin CSTextLibCtl.sidbEdit sdb_slab_thk 
               Height          =   315
               Left            =   10710
               TabIndex        =   11
               Top             =   120
               Visible         =   0   'False
               Width           =   1365
               _Version        =   262145
               _ExtentX        =   2408
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               Undo            =   0
               Data            =   0
            End
            Begin InDate.ULabel ULabel6 
               Height          =   315
               Left            =   270
               Top             =   120
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               Caption         =   "变更炼钢厂"
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
               ForeColor       =   0
            End
            Begin Threed.SSOption opt_plt_no2 
               Height          =   285
               Left            =   3210
               TabIndex        =   13
               Top             =   150
               Width           =   1515
               _ExtentX        =   2672
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
               Caption         =   "老炼钢厂"
            End
            Begin CSTextLibCtl.sidbEdit sdb_slab_wid 
               Height          =   315
               Left            =   12090
               TabIndex        =   14
               Top             =   120
               Visible         =   0   'False
               Width           =   1365
               _Version        =   262145
               _ExtentX        =   2408
               _ExtentY        =   556
               _StockProps     =   125
               Text            =   " 0.00"
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
               Undo            =   0
               Data            =   0
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   3825
            Left            =   0
            TabIndex        =   9
            Top             =   585
            Width           =   15210
            _Version        =   393216
            _ExtentX        =   26829
            _ExtentY        =   6747
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
            MaxCols         =   32
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CEG2050C.frx":00A4
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3840
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   15210
         _Version        =   393216
         _ExtentX        =   26829
         _ExtentY        =   6773
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
         MaxCols         =   0
         MaxRows         =   20
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         SpreadDesigner  =   "CEG2050C.frx":1265
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   225
      Left            =   14280
      Top             =   570
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   397
      Caption         =   "重量/炉次数"
      Alignment       =   1
      BackColor       =   14737632
      BackgroundStyle =   1
      BorderEffect    =   0
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel3 
      Height          =   705
      Left            =   150
      Top             =   90
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   1244
      Caption         =   ""
      Alignment       =   1
      BackColor       =   12640511
      BackgroundStyle =   1
      BorderEffect    =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin Threed.SSCommand cmd_slab_design_change 
      Height          =   405
      Left            =   12570
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   420
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "设计标准变更"
      BevelWidth      =   3
   End
End
Attribute VB_Name = "CEG2050C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       DAILY SCHEDULE
'-- Sub_System Name
'-- Program Name      SMS CHANGE
'-- Program ID        CEG2050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2007.10.24
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
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Active_Row As Integer           'Active Row
Dim Active_Col As Integer           'Active Col

Private Sub Form_Define()

    Dim iCol As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_cnt1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_wgt1, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_cnt2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_slab_wgt2, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
     Call Gp_Ms_Collection(txt_sms_plt, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_slab_thk, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(sdb_slab_wid, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
         
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, False)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, False)
    
    For iCol = 4 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CEG2050C.P_REFER2", Key:="P-R"
    sc2.Add Item:="CEG2050C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:="CEG2050C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, SpreadHeader + (ss1.RowHeaderCols - 1), True)

End Sub

Private Sub Sp_Setting()

    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 3)) = 16
    ss1.ColWidth(SpreadHeader + (ss1.RowHeaderCols - 2)) = 5
    ss1.MaxCols = 0

End Sub

Private Sub cmd_slab_design_change_Click()
    
    If ss1.MaxCols - 1 = Active_Col Or ss1.MaxCols - 2 = Active_Col Or _
        ss1.MaxCols - 3 = Active_Col Or ss1.MaxCols = Active_Col Then Exit Sub
    
    If Active_Col Mod 3 <> 2 And Active_Col Mod 3 <> 1 Then Exit Sub
    
    Load Slab_Design_Change

    If Active_Row > 0 And Active_Col > 0 Then
    
        ss1.Row = 0
        ss1.Col = Active_Col
        Slab_Design_Change.sdb_slab_thk_fr.Value = ss1.Text
        Slab_Design_Change.sdb_slab_thk_to.Value = ss1.Text
    
        ss1.Row = Active_Row
        ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 2)
        Slab_Design_Change.sdb_slab_wid_fr.Value = ss1.Value
        Slab_Design_Change.sdb_slab_wid_to.Value = ss1.Value
    
        ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 1)
        Slab_Design_Change.txt_stlgrd.Text = ss1.Text
        ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 3)
        Slab_Design_Change.txt_stlgrd_name.Text = ss1.Text
    
        Slab_Design_Change.sdb_slab_len_fr.Value = 0
        Slab_Design_Change.sdb_slab_len_to.Value = 9999999
        
        Slab_Design_Change.sdb_asroll_cnt_fr.Value = 0
        Slab_Design_Change.sdb_asroll_cnt_to.Value = 99
        
        Slab_Design_Change.sdb_prod_thk_fr.Value = 0
        Slab_Design_Change.sdb_prod_thk_to.Value = 9999.99
        
        Slab_Design_Change.sdb_prod_wid_fr.Value = 0
        Slab_Design_Change.sdb_prod_wid_to.Value = 9999.99
        
        Slab_Design_Change.sdb_prod_len_fr.Value = 0
        Slab_Design_Change.sdb_prod_len_to.Value = 9999999.9
        
        Slab_Design_Change.cbo_prod_cnt.Text = "0"
        
    End If
    
    Slab_Design_Change.P_MODE = "W"
    
    Slab_Design_Change.Show 1

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

    Dim sStatus As String
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    If Mid(sAuthority, 3, 1) <> "1" Then
        cmd_slab_design_change.Enabled = False
    End If
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Sp_Setting
    
    sStatus = Gf_CodeFind(M_CN1, "SELECT PRC_STS FROM ZP_JOB WHERE PLT = 'C3' AND MAIN_PGMID = 'CEG2000P'")
    If sStatus <> "C" Then
        ULabel3.Caption = "坯料使用计划进行中...!!"
    Else
        sStatus = Gf_CodeFind(M_CN1, "SELECT PRC_STS FROM ZP_JOB WHERE PLT = 'C3' AND MAIN_PGMID = 'CEH1000P'")
        If sStatus <> "C" Then
            ULabel3.Caption = "坯料分段作业指示进行中...!!"
        Else
            sStatus = Gf_CodeFind(M_CN1, "SELECT PRC_STS FROM ZP_JOB WHERE PLT = 'C3' AND MAIN_PGMID = 'CED1000P'")
            If sStatus <> "C" Then
                ULabel3.Caption = "轧钢工序计划进行中...!!"
            Else
                ULabel3.Caption = "没有可作业对象...!!"
                sAuthority = "0000"
            End If
        End If
    End If
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet

    txt_plt.Text = "C3"
    Call txt_plt_KeyUp(0, 0)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "E-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "E-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Spl_SizeSet(SSSplitter1, "E-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "E-System.INI", Me.Name)
    
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

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
        Call Gf_Sp_Cls(sc2)
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        txt_plt.Text = "C3"
        Call txt_plt_KeyUp(0, 0)
        opt_plt_no1.Value = True
        ss1.MaxCols = 0
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Ref()

    Dim sQuery1 As String   'Header Display
    Dim sQuery2 As String   'Data Display
    Dim sQuery3 As String   'STLGRD SUM Display
    Dim sQuery4 As String   'WID, THK SUM Display
    Dim sQuery5 As String   'TOTAL SUM Display
    Dim SMESG As String
    
    'Header Display
    sQuery1 = "SELECT  DISTINCT  SLAB_THK "
    sQuery1 = sQuery1 + "  FROM  EP_SLAB_EDT3 "
    sQuery1 = sQuery1 + " WHERE  SLAB_EDT_FL  IN ('N','H') "
    sQuery1 = sQuery1 + " ORDER  BY SLAB_THK ASC "
    
    '炉次编制量标准 B1:150 B3:30
    
    'Data Display
    sQuery2 = " {call CEG2050C.P_DATA ()} "
    
    'STLGRD, WID SUM Display
    sQuery3 = " {call CEG2050C.P_STLGRD_WID ()} "
    
    'THK SUM Display
    sQuery4 = " {call CEG2050C.P_THK ()} "
    
    'SUM Display
    sQuery5 = " {call CEG2050C.P_TOTAL ()} "
    
    SMESG = Gf_Ms_NeceCheck(nControl)
    If SMESG = "OK" Then
    
        SMESG = Gf_Ms_NeceCheck2(mControl)
        If SMESG = "OK" Then

            'Header Display
            Call Sp_Header_Refer1(ss1, sQuery1)      'Header Display
        
            'Data Display
            If Sp_Data_Refer1(ss1, sQuery2) Then     'SLAB Data Display
                Call Sp_Data_Refer2(ss1, sQuery3)    'STLGRD SUM Display
                Call Sp_Data_Refer3(ss1, sQuery4)    'WID, THK SUM Display
                Call Sp_Data_Refer4(ss1, sQuery5)    'TOTAL SUM Display
                ss1.OperationMode = OperationModeNormal
                opt_plt_no1.Value = True
                Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
                sdb_slab_cnt1.Value = Gf_FloatFind(M_CN1, "SELECT COUNT(SLAB_EDT_SEQ) FROM EP_SLAB_EDT3 WHERE SMS_PLT = 'B1' AND SLAB_EDT_FL  IN ('N','H') ")
                sdb_slab_cnt2.Value = Gf_FloatFind(M_CN1, "SELECT COUNT(SLAB_EDT_SEQ) FROM EP_SLAB_EDT3 WHERE SMS_PLT = 'B3' AND SLAB_EDT_FL  IN ('N','H') ")
                sdb_slab_wgt1.Value = Gf_FloatFind(M_CN1, "SELECT SUM(NVL(SLAB_WGT,0)) FROM EP_SLAB_EDT3 WHERE SMS_PLT = 'B1' AND SLAB_EDT_FL  IN ('N','H') ")
                sdb_slab_wgt2.Value = Gf_FloatFind(M_CN1, "SELECT SUM(NVL(SLAB_WGT,0)) FROM EP_SLAB_EDT3 WHERE SMS_PLT = 'B3' AND SLAB_EDT_FL  IN ('N','H') ")
            End If
            
        Else
            Call Gp_MsgBoxDisplay(Trim(SMESG) + "长度不正确", "I")
        End If
    
    Else
        Call Gp_MsgBoxDisplay(Trim(SMESG) + "必须输入", "I")
    End If

End Sub

Public Sub Form_Pro()
    
    If Gf_Sp_Process(M_CN1, sc2, Mc2) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        Call Form_Ref
    End If
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, sc2)
      
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

Private Sub opt_plt_no1_Click(Value As Integer)

    If opt_plt_no1.Value Then
        opt_plt_no1.ForeColor = &HFF&
        opt_plt_no2.ForeColor = &H80000012
    Else
        opt_plt_no1.ForeColor = &H80000012
    End If
    
End Sub

Private Sub opt_plt_no2_Click(Value As Integer)
    
    If opt_plt_no2.Value Then
        opt_plt_no2.ForeColor = &HFF&
        opt_plt_no1.ForeColor = &H80000012
    Else
        opt_plt_no2.ForeColor = &H80000012
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

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If Row < 1 Or ss1.MaxRows = Row Then
        Active_Row = 0
        Active_Col = 0
        Exit Sub
    End If
    
    If ss1.MaxCols - 1 = Col Or ss1.MaxCols - 2 = Col Or ss1.MaxCols = Col Then
        Active_Row = 0
        Active_Col = 0
        Exit Sub
    End If
    
    If Col Mod 3 = 0 Then
        Active_Row = 0
        Active_Col = 0
        Exit Sub
    End If
    
    If Col Mod 3 = 1 Then
        txt_sms_plt.Text = "B1"
        opt_plt_no2.Value = True
    Else
        txt_sms_plt.Text = "B3"
        opt_plt_no1.Value = True
    End If
    
    Active_Row = Row
    Active_Col = Col
    
    ss1.Col = Col
    ss1.Row = SpreadHeader + (ss1.ColHeaderRows - 2)
    sdb_slab_thk.Value = Val(ss1.Text)
    
    ss1.Row = Row
    ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 1)
    txt_stlgrd.Text = ss1.Text
    
    ss1.Col = SpreadHeader + (ss1.RowHeaderCols - 2)
    sdb_slab_wid.Value = Val(ss1.Text)
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call MenuTool_ReSet
    ss2.OperationMode = OperationModeNormal

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If Row < 1 Then Exit Sub
    
    ss2.Col = 0
    ss2.Row = Row
    
    If ss2.Text = "" Then
        ss2.Col = 2
'        If ss2.Text <> "*" Then
'            Call Gp_MsgBoxDisplay("不能变更炼钢厂", "I", Me.Caption)
'            Exit Sub
'        End If
        
        ss2.Col = 0:    ss2.Text = "Update"
        ss2.Col = 3
        If opt_plt_no1.Value Then
            ss2.Text = "B1"
        Else
            ss2.Text = "B3"
        End If
        
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row, , &HFFFF80)
    Else
        ss2.Col = 0:    ss2.Text = ""
        ss2.Col = 3:    ss2.Text = ""
        Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row)
    End If

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
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
    End If
    
End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If Row > 0 Then
        Set Active_Spread = Me.ss2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = True
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
        DD.rControl.Add Item:=txt_plt_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_name.Text = ""
    End If

End Sub

Private Function Sp_Header_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Header_Refer1 = True
        
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer1 = False
            '.ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 3
            For iCol = 0 To .MaxCols - 1 Step 3
            
                For iColCnt = 1 To 3
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 2)
                    .Col = iCol + iColCnt
                    
                    If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(0, iCnt))
                    End If
                    
                    .ColWidth(iCol + iColCnt) = 8
    
                    .Col = iCol + iColCnt: .Col2 = iCol + iColCnt
                    .Row = 1: .Row2 = -1
                    .BlockMode = True
                    .TypeHAlign = TypeHAlignCenter
                    .TypeVAlign = TypeVAlignCenter
                    .BlockMode = False
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 1)
                    .Col = iCol + iColCnt
                    
                    Select Case iColCnt
                        Case 1
                            .Text = "板卷厂"
                        Case 2
                            .Text = "老炼厂"
                        Case 3
                            .Text = "合计"
                    End Select
                    
                    If iColCnt = 3 Then
                        Call Gp_Sp_ColHidden(ss1, .Col, True)
                    End If
                    
                Next iColCnt
                
                iCnt = iCnt + 1
                
            Next iCol
            
            '合计 Col
            For iColCnt = 1 To 3
                
                .MaxCols = .MaxCols + 1
                .Col = .MaxCols
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                .Text = "合计(t)"
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                
                Select Case iColCnt
                    Case 1
                        .Text = "板卷厂"
                    Case 2
                        .Text = "老炼厂"
                    Case 3
                        .Text = "合计"
                End Select
                    
                .ColWidth(.Col) = 12
                    
                .Col = .MaxCols: .Col2 = .MaxCols
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .TypeHAlign = TypeHAlignCenter
                .TypeVAlign = TypeVAlignCenter
                .BlockMode = False
                
            Next iColCnt
            
        End If
        
        .BlockMode = True
        .Col = .MaxCols:  .Col2 = .MaxCols
        .Row = 1: .Row2 = -1
        .ForeColor = &HFF&  '&H00FF0000&
        .BlockMode = False
        
        For iColCnt = 3 To .MaxCols - 3 Step 3
            .BlockMode = True
            .Col = iColCnt:  .Col2 = iColCnt
            .Row = 1: .Row2 = -1
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iColCnt
        
        .BlockMode = True
        .Row = SpreadHeader + (.ColHeaderRows - 2)
        .Col = 1
        .Row2 = SpreadHeader + (.ColHeaderRows - 2)
        .Col2 = .MaxCols - 3
        .RowMerge = MergeAlways
        '.ColMerge = MergeAlways
        .BlockMode = False
        
        .BlockMode = True
        .Row = SpreadHeader + (.ColHeaderRows - 2)
        .Col = .MaxCols - 2
        .Row2 = SpreadHeader + (.ColHeaderRows - 1)
        .Col2 = .MaxCols - 2
        .RowMerge = MergeAlways
        ''.ColMerge = MergeAlways
        .BlockMode = False
        
        .ReDraw = True
        .Refresh
        
        Screen.MousePointer = vbDefault
        
    End With
        
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Private Function Sp_Data_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim iBas As Integer
    Dim iCot As Integer
    
    Dim sCol_a As String
    Dim sCol_b As String
    Dim sStlgrd As String
    Dim sWid As String
    
    Dim ColSum As Double
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Data_Refer1 = True
        .ReDraw = False
        .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer1 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)

                If iCnt = 0 Or sStlgrd <> Trim(ArrayRecords(2, iCnt)) Or sWid <> Trim(ArrayRecords(1, iCnt)) Then
                    sStlgrd = Trim(ArrayRecords(2, iCnt))
                    sWid = Trim(ArrayRecords(1, iCnt))
                    .MaxRows = .MaxRows + 1
                    .Row = .MaxRows
                    .Col = SpreadHeader + (.RowHeaderCols - 3)
                    .Text = Trim(ArrayRecords(0, iCnt))
                    .Col = SpreadHeader + (.RowHeaderCols - 2)
                    .Text = Trim(ArrayRecords(1, iCnt))
                    .Col = SpreadHeader + (.RowHeaderCols - 1)
                    .Text = Trim(ArrayRecords(2, iCnt))
                End If
                
                For iCol = 1 To .MaxCols - 1 Step 3
                
                    .Col = iCol
                    .Row = SpreadHeader + (.ColHeaderRows - 2)
                    
                    If .Text = Trim(ArrayRecords(3, iCnt)) Then

                        .Row = .MaxRows
                        
                        If VarType(ArrayRecords(4, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(4, iCnt)) = "0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(4, iCnt))
                            End If
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(5, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(5, iCnt)) = "0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(5, iCnt))
                            End If
                        End If
                        
                        .Col = iCol + 2
                        If VarType(ArrayRecords(6, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(6, iCnt))
                        End If
                    End If
                        
                Next iCol
                
            Next iCnt
            
        End If
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 0
        .Text = "合计(t)"
        
        Call Gp_Sp_EvenRowBackcolor(sPname, 1)
        
        .BlockMode = True
        .Row = .MaxRows:  .Row2 = .MaxRows
        .Col = 1: .Col2 = -1
        .ForeColor = &HFF&
        .BlockMode = False
        
        For iCol = 3 To .MaxCols - 3 Step 3
            .BlockMode = True
            .Col = iCol:  .Col2 = iCol
            .Row = .MaxRows: .Row2 = .MaxRows
            .ForeColor = &HFF0000
            .BlockMode = False
        Next iCol
        
        .ReDraw = True
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Private Function Sp_Data_Refer2(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay2_Error

    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim sStlgrd As String
    Dim sWid As String
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Data_Refer2 = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer2 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)
                
                For iRow = 1 To .MaxRows
                    
                    .Row = iRow
                    .Col = SpreadHeader + (.RowHeaderCols - 1)
                    sStlgrd = .Text
                    .Col = SpreadHeader + (.RowHeaderCols - 2)
                    sWid = .Text
                    
                    If sStlgrd = Trim(ArrayRecords(1, iCnt)) And sWid = Trim(ArrayRecords(2, iCnt)) Then
    
                        .Col = .MaxCols - 2
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(3, iCnt)) = "0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(3, iCnt))
                            End If
                        End If
                        
                        .Col = .MaxCols - 1
                        If VarType(ArrayRecords(4, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(4, iCnt)) = "0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(4, iCnt))
                            End If
                        End If
                        
                        .Col = .MaxCols
                        If VarType(ArrayRecords(5, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(5, iCnt)) = "0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(5, iCnt))
                            End If
                        End If
                        
                        Exit For
                        
                    End If
                    
                Next iRow

            Next iCnt
                
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay2_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer2 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay2_Error : " & Error)
    
End Function

Private Function Sp_Data_Refer3(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay3_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Data_Refer3 = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer3 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            For iCnt = 0 To UBound(ArrayRecords, 2)

                For iCol = 1 To .MaxCols - 1 Step 3
                
                    .Col = iCol
                    .Row = SpreadHeader + (.ColHeaderRows - 2)
                    
                    If .Text = Trim(ArrayRecords(0, iCnt)) Then
                        .Row = .MaxRows
                        
                        .Col = iCol
                        If VarType(ArrayRecords(1, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(1, iCnt)) = "0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(1, iCnt))
                            End If
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(2, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(2, iCnt)) = "0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(2, iCnt))
                            End If
                        End If
                        
                        .Col = iCol + 2
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Trim(ArrayRecords(3, iCnt)) = "0/0" Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(3, iCnt))
                            End If
                        End If
                        
                        Exit For
                        
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
        .ReDraw = True
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay3_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer3 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay3_Error : " & Error)
    
End Function

Private Function Sp_Data_Refer4(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay4_Error

    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Data_Refer4 = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer4 = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
                            
            .Row = .MaxRows
            
            .Col = .MaxCols - 2
            If VarType(ArrayRecords(0, 0)) = vbNull Then
                .Text = ""
            Else
                If Trim(ArrayRecords(0, 0)) = "0/0" Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, 0))
                End If
            End If
            
            .Col = .MaxCols - 1
            If VarType(ArrayRecords(1, 0)) = vbNull Then
                .Text = ""
            Else
                If Trim(ArrayRecords(1, 0)) = "0/0" Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, 0))
                End If
            End If
            
            .Col = .MaxCols
            If VarType(ArrayRecords(2, 0)) = vbNull Then
                .Text = ""
            Else
                If Trim(ArrayRecords(2, 0)) = "0/0" Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(2, 0))
                End If
            End If
            
        End If
        
        .ReDraw = True
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay4_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer4 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay3_Error : " & Error)
    
End Function

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub
