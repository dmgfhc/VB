VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form DGA1140C 
   Caption         =   "板卷火剪切实绩查询及修改界面_DGA1140C"
   ClientHeight    =   6330
   ClientLeft      =   180
   ClientTop       =   3660
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   13350
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9105
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   16060
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      PaneTree        =   "DGA1140C.frx":0000
      Begin Threed.SSFrame SSFrame1 
         Height          =   1020
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15210
         _ExtentX        =   26829
         _ExtentY        =   1799
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.ComboBox cbo_plt 
            Height          =   300
            ItemData        =   "DGA1140C.frx":0052
            Left            =   13410
            List            =   "DGA1140C.frx":005F
            TabIndex        =   15
            Top             =   540
            Width           =   795
         End
         Begin VB.TextBox txt_lot_no 
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
            Left            =   10200
            TabIndex        =   14
            Top             =   540
            Width           =   1365
         End
         Begin VB.TextBox txt_prc_line 
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
            Left            =   30
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.ComboBox cbo_group 
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
            ItemData        =   "DGA1140C.frx":006F
            Left            =   13410
            List            =   "DGA1140C.frx":0082
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   150
            Width           =   795
         End
         Begin VB.ComboBox cbo_shift 
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
            ItemData        =   "DGA1140C.frx":0094
            Left            =   10200
            List            =   "DGA1140C.frx":00A4
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   150
            Width           =   795
         End
         Begin VB.TextBox txt_mat_no 
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
            Left            =   5535
            TabIndex        =   2
            Top             =   540
            Width           =   1770
         End
         Begin InDate.ULabel ULabel5 
            Height          =   315
            Left            =   3975
            Top             =   150
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            Caption         =   "作业日期"
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
         Begin InDate.UDate sdt_wrk_date_fr 
            Height          =   315
            Left            =   5535
            TabIndex        =   6
            Tag             =   "作业日期"
            Top             =   150
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
            MaxLength       =   10
         End
         Begin InDate.UDate sdt_wrk_date_to 
            Height          =   315
            Left            =   7290
            TabIndex        =   7
            Tag             =   "作业日期"
            Top             =   150
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
            MaxLength       =   10
         End
         Begin InDate.ULabel ULabel34 
            Height          =   315
            Left            =   9090
            Top             =   150
            Width           =   1080
            _ExtentX        =   1905
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
         Begin InDate.ULabel ULabel16 
            Height          =   315
            Left            =   3975
            Top             =   540
            Width           =   1530
            _ExtentX        =   2699
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
         Begin Threed.SSOption opt_htn_plt1 
            Height          =   285
            Left            =   270
            TabIndex        =   8
            Top             =   150
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   255
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "1号热处理线"
            Value           =   -1
         End
         Begin Threed.SSOption opt_htn_plt2 
            Height          =   285
            Left            =   2040
            TabIndex        =   9
            Top             =   120
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   503
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "2号热处理线"
         End
         Begin Threed.SSOption opt_htn_plt3 
            Height          =   285
            Left            =   240
            TabIndex        =   12
            Top             =   540
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   503
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "3号修建部切割"
         End
         Begin Threed.SSOption opt_htn_plt4 
            Height          =   405
            Left            =   2040
            TabIndex        =   13
            Top             =   480
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   714
            _Version        =   196609
            Font3D          =   1
            ForeColor       =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "4700热处理线"
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Left            =   9090
            Top             =   540
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            Caption         =   "垛位号"
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
         Begin InDate.ULabel ULabel35 
            Height          =   315
            Left            =   12285
            Top             =   150
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            Caption         =   "班别"
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
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   12285
            Top             =   540
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            Caption         =   "生产厂"
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
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   120
            Left            =   7110
            TabIndex        =   10
            Top             =   255
            Width           =   195
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   8025
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1080
         Width           =   15210
         _Version        =   393216
         _ExtentX        =   26829
         _ExtentY        =   14155
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
         MaxCols         =   49
         MaxRows         =   20
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "DGA1140C.frx":00B3
      End
   End
End
Attribute VB_Name = "DGA1140C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   HTM System
'-- Program Name      热处理冷矫直实绩查询
'-- Program ID        DGA1120C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim.Sung.Ho
'-- Coder             Kim.Sung.Ho
'-- Date              2007.12.29
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

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_prc_line, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdt_wrk_date_fr, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdt_wrk_date_to, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(cbo_shift, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(cbo_group, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_lot_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(cbo_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
               
     'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, "p", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
  
 
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="DGA1140C.P_REFER", Key:="P-R"
    sc1.Add Item:="DGA1140C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="DGA1140C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

    
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

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    txt_prc_line.Text = "3"
    sdt_wrk_date_fr.RawData = Mid(sdt_wrk_date_fr.RawData, 1, 6) & "01"
    
'    Call Gp_Sp_HdColColor(sc1.Item("Spread"), 21)
'    Call Gp_Sp_HdColColor(sc1.Item("Spread"), 23)
'    Call Gp_Sp_HdColColor(sc1.Item("Spread"), 25)
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "DG-System.INI", Me.Name)
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "DG-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
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
    
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        opt_htn_plt1.Value = True
        opt_htn_plt2.Value = False
        opt_htn_plt1.ForeColor = &HFF&
        opt_htn_plt2.ForeColor = &H0&
        txt_prc_line.Text = "1"
        sdt_wrk_date_fr.RawData = Mid(sdt_wrk_date_fr.RawData, 1, 6) & "01"
    End If

End Sub

Public Sub Form_Ref()

  Dim I As Integer
  Dim iRow As Integer
  Dim iCol As Integer
  

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        
        For iRow = 1 To ss1.MaxRows
    
               ss1.Row = iRow
               ss1.Col = 47  '20160909 加2列
                If ss1.Text = "Y" Then
                  For I = 1 To ss1.MaxCols
                       ss1.Col = I
                       ss1.ForeColor = &HC000&
                  Next
                End If
          
        Next iRow
        
        
    End If
    
End Sub

Public Sub Form_Pro()

  Dim I As Integer
  Dim iRow As Integer
  Dim iCol As Integer

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        
        For iRow = 1 To ss1.MaxRows
    
               ss1.Row = iRow
               ss1.Col = 47 '20160909 加2列
                If ss1.Text = "Y" Then
                  For I = 1 To ss1.MaxCols
                       ss1.Col = I
                       ss1.ForeColor = &HC000&
                  Next
                End If
        
       
        Next iRow
    End If
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

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

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Sub opt_htn_plt1_Click(Value As Integer)

    If opt_htn_plt1 Then
        opt_htn_plt1.ForeColor = &HFF&
        opt_htn_plt2.ForeColor = &H80000012
        opt_htn_plt3.ForeColor = &H80000012
        opt_htn_plt4.ForeColor = &H80000012
        txt_prc_line.Text = "3"
    End If
    
End Sub

Private Sub opt_htn_plt2_Click(Value As Integer)

    If opt_htn_plt2 Then
        opt_htn_plt2.ForeColor = &HFF&
        opt_htn_plt1.ForeColor = &H80000012
        opt_htn_plt3.ForeColor = &H80000012
        opt_htn_plt4.ForeColor = &H80000012
        txt_prc_line.Text = "4"
    End If
    
End Sub
Private Sub opt_htn_plt3_Click(Value As Integer)

    If opt_htn_plt3 Then
        opt_htn_plt3.ForeColor = &HFF&
        opt_htn_plt1.ForeColor = &H80000012
        opt_htn_plt2.ForeColor = &H80000012
        opt_htn_plt4.ForeColor = &H80000012
        txt_prc_line.Text = "5"
    End If
    
End Sub

Private Sub opt_htn_plt4_Click(Value As Integer)
     If opt_htn_plt4 Then
        opt_htn_plt4.ForeColor = &HFF&
        opt_htn_plt1.ForeColor = &H80000012
        opt_htn_plt2.ForeColor = &H80000012
        opt_htn_plt3.ForeColor = &H80000012
        txt_prc_line.Text = "2"
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

'Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'
'    If Gf_Sc_Authority(sAuthority, "U") Then
'        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), ss1.MaxCols - 2)
'    End If
'
'End Sub
Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
 If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)

         ss1.Col = 0

         ss1.Row = ss1.ActiveRow

        If ss1.Text = "Update" Then
    
           ss1.Col = 31 '20160909 加2列
  
           ss1.Text = sUserID
    
        End If
  End If

End Sub


Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), ss1.MaxCols - 2)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True
    
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol
    
        Case 27 '20160609 加2
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.rControl.Add Item:=ss1.ActiveCol
                DD.rControl.Add Item:=ss1.ActiveCol + 1
                
                DD.nameType = "2"
                Call Gf_EmpID_DD(M_CN1, KeyCode)
                
            Else
            
                ss1.Col = ss1.ActiveCol

                 If Len(Trim(ss1.Text)) = ss1.TypeMaxEditLen Then
                        sTemp_Code = ss1.Text
                        ss1.Col = ss1.ActiveCol + 1
                        ss1.Text = Gf_EmpNameFind(M_CN1, Trim(sTemp_Code))
                    Else
                        ss1.Col = ss1.ActiveCol + 1
                        ss1.Text = ""
                 End If
                
            End If
            
    End Select
    
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

