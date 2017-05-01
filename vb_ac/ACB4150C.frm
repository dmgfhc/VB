VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB4150C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "录入评审对象_ACB4150C"
   ClientHeight    =   9225
   ClientLeft      =   255
   ClientTop       =   1635
   ClientWidth     =   15315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15315
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9135
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   16113
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACB4150C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter3 
         Height          =   2910
         Left            =   0
         TabIndex        =   3
         Top             =   6225
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   5133
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "ACB4150C.frx":0052
         Begin FPSpread.vaSpread ss2 
            Height          =   2295
            Left            =   0
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   615
            Width           =   15225
            _Version        =   393216
            _ExtentX        =   26855
            _ExtentY        =   4048
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
            MaxCols         =   6
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACB4150C.frx":00A4
         End
         Begin Threed.SSPanel SSPanel2 
            Height          =   585
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   1032
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_title_reason_comm 
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
               Left            =   7125
               MaxLength       =   200
               TabIndex        =   19
               Tag             =   "处理代码"
               Top             =   125
               Width           =   7920
            End
            Begin VB.TextBox txt_title_reason_nm 
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
               Left            =   2085
               MaxLength       =   60
               TabIndex        =   18
               Tag             =   "处理代码"
               Top             =   125
               Width           =   3480
            End
            Begin VB.TextBox txt_title_reason_cd 
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
               Left            =   1530
               MaxLength       =   4
               TabIndex        =   17
               Tag             =   "处理代码"
               Top             =   125
               Width           =   555
            End
            Begin InDate.ULabel ULabel2 
               Height          =   315
               Left            =   120
               Top             =   120
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "代表原因"
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
            Begin InDate.ULabel ULabel9 
               Height          =   315
               Left            =   5730
               Top             =   120
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "代表原因详细"
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
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   6165
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   10874
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         PaneTree        =   "ACB4150C.frx":04F1
         Begin Threed.SSPanel SSPanel1 
            Height          =   945
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   1667
            _Version        =   196609
            BackColor       =   14737632
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_prc_line 
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
               Height          =   315
               Left            =   8040
               MaxLength       =   1
               TabIndex        =   20
               Tag             =   "连铸线"
               Top             =   120
               Width           =   555
            End
            Begin VB.TextBox txt_proc_cd 
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
               Height          =   315
               Left            =   5190
               MaxLength       =   3
               TabIndex        =   13
               Tag             =   "进程状态"
               Top             =   120
               Width           =   765
            End
            Begin VB.TextBox txt_loc 
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
               Left            =   5190
               MaxLength       =   7
               TabIndex        =   12
               Tag             =   "垛位号"
               Top             =   510
               Width           =   1155
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
               Left            =   1620
               MaxLength       =   2
               TabIndex        =   11
               Tag             =   "堆放仓库"
               Top             =   510
               Width           =   510
            End
            Begin VB.TextBox txt_cur_inv_nm 
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
               Left            =   2130
               TabIndex        =   10
               Tag             =   "堆放仓库"
               Top             =   510
               Width           =   1380
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
               Left            =   12690
               MaxLength       =   11
               TabIndex        =   9
               Tag             =   "订单号"
               Top             =   120
               Width           =   1380
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
               Left            =   14070
               TabIndex        =   8
               Tag             =   "订单号"
               Top             =   120
               Width           =   750
            End
            Begin VB.TextBox txt_slab_no 
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
               Left            =   1620
               MaxLength       =   10
               TabIndex        =   0
               Tag             =   "板坯号"
               Top             =   125
               Width           =   1245
            End
            Begin InDate.ULabel ULabel3 
               Height          =   315
               Left            =   210
               Top             =   120
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "板坯号"
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
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Left            =   6630
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
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.UDate dpt_prod_fr 
               Height          =   315
               Left            =   8040
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
            Begin InDate.UDate dpt_prod_to 
               Height          =   315
               Left            =   9570
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
            Begin InDate.ULabel ULabel5 
               Height          =   315
               Left            =   11280
               Top             =   120
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
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel12 
               Height          =   315
               Left            =   210
               Top             =   510
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "堆放仓库"
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
            Begin InDate.ULabel ULabel6 
               Height          =   315
               Left            =   3780
               Top             =   510
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "垛位号"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   3780
               Top             =   120
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "进程状态"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16711680
            End
            Begin InDate.ULabel ULabel7 
               Height          =   315
               Left            =   6630
               Top             =   120
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "连铸线"
               Alignment       =   1
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "~"
               Height          =   120
               Left            =   9450
               TabIndex        =   7
               Top             =   600
               Width           =   90
            End
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   5190
            Left            =   0
            TabIndex        =   14
            Top             =   975
            Width           =   15225
            _Version        =   393216
            _ExtentX        =   26855
            _ExtentY        =   9155
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
            MaxCols         =   31
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACB4150C.frx":0543
         End
      End
   End
End
Attribute VB_Name = "ACB4150C"
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
'-- Program Name      SLAB DELIBERATION REASON EVENT PROCESS
'-- Program ID        ACB4150C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2009.9.25
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

Private Sub Form_Define()

    Dim I As Integer
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_proc_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dpt_prod_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dpt_prod_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_cur_inv, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_cur_inv_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_loc, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ord_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(cbo_ord_item, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_prc_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For I = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, I, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next I
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB4150C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
        
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", "n", "m", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", "n", "m", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACB4150C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc"
    
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss2, 5, True)
    Call Gp_Sp_ColHidden(ss2, 6, True)

End Sub

Private Function Sp_Process(Conn As ADODB.Connection, Sc As Collection) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim sSLAB_NO As String
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim lRow, lRow2 As Long
    
    Dim adoCmd As ADODB.Command

    Sp_Process = True
    
    Screen.MousePointer = vbHourglass
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Sp_Process = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To 6
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    For lRow = 1 To ss1.MaxRows
    
        ss1.Row = lRow
        ss1.Col = 0
        
        If ss1.Text <> "" Then
        
            ss1.Col = 1
            sSLAB_NO = ss1.Text
    
            For lRow2 = 1 To ss2.MaxRows
            
                adoCmd.Parameters(0).Value = "I"
                adoCmd.Parameters(1).Value = sSLAB_NO
                adoCmd.Parameters(2).Value = txt_title_reason_cd.Text
                adoCmd.Parameters(3).Value = txt_title_reason_comm.Text
                
                ss2.Row = lRow2
                ss2.Col = 1     'REASON_CD
                adoCmd.Parameters(4).Value = ss2.Text
                ss2.Col = 3     'REASON_COMM
                adoCmd.Parameters(5).Value = ss2.Text
                ss2.Col = 5     'EMP_ID
                adoCmd.Parameters(6).Value = ss2.Text
                
                adoCmd.Execute
                    
                'Error Check
                If adoCmd("Error") <> "0" Then
                
                    ret_Result_ErrCode = adoCmd("Error")
                    ret_Result_ErrMsg = adoCmd("Messg")
            
                    sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                    
                    Call Gp_Sp_RowColor(ss1, lRow, , vbYellow)
                    Call Gp_MsgBoxDisplay(sErrMessg)
                    
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    
                    Conn.RollbackTrans
                    Sp_Process = False
                    Exit Function
            
                End If
                
            Next lRow2
            
        End If
        
    Next lRow
    
    Conn.CommitTrans
    
    Sc.Item("Spread").ReDraw = True
    
    If iProcessCount > 0 Then
        
        MDIMain.StatusBar1.Panels(1) = "提示信息：成功处理了" & iProcessCount & "条记录"
        
    End If
            
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("Sp_Process Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

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
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "C-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(sc2.Item("Spread"), 1)
    
    dpt_prod_fr.RawData = ""
    dpt_prod_to.RawData = ""
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "C-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "C-System.INI", Me.Name)
    
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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            rControl(1).SetFocus
            dpt_prod_fr.RawData = ""
            dpt_prod_to.RawData = ""
            txt_title_reason_cd.Text = ""
            txt_title_reason_nm.Text = ""
            txt_title_reason_comm.Text = ""
        End If
    End If

End Sub

Public Sub Form_Ref()

    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    If Len(txt_slab_no.Text) < 8 And txt_ord_no.Text = "" And (Replace(dpt_prod_fr.RawData, "_", "") = "" Or Replace(dpt_prod_to.RawData, "_", "") = "") Then
        Call Gp_MsgBoxDisplay("板坯号或生产日期或订单号必须输入", "I", "错误提示")
        Exit Sub
    End If
    
    If Replace(dpt_prod_fr.RawData, "_", "") <> "" Or Replace(dpt_prod_to.RawData, "_", "") <> "" Then
    
        If Len(Replace(dpt_prod_fr.RawData, "_", "")) <> 8 Or Len(Replace(dpt_prod_to.RawData, "_", "")) <> 8 Then
            Call Gp_MsgBoxDisplay("生产日期错误", "I", "错误提示")
            Exit Sub
        Else
            If DateDiff("D", CDate(dpt_prod_fr.Text), CDate(dpt_prod_to.Text)) > 3 Then
                Call Gp_MsgBoxDisplay("生产日期期限不能超过3天", "I", "错误提示")
                Exit Sub
            End If
        End If
    
    End If
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
        Call Gf_Sp_Cls(sc2)
        txt_title_reason_cd.Text = ""
        txt_title_reason_nm.Text = ""
        txt_title_reason_comm.Text = ""
    End If
            
End Sub

Public Sub Form_Pro()
        
    If Sp_Process(M_CN1, sc2) Then
    
        If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
        End If
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        ss2.MaxRows = 0
        txt_title_reason_cd.Text = ""
        txt_title_reason_nm.Text = ""
        txt_title_reason_comm.Text = ""
        
    End If

End Sub

Public Sub Form_Ins()
    
    If ss1.MaxRows <= 0 Then Exit Sub
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 5)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 5)
    
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

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("Sc"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    Dim I As Integer
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
    
'    For I = BlockRow To BlockRow2
'
'        ss1.Row = I
'        ss1.Col = 0
'
'        If ss1.Text <> "选择" Then
'            ss1.Col = 0
'            ss1.Text = "选择"
'            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, I, I, , &HFFFF80)
'        Else
'           ss1.Col = 0
'           ss1.Text = ""
'           Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, I, I)
'        End If
'
'    Next I
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 0
        
    If ss1.Text <> "选择" Then
        ss1.Col = 0
        ss1.Text = "选择"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
    Else
       ss1.Col = 0
       ss1.Text = ""
       Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
    End If

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 3
    If ss1.Text = "" Then Exit Sub
    
    ss1.Col = 1
    ACB4140C.txt_slab_no1.Text = ss1.Text
    ACB4140C.opt_all.Value = True
    ACB4140C.opt_all.ForeColor = &HFF&
    ACB4140C.opt_in_wait.ForeColor = &H80000012
    ACB4140C.opt_wait.ForeColor = &H80000012
    ACB4140C.opt_complete.ForeColor = &H80000012
    ACB4140C.txt_rec_sts.Text = "A"

    Call ACB4140C.Form_Ref
    Call ACB4140C.ss1_Click(1, 1)
    
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

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    Dim I As Integer
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
    End If
    
    If Mode = 1 And Col = 4 Then
    
        If ss2.Text = "0" Or ss2.Text = "" Then
        
            For I = 1 To ss2.MaxRows
            
                ss2.Row = I
                ss2.Col = 4
                
                If I <> Row Then
                    ss2.Text = "0"
                End If
            
            Next I
            
            ss2.Row = Row
            
            ss2.Col = 1
            txt_title_reason_cd.Text = ss2.Text
            ss2.Col = 2
            txt_title_reason_nm.Text = ss2.Text
            ss2.Col = 3
            txt_title_reason_comm.Text = ss2.Text
            
            ss2.Col = 4
            ss2.Tag = ss2.Text
            
        Else
        
            txt_title_reason_cd.Text = ""
            txt_title_reason_nm.Text = ""
            txt_title_reason_comm.Text = ""
        
        End If
        
    End If
    
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 5)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code, sQuery As String

    If ss2.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss2.ActiveCol
    
        Case 1    'REASON_CD
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss2
                
                DD.sWitch = "SP"
                DD.sKey = "C0017"
                DD.rControl.Add Item:=1
                DD.rControl.Add Item:=2
                
                DD.nameType = "2"
                'DD.sWhere = "AND CD  <>  '9090' "
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            Else
            
                ss2.Col = ss2.ActiveCol
                
                If Len(Trim(ss2.Text)) = ss2.TypeMaxEditLen Then
                
                    sTemp_Code = ss2.Text
                    ss2.Col = 2
                    ss2.Text = Gf_ComnNameFind(M_CN1, "C0017", Trim(sTemp_Code), 2)
                    
                Else
                
                    ss2.Col = 2
                    ss2.Text = ""
                    
                End If
            
            End If
            
            ss2.Col = 1
            
'            If ss2.Text = "9090" Then
'                ss2.Col = 1
'                ss2.Text = ""
'                ss2.Col = 2
'                ss2.Text = ""
'            End If
            
    End Select

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
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub txt_cur_inv_DblClick()

    Call txt_cur_inv_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_cur_inv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
       DD.sWitch = "MS"
       DD.sKey = "C0013"

       DD.rControl.Add Item:=txt_cur_inv
       DD.rControl.Add Item:=txt_cur_inv_nm

       DD.nameType = "2"
       Call Gf_Common_DD(M_CN1, KeyCode)
    
    Else
    
        If Len(Trim(txt_cur_inv.Text)) = txt_cur_inv.MaxLength Then
            txt_cur_inv_nm.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv.Text, 2)
        Else
            txt_cur_inv_nm.Text = ""
        End If

    End If
    
End Sub

Private Sub txt_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String

    If Len(Trim(txt_ord_no.Text)) = txt_ord_no.MaxLength Then
    
        If cbo_ord_item.Text <> "" Then Exit Sub
        
        txt_ord_no.Text = StrConv(txt_ord_no.Text, vbUpperCase)
        
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(txt_ord_no.Text) & "'"
        Call Gf_ComboAdd(M_CN1, cbo_ord_item, sQuery)

    Else
        cbo_ord_item.Clear
    End If


End Sub

Private Sub txt_title_reason_cd_DblClick()

    Call txt_title_reason_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_title_reason_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0017"
        DD.rControl.Add Item:=txt_title_reason_cd
        DD.rControl.Add Item:=txt_title_reason_nm
        
        DD.nameType = "2"
        'DD.sWhere = "AND CD  <>  '9090' "
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_title_reason_cd)) = txt_title_reason_cd.MaxLength Then
            txt_title_reason_nm.Text = Gf_ComnNameFind(M_CN1, "C0017", Trim(txt_title_reason_cd.Text), 2)
        Else
            txt_title_reason_nm.Text = ""
        End If
        
    End If
    
'    If txt_title_reason_cd.Text = "9090" Then
'        txt_title_reason_cd.Text = ""
'        txt_title_reason_nm.Text = ""
'    End If

End Sub

