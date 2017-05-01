VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGE2030C 
   Caption         =   "钢板垛位变更及查询界面_CGE2030C"
   ClientHeight    =   9660
   ClientLeft      =   195
   ClientTop       =   1500
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   14355
   WindowState     =   2  'Maximized
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
      Left            =   5340
      MaxLength       =   7
      TabIndex        =   17
      Top             =   120
      Width           =   990
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
      Height          =   315
      Left            =   8610
      MaxLength       =   10
      TabIndex        =   16
      Top             =   120
      Width           =   990
   End
   Begin VB.TextBox txt_plate_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11975
      MaxLength       =   14
      TabIndex        =   15
      Top             =   90
      Width           =   1830
   End
   Begin VB.TextBox txt_o_t_addr 
      BackColor       =   &H00E0E0E0&
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
      Left            =   10005
      TabIndex        =   14
      Top             =   735
      Width           =   990
   End
   Begin VB.TextBox txt_o_t_addr_nm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11625
      TabIndex        =   13
      Top             =   735
      Width           =   3525
   End
   Begin VB.TextBox txt_o_f_addr_nm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2910
      TabIndex        =   12
      Top             =   735
      Width           =   3525
   End
   Begin VB.TextBox txt_o_f_addr 
      BackColor       =   &H00E0E0E0&
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
      Left            =   1290
      TabIndex        =   11
      Top             =   735
      Width           =   990
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
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   1350
   End
   Begin VB.TextBox text_cur_inv_code 
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
      Left            =   1425
      MaxLength       =   2
      TabIndex        =   9
      Top             =   120
      Width           =   420
   End
   Begin VB.TextBox txt_plate_cnt 
      Height          =   330
      Left            =   585
      TabIndex        =   8
      Text            =   " "
      Top             =   9420
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt_p_row 
      Height          =   330
      Left            =   1095
      TabIndex        =   7
      Text            =   " "
      Top             =   9420
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt_location1 
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
      Left            =   4680
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   2
      Top             =   9465
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txt_location3 
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
      Left            =   6990
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   1
      Top             =   9465
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox txt_location2 
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
      Left            =   5835
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Top             =   9465
      Visible         =   0   'False
      Width           =   1155
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8070
      Left            =   60
      TabIndex        =   3
      Top             =   1125
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   14235
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   1
      PaneTree        =   "CGE2030C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   8040
         Left            =   15
         TabIndex        =   4
         Top             =   15
         Width           =   7545
         _Version        =   393216
         _ExtentX        =   13309
         _ExtentY        =   14182
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
         MaxCols         =   14
         MaxRows         =   50
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGE2030C.frx":0052
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   8040
         Left            =   7620
         TabIndex        =   20
         Top             =   15
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   14182
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   1
         PaneTree        =   "CGE2030C.frx":09D8
         Begin Threed.SSFrame SSFrame1 
            Height          =   8010
            Left            =   15
            TabIndex        =   21
            Top             =   15
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   14129
            _Version        =   196609
            BackColor       =   14737632
            Begin VB.TextBox txt_plate_num 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   90
               TabIndex        =   22
               Top             =   450
               Width           =   930
            End
            Begin Threed.SSCommand ssc_move 
               Height          =   585
               Left            =   90
               TabIndex        =   23
               Top             =   1080
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   1032
               _Version        =   196609
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "移 动"
            End
            Begin Threed.SSCommand ssc_can 
               Height          =   585
               Left            =   90
               TabIndex        =   24
               Top             =   2550
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   1032
               _Version        =   196609
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "取 消"
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   8010
            Left            =   1200
            TabIndex        =   25
            Top             =   15
            Width           =   6345
            _Version        =   393216
            _ExtentX        =   11192
            _ExtentY        =   14129
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
            MaxCols         =   14
            MaxRows         =   50
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "CGE2030C.frx":0A2A
         End
      End
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   1650
      TabIndex        =   5
      Top             =   9435
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      _Version        =   196609
      ForeColor       =   255
      Caption         =   "在线钢板"
   End
   Begin Threed.SSCommand cmd_Loc_Search 
      Height          =   375
      Left            =   3585
      TabIndex        =   6
      Top             =   9435
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      Caption         =   "垛位查询"
   End
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   11010
      Top             =   735
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
      Caption         =   "名称"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   2295
      Top             =   735
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   556
      Caption         =   "名称"
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
   Begin Threed.SSCheck Chk_ss1 
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   735
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   2
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
      Caption         =   "起始垛位"
      Value           =   1
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   285
      Left            =   8835
      TabIndex        =   19
      Top             =   735
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   8421504
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
      Caption         =   "目的垛位"
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   10680
      Top             =   120
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "钢板号"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   4050
      Top             =   120
      Width           =   1260
      _ExtentX        =   2223
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
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   7320
      Top             =   120
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "目的垛位"
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
   Begin InDate.ULabel ULabel25 
      Height          =   315
      Left            =   120
      Top             =   120
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   6360
      Top             =   120
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      Caption         =   "=====>"
      Alignment       =   1
      BackColor       =   16761087
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
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   90
      X2              =   15195
      Y1              =   570
      Y2              =   585
   End
End
Attribute VB_Name = "CGE2030C"
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
'-- Program Name      钢板入库，垛位变更及库存查询界面
'-- Program ID        CGE2030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2008.04.24
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
Public sOth As String

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
Dim Sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Click_YN As Boolean
Const SS1_USERID = 14
Const SS2_USERID = 14

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_f_addr, "p", " ", " ", " ", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
Call Gp_Ms_Collection(text_cur_inv_code, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_t_addr, "p", " ", " ", " ", "r", "a", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       
    'MASTER Collection
    Mc1.Add Item:="CGE2030C.P_REFER", Key:="P-R"
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    Call Gp_Sp_Collection(ss2, 1, " ", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", "a", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGE2030C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="CGE2030C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="CGE2030C.P_SREFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="CGE2030C.P_MODIFY", Key:="P-M"
    Sc2.Add Item:="CGE2030C.P_ONEROW", Key:="P-O"
    Sc2.Add Item:="CGE2030C.P_SREFER2", Key:="P-R"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "◎"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Click_YN = False
    
End Sub

Private Sub cmd_Loc_Search_Click()
    Dim OutParam(3, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    If Trim(TXT_PLATE_NO.Text) = "" Then
        Call Gp_MsgBoxDisplay("必须输入钢板号", "", "错误提示")
        Exit Sub
    End If
    
    On Error Resume Next

    Screen.MousePointer = vbHourglass
    
    txt_location1.Text = ""
    txt_location2.Text = ""
    txt_location3.Text = ""
        
    'Return loaction1 Parameter
    OutParam(1, 1) = "arg_loaction1"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 10

    'Return loaction2 Parameter
    OutParam(2, 1) = "arg_loaction2"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 10
    
    'Return loaction3 Parameter
    OutParam(3, 1) = "arg_loaction3"
    OutParam(3, 2) = adVarChar
    OutParam(3, 3) = adParamOutput
    OutParam(3, 4) = 10
        
    sQuery = "{call AFL2010P ('PP','" & Trim(TXT_PLATE_NO.Text) & "',?,?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(3, 1), OutParam(3, 2), OutParam(3, 3), OutParam(3, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If Left(adoCmd("arg_loaction1"), 3) = "NOT" Then
        Call Gp_MsgBoxDisplay("垛位查询失败，请确认")
    Else
        txt_location1.Text = adoCmd("arg_loaction1")
        txt_location2.Text = adoCmd("arg_loaction2")
        txt_location3.Text = adoCmd("arg_loaction3")
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

    If sOth <> "" Then
       Call Form_Ref
       sOth = ""
    End If
    
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
    'Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(Sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(Sc2)
    
    sc1.Item("Spread").RetainSelBlock = False
    Sc2.Item("Spread").RetainSelBlock = False
        
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc2.Item("Spread"), "G-System.INI", Me.Name)
    
    If App.Title = "CG" Then
        text_cur_inv_code = "ZB"
    ElseIf App.Title = "EG" Then
        text_cur_inv_code = "WG"
    End If
'       text_cur_inv_code.Text = "ZB"
    
    SSSplitter2.Panes(0).LockWidth = True
    
    Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc2.Item("Spread"), "G-System.INI", Me.Name)
    
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
    Set Sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

'    If cbo_sel.Text = "I" Then
'       ss1.Row = ss1.ActiveRow
'       ss1.Col = 1
'       ss1.Text = "1"
'    ElseIf cbo_sel = "O" Then
'        ss1.Row = ss1.ActiveRow
'        ss1.Col = 1
'        ss1.Text = "2"
'    End If
'
'   Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    
    Call Form_Ref
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Sc2) Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            rControl(1).SetFocus
            txt_o_f_addr.Text = ""
            txt_o_f_addr_nm.Text = ""
            txt_o_t_addr.Text = ""
            txt_o_t_addr_nm.Text = ""
        End If
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim iRow  As Integer
    Dim sRow As Integer
    Dim tRow As Integer
    Dim sMesg As String
    
    If txt_f_addr <> "" And txt_t_addr <> "" Then
        If txt_f_addr = txt_t_addr Then
           sMesg = "  原,目标位置相同 !!!  "
           GoTo Refer_Err
        End If
    End If
    
    txt_plate_cnt = ""
    TXT_PLATE_NUM = ""
    txt_p_row = ""
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

         If Len(TXT_PLATE_NO.Text) >= 12 Then
            Call Gf_Ms_Refer(M_CN1, Mc1)
         End If

         If Gf_Sp_Refer(M_CN1, Sc2, Mc2, , , False) Or Gf_Sp_Refer(M_CN1, sc1, Mc1, , , False) Then
            sc1.Item("Spread").OperationMode = OperationModeNormal
            Sc2.Item("Spread").OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
         End If

    With ss1
    
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
         
         If Len(TXT_PLATE_NO.Text) >= 12 Then

            For iRow = 1 To .MaxRows
               .ROW = iRow
               .Col = 2
                If Trim(.Text) = TXT_PLATE_NO.Text Then
                   ss1.BackColor = &HFFC0FF
                   Exit For
                End If
            Next iRow

         End If
         
    End With
    
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
    
    Exit Sub

Refer_Err:
 
    Call Gp_MsgBoxDisplay(sMesg)

End Sub

Public Sub Form_Pro()


    Dim iRow  As Integer
    Dim sRow As Integer
    Dim tRow As Integer

    If Click_YN = True Then
       
       If ss2.MaxRows > 0 Then
       
         For iRow = 1 To ss2.MaxRows
            ss2.ROW = iRow
            ss2.Col = 0
             If ss2.Text = "Delete" Or ss2.Text = "Input" Or ss2.Text = "Update" Then
                ss2.Col = SS2_USERID
                ss2.Text = sUserID
             End If
         Next iRow

       End If
       
       If Gf_Mill_Process(M_CN1, Sc2, Mc2, , "P") Then
          If Gf_Sp_Refer(M_CN1, sc1, Mc1, , , False) Then  'and Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False) Then
             sc1.Item("Spread").OperationMode = OperationModeNormal
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
          End If
       End If
       
    Else
    
       If ss1.MaxRows > 0 Then
       
         For iRow = 1 To ss1.MaxRows
            ss1.ROW = iRow
            ss1.Col = 0
             If ss1.Text = "Delete" Or ss1.Text = "Input" Or ss1.Text = "Update" Then
                ss1.Col = SS1_USERID
                ss1.Text = sUserID
             End If
         Next iRow

       End If
    
       If Gf_Mill_Process(M_CN1, sc1, Mc1, , "P") Then
          Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       End If
    End If
    
    With ss1
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
    
    txt_plate_cnt = ""
    TXT_PLATE_NUM = ""
    txt_p_row = ""
       
End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Cpy()

'    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_ColumnsSort()

  '  Spread_ColSort.Show 1
    
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
    
    Dim i As Long
    
    With Proc_Sc("Sc").Item("Spread")
        
        If .MaxRows < 1 Then Exit Sub
        If .SelBlockRow < 1 Then Exit Sub
        
        For i = .SelBlockRow To .SelBlockRow2
            .ROW = i
            .Col = 2
            If Trim(.Text) <> "" Then
                .Col = 0
                If Trim(.Text) = "" Then
                    .Text = "Delete"
                End If
            End If
        Next i
        
    End With
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
        
    Dim iPoint As Integer
    Dim iLastRow As Integer
    Dim iMove As Integer
    Dim iCnt As Integer
    Dim iLastVal As Integer
    
    Dim plate_no As String
    Dim iPlate_cnt As Integer
    Dim iPlate_wgt As Double
    
    Dim tRow  As Integer
    
    If ss1.MaxRows <= 0 Then Exit Sub
    If ROW <= 0 Then Exit Sub
    
    If Col = 0 Then
    
            Call ss1_DblClick(1, ROW)
    
    Else
           
            lBlkcol1 = 0
            lBlkcol2 = 0
            lBlkrow1 = 0
            lBlkrow2 = 0
            
            txt_plate_cnt = ""
            TXT_PLATE_NUM = ""
            
            ss1.Col = 2
            ss1.ROW = ROW
            
            If Trim(ss1.Text) = "" Then
               txt_plate_cnt = ""
               TXT_PLATE_NUM = ""
               txt_p_row = ""
               Exit Sub
            End If
            
            ss1.Col = 1
            ss1.ROW = ROW
            txt_p_row = ROW
        
            iPoint = ss1.Text
            
            For iCnt = ROW To 1 Step -1
                ss1.Col = 2
                ss1.ROW = iCnt
                If Trim(ss1.Text) = "" Then
                   iLastRow = iCnt + 1
                   ss1.Col = 1
                   ss1.ROW = iLastRow
                   iLastVal = ss1.Text
                   iMove = iLastVal - iPoint + 1
                   txt_plate_cnt = iMove
                   TXT_PLATE_NUM = iMove
                   iCnt = 1
                 End If
            Next iCnt
            
            If txt_plate_cnt = "" Then
               txt_plate_cnt = ROW
               TXT_PLATE_NUM = ROW
            End If
            
    End If
       
    Exit Sub
    
ss1_Click_error:
    Call Gp_MsgBoxDisplay(" Not allowed Select Row", "I")
    
 End Sub

   
Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)

    Dim plate_no As String
    Dim iCnt As Integer
    Dim iPlate_cnt As Integer
    Dim iPlate_wgt As Double
    
    Dim tRow  As Integer
    
    Dim delete As String
    
    delete = ""

    If ss1.MaxRows < 1 Then Exit Sub
    
    If Col > 0 Then
    
        iPlate_cnt = 0
        iPlate_wgt = 0
        
            ss1.ROW = ROW
            ss1.Col = 0
            If ss1.Text = "Delete" Or ss1.Text = "Input" Or ss1.Text = "Update" Then
                delete = "Y"
            End If
            
            ss1.Col = 2
            plate_no = Trim(ss1.Text)
        
            If ss2.MaxRows = 0 Or plate_no = "" Then
               Exit Sub
            End If
            
            If delete = "Y" Then
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

                With ss1
                       .Col = 0
                       .Text = ""
                       .Col = 2
                       .BackColor = &H80000005
                End With
                Exit Sub
            End If
        
            ss1.ROW = ROW
            ss1.Col = 0
            ss1.Text = "Delete"
            ss1.Col = 2
            ss1.BackColor = &HFFC0FF
            
            If Click_YN = False Then
               Click_YN = True
            End If
                            
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
                            .Col = 2
                            .BackColor = &HFFC0FF
                            .Col = SS2_USERID
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
                        .Col = 2
                        .BackColor = &HFFC0FF
                        .Col = SS2_USERID
                        .Text = sUserID
                         If tRow > 1 Then
                         Call .SetActiveCell(1, tRow - 1)
                         End If
                         Exit Sub
                     End If
                     
                End If
                 
            End With
                
'    Else
'
'            With ss1
'                 If Col = 2 Then
'                    .Row = Row + 1
'                    .Col = 2
'                    If Trim(.Text) = "" And .Row <> .MaxRows + 1 Then Exit Sub
'                    .Row = Row
'                    If Trim(.Text) = "" Then
'                       .Col = 0
'                       .Text = "Input"
'                       .Col = 14
'                       .Text = sUserID
'                    Else
'                       .Col = 0
'                       .Text = "Update"
'                       .Col = 14
'                       .Text = sUserID
'                    End If
'                 End If
'            End With
    
    End If
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    ss1.ROW = ROW + 1
    ss1.Col = 2
    If Trim(ss1.Text) = "" Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
     
    End If
        
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

'    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
'
'    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
'
'    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
'        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
'    End If
'
'    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)
    
    'Call Gp_Sp_Sort(Sc2.Item("Spread"), Col, Row)
    
    If ROW <= 0 Then Exit Sub
        
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    ss2.Col = 2
    ss2.ROW = ROW

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

Private Sub ss2_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    ss2.ROW = ROW + 1
    ss2.Col = 2
    If Trim(ss2.Text) = "" Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
     
    End If
    
End Sub

'Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
'
'    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
'
'    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
'        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
'    End If
'
'    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True
'
'End Sub

Private Sub ss2_LostFocus()

    txt_plate_cnt = ""
    TXT_PLATE_NUM = ""

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

'Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'
'    If Row > 0 Then
'        Set Active_Spread = Me.ss2
'        PopupMenu MDIMain.PopUp_Spread
'    End If
'
'End Sub

Private Sub Chk_ss1_Click(Value As Integer)
    
    If Chk_ss1.Value = ssCBUnchecked Then
       If Chk_ss2.Value = ssCBUnchecked Then
            Chk_ss1.Value = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Gf_Sp_Change(Proc_Sc, sc1) Then
        Chk_ss1.ForeColor = &HFF&
        Chk_ss2.ForeColor = &H808080
        Chk_ss2.Value = ssCBUnchecked
    Else
        Chk_ss1.Value = ssCBUnchecked
        Chk_ss2.Value = ssCBChecked
    End If
        
End Sub

Private Sub Chk_ss2_Click(Value As Integer)
    
    If Chk_ss2.Value = ssCBUnchecked Then
        If Chk_ss1.Value = ssCBUnchecked Then
            Chk_ss2.Value = ssCBChecked
        End If
        Exit Sub
    End If
    
    If Gf_Sp_Change(Proc_Sc, Sc2) Then
        Chk_ss1.ForeColor = &H808080
        Chk_ss2.ForeColor = &HFF&
        Chk_ss1.Value = ssCBUnchecked
    Else
        Chk_ss2.Value = ssCBUnchecked
        Chk_ss1.Value = ssCBChecked
    End If
        
End Sub

Private Sub ssc_can_Click()
    Dim iRow As Integer
    Dim sRow As Integer
    
   If Gf_Sp_Refer(M_CN1, sc1, Mc1, , , False) And Gf_Sp_Refer(M_CN1, Sc2, Mc2, , , False) Then
   
        sc1.Item("Spread").OperationMode = OperationModeNormal
        Sc2.Item("Spread").OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        If Click_YN = True Then
           Click_YN = False
        End If
   
        With ss1
              For iRow = 1 To .MaxRows
                 .ROW = iRow
                 .Col = 2
                  If Trim(.Text) <> "" Then
                     sRow = iRow
                     Exit For
                  End If
                  sRow = .MaxRows
              Next iRow
              
              sRow = sRow + 15
              If sRow > .MaxRows Then
                 sRow = .MaxRows
              End If
              
              Call .SetActiveCell(1, sRow)
         End With
         
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
              
              sRow = sRow + 15
              If sRow > .MaxRows Then
                 sRow = .MaxRows
              End If
              
              Call .SetActiveCell(1, sRow)
         End With
   End If
   
   txt_plate_cnt = ""
   TXT_PLATE_NUM = ""
   txt_p_row = ""
 
End Sub

Private Sub ssc_move_Click()

    Dim i As Integer
    Dim iCnt As Integer
    Dim iGap As Integer
    Dim ifCnt As Integer
    Dim ifWid As Integer
    Dim ifLen As Integer
    Dim isCnt As Integer
    Dim iVal1 As Integer
    Dim iVal2 As Integer
    Dim isLen As Integer
    Dim iRow2  As Integer
    Dim iMoveCnt  As Integer
    Dim iFromRow  As Integer
    Dim iToStaRow  As Integer
    Dim sMsg  As String
    
 
    If ss2.MaxRows <= 0 Then Exit Sub
  
    If txt_t_addr = "" Then
       sMsg = "  必须选择目的垛位  !!!  "
       GoTo MOVE_CLICK_ERROR
    End If
    
    For iCnt = ss1.MaxRows To 1 Step -1
        ss1.Col = 0
        ss1.ROW = iCnt
        If Trim(ss1.Text) = "Delete" Then
           sMsg = "  请取消本次操作，重新进行操作  !!!  "
           GoTo MOVE_CLICK_ERROR
        End If
    Next

    If txt_plate_cnt = "" Then
       Exit Sub
    End If
  
    If Click_YN = False Then
       Click_YN = True
    End If
    
    If TXT_PLATE_NUM > txt_plate_cnt Or TXT_PLATE_NUM < 1 Then
           sMsg = "  移动块数错误，请确认后重新进行操作  !!!  "
           GoTo MOVE_CLICK_ERROR
    End If
    
    iFromRow = txt_p_row
    iMoveCnt = TXT_PLATE_NUM

    For iCnt = 1 To ss2.MaxRows
        ss2.Col = 2
        ss2.ROW = iCnt
        If ss2.Text <> "" Then
           iToStaRow = iCnt - 1
           iCnt = ss2.MaxRows
        Else
           iToStaRow = iCnt
        End If
    Next
    
    If iMoveCnt > iToStaRow Then
       sMsg = "  目的垛位没有足够位置放置所移钢板，重新进行操作  !!!  "
       GoTo MOVE_CLICK_ERROR
    End If
       
   ss1.SetSelection 2, iFromRow - iMoveCnt + 1, 2, iFromRow
   ss1.ClipboardCopy
     
   ss2.SetSelection 2, iToStaRow - iMoveCnt + 1, 2, iToStaRow
   ss2.ClipboardPaste
      
    With ss1
    
        For iCnt = iFromRow - iMoveCnt + 1 To iFromRow Step 1
          .ROW = iCnt
          .Col = 0
           ss1.Text = "Delete"
           For i = 1 To 2  '.MaxCols
               .Col = i
               .BackColor = &HFFC0FF
           Next
        Next
        
    End With

    With ss2

        For iCnt = iToStaRow - iMoveCnt + 1 To iToStaRow Step 1
          .ROW = iCnt
          .Col = 0
           ss2.Text = "Input"
          .ROW = iCnt
          .Col = 3
           ss2.Text = txt_t_addr

           For i = 1 To 2 ' .MaxCols
              .Col = i
              .BackColor = &HFFC0FF
           Next i
'              .Col = 1
'               ss2.Text = ss1.MaxRows - iCnt + 1
        Next iCnt

    End With

    Exit Sub
    
'Chk_ss1.Value = ssCBChecked
MOVE_CLICK_ERROR:
    Call Gp_MsgBoxDisplay(sMsg)

End Sub

Private Sub SSCommand1_Click()
       CGE2020C.Show
       CGE2020C.SetFocus
End Sub

Private Sub text_cur_inv_code_Change()

    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
    Else
      text_cur_inv.Text = ""
    End If
    
End Sub

Private Sub text_cur_inv_code_DblClick()
    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
End Sub
Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
     
        If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
        Else
          text_cur_inv.Text = ""
        End If
        
    End If
End Sub

Private Sub txt_f_addr_DblClick()
     Call txt_f_addr_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_location1_DblClick()
    txt_t_addr.Text = Mid(txt_location1.Text, 1, 7)
    Call Form_Ref
End Sub

Private Sub txt_location2_DblClick()
    txt_t_addr.Text = Mid(txt_location2.Text, 1, 7)
    Call Form_Ref
End Sub

Private Sub txt_location3_DblClick()
    txt_t_addr.Text = Mid(txt_location3.Text, 1, 7)
    Call Form_Ref

End Sub

Private Sub txt_t_addr_DblClick()
     Call txt_t_addr_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_f_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        If text_cur_inv_code.Text = "ZB" Then
           DD.sKey = "F0037"
        ElseIf text_cur_inv_code.Text = "WG" Then
           DD.sKey = "F0036"
        ElseIf text_cur_inv_code.Text = "52" Then
           DD.sKey = "F0038"
        Else
           DD.sKey = "X"
        End If
        txt_f_addr.Text = "P"
        DD.rControl.Add Item:=txt_f_addr
        DD.rControl.Add Item:=txt_o_f_addr_nm
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        txt_o_f_addr.Text = txt_f_addr.Text
        
        Exit Sub
        
    End If

    If Len(Trim(txt_f_addr)) = txt_f_addr.MaxLength Then
        txt_o_f_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0033", Trim(txt_f_addr.Text), 2)
        txt_o_f_addr.Text = txt_f_addr.Text
    Else
        txt_o_f_addr_nm.Text = ""
        txt_o_f_addr.Text = ""
    End If

End Sub

Private Sub txt_t_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        If text_cur_inv_code.Text = "ZB" Then
           DD.sKey = "F0037"
        ElseIf text_cur_inv_code.Text = "WG" Then
           DD.sKey = "F0036"
        ElseIf text_cur_inv_code.Text = "52" Then
           DD.sKey = "F0038"
        Else
           DD.sKey = "X"
        End If
        txt_t_addr.Text = "P"
        DD.rControl.Add Item:=txt_t_addr
        DD.rControl.Add Item:=txt_o_t_addr_nm
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        txt_o_t_addr.Text = txt_t_addr.Text
        
        Exit Sub
        
    End If

    If Len(Trim(txt_t_addr)) = txt_t_addr.MaxLength Then
        txt_o_t_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0033", Trim(txt_t_addr.Text), 2)
        txt_o_t_addr.Text = txt_t_addr.Text
    Else
        txt_o_t_addr_nm.Text = ""
        txt_o_t_addr.Text = ""
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

Private Sub ULabel11_DblClick()

    Dim sMsg As String
    Dim mResult As String
    
    If Gf_Sp_ProceExist(sc1.Item("Spread"), True) Then Exit Sub
    
    If text_cur_inv.Text = "" Then
       sMsg = "请正确选择当前库"
       mResult = MsgBox(sMsg, vbYesNo, "重要提示")
       Exit Sub
    End If
    
    If txt_t_addr.Text <> "" Then
       sMsg = "确定对垛位（" + txt_t_addr.Text + "）进行调整吗？"
       mResult = MsgBox(sMsg, vbYesNo, "重要提示")
       If mResult = vbYes Then
           If Gp_LOC_Exec(text_cur_inv_code.Text, txt_t_addr.Text) = "" Then
              MsgBox ("垛位调整完毕 ！")
              Call Form_Ref
           Else
              MsgBox (" 垛位调整失败！")
           End If
       End If
       Exit Sub
    End If
    
End Sub

Private Sub ULabel6_DblClick()

    Dim sMsg As String
    Dim mResult As String
    
    If Gf_Sp_ProceExist(sc1.Item("Spread"), True) Then Exit Sub
    
    If text_cur_inv.Text = "" Then
       sMsg = "请正确选择当前库"
       mResult = MsgBox(sMsg, vbYesNo, "重要提示")
       Exit Sub
    End If
    
    If txt_f_addr.Text <> "" Then
       sMsg = "确定对垛位（" + txt_f_addr.Text + "）进行调整吗？"
       mResult = MsgBox(sMsg, vbYesNo, "重要提示")
       If mResult = vbYes Then
           If Gp_LOC_Exec(text_cur_inv_code.Text, txt_f_addr.Text) = "" Then
              MsgBox ("垛位调整完毕 ！")
              Call Form_Ref
           Else
              MsgBox (" 垛位调整失败！")
           End If
       End If
       Exit Sub
    End If
    
End Sub
