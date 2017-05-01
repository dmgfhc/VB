VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB6060C 
   Caption         =   "板坯库库存修改及查询_ACB6060C"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   16020
   WindowState     =   2  'Maximized
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
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   135
      Width           =   1140
   End
   Begin VB.ComboBox cbo_inv 
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
      ItemData        =   "ACB6060C.frx":0000
      Left            =   1005
      List            =   "ACB6060C.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "连铸机号"
      Top             =   135
      Width           =   870
   End
   Begin VB.TextBox txt_MV_LST_NO 
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
      Left            =   13140
      TabIndex        =   23
      Top             =   135
      Width           =   2100
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
      Left            =   13725
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   22
      Top             =   -15
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_slab_no 
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
      Left            =   10065
      MaxLength       =   10
      TabIndex        =   3
      Top             =   135
      Width           =   1455
   End
   Begin VB.TextBox txt_f_addr 
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
      IMEMode         =   3  'DISABLE
      Left            =   4830
      MaxLength       =   7
      TabIndex        =   1
      Top             =   135
      Width           =   975
   End
   Begin VB.TextBox txt_t_addr 
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
      IMEMode         =   3  'DISABLE
      Left            =   7455
      MaxLength       =   7
      TabIndex        =   2
      Top             =   135
      Width           =   975
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
      Left            =   12255
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   17
      Top             =   -75
      Visible         =   0   'False
      Width           =   1455
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
      Left            =   10770
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   16
      Top             =   -90
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cbo_ccm_line 
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
      ItemData        =   "ACB6060C.frx":0016
      Left            =   5655
      List            =   "ACB6060C.frx":0018
      TabIndex        =   15
      Tag             =   "连铸机号"
      Top             =   -120
      Visible         =   0   'False
      Width           =   615
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1020
      Left            =   150
      TabIndex        =   0
      Top             =   510
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   1799
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox txt_o_f_addr 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txt_o_f_addr_nm 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   2385
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   3570
      End
      Begin VB.TextBox txt_o_t_addr 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   10395
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txt_o_t_addr_nm 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Width           =   3570
      End
      Begin VB.TextBox txt_slab_cnt 
         Height          =   330
         Left            =   6435
         TabIndex        =   7
         Text            =   " "
         Top             =   555
         Width           =   465
      End
      Begin VB.TextBox txt_p_row 
         Enabled         =   0   'False
         Height          =   330
         Left            =   8520
         TabIndex        =   6
         Text            =   " "
         Top             =   555
         Width           =   465
      End
      Begin Threed.SSOption opt_Left_Right 
         Height          =   285
         Left            =   210
         TabIndex        =   5
         Top             =   630
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "左边->右边"
         Value           =   -1
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   120
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         Caption         =   "起始垛位号"
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
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   9120
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         Caption         =   "目的垛位号"
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
      Begin Threed.SSCommand ssc_can 
         Height          =   330
         Left            =   7725
         TabIndex        =   12
         Top             =   555
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&取消"
      End
      Begin Threed.SSCommand ssc_move 
         Height          =   330
         Left            =   6930
         TabIndex        =   13
         Top             =   555
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "&移动"
      End
      Begin Threed.SSOption opt_Right_Left 
         Height          =   285
         Left            =   13410
         TabIndex        =   14
         Top             =   630
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "左边<-右边 "
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   8790
      Top             =   135
      Width           =   1245
      _ExtentX        =   2196
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   3555
      Top             =   135
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "起始垛位号"
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
      Left            =   6180
      Top             =   135
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "目的垛位号"
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7650
      Left            =   120
      TabIndex        =   18
      Top             =   1575
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   13494
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACB6060C.frx":001A
      Begin FPSpread.vaSpread ss1 
         Height          =   7650
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   7695
         _Version        =   393216
         _ExtentX        =   13573
         _ExtentY        =   13494
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
         MaxCols         =   16
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB6060C.frx":006C
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   7650
         Left            =   7755
         TabIndex        =   20
         Top             =   0
         Width           =   7350
         _Version        =   393216
         _ExtentX        =   12965
         _ExtentY        =   13494
         _StockProps     =   64
         Enabled         =   0   'False
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
         MaxCols         =   16
         MaxRows         =   2
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB6060C.frx":084A
      End
   End
   Begin Threed.SSCommand cmd_Loc_Search 
      Height          =   315
      Left            =   9465
      TabIndex        =   21
      Top             =   -180
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "垛位查询"
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   11865
      Top             =   135
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "移拨码单号"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   165
      Top             =   135
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      Caption         =   "仓库"
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
Attribute VB_Name = "ACB6060C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       System Management
'-- Sub_System Name   Code Management
'-- Program Name      Common Code
'-- Program ID        ACB6060C
'-- Document No       Q-00-0010(Specification)
'-- Designer          YIDUJUN
'-- Coder             YIDUJUN
'-- Date              2010.5.25
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
Public iMoveCnt As Integer          'Move Slabs Count
Public iFromRow As Integer          'From Slab Row
Public iToStaRow As Integer         'To Slab Row
Public TopSlabRow As Integer
Public TopSlabNo As String
Public Active_LForm As String       'Form Active AFL2040C
Public S1_Click As String
Public Max_Rows As Integer
Public To_Bedseq As String

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
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sChkFlag As String

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_f_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_t_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(cbo_ccm_line, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_MV_LST_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(cbo_inv, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, " ", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    Call Gp_Sp_Collection(ss2, 1, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB6060C.P_MODIFY1", Key:="P-M"
    sc1.Add Item:="ACB6060C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ACB6060C.P_MODIFY1", Key:="P-M"
    sc2.Add Item:="ACB6060C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    
    Call Gp_Sp_ColHidden(ss1, 15, True)
    Call Gp_Sp_ColHidden(ss2, 15, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cbo_inv_Change()
    If Len(Trim(cbo_inv.Text)) = 2 Then
          text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", cbo_inv.Text, 2)
          Exit Sub
    Else
          text_cur_inv.Text = ""
    End If
End Sub

Private Sub cbo_inv_Click()
    If Len(Trim(cbo_inv.Text)) = 2 Then
          text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", cbo_inv.Text, 2)
          Exit Sub
    Else
          text_cur_inv.Text = ""
    End If
End Sub

Private Sub Form_Activate()
    If Active_LForm = "1" Then
       Call Form_Ref
       Active_LForm = ""
    End If
    
    ss1.MaxRows = Max_Rows
    ss1.Row = Max_Rows
    ss1.Col = 1
    ss1.Action = ActionActiveCell
    
    ss2.MaxRows = Max_Rows
    ss2.Row = Max_Rows
    ss2.Col = 1
    ss2.Action = ActionActiveCell
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(7).Enabled = False
    
    If cbo_inv = "XK" Then
        txt_o_f_addr.Text = txt_f_addr.Text
        txt_o_f_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0034", Trim(txt_f_addr.Text), 2)
        
        txt_o_t_addr.Text = txt_t_addr.Text
        txt_o_t_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0034", Trim(txt_t_addr.Text), 2)
    ElseIf cbo_inv = "52" Then
        txt_o_f_addr.Text = txt_f_addr.Text
        txt_o_f_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0035", Trim(txt_f_addr.Text), 2)
        
        txt_o_t_addr.Text = txt_t_addr.Text
        txt_o_t_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0035", Trim(txt_t_addr.Text), 2)
    End If
    
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Max_Rows = 1000
    
    Screen.MousePointer = vbHourglass
    
    'cbo_ccm_line控件下拉框添加三个选项：1,2,3
    cbo_ccm_line.AddItem "1"
    cbo_ccm_line.AddItem "2"
    cbo_ccm_line.AddItem "3"
    
    'cbo_inv控件下拉框添加两个选项：XK,52
'    cbo_inv.AddItem "XK"
'    cbo_inv.AddItem "52"
    cbo_inv.ListIndex = 0
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(7).Enabled = False
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    sc1.Item("Spread").RetainSelBlock = False
    sc2.Item("Spread").RetainSelBlock = False
    
    Call Gp_Spl_SizeGet(SSSplitter1, "F-System.INI", Me.Name, "W")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "F-System.INI", Me.Name)
 
    
    sChkFlag = "ON"
    
    Screen.MousePointer = vbDefault
              
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "F-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    
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


Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("pControl"))
            txt_o_f_addr = ""
            txt_o_f_addr_nm = ""
            txt_o_t_addr = ""
            txt_o_t_addr_nm = ""
            txt_slab_cnt = ""
            txt_p_row = ""
            txt_location1 = ""
            txt_location2 = ""
            txt_location3 = ""
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            MDIMain.MenuTool.Buttons(12).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(7).Enabled = False
            
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
                
                ss1.MaxRows = Max_Rows
                ss1.Row = Max_Rows
                ss1.Col = 1
                ss1.Action = ActionActiveCell
                
                ss2.MaxRows = Max_Rows
                ss2.Row = Max_Rows
                ss2.Col = 1
                ss2.Action = ActionActiveCell
                
                txt_f_addr.SetFocus
        End If
    End If

End Sub

Public Sub Form_Ref()
    Dim iCnt     As Integer
    Dim sFromLoc As String
    Dim sToLoc   As String
    Dim oSpr     As Variant
    
    On Error GoTo Refer_Err

    Dim iRow, iCol, MaxCnt, iStemp As Integer
    Dim sMsg, SMESG, sTemp, sQuery As String
    
    If cbo_inv.Text = "" Then
        MsgBox "请选择仓库！", vbCritical, "系统提示信息"
        Exit Sub
    End If
    
    If Trim(txt_f_addr) <> "" And Trim(txt_f_addr) <> "S0X0101" And Trim(txt_f_addr) <> "S0Y0101" Then
        sQuery = "SELECT * FROM FP_STDYARD WHERE LOCATION = '" + txt_f_addr + "' AND YARD_KND = '" + cbo_inv + "'"
        If Gf_FloatFind(M_CN1, sQuery) = 0 Then
           
           MsgBox txt_f_addr.Tag & "不正确，请重新输入！", vbCritical, "系统提示信息"
           Exit Sub
        End If
    End If
    
    If Trim(txt_t_addr) <> "" And Trim(txt_t_addr) <> "S0X0101" And Trim(txt_t_addr) <> "S0Y0101" Then
        sQuery = "SELECT * FROM FP_STDYARD WHERE LOCATION = '" + txt_t_addr + "' AND YARD_KND = '" + cbo_inv + "'"
        If Gf_FloatFind(M_CN1, sQuery) = 0 Then
           
           MsgBox txt_t_addr.Tag & "不正确，请重新输入！", vbCritical, "系统提示信息"
           Exit Sub
        End If
    End If
    
    If opt_Left_Right Then
        sFromLoc = txt_f_addr
        sToLoc = txt_t_addr
    Else
        sFromLoc = txt_t_addr
        sToLoc = txt_f_addr
    End If
    
    If sFromLoc <> "" And sToLoc <> "" Then
       
        If sFromLoc = sToLoc Then
           sMsg = "起始垛位号和目的垛位号相同！请重新选择起始垛位号和目的垛位号！"
           GoTo Refer_Err
        ElseIf sFromLoc <> "S0X0101" And sFromLoc <> "S0Y0101" Then
           If txt_slab_no <> "" Then
              sMsg = "不是在线板坯入库，板坯号和目的垛位号不能同时输入！"
              GoTo Refer_Err
           End If
        End If
    End If
            
    If txt_f_addr = "S0X0101" Or txt_f_addr = "S0Y0101" Then
       Call Gp_Sp_ColHidden(ss1, 1, True)
    Else
       Call Gp_Sp_ColHidden(ss1, 1, False)
    End If
    
    If txt_t_addr = "S0X0101" Or txt_t_addr = "S0Y0101" Then
       Call Gp_Sp_ColHidden(ss2, 1, True)
    Else
       Call Gp_Sp_ColHidden(ss2, 1, False)
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    ss1.MaxRows = 0
    ss2.MaxRows = 0
    
    SMESG = Gf_Ms_NeceCheck(nControl)
    If SMESG = "OK" Then
    
         SMESG = Gf_Ms_NeceCheck2(mControl)
        If SMESG = "OK" Then
       
            If Gf_Sp_Refer(M_CN1, sc1, Mc1, Nothing, Nothing, False) Then
                sc1.Item("Spread").OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                MDIMain.MenuTool.Buttons(11).Enabled = False
                MDIMain.MenuTool.Buttons(12).Enabled = False
                MDIMain.MenuTool.Buttons(7).Enabled = False
                MaxCnt = ss1.MaxRows
                ss1.MaxRows = Max_Rows
                ss1.Row = Max_Rows
                ss1.Col = 1
                ss1.Action = ActionActiveCell
                    
                With ss1
'                     For iRow = .MaxCols To 1 Step -1
'                         .Row = iRow
'                         .Col = 2
'                         If .Text = "" Then
'                            For iCol = 1 To .MaxCols
'                                .Col = iCol
'                                .Text = ""
'                            Next iCol
'                         End If
'                     Next iRow
                
                     For iRow = MaxCnt To 1 Step -1
                         .Row = iRow
                         For iCol = 1 To .MaxCols
                             .Col = iCol
                             sTemp = .Text
                             .Text = ""
                             .Row = .Row + Max_Rows - MaxCnt
                             .Text = sTemp
                             .Row = iRow
                         Next iCol
                     Next iRow
                     
                     For iRow = 1 To .MaxRows
                         .Row = iRow
                         .Col = 16
                         .Text = cbo_inv.Text
                     Next iRow
                                         
                    If (sFromLoc = "S0X0101" Or sFromLoc = "S0Y0101" Or Trim(sFromLoc) = "") And txt_slab_no <> "" Then
                       .Col = 2
                       For iRow = Max_Rows - MaxCnt + 1 To Max_Rows
                           .Row = iRow
                           If .Text = txt_slab_no Then
                              .Col = 0
                              .Text = "Delete"
                              iFromRow = .Row
                              iMoveCnt = 1
                              txt_slab_cnt = 1
                              For iCol = 1 To 15
                                  .Col = iCol
                                  .BackColor = &HFF
                              Next iCol
                              
                              Exit For
                           End If
                       Next iRow
                    End If
                End With
            Else
                
                sc1.Item("Spread").MaxRows = Max_Rows
                sc1.Item("Spread").Row = Max_Rows
                sc1.Item("Spread").Col = 1
                sc1.Item("Spread").Action = ActionActiveCell
                With ss1
                     For iRow = 1 To .MaxRows
                         .Row = iRow
                         .Col = 16
                         .Text = cbo_inv.Text
                     Next iRow
                End With
            End If
            
            
            If Gf_Sp_Refer(M_CN1, sc2, Mc1, Nothing, Nothing, False) Then
                sc2.Item("Spread").OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                MDIMain.MenuTool.Buttons(12).Enabled = False
                MDIMain.MenuTool.Buttons(11).Enabled = False
                MDIMain.MenuTool.Buttons(7).Enabled = False
                MaxCnt = ss2.MaxRows
                
                sc2.Item("Spread").MaxRows = Max_Rows
                sc2.Item("Spread").Row = Max_Rows
                sc2.Item("Spread").Col = 1
                sc2.Item("Spread").Action = ActionActiveCell
                
                With ss2
                     For iRow = MaxCnt To 1 Step -1
                         .Row = iRow
                         For iCol = 1 To .MaxCols
                             .Col = iCol
                             sTemp = .Text
                             .Text = ""
                             .Row = .Row + Max_Rows - MaxCnt
                             .Text = sTemp
                             .Row = iRow
                         Next iCol
                     Next iRow
                     
                     For iRow = 1 To .MaxRows
                         .Row = iRow
                         .Col = 16
                         .Text = cbo_inv.Text
                     Next iRow
  
                    If txt_slab_no <> "" Then
                         sQuery = "SELECT * FROM FP_SLABYARD WHERE SLAB_NO = '" + txt_slab_no + "' AND YARD_KND = '" + cbo_inv + "' "
                         If Gf_FloatFind(M_CN1, sQuery) = 0 Then
                            .Row = Max_Rows - MaxCnt
                            .Col = 0
                            .Text = "Input"
                            .Col = 1
                            .Row = Max_Rows - MaxCnt + 1
                             iStemp = .Text
                            .Row = Max_Rows - MaxCnt
                             
                             iMoveCnt = 1
                             iToStaRow = .Row
                             txt_slab_cnt = 1
                            
                            .Text = iStemp + 1
                            .Col = 2
                            .Text = txt_slab_no.Text
                            .Col = 3
                            .Text = sToLoc
                            .Col = 15
                            .Text = sFromLoc
                            ssc_can.Enabled = True
                        
                             For iCol = 1 To 14
                                .Col = iCol
                                .BackColor = &HFF
                             Next iCol
                             
                             MDIMain.MenuTool.Buttons(4).Enabled = True
                         Else
                            
                            If sToLoc = "" Then
                               .Col = 2
                                For iRow = Max_Rows To Max_Rows - MaxCnt + 1 Step -1
                                    .Row = iRow
                                     If .Text = txt_slab_no Then
                                        .SetSelection 2, .Row, 2, .Row
                                        .ForeColor = &HFF
                                     End If
                                Next iRow
                            End If
                            
                            If sFromLoc = "S0X0101" Or sFromLoc = "S0Y0101" Then
                               MsgBox "板坯 " + txt_slab_no + " 已经在库中，不需再做入库处理！", vbInformation, "系统提示信息"
                               txt_slab_no = ""
                            End If
                            
                            Exit Sub
                         End If
                    End If
                End With
            Else
                
                sc2.Item("Spread").MaxRows = Max_Rows
                sc2.Item("Spread").Row = Max_Rows
                sc2.Item("Spread").Col = 1
                sc2.Item("Spread").Action = ActionActiveCell
                With ss2
                     For iRow = 1 To .MaxRows
                         .Row = iRow
                         .Col = 16
                         .Text = cbo_inv.Text
                     Next iRow
                End With
                
                If txt_slab_no <> "" And (sFromLoc = "S0X0101" Or sFromLoc = "S0Y0101" Or Trim(sFromLoc) = "") Then
                   sQuery = "SELECT * FROM FP_SLABYARD WHERE SLAB_NO = " + txt_slab_no + "  AND YARD_KND = '" + cbo_inv + "'"
                   If Gf_FloatFind(M_CN1, sQuery) = 0 Then

                        With ss2
                             .Row = Max_Rows
                             .Col = 0
                             .Text = "Input"
                             .Col = 1
                             .Text = "1"
                             .Col = 2
                             .Text = txt_slab_no.Text
                             .Col = 3
                             .Text = sToLoc
                             .Col = 15
                             .Text = sFromLoc
                             
                             iMoveCnt = 1
                             iToStaRow = .Row
                             txt_slab_cnt = 1
                             
                             For iCol = 1 To 14
                                .Col = iCol
                                .BackColor = &HFF
                             Next iCol
                             
                             For iRow = 1 To .MaxRows
                                 .Row = iRow
                                 .Col = 16
                                 .Text = cbo_inv.Text
                             Next iRow
                             
                             MDIMain.MenuTool.Buttons(4).Enabled = True
                        End With
                        ssc_can.Enabled = True
                   Else
                        MsgBox "板坯 " + txt_slab_no + " 已经在库中，不需再做入库处理！", vbInformation, "系统提示信息"
                        txt_slab_no = ""
                        Exit Sub
                   End If
                
                Else
'                   If sToLoc <> "" And sToLoc <> "S0A0101" And sToLoc <> "S0Q0101" And opt_Left_Right.Enabled = True Then
'                      MsgBox "垛位 " + sToLoc + " 没有板坯！", vbInformation, "系统提示信息"
'                   ElseIf (sToLoc = "S0A0101" Or sToLoc = "S0Q0101") Then
'                      MsgBox "没有在线板坯等待入库！", vbInformation, "系统提示信息"
'                   End If
                End If

            End If
                        
        Else
            SMESG = SMESG + "长度不正确"
            Call Gp_MsgBoxDisplay(SMESG)
        End If
    
    Else
        SMESG = SMESG + "必须输入"
        Call Gp_MsgBoxDisplay(SMESG)
        
    End If
    
    If txt_slab_no = "" Then
       txt_slab_cnt = ""
       ssc_move.Enabled = False
       ssc_can.Enabled = False
       txt_p_row = ""
    End If
    
      
    
    
'    txt_MV_LST_NO.Text = ""
    
''     If iToStaRow = 0 And AFL2040C.Active_CForm <> "" Then
'    If AFL2040C.Active_CForm <> "" Then
'           'iFromRow = txt_p_row
'       If Trim(txt_slab_cnt) <> "" Then
'          iMoveCnt = CInt(txt_slab_cnt)
'       End If
'
'       For iCnt = Max_Rows To 1 Step -1
'           ss2.Col = 1
'           ss2.Row = iCnt
'           If ss2.Text = "" Then
'              iToStaRow = iCnt + 1
'              iCnt = 1
'              Exit For
'           End If
'       Next iCnt
'
'        ss2.Row = iToStaRow - iMoveCnt + 1
'        ss2.Col = 1
'        To_Bedseq = Format(ss2.Text, "0#")
'
'        sQuery = "SELECT MAX_CNT FROM FP_STDYARD WHERE LOCATION ='" + sToLoc + "' AND YARD_KND = '00'"
'        ssc_move.Enabled = False
'        'If ss2.Text <> "" Then
'            If sToLoc <> "S0A0101" And To_Bedseq > Gf_FloatFind(M_CN1, sQuery) Then
'               MsgBox "已超出目的垛位的存放能力！当前操作无法继续！", vbCritical, "系统提示信息"
'               Call ssc_can_Click
'               Exit Sub
'            End If
'        'End If
'
'     End If
     
    Exit Sub

Refer_Err:
    Call Gp_MsgBoxDisplay(sMsg)
 
End Sub

Public Function Sp_Refer(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                            Optional MsgChk As Boolean = True) As Boolean

On Error GoTo SpreadRef_Error

    Dim sQuery As String
    Dim sMsg As String

'    If MsgChk Then
'        If Gf_Sp_ProceExist(Sc.Item("Spread")) Then
'            Gf_Sp_Refer = True
'            Exit Function
'        End If
'    End If

    If Not MC Is Nothing Then
        Sp_Refer = Sp_Display(Conn, Sc.Item("Spread"), Gf_Ms_MakeQuery(Sc.Item("P-R"), "R", MC("pControl")), _
                                    Sc.Item("pColumn"), MsgChk)
        If Sp_Refer Then Call Gp_Ms_ControlLock(MC!lControl, True)
    Else
         Sp_Refer = Sp_Display(Conn, Sc.Item("Spread"), Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-R"), _
                                    "R", Sc.Item("aColumn"), 1), Sc.Item("pColumn"), MsgChk)
    End If
    
    If Sp_Refer Then
        'Sc!Spread.SetFocus
    End If
        
    Exit Function
    
SpreadRef_Error:

    Call Gp_MsgBoxDisplay("Failed on data inquiry")
     Sp_Refer = False

End Function

Public Function Sp_Display(Conn As ADODB.Connection, sPname As vaSpread, sQuery As String, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo SpreadDisplay_Error
    
    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
 
    With sPname

         Sp_Display = True
        
        .ReDraw = False
        .MaxRows = 0
        .MaxRows = Max_Rows: iCount = 0
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then

            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")

              Sp_Display = False
            .ReDraw = True

            AdoRs.Close
            Set AdoRs = Nothing

            Screen.MousePointer = vbDefault

            Exit Function

        End If

        ArrayRecords = AdoRs.GetRows

        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 1) <> 0 Then
        
            '.MaxRows = UBound(ArrayRecords, 2) + 1
        
            For iRowCount = 0 To UBound(ArrayRecords, 2)
            
                .Row = Max_Rows - iRowCount
                
                For iColcount = 0 To .MaxCols - 1
                
                    .Col = iColcount + 1
                    
                    Select Case .CellType
                    
                        Case SS_CELL_TYPE_CHECKBOX
                            If VarType(ArrayRecords(iColcount, iRowCount)) <> vbNull Or _
                               Trim(ArrayRecords(iColcount, iRowCount)) = "1" Then
                                .Text = Trim(ArrayRecords(iColcount, iRowCount))
                            End If
                            
                        Case SS_CELL_TYPE_COMBOBOX
                            If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or _
                               Trim(ArrayRecords(iColcount, iRowCount)) = "" Then
                                .Value = 0
                            Else
                                .Value = Trim(ArrayRecords(iColcount, iRowCount))
                            End If
                            
                        Case SS_CELL_TYPE_DATE
                            If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                                .Text = ""
                            Else
                                .Text = Mid(Trim(ArrayRecords(iColcount, iRowCount)), 1, 4) & "-" & _
                                        Mid(Trim(ArrayRecords(iColcount, iRowCount)), 5, 2) & "-" & _
                                        Mid(Trim(ArrayRecords(iColcount, iRowCount)), 7, 2)
                            End If
                            
                        Case Else
                            If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(iColcount, iRowCount))
                            End If
                            
                    End Select
                    
                Next iColcount
                
            Next iRowCount
            
        End If
        
        If Not lColumn Is Nothing Then
            
            'lControl Lock
            For iCount = 1 To lColumn.Count

                .Protect = True
                .Col = lColumn(iCount): .Col2 = lColumn(iCount)
                .Row = 1: .Row2 = .MaxRows
                .BlockMode = True: .Lock = True
                .BlockMode = False

            Next iCount

        End If
        
        .ReDraw = True
        
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
     Sp_Display = False
    Screen.MousePointer = vbDefault

End Function

Public Sub Form_Pro()
    Dim SlabNo      As String
    Dim I           As Integer
    Dim sQuery      As String
    Dim sFromLoc    As String

    txt_slab_no = ""
    
    If (ss1.Enabled = True And ssc_move.Enabled = True) Or (ss2.Enabled = True And ssc_move.Enabled = False) Then
        If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           For I = 1 To ss2.MaxRows
               ss2.Col = 0
               ss2.Row = I
               ss2.Text = ""
           Next I
        End If
    End If
    
    If (ss2.Enabled = True And ssc_move.Enabled = True) Or (ss1.Enabled = True And ssc_move.Enabled = False) Then
        If Gf_Sp_Process(M_CN1, sc2, Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            S1_Click = ""
            For I = 1 To ss1.MaxRows
               ss1.Col = 0
               ss1.Row = I
               ss1.Text = ""
           Next I
        End If
    End If
        
        MDIMain.MenuTool.Buttons(12).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(7).Enabled = False
        SlabNo = txt_slab_no
        txt_slab_no = ""
            
        If opt_Left_Right Then
            sFromLoc = txt_f_addr
        Else
            sFromLoc = txt_t_addr
        End If
            
        If sFromLoc = "S0X0101" Or sFromLoc = "S0Y0101" Then
           MDIMain.StatusBar1.Panels(1).Text = "板坯 " + SlabNo + " 成功入库！"
        Else
           MDIMain.StatusBar1.Panels(1).Text = "您所选板坯的垛位成功变更！"
        End If
        
    Call Form_Ref
End Sub


Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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
ss1.Row = 0
ss1.Col = 0
If ss1.Text = "◎" Then
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
Else
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End If

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    ss1.Col = 0
    ss1.Row = 0
    
    If ss1.Text = "◎" Then
        Call Gp_Sp_Del(Proc_Sc("Sc"))
    Else
        Call Gp_Sp_Del(Proc_Sc("Sc2"))
    End If
    txt_slab_no = ""
End Sub

Public Sub Spread_Can()
    ss1.Col = 0
    ss1.Row = 0
    If ss1.Text = "◎" Then
        Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
    Else
        Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc2"))
    End If
End Sub

Private Sub opt_Left_Right_Click(Value As Integer)

    If sChkFlag = "ON" Then
        If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread"), False) Then
            sChkFlag = ""
            opt_Right_Left.Value = True
            sChkFlag = "ON"
            Exit Sub
        End If
        
        Call rowEdit
        
        ULabel3.Caption = "起始垛位号"
        ULabel4.Caption = "起始垛位号"
        txt_f_addr.Tag = "起始垛位号"
        
        ULabel6.Caption = "目的垛位号"
        ULabel8.Caption = "目的垛位号"
        txt_t_addr.Tag = "目的垛位号"
        
        opt_Left_Right.ForeColor = &HFF&
        opt_Right_Left.ForeColor = &H808080
        ss1.Col = 0
        ss1.Row = 0
        ss1.Text = "◎"
        ss2.Col = 0
        ss2.Row = 0
        ss2.Text = ""
        
        ss1.Enabled = True
        ss2.Enabled = False
    End If
    
End Sub

Private Sub opt_Right_Left_Click(Value As Integer)

    If sChkFlag = "ON" Then
        If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread"), False) Then
            sChkFlag = ""
            opt_Left_Right.Value = True
            sChkFlag = "ON"
            Exit Sub
        End If
        
        Call rowEdit
        
        ULabel3.Caption = "目的垛位号"
        ULabel4.Caption = "目的垛位号"
        txt_f_addr.Tag = "目的垛位号"
        
        ULabel6.Caption = "起始垛位号"
        ULabel8.Caption = "起始垛位号"
        txt_t_addr.Tag = "起始垛位号"
        
        opt_Right_Left.ForeColor = &HFF&
        opt_Left_Right.ForeColor = &H808080
        ss1.Col = 0
        ss1.Row = 0
        ss1.Text = ""
        ss2.Col = 0
        ss2.Row = 0
        ss2.Text = "◎"
        
        ss1.Enabled = False
        ss2.Enabled = True
    End If
    
End Sub

Private Sub rowEdit()
    txt_slab_no = ""
    ss1.MaxRows = 0
    ss2.MaxRows = 0
    Call Form_Ref
    
End Sub

Private Sub opt_Right_Left_DblClick(Value As Integer)

        Dim sMsg As String
    Dim mResult As String
    
    If Gf_Sp_ProceExist(sc1.Item("Spread"), True) Then Exit Sub
    
    If txt_f_addr.Text <> "" Then
       sMsg = "确定对垛位（" + txt_f_addr.Text + "）进行垛层调整吗？"
       mResult = MsgBox(sMsg, vbYesNo, "系统提示信息")
       If mResult = vbYes Then
           If Gp_LOC_Exec(Trim(txt_f_addr.Text)) = "" Then
              MsgBox ("垛位调整完毕 ！"), vbInformation, "系统提示信息"
              Call Form_Ref
           Else
              MsgBox (" 垛位调整失败！"), vbCritical, "系统提示信息"
           End If
       End If
       Exit Sub
    End If


End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
    
    If opt_Right_Left Then Exit Sub
    Call ssc_Upd_Process(ss1, txt_f_addr, Row)
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If opt_Right_Left Then Exit Sub
    Call ssc_Upd_Process(ss1, txt_f_addr, Row)
End Sub

Private Sub ssc_Upd_Process(oSpr As vaSpread, sText As Variant, ByVal Row As Long)

    Dim iCurrRowVal As Integer

    With oSpr

        If Gf_Sc_Authority(sAuthority, "U") Then
            
            .Row = Row
            .Col = 0
            .Text = "Update"
                
            If Row = .MaxRows Then
                .Col = 1
                .Value = 1
            Else
                .Row = Row + 1
                .Col = 1
                iCurrRowVal = Val(.Value & "")
                
                .Row = Row
                .Value = iCurrRowVal + 1
            End If
            
            .Col = 3
            .Text = Trim(sText.Text)
            .Col = 14
            .Text = sUserID
        End If
    
    End With
End Sub


Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
        
    Dim iCnt, I As Integer
    Dim sFlag As String
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If opt_Right_Left Then Exit Sub
    
    With ss1
       .Col = 0
       For I = Max_Rows To 1 Step -1
           .Row = I
           If .Text = "Delete" Then
              sFlag = "Y"
              Exit For
           End If
       Next I
       
        .Col = 2
        .Row = Row
        
        If (.Text <> "") And sFlag <> "Y" Then
           S1_Click = "1"
           txt_slab_no = .Text
           
        ElseIf (.Text = "") And sFlag <> "Y" Then
           txt_slab_no = ""
           txt_slab_cnt = ""
           txt_p_row = Row
           ssc_move.Enabled = False
           ssc_can.Enabled = False
        End If
        
        If sFlag <> "Y" Then
           txt_p_row = Row
        End If
        
        If (txt_f_addr <> "S0X0101" And txt_f_addr <> "S0Y0101") And sFlag <> "Y" Then
            For iCnt = Row To 1 Step -1
                .Col = 2
                .Row = iCnt
                If .Text <> "" Then
                   I = I + 1
                ElseIf .Text = "" Then
                   
                   If .Row <> Max_Rows Then
                      .Row = .Row + 1
                   End If
                   
                   TopSlabNo = .Text
                   TopSlabRow = .Row
                   Exit For
                End If
            Next iCnt
            
            If I <> 0 Then
               txt_slab_cnt = I
            End If
        
        ElseIf txt_f_addr = "S0X0101" Or txt_f_addr = "S0Y0101" Then
        
            txt_slab_cnt = 1
            txt_p_row = .ActiveRow
            .Row = .ActiveRow
            .Col = 2
             txt_slab_no = .Text
             
    '          ss1.SetFocus
    '          ss1.SetSelection 2, .Row, 2, .Row
           
        End If
        
        If txt_slab_no <> "" And txt_slab_no <> "0" And sFlag <> "Y" Then
           ssc_move.Enabled = True
        End If
    
    End With

    If txt_slab_cnt <> "" And txt_slab_cnt <> "0" Then
       ssc_can.Enabled = True
    End If
     
    Exit Sub

ss1_Click_error:
    Call Gp_MsgBoxDisplay(" Not allowed Select Row", "I")
    
 End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
   txt_slab_no = ""
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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

Private Sub ss2_Change(ByVal Col As Long, ByVal Row As Long)
    
    If opt_Left_Right Then Exit Sub
    Call ssc_Upd_Process(ss2, txt_t_addr, Row)
    
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If opt_Left_Right Then Exit Sub
    Call ssc_Upd_Process(ss2, txt_t_addr, Row)
    
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim iCnt, I As Integer
    Dim sFlag As String
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If opt_Left_Right Then Exit Sub
   
    With ss2
       .Col = 0
       For I = Max_Rows To 1 Step -1
           .Row = I
           If .Text = "Delete" Then
              sFlag = "Y"
              Exit For
           End If
       Next I
       
        .Col = 2
        .Row = Row
        
        If (.Text <> "") And sFlag <> "Y" Then
           S1_Click = "1"
           txt_slab_no = .Text
           
        ElseIf (.Text = "") And sFlag <> "Y" Then
           txt_slab_no = ""
           txt_slab_cnt = ""
           txt_p_row = Row
           ssc_move.Enabled = False
           ssc_can.Enabled = False
        End If
        
        If sFlag <> "Y" Then
           txt_p_row = Row
        End If
        
        If (txt_t_addr <> "S0X0101" And txt_t_addr <> "S0Y0101") And sFlag <> "Y" Then
            For iCnt = Row To 1 Step -1
                .Col = 2
                .Row = iCnt
                If .Text <> "" Then
                   I = I + 1
                ElseIf .Text = "" Then
                   
                   If .Row <> Max_Rows Then
                      .Row = .Row + 1
                   End If
                   
                   TopSlabNo = .Text
                   TopSlabRow = .Row
                   Exit For
                End If
            Next iCnt
            
            If I <> 0 Then
               txt_slab_cnt = I
            End If
        
        ElseIf txt_t_addr = "S0X0101" Or txt_t_addr = "S0Y0101" Then
        
            txt_slab_cnt = 1
            txt_p_row = .ActiveRow
            .Row = .ActiveRow
            .Col = 2
             txt_slab_no = .Text
           
        End If
        
        If txt_slab_no <> "" And txt_slab_no <> "0" And sFlag <> "Y" Then
           ssc_move.Enabled = True
        End If
    
    End With

    If txt_slab_cnt <> "" And txt_slab_cnt <> "0" Then
       ssc_can.Enabled = True
    End If
     
    Exit Sub

ss2_Click_error:
    Call Gp_MsgBoxDisplay(" Not allowed Select Row", "I")
    

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
   txt_slab_no = ""
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        'Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
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
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub


Private Sub ssc_can_Click()

    If opt_Left_Right Then
        Call ssc_can_Process(ss1, ss2)
    Else
        Call ssc_can_Process(ss2, ss1)
    End If
         
End Sub

Private Sub ssc_move_Click()

    If opt_Left_Right Then
        Call ssc_move_Process(ss1, ss2)
    Else
        Call ssc_move_Process(ss2, ss1)
    End If

End Sub

Private Sub ssc_can_Process(oSpr1 As vaSpread, oSpr2 As vaSpread)

    Dim I As Integer
    Dim iCnt As Integer

    oSpr2.SetSelection 1, iToStaRow - iMoveCnt + 1, 15, iToStaRow
    oSpr2.ClipboardCut
  
    With oSpr1
      
      For iCnt = iFromRow - iMoveCnt + 1 To iFromRow Step 1
         .Row = iCnt
         .Col = 0
         .Text = ""
         For I = 1 To 3
          .Col = I
          .BackColor = &HC0FFFF
         Next I
         
         For I = 4 To .MaxCols
          .Col = I
          .BackColor = &HFFFFFF
         Next I
      Next
    End With
    
    With oSpr2
       
      For iCnt = iToStaRow To iToStaRow - iMoveCnt + 1 Step -1
         .Row = iCnt
         .Col = 0
         .Text = ""
         For I = 1 To 3
          .Col = I
          .BackColor = &HC0FFFF
          .Text = ""
         Next I
         
         For I = 4 To .MaxCols
          .Col = I
          .BackColor = &HFFFFFF
         Next
    
      Next
    End With
    
    S1_Click = ""
    txt_p_row = ""
    txt_slab_no = ""
    txt_slab_cnt = ""
    To_Bedseq = ""
    iToStaRow = 0
    iMoveCnt = 0
    
    oSpr2.MaxRows = Max_Rows
    oSpr2.Row = Max_Rows
    oSpr2.Col = 1
    oSpr2.Action = ActionActiveCell
    
    ssc_can.Enabled = False
         
End Sub

Private Sub ssc_move_Process(oSpr1 As vaSpread, oSpr2 As vaSpread)

    Dim I As Integer
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
    Dim sMsg, sSeq, sQuery As String
    Dim sTempSlabNo As String
    Dim sBedSeq As String
    Dim sFromLoc As String
    Dim sToLoc   As String
    
    If opt_Left_Right Then
        sFromLoc = txt_f_addr
        sToLoc = txt_t_addr
    Else
        sFromLoc = txt_t_addr
        sToLoc = txt_f_addr
    End If

    If sToLoc = "" Then
       sMsg = "请输入目的垛位号！"
       GoTo MOVE_CLICK_ERROR
    End If
    
    If Val(txt_slab_cnt & "") < 1 Or Val(txt_p_row & "") < 1 Then Exit Sub
    
    iFromRow = txt_p_row
    iMoveCnt = txt_slab_cnt
    
    For iCnt = Max_Rows To 1 Step -1
        oSpr2.Col = 1
        oSpr2.Row = iCnt
       If oSpr2.Text = "" Then
          iToStaRow = iCnt
          iCnt = 1
          Exit For
       End If
    Next iCnt
  
    iRow2 = iToStaRow + 1
    oSpr2.Row = iRow2
    oSpr2.Col = 2
    sTempSlabNo = oSpr2.Text
    oSpr2.Col = 6
    If (sTempSlabNo <> "" And oSpr2.Text = "") And iRow2 <> Max_Rows + 1 Then
       sMsg = "目的垛位上顶层板坯的宽度不存在！当前操作无法继续进行！"
         GoTo MOVE_CLICK_ERROR
    Else
       If oSpr2.Text <> "" Then
          iVal2 = oSpr2.Text
       End If
    End If
    
    oSpr1.Col = 6
    oSpr1.Row = iFromRow
    If oSpr1.Text = "" Then
       sMsg = "起始垛位上要移动的板坯宽度不存在！当前操作无法继续进行！"
       GoTo MOVE_CLICK_ERROR
    Else
       If oSpr1.Text <> "" Then
          iVal1 = oSpr1.Text
       End If
    End If
    
    '   If (iVal1 > iVal2) And iVal1 <> 0 And iVal2 <> 0 Then
    '      iGap = iVal1 - iVal2
    '      If iGap > 100 Then
    '         sMsg = "起始垛位上要移动的板坯过宽！该移动无法实现！"
    '         GoTo MOVE_CLICK_ERROR
    '      End If
    '   End If
    
    
    ' From Address slab  move to To Address check length
    iRow2 = iToStaRow + 1
    oSpr2.Row = iRow2
    oSpr2.Col = 2
    sTempSlabNo = oSpr2.Text
    oSpr2.Col = 7
    
    If (sTempSlabNo <> "" And oSpr2.Text = "") And iRow2 <> Max_Rows + 1 Then
       sMsg = "目的垛位上顶层板坯的长度不存在！当前操作无法继续进行！"
       GoTo MOVE_CLICK_ERROR
    Else
       If oSpr2.Text <> "" Then
          iVal2 = oSpr2.Text
       End If
    End If
    
    oSpr1.Col = 7
    oSpr1.Row = iFromRow
    If oSpr1.Text = "" Then
       sMsg = "起始垛位上要移动的板坯长度不存在！当前操作无法继续进行！"
       GoTo MOVE_CLICK_ERROR
    Else
       If oSpr1.Text <> "" Then
          iVal1 = oSpr1.Text
       End If
    End If
    
    '   If (iVal1 > iVal2) And iVal1 <> 0 And iVal2 <> 0 Then
    '      iGap = iVal1 - iVal2
    '      If iGap > 1000 Then
    '         sMsg = "起始垛位上要移动的板坯过长！该移动无法实现！"
    '         GoTo MOVE_CLICK_ERROR
    '      End If
    '   End If
    
    For iCnt = iFromRow To iFromRow - iMoveCnt + 2 Step -1
       oSpr1.Col = 6
       oSpr1.Row = iCnt
       If oSpr1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯宽度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal1 = oSpr1.Text
       
       oSpr1.Col = 6
       oSpr1.Row = iCnt - 1
       If oSpr1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯宽度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal2 = oSpr1.Text
       
    '       If iVal1 < iVal2 Then
    '          iGap = iVal2 - iVal1
    '          If iGap > 100 Then
    '             sMsg = "起始垛位上要移动的板坯宽度不符合堆放标准！"
    '              GoTo MOVE_CLICK_ERROR
    '          End If
    '       End If
       
       oSpr1.Col = 7
       oSpr1.Row = iCnt
       If oSpr1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯长度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal1 = oSpr1.Text
       oSpr1.Col = 7
       oSpr1.Row = iCnt - 1
       If oSpr1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯长度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal2 = oSpr1.Text
    '       If iVal1 < iVal2 Then
    '          iGap = iVal2 - iVal1
    '          If iGap > 1000 Then
    '             sMsg = "起始垛位上要移动的板坯长度不符合堆放标准！"
    '              GoTo MOVE_CLICK_ERROR
    '          End If
    '       End If
    Next
    
    oSpr1.SetSelection 1, iFromRow - iMoveCnt + 1, 15, iFromRow
    oSpr1.ClipboardCopy
     
    oSpr2.SetSelection 1, iToStaRow - iMoveCnt + 1, 15, iToStaRow
    oSpr2.ClipboardPaste
    
    '    If sFromLoc = "S0A0101" Then
    '        With oSpr1
    '            .Row = iFromRow
    '            .Col = 0
    '            oSpr1.Text = "Delete"
    '            For iCnt = 1 To .MaxCols
    '             .Col = iCnt
    '             .BackColor = &HFF
    '            Next
    '        End With
    '    Else
        With oSpr1
            For iCnt = iFromRow - iMoveCnt + 1 To iFromRow
              .Row = iCnt
              .Col = 0
              oSpr1.Text = "Delete"
              For I = 1 To .MaxCols
                .Col = I
                .BackColor = &HFF
              Next
            Next
        End With
    '    End If
    
    With oSpr2
       
        For iCnt = iToStaRow To iToStaRow - iMoveCnt + 1 Step -1
              .Row = iCnt
              .Col = 0
              .Text = "Input"
    
              .Col = 3
              .Text = sToLoc
    
              .Col = 14
              .Text = sUserID
              
              .Col = 15
              .Text = sFromLoc
             
              For I = 1 To .MaxCols
                .Col = I
                .BackColor = &HFF
              Next
              .Col = 1
              
              If .Row <> Max_Rows Then
                 .Row = iCnt + 1
                  sSeq = CInt(oSpr2.Text) + 1
                  .Row = iCnt
                  .Text = sSeq
              Else
                  .Text = "1"
              End If
        Next
        
        oSpr2.Row = iToStaRow - iMoveCnt + 1
        oSpr2.Col = 1
        If Len(Trim(oSpr2.Text)) = 1 Then
           To_Bedseq = "0" + oSpr2.Text
        ElseIf Len(Trim(oSpr2.Text)) = 2 Then
           To_Bedseq = oSpr2.Text
        End If
        
        sQuery = "SELECT MAX_CNT FROM FP_STDYARD WHERE LOCATION ='" + sToLoc + "' AND YARD_KND = '" + cbo_inv + "'"
        ssc_move.Enabled = False
        
        oSpr2.Row = 1
        oSpr2.Col = 2
        'If oSpr2.Text <> "" Then
            If sToLoc <> "S0X0101" And sToLoc <> "S0Y0101" And To_Bedseq > Gf_FloatFind(M_CN1, sQuery) Then
               MsgBox "已超出目的垛位的存放能力！当前操作无法继续！", vbCritical, "系统提示信息"
               Call ssc_can_Click
               Exit Sub
            End If
        'End If
    End With
    Exit Sub
    'Chk_oSpr1.Value = ssCBChecked
    
MOVE_CLICK_ERROR:
    Call Gp_MsgBoxDisplay(sMsg)

End Sub


Private Sub txt_f_addr_Change()
Dim sQuery As String

    If cbo_inv.Text = "" Then
        MsgBox "请先选择仓库！", vbCritical, "系统提示信息"
        Exit Sub
    End If
    
    If Len(txt_f_addr) = 7 And Trim(txt_f_addr) <> "S0X0101" And Trim(txt_f_addr) <> "S0Y0101" Then
       sQuery = "SELECT * FROM FP_STDYARD WHERE LOCATION = '" + txt_f_addr + "' AND YARD_KND = '" + cbo_inv + "'"
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
       
          MsgBox txt_f_addr.Tag & "不正确，请重新输入！", vbCritical, "系统提示信息"
          Exit Sub
       End If
       cbo_ccm_line.Visible = False
'    ElseIf Trim(txt_f_addr) = "S0A0101" Then
'       cbo_ccm_line.Visible = True
'    ElseIf Trim(txt_f_addr) <> "S0A0101" Then
'       cbo_ccm_line.Visible = False
    End If
End Sub

Private Sub txt_f_addr_DblClick()

    Call txt_f_addr_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_f_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If cbo_inv.Text = "" Then
        MsgBox "请先选择仓库！", vbCritical, "系统提示信息"
        Exit Sub
    End If
    
    If cbo_inv.Text = "XK" Then
    
        If KeyCode = vbKeyF4 Then
        
            txt_f_addr.Text = "S"
            DD.sWitch = "MS"
     '       DD.sKey = "F0009"
            DD.sKey = "F0034"
            DD.rControl.Add Item:=txt_f_addr
            DD.rControl.Add Item:=txt_o_f_addr_nm
            
            DD.nameType = "2"
            
            Call Gf_Common_DD(M_CN1, KeyCode)
            txt_o_f_addr.Text = txt_f_addr.Text
            Exit Sub
            
        End If
    
    
        If Len(Trim(txt_f_addr)) = txt_f_addr.MaxLength Then
            txt_o_f_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0034", Trim(txt_f_addr.Text), 2)
        Else
            txt_o_f_addr_nm.Text = ""
        End If
        
        If Len(Trim(txt_f_addr)) = 7 Then
           txt_o_f_addr.Text = txt_f_addr.Text
        Else
           txt_o_f_addr.Text = ""
        End If
        
    ElseIf cbo_inv.Text = "52" Then
    
        If KeyCode = vbKeyF4 Then
        
            txt_f_addr.Text = "S"
            DD.sWitch = "MS"
     '       DD.sKey = "F0009"
            DD.sKey = "F0035"
            DD.rControl.Add Item:=txt_f_addr
            DD.rControl.Add Item:=txt_o_f_addr_nm
            
            DD.nameType = "2"
            
            Call Gf_Common_DD(M_CN1, KeyCode)
            txt_o_f_addr.Text = txt_f_addr.Text
            Exit Sub
            
        End If
    
    
        If Len(Trim(txt_f_addr)) = txt_f_addr.MaxLength Then
            txt_o_f_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0035", Trim(txt_f_addr.Text), 2)
        Else
            txt_o_f_addr_nm.Text = ""
        End If
        
        If Len(Trim(txt_f_addr)) = 7 Then
           txt_o_f_addr.Text = txt_f_addr.Text
        Else
           txt_o_f_addr.Text = ""
        End If
        
    End If

End Sub

Public Function Sp_Process(Conn As ADODB.Connection, Scc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim dTempFloat As Double
    
    Dim SMESG As String
    Dim sTemp As String
    Dim ProcessChk As String
    
    Dim adoCmd As ADODB.Command

    Sp_Process = True
    
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function Or iCount = 0
    If Scc.Item("Spread").MaxRows < 1 Or Scc.Item("iColumn").Count = 0 Then
        Sp_Process = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    Scc.Item("Spread").ReDraw = False
    
    'NeceCheck
    For iCount = 1 To Scc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Scc.Item("Spread"), 0, iCount))
            
            Case "Input", "Update"
            
                If Not MC Is Nothing Then
                    Call Gp_Sp_Move(iCount, Scc, MC)
                End If
                
                'Maxlength Check
                SMESG = Gf_Sp_NeceCheck2(Scc.Item("Spread"), Scc.Item("mColumn"), iCount, Scc.Item("nColumn"))
                        
                If Trim(SMESG) = "OK" Then
                    
                ElseIf Mid(SMESG, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Scc.Item("Spread"), iCount, , vbYellow)
                    SMESG = Mid(SMESG, 6, Len(SMESG))
                    SMESG = SMESG + "长度不正确"
                    Call Gp_MsgBoxDisplay(SMESG)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Sp_Process = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Scc.Item("Spread"), iCount, , vbYellow)
                    SMESG = SMESG + "必须输入"
                    Call Gp_MsgBoxDisplay(SMESG)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Sp_Process = False
                    Exit Function
                End If
        
        End Select
    
    Next iCount
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Process = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Scc.Item("P-M")
    
    Conn.BeginTrans
    
    'Ceate Parameter (Input) iType + iColumn
    For iCount = 0 To Scc.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Ceate Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    For iCount = 1 To Scc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        
        Select Case Trim(Gf_Sp_RcvData(Scc.Item("Spread"), 0, iCount))
        
            Case "Input"
            
                adoCmd.Parameters(0).Value = "I"
                ProcessChk = "YES"
                
            Case "Update"
            
                adoCmd.Parameters(0).Value = "U"
                ProcessChk = "YES"
                
            Case "Delete"
            
                adoCmd.Parameters(0).Value = "D"
                ProcessChk = "YES"
            
        End Select
          
        If ProcessChk = "YES" Then
            
            'Parameters Setting
            For iCol = 1 To Scc.Item("iColumn").Count
            
                Scc.Item("Spread").Col = Scc.Item("iColumn").Item(iCol)
                
                Select Case Scc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Scc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempFloat = Scc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempFloat)
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Scc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempInt = Scc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempInt)
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Scc.Item("Spread").Text = "1" Then
                            adoCmd.Parameters(iCol).Value = "1"
                        Else
                            adoCmd.Parameters(iCol).Value = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Scc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = "0"
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Scc.Item("Spread").Value))
                        End If
                        
                     Case SS_CELL_TYPE_DATE
                        If Trim(Scc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Mid(Trim(Scc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Scc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Scc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Scc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).Value = Trim(sTemp)
                        
                End Select
           
            Next iCol
                           
            iProcessCount = iProcessCount + 1
            
            adoCmd.Execute
            
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Scc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Sp_Process = False
    
                Exit Function
        
             End If
        
        End If
        
    Next iCount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For iCount = 1 To Scc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Scc.Item("Spread"), 0, iCount))
        
            Case "Input", "Update"
            
                Call Gp_Sp_SendData(Scc.Item("Spread"), "", 0, iCount)
                
            Case "Delete"
                
                Call Gp_Sp_SendData(Scc.Item("Spread"), "", 0, iCount)
                Call Gp_Sp_DeleteRow(Scc.Item("Spread"), iCount)
                iCount = iCount - 1
            
        End Select
        
    Next iCount
    
    Scc.Item("Spread").ReDraw = True
    
  '  If iProcessCount > 0 Then
  '      If Not Mc Is Nothing Then
  '          If RefChek = False Then Sp_Process = Sp_Display(Conn, Sc.Item("Spread"), _
  '                                                  Gf_Ms_MakeQuery(Sc.Item("P-R"), "R", Mc("pControl")), Sc.Item("pColumn"), False)
  '      Else
  '          If RefChek = False Then Sp_Process = Sp_Display(Conn, Sc.Item("Spread"), _
  '                         Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-R"), "R", Sc.Item("aColumn"), 1), Sc.Item("pColumn"), False)
  '      End If
  '
  '      MDIMain.StatusBar1.Panels(1) = "Message : Data that handle is " & iProcessCount & " items"
  '      'Call Gp_MsgBoxDisplay("Data that handle is " & iProcessCount & " items", "I")
  '
  '  End If
            
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            Call Gp_Ms_ControlLock(MC.Item("lControl"), True)
        End If
    Else
        Sp_Process = False
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Function

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans
    
    Sp_Process = False
    
    ERR.Raise ERR.Number, ERR.Description

End Function

Private Sub txt_location1_DblClick()
    txt_t_addr.Text = txt_location1.Text
    Call Form_Ref
End Sub

Private Sub txt_location2_DblClick()
    txt_t_addr.Text = txt_location2.Text
    Call Form_Ref
End Sub

Private Sub txt_location3_DblClick()
    txt_t_addr.Text = txt_location3.Text
    Call Form_Ref
End Sub


Private Sub txt_slab_no_Change()
    Dim sQuery As String

    If Len(txt_slab_no) = 10 Then
       sQuery = "SELECT * FROM FP_SLAB WHERE SLAB_NO = '" + txt_slab_no + "'"
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
       
        MsgBox "该板坯不存在，板坯号无效！", vbCritical, "系统提示信息"
        If txt_t_addr <> "" Then
           txt_slab_no = ""
        Else
           Exit Sub
        End If
       End If
    End If
    
    If Len(txt_slab_no) = 10 And txt_f_addr <> "S0X0101" And txt_f_addr <> "S0Y0101" And S1_Click <> "1" Then
        txt_t_addr = ""
        txt_o_t_addr = ""
        txt_o_t_addr_nm = ""
    End If
End Sub

Private Sub txt_t_addr_Change()

Dim sQuery As String

    If cbo_inv.Text = "" Then
            MsgBox "请先选择仓库！", vbCritical, "系统提示信息"
            Exit Sub
    End If
    
    If Len(txt_t_addr) = 7 And Trim(txt_t_addr) <> "S0X0101" And Trim(txt_t_addr) <> "S0Y0101" Then
       sQuery = "SELECT * FROM FP_STDYARD WHERE LOCATION = '" + txt_t_addr + "' AND YARD_KND = '" + cbo_inv + "' "
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
       
          MsgBox txt_t_addr.Tag & "不正确，请重新输入！", vbCritical, "系统提示信息"
          Exit Sub
       End If
    End If
   
   If txt_f_addr <> "S0X0101" And txt_f_addr <> "S0Y0101" And Len(txt_t_addr) = 7 Then
      txt_slab_no = ""
   End If
End Sub

Private Sub txt_t_addr_DblClick()

    Call txt_t_addr_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_t_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If cbo_inv.Text = "" Then
            MsgBox "请先选择仓库！", vbCritical, "系统提示信息"
            Exit Sub
    End If
    
    If cbo_inv.Text = "XK" Then
        
        If KeyCode = vbKeyF4 Then
        
            txt_t_addr.Text = "S"
            DD.sWitch = "MS"
       '     DD.sKey = "F0009"
            DD.sKey = "F0034"
            DD.rControl.Add Item:=txt_t_addr
            DD.rControl.Add Item:=txt_o_t_addr_nm
            
            DD.nameType = "2"
            
            Call Gf_Common_DD(M_CN1, KeyCode)
            txt_o_t_addr.Text = txt_t_addr.Text
            Exit Sub
            
        End If
    
        If Len(Trim(txt_t_addr)) = txt_t_addr.MaxLength Then
            txt_o_t_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0034", Trim(txt_t_addr.Text), 2)
        Else
            txt_o_t_addr_nm.Text = ""
        End If
          
    
        If Len(Trim(txt_t_addr)) = 7 Then
           txt_o_t_addr.Text = txt_t_addr.Text
           If txt_f_addr <> "S0X0101" And txt_f_addr <> "S0Y0101" Then
              txt_slab_no = ""
           End If
        Else
           txt_o_t_addr.Text = ""
        End If
        
    ElseIf cbo_inv.Text = "52" Then
    
        If KeyCode = vbKeyF4 Then
        
            txt_t_addr.Text = "S"
            DD.sWitch = "MS"
       '     DD.sKey = "F0009"
            DD.sKey = "F0035"
            DD.rControl.Add Item:=txt_t_addr
            DD.rControl.Add Item:=txt_o_t_addr_nm
            
            DD.nameType = "2"
            
            Call Gf_Common_DD(M_CN1, KeyCode)
            txt_o_t_addr.Text = txt_t_addr.Text
            Exit Sub
            
        End If
    
        If Len(Trim(txt_t_addr)) = txt_t_addr.MaxLength Then
            txt_o_t_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0035", Trim(txt_t_addr.Text), 2)
        Else
            txt_o_t_addr_nm.Text = ""
        End If
          
    
        If Len(Trim(txt_t_addr)) = 7 Then
           txt_o_t_addr.Text = txt_t_addr.Text
           If txt_f_addr <> "S0X0101" And txt_f_addr <> "S0Y0101" Then
              txt_slab_no = ""
           End If
        Else
           txt_o_t_addr.Text = ""
        End If
        
    End If
    

End Sub

Private Sub cmd_Loc_Search_Click()
    
    Dim OutParam(3, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    If Trim(txt_slab_no.Text) = "" Then
        Call Gp_MsgBoxDisplay("必须输入板坯号", "", "错误提示")
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
        
    sQuery = "{call AFL2010P ('SL','" & Trim(txt_slab_no.Text) & "',?,?,?)}"
    
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

Private Sub ULabel3_DblClick()

    Dim sMsg As String
    Dim mResult As String
    
    If Gf_Sp_ProceExist(sc1.Item("Spread"), True) Then Exit Sub
    
    If txt_f_addr.Text <> "" Then
       sMsg = "确定对垛位（" + txt_f_addr.Text + "）进行垛层调整吗？"
       mResult = MsgBox(sMsg, vbYesNo, "系统提示信息")
       If mResult = vbYes Then
           If Gp_LOC_Exec(Trim(txt_f_addr.Text)) = "" Then
              MsgBox ("垛位调整完毕 ！"), vbInformation, "系统提示信息"
              Call Form_Ref
           Else
              MsgBox (" 垛位调整失败！"), vbCritical, "系统提示信息"
           End If
       End If
       Exit Sub
    End If

End Sub

Public Function Gp_LOC_Exec(sAddr As String) As String

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

    sQuery = "{call ACB6060C.P_LOC_TUN ('" + cbo_inv + "','" + sAddr + "',?)}"

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
    ERR.Raise ERR.Number, ERR.Description & sQuery

End Function


Private Sub ULabel6_DblClick()

    Dim sMsg As String
    Dim mResult As String
    
    If Gf_Sp_ProceExist(sc2.Item("Spread"), True) Then Exit Sub
    
    If txt_t_addr.Text <> "" Then
       sMsg = "确定对垛位（" + txt_t_addr.Text + "）进行垛层调整吗？"
       mResult = MsgBox(sMsg, vbYesNo, "系统提示信息")
       If mResult = vbYes Then
           If Gp_LOC_Exec(Trim(txt_t_addr.Text)) = "" Then
              MsgBox ("垛位调整完毕 ！"), vbInformation, "系统提示信息"
              Call Form_Ref
           Else
              MsgBox (" 垛位调整失败！"), vbCritical, "系统提示信息"
           End If
       End If
       Exit Sub
    End If


End Sub


