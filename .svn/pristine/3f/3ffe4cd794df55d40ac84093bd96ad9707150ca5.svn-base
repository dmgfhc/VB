VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGA2010C 
   BackColor       =   &H80000013&
   Caption         =   "板坯库库存修改及查询界面_CGA2010C"
   ClientHeight    =   9480
   ClientLeft      =   135
   ClientTop       =   1440
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   14610
   WindowState     =   2  'Maximized
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
      ItemData        =   "CGA2010C.frx":0000
      Left            =   2280
      List            =   "CGA2010C.frx":000A
      TabIndex        =   20
      Tag             =   "连铸机号"
      Top             =   150
      Visible         =   0   'False
      Width           =   615
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
      Left            =   11715
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   18
      Top             =   135
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
      Left            =   12870
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   17
      Top             =   135
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
      Left            =   14040
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   16
      Top             =   135
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1050
      Left            =   135
      TabIndex        =   3
      Top             =   495
      Width           =   15075
      Begin VB.OptionButton opt_Left_Right 
         BackColor       =   &H00E0E0E0&
         Caption         =   "左边->右边"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   690
         Width           =   1410
      End
      Begin Threed.SSFrame SSFrame1 
         Height          =   345
         Left            =   1650
         TabIndex        =   22
         Top             =   630
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   609
         _Version        =   196609
         BackColor       =   14737632
         Begin VB.OptionButton opt_sequence 
            BackColor       =   &H00E0E0E0&
            Caption         =   "一列"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   60
            Width           =   840
         End
         Begin VB.OptionButton opt_sequence 
            BackColor       =   &H00E0E0E0&
            Caption         =   "二列"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   870
            TabIndex        =   26
            Top             =   60
            Width           =   840
         End
         Begin VB.OptionButton opt_sequence 
            BackColor       =   &H00E0E0E0&
            Caption         =   "三列"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   25
            Top             =   60
            Width           =   840
         End
         Begin VB.OptionButton opt_sequence 
            BackColor       =   &H00E0E0E0&
            Caption         =   "四列"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   2490
            TabIndex        =   24
            Top             =   60
            Width           =   840
         End
      End
      Begin VB.TextBox txt_sequence 
         Height          =   330
         Left            =   5070
         TabIndex        =   21
         Text            =   " "
         Top             =   660
         Width           =   465
      End
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
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   2460
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
         Left            =   9765
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   12285
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   2640
      End
      Begin VB.OptionButton opt_Right_Left 
         BackColor       =   &H00E0E0E0&
         Caption         =   "左边<-右边 "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13350
         TabIndex        =   6
         Top             =   735
         Width           =   1515
      End
      Begin VB.TextBox txt_slab_cnt 
         Height          =   330
         Left            =   6120
         TabIndex        =   5
         Text            =   " "
         Top             =   615
         Width           =   465
      End
      Begin VB.TextBox txt_p_row 
         Enabled         =   0   'False
         Height          =   330
         Left            =   8190
         TabIndex        =   4
         Text            =   " "
         Top             =   615
         Width           =   465
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   2790
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "垛位名称"
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
         Left            =   90
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "起始垛位号"
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   11100
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "垛位名称"
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
         Left            =   8355
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "目的垛位号"
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
      Begin Threed.SSCommand ssc_can 
         Height          =   330
         Left            =   7395
         TabIndex        =   11
         Top             =   615
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         Enabled         =   0   'False
         Caption         =   "&取消"
      End
      Begin Threed.SSCommand ssc_move 
         Height          =   330
         Left            =   6600
         TabIndex        =   12
         Top             =   615
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   255
         Enabled         =   0   'False
         Caption         =   "&移动"
      End
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
      Left            =   4680
      MaxLength       =   7
      TabIndex        =   1
      Top             =   135
      Width           =   975
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
      Left            =   1320
      MaxLength       =   7
      TabIndex        =   0
      Top             =   150
      Width           =   975
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
      Left            =   7125
      MaxLength       =   10
      TabIndex        =   2
      Top             =   135
      Width           =   1455
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   5910
      Top             =   135
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Caption         =   "板坯号"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   105
      Top             =   135
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Caption         =   "起始垛位号"
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
      Left            =   3480
      Top             =   135
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Caption         =   "目的垛位号"
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
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7650
      Left            =   105
      TabIndex        =   13
      Top             =   1575
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   13494
      _Version        =   196609
      PaneTree        =   "CGA2010C.frx":0014
      Begin FPSpread.vaSpread ss1 
         Height          =   7590
         Left            =   30
         TabIndex        =   14
         Top             =   30
         Width           =   7215
         _Version        =   393216
         _ExtentX        =   12726
         _ExtentY        =   13388
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
         MaxCols         =   19
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGA2010C.frx":0066
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   7590
         Left            =   7335
         TabIndex        =   15
         Top             =   30
         Width           =   7740
         _Version        =   393216
         _ExtentX        =   13653
         _ExtentY        =   13388
         _StockProps     =   64
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CGA2010C.frx":1E9E
      End
   End
   Begin Threed.SSCommand cmd_Loc_Search 
      Height          =   315
      Left            =   10650
      TabIndex        =   19
      Top             =   135
      Width           =   1050
      _ExtentX        =   1852
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
End
Attribute VB_Name = "CGA2010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       中板系统
'-- Sub_System Name   板坏库管理
'-- Program Name      Yard Position change
'-- Program ID        CGA2010C
'-- Document No
'-- Designer          Shin.c.s
'-- Coder             Shin.c.s
'-- Date              2007.7.26
'-- Description       Yard Position change
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
Public Active_CForm As String       'Form Active CGA2010c
Public Active_LForm As String       'Form Active CGA2010c
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
Dim sCross_Seq As String

Private Sub Form_Define()
        
   'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_f_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_t_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(cbo_ccm_line, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, " ", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    Call Gp_Sp_Collection(ss2, 1, " ", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, False)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, False)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2, False)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGA2010C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="CGA2010C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGA2010C.P_MODIFY", Key:="P-M"
    sc2.Add Item:="CGA2010C.P_REFER2", Key:="P-R"
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
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "◎"
    
    Call Gp_Sp_ColHidden(ss1, 19, True)
    Call Gp_Sp_ColHidden(ss2, 19, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub


Private Sub Form_Activate()
    If Active_LForm = "1" Then
       Call Form_Ref
       Active_LForm = ""
    End If
    
'    ss1.MaxRows = Max_Rows
'    ss1.Row = Max_Rows
'    ss1.Col = 1
'    ss1.Action = ActionActiveCell
'
'    ss2.MaxRows = Max_Rows
'    ss2.Row = Max_Rows
'    ss2.Col = 1
'    ss2.Action = ActionActiveCell
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(7).Enabled = False
    txt_o_f_addr.Text = txt_f_addr.Text
    txt_o_f_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0033", Trim(txt_f_addr.Text), 2)
    
    txt_o_t_addr.Text = txt_t_addr.Text
    txt_o_t_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0033", Trim(txt_t_addr.Text), 2)
    
    
    opt_sequence(0).Enabled = False
    opt_sequence(1).Enabled = False
    opt_sequence(2).Enabled = False
    opt_sequence(3).Enabled = False
         
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Max_Rows = 200
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    MDIMain.MenuTool.Buttons(12).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(7).Enabled = False
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    sc1.Item("Spread").RetainSelBlock = False
    sc2.Item("Spread").RetainSelBlock = False
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    sChkFlag = "ON"
    
    opt_Left_Right.Value = True
    Call opt_Left_Right_Click
    
    Screen.MousePointer = vbDefault
              
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
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
                ss1.ROW = Max_Rows
                ss1.Col = 1
                ss1.Action = ActionActiveCell
                
                ss2.MaxRows = Max_Rows
                ss2.ROW = Max_Rows
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
    Dim sMsg, sMesg, sTemp, sQuery As String
    
    If Len(Trim(txt_f_addr)) = 7 Then
       sQuery = "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'F0033' AND CD = '" & txt_f_addr.Text & "'"
        If Gf_FloatFind(M_CN1, sQuery) = 0 Then
           MsgBox txt_f_addr.Tag & "不正确，请重新输入！", vbCritical, "系统提示信息"
           Exit Sub
        End If
    End If
    
    If Len(Trim(txt_t_addr)) = 7 Then
       sQuery = "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'F0033' AND CD = '" & txt_t_addr.Text & "'"
        If Gf_FloatFind(M_CN1, sQuery) = 0 Then
           MsgBox txt_t_addr.Tag & "不正确，请重新输入！", vbCritical, "系统提示信息"
           Exit Sub
        End If
    End If
    
    If opt_Left_Right Then
        sFromLoc = txt_f_addr
        sToLoc = txt_t_addr
    ElseIf opt_Right_Left Then
        sFromLoc = txt_t_addr
        sToLoc = txt_f_addr
    Else
        sFromLoc = txt_f_addr
        sToLoc = txt_t_addr
    End If
    
    If sFromLoc <> "" And sToLoc <> "" Then
       
        If sFromLoc = sToLoc Then
           sMsg = "起始垛位号和目的垛位号相同！请重新选择起始垛位号和目的垛位号！"
           GoTo Refer_Err
        ElseIf Mid(sFromLoc, 1, 2) <> "S0" Then
           If TXT_SLAB_NO <> "" Then
              sMsg = "不是在线板坯入库，板坯号和目的垛位号不能同时输入！"
              GoTo Refer_Err
           End If
        End If
    End If
            
'    If txt_f_addr = "S0L0101" Or txt_f_addr = "S0C0101" Or txt_f_addr = "S0Q0101" Then
'       Call Gp_Sp_ColHidden(ss1, 1, True)
'       Call Gp_Sp_ColHidden(ss1, 2, True)
'    Else
'       Call Gp_Sp_ColHidden(ss1, 1, False)
'       Call Gp_Sp_ColHidden(ss1, 2, False)
'    End If
    
'    If txt_t_addr = "S0L0101" Or txt_t_addr = "S0C0101" Or txt_t_addr = "S0Q0101" Then
'       Call Gp_Sp_ColHidden(ss2, 1, True)
'       Call Gp_Sp_ColHidden(ss2, 2, True)
'    Else
'       Call Gp_Sp_ColHidden(ss2, 1, False)
'       Call Gp_Sp_ColHidden(ss2, 2, False)
'    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    ss1.MaxRows = 0
    ss2.MaxRows = 0
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
    
         sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then
       
            If Gf_Sp_Refer(M_CN1, sc1, Mc1, Nothing, Nothing, False) Then
                sc1.Item("Spread").OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                MDIMain.MenuTool.Buttons(11).Enabled = False
                MDIMain.MenuTool.Buttons(12).Enabled = False
                MDIMain.MenuTool.Buttons(7).Enabled = False
                MaxCnt = ss1.MaxRows
                If Mid(sFromLoc, 1, 2) = "S0" Then
                      'ss1.MaxRows = MaxCnt
                      ss1.MaxRows = ss1.MaxRows
                      Max_Rows = ss1.MaxRows
                Else
                   sQuery = "SELECT COUNT(*) FROM FP_SLABYARD WHERE YARD_ADDR = '" + txt_f_addr + "' AND YARD_KND = 'ZB'"
                   Max_Rows = Gf_FloatFind(M_CN1, sQuery)
                   
                   'Max_Rows = 200
                   ss1.MaxRows = Max_Rows
                End If
                
                ss1.ROW = Max_Rows
                ss1.Col = 1
                ss1.Action = ActionActiveCell
                    
                With ss1
                     For iRow = MaxCnt To 1 Step -1
                         .ROW = iRow
                         For iCol = 1 To .MaxCols
                             .Col = iCol
                             sTemp = .Text
                             .Text = ""
                             .ROW = .ROW + Max_Rows - MaxCnt
                             .Text = sTemp
                             .ROW = iRow
                         Next iCol
                     Next iRow
                                         
                    If (Mid(sFromLoc, 1, 2) = "S0" Or Trim(sFromLoc) = "") And TXT_SLAB_NO <> "" Then
                       .Col = 3
                       For iRow = Max_Rows - MaxCnt + 1 To Max_Rows
                           .ROW = iRow
                           If .Text = TXT_SLAB_NO Then
                              .Col = 0
                              .Text = "Delete"
                              iFromRow = .ROW
                              iMoveCnt = 1
                              txt_slab_cnt = 1
                              For iCol = 1 To 17
                                  .Col = iCol
                                  .BackColor = &HFF
                              Next iCol
                              
                              Exit For
                           End If
                       Next iRow
                    End If
                End With
                Call Spread_Color_Set(ss1)
            Else
                
                sQuery = "SELECT COUNT(*) FROM FP_SLABYARD WHERE YARD_ADDR = '" + txt_f_addr + "' AND YARD_KND = 'ZB'"
                Max_Rows = Gf_FloatFind(M_CN1, sQuery)
                sc1.Item("Spread").MaxRows = Max_Rows
                sc1.Item("Spread").ROW = Max_Rows
                sc1.Item("Spread").Col = 1
                sc1.Item("Spread").Action = ActionActiveCell
                If txt_f_addr <> "" And Mid(txt_f_addr, 1, 2) <> "S0" And sFromLoc = txt_f_addr Then
                   MsgBox "垛位 " + sFromLoc + " 没有板坯！", vbInformation, "系统提示信息"
                ElseIf Mid(txt_f_addr, 1, 2) = "S0" And sFromLoc = txt_f_addr Then
                   MsgBox "没有在线板坯等待入库！", vbInformation, "系统提示信息"
                End If
            End If
            
            
            If Gf_Sp_Refer(M_CN1, sc2, Mc1, Nothing, Nothing, False) Then
                sc2.Item("Spread").OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                MDIMain.MenuTool.Buttons(12).Enabled = False
                MDIMain.MenuTool.Buttons(11).Enabled = False
                MDIMain.MenuTool.Buttons(7).Enabled = False
                MaxCnt = ss2.MaxRows
                
                If Mid(sToLoc, 1, 2) = "S0" Then
                   'sc2.Item("Spread").MaxRows = MaxCnt
                   sc2.Item("Spread").MaxRows = ss2.MaxRows * 2
                   Max_Rows = ss2.MaxRows
                Else
                   sQuery = "SELECT COUNT(*) FROM FP_SLABYARD WHERE YARD_ADDR = '" + txt_t_addr + "' AND YARD_KND = 'ZB'"
                   Max_Rows = Gf_FloatFind(M_CN1, sQuery)
                   'Max_Rows = 200
                   sc2.Item("Spread").MaxRows = Max_Rows
                End If
                
                sc2.Item("Spread").ROW = Max_Rows
                sc2.Item("Spread").Col = 1
                sc2.Item("Spread").Action = ActionActiveCell
                
                With ss2
                     For iRow = MaxCnt To 1 Step -1
                         .ROW = iRow
                         For iCol = 1 To .MaxCols
                             .Col = iCol
                             sTemp = .Text
                             .Text = ""
                             .ROW = .ROW + Max_Rows - MaxCnt
                             .Text = sTemp
                             .ROW = iRow
                         Next iCol
                     Next iRow
  
                    If TXT_SLAB_NO <> "" Then
'                         sQuery = "SELECT * FROM FP_SLABYARD WHERE SLAB_NO = '" + txt_slab_no + "'"
'                         If Gf_FloatFind(M_CN1, sQuery) = 0 Then
'                            .Row = Max_Rows - MaxCnt
'                            .Col = 0
'                            .Text = "Input"
'                            .Col = 1
'                            .Row = Max_Rows - MaxCnt + 1
'                             iStemp = .Text
'                            .Row = Max_Rows - MaxCnt
'
'                             iMoveCnt = 1
'                             iToStaRow = .Row
'                             txt_slab_cnt = 1
'
'                            .Text = iStemp + 1
'                            .Col = 3
'                            .Text = txt_slab_no.Text
'                            .Col = 4
'                            .Text = sToLoc
'                            .Col = 16
'                            .Text = sFromLoc
'                            ssc_can.Enabled = True
'
'                             For iCol = 1 To 14
'                                .Col = iCol
'                                .BackColor = &HFF
'                             Next iCol
'
'                             MDIMain.MenuTool.Buttons(4).Enabled = True
'                         Else
                            
                            If sToLoc = "" Then
                               .Col = 3
                                For iRow = Max_Rows To Max_Rows - MaxCnt + 1 Step -1
                                    .ROW = iRow
                                     If .Text = TXT_SLAB_NO Then
                                        .SetSelection 2, .ROW, 2, .ROW
                                        .ForeColor = &HFF
                                     End If
                                Next iRow
                            End If
                            
                            If Mid(sFromLoc, 1, 2) = "S0" Then
                               MsgBox "板坯 " + TXT_SLAB_NO + " 已经在库中，不需再做入库处理！", vbInformation, "系统提示信息"
                               TXT_SLAB_NO = ""
                            End If
                            
                         '   Exit Sub
                         'End If
                    End If
                End With
                Call Spread_Color_Set(ss2)
            Else
                sQuery = "SELECT COUNT(*) FROM FP_SLABYARD WHERE YARD_ADDR = '" + txt_t_addr + "' AND YARD_KND = 'ZB'"
                Max_Rows = Gf_FloatFind(M_CN1, sQuery)
                sc2.Item("Spread").MaxRows = Max_Rows
                sc2.Item("Spread").ROW = Max_Rows
                sc2.Item("Spread").Col = 1
                sc2.Item("Spread").Action = ActionActiveCell
                
                If TXT_SLAB_NO <> "" And (Mid(sFromLoc, 1, 2) = "S0" Or Trim(sFromLoc) = "") Then
                   sQuery = "SELECT * FROM FP_SLABYARD WHERE SLAB_NO = " + TXT_SLAB_NO
                   If Gf_FloatFind(M_CN1, sQuery) = 0 Then

                        With ss2
                             .ROW = Max_Rows
                             .Col = 0
                             .Text = "Input"
                             .Col = 1
                             .Text = "1"
                             .Col = 3
                             .Text = TXT_SLAB_NO.Text
                             .Col = 4
                             .Text = sToLoc
                             .Col = 19
                             .Text = sFromLoc
                             
                             iMoveCnt = 1
                             iToStaRow = .ROW
                             txt_slab_cnt = 1
                             
                             For iCol = 1 To 16
                                .Col = iCol
                                .BackColor = &HFF
                             Next iCol
                             
                             MDIMain.MenuTool.Buttons(4).Enabled = True
                        End With
                        ssc_can.Enabled = True
                   Else
                        MsgBox "板坯 " + TXT_SLAB_NO + " 已经在库中，不需再做入库处理！", vbInformation, "系统提示信息"
                        TXT_SLAB_NO = ""
                        Exit Sub
                   End If
                
                Else
                   If sToLoc <> "" And Mid(sToLoc, 1, 2) <> "S0" Then
                      MsgBox "垛位 " + sToLoc + " 没有板坯！", vbInformation, "系统提示信息"
                   ElseIf Mid(sToLoc, 1, 2) = "S0" Then
                      MsgBox "没有在线板坯等待入库！", vbInformation, "系统提示信息"
                   End If
                End If

            End If
                        
        Else
            sMesg = sMesg + "长度不正确"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    Else
        sMesg = sMesg + "必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
        
    End If
    
    If TXT_SLAB_NO = "" Then
       txt_slab_cnt = ""
       ssc_move.Enabled = False
           opt_sequence(0).Enabled = False
           opt_sequence(1).Enabled = False
           opt_sequence(2).Enabled = False
           opt_sequence(3).Enabled = False
       ssc_can.Enabled = False
       txt_p_row = ""
    End If
    
    
    If iToStaRow = 0 And CGA2010C.Active_CForm <> "" Then
        If CGA2010C.Active_CForm <> "" Then
            iFromRow = txt_p_row
            If Trim(txt_slab_cnt) <> "" Then
               iMoveCnt = CInt(txt_slab_cnt)
            End If
        
            For iCnt = Max_Rows To 1 Step -1
                ss2.Col = 1
                ss2.ROW = iCnt
                If ss2.Text = "" Then
                   iToStaRow = iCnt + 1
                   iCnt = 1
                   Exit For
                End If
            Next iCnt
        
            ss2.ROW = iToStaRow - iMoveCnt + 1
            ss2.Col = 1
            To_Bedseq = Format(ss2.Text, "0#")
            
            sQuery = "SELECT MAX_CNT FROM FP_STDYARD WHERE LOCATION ='" + sToLoc + "'"
            ssc_move.Enabled = False
            opt_sequence(0).Enabled = False
            opt_sequence(1).Enabled = False
            opt_sequence(2).Enabled = False
            opt_sequence(3).Enabled = False
            If ss2.Text <> "" Then
                If Mid(sToLoc, 1, 2) <> "S0" And CInt(To_Bedseq) > Gf_FloatFind(M_CN1, sQuery) Then
                   MsgBox "已超出目的垛位的存放能力！当前操作无法继续！", vbCritical, "系统提示信息"
                   Call ssc_can_Click
                   Exit Sub
                End If
            End If
    
         End If
     End If
     
     Exit Sub

Refer_Err:

    Call Gp_MsgBoxDisplay(sMsg)
 
End Sub

Public Sub Form_Pro()
    Dim SlabNo      As String
    Dim i           As Integer
    Dim sQuery      As String
    Dim sFromLoc    As String

    TXT_SLAB_NO = ""
    
If opt_Right_Left.Value = True Then
    If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       For i = 1 To ss2.MaxRows
           Call Gp_Sp_SendData(ss2, "", 0, i)
       Next i
    End If
ElseIf opt_Left_Right.Value = True Then
    
    If Gf_Sp_Process(M_CN1, sc2, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        MDIMain.MenuTool.Buttons(12).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(7).Enabled = False
        For i = 1 To ss1.MaxRows
            Call Gp_Sp_SendData(ss1, "", 0, i)
        Next i
        SlabNo = TXT_SLAB_NO
        TXT_SLAB_NO = ""
        
        If opt_Left_Right Then
            sFromLoc = txt_f_addr
        Else
            sFromLoc = txt_t_addr
        End If
        
        If Mid(sFromLoc, 1, 2) = "S0" Then
            MDIMain.StatusBar1.Panels(1).Text = "板坯 " + SlabNo + " 已成功入库！"
        Else
            MDIMain.StatusBar1.Panels(1).Text = "您所选板坯的库位已成功变更！"
        End If
        S1_Click = ""
    End If
End If
    
    Call Form_Ref
End Sub


Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
End Sub
Public Sub Spread_Color_Set(MapSpr As Variant)
Dim iCnt As Long
Dim tmFromRow As Long
Dim tmToRow As Long
Dim tmFromVal As String
Dim tmToVal As String
Dim ColFlag As String

    'Dan sequnence check
    MapSpr.ROW = MapSpr.MaxRows
    MapSpr.Col = 1
    tmFromVal = MapSpr.Text
    tmFromRow = MapSpr.MaxRows
    ColFlag = "0"
    If tmFromVal = "" Then Exit Sub
    
    For iCnt = MapSpr.MaxRows To 1 Step -1
        MapSpr.ROW = iCnt
        MapSpr.Col = 1
        If Trim(MapSpr.Text) <> "" Then
            If MapSpr.Text <> tmFromVal Then
               tmToRow = iCnt + 1
               
                MapSpr.Col = 1: MapSpr.Col2 = MapSpr.MaxCols
                MapSpr.ROW = tmToRow: MapSpr.Row2 = tmFromRow
                
                If ColFlag = "0" Then
                    MapSpr.BlockMode = True
                    MapSpr.ForeColor = &H0&
                    MapSpr.BackColor = &H80000013
                    MapSpr.BlockMode = False
                    
                    ColFlag = "1"
                Else
                    ColFlag = "0"
                End If
                tmFromRow = iCnt
                MapSpr.ROW = iCnt
                MapSpr.Col = 1
                tmFromVal = MapSpr.Text
            End If
        Else
                MapSpr.Col = 1: MapSpr.Col2 = MapSpr.MaxCols
                MapSpr.ROW = iCnt + 1: MapSpr.Row2 = tmFromRow
                
                If ColFlag = "0" Then
                    MapSpr.BlockMode = True
                    MapSpr.ForeColor = &H0&
                    MapSpr.BackColor = &H80000013
                    MapSpr.BlockMode = False
                    
                    ColFlag = "1"
                Else
                    ColFlag = "0"
                End If
                Exit For
       End If
    Next
        

    
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
    ss1.Col = 0
    ss1.ROW = 0
    
    If ss1.Text = "◎" Then
        Call Gp_Sp_Del(Proc_Sc("Sc"))
    Else
        Call Gp_Sp_Del(Proc_Sc("Sc2"))
    End If
    TXT_SLAB_NO = ""
End Sub

Public Sub Spread_Can()
    ss1.Col = 0
    ss1.ROW = 0
    If ss1.Text = "◎" Then
        Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
    Else
        Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc2"))
    End If
End Sub

Private Sub opt_Left_Right_Click()
Dim sQuery As String

    If sChkFlag = "ON" Then
        If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
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
        ss1.ROW = 0
        ss1.Text = "◎"
        ss2.Col = 0
        ss2.ROW = 0
        ss2.Text = ""
        
        ss1.Enabled = True
        ss2.Enabled = False
    End If
    
    If Len(txt_t_addr) = 7 And Mid(Trim(txt_t_addr), 1, 2) <> "S0" Then
       sQuery = "SELECT MAX(CROSS_SEQ) FROM FP_SLABYARD WHERE YARD_ADDR = '" + txt_t_addr + "' AND YARD_KND = 'ZB'"
       sCross_Seq = Gf_FloatFind(M_CN1, sQuery)
       If sCross_Seq = "0" Then sCross_Seq = "2"
       opt_sequence(CInt(sCross_Seq) - 1).Value = True
   End If
End Sub

Private Sub opt_Right_Left_Click()
Dim sQuery As String

    If sChkFlag = "ON" Then
        If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
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
        ss1.ROW = 0
        ss1.Text = ""
        ss2.Col = 0
        ss2.ROW = 0
        ss2.Text = "◎"
        
        ss1.Enabled = False
        ss2.Enabled = True
    End If
    
    If Len(txt_f_addr) = 7 And Mid(Trim(txt_f_addr), 1, 2) <> "S0" Then
       sQuery = "SELECT MAX(CROSS_SEQ) FROM FP_SLABYARD WHERE YARD_ADDR = '" + txt_f_addr + "' AND YARD_KND = 'ZB'"
       sCross_Seq = Gf_FloatFind(M_CN1, sQuery)
       If sCross_Seq = "0" Then sCross_Seq = "2"
       opt_sequence(CInt(sCross_Seq) - 1).Value = True
   End If
   
End Sub

Private Sub rowEdit()
    TXT_SLAB_NO = ""
End Sub

Private Sub opt_sequence_Click(Index As Integer)
    txt_sequence = Index
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Change(ByVal Col As Long, ByVal ROW As Long)
    
    If opt_Right_Left Then Exit Sub
    Call ssc_Upd_Process(ss1, txt_f_addr, ROW)
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If opt_Right_Left Then Exit Sub
    If Mode <> 1 Then
       Call ssc_Upd_Process(ss1, txt_f_addr, ROW)
    End If
End Sub

Private Sub ssc_Upd_Process(oSpr As vaSpread, sText As Variant, ByVal ROW As Long)

    Dim iCurrRowVal As Integer

    With oSpr

        If Gf_Sc_Authority(sAuthority, "U") Then
            
            .ROW = ROW
            .Col = 0
            .Text = "Update"
                
            If ROW = .MaxRows Then
                .Col = 1
                .Value = 1
            Else
                .ROW = ROW + 1
                .Col = 1
                iCurrRowVal = Val(.Value & "")
                
                .ROW = ROW
                .Value = iCurrRowVal + 1
            End If
            
            .Col = 5
            .Text = Trim(sText.Text)
            .Col = 18
            .Text = sUserID
        End If
    
    End With
End Sub


Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
        
    Dim iCnt, i As Integer
    Dim sFlag As String
    Dim ACTROW As Integer
    
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If opt_Right_Left Then Exit Sub
    
    ACTROW = ROW
    With ss1
       .Col = 0
       For i = Max_Rows To 1 Step -1
           .ROW = i
           If .Text = "Delete" Then
              sFlag = "Y"
              Exit For
           End If
       Next i
       
        .Col = 3
        .ROW = ROW
        
        If (.Text <> "") And sFlag <> "Y" Then
           S1_Click = "1"
           TXT_SLAB_NO = .Text
           
        ElseIf (.Text = "") And sFlag <> "Y" Then
           TXT_SLAB_NO = ""
           txt_slab_cnt = ""
           txt_p_row = ROW
           ssc_move.Enabled = False
           ssc_can.Enabled = False
           opt_sequence(0).Enabled = False
           opt_sequence(1).Enabled = False
           opt_sequence(2).Enabled = False
           opt_sequence(3).Enabled = False
        End If
        
        If sFlag <> "Y" Then
           txt_p_row = ROW
        End If
        
        If Mid(txt_f_addr, 1, 2) <> "S0" And sFlag <> "Y" Then
            For iCnt = ROW To 1 Step -1
                .Col = 5
                .ROW = iCnt
                If .Text <> "" Then
                   i = i + 1
                ElseIf .Text = "" Then
                   
                   If .ROW <> Max_Rows Then
                      .ROW = .ROW + 1
                   End If
                   
                   TopSlabNo = .Text
                   TopSlabRow = .ROW
                   Exit For
                End If
            Next iCnt
            
            iCnt = 0
            For i = ACTROW To 1 Step -1
                ss1.ROW = i
                ss1.Col = 3
                If ss1.Text <> "" Then
                   iCnt = iCnt + 1
                Else
                   txt_slab_cnt = iCnt
                   Exit For
                End If
                    
                If i = 1 Then
                   txt_slab_cnt = iCnt
                End If
            Next
'            If I <> 0 Then
'               txt_slab_cnt = I
'            End If
        
        ElseIf Mid(txt_f_addr, 1, 2) = "S0" Then

            txt_slab_cnt = 1
            txt_p_row = .ActiveRow
            .ROW = .ActiveRow
            .Col = 3
             TXT_SLAB_NO = .Text

        End If
        
        If TXT_SLAB_NO <> "" And TXT_SLAB_NO <> "0" And sFlag <> "Y" Then
           ssc_move.Enabled = True
           opt_sequence(0).Enabled = True
           opt_sequence(1).Enabled = True
           opt_sequence(2).Enabled = True
           opt_sequence(3).Enabled = True
        End If
    
    End With

    If txt_slab_cnt <> "" And txt_slab_cnt <> "0" Then
       ssc_can.Enabled = True
    End If
     
    Exit Sub

ss1_Click_error:
    Call Gp_MsgBoxDisplay(" Not allowed Select Row", "I")
    
 End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
   TXT_SLAB_NO = ""
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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

Private Sub ss2_Change(ByVal Col As Long, ByVal ROW As Long)
    
    If opt_Left_Right Then Exit Sub
    Call ssc_Upd_Process(ss2, txt_t_addr, ROW)
    
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If opt_Left_Right Then Exit Sub
    Call ssc_Upd_Process(ss2, txt_t_addr, ROW)
    
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)
    
    Dim iCnt, i As Integer
    Dim sFlag As String
    Dim ACTROW As Integer
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If opt_Left_Right Then Exit Sub
   
    ACTROW = ROW
    With ss2
       .Col = 0
       For i = Max_Rows To 1 Step -1
           .ROW = i
           If .Text = "Delete" Then
              sFlag = "Y"
              Exit For
           End If
       Next i
       
        .Col = 3
        .ROW = ROW
        
        If (.Text <> "") And sFlag <> "Y" Then
           S1_Click = "1"
           TXT_SLAB_NO = .Text
           
        ElseIf (.Text = "") And sFlag <> "Y" Then
           TXT_SLAB_NO = ""
           txt_slab_cnt = ""
           txt_p_row = ROW
           ssc_move.Enabled = False
           opt_sequence(0).Enabled = False
           opt_sequence(1).Enabled = False
           opt_sequence(2).Enabled = False
           opt_sequence(3).Enabled = False
           ssc_can.Enabled = False
        End If
        
        If sFlag <> "Y" Then
           txt_p_row = ROW
        End If
        
        If Mid(txt_t_addr, 1, 2) <> "S0" And sFlag <> "Y" Then
            For iCnt = ROW To 1 Step -1
                .Col = 3
                .ROW = iCnt
                If .Text <> "" Then
                   i = i + 1
                ElseIf .Text = "" Then
                   
                   If .ROW <> Max_Rows Then
                      .ROW = .ROW + 1
                   End If
                   
                   TopSlabNo = .Text
                   TopSlabRow = .ROW
                   Exit For
                End If
            Next iCnt
            
'            If I <> 0 Then
'               txt_slab_cnt = I
'            End If
            
            iCnt = 0
            For i = ACTROW To 1 Step -1
                ss2.ROW = i
                ss2.Col = 3
                If ss2.Text <> "" Then
                   iCnt = iCnt + 1
                Else
                   txt_slab_cnt = iCnt
                   Exit For
                End If
                    
                If i = 1 Then
                   txt_slab_cnt = iCnt
                End If
            Next
        
        ElseIf Mid(txt_t_addr, 1, 2) = "S0" Then
        
            txt_slab_cnt = 1
            txt_p_row = .ActiveRow
            .ROW = .ActiveRow
            .Col = 3
             TXT_SLAB_NO = .Text
           
        End If
        
        If TXT_SLAB_NO <> "" And TXT_SLAB_NO <> "0" And sFlag <> "Y" Then
           ssc_move.Enabled = True
           opt_sequence(0).Enabled = True
           opt_sequence(1).Enabled = True
           opt_sequence(2).Enabled = True
           opt_sequence(3).Enabled = True
        End If
    
    End With

    If txt_slab_cnt <> "" And txt_slab_cnt <> "0" Then
       ssc_can.Enabled = True
    End If
     
    Exit Sub

ss2_Click_error:
    Call Gp_MsgBoxDisplay(" Not allowed Select Row", "I")
    

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal ROW As Long)
   TXT_SLAB_NO = ""
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

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If ROW > 0 Then
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
'If Trim(txt_t_addr.Text) = "S0C0101" Or Trim(txt_t_addr.Text) = "S0L0101" Or Trim(txt_t_addr.Text) = "S0Q0101" Then
'   MsgBox "目的垛位号不正确！", vbCritical, "系统提示"
'   Exit Sub
'End If
Dim sQuery As String
    If Len(txt_t_addr) = 7 And Mid(Trim(txt_t_addr), 1, 2) <> "S0" Then
       sQuery = "SELECT * FROM FP_STDYARD WHERE LOCATION = '" + txt_t_addr + "' AND YARD_KND = 'ZB'"
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
       
          MsgBox txt_t_addr.Tag & "目的垛位号不正确，请重新输入！", vbCritical, "系统提示信息"
          Exit Sub
       End If
    End If


    If opt_Left_Right Then
        Call ssc_move_Process(ss1, ss2)
    Else
        Call ssc_move_Process(ss2, ss1)
    End If

End Sub

Private Sub ssc_can_Process(oSpr1 As vaSpread, oSpr2 As vaSpread)

    Dim i As Integer
    Dim iCnt As Integer

    oSpr2.SetSelection 1, iToStaRow - iMoveCnt + 1, 16, iToStaRow
    oSpr2.ClipboardCut
  
    With oSpr1
      
      For iCnt = iFromRow - iMoveCnt + 1 To iFromRow Step 1
         .ROW = iCnt
         .Col = 0
         .Text = ""
         For i = 1 To 5
          .Col = i
          .BackColor = &HC0FFFF
         Next i

         For i = 6 To .MaxCols
          .Col = i
          .BackColor = &HFFFFFF
         Next i
      Next
    End With
    Call Spread_Color_Set(oSpr1)

    With oSpr2

      For iCnt = iToStaRow To iToStaRow - iMoveCnt + 1 Step -1
         .ROW = iCnt
         .Col = 0
         .Text = ""
         For i = 1 To 5
          .Col = i
          .BackColor = &HC0FFFF
         Next i
         
         For i = 3 To .MaxCols
            If i <> 5 Then
             .Col = i
             .Text = ""
             End If
         Next i
         
         For i = 6 To .MaxCols
          .Col = i
          .BackColor = &HFFFFFF
         Next
      Next
    End With
    Call Spread_Color_Set(oSpr2)
    
    S1_Click = ""
    txt_p_row = ""
    TXT_SLAB_NO = ""
    txt_slab_cnt = ""
    To_Bedseq = ""
    iToStaRow = 0
    iMoveCnt = 0
    
    'oSpr2.MaxRows = Max_Rows
    'oSpr2.ROW = Max_Rows
    'oSpr2.Col = 1
    'oSpr2.Action = ActionActiveCell
    
    ssc_can.Enabled = False

         
End Sub

Private Sub ssc_move_Process(oSpr1 As vaSpread, oSpr2 As vaSpread)

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
    Dim sMsg, sSeq, sQuery As String
    Dim sTempSlabNo As String
    Dim sBedSeq As String
    Dim sFromLoc As String
    Dim sToLoc   As String
    Dim tmpSeqNo As Integer
    Dim fromCnt As Integer
    Dim toCnt As Integer
    
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
    
    toCnt = 0
    For i = 1 To oSpr2.MaxRows
        oSpr2.ROW = i
        oSpr2.Col = 3
        If oSpr2.Text = "" Then
           toCnt = toCnt + 1
        End If
    Next i

    
    If CInt(txt_slab_cnt) > toCnt And Mid(sToLoc, 1, 2) <> "S0" Then
       sMsg = "已超出目的垛位的存放能力...！"
         GoTo MOVE_CLICK_ERROR
    End If
    
    
    If Val(txt_slab_cnt & "") < 1 Or Val(txt_p_row & "") < 1 Then Exit Sub
    
    iFromRow = txt_p_row
    iMoveCnt = txt_slab_cnt
    
'    For iCnt = Max_Rows To 1 Step -1
'        oSpr2.Col = 3
'        oSpr2.ROW = iCnt
'       If oSpr2.Text = "" Then
'          iToStaRow = iCnt
'          iCnt = 1
'          Exit For
'       End If
'    Next iCnt
    
    For iCnt = 1 To Max_Rows
        oSpr2.Col = 3
        oSpr2.ROW = iCnt
       If oSpr2.Text = "" Then
          iToStaRow = iCnt
          'iCnt = 1
       Else
          Exit For
       End If
    Next iCnt
    
    iRow2 = iToStaRow + 1
    oSpr2.ROW = iRow2
    oSpr2.Col = 3
    sTempSlabNo = oSpr2.Text
    oSpr2.Col = 8
    If (sTempSlabNo <> "" And oSpr2.Text = "") And iRow2 <> Max_Rows + 1 Then
       sMsg = "目的垛位上顶层板坯的宽度不存在！当前操作无法继续进行！"
         GoTo MOVE_CLICK_ERROR
    Else
       If oSpr2.Text <> "" Then
          iVal2 = oSpr2.Text
       End If
    End If
    
    oSpr1.Col = 8
    oSpr1.ROW = iFromRow
    If oSpr1.Text = "" Then
       sMsg = "起始垛位上要移动的板坯宽度不存在！当前操作无法继续进行！"
       GoTo MOVE_CLICK_ERROR
    Else
       If oSpr1.Text <> "" Then
          iVal1 = oSpr1.Text
       End If
    End If
    

    iRow2 = iToStaRow + 1
    oSpr2.ROW = iRow2
    oSpr2.Col = 3
    sTempSlabNo = oSpr2.Text
    oSpr2.Col = 9
    
    If (sTempSlabNo <> "" And oSpr2.Text = "") And iRow2 <> Max_Rows + 1 Then
       sMsg = "目的垛位上顶层板坯的长度不存在！当前操作无法继续进行！"
       GoTo MOVE_CLICK_ERROR
    Else
       If oSpr2.Text <> "" Then
          iVal2 = oSpr2.Text
       End If
    End If
    
    oSpr1.Col = 9
    oSpr1.ROW = iFromRow
    If oSpr1.Text = "" Then
       sMsg = "起始垛位上要移动的板坯长度不存在！当前操作无法继续进行！"
       GoTo MOVE_CLICK_ERROR
    Else
       If oSpr1.Text <> "" Then
          iVal1 = oSpr1.Text
       End If
    End If
    
    
    For iCnt = iFromRow To iFromRow - iMoveCnt + 2 Step -1
       oSpr1.Col = 8
       oSpr1.ROW = iCnt
       If oSpr1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯宽度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal1 = oSpr1.Text
       
       oSpr1.Col = 8
       oSpr1.ROW = iCnt - 1
       If oSpr1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯宽度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal2 = oSpr1.Text
       
       oSpr1.Col = 9
       oSpr1.ROW = iCnt
       If oSpr1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯长度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal1 = oSpr1.Text
       oSpr1.Col = 9
       oSpr1.ROW = iCnt - 1
       If oSpr1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯长度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal2 = oSpr1.Text

    Next
    
    oSpr1.SetSelection 1, iFromRow - iMoveCnt + 1, 18, iFromRow
    oSpr1.ClipboardCopy
     
'    oSpr2.SetSelection 1, 39, 15, 40
    oSpr2.SetSelection 1, iToStaRow - iMoveCnt + 1, 18, iToStaRow
    oSpr2.ClipboardPaste
    

    With oSpr1
        For iCnt = iFromRow - iMoveCnt + 1 To iFromRow
          .ROW = iCnt
          .Col = 0
          oSpr1.Text = "Delete"
          For i = 1 To .MaxCols
            .Col = i
            .BackColor = &HFF
          Next
        Next
    End With

    
    With oSpr2
       
        For iCnt = iToStaRow To iToStaRow - iMoveCnt + 1 Step -1
              .ROW = iCnt
              .Col = 0
              .Text = "Input"
    
              .Col = 5
              .Text = sToLoc
    
              .Col = 18
              .Text = sUserID
              
              .Col = 19
              .Text = sFromLoc
             
              For i = 1 To .MaxCols
                .Col = i
                .BackColor = &HFF
              Next
'              .Col = 1
              
'              If .Row <> Max_Rows Then
'                 .Row = iCnt + 1
'                  sSeq = CInt(oSpr2.Text) + 1
'                  .Row = iCnt
'                  .Text = sSeq
'              Else
'                  .Text = "1"
'              End If
        Next
        
        oSpr2.ROW = iToStaRow - iMoveCnt + 1
        oSpr2.Col = 1
        If Len(Trim(oSpr2.Text)) = 1 Then
           To_Bedseq = "0" + oSpr2.Text
        ElseIf Len(Trim(oSpr2.Text)) = 2 Then
           To_Bedseq = oSpr2.Text
        End If
        
        sQuery = "SELECT MAX_CNT FROM FP_STDYARD WHERE LOCATION ='" + sToLoc + "'"
        ssc_move.Enabled = False
        opt_sequence(0).Enabled = False
        opt_sequence(1).Enabled = False
        opt_sequence(2).Enabled = False
        opt_sequence(3).Enabled = False
        
        
'        oSpr2.Row = 1
'        oSpr2.Col = 3
'        'If oSpr2.Text <> "" Then
'            If sToLoc <> "S0C0101" And sToLoc <> "S0L0101" And sToLoc <> "S0Q0101" And CInt(To_Bedseq) > Gf_FloatFind(M_CN1, sQuery) Then
'               MsgBox "已超出目的垛位的存放能力！当前操作无法继续！", vbCritical, "系统提示信息"
'               Call ssc_can_Click
'               Exit Sub
'            End If
'        'End If
    End With

    'Chk_oSpr1.Value = ssCBChecked
    
    'Dan sequnence check
    iCnt = 0
    For iCnt = oSpr2.MaxRows To 1 Step -1
        oSpr2.ROW = iCnt
        oSpr2.Col = 0
        If oSpr2.Text = "Input" Then
            Select Case txt_sequence
                Case "0"
                     oSpr2.ROW = iCnt
                     oSpr2.Col = 2
                     oSpr2.Text = "1"
                Case "1"
                     If iCnt = oSpr2.MaxRows Then
                        oSpr2.ROW = iCnt
                        oSpr2.Col = 1
                        oSpr2.Text = 1
                        oSpr2.Col = 1
                        oSpr2.Text = 1
                        oSpr2.Col = 2
                        oSpr2.Text = "1"
                     Else
                        oSpr2.ROW = iCnt + 1
                        oSpr2.Col = 2
                        If oSpr2.Text = "1" Then
                            oSpr2.Col = 2
                            oSpr2.ROW = iCnt
                            oSpr2.Text = "2"
                            
                            oSpr2.ROW = iCnt + 1
                            oSpr2.Col = 1
                            tmpSeqNo = oSpr2.Value
                            oSpr2.ROW = iCnt
                            oSpr2.Col = 1
                            oSpr2.Value = tmpSeqNo
                            
                        Else
                            oSpr2.Col = 2
                            oSpr2.ROW = iCnt
                            oSpr2.Text = "1"
                            
                            oSpr2.ROW = iCnt + 1
                            oSpr2.Col = 1
                            tmpSeqNo = oSpr2.Text + 1
                            oSpr2.ROW = iCnt
                            oSpr2.Col = 1
                            oSpr2.Value = tmpSeqNo
                        End If
                     End If
                 
                Case "2"
                     If iCnt = oSpr2.MaxRows Then
                        oSpr2.ROW = iCnt
                        oSpr2.Col = 1
                        oSpr2.Text = 1
                        oSpr2.Col = 2
                        oSpr2.Text = "1"
                     Else
                        oSpr2.ROW = iCnt + 1
                        oSpr2.Col = 2
                        If oSpr2.Text = "1" Then
                            oSpr2.Col = 2
                            oSpr2.ROW = iCnt
                            oSpr2.Text = "2"
                            
                            oSpr2.ROW = iCnt + 1
                            oSpr2.Col = 1
                            tmpSeqNo = oSpr2.Value
                            oSpr2.ROW = iCnt
                            oSpr2.Col = 1
                            oSpr2.Value = tmpSeqNo
                        ElseIf oSpr2.Text = "2" Then
                            oSpr2.Col = 2
                            oSpr2.ROW = iCnt
                            oSpr2.Text = "3"
                            
                            oSpr2.ROW = iCnt + 1
                            oSpr2.Col = 1
                            tmpSeqNo = oSpr2.Value
                            oSpr2.ROW = iCnt
                            oSpr2.Col = 1
                            oSpr2.Value = tmpSeqNo
                        ElseIf oSpr2.Text = "3" Then
                            oSpr2.Col = 2
                            oSpr2.ROW = iCnt
                            oSpr2.Text = "1"
                            
                            oSpr2.ROW = iCnt + 1
                            oSpr2.Col = 1
                            tmpSeqNo = oSpr2.Value + 1
                            oSpr2.ROW = iCnt
                            oSpr2.Col = 1
                            oSpr2.Value = tmpSeqNo
                        End If
                     End If
                
                Case "3"
                     If iCnt = oSpr2.MaxRows Then
                        oSpr2.ROW = iCnt
                        oSpr2.Col = 1
                        oSpr2.Text = 1
                        oSpr2.Col = 2
                        oSpr2.Text = "1"
                     Else
                        oSpr2.ROW = iCnt + 1
                        oSpr2.Col = 2
                        If oSpr2.Text = "1" Then
                            oSpr2.Col = 2
                            oSpr2.ROW = iCnt
                            oSpr2.Text = "2"
                            
                            oSpr2.ROW = iCnt + 1
                            oSpr2.Col = 1
                            tmpSeqNo = oSpr2.Value
                            oSpr2.ROW = iCnt
                            oSpr2.Col = 1
                            oSpr2.Value = tmpSeqNo
                        ElseIf oSpr2.Text = "2" Then
                            oSpr2.Col = 2
                            oSpr2.ROW = iCnt
                            oSpr2.Text = "3"
                            
                            oSpr2.ROW = iCnt + 1
                            oSpr2.Col = 1
                            tmpSeqNo = oSpr2.Value
                            oSpr2.ROW = iCnt
                            oSpr2.Col = 1
                            oSpr2.Value = tmpSeqNo
                        ElseIf oSpr2.Text = "3" Then
                            oSpr2.Col = 2
                            oSpr2.ROW = iCnt
                            oSpr2.Text = "4"
                            
                            oSpr2.ROW = iCnt + 1
                            oSpr2.Col = 1
                            tmpSeqNo = oSpr2.Value
                            oSpr2.ROW = iCnt
                            oSpr2.Col = 1
                            oSpr2.Value = tmpSeqNo
                        ElseIf oSpr2.Text = "4" Then
                            oSpr2.Col = 2
                            oSpr2.ROW = iCnt
                            oSpr2.Text = "1"
                            
                            oSpr2.ROW = iCnt + 1
                            oSpr2.Col = 1
                            tmpSeqNo = oSpr2.Value + 1
                            oSpr2.ROW = iCnt
                            oSpr2.Col = 1
                            oSpr2.Value = tmpSeqNo
                        End If
                     End If
                     
                
            End Select
        End If
    Next
    
    
    
Exit Sub
    
MOVE_CLICK_ERROR:
    Call Gp_MsgBoxDisplay(sMsg)

End Sub


Private Sub txt_f_addr_Change()
Dim sQuery As String
    If Len(txt_f_addr) = 7 Then
       sQuery = "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'F0033' AND CD = '" & txt_f_addr.Text & "'"
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
          MsgBox txt_f_addr.Tag & "不正确，请重新输入！", vbCritical, "系统提示信息"
          Exit Sub
       End If
'       opt_Right_Left.Enabled = True
       If Mid(Trim(txt_f_addr), 1, 2) = "S0" Then
          opt_Left_Right.Value = True
          opt_Right_Left.Enabled = False
       End If
    End If
End Sub

Private Sub txt_f_addr_DblClick()
    Call txt_f_addr_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_f_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        txt_f_addr.Text = "S"
        DD.sWitch = "MS"
        DD.sKey = "F0033"
        DD.rControl.Add Item:=txt_f_addr
        DD.rControl.Add Item:=txt_o_f_addr_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)
        txt_o_f_addr.Text = txt_f_addr.Text
        Exit Sub

    End If


    If Len(Trim(txt_f_addr)) = txt_f_addr.MaxLength Then
        txt_o_f_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0033", Trim(txt_f_addr.Text), 2)
    Else
        txt_o_f_addr_nm.Text = ""
    End If

    If Len(Trim(txt_f_addr)) = 7 Then
       txt_o_f_addr.Text = txt_f_addr.Text
    Else
       txt_o_f_addr.Text = ""
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
    
    Dim sMesg As String
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
                sMesg = Gf_Sp_NeceCheck2(Scc.Item("Spread"), Scc.Item("mColumn"), iCount, Scc.Item("nColumn"))
                        
                If Trim(sMesg) = "OK" Then
                    
                ElseIf Mid(sMesg, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Scc.Item("Spread"), iCount, , vbYellow)
                    sMesg = Mid(sMesg, 6, Len(sMesg))
                    sMesg = sMesg + "长度不正确"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Sp_Process = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Scc.Item("Spread"), iCount, , vbYellow)
                    sMesg = sMesg + "必须输入"
                    Call Gp_MsgBoxDisplay(sMesg)
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
    
    Err.Raise Err.Number, Err.Description

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

    If Len(TXT_SLAB_NO) = 10 Then
       sQuery = "SELECT * FROM FP_SLAB WHERE SLAB_NO = '" + TXT_SLAB_NO + "'"
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
       
        MsgBox "该板坯不存在，板坯号无效！", vbCritical, "系统提示信息"
        If txt_t_addr <> "" Then
           TXT_SLAB_NO = ""
        Else
           Exit Sub
        End If
       End If
    End If
    
    If Len(TXT_SLAB_NO) = 10 And Mid(txt_f_addr, 1, 2) <> "S0" And S1_Click <> "1" Then
        txt_t_addr = ""
        txt_o_t_addr = ""
        txt_o_t_addr_nm = ""
    End If
End Sub

Private Sub txt_t_addr_Change()

Dim sQuery As String
    If Len(txt_t_addr) = 7 Then
       sQuery = "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'F0033' AND CD = '" & txt_t_addr.Text & "'"
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
          MsgBox txt_t_addr.Tag & "不正确，请重新输入！", vbCritical, "系统提示信息"
          Exit Sub
       End If
    End If
   
   If Len(txt_t_addr) = 7 And Mid(txt_f_addr, 1, 2) <> "S0" Then
      TXT_SLAB_NO = ""
      sQuery = "SELECT MAX(CROSS_SEQ) FROM FP_SLABYARD WHERE YARD_ADDR = '" + txt_t_addr + "' AND YARD_KND = 'ZB'"
      sCross_Seq = Gf_FloatFind(M_CN1, sQuery)
      If sCross_Seq = "0" Then sCross_Seq = "2"
      opt_sequence(CInt(sCross_Seq) - 1).Value = True
   ElseIf Mid(Trim(txt_t_addr), 1, 2) = "S0" Then
      opt_Left_Right.Value = True
      opt_Right_Left.Value = False
   End If
End Sub

Private Sub txt_t_addr_DblClick()
    Call txt_t_addr_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_t_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        txt_t_addr.Text = "S"
        DD.sWitch = "MS"
        DD.sKey = "F0033"
        DD.rControl.Add Item:=txt_t_addr
        DD.rControl.Add Item:=txt_o_t_addr_nm

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)
        txt_o_t_addr.Text = txt_t_addr.Text
        Exit Sub

    End If

    If Len(Trim(txt_t_addr)) = txt_t_addr.MaxLength Then
        txt_o_t_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0033", Trim(txt_t_addr.Text), 2)
    Else
        txt_o_t_addr_nm.Text = ""
    End If


    If Len(Trim(txt_t_addr)) = 7 Then
       txt_o_t_addr.Text = txt_t_addr.Text
       If Mid(txt_f_addr, 1, 2) <> "S0" Then
          TXT_SLAB_NO = ""
       End If
    Else
       txt_o_t_addr.Text = ""
    End If

End Sub

Private Sub cmd_Loc_Search_Click()
    
    Dim OutParam(3, 4) As Variant
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    If Trim(TXT_SLAB_NO.Text) = "" Then
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
        
    sQuery = "{call AFL2010P ('SL','" & Trim(TXT_SLAB_NO.Text) & "',?,?,?)}"
    
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
