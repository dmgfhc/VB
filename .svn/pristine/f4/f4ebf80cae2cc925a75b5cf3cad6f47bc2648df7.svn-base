VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFL2030C 
   Caption         =   "移送板坯再入库实绩录入界面_AFL2030C"
   ClientHeight    =   7725
   ClientLeft      =   195
   ClientTop       =   1320
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   8535
      Left            =   135
      TabIndex        =   8
      Top             =   615
      Width           =   15075
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         Left            =   8955
         Locked          =   -1  'True
         TabIndex        =   11
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
         Left            =   11475
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   2640
      End
      Begin Threed.SSCheck Chk_ss2 
         Height          =   330
         Left            =   7650
         TabIndex        =   9
         Top             =   675
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "目的垛位"
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   7440
         Left            =   90
         TabIndex        =   14
         Top             =   1035
         Width           =   7380
         _Version        =   393216
         _ExtentX        =   13018
         _ExtentY        =   13123
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
         MaxCols         =   13
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFL2030C.frx":0000
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   2790
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "移送库名称"
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
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   90
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         Caption         =   "已移送库 "
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   10290
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
         Left            =   7545
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   7440
         Left            =   7545
         TabIndex        =   15
         Top             =   1035
         Width           =   7380
         _Version        =   393216
         _ExtentX        =   13018
         _ExtentY        =   13123
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
         MaxCols         =   15
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFL2030C.frx":1C53
      End
      Begin Threed.SSCheck SSCheck1 
         Height          =   330
         Left            =   90
         TabIndex        =   18
         Top             =   675
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   255
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "已移送板坯"
         Value           =   1
      End
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
      Left            =   8355
      MaxLength       =   10
      TabIndex        =   6
      Top             =   165
      Width           =   1455
   End
   Begin VB.TextBox txt_p_row 
      Height          =   315
      Left            =   14745
      TabIndex        =   5
      Text            =   " "
      Top             =   165
      Width           =   465
   End
   Begin VB.TextBox txt_slab_cnt 
      Height          =   315
      Left            =   12600
      TabIndex        =   4
      Text            =   " "
      Top             =   165
      Width           =   465
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
      Left            =   1425
      MaxLength       =   2
      TabIndex        =   0
      Top             =   165
      Width           =   840
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
      Left            =   11295
      MaxLength       =   7
      TabIndex        =   3
      Top             =   165
      Width           =   975
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   330
      Left            =   465
      TabIndex        =   7
      Top             =   1290
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "起始垛位"
      Value           =   1
   End
   Begin Threed.SSCommand ssc_can 
      Height          =   315
      Left            =   13935
      TabIndex        =   16
      Top             =   165
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      Enabled         =   0   'False
      Caption         =   "&取消"
   End
   Begin Threed.SSCommand ssc_move 
      Height          =   315
      Left            =   13095
      TabIndex        =   17
      Top             =   165
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      Enabled         =   0   'False
      Caption         =   "&移动"
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   7155
      Top             =   165
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
      Left            =   225
      Top             =   165
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Caption         =   "移送库"
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
      Left            =   10080
      Top             =   165
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.UDate txt_to_DATE 
      Height          =   315
      Left            =   5445
      TabIndex        =   2
      Tag             =   "移送日期"
      Top             =   165
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
   End
   Begin InDate.UDate txt_from_DATE 
      Height          =   315
      Left            =   3690
      TabIndex        =   1
      Tag             =   "移送日期"
      Top             =   165
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
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   2580
      Top             =   165
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "移送日期"
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
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5220
      TabIndex        =   19
      Top             =   195
      Width           =   255
   End
End
Attribute VB_Name = "AFL2030C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      RETURN SLAB INPUTTING
'-- Program ID        AFL2030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2005.3.23
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

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      'Call Gp_Ms_Collection(txt_slab_no, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_f_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_t_addr, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_from_DATE, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_to_DATE, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, " ", "n", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    'Sc1.Add Item:="AFL2030C.P_MODIFY1", Key:="P-M"
    sc1.Add Item:="AFL2030C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFL2030C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:="AFL2030C.P_REFER2", Key:="P-R"
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
    
    Call Gp_Sp_ColHidden(ss1, 13, True)
    Call Gp_Sp_ColHidden(ss2, 15, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Form_Activate()
    
'    ss2.MaxRows = Max_Rows
'    ss2.Row = Max_Rows
'    ss2.Col = 1
'    ss2.Action = ActionActiveCell
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(7).Enabled = False
    txt_o_f_addr.Text = txt_f_addr.Text
    txt_o_f_addr_nm.Text = Gf_ComnNameFind(M_CN1, "C0013", Trim(txt_f_addr.Text), 2)
    
    txt_o_t_addr.Text = txt_t_addr.Text
    txt_o_t_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0009", Trim(txt_t_addr.Text), 2)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

Max_Rows = 60
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(7).Enabled = False
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    sc1.Item("Spread").RetainSelBlock = False
    sc2.Item("Spread").RetainSelBlock = False
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "F-System.INI", Me.Name)
    
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
            txt_from_DATE.Text = ""
            txt_to_DATE.Text = ""
            txt_slab_no.Text = ""
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
                 MDIMain.MenuTool.Buttons(8).Enabled = False
                 MDIMain.MenuTool.Buttons(9).Enabled = False
                 MDIMain.MenuTool.Buttons(7).Enabled = False
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
                
'                ss2.MaxRows = Max_Rows
'                ss2.Row = Max_Rows
'                ss2.Col = 1
'                ss2.Action = ActionActiveCell
                
                txt_f_addr.SetFocus
        End If
    End If

End Sub

Public Sub Form_Ref()
Dim iCnt As Integer

On Error GoTo Refer_Err

    Dim iRow, iCol, MaxCnt, iStemp As Integer
    Dim sMsg, sMesg, sTemp, sQuery As String
Call ssc_can_Click
     
If Trim(txt_f_addr) <> "" Then
    sQuery = "SELECT * FROM ZP_CD WHERE CD_MANA_NO = 'C0013'"
    If Gf_FloatFind(M_CN1, sQuery) = 0 Then
       
       MsgBox "移送库不正确，请重新输入！", vbCritical, "系统提示信息"
       Exit Sub
    End If
End If

If Trim(txt_t_addr) <> "" And Trim(txt_t_addr) <> "S0A0101" Then
    sQuery = "SELECT * FROM FP_STDYARD WHERE LOCATION = '" + txt_t_addr + "' AND YARD_KND = '00'"
    If Gf_FloatFind(M_CN1, sQuery) = 0 Then
       
       MsgBox "目的垛位号不正确，请重新输入！", vbCritical, "系统提示信息"
       Exit Sub
    End If
End If

'    If (txt_f_addr <> "" And Gf_FloatFind(M_CN1, "SELECT COUNT(*) FROM FP_STDYARD WHERE LOCATION = '" + txt_f_addr + "'") = 0) Or _
'       (txt_t_addr <> "" And Gf_FloatFind(M_CN1, "SELECT COUNT(*) FROM FP_STDYARD WHERE LOCATION = '" + txt_t_addr + "'") = 0) Then
'
'
'       MsgBox "请输入正确的垛位号！", vbCritical, "系统提示信息"
'       Exit Sub
'    Else
'    End If

    
    If txt_t_addr = "S0A0101" Then
       Call Gp_Sp_ColHidden(ss2, 1, True)
    Else
       Call Gp_Sp_ColHidden(ss2, 1, False)
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
    
         sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then
       
            If Gf_Sp_Refer(M_CN1, sc1, Mc1, Nothing, Nothing, False) Then
                sc1.Item("Spread").OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                MDIMain.MenuTool.Buttons(9).Enabled = False
                MDIMain.MenuTool.Buttons(8).Enabled = False
                MDIMain.MenuTool.Buttons(7).Enabled = False
                MaxCnt = ss1.MaxRows
                
'                ss1.MaxRows = Max_Rows
'                ss1.Row = Max_Rows
'                ss1.Col = 1
'                ss1.Action = ActionActiveCell
                
                
                With ss1
                     .Col = 13
                     For iRow = MaxCnt To 1 Step -1
                         .Row = iRow
                         .Text = sUserID
                     Next iRow


                    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)

                End With
            Else
                
                sc1.Item("Spread").MaxRows = Max_Rows
                sc1.Item("Spread").Row = Max_Rows
                sc1.Item("Spread").Col = 1
                sc1.Item("Spread").ACTION = ActionActiveCell

            End If
            
            
            If Gf_Sp_Refer(M_CN1, sc2, Mc1, Nothing, Nothing, False) Then
                sc2.Item("Spread").OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                MDIMain.MenuTool.Buttons(9).Enabled = False
                MDIMain.MenuTool.Buttons(8).Enabled = False
                MDIMain.MenuTool.Buttons(7).Enabled = False
                MaxCnt = ss2.MaxRows
                
                sc2.Item("Spread").MaxRows = Max_Rows
                sc2.Item("Spread").Row = Max_Rows
                sc2.Item("Spread").Col = 1
                sc2.Item("Spread").ACTION = ActionActiveCell
                
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
  
                    If txt_slab_no <> "" Then
                         sQuery = "SELECT * FROM FP_SLABYARD WHERE SLAB_NO = '" + txt_slab_no + "'"
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
                            .Text = txt_t_addr.Text
                            .Col = 15
                            .Text = txt_f_addr.Text
                            ssc_can.Enabled = True
                        
                             For iCol = 1 To 14
                                .Col = iCol
                                .BackColor = &HFF
                             Next iCol
                             
                             MDIMain.MenuTool.Buttons(4).Enabled = True
                         Else
                            
                            If txt_t_addr = "" Then
                               .Col = 2
                                For iRow = Max_Rows To Max_Rows - MaxCnt + 1 Step -1
                                    .Row = iRow
                                     If .Text = txt_slab_no Then
                                        .SetSelection 2, .Row, 2, .Row
                                        .ForeColor = &HFF
                                     End If
                                Next iRow
                            End If
                            
                            Exit Sub
                         End If
                    End If
                End With
            Else
                
                sc2.Item("Spread").MaxRows = Max_Rows
                sc2.Item("Spread").Row = Max_Rows
                sc2.Item("Spread").Col = 1
                sc2.Item("Spread").ACTION = ActionActiveCell
                
                If txt_slab_no <> "" And Trim(txt_f_addr) <> "" Then
                   sQuery = "SELECT * FROM FP_SLABYARD WHERE SLAB_NO = " + txt_slab_no
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
                             .Text = txt_t_addr.Text
                             .Col = 15
                             .Text = txt_f_addr.Text
                             
                             iMoveCnt = 1
                             iToStaRow = .Row
                             txt_slab_cnt = 1
                             
                             For iCol = 1 To 14
                                .Col = iCol
                                .BackColor = &HFF
                             Next iCol
                             
                             MDIMain.MenuTool.Buttons(4).Enabled = True
                        End With
                        ssc_can.Enabled = True
                   Else
                        MsgBox "板坯 " + txt_slab_no + " 已经在库中，不需再做入库处理！", vbInformation, "系统提示信息"
                        txt_slab_no = ""
                        Exit Sub
                   End If
                
                Else
                   If txt_t_addr <> "" And txt_t_addr <> "S0A0101" Then
                      MsgBox "垛位 " + txt_t_addr + " 没有板坯！", vbInformation, "系统提示信息"
                   ElseIf txt_t_addr = "S0A0101" Then
                      MsgBox "没有在线板坯等待入库！", vbInformation, "系统提示信息"
                   End If
                End If

            End If
            
            Call Gp_Sp_ColGet(sc2.Item("Spread"), "F-System.INI", Me.Name)
            
        Else
            sMesg = sMesg + "长度不正确"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    
    Else
        sMesg = sMesg + "必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
        
    End If
    
    If txt_slab_no = "" Then
       txt_slab_cnt = ""
       ssc_move.Enabled = False
       ssc_can.Enabled = False
       txt_p_row = ""
    End If
    
'     If iToStaRow = 0 And AFL2040C.Active_CForm <> "" Then

       If Trim(txt_slab_cnt) <> "" Then
          iMoveCnt = CInt(txt_slab_cnt)
       End If
       
       For iCnt = Max_Rows To 1 Step -1
           ss2.Col = 1
           ss2.Row = iCnt
           If ss2.Text = "" Then
              iToStaRow = iCnt + 1
              iCnt = 1
              Exit For
           End If
       Next iCnt

        ss2.Row = iToStaRow - iMoveCnt + 1
        ss2.Col = 1
        If Len(Trim(ss2.Text)) = 1 Then
           To_Bedseq = "0" + ss2.Text
        ElseIf Len(Trim(ss2.Text)) = 2 Then
           To_Bedseq = ss2.Text
        End If
        
        sQuery = "SELECT MAX_CNT FROM FP_STDYARD WHERE LOCATION ='" + txt_t_addr + "' AND YARD_KND = '00'"
        ssc_move.Enabled = False
        'If ss2.Text <> "" Then
            If txt_t_addr <> "S0A0101" And To_Bedseq > Gf_FloatFind(M_CN1, sQuery) Then
               MsgBox "已超出目的垛位的存放能力！当前操作无法继续！", vbCritical, "系统提示信息"
               Call ssc_can_Click
               Exit Sub
            End If
        'End If
     
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
    
    Dim icount As Integer
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
        .MaxRows = Max_Rows: icount = 0
        
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
                                .VALUE = 0
                            Else
                                .VALUE = Trim(ArrayRecords(iColcount, iRowCount))
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
            For icount = 1 To lColumn.Count

                .Protect = True
                .Col = lColumn(icount): .Col2 = lColumn(icount)
                .Row = 1: .Row2 = .MaxRows
                .BlockMode = True: .Lock = True
                .BlockMode = False

            Next icount

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
Dim SlabNo As String
Dim i As Integer
Dim sQuery As String

ss2.Row = 0
ss2.Col = 0
If (Chk_ss2.VALUE = ssCBUnchecked And ss2.Text <> "") Or (Chk_ss2.VALUE = ssCBChecked And ss2.Text <> "◎") Then
   MsgBox "目的垛位已达最大板坯数或垛层不正确！", vbCritical, "系统提示信息"
   Call ssc_can_Click
   Call Form_Ref
   ss1.Row = 0
   ss2.Row = 0
   For i = 0 To ss1.MaxCols
       ss1.Col = i
       ss2.Col = i
       If i = 0 Then
          ss2.Text = ""
       Else
          ss2.Text = ss1.Text
       End If
       
   Next i
   Exit Sub
End If
    
'        If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then
'           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'        End If
        
        If Gf_Sp_Process(M_CN1, sc2, Mc1) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           MDIMain.MenuTool.Buttons(9).Enabled = False
           MDIMain.MenuTool.Buttons(8).Enabled = False
           MDIMain.MenuTool.Buttons(7).Enabled = False
           SlabNo = txt_slab_no
           txt_slab_no = ""
           
           If txt_f_addr <> "" Then
              MDIMain.StatusBar1.Panels(1).Text = "板坯 " + SlabNo + " 已成功入库！"
           End If
           S1_Click = ""

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
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("Sc"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
        
    Dim iCnt, i As Integer
    Dim sFlag As String
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
   
With ss1
   .Col = 0
   For i = Max_Rows To 1 Step -1
       .Row = i
       If .Text = "Delete" Then
          sFlag = "Y"
          Exit For
       End If
   Next i
   
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
    
    If sFlag = "Y" Then
            For iCnt = Row To 1 Step -1
                .Col = 2
                .Row = iCnt
                If .Text <> "" Then
                   i = i + 1
                ElseIf .Text = "" Then
                   
                   If .Row <> Max_Rows Then
                      .Row = .Row + 1
                   End If
                   
                   TopSlabNo = .Text
                   TopSlabRow = .Row
                   Exit For
                End If
            Next iCnt
       
            If i <> 0 Then
               txt_slab_cnt = i
            End If
    
    Else
    
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



'Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'
'    Dim iCurrRow    As Integer
'    Dim iCurrRowVal As Integer
'    Dim iCalCurrRow  As Integer
'    Dim iCalChgRow     As Integer
'    Dim iChgRow     As Integer
'    Dim iCnt      As Integer
'
'    If Gf_Sc_Authority(sAuthority, "U") Then
'       ' Proc_Sc("SC").Item("Spread").Row = Row
'       ' Proc_Sc("SC").Item("Spread").Col = 0
'       ' iCurrRow = Proc_Sc("SC").Item("Spread").Text
'
'        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'
'        Proc_Sc("SC").Item("Spread").Row = Row
'
'        Proc_Sc("SC").Item("Spread").Col = 0
'
'        If Proc_Sc("SC").Item("Spread").Text = "Update" Then
'
'           Proc_Sc("SC").Item("Spread").Row = Row
'           Proc_Sc("SC").Item("Spread").Col = 1
'
'           If Proc_Sc("SC").Item("Spread").Value = "" Then
'
'            iCurrRowVal = Max_Rows - Row + 1
'
'           Else
'                iCurrRowVal = Proc_Sc("SC").Item("Spread").Value
'
'           End If
'
'           iCalCurrRow = Max_Rows - Row + 1
'
'           If iCurrRowVal <> iCalCurrRow Then
'
'            If iCalCurrRow < iCurrRowVal Then
'
'               iCalChgRow = iCurrRowVal - iCalCurrRow
'               iChgRow = Row - iCalChgRow
'
'               For iCnt = iChgRow To Row - 1 Step 1
'
'                 Proc_Sc("SC").Item("Spread").Row = iCnt
'                 Proc_Sc("SC").Item("Spread").Col = 1
'                 Proc_Sc("SC").Item("Spread").Text = Proc_Sc("SC").Item("Spread").Text - 1
'                 Proc_Sc("SC").Item("Spread").Row = iCnt
'                 Proc_Sc("SC").Item("Spread").Col = 0
'                 Proc_Sc("SC").Item("Spread").Text = "Update"
'
'               Next
'            Else
'
'               iCalChgRow = iCalCurrRow - iCurrRowVal
'               iChgRow = Row + iCalChgRow
'
'               For iCnt = iChgRow To Row + 1 Step -1
'
'                  Proc_Sc("SC").Item("Spread").Row = iCnt
'                  Proc_Sc("SC").Item("Spread").Col = 1
'                  Proc_Sc("SC").Item("Spread").Text = iCurrRowVal + 1
'
'                  Proc_Sc("SC").Item("Spread").Row = iCnt
'                  Proc_Sc("SC").Item("Spread").Col = 0
'                  Proc_Sc("SC").Item("Spread").Text = "Update"
'               Next
'            End If
'           End If
'         End If
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 14)   ' 荐沥且锭 淬寸磊 ID
'
'    End If
'End Sub

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

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
   txt_slab_no = ""
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
'    Dim iCurrRow    As Integer
'    Dim iCurrRowVal As Integer
'    Dim iCalCurrRow  As Integer
'    Dim iCalChgRow     As Integer
'    Dim iChgRow     As Integer
'    Dim iCnt      As Integer
'
'    If Gf_Sc_Authority(sAuthority, "U") Then
'       ' Proc_Sc("SC").Item("Spread").Row = Row
'       ' Proc_Sc("SC").Item("Spread").Col = 0
'       ' iCurrRow = Proc_Sc("SC").Item("Spread").Text
'
'        Call Gp_Sp_UpdateMake(Proc_Sc("SC2")("Spread"), Mode)
'
'        Proc_Sc("SC2").Item("Spread").Row = Row
'
'        Proc_Sc("SC2").Item("Spread").Col = 0
'
'        If Proc_Sc("SC2").Item("Spread").Text = "Update" Then
'
'           Proc_Sc("SC2").Item("Spread").Row = Row
'           Proc_Sc("SC2").Item("Spread").Col = 1
'
'           iCurrRowVal = Proc_Sc("SC2").Item("Spread").Text
'
'           iCalCurrRow = Max_Rows - Row + 1
'
'           If iCurrRowVal <> iCalCurrRow Then
'
'            If iCalCurrRow < iCurrRowVal Then
'
'               iCalChgRow = iCurrRowVal - iCalCurrRow
'               iChgRow = Row - iCalChgRow
'
'               For iCnt = iChgRow To Row - 1 Step 1
'
'                 Proc_Sc("SC2").Item("Spread").Row = iCnt
'                 Proc_Sc("SC2").Item("Spread").Col = 1
'                 Proc_Sc("SC2").Item("Spread").Text = Proc_Sc("SC2").Item("Spread").Text - 1
'                 Proc_Sc("SC2").Item("Spread").Row = iCnt
'                 Proc_Sc("SC2").Item("Spread").Col = 0
'                 Proc_Sc("SC2").Item("Spread").Text = "Update"
'
'               Next
'            Else
'
'               iCalChgRow = iCalCurrRow - iCurrRowVal
'               iChgRow = Row + iCalChgRow
'
'               For iCnt = iChgRow To Row + 1 Step -1
'
'                 Proc_Sc("SC2").Item("Spread").Row = iCnt
'                 Proc_Sc("SC2").Item("Spread").Col = 1
'                 Proc_Sc("SC2").Item("Spread").Text = iCurrRowVal + 1
'
'                 Proc_Sc("SC2").Item("Spread").Row = iCnt
'                 Proc_Sc("SC2").Item("Spread").Col = 0
'                 Proc_Sc("SC2").Item("Spread").Text = "Update"
'               Next
'            End If
'           End If
'         End If
'         Call Gp_Sp_InAuthority(Proc_Sc("Sc2"), 14)   ' 荐沥且锭 淬寸磊 ID
'
'    End If
    
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

Private Sub Chk_ss1_Click(VALUE As Integer)
    
    If Chk_ss1.VALUE = ssCBUnchecked Then
       If Chk_ss2.VALUE = ssCBUnchecked Then
            Chk_ss1.VALUE = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Gf_Sp_Change(Proc_Sc, sc1) Then
        Chk_ss1.ForeColor = &HFF&
        Chk_ss2.ForeColor = &H808080
        Chk_ss2.VALUE = ssCBUnchecked
    Else
        Chk_ss1.VALUE = ssCBUnchecked
        Chk_ss2.VALUE = ssCBChecked
    End If
        
End Sub

Private Sub Chk_ss2_Click(VALUE As Integer)
    
    If Chk_ss2.VALUE = ssCBUnchecked Then
        If Chk_ss1.VALUE = ssCBUnchecked Then
            Chk_ss2.VALUE = ssCBChecked
        End If
        Exit Sub
    End If
    
    If Gf_Sp_Change(Proc_Sc, sc2) Then
        Chk_ss1.ForeColor = &H808080
        Chk_ss2.ForeColor = &HFF&
        Chk_ss1.VALUE = ssCBUnchecked
    Else
        Chk_ss2.VALUE = ssCBUnchecked
        Chk_ss1.VALUE = ssCBChecked
    End If
        
End Sub

Private Sub ssc_can_Click()

    Dim i As Integer
    Dim iCnt As Integer

    'For iCnt = iToStaRow To iToStaRow - iMoveCnt + 1 Step -1
        ss2.SetSelection 1, iToStaRow - iMoveCnt + 1, 15, iToStaRow
        ss2.ClipboardCut
    'Next
  
    With ss1
      
      For iCnt = iFromRow - iMoveCnt + 1 To iFromRow Step 1
         .Row = iCnt
         .Col = 0
         ss1.Text = ""
         For i = 1 To 3
          .Col = i
          .BackColor = &HC0FFFF
         Next i
         
         For i = 4 To .MaxCols
          .Col = i
          .BackColor = &HFFFFFF
         Next i
      Next
    End With
    
    With ss2
       
      For iCnt = iToStaRow To iToStaRow - iMoveCnt + 1 Step -1
         .Row = iCnt
         .Col = 0
         ss2.Text = ""
         For i = 1 To 3
          .Col = i
          .BackColor = &HC0FFFF
         Next i
         
         For i = 4 To .MaxCols
          .Col = i
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
    
    ssc_can.Enabled = False
         
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
    Dim sMsg, sSeq, sQuery As String
    Dim sTempSlabNo As String
    Dim sBedSeq As String
    
   If txt_t_addr = "" Then
      sMsg = "请输入目的垛位号！"
      GoTo MOVE_CLICK_ERROR
   End If
   iFromRow = txt_p_row
   iMoveCnt = txt_slab_cnt
   
   For iCnt = Max_Rows To 1 Step -1
       ss2.Col = 1
       ss2.Row = iCnt
      If ss2.Text = "" Then
         iToStaRow = iCnt
         iCnt = 1
         Exit For
      End If
   Next iCnt
    
  ' From Address slab  move to To Address check width
  If (Chk_ss2.VALUE = ssCBUnchecked And ss2.Text <> "") Or (Chk_ss2.VALUE = ssCBChecked And ss2.Text <> "◎") Then
        MsgBox "目的垛位已达最大板坯数或垛层不正确！", vbCritical, "系统提示信息"
        Call ssc_can_Click
        Call Form_Ref
        ss1.Row = 0
        ss2.Row = 0
        For i = 0 To ss1.MaxCols
            ss1.Col = i
            ss2.Col = i
            If i = 0 Then
               ss2.Text = ""
            Else
               ss2.Text = ss1.Text
            End If
            
        Next i
        Exit Sub
  End If
    
   iRow2 = iToStaRow + 1
   ss2.Row = iRow2
   ss2.Col = 1
   sTempSlabNo = ss2.Text
   ss2.Col = 6
   If (sTempSlabNo <> "" And ss2.Text = "") And iRow2 <> Max_Rows + 1 Then
       sMsg = "目的垛位上顶层板坯的宽度不存在！当前操作无法继续进行！"
         GoTo MOVE_CLICK_ERROR
   Else
       If ss2.Text <> "" Then
          iVal2 = ss2.Text
       End If
   End If
   
   ss1.Col = 5
   ss1.Row = iFromRow
   If ss1.Text = "" Then
       sMsg = "起始垛位上要移动的板坯宽度不存在！当前操作无法继续进行！"
       GoTo MOVE_CLICK_ERROR
   Else
       If ss1.Text <> "" Then
          iVal1 = ss1.Text
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
   ss2.Row = iRow2
   ss2.Col = 2
   sTempSlabNo = ss2.Text
   ss2.Col = 7
   
   If (sTempSlabNo <> "" And ss2.Text = "") And iRow2 <> Max_Rows + 1 Then
       sMsg = "目的垛位上顶层板坯的长度不存在！当前操作无法继续进行！"
       GoTo MOVE_CLICK_ERROR
   Else
       If ss2.Text <> "" Then
          iVal2 = ss2.Text
       End If
   End If
   
   ss1.Col = 6
   ss1.Row = iFromRow
   If ss1.Text = "" Then
       sMsg = "起始垛位上要移动的板坯长度不存在！当前操作无法继续进行！"
       GoTo MOVE_CLICK_ERROR
   Else
       If ss1.Text <> "" Then
          iVal1 = ss1.Text
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
       ss1.Col = 5
       ss1.Row = iCnt
       If ss1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯宽度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal1 = ss1.Text
       
       ss1.Col = 5
       ss1.Row = iCnt - 1
       If ss1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯宽度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal2 = ss1.Text
       
'       If iVal1 < iVal2 Then
'          iGap = iVal2 - iVal1
'          If iGap > 100 Then
'             sMsg = "起始垛位上要移动的板坯宽度不符合堆放标准！"
'              GoTo MOVE_CLICK_ERROR
'          End If
'       End If
       
       ss1.Col = 6
       ss1.Row = iCnt
       If ss1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯长度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal1 = ss1.Text
       ss1.Col = 6
       ss1.Row = iCnt - 1
       If ss1.Text = "" Then
          sMsg = "起始垛位上要移动的板坯长度不存在！"
          GoTo MOVE_CLICK_ERROR
       End If
       iVal2 = ss1.Text
'       If iVal1 < iVal2 Then
'          iGap = iVal2 - iVal1
'          If iGap > 1000 Then
'             sMsg = "起始垛位上要移动的板坯长度不符合堆放标准！"
'              GoTo MOVE_CLICK_ERROR
'          End If
'       End If
   Next
   
   ss1.SetSelection 2, iFromRow - iMoveCnt + 1, 13, iFromRow
   ss1.ClipboardCopy
     
   ss2.SetSelection 3, iToStaRow - iMoveCnt + 1, 14, iToStaRow
   ss2.ClipboardPaste
    
'    If txt_f_addr = "S0A0101" Then
'        With ss1
'            .Row = iFromRow
'            .Col = 0
'            ss1.Text = "Delete"
'            For iCnt = 1 To .MaxCols
'             .Col = iCnt
'             .BackColor = &HFF
'            Next
'        End With
'    Else
        With ss1
            For iCnt = iFromRow - iMoveCnt + 1 To iFromRow
              .Row = iCnt
              .Col = 0
              ss1.Text = "Delete"
              For i = 1 To .MaxCols
                .Col = i
                .BackColor = &HFF
              Next
            Next
        End With
'    End If

    With ss2
       
        For iCnt = iToStaRow To iToStaRow - iMoveCnt + 1 Step -1
              .Row = iCnt
              .Col = 0
              .Text = "Input"
    
              .Col = 2
              .Text = txt_t_addr
    
              .Col = 14
              .Text = sUserID
              
              .Col = 15
              .Text = txt_f_addr
             
              For i = 1 To .MaxCols
                .Col = i
                .BackColor = &HFF
              Next
              .Col = 1
              
              If .Row <> Max_Rows Then
                 .Row = iCnt + 1
                  sSeq = CInt(ss2.Text) + 1
                  .Row = iCnt
                  .Text = sSeq
              Else
                  .Text = "1"
              End If
        Next
        
        ss2.Row = iToStaRow - iMoveCnt + 1
        ss2.Col = 1
        If Len(Trim(ss2.Text)) = 1 Then
           To_Bedseq = "0" + ss2.Text
        ElseIf Len(Trim(ss2.Text)) = 2 Then
           To_Bedseq = ss2.Text
        End If
        
        sQuery = "SELECT MAX_CNT FROM FP_STDYARD WHERE LOCATION ='" + txt_t_addr + "' AND YARD_KND = '00'"
        ssc_move.Enabled = False
        
        ss2.Row = 1
        ss2.Col = 2
        'If ss2.Text <> "" Then
            If txt_t_addr <> "S0A0101" And To_Bedseq > Gf_FloatFind(M_CN1, sQuery) Then
               MsgBox "已超出目的垛位的存放能力！当前操作无法继续！", vbCritical, "系统提示信息"
               Call ssc_can_Click
               Exit Sub
            End If
        'End If
    End With
    Exit Sub
    'Chk_ss1.Value = ssCBChecked
    
MOVE_CLICK_ERROR:
    Call Gp_MsgBoxDisplay(sMsg)

End Sub

Private Sub txt_f_addr_Change()
Dim sQuery As String
    If Len(txt_f_addr) = 7 And Trim(txt_f_addr) <> "S0A0101" Then
       sQuery = "SELECT * FROM FP_STDYARD WHERE LOCATION = '" + txt_f_addr + "' AND YARD_KND = '00'"
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
       
          MsgBox "起始垛位号不正确，请重新输入！", vbCritical, "系统提示信息"
       End If
    End If
End Sub

Private Sub txt_f_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
        DD.rControl.Add Item:=txt_f_addr
        DD.rControl.Add Item:=txt_o_f_addr_nm
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        txt_o_f_addr.Text = txt_f_addr.Text
        Exit Sub
        
    End If


    If Len(Trim(txt_f_addr)) = txt_f_addr.MaxLength Then
        txt_o_f_addr_nm.Text = Gf_ComnNameFind(M_CN1, "C0013", Trim(txt_f_addr.Text), 2)
    Else
        txt_o_f_addr_nm.Text = ""
    End If
    
    If Len(Trim(txt_f_addr)) = 2 Then
       txt_o_f_addr.Text = txt_f_addr.Text
    Else
       txt_o_f_addr.Text = ""
    End If

End Sub

Public Function Sp_Process(Conn As ADODB.Connection, Scc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, icount, iProcessCount As Integer
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
    For icount = 1 To Scc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Scc.Item("Spread"), 0, icount))
            
            Case "Input", "Update"
            
                If Not MC Is Nothing Then
                    Call Gp_Sp_Move(icount, Scc, MC)
                End If
                
                'Maxlength Check
                sMesg = Gf_Sp_NeceCheck2(Scc.Item("Spread"), Scc.Item("mColumn"), icount, Scc.Item("nColumn"))
                        
                If Trim(sMesg) = "OK" Then
                    
                ElseIf Mid(sMesg, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Scc.Item("Spread"), icount, , vbYellow)
                    sMesg = Mid(sMesg, 6, Len(sMesg))
                    sMesg = sMesg + "长度不正确"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Sp_Process = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Scc.Item("Spread"), icount, , vbYellow)
                    sMesg = sMesg + "必须输入"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Sp_Process = False
                    Exit Function
                End If
        
        End Select
    
    Next icount
    
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
    For icount = 0 To Scc.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next icount
    
    'Ceate Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    For icount = 1 To Scc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        
        Select Case Trim(Gf_Sp_RcvData(Scc.Item("Spread"), 0, icount))
        
            Case "Input"
            
                adoCmd.Parameters(0).VALUE = "I"
                ProcessChk = "YES"
                
            Case "Update"
            
                adoCmd.Parameters(0).VALUE = "U"
                ProcessChk = "YES"
                
            Case "Delete"
            
                adoCmd.Parameters(0).VALUE = "D"
                ProcessChk = "YES"
            
        End Select
          
        If ProcessChk = "YES" Then
            
            'Parameters Setting
            For iCol = 1 To Scc.Item("iColumn").Count
            
                Scc.Item("Spread").Col = Scc.Item("iColumn").Item(iCol)
                
                Select Case Scc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Scc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = 0
                        Else
                            dTempFloat = Scc.Item("Spread").Text
                            adoCmd.Parameters(iCol).VALUE = STR(dTempFloat)
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Scc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = 0
                        Else
                            dTempInt = Scc.Item("Spread").Text
                            adoCmd.Parameters(iCol).VALUE = STR(dTempInt)
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Scc.Item("Spread").Text = "1" Then
                            adoCmd.Parameters(iCol).VALUE = "1"
                        Else
                            adoCmd.Parameters(iCol).VALUE = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Scc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = "0"
                        Else
                            adoCmd.Parameters(iCol).VALUE = Trim(STR(Scc.Item("Spread").VALUE))
                        End If
                        
                     Case SS_CELL_TYPE_DATE
                        If Trim(Scc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).VALUE = ""
                        Else
                            adoCmd.Parameters(iCol).VALUE = Mid(Trim(Scc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Scc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Scc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Scc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).VALUE = Trim(sTemp)
                        
                End Select
           
            Next iCol
                           
            iProcessCount = iProcessCount + 1
            
            adoCmd.Execute
            
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Scc.Item("Spread"), icount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Sp_Process = False
    
                Exit Function
        
             End If
        
        End If
        
    Next icount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For icount = 1 To Scc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Scc.Item("Spread"), 0, icount))
        
            Case "Input", "Update"
            
                Call Gp_Sp_SendData(Scc.Item("Spread"), "", 0, icount)
                
            Case "Delete"
                
                Call Gp_Sp_SendData(Scc.Item("Spread"), "", 0, icount)
                Call Gp_Sp_DeleteRow(Scc.Item("Spread"), icount)
                icount = icount - 1
            
        End Select
        
    Next icount
    
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
    
    If Len(txt_slab_no) = 10 And txt_f_addr <> "S0A0101" And S1_Click <> "1" Then
        txt_t_addr = ""
        txt_o_t_addr = ""
        txt_o_t_addr_nm = ""
    End If
End Sub

Private Sub txt_t_addr_Change()

Dim sQuery As String
    If Len(txt_t_addr) = 7 And Trim(txt_t_addr) <> "S0A0101" Then
       sQuery = "SELECT * FROM FP_STDYARD WHERE LOCATION = '" + txt_t_addr + "' AND YARD_KND = '00'"
       If Gf_FloatFind(M_CN1, sQuery) = 0 Then
       
          MsgBox "目的垛位号不正确，请重新输入！", vbCritical, "系统提示信息"
          Exit Sub
       End If
    End If
   
   If txt_f_addr <> "S0A0101" And Len(txt_t_addr) = 7 Then
      txt_slab_no = ""
   End If
End Sub

Private Sub txt_t_addr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        txt_t_addr.Text = "S"
        DD.sWitch = "MS"
        DD.sKey = "F0009"
        DD.rControl.Add Item:=txt_t_addr
        DD.rControl.Add Item:=txt_o_t_addr_nm
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        txt_o_t_addr.Text = txt_t_addr.Text
        Exit Sub
        
    End If

    If Len(Trim(txt_t_addr)) = txt_t_addr.MaxLength Then
        txt_o_t_addr_nm.Text = Gf_ComnNameFind(M_CN1, "F0009", Trim(txt_t_addr.Text), 2)
    Else
        txt_o_t_addr_nm.Text = ""
    End If
      

    If Len(Trim(txt_t_addr)) = 7 Then
       txt_o_t_addr.Text = txt_t_addr.Text
       If txt_f_addr <> "S0A0101" Then
          txt_slab_no = ""
       End If
    Else
       txt_o_t_addr.Text = ""
    End If

End Sub



