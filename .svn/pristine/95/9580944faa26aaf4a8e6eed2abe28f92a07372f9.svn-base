VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form DED1021C 
   Caption         =   "热处理对象查询及选定(钢板)_DED1021C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_cur_inv 
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
      Left            =   1410
      MaxLength       =   2
      TabIndex        =   20
      Top             =   90
      Width           =   495
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
      Left            =   1935
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   90
      Width           =   2010
   End
   Begin VB.TextBox Text_LOC 
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
      Left            =   13560
      MaxLength       =   15
      TabIndex        =   7
      Tag             =   "CD_MANA_NO"
      Top             =   90
      Width           =   1515
   End
   Begin VB.TextBox txt_cust_cd 
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
      Left            =   5640
      MaxLength       =   6
      TabIndex        =   6
      Top             =   540
      Width           =   915
   End
   Begin VB.TextBox txt_cust_cd_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6555
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "客户"
      Top             =   540
      Width           =   3645
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
      Left            =   1410
      MaxLength       =   18
      TabIndex        =   4
      Tag             =   "标准号"
      Top             =   540
      Width           =   2535
   End
   Begin VB.TextBox TXT_ORD_NO 
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
      Left            =   9870
      MaxLength       =   11
      TabIndex        =   3
      Top             =   90
      Width           =   1530
   End
   Begin VB.ComboBox CBO_ORD_ITEM 
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
      Left            =   11430
      TabIndex        =   2
      Top             =   90
      Width           =   720
   End
   Begin VB.TextBox txt_plt 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
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
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "生产工厂"
      Top             =   90
      Width           =   570
   End
   Begin VB.TextBox txt_plt_name 
      CausesValidation=   0   'False
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
      Left            =   6210
      TabIndex        =   0
      Tag             =   "机号"
      Top             =   90
      Width           =   1950
   End
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Left            =   8490
      Top             =   90
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
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   0
      Left            =   4320
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "生产工厂"
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
   Begin InDate.ULabel ULabel22 
      Height          =   315
      Index           =   1
      Left            =   90
      Top             =   540
      Width           =   1305
      _ExtentX        =   2302
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   4320
      Top             =   540
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "客户"
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
      Left            =   12480
      Top             =   90
      Width           =   1050
      _ExtentX        =   1852
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   570
      Left            =   90
      TabIndex        =   8
      Top             =   900
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   1005
      _Version        =   196609
      BackColor       =   14737918
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.ComboBox CBO_ORD_ITEM1 
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
         Left            =   14940
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox TXT_ORD_NO1 
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
         Left            =   14310
         MaxLength       =   11
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.ComboBox Cbo_n 
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
         ItemData        =   "DED1021C.frx":0000
         Left            =   9240
         List            =   "DED1021C.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Tag             =   "进程状态"
         Top             =   120
         Width           =   765
      End
      Begin VB.ComboBox Cbo_t 
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
         ItemData        =   "DED1021C.frx":0004
         Left            =   13350
         List            =   "DED1021C.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "进程状态"
         Top             =   120
         Width           =   765
      End
      Begin VB.ComboBox Cbo_q 
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
         ItemData        =   "DED1021C.frx":0008
         Left            =   11250
         List            =   "DED1021C.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "进程状态"
         Top             =   120
         Width           =   765
      End
      Begin VB.TextBox txt_cur_inv1 
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
         Left            =   5880
         MaxLength       =   2
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox text_cur_inv1 
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
         Left            =   6405
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   1440
      End
      Begin Threed.SSCheck chk_sel 
         Height          =   345
         Left            =   90
         TabIndex        =   16
         Top             =   120
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         _Version        =   196609
         Font3D          =   1
         BackColor       =   12632319
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "批次取消/选择"
      End
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   1770
         Top             =   120
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         Caption         =   "期限日期"
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
         ForeColor       =   255
      End
      Begin InDate.UDate SDT_PROD_DATE 
         Height          =   315
         Left            =   2910
         TabIndex        =   17
         Tag             =   "日期"
         Top             =   90
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   2
         Left            =   8370
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         Caption         =   "N;正火"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   3
         Left            =   12510
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         Caption         =   "T;回火"
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
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   4
         Left            =   10410
         Top             =   120
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         Caption         =   "Q;淬火"
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
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   4830
         Top             =   120
         Width           =   1065
         _ExtentX        =   1879
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
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   7695
      Left            =   90
      TabIndex        =   18
      Top             =   1500
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   13573
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
      MaxCols         =   24
      MaxRows         =   30
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "DED1021C.frx":000C
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   90
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "堆放仓库"
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
End
Attribute VB_Name = "DED1021C"
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
'-- Program Name      热处理装炉实绩查询_装炉时间
'-- Program ID        DGA1090C
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
    
    Dim lCol As Integer

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
    
           Call Gp_Ms_Collection(txt_cur_inv, "p", "N", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(Text_LOC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(CBO_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_stdspec_chg, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="DED1021C.P_REFER", Key:="P-R"
    sc1.Add Item:="DED1021C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
'
'    Call Gp_Sp_ColHidden(ss1, 2, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub
Private Sub chk_sel_Click(Value As Integer)

    Dim iRow As Integer
    Dim TIME As String
    
    TIME = Format(Now, "YYYY-MM-DD")
    
    If SDT_PROD_DATE.RawData = "" Then
           MsgBox "请输入期限日期......!"
           Exit Sub
    End If
    
    If SDT_PROD_DATE.Text < TIME Then
      MsgBox "输入的期限日期不能小于当前系统日期......!"
      Exit Sub
    End If

    If txt_cur_inv1.Text = "" Then
       MsgBox "请输入堆放仓库......!"
       Exit Sub
    End If
    If txt_cur_inv1.Text <> "00" And txt_cur_inv1.Text <> "WG" And txt_cur_inv1.Text <> "WD" Then
          MsgBox "请输入正确的堆放仓库......!"
          Exit Sub
    End If
    
    
    If chk_sel Then
        For iRow = 1 To ss1.MaxRows
            ss1.Row = iRow
            ss1.Col = 3
            If ss1.Text <> txt_cur_inv1 Then
               ss1.Col = 0

            Else
                ss1.Col = 0
                ss1.Text = "Update"
       
                  ss1.Col = 3
                  If ss1.Text = txt_cur_inv1 Then
                  
                       ss1.Col = 5
                       If Len(ss1.Text) > 2 Then
                          ss1.Col = 7
                          ss1.Text = ""
                       ElseIf Len(ss1.Text) = 2 Then
                           ss1.Col = 5
                           If Mid(ss1.Text, 1, 1) = "N" Then
                           ss1.Col = 7
                           ss1.Text = Cbo_n
                        
                            ElseIf Mid(ss1.Text, 1, 1) = "Q" Then
                                  ss1.Col = 7
                                  ss1.Text = Cbo_q
                            ElseIf Mid(ss1.Text, 1, 1) = "T" Then
                                  ss1.Col = 7
                                  ss1.Text = Cbo_t
                            End If
                            
                         ElseIf Len(ss1.Text) = 0 Then
                         
                            ss1.Col = 7
                            ss1.Text = ""
                            
                     End If
                     
                     
                   ss1.Col = 8
                   If Len(ss1.Text) > 2 Then
                          ss1.Col = 10
                          ss1.Text = ""
                          
                    ElseIf Len(ss1.Text) = 2 Then
                    
                          ss1.Col = 5
                          If Len(ss1.Text) < 2 Then
                              ss1.Col = 10
                              ss1.Text = ""

                           ElseIf Len(ss1.Text) >= 2 Then
                           
                                    ss1.Col = 8
                            
                                    If Mid(ss1.Text, 1, 1) = "N" Then
                                          ss1.Col = 10
                                          ss1.Text = Cbo_n
                                          
                                     ElseIf Mid(ss1.Text, 1, 1) = "Q" Then
                                          ss1.Col = 10
                                          ss1.Text = Cbo_q
                                     ElseIf Mid(ss1.Text, 1, 1) = "T" Then
                                          ss1.Col = 10
                                          ss1.Text = Cbo_t
                                     End If
                               End If
                                     
                    ElseIf Len(ss1.Text) < 2 Then
                         
                         ss1.Col = 10
                         ss1.Text = ""
                    End If
                    
                         
                   ss1.Col = 11
                   If Len(ss1.Text) > 2 Then
                          ss1.Col = 13
                          ss1.Text = ""

                    ElseIf Len(ss1.Text) = 2 Then

                                  ss1.Col = 8
                               If Len(ss1.Text) < 2 Then
                                  ss1.Col = 13
                                  ss1.Text = ""

                                ElseIf Len(ss1.Text) >= 2 Then
                                
                                      ss1.Col = 11
                                      If Mid(ss1.Text, 1, 1) = "N" Then
                                            ss1.Col = 13
                                            ss1.Text = Cbo_n

                                       ElseIf Mid(ss1.Text, 1, 1) = "Q" Then
                                            ss1.Col = 13
                                            ss1.Text = Cbo_q
                                       ElseIf Mid(ss1.Text, 1, 1) = "T" Then
                                            ss1.Col = 13
                                            ss1.Text = Cbo_t
                                       End If
                                 End If

                    ElseIf Len(ss1.Text) < 2 Then

                         ss1.Col = 13
                         ss1.Text = ""
                    End If
                  
                     ss1.Col = 23
                     ss1.Text = SDT_PROD_DATE.Text
                     ss1.Col = 24
                    ss1.Text = sUserID
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)
                End If
            End If
          
        Next iRow
    Else
        For iRow = 1 To ss1.MaxRows
            ss1.Row = iRow
            ss1.Col = 0
            ss1.Text = ""
            ss1.Col = 7
            ss1.Text = ""
            ss1.Col = 10
            ss1.Text = ""
            ss1.Col = 13
            ss1.Text = ""
            ss1.Col = 23
            ss1.Text = ""
            ss1.Col = 24
            ss1.Text = ""
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow)
        Next iRow

    End If
    
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
 
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "EG-System.INI", Me.Name)
        
    Screen.MousePointer = vbDefault
    
    txt_cur_inv.Text = "00"
    text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", Trim(txt_cur_inv.Text), "1")
    
    txt_cur_inv1.Text = "00"
    text_cur_inv1.Text = Gf_ComnNameFind(M_CN1, "C0013", Trim(txt_cur_inv1.Text), "1")

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "EG-System.INI", Me.Name)
    
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
        
        SDT_PROD_DATE.RawData = ""
        txt_cur_inv1.Text = ""
        text_cur_inv.Text = ""
        chk_sel.Value = ssCBUnchecked
    End If

End Sub

Public Sub Form_Ref()
Dim CNT As Integer
Dim WGT As Double
Dim i As Integer
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
       
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
    
End Sub

Public Sub Form_Pro()
  Dim iRow As Integer
Dim aa As Integer


    For iRow = 0 To ss1.MaxRows

                ss1.Row = iRow
                ss1.Col = 0

        If ss1.Text = "Update" Then


            ss1.Col = 3
            If ss1.Text = txt_cur_inv1 Then

                                  ss1.Col = 8
                                  If Len(ss1.Text) > 2 Then
                                     ss1.Col = 10
                                     ss1.Text = ""

                                  ElseIf Len(ss1.Text) = 2 Then
                                     
                                            ss1.Col = 5
                                            If Len(ss1.Text) = 2 Then
                                            ss1.Col = 7
                                              If ss1.Text = "" Then
                                                 ss1.Col = 10
                                                 MsgBox "请先做热处理1......!"
                                                 Exit Sub
                                            End If
                                            End If
                                  ElseIf Len(ss1.Text) < 2 Then
                                         
                                          ss1.Col = 10
                                          ss1.Text = ""
'
                                          
                                   End If
                                   
                                  ss1.Col = 11
                                  If Len(ss1.Text) > 2 Then
                                     ss1.Col = 13
                                     ss1.Text = ""

                                  ElseIf Len(ss1.Text) = 2 Then
                                          
                                            ss1.Col = 8
                                            If Len(ss1.Text) = 2 Then
                                            ss1.Col = 10
                                            If ss1.Text = "" Then
                                               ss1.Col = 13
                                               MsgBox "请先做热处理2......!"
                                               Exit Sub
                                            End If
                                            End If
                                  ElseIf Len(ss1.Text) < 2 Then
                                         
                                          ss1.Col = 13
                                          ss1.Text = ""
'
                                          
                                   End If
   
             End If

         End If

      Next iRow

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
    
End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

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
'        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Sub SDT_PROD_DATE_LostFocus()
  Dim iRow As Integer
  Dim TIME As String


       For iRow = 0 To ss1.MaxRows
           ss1.Row = iRow
           ss1.Col = 0
          If ss1.Text = "Update" Then
             ss1.Col = 23
             ss1.Text = SDT_PROD_DATE.Text
          End If
       Next iRow
       
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim iRow As Integer
    Dim i As Integer
    Dim TIME As String
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

   If BlockRow < 0 Then Exit Sub

    If SDT_PROD_DATE.RawData = "" Then
           MsgBox "请输入期限日期......!"
           Exit Sub
    End If
    
    
    TIME = Format(Now, "YYYY-MM-DD")
    
    If SDT_PROD_DATE.Text < TIME Then
      MsgBox "输入的期限日期不能小于当前系统日期......!"
      Exit Sub
    End If

    If txt_cur_inv1.Text = "" Then
       MsgBox "请输入堆放仓库......!"
       Exit Sub
    End If
    If txt_cur_inv1.Text <> "00" And txt_cur_inv1.Text <> "WG" And txt_cur_inv1.Text <> "WD" Then
          MsgBox "只堆放仓库--> '00' / 'WD' / 'WG'"
          Exit Sub
    End If
    
   

'    If Gf_Sc_Authority(sAuthority, "U") Then

        For iRow = BlockRow To BlockRow2

            ss1.Row = iRow
            ss1.Col = 0

            If ss1.Text = "Update" Then
                ss1.Text = ""
                ss1.Col = 7
                ss1.Text = ""
                ss1.Col = 10
                ss1.Text = ""
                ss1.Col = 13
                ss1.Text = ""
                ss1.Col = 23
                ss1.Text = ""
                ss1.Col = 24
                ss1.Text = ""
         Else
                ss1.Col = 3
                If ss1.Text <> txt_cur_inv1 Then
                   ss1.Col = 0
                Else
                ss1.Col = 0
                ss1.Text = "Update"
               
                  ss1.Col = 3
                  If ss1.Text = txt_cur_inv1 Then
                  
                           ss1.Col = 5
                       If Len(ss1.Text) > 2 Then
                          ss1.Col = 7
                          ss1.Text = ""
                       ElseIf Len(ss1.Text) = 2 Then
                           ss1.Col = 5
                           If Mid(ss1.Text, 1, 1) = "N" Then
                           ss1.Col = 7
                           ss1.Text = Cbo_n
                        
                            ElseIf Mid(ss1.Text, 1, 1) = "Q" Then
                                  ss1.Col = 7
                                  ss1.Text = Cbo_q
                            ElseIf Mid(ss1.Text, 1, 1) = "T" Then
                                  ss1.Col = 7
                                  ss1.Text = Cbo_t
                            End If
                            
                         ElseIf Len(ss1.Text) = 0 Then
                         
                            ss1.Col = 7
                            ss1.Text = ""
                            
                     End If
                     
                     
                   ss1.Col = 8
                   If Len(ss1.Text) > 2 Then
                          ss1.Col = 10
                          ss1.Text = ""
                          
                    ElseIf Len(ss1.Text) = 2 Then
                    
                          ss1.Col = 5
                          If Len(ss1.Text) < 2 Then
                              ss1.Col = 10
                              ss1.Text = ""

                           ElseIf Len(ss1.Text) >= 2 Then
                           
                                    ss1.Col = 8
                            
                                    If Mid(ss1.Text, 1, 1) = "N" Then
                                          ss1.Col = 10
                                          ss1.Text = Cbo_n
                                          
                                     ElseIf Mid(ss1.Text, 1, 1) = "Q" Then
                                          ss1.Col = 10
                                          ss1.Text = Cbo_q
                                     ElseIf Mid(ss1.Text, 1, 1) = "T" Then
                                          ss1.Col = 10
                                          ss1.Text = Cbo_t
                                     End If
                               End If
                                     
                    ElseIf Len(ss1.Text) < 2 Then
                         
                         ss1.Col = 10
                         ss1.Text = ""
                    End If
                    
                         
                   ss1.Col = 11
                   If Len(ss1.Text) > 2 Then
                          ss1.Col = 13
                          ss1.Text = ""

                    ElseIf Len(ss1.Text) = 2 Then

                                  ss1.Col = 8
                               If Len(ss1.Text) < 2 Then
                                  ss1.Col = 13
                                  ss1.Text = ""

                                ElseIf Len(ss1.Text) >= 2 Then
                                
                                      ss1.Col = 11
                                      If Mid(ss1.Text, 1, 1) = "N" Then
                                            ss1.Col = 13
                                            ss1.Text = Cbo_n

                                       ElseIf Mid(ss1.Text, 1, 1) = "Q" Then
                                            ss1.Col = 13
                                            ss1.Text = Cbo_q
                                       ElseIf Mid(ss1.Text, 1, 1) = "T" Then
                                            ss1.Col = 13
                                            ss1.Text = Cbo_t
                                       End If
                                 End If

                    ElseIf Len(ss1.Text) < 2 Then

                         ss1.Col = 13
                         ss1.Text = ""
                    End If
                  
                  
                        ss1.Col = 23
                        ss1.Text = SDT_PROD_DATE.Text
                        ss1.Col = 24
                        ss1.Text = sUserID
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)
               End If
            End If
           End If
        Next iRow



End Sub

'Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
'
'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0
'
'End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    If Row > 0 And (Col = 38 Or Col = 39) Then
       ss1.Row = Row
       ss1.Col = Col
       ss1.Value = Gf_DTSet(M_CN1, , "X")
    End If
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        ss1.Row = ss1.ActiveRow
        ss1.Col = 36
        ss1.Text = sUserID
    End If

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

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
       Set Active_Spread = Me.ss1
       PopupMenu MDIMain.PopUp_Spread
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

    Else

        If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_name.Text = ""
        End If
    
    End If
    
End Sub

Private Sub txt_stdspec_chg_DblClick()
    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stdspec_chg

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

    End If
    
End Sub
Private Sub txt_cust_cd_DblClick()

    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_cd_name

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_cd_name.Text = ""
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
Private Sub txt_ord_no_LostFocus()

    If TXT_ORD_NO.Text <> "" Then
       If (Len(TXT_ORD_NO.Text) < TXT_ORD_NO.MaxLength) Then
          Call Gp_MsgBoxDisplay("订单号输入未完成！")
          CBO_ORD_ITEM.Text = ""
          TXT_ORD_NO.SetFocus
       End If
    End If

End Sub

Private Sub txt_cur_inv1_Change()
 
  If Len(Trim(txt_cur_inv1.Text)) = txt_cur_inv1.MaxLength Then
          text_cur_inv1.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv1.Text, 2)
          Exit Sub
    Else
          text_cur_inv1.Text = ""
    End If
End Sub

Private Sub txt_cur_inv1_DblClick()

    Call txt_cur_inv1_KeyUp(vbKeyF4, 0)
    
End Sub
Private Sub txt_cur_inv1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
    
        DD.rControl.Add Item:=txt_cur_inv1
        DD.rControl.Add Item:=text_cur_inv1
        
    
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
       
        If Len(Trim(txt_cur_inv1.Text)) = txt_cur_inv1.MaxLength Then
            text_cur_inv1.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv1.Text, 2)
            Exit Sub
        Else
            text_cur_inv1.Text = ""
        End If
    End If
End Sub
Private Sub txt_cur_inv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
    
        DD.rControl.Add Item:=txt_cur_inv
        DD.rControl.Add Item:=text_cur_inv
        
    
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
       
        If Len(Trim(txt_cur_inv.Text)) = txt_cur_inv.MaxLength Then
            text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv.Text, 2)
            Exit Sub
        Else
            text_cur_inv.Text = ""
        End If
    End If
End Sub

Private Sub txt_cur_inv_Change()



    If txt_cur_inv.Text = "00" Then
    txt_cur_inv1.Text = "00"
    Cbo_n.Clear
    
    Cbo_n.AddItem ""
    Cbo_n.AddItem "1"
    Cbo_n.AddItem "2"
    Cbo_n.ListIndex = 1
    
    Cbo_q.Clear
    Cbo_q.AddItem ""
    Cbo_q.AddItem "1"
'    Cbo_q.AddItem "2"
    Cbo_q.ListIndex = 1
    
    Cbo_t.Clear
    Cbo_t.AddItem ""
    Cbo_t.AddItem "1"
    Cbo_t.AddItem "2"
    Cbo_t.ListIndex = 1
   

  ElseIf txt_cur_inv.Text = "WG" Then
     txt_cur_inv1.Text = "WG"
     Cbo_n.Clear
     Cbo_n.AddItem ""
     Cbo_n.AddItem "1"
     Cbo_n.AddItem "2"
     Cbo_n.ListIndex = 2
     Cbo_q.Clear
     Cbo_q.AddItem ""
     Cbo_q.AddItem "1"
'     Cbo_q.AddItem "2"
     Cbo_q.ListIndex = 0
     
     
     Cbo_t.Clear
     Cbo_t.AddItem ""
     Cbo_t.AddItem "1"
     Cbo_t.AddItem "2"
     Cbo_t.ListIndex = 2
     
  ElseIf txt_cur_inv.Text = "WD" Then
     txt_cur_inv1.Text = "WD"
     Cbo_n.Clear
     Cbo_n.AddItem ""
     Cbo_n.AddItem "1"
     Cbo_n.AddItem "2"
     Cbo_n.ListIndex = 2
     Cbo_q.Clear
     Cbo_q.AddItem ""
     Cbo_q.AddItem "1"
'     Cbo_q.AddItem "2"
     Cbo_q.ListIndex = 0
     
     
     Cbo_t.Clear
     Cbo_t.AddItem ""
     Cbo_t.AddItem "1"
     Cbo_t.AddItem "2"
     Cbo_t.ListIndex = 2
    
  End If
  
   If Len(Trim(txt_cur_inv.Text)) = txt_cur_inv.MaxLength Then
          text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_cur_inv.Text, 2)
          Exit Sub
   Else
          text_cur_inv.Text = ""
   End If


  
End Sub

Private Sub txt_cur_inv_DblClick()

   
   
    If txt_cur_inv.Text = "00" Then
       txt_cur_inv1.Text = "00"
   
    Cbo_n.Clear
    Cbo_n.AddItem "1"
    Cbo_n.AddItem "2"
    Cbo_n.ListIndex = 0
    
    Cbo_q.Clear
    Cbo_q.AddItem "1"
    Cbo_q.AddItem "2"
    Cbo_q.ListIndex = 0
    Cbo_t.Clear
    Cbo_t.AddItem "1"
    Cbo_t.AddItem "2"
    Cbo_t.ListIndex = 0
     
  ElseIf txt_cur_inv.Text = "WG" Then
     txt_cur_inv1.Text = "WG"
     Cbo_n.Clear
     Cbo_n.AddItem "1"
     Cbo_n.AddItem "2"
     Cbo_n.ListIndex = 1
     Cbo_q.Clear
     Cbo_q.AddItem "1"
     Cbo_q.AddItem "2"
     Cbo_q.ListIndex = 1
     Cbo_t.Clear
     Cbo_t.AddItem "1"
     Cbo_t.AddItem "2"
     Cbo_t.ListIndex = 1
     
   ElseIf txt_cur_inv.Text = "WD" Then
     txt_cur_inv1.Text = "WD"
     Cbo_n.Clear
     Cbo_n.AddItem "1"
     Cbo_n.AddItem "2"
     Cbo_n.ListIndex = 1
     Cbo_q.Clear
     Cbo_q.AddItem "1"
     Cbo_q.AddItem "2"
     Cbo_q.ListIndex = 1
     Cbo_t.Clear
     Cbo_t.AddItem "1"
     Cbo_t.AddItem "2"
     Cbo_t.ListIndex = 1
     
  
  End If
  
   Call txt_cur_inv_KeyUp(vbKeyF4, 0)
    
     
End Sub


Private Sub Cbo_n_Click()

 Dim iRow As Integer
  
        For iRow = 0 To ss1.MaxRows
               ss1.Row = iRow
               ss1.Col = 0
            If ss1.Text = "Update" Then
               ss1.Col = 3
                If ss1.Text = txt_cur_inv1.Text Then

                      ss1.Col = 5
                      If Mid(ss1.Text, 1, 1) = "N" Then
                          ss1.Col = 7
                          ss1.Text = Cbo_n
                      End If
                      ss1.Col = 8
                      If Mid(ss1.Text, 1, 1) = "N" Then
                          ss1.Col = 10
                          ss1.Text = Cbo_n

                      End If

                       ss1.Col = 11
                      If Mid(ss1.Text, 1, 1) = "N" Then
                          ss1.Col = 13
                          ss1.Text = Cbo_n

                      End If

                 End If

            End If


        Next iRow

End Sub
Private Sub Cbo_q_Click()

 Dim iRow As Integer
 
      For iRow = 0 To ss1.MaxRows
               ss1.Row = iRow
               ss1.Col = 0
            If ss1.Text = "Update" Then
               ss1.Col = 3
                If ss1.Text = txt_cur_inv1.Text Then
                
                           ss1.Col = 5
                           If Mid(ss1.Text, 1, 1) = "Q" Then
                               ss1.Col = 7
                               ss1.Text = Cbo_q
                           End If

                           ss1.Col = 8
                           If Mid(ss1.Text, 1, 1) = "Q" Then
                              ss1.Col = 10
                              ss1.Text = Cbo_q
                           End If

                           ss1.Col = 11
                           If Mid(ss1.Text, 1, 1) = "Q" Then
                              ss1.Col = 13
                              ss1.Text = Cbo_q
                           End If
                       
                 End If
                 
            End If
             

        Next iRow

  
End Sub
Private Sub Cbo_t_Click()

 Dim iRow As Integer
    
        For iRow = 0 To ss1.MaxRows
               ss1.Row = iRow
               ss1.Col = 0
            If ss1.Text = "Update" Then
               ss1.Col = 3
                If ss1.Text = txt_cur_inv1.Text Then
                
                      ss1.Col = 5
                      If Mid(ss1.Text, 1, 1) = "T" Then
                          ss1.Col = 7
                          ss1.Text = Cbo_t
                      End If
                      
                      ss1.Col = 8
                      If Mid(ss1.Text, 1, 1) = "T" Then
                          ss1.Col = 10
                          ss1.Text = Cbo_t
                      End If
                      
                       ss1.Col = 11
                      If Mid(ss1.Text, 1, 1) = "T" Then
                          ss1.Col = 13
                          ss1.Text = Cbo_t
                      End If
                
                 End If
                 
            End If
             

        Next iRow

End Sub
Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow As Integer
    Dim TIME As String
    
 
    If ss1.MaxRows < 1 Then Exit Sub
    
    lBlkcol1 = Col
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

  TIME = Format(Now, "YYYY-MM-DD")

  If Row < 0 Then Exit Sub
  
  If SDT_PROD_DATE.RawData = "" Then
         MsgBox "请输入期限日期......!"
         Exit Sub
  End If
    
  If SDT_PROD_DATE.Text < TIME Then
      MsgBox "输入的期限日期不能小于当前系统日期......!"
      Exit Sub
  End If

  If txt_cur_inv1.Text = "" Then
     MsgBox "请输入堆放仓库......!"
     Exit Sub
  End If
  If txt_cur_inv1.Text <> "00" And txt_cur_inv1.Text <> "WG" And txt_cur_inv1.Text <> "WD" Then
        MsgBox "请输入正确的堆放仓库......!"
        Exit Sub
  End If

     
 If Col = 1 Then
     
    ss1.Row = Row
    ss1.Col = 0
  
    
   
    If ss1.Text = "Update" Then
        
        ss1.Text = Row
        ss1.Col = 7
        ss1.Text = ""
        ss1.Col = 10
        ss1.Text = ""
        ss1.Col = 13
        ss1.Text = ""
        ss1.Col = 23
        ss1.Text = ""
        ss1.Col = 24
        ss1.Text = ""
        
    Else
    
            ss1.Col = 3
            If ss1.Text <> txt_cur_inv1 Then
               ss1.Col = 0

            Else
                ss1.Col = 0
                ss1.Text = "Update"
       
                  ss1.Col = 3
                  If ss1.Text = txt_cur_inv1 Then
                  
                        ss1.Col = 5
                       If Len(ss1.Text) > 2 Then
                          ss1.Col = 7
                          ss1.Text = ""
                       ElseIf Len(ss1.Text) = 2 Then
                           ss1.Col = 5
                           If Mid(ss1.Text, 1, 1) = "N" Then
                           ss1.Col = 7
                           ss1.Text = Cbo_n
                        
                            ElseIf Mid(ss1.Text, 1, 1) = "Q" Then
                                  ss1.Col = 7
                                  ss1.Text = Cbo_q
                            ElseIf Mid(ss1.Text, 1, 1) = "T" Then
                                  ss1.Col = 7
                                  ss1.Text = Cbo_t
                            End If
                            
                         ElseIf Len(ss1.Text) = 0 Then
                         
                            ss1.Col = 7
                            ss1.Text = ""
                            
                     End If
                     
                     
                   ss1.Col = 8
                   If Len(ss1.Text) > 2 Then
                          ss1.Col = 10
                          ss1.Text = ""
                          
                    ElseIf Len(ss1.Text) = 2 Then
                    
                          ss1.Col = 5
                          If Len(ss1.Text) < 2 Then
                              ss1.Col = 10
                              ss1.Text = ""

                           ElseIf Len(ss1.Text) >= 2 Then
                           
                                    ss1.Col = 8
                            
                                    If Mid(ss1.Text, 1, 1) = "N" Then
                                          ss1.Col = 10
                                          ss1.Text = Cbo_n
                                          
                                     ElseIf Mid(ss1.Text, 1, 1) = "Q" Then
                                          ss1.Col = 10
                                          ss1.Text = Cbo_q
                                     ElseIf Mid(ss1.Text, 1, 1) = "T" Then
                                          ss1.Col = 10
                                          ss1.Text = Cbo_t
                                     End If
                               End If
                                     
                    ElseIf Len(ss1.Text) < 2 Then
                         
                         ss1.Col = 10
                         ss1.Text = ""
                    End If
                    
                         
                   ss1.Col = 11
                   If Len(ss1.Text) > 2 Then
                          ss1.Col = 13
                          ss1.Text = ""

                    ElseIf Len(ss1.Text) = 2 Then

                                  ss1.Col = 8
                               If Len(ss1.Text) < 2 Then
                                  ss1.Col = 13
                                  ss1.Text = ""

                                ElseIf Len(ss1.Text) >= 2 Then
                                
                                      ss1.Col = 11
                                      If Mid(ss1.Text, 1, 1) = "N" Then
                                            ss1.Col = 13
                                            ss1.Text = Cbo_n

                                       ElseIf Mid(ss1.Text, 1, 1) = "Q" Then
                                            ss1.Col = 13
                                            ss1.Text = Cbo_q
                                       ElseIf Mid(ss1.Text, 1, 1) = "T" Then
                                            ss1.Col = 13
                                            ss1.Text = Cbo_t
                                       End If
                                 End If

                    ElseIf Len(ss1.Text) < 2 Then

                         ss1.Col = 13
                         ss1.Text = ""
                    End If
                  
                      ss1.Col = 23
                      ss1.Text = SDT_PROD_DATE.Text
                      ss1.Col = 24
                      ss1.Text = sUserID
                      Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)
             End If
           End If
    End If
  End If
End Sub
