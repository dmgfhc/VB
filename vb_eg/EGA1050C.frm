VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form EGA1050C 
   Caption         =   "火切实绩查询及修改_EGA1050C"
   ClientHeight    =   10395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   15900
      Top             =   1710
   End
   Begin VB.TextBox txt_PrcLine 
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
      Height          =   330
      Left            =   0
      MaxLength       =   10
      TabIndex        =   10
      Tag             =   "产线别"
      Text            =   "5"
      Top             =   0
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txt_Loc 
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
      Height          =   330
      Left            =   13530
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "生产工厂"
      Top             =   240
      Width           =   1770
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   465
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "GulimChe"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox TXT_MILL_LOT_NO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4065
         MaxLength       =   14
         TabIndex        =   1
         Top             =   90
         Width           =   1830
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   2835
         Top             =   90
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         Caption         =   "轧批号"
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
      Begin VB.TextBox TXT_INQNO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8010
         MaxLength       =   14
         TabIndex        =   2
         Tag             =   "材料号"
         Top             =   90
         Width           =   1890
      End
      Begin VB.TextBox txt_WkPlt 
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
         Left            =   1485
         Locked          =   -1  'True
         TabIndex        =   6
         Tag             =   "生产工厂"
         Text            =   "C3"
         Top             =   90
         Width           =   420
      End
      Begin VB.TextBox TXT_PRODCD 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   14700
         TabIndex        =   5
         Tag             =   "产品代码"
         Top             =   -30
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox TXT_MPLATE_NO 
         Height          =   330
         Left            =   14715
         MaxLength       =   14
         TabIndex        =   4
         Tag             =   "材料号"
         Top             =   90
         Visible         =   0   'False
         Width           =   300
      End
      Begin InDate.ULabel ULabel10 
         Height          =   300
         Left            =   240
         Top             =   90
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         Caption         =   "生产工厂"
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
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   6780
         Top             =   90
         Width           =   1200
         _ExtentX        =   2117
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
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   13740
         TabIndex        =   7
         Top             =   225
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   12300
      Top             =   240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "垛位"
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
   Begin SSSplitter.SSSplitter SP1 
      Height          =   7635
      Left            =   90
      TabIndex        =   8
      Top             =   1500
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   13467
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "EGA1050C.frx":0000
      Begin FPSpread.vaSpread ss2 
         Height          =   3495
         Left            =   0
         TabIndex        =   9
         Top             =   4140
         Width           =   15210
         _Version        =   393216
         _ExtentX        =   26829
         _ExtentY        =   6165
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "EGA1050C.frx":0052
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   4050
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   15210
         _Version        =   393216
         _ExtentX        =   26829
         _ExtentY        =   7144
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColsFrozen      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   40
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "EGA1050C.frx":0909
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   375
      Left            =   4410
      TabIndex        =   11
      Top             =   660
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1#侧标"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   60
         Width           =   975
      End
      Begin VB.Shape tcpStatus 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   1140
         Shape           =   3  'Circle
         Top             =   75
         Width           =   285
      End
      Begin VB.Label tcpMsg 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1560
         TabIndex        =   13
         Top             =   105
         Width           =   2055
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   15900
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "191.168.1.100"
      RemotePort      =   5080
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   375
      Left            =   8730
      TabIndex        =   14
      Top             =   660
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2#侧标"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   60
         Width           =   975
      End
      Begin VB.Label tcpMsg2 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1560
         TabIndex        =   16
         Top             =   105
         Width           =   2055
      End
      Begin VB.Shape tcpStatus2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   1140
         Shape           =   3  'Circle
         Top             =   75
         Width           =   285
      End
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   16530
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "191.168.1.100"
      RemotePort      =   5080
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   375
      Left            =   90
      TabIndex        =   17
      Top             =   660
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "表喷"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   18
         Top             =   60
         Width           =   975
      End
      Begin VB.Label tcpMsg3 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1560
         TabIndex        =   19
         Top             =   105
         Width           =   2055
      End
      Begin VB.Shape tcpStatus3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   1140
         Shape           =   3  'Circle
         Top             =   75
         Width           =   285
      End
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   17160
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "191.168.1.100"
      RemotePort      =   5080
   End
   Begin Threed.SSFrame SSFrame5 
      Height          =   375
      Left            =   90
      TabIndex        =   21
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.CheckBox chk_Cond 
         BackColor       =   &H00E0E0E0&
         Caption         =   "钢印"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   22
         Top             =   60
         Width           =   975
      End
      Begin VB.Shape tcpStatus4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         Height          =   225
         Left            =   1140
         Shape           =   3  'Circle
         Top             =   75
         Width           =   285
      End
      Begin VB.Label tcpMsg4 
         BackColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1560
         TabIndex        =   23
         Top             =   105
         Width           =   2055
      End
   End
   Begin MSWinsockLib.Winsock Winsock4 
      Left            =   16530
      Top             =   1710
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "191.168.1.100"
      RemotePort      =   5080
   End
   Begin Threed.SSCommand SSSend1 
      Height          =   405
      Left            =   12810
      TabIndex        =   24
      Top             =   1020
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   714
      _Version        =   196609
      ForeColor       =   255
      Caption         =   "标印发送"
   End
End
Attribute VB_Name = "EGA1050C"
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
'-- Program Name      火切实绩查询及修改
'-- Program ID        EGA1050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2010.7.20
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER            DATE          EDITOR       DESCRIPTION
'-- 1.01           2010.7.20     GUOLI
'-- 1.02           2011.7.14     LIQIAN       侧喷信息发送侧喷设备
'-- 1.03           2011.11.25    LIQIAN       标印信息发送表喷设备
'-- 1.04           2012.08.08    LIQIAN       钢印信息发送至钢印机
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public sDateTime As String          'Active Form Time Setting
Public sQuery_Rt As String          'Active Form sQuery Setting

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



Dim Mc1 As New Collection           'Master Collectionn
Dim Mc2 As New Collection           'Master Collectionn

Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_PLATE_NO = 1
Const SPD_UST = 9
Const SPD_DS_CUT_END_DATE = 12
Const SPD_DATE = 13
Const SPD_THK = 14
Const SPD_WID = 15
Const SPD_LEN = 16
Const SPD_WGT = 17
Const SPD_DS_LAST_YN = 18
Const SPD_SURF_GRD = 16
Const SPD_TRIM_FL = 17
Const SPD_DS_KNIFE_GAP = 18
Const SPD_EMP_CD = 21


Const SPD_PROC_CD = 23
Const SPD_END_USE = 24
Const SPD_GROUP = 25

'取钢种发送侧喷设备, 2011-07-14 , by Liqian
Const SPD_STLGRD_CP = 33
'取标准判断是否船板, 2011-07-14 , by Liqian
Const SPD_APLY_STDSPEC = 34
'取客户发送侧喷设备, 2011-07-14 , by Liqian
Const SPD_CUST_CD = 35
'取打印标准发送侧喷设备, 2011-07-14 , by Liqian
Const SPD_STDSPEC_YY = 36
Const SPD_VESSEL_NO = 37
Const SPD_SIDEMARK = 38 '侧喷加喷
Const SPD_SEALMEMO = 39 '加冲钢印
Const SPD_CUST_CD_SHORT = 40


Dim sQuery   As String
Dim sLoopFl  As String

Dim Screen_Fl As Boolean

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"
       
    'MASTER Collection
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_WkPlt, "p", " ", " ", " ", " ", " ", "l", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(TXT_INQNO, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
  Call Gp_Ms_Collection(TXT_MILL_LOT_NO, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          
    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=cControl1, Key:="cControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="EGA1050C.P_REFER1", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(TXT_MPLATE_NO, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(TXT_PRODCD, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
        
    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", "n", "m", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    '取钢种发送侧喷设备, 2011-07-14 , by Liqian
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    '取标准发送侧喷设备, 2011-07-14 , by Liqian
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    '取客户发送侧喷设备, 2011-07-14 , by Liqian
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    '取打印标准种发送侧喷设备, 2011-07-14 , by Liqian
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '侧喷加喷 20150122
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '加冲钢印 20150122
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="EGA1050C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="EGA1050C.P_SREFER1", Key:="P-R"
    sc1.Add Item:="EGA1050C.P_ONEROW", Key:="P-O"
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
    
    Call Gp_Sp_ColHidden(ss1, 27, True)
'    Call Gp_Sp_ColHidden(ss1, 28, True)
        
    Screen_Fl = False
     
End Sub

Public Sub Form_Exc()
ss1.ROW = 0
ss1.Col = 0
If ss1.Text = "◎" Then
    Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End If

ss2.ROW = 0
ss2.Col = 0
If ss2.Text = "◎" Then
    Call Gp_Sp_Excel(Me, ss2, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End If

End Sub

Private Sub chk_Cond_Click(Index As Integer)

    Dim strState As String
    Dim strState2 As String
    Dim strState3 As String
    Dim strState4 As String

       If Index = 0 Then
          If chk_Cond(0) = 1 Then
              Winsock1.Connect
          Else
              Winsock1.Close
              strState = "连接断线"
              tcpStatus.BackColor = &HFF&
              chk_Cond(0).ForeColor = &HFF&
              tcpMsg.Caption = "侧喷机状态 : " & strState
          End If
       End If
       
       If Index = 1 Then
          If chk_Cond(1) = 1 Then
             Winsock2.Connect
          Else
             Winsock2.Close
             strState2 = "连接断线"
             tcpStatus2.BackColor = &HFF&
             chk_Cond(1).ForeColor = &HFF&
             tcpMsg2.Caption = "侧喷机状态 : " & strState2
       End If
     End If
     
     If Index = 2 Then
          If chk_Cond(2) = 1 Then
             Winsock3.Connect
          Else
             Winsock3.Close
             strState3 = "连接断线"
             tcpStatus3.BackColor = &HFF&
             chk_Cond(2).ForeColor = &HFF&
             tcpMsg3.Caption = "标印机状态 : " & strState3
       End If
     End If
     
     If Index = 3 Then
          If chk_Cond(3) = 1 Then
             Winsock4.Connect
          Else
             Winsock4.Close
             strState4 = "连接断线"
             tcpStatus4.BackColor = &HFF&
             chk_Cond(3).ForeColor = &HFF&
             tcpMsg4.Caption = "钢印机状态 : " & strState4
       End If
     End If
    
End Sub

Private Sub SSSend1_Click()

Dim iCount      As Integer

Dim sPlateNo    As String
Dim sThk As String      '厚
Dim sWid As String      '宽
Dim sLen As String      '长

Dim sStlgrd As String   '钢种
Dim sStdspec As String  '标准
Dim sCustCD As String   '客户
Dim sStdspec_yy As String '打印标准
Dim sProddate As String
Dim sGroup As String
Dim sVessel_no As String
Dim sCustCD_short As String
Dim sWgt As String
Dim sUST As String
Dim sIDEMARK As String
Dim sEALMEMO As String

    With ss1
        For iCount = 1 To .MaxRows
             .Col = 0
             .ROW = iCount
    
                .Col = SPD_PLATE_NO:         sPlateNo = .Text
                .Col = SPD_THK:              sThk = Trim(str(.Text))
                .Col = SPD_WID:              sWid = Trim(str(.Text))
                .Col = SPD_LEN:              sLen = Trim(str(.Text))
                .Col = SPD_STLGRD_CP:        sStlgrd = .Text
                .Col = SPD_APLY_STDSPEC:     sStdspec = .Text
                .Col = SPD_CUST_CD:          sCustCD = .Text
                .Col = SPD_STDSPEC_YY:       sStdspec_yy = .Text
                .Col = SPD_DATE:             sProddate = .Text
                 sProddate = Mid(sProddate, 1, 10)
                .Col = SPD_GROUP:            sGroup = .Text
                .Col = SPD_VESSEL_NO:        sVessel_no = .Text
                .Col = SPD_SIDEMARK:         sIDEMARK = .Text
                .Col = SPD_SEALMEMO:         sEALMEMO = .Text
                .Col = SPD_CUST_CD_SHORT:    sCustCD_short = .Text
                 sCustCD_short = Trim(sCustCD_short)
                .Col = SPD_WGT:              sWgt = .Text
                 If Mid(sWgt, 1, 1) = "." Then
                            sWgt = "0" & sWgt
                 End If
                .Col = SPD_UST:              sUST = .Text
                .Col = 0
                If (chk_Cond(0) Or chk_Cond(1) Or chk_Cond(2)) Then
                    Call Cmd_SEND(sPlateNo, sThk, sWid, sLen, sStlgrd, sStdspec, sCustCD, sStdspec_yy, sProddate, sGroup, sVessel_no, sIDEMARK, sEALMEMO, sCustCD_short, sWgt, sUST)
                End If
        Next iCount
    End With
End Sub

Private Sub Timer1_Timer()

    'sckClosed            0 缺省的。--关闭 没有的
    'sckOpen              1 打开 --打开的
    'sckListening         2 侦听 --察看有没有请求进入的
    'sckConnectionPending 3 连接挂起
    'sckResolvingHost     4 识别主机
    'sckHostResolved      5 已识别主机
    'sckConnecting        6 正在连接
    'sckConnected         7 已连接
    'sckClosing           8 同级人员正在关闭连接 -说明对方关闭了你连接
    'sckError             9 错误
    
    Dim strState As String
    Dim strState2 As String
    Dim strState3 As String
    Dim strState4 As String
    
    If chk_Cond(0) <> 1 And chk_Cond(1) <> 1 And chk_Cond(2) <> 1 And chk_Cond(3) <> 1 Then
       Exit Sub
    Else
    
        If chk_Cond(0) = 1 Then
            
            Select Case Winsock1.State
                Case 0
                    strState = "连接关闭"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 1
                    strState = "连接打开"
                Case 2
                    strState = "连接保留"
                Case 3
                    strState = "Close"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 4
                    strState = "Find Host...."
                Case 5
                    strState = "找到主机"
                Case 6
                    strState = "正在连接"
                Case 7
                    strState = "连接正常"
                    tcpStatus.BackColor = &HC000&
                    chk_Cond(0).ForeColor = &HC000&
                Case 8
                    strState = "连接断线"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
                Case 9
                    strState = "连接错误"
                    tcpStatus.BackColor = &HFF&
                    chk_Cond(0).ForeColor = &HFF&
            Case Else
                strState = "StateNum:" & Winsock1.State
                tcpStatus.BackColor = &HFF&
                chk_Cond(0).ForeColor = &HFF&
            End Select

            tcpMsg.Caption = "侧喷机状态 : " & strState

        End If
        
        If chk_Cond(1) = 1 Then

            Select Case Winsock2.State
                Case 0
                    strState2 = "连接关闭"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(1).ForeColor = &HFF&
                Case 1
                    strState2 = "连接打开"
                Case 2
                    strState2 = "连接保留"
                Case 3
                    strState2 = "Close"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(1).ForeColor = &HFF&
                Case 4
                    strState2 = "Find Host...."
                Case 5
                    strState2 = "找到主机"
                Case 6
                    strState2 = "正在连接"
                Case 7
                    strState2 = "连接正常"
                    tcpStatus2.BackColor = &HC000&
                    chk_Cond(1).ForeColor = &HC000&
                Case 8
                    strState2 = "连接断线"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(1).ForeColor = &HFF&
                Case 9
                    strState2 = "连接错误"
                    tcpStatus2.BackColor = &HFF&
                    chk_Cond(1).ForeColor = &HFF&
            Case Else
                strState2 = "StateNum:" & Winsock2.State
                tcpStatus2.BackColor = &HFF&
                chk_Cond(1).ForeColor = &HFF&
            End Select

            tcpMsg2.Caption = "侧喷机状态 : " & strState2

        End If
        
        '2011-11-25  标印机
        If chk_Cond(2) = 1 Then

            Select Case Winsock3.State
                Case 0
                    strState3 = "连接关闭"
                    tcpStatus3.BackColor = &HFF&
                    chk_Cond(2).ForeColor = &HFF&
                Case 1
                    strState3 = "连接打开"
                Case 2
                    strState3 = "连接保留"
                Case 3
                    strState3 = "Close"
                    tcpStatus3.BackColor = &HFF&
                    chk_Cond(2).ForeColor = &HFF&
                Case 4
                    strState3 = "Find Host...."
                Case 5
                    strState3 = "找到主机"
                Case 6
                    strState3 = "正在连接"
                Case 7
                    strState3 = "连接正常"
                    tcpStatus3.BackColor = &HC000&
                    chk_Cond(2).ForeColor = &HC000&
                Case 8
                    strState3 = "连接断线"
                    tcpStatus3.BackColor = &HFF&
                    chk_Cond(2).ForeColor = &HFF&
                Case 9
                    strState3 = "连接错误"
                    tcpStatus3.BackColor = &HFF&
                    chk_Cond(2).ForeColor = &HFF&
            Case Else
                strState3 = "StateNum:" & Winsock3.State
                tcpStatus3.BackColor = &HFF&
                chk_Cond(2).ForeColor = &HFF&
            End Select

            tcpMsg3.Caption = "标印机状态 : " & strState3

        End If
        
         '2012-08-08  钢印机
                 If chk_Cond(3) = 1 Then

            Select Case Winsock4.State
                Case 0
                    strState4 = "连接关闭"
                    tcpStatus4.BackColor = &HFF&
                    chk_Cond(3).ForeColor = &HFF&
                Case 1
                    strState4 = "连接打开"
                Case 2
                    strState4 = "连接保留"
                Case 3
                    strState4 = "Close"
                    tcpStatus4.BackColor = &HFF&
                    chk_Cond(3).ForeColor = &HFF&
                Case 4
                    strState4 = "Find Host...."
                Case 5
                    strState4 = "找到主机"
                Case 6
                    strState4 = "正在连接"
                Case 7
                    strState4 = "连接正常"
                    tcpStatus4.BackColor = &HC000&
                    chk_Cond(3).ForeColor = &HC000&
                Case 8
                    strState4 = "连接断线"
                    tcpStatus4.BackColor = &HFF&
                    chk_Cond(3).ForeColor = &HFF&
                Case 9
                    strState4 = "连接错误"
                    tcpStatus4.BackColor = &HFF&
                    chk_Cond(3).ForeColor = &HFF&
            Case Else
                strState4 = "StateNum:" & Winsock4.State
                tcpStatus4.BackColor = &HFF&
                chk_Cond(3).ForeColor = &HFF&
            End Select

            tcpMsg4.Caption = "钢印机状态 : " & strState4

        End If
        
    End If
    
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
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))

    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    ss1.RowHeight(-1) = 13.5

    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "EG-System.INI", Me.Name)
    
    ''  2012-02-15  中板厂厚板标识查询,在此界面导出标识内容打印,把隐藏内容显示
'    '取钢种发送侧喷设备, 2011-07-14 , by Liqian ,画面隐藏该列
'    Call Gp_Sp_ColHidden(ss1, SPD_STLGRD_CP, True)
'    '取标准判断是否船板, 2011-07-14 , by Liqian ,画面隐藏该列
'    Call Gp_Sp_ColHidden(ss1, SPD_APLY_STDSPEC, True)
'    '取钢种发送侧喷设备, 2011-07-14 , by Liqian ,画面隐藏该列
'    Call Gp_Sp_ColHidden(ss1, SPD_CUST_CD, True)
'    '取打印标准发送侧喷设备, 2011-07-14 , by Liqian ,画面隐藏该列
'    Call Gp_Sp_ColHidden(ss1, SPD_STDSPEC_YY, True)
    
    txt_WkPlt = "C3"
    
    Screen.MousePointer = vbDefault
    
     '本地测试用...
'    Winsock1.RemoteHost = "172.18.57.76"
'    Winsock1.RemotePort = "9099"
''
'    Winsock2.RemoteHost = "172.18.57.76"
'    Winsock2.RemotePort = "9099"
'
'    '标印机
'    Winsock3.RemoteHost = "172.18.43.113"
'    Winsock3.RemotePort = "34242"
'
''    '钢印机
'    Winsock4.RemoteHost = "172.18.57.76"
'    Winsock4.RemotePort = "9099"
    
    '侧喷端口,需要提供...
    Winsock1.RemoteHost = Gf_ComnNameFind(M_CN1, "G0040", "03", 1)
    Winsock1.RemotePort = Gf_ComnNameFind(M_CN1, "G0040", "03", 2)
    Winsock2.RemoteHost = Gf_ComnNameFind(M_CN1, "G0040", "04", 1)
    Winsock2.RemotePort = Gf_ComnNameFind(M_CN1, "G0040", "04", 2)
    '表喷主机
    Winsock3.RemoteHost = Gf_ComnNameFind(M_CN1, "G0040", "05", 1)
    Winsock3.RemotePort = Gf_ComnNameFind(M_CN1, "G0040", "05", 2)

    '钢印主机
    Winsock4.RemoteHost = Gf_ComnNameFind(M_CN1, "EG001", "01", 1)
    Winsock4.RemotePort = Gf_ComnNameFind(M_CN1, "EG001", "01", 2)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "EG-System.INI", Me.Name)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If


    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing

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

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
        If Gf_Sp_Cls(sc2) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
            Call Gp_Ms_ControlLock(Mc2("lControl"), False)
            
'            lbl_moplate_wgt.Caption = ""
            
        End If
    End If
End Sub

Public Sub Form_Ref()
    
    On Error GoTo Refer_Err
    
    Dim iCount As Integer
    
    ss1.MaxRows = 0
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False)
    ss1.OperationMode = OperationModeNormal
    ss2.OperationMode = OperationModeNormal
    If ss2.MaxRows > 0 Then
       Call ss2_DblClick(1, 1)
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       ss2.Col = 0
       ss2.ROW = 0
       ss2.Text = "◎"
    End If
    
Refer_Err:
       
End Sub

Public Sub Form_Pro()

Dim iCount      As Integer
Dim START_FOR   As Integer
Dim sDateFrom   As String
Dim sDateTo     As String
Dim sPlateNo    As String

Dim sThk As String      '厚
Dim sWid As String      '宽
Dim sLen As String      '长

Dim sStlgrd As String   '钢种
Dim sStdspec As String  '标准
Dim sCustCD As String   '客户
Dim sStdspec_yy As String '打印标准
Dim sProddate As String
Dim sGroup As String
Dim sVessel_no As String
Dim sCustCD_short As String
Dim sWgt As String
Dim sUST As String
Dim sIDEMARK As String
Dim sEALMEMO As String
    
Dim inum As Integer
Dim lRow As Integer
    
    For iCount = 1 To ss1.MaxRows
        ss1.ROW = iCount
'        ss1.Col = 0
'        ss1.Text = "Update"
        ss1.Col = 18
        If ss1.Value = 1 Then
            START_FOR = iCount
            Exit For
        End If
    Next

    If START_FOR < ss1.MaxRows Then
        START_FOR = START_FOR + 1
    End If
    
    If Gf_Sp_Pro(M_CN1, Proc_Sc("SC"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
    End If
        
     With ss1
        For iCount = 1 To .MaxRows
             .Col = 0
             .ROW = iCount
             If .Text = "Update" Or .Text = "Insert" Then
             
                .Col = SPD_PLATE_NO:         sPlateNo = .Text
                .Col = SPD_THK:              sThk = Trim(str(.Text))
                .Col = SPD_WID:              sWid = Trim(str(.Text))
                .Col = SPD_LEN:              sLen = Trim(str(.Text))
                .Col = SPD_STLGRD_CP:        sStlgrd = .Text
                .Col = SPD_APLY_STDSPEC:     sStdspec = .Text
                .Col = SPD_CUST_CD:          sCustCD = .Text
                .Col = SPD_STDSPEC_YY:       sStdspec_yy = .Text
                .Col = SPD_DATE:             sProddate = .Text
                 sProddate = Mid(sProddate, 1, 10)
                .Col = SPD_GROUP:            sGroup = .Text
                .Col = SPD_VESSEL_NO:        sVessel_no = .Text
                .Col = SPD_SIDEMARK:         sIDEMARK = .Text
                .Col = SPD_SEALMEMO:         sEALMEMO = .Text
                .Col = SPD_CUST_CD_SHORT:    sCustCD_short = .Text
                 sCustCD_short = Trim(sCustCD_short)
                .Col = SPD_WGT:              sWgt = .Text
                 If Mid(sWgt, 1, 1) = "." Then
                            sWgt = "0" & sWgt
                 End If
                .Col = SPD_UST:              sUST = .Text
                .Col = 0
                If (chk_Cond(0) Or chk_Cond(1) Or chk_Cond(2)) Then
                    Call Cmd_SEND(sPlateNo, sThk, sWid, sLen, sStlgrd, sStdspec, sCustCD, sStdspec_yy, sProddate, sGroup, sVessel_no, sIDEMARK, sEALMEMO, sCustCD_short, sWgt, sUST)
                End If
'                Exit For
             End If
        Next iCount
    End With
    
    Call Form_Ref
    
End Sub

Private Sub Cmd_SEND(iSplate_no As String, iThk As String, iWid As String, iLen As String, iStlgrd As String, iStdspec As String, iCustCD As String, iStdspec_yy As String, iProddate As String, iGroup As String, iVessel_no As String, iSidemark As String, iSealmemo As String, iCustCD_short As String, iWGT As String, iUST As String)
    Dim sMesg As String

    Dim sPlate_no As String             '钢板号
    Dim sSize_Str As String             '规格
    Dim sStlgrd As String               '钢种
    Dim sStdspec As String              '标准
    Dim sCustCD As String               '客户
    Dim sStdspec_yy As String           '打印标准
    Dim sSend_Str As String * 55        '侧喷内容
    Dim sNum As String
    Dim sSpec_Str As String
    
    '标印接口
    Dim Header As String * 2           '钢印、标印共用
    Dim sPlateNo As String * 14        '钢印、标印共用
    Dim sLen As String * 10
    Dim sWid As String * 9
    Dim sThk As String * 9
    Dim sLogo1 As String * 1           '钢印、标印共用
    Dim sLogo2 As String * 1           '钢印、标印共用
    Dim sPaint_L1 As String * 46
    Dim sPaint_L2 As String * 46
    Dim sPaint_L3 As String * 46
    Dim sPaint_L4 As String * 46
    Dim sSend_Data As String
    
    '钢印接口
    Dim sSend_Punch As String
    Dim sStlgrd1 As String * 16        '钢印钢种
    Dim sNGPlateNo As String * 16      '钢印NG+子板号
    
    Dim Nisco As String
    Dim sFlag As String
    Dim sWgt As String
    Dim sUST As String
    Dim sProd_Date As String
    Dim sGroup As String
    Dim sSpec As String
    Dim sSpec1 As String
    Dim sCUST_CD_SHORT As String
    Dim sVessel_no As String
    Dim sSidemark As String
    Dim sSealmemo As String
    
    sPlate_no = iSplate_no
    sSize_Str = iThk + "X" + iWid + "X" + iLen
    
    '表喷
    Header = "MD"
    sPlateNo = iSplate_no
    sLen = iLen
    sWid = iWid
    sThk = iThk
    sWgt = iWGT
    sProd_Date = iProddate
    sCUST_CD_SHORT = iCustCD_short
    sVessel_no = iVessel_no
    sUST = iUST
    
    sGroup = Trim(iGroup)
    
    If sGroup <> "A" And sGroup <> "B" And sGroup <> "C" And sGroup <> "D" Then
        sMesg = " 班别错误，请确认是否正确输入班别"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    sStdspec = iStdspec
    sSpec = iStlgrd
    sStdspec_yy = iStdspec_yy
    sNum = InStr(sStdspec, "-")
    If sNum = 0 Then
        sNum = Len(sStdspec)
    End If
    sSpec_Str = Mid(sStdspec, 1, (sNum - 1))
    
    If sStdspec = "ZH-ABS-A36" Then
       sSpec_Str = "ABS"
    End If
    
    sSpec1 = sStdspec_yy

    Select Case sSpec_Str

           Case "ABS"                                 '美国船级社
                sStlgrd = sStdspec_yy
                sLogo1 = 9
                sPaint_L2 = sSpec1
           Case "CCS"                                 '中国
                sStlgrd = sStdspec_yy
                sLogo1 = 1
                sPaint_L2 = sSpec1
           Case "GL"                                  '德国
                sStlgrd = sStdspec_yy
                sLogo1 = 8
                sPaint_L2 = sSpec1
           Case "BV"                                  '法国
                sStlgrd = sStdspec_yy
                sLogo1 = 2
                sPaint_L2 = sSpec1
           Case "DNV"                                 '挪威
                sStlgrd = sStdspec_yy
                sLogo1 = 4
                sPaint_L2 = sSpec1
           Case "KR"                                  '韩国
                sStlgrd = sStdspec_yy
                sLogo1 = 7
                sPaint_L2 = sSpec1
           Case "LR"                                  '英国
                sStlgrd = sStdspec_yy
                sLogo1 = 6
                sPaint_L2 = sSpec1
           Case "RINA"                                '意大利
                sStlgrd = sStdspec_yy
                sLogo1 = 3
                sPaint_L2 = sSpec1
           Case "NK"                                  '日本
                sStlgrd = sStdspec_yy
                sLogo1 = 5
                sPaint_L2 = sSpec1
           Case "IRS"                                 '印度
                sStlgrd = sStdspec_yy
                sLogo1 = ""
                sPaint_L2 = sSpec1
           Case Else
                sStlgrd = iStlgrd
                sLogo1 = ""
                sSpec1 = sSpec + " " + sStdspec_yy
                sPaint_L2 = sSpec1
    End Select
    
    sCustCD = iCustCD
    
    sSidemark = iSidemark  '侧喷加喷 20150122
    sSealmemo = iSealmemo  '加冲钢印 20150122
    
    Nisco = "NG"
    sFlag = "X"
    
    '编辑探伤信息
    '如果钢板要求探伤，喷印第四行加喷 T
    If sUST = "" Or sUST = "/" Or sUST = "X" Then
       sUST = ""
    Else
       sUST = "T"
    End If
    
    '侧标信息
    'sSend_Str = sPlate_no + " " + sSize_Str + " " + sStlgrd + " " + sCustCD
     sSend_Str = sPlate_no + " " + sStlgrd + " " + sSize_Str + " " + sSidemark + " " + sCustCD
    
    '有重量标识要求的编辑重量信息
    If iStdspec_yy = "GB 713-2008" Or iStdspec_yy = "GB 3531-2008" Or iStdspec_yy = "GB 19189-2003" Then
        sWgt = "  T.W. " & sWgt & " t"
    Else
        sWgt = ""
    End If
    
    sLogo2 = ""
    'sPaint_L1 = Nisco + " " + sPlate_no + " " + sWgt
     sPaint_L1 = sPlate_no + " " + sWgt
    
    ' 编辑标印第3行内容
    sPaint_L3 = sSize_Str + " " + sProd_Date + " " + sGroup
    
    ' 编辑标印第4行内容
    If sCUST_CD_SHORT <> "" Then
       If sUST <> "" Then
       sPaint_L4 = sCUST_CD_SHORT + " " + sUST + " " + sVessel_no
       Else
       sPaint_L4 = sCUST_CD_SHORT + " " + sVessel_no
       End If
    Else
       If sUST <> "" Then
       sPaint_L4 = sUST + " " + sVessel_no
       Else
       sPaint_L4 = sVessel_no
       End If
    End If
    
    '表喷信息
    sSend_Data = Header + sPlateNo + sLen + sWid + sThk + sLogo1 + sLogo2 + sPaint_L1 + sPaint_L2 + sPaint_L3 + sPaint_L4
    
    '钢印信息
    sStlgrd1 = sStlgrd
    'sNGPlateNo = Nisco + sPlateNo
     sNGPlateNo = sPlateNo
     
    sSend_Punch = Header + sPlateNo + sLogo1 + sLogo2 + sStlgrd1 + sNGPlateNo + sSealmemo

    If chk_Cond(0) = 1 Or chk_Cond(1) = 1 Or chk_Cond(2) = 1 Then
    
       If chk_Cond(0) = 1 Then
          Winsock1.SendData sSend_Str
       End If
       If chk_Cond(1) = 1 Then
          Winsock2.SendData sSend_Str
       End If
       If chk_Cond(2) = 1 Then
          Winsock3.SendData sSend_Data
       End If
       
       If chk_Cond(3) = 1 Then
          Winsock4.SendData sSend_Punch
       End If
       
    End If

End Sub

Public Sub Form_Ins()
Dim dThk        As Double
Dim dWid        As Double
Dim dLen        As Double
Dim dWgt        As Double
Dim lRow        As Long
Dim sPlateNo    As String
Dim sClipText   As String
Dim cCUTDATE As String
Dim cTRIM As String
Dim cHEAD As String
Dim cTAIL As String
Dim cSAMP1 As String
Dim cSAMP2 As String

Dim iIdr        As Integer
Dim iCount As Integer

    sPlateNo = ""

    With ss1
        If .MaxRows = 0 Then
           If Len(TXT_MPLATE_NO.Text) = 12 Then
               Call Gp_Sp_Ins(Proc_Sc("Sc"))
              .ROW = 1
              .Col = 1
              .Text = TXT_MPLATE_NO.Text & "01"
              .Col = 27
              .Text = txt_WkPlt
              .Col = 28
              .Text = txt_PrcLine
              .Col = 29
              .Text = Trim(txt_loc.Text)
              .Col = 26
              .Text = sUserID
           Else
               Call Gp_MsgBoxDisplay("请正确输入母板号 ！")
           End If
           Exit Sub
        End If
        For iCount = .ActiveRow To .MaxRows
            .ROW = iCount
            .Col = 1
            If Left(.Text, 12) = Left(sPlateNo, 12) Or sPlateNo = "" Then
               sPlateNo = .Text
               lRow = iCount
            Else
               Exit For
            End If
        Next iCount
    End With

    sPlateNo = ""

    Call ss1.SetActiveCell(1, lRow)
    Call Gp_Sp_Ins(Proc_Sc("Sc"))

    With ss1
        .ReDraw = False
        .ROW = .ActiveRow
        .Col = 27
        .Text = txt_WkPlt
        .Col = 28
        .Text = txt_PrcLine
        .Col = 29
        .Text = Trim(txt_loc.Text)
        .Col = 26
        .Text = sUserID

        If lRow > 0 Then
            .ROW = lRow
            .Col = SPD_PLATE_NO:      sPlateNo = .Text
            .Col = SPD_THK:           dThk = Val(.Text & "")
            .Col = SPD_WID:           dWid = Val(.Text & "")
            .Col = SPD_LEN:           dLen = Val(.Text & "")
            .Col = SPD_WGT:           dWgt = Val(.Text & "")
        Else
            sPlateNo = TXT_MPLATE_NO.Text & "00"
        End If

        .ROW = lRow + 1
        .Col = SPD_PLATE_NO:      .Text = sPlateNo
        .Col = SPD_THK:           .Text = dThk
        .Col = SPD_WID:           .Text = dWid
        .Col = SPD_LEN:           .Text = dLen
        .Col = SPD_WGT:           .Text = dWgt
        .Col = 0: .Text = "Input"
        .Col = SPD_PLATE_NO: .Text = Left(.Text, 12) + Format((Val(Mid(.Text, 13, 2)) + 1), "00")

         Call .SetActiveCell(1, .ROW)
        .ReDraw = True
    End With
    
    ss1.ROW = 1
    ss1.Col = 13
    cCUTDATE = ss1.Text
    ss1.Col = 19
    cTRIM = ss1.Value
    ss1.Col = 20
    cHEAD = ss1.Value
    ss1.Col = 21
    cTAIL = ss1.Value
    ss1.Col = 22
    cSAMP1 = ss1.Value
    ss1.Col = 23
    cSAMP2 = ss1.Value

    For iIdr = 1 To ss1.MaxRows
        ss1.ROW = iIdr
        ss1.Col = 0
        If ss1.Text = "Input" Then
            ss1.Col = 13
            ss1.Text = cCUTDATE
            ss1.Col = 18
            ss1.Value = 0
            
            ss1.Col = 19
            ss1.Value = cTRIM
            ss1.Col = 20
            ss1.Value = cHEAD
            ss1.Col = 21
            ss1.Value = cTAIL
            ss1.Col = 22
            ss1.Value = cSAMP1
            ss1.Col = 23
            ss1.Value = cSAMP2
            
            ss1.Col = 24
            ss1.Text = Gf_ShiftSet3(M_CN1)
            ss1.Col = 25
            ss1.Text = Gf_GroupSet(M_CN1, Gf_ShiftSet3(M_CN1), Gf_DTSet(M_CN1, , "X"))
            ss1.Col = 26
            ss1.Text = sUserID
            ss1.Col = 27
            ss1.Text = txt_WkPlt
            ss1.Col = 28
            ss1.Text = txt_PrcLine
            ss1.Col = 29
            ss1.Text = Trim(txt_loc.Text)
            If iIdr = ss1.MaxRows Then
               ss1.ROW = iIdr
               ss1.Col = 18
               ss1.Value = 1
            End If
        End If
    Next iIdr


End Sub

Private Function Cal_Plate_Wgt(sMode As String, sEndUseCd As String, sStlgrd As String, _
                                dThk As Double, dWid As Double, dLen As Double) As Double

    Dim RS  As New ADODB.Recordset
    
    Cal_Plate_Wgt = 0
    
    sQuery = "SELECT  Gf_Cal_Plate_Wgt('" & sMode & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & sEndUseCd & "'" & vbCrLf
    sQuery = sQuery & "             ,'" & sStlgrd & "'" & vbCrLf
    sQuery = sQuery & "             ," & dThk & vbCrLf
    sQuery = sQuery & "             ," & dWid & vbCrLf
    sQuery = sQuery & "             ," & dLen & vbCrLf
    sQuery = sQuery & "             ,0 )" & vbCrLf
    sQuery = sQuery & "       FROM  DUAL " & vbCrLf
    RS.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If RS.EOF = False Then
        Cal_Plate_Wgt = Val(RS(0).Value & "")
    End If
    
    RS.Close
    Set RS = Nothing
     
End Function

Public Sub Spread_Can()
    ss1.Col = 0
    ss1.ROW = ss1.ActiveRow
    Select Case Trim(ss1.Text)
        Case "Input"
              Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
        Case Else
             ss1.Text = ""
    End Select
End Sub
Public Sub Spread_Del()

    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Public Sub Spread_Cpy()
    Call Gp_Sp_Copy(Proc_Sc("Sc"))
End Sub

Public Sub Spread_Pst()
    Call Gp_Sp_Paste(Proc_Sc("Sc"))
End Sub


Public Sub Spread_Forzens_Setting()
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
End Sub

Public Sub Spread_Forzens_Cancel()
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
End Sub




Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub


Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal ROW As Long, ByVal ButtonDown As Integer)
Dim FOR_CNT
Dim START_FOR As Integer
    If Col <> 18 Then Exit Sub
    If ButtonDown = 0 Then Exit Sub
    For FOR_CNT = 1 To ss1.MaxRows
        If FOR_CNT <> ROW Then
            ss1.Col = 18
            ss1.ROW = FOR_CNT
            ss1.Value = 0
        End If
    Next
       
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    ss1.Col = 0
    ss1.ROW = 0
    ss1.Text = "◎"
    
    ss2.Col = 0
    ss2.ROW = 0
    ss2.Text = ""
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
Dim sDate     As String
Dim sDateTo   As String
Dim FOR_CNT   As Long
Dim tmpYYMMDD As String

    If ROW < 1 Then Exit Sub
    If Col < 11 Then Exit Sub
    
    tmpYYMMDD = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL")
    
    ss1.ROW = ROW: ss1.Col = Col
    
    'For FOR_CNT = 1 To ss1.MaxRows
    
        ss1.ROW = ROW
        If ss1.Col = 13 Then
           ss1.Text = tmpYYMMDD
        End If
        ss1.Col = 29
        ss1.Text = Trim(txt_loc.Text)
        ss1.Col = 28
        ss1.Text = txt_PrcLine.Text
        ss1.Col = 27
        ss1.Text = txt_WkPlt
        ss1.Col = 26
        ss1.Text = sUserID
        Call ss1_Row_Edit(ROW)
    'Next
    
End Sub

Private Sub ss1_LostFocus()
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal ROW As Long)
    ss1.Col = 0
    ss1.ROW = 0
    ss1.Text = ""
    
    ss2.Col = 0
    ss2.ROW = 0
    ss2.Text = "◎"
End Sub

Private Sub ss2_LostFocus()
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

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    If ROW > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If
End Sub

Private Sub ss1_EditChange(ByVal Col As Long, ByVal ROW As Long)
    
    Dim dThk        As Double
    Dim dWid        As Double
    Dim dLen        As Double
    Dim sEndUseCd   As String
    Dim sStlgrd     As String
    
    If ROW < 1 Then Exit Sub
    
    ss1.ROW = ROW
            
    If Col = SPD_THK Or Col = SPD_WID Or Col = SPD_LEN Then
        ss1.Col = SPD_THK:  dThk = Val(ss1.Text & "")
        ss1.Col = SPD_WID:  dWid = Val(ss1.Text & "")
        ss1.Col = SPD_LEN:  dLen = Val(ss1.Text & "")
        ss1.Col = SPD_END_USE:   sEndUseCd = Trim(ss1.Text)
        ss1.Col = SPD_STLGRD_CP:    sStlgrd = Trim(ss1.Text)
        If dThk > 0 And dWid > 0 And dLen > 0 Then
            ss1.Col = SPD_WGT
            ss1.Text = Cal_Plate_Wgt("WGT", sEndUseCd, sStlgrd, dThk, dWid, dLen)
        End If
    End If
    
    Call ss1_Row_Edit(ROW)
End Sub

Private Sub ss1_Change(ByVal Col As Long, ByVal ROW As Long)
    If ROW < 1 Then Exit Sub
       
    Call ss1_Row_Edit(ROW)

End Sub
Private Sub ss1_Data_Edit()
    Dim iIdr        As Integer
    Dim iThk        As Long
    Dim iWid        As Long
    Dim iLen        As Long
    Dim iWGT        As Double
    Dim ROW         As Long
    Dim sDate       As String
    Dim sDateTo     As String
    
    For iIdr = 1 To ss1.MaxRows
        ss1.ROW = iIdr
        ss1.Col = 24
        ss1.Text = Gf_ShiftSet3(M_CN1)
        ss1.Col = 25
        ss1.Text = Gf_GroupSet(M_CN1, Gf_ShiftSet3(M_CN1), Gf_DTSet(M_CN1, , "X"))
        ss1.Col = 26
        ss1.Text = sUserID
        ss1.Col = 27
        ss1.Text = txt_WkPlt
        ss1.Col = 28
        ss1.Text = txt_PrcLine.Text
        
    Next iIdr
    
    
End Sub
'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Pro
'   2.Name         : Spread Data Process
'   3.Input  Value : Conn Connection, Sc Collection, Mc Collection, {RefChek Boolean}
'   4.Return Value : Boolean
'   5.Writer       : 杨猛
'   6.Create Date  : 2010. 12 .09
'   7.Modify Date  :
'   8.Comment      : Spread Data Process
'---------------------------------------------------------------------------------------
Public Function Gf_Sp_Pro(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                              Optional RefChek As Boolean = False) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim dTempFloat As Double
    
    Dim sMesg As String
    Dim sTemp As String
    Dim ProcessChk As String
    Dim DelYN As Boolean
    Dim Msg_Count As Integer
    Dim Msg_Yes As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command

    Gf_Sp_Pro = True
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function Or iCount = 0
    If Sc.Item("Spread").MaxRows < 1 Or Sc.Item("iColumn").Count = 0 Then
        Gf_Sp_Pro = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    Sc.Item("Spread").ReDraw = False
    
    'NeceCheck
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
            
            Case "Input", "Update"
            
                If Not MC Is Nothing Then
                    Call Gp_Sp_Move(iCount, Sc, MC)
                End If
                
                'Maxlength Check
                sMesg = Gf_Sp_NeceCheck2(Sc.Item("Spread"), Sc.Item("mColumn"), iCount, Sc.Item("nColumn"))
                        
                If Trim(sMesg) = "OK" Then
                    
                ElseIf Mid(sMesg, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = Mid(sMesg, 6, Len(sMesg))
                    sMesg = sMesg + "长度不正确"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Pro = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = sMesg + "必须输入"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Pro = False
                    Exit Function
                End If
        
        End Select
    
    Next iCount
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Gf_Sp_Pro = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To Sc.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    Msg_Count = 1
    For iCount = 1 To Sc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        DelYN = False
        
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input"
                adoCmd.Parameters(0).Value = "I"
                ProcessChk = "YES"
                
            Case "Update"
                adoCmd.Parameters(0).Value = "U"
                ProcessChk = "YES"
                
            Case "Delete"
                adoCmd.Parameters(0).Value = "D"
                If Msg_Count = 1 Then
                   DelYN = Gf_MessConfirm("您确定要删除状态为[Delete]的数据吗？", "Q")
                   If DelYN Then Msg_Yes = "yes"
                   Msg_Count = Msg_Count + 1
                End If
                If Msg_Yes = "yes" Then DelYN = True
        End Select
          
        If ProcessChk = "YES" Or DelYN Then
            
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").Count
            
                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                
                Select Case Sc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempFloat = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Trim(str(dTempFloat))
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempInt = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Trim(str(dTempInt))
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Sc.Item("Spread").Value = "1" Then
                            adoCmd.Parameters(iCol).Value = "1"
                        Else
                            adoCmd.Parameters(iCol).Value = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = "0"
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc.Item("Spread").Value) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Sc.Item("Spread").Value)
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
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
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Gf_Sp_Pro = False
                Exit Function
        
             End If
        
        End If
        
    Next iCount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 0, iCount))
        
            Case "Input", "Update"
            
                sQuery = Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-O"), "O", Sc.Item("pColumn"), iCount)
                Call Gp_Sp_OneRowDisplay(Conn, sQuery, Sc.Item("Spread"), iCount)
                
            Case "Delete"
                If DelYN Then
                   Call Gp_Sp_SendData(Sc.Item("Spread"), "", 0, iCount)
                   Call Gp_Sp_DeleteRow(Sc.Item("Spread"), iCount)
                   iCount = iCount - 1
                End If
        End Select
        
    Next iCount
    
    Sc.Item("Spread").ReDraw = True
            
    If iProcessCount > 0 Then
        If Not MC Is Nothing Then
            Call Gp_Ms_ControlLock(MC.Item("lControl"), True)
        End If
    Else
        Gf_Sp_Pro = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:
    
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Gf_Sp_Pro = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Pro Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
'        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Sub ss1_Row_Edit(ByVal ROW As Long)
    
    ss1.Col = 0
    ss1.ROW = ROW
    If Trim(ss1.Text) <> "Input" And _
       Trim(ss1.Text) <> "Update" And _
       Trim(ss1.Text) <> "Delete" Then
       ss1.Text = "Update"
    End If
    
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal ROW As Long)
    
    Dim iCount As Integer

    If ROW < 1 Then Exit Sub
    
    ss2.ROW = ROW
    ss2.Col = 1
    If ss2.Text <> "" Then
       TXT_MPLATE_NO.Text = ss2.Text
    End If
    ss2.Col = 5
    TXT_PRODCD.Text = ss2.Text
    
    If Trim(TXT_MPLATE_NO.Text) <> "" Then
        Call Gf_Sp_Refer(M_CN1, sc1, Mc2, Mc2("nControl"), Mc2("mControl"), False)
        ss1.OperationMode = OperationModeNormal
        Call ss1_Data_Edit
        
        ss1.Col = 20
        ss1.ROW = ss1.MaxRows
        ss1.Value = 1
        ss1.Col = 13
        If ss1.Text = "" Then
           ss1.Col = 18
           ss1.Value = 1
        End If
        
    End If
    
    ss2.ROW = ROW
    ss2.Col = 5
    If ss2.Text = "MP" Then
        MDIMain.MenuTool.Buttons(7).Enabled = True
        MDIMain.MenuTool.Buttons(8).Enabled = True
        MDIMain.MenuTool.Buttons(9).Enabled = True
        MDIMain.MenuTool.Buttons(11).Enabled = True
        MDIMain.MenuTool.Buttons(12).Enabled = True
    ElseIf ss2.Text = "PP" Then
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(9).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    End If

End Sub

Private Sub TXT_INQNO_KeyUp(KeyCode As Integer, Shift As Integer)
     TXT_MPLATE_NO.Text = Trim(TXT_INQNO.Text)
End Sub

'Private Sub txt_plt_DblClick()
'    Call txt_plt_KeyUp(vbKeyF4, 0)
'End Sub
'
'Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "C0001"
'        DD.rControl.Add Item:=txt_plt
'        DD.rControl.Add Item:=txt_plt_name
'
'        DD.nameType = "2"
'        Call Gf_Common_DD(M_CN1, KeyCode)
'        Exit Sub
'
'    End If
'
'    If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
'        txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
'    Else
'        txt_plt_name.Text = ""
'    End If
'End Sub


'
'Private Sub txt_WkPlt_Change()
'    cbo_PrcLine.Clear
'
'    If txt_WkPlt = "C1" Then
'       cbo_PrcLine.AddItem "一号线"
'       cbo_PrcLine.AddItem "二号线"
'    Else
'       cbo_PrcLine.AddItem "一号线"
'    End If
'    cbo_PrcLine.ListIndex = 0
'End Sub


