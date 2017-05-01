VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQF0100C 
   Caption         =   "板坯试验实绩录入_AQF0100C"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   11085
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   3360
      Left            =   150
      TabIndex        =   4
      Top             =   5790
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   5927
      _Version        =   196609
      ForeColor       =   255
      Caption         =   "-"
      Begin VB.TextBox txt_FACT_CD6 
         Height          =   315
         Left            =   6540
         MaxLength       =   1
         TabIndex        =   24
         Top             =   2850
         Width           =   420
      End
      Begin VB.TextBox txt_FACT_Name6 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6975
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2850
         Width           =   7140
      End
      Begin VB.TextBox txt_FACT_CD5 
         Height          =   315
         Left            =   6540
         MaxLength       =   1
         TabIndex        =   22
         Top             =   2490
         Width           =   420
      End
      Begin VB.TextBox txt_FACT_Name5 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6975
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   21
         Top             =   2490
         Width           =   7140
      End
      Begin VB.TextBox txt_FACT_CD4 
         Height          =   315
         Left            =   6540
         MaxLength       =   1
         TabIndex        =   20
         Top             =   2130
         Width           =   420
      End
      Begin VB.TextBox txt_FACT_Name4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6975
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2130
         Width           =   7140
      End
      Begin VB.TextBox txt_FACT_CD3 
         Height          =   315
         Left            =   6540
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1770
         Width           =   420
      End
      Begin VB.TextBox txt_FACT_Name3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6975
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1770
         Width           =   7140
      End
      Begin VB.TextBox txt_FACT_CD2 
         Height          =   315
         Left            =   6555
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1410
         Width           =   420
      End
      Begin VB.TextBox txt_FACT_Name2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6990
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   15
         Top             =   1410
         Width           =   7140
      End
      Begin VB.TextBox txt_FACT_CD1 
         Height          =   315
         Left            =   6555
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1050
         Width           =   420
      End
      Begin VB.TextBox txt_FACT_Name1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6990
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1050
         Width           =   7140
      End
      Begin VB.TextBox txt_loc_name 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1635
         MaxLength       =   10
         TabIndex        =   12
         Top             =   780
         Width           =   990
      End
      Begin VB.TextBox txt_smp_loc 
         Height          =   315
         Left            =   1305
         MaxLength       =   1
         TabIndex        =   11
         Top             =   780
         Width           =   330
      End
      Begin VB.TextBox txt_macro 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1875
         Width           =   1335
      End
      Begin VB.TextBox txt_s_print 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   390
         MaxLength       =   10
         TabIndex        =   9
         Top             =   1875
         Width           =   1335
      End
      Begin VB.TextBox txt_Slab_NO 
         Height          =   315
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   8
         Top             =   345
         Width           =   990
      End
      Begin VB.TextBox txt_STLGRD 
         Enabled         =   0   'False
         Height          =   315
         Left            =   375
         MaxLength       =   11
         TabIndex        =   7
         Top             =   2625
         Width           =   1155
      End
      Begin VB.TextBox txt_STLGRD_Detail 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1545
         TabIndex        =   6
         Top             =   2625
         Width           =   1530
      End
      Begin VB.TextBox txt_DCS_CD 
         Height          =   285
         Left            =   3090
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1905
         Visible         =   0   'False
         Width           =   855
      End
      Begin InDate.ULabel ULabel14_shiyandengji 
         Height          =   315
         Left            =   390
         Tag             =   "订单号"
         Top             =   1215
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "试验等级"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel15 
         Height          =   315
         Left            =   390
         Top             =   1545
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "硫印"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel17 
         Height          =   315
         Left            =   1740
         Top             =   1545
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Caption         =   "低倍"
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel14 
         Height          =   315
         Left            =   390
         Tag             =   "订单号"
         Top             =   2295
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "实绩钢种"
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
         Left            =   3105
         Top             =   1545
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         Caption         =   "判定代码"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   390
         Top             =   345
         Width           =   915
         _ExtentX        =   1614
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
         ForeColor       =   -2147483641
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   390
         Top             =   780
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         Caption         =   "取样位置"
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
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   5415
         Top             =   1050
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "缺陷类型 1"
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
      Begin InDate.ULabel ULabel18 
         Height          =   315
         Left            =   5415
         Top             =   345
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   556
         Caption         =   "缺陷种类"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   5415
         Top             =   690
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "缺陷类型"
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
         Left            =   6555
         Top             =   690
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   556
         Caption         =   "代码"
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
      Begin InDate.ULabel ULabel7 
         Height          =   315
         Left            =   7005
         Top             =   690
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   556
         Caption         =   "缺陷名称"
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
         ForeColor       =   0
      End
      Begin InDate.ULabel ULabel8 
         Height          =   315
         Left            =   5415
         Top             =   1410
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "缺陷类型 2"
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
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   5400
         Top             =   1770
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "缺陷类型 3"
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
         Left            =   5400
         Top             =   2130
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "缺陷类型 4"
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
         Left            =   5400
         Top             =   2490
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "缺陷类型 5"
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
      Begin InDate.ULabel ULabel12 
         Height          =   315
         Left            =   5400
         Top             =   2850
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         Caption         =   "缺陷类型 6"
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
   Begin InDate.UDate txt_Date 
      Height          =   330
      Left            =   4185
      TabIndex        =   3
      Top             =   150
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   5070
      Left            =   120
      TabIndex        =   2
      Top             =   585
      Width           =   14955
      _Version        =   393216
      _ExtentX        =   26379
      _ExtentY        =   8943
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
      MaxCols         =   25
      MaxRows         =   13
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQF0100C.frx":0000
   End
   Begin VB.TextBox txt_Cast_NO 
      Height          =   315
      Left            =   7560
      MaxLength       =   18
      TabIndex        =   1
      Tag             =   "标准代号"
      Top             =   150
      Width           =   1335
   End
   Begin VB.TextBox txt_Charge_NO 
      Height          =   315
      Left            =   1470
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "标准代号"
      Top             =   150
      Width           =   840
   End
   Begin InDate.ULabel ULabel3_charge_no 
      Height          =   315
      Left            =   120
      Top             =   150
      Width           =   1335
      _ExtentX        =   2355
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
   Begin InDate.ULabel ULabel3_prod_date 
      Height          =   315
      Left            =   2760
      Top             =   150
      Width           =   1335
      _ExtentX        =   2355
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
   Begin InDate.ULabel ULabel3_cast_no 
      Height          =   315
      Left            =   6210
      Top             =   150
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "连浇炉数号"
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
   Begin VB.Line Line1 
      X1              =   120
      X2              =   15100
      Y1              =   520
      Y2              =   520
   End
End
Attribute VB_Name = "AQF0100C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   板坯取样
'-- Program Name      板坯取样实绩查询录入
'-- Program ID        AQF0100C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2006.01.11
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

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

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

Dim lngCurRow As Long
Dim bClicked As Boolean
Private Sub Form_Define()
      
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_charge_no, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(txt_Date, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(txt_Cast_NO, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)

    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", "n", "m", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", "n", "m", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AQF0100C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:="AQF0100C.P_SREFER1", Key:="P-R"
    sc1.Add Item:="AQF0100C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
      
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    Call Gp_Sp_ColHidden(ss1, 20, True)
    Call Gp_Sp_ColHidden(ss1, 23, True)
    MDIMain.MenuTool.Buttons(4).Enabled = True                  'Save
    MDIMain.MenuTool.Buttons(7).Enabled = True                  'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = True                  'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = True                  'Row Cancel
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
    
    
    Call Gp_Ms_Cls(Mc1("pControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    'Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gf_Sp_Cls(sc1)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    txt_Date.RawData = ""
    Call CtrlLock
    Screen.MousePointer = vbDefault
    'Call Combo_Set
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
          
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

Public Sub Spread_Can()
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc1"))
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC1")) Then
        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call pContro1(1).SetFocus
        txt_Date.RawData = ""
        Call CtrlCls
        Call CtrlLock
    End If
    ss1.Row = 0
    ss1.Col = 10: ss1.Text = "取样位置"
    ss1.Col = 11: ss1.Text = "判定代码"
    ss1.Col = 12: ss1.Text = "硫印等级"
    ss1.Col = 13: ss1.Text = "低倍等级"
End Sub

Public Sub Form_Ref()

    
    If Gf_Sp_ProceExist(ss1) Then Exit Sub
    'Call Gf_Sp_Cls(sc1)
    Call CtrlCls
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        If ss1.MaxRows < 1 Then
             Exit Sub
        End If
        bClicked = True
        Call CtrlLock(False)
        lngCurRow = 1
        
        Call Sp_To_Ctrl(lngCurRow)
        bClicked = True
        Call ss1.SetActiveCell(1, lngCurRow)
        Call ss1_Click(1, lngCurRow)
    End If
    'Call Gf_Sp_Cls(sc2)
    MDIMain.MenuTool.Buttons(4).Enabled = True                  'Save
    MDIMain.MenuTool.Buttons(7).Enabled = True                  'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = True                  'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = True                  'Row Cancel
   
End Sub

Public Sub Form_Ins()
    Call Gp_Sp_Ins(Proc_Sc("Sc1"))
    lngCurRow = ss1.ActiveRow
    If lngCurRow > 1 Then
    End If
    Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), 20)
    Call CtrlCls
End Sub
Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("Sc1"))

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub


Public Sub Form_Pro()
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc1"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        MDIMain.MenuTool.Buttons(4).Enabled = True                  'Save
        MDIMain.MenuTool.Buttons(7).Enabled = True                  'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = True                  'Row Delete
        MDIMain.MenuTool.Buttons(9).Enabled = True                  'Row Cancel
        Call Form_Ref
    End If
End Sub
            

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim strTag As String
    If ss1.MaxRows < 1 Or Row < 1 Then
        Exit Sub
    End If
    bClicked = True
    ss1.Col = 0: ss1.Row = Row: strTag = ss1.Text
    ss1.SetFocus
    lngCurRow = ss1.ActiveRow
    Call CtrlCls
    Call Sp_To_Ctrl(lngCurRow)
    Call CtrlLock(False)
    ss1.Col = 0: ss1.Row = Row: ss1.Text = strTag
    bClicked = False
End Sub


Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Debug.Print KeyCode

    Select Case KeyCode
    Case 33, 34, 38, 40
        bClicked = True
        Call CtrlCls
        ss1.SetFocus
        lngCurRow = ss1.ActiveRow
        Call Sp_To_Ctrl(lngCurRow)
        Call CtrlLock(False)
        bClicked = False
    End Select
End Sub


Private Sub txt_smp_loc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        txt_smp_loc.Text = ""
        DD.sWitch = "MS"
        DD.sKey = "Q0021"
        DD.rControl.Add Item:=txt_smp_loc
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        ss1.Col = 10: ss1.Row = lngCurRow: ss1.Text = UCase(txt_smp_loc.Text)
        If Len(Trim(txt_smp_loc.Text)) = 1 Then
            txt_loc_name.Text = Gf_ComnNameFind(M_CN1, "Q0021", Trim(txt_smp_loc.Text), 1)
            
        Else
            txt_loc_name.Text = ""
        End If
        Call Sp_Change(bClicked)
    End If
End Sub
Private Sub txt_SMP_LOC_Change()

   txt_smp_loc.Text = UCase(txt_smp_loc.Text)
   If bClicked = False Then
        ss1.Row = lngCurRow: ss1.Col = 10
   End If
   Select Case Trim(txt_smp_loc.Text)
          Case "M"
                txt_loc_name.Text = "中部"
                ss1.Text = "M"
          Case "T"
                txt_loc_name.Text = "头部"
                ss1.Text = "T"
          Case "B"
                txt_loc_name.Text = "尾部"
                ss1.Text = "B"
          Case Else
                txt_loc_name.Text = ""
                txt_smp_loc.Text = ""
                ss1.Text = ""
   End Select
   
   Call Sp_Change(bClicked)
End Sub
Private Sub txt_STLGRD_Change()
Dim sQuery As String
    If Len(Trim(txt_STLGRD.Text)) = 11 Then
       sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(txt_STLGRD.Text) + "'"
       
       txt_STLGRD_DETAIL = Gf_FloatFind(M_CN1, sQuery)
    Else
        txt_STLGRD_DETAIL = ""
    End If
End Sub

Private Sub txt_STLGRD_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim sQuery As String
    If KeyCode = vbKeyF4 Then
    
        DD.nameType = "1"
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_STLGRD
        
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
        If Len(Trim(txt_STLGRD.Text)) = 11 Then
           sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(txt_STLGRD.Text) + "'"
           txt_STLGRD_DETAIL = Gf_FloatFind(M_CN1, sQuery)
        Else
            txt_STLGRD_DETAIL = ""
        End If
    End If
End Sub

Private Sub Defact_CD_Set()
    ss1.Row = lngCurRow
    ss1.Col = 14: txt_FACT_CD1.Text = Trim(ss1.Text)
    ss1.Col = 15: txt_FACT_CD2.Text = Trim(ss1.Text)
    ss1.Col = 16: txt_FACT_CD3.Text = Trim(ss1.Text)
    ss1.Col = 17: txt_FACT_CD4.Text = Trim(ss1.Text)
    ss1.Col = 18: txt_FACT_CD5.Text = Trim(ss1.Text)
    ss1.Col = 19: txt_FACT_CD6.Text = Trim(ss1.Text)
End Sub

Private Sub txt_DCS_CD_Change()
    If bClicked = True Then Exit Sub
    
    ss1.Row = lngCurRow: ss1.Col = 11
    ss1.Text = txt_DCS_CD.Text
    Call Sp_Change(bClicked)
End Sub

Private Sub txt_FACT_CD1_Change()
    Dim sQuery As String
    
    If Len(Trim(txt_FACT_CD1.Text)) > 0 Then
       sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD1.Text) + "' AND CD_MANA_NO = 'Q0063'"
       txt_FACT_Name1.Text = Gf_FloatFind(M_CN1, sQuery)
       If bClicked = False Then
            If Len(Trim(txt_FACT_Name1.Text)) > 1 Then
                 ss1.Row = lngCurRow: ss1.Col = 14: ss1.Text = txt_FACT_CD1.Text
            Else
                 ss1.Row = lngCurRow: ss1.Col = 14: ss1.Text = ""
            End If
        End If
    Else
        txt_FACT_Name1.Text = ""
    End If
    Call Sp_Change(bClicked)
End Sub

Private Sub txt_FACT_CD1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim str_orgin As String
    Dim sQuery As String
    
    If KeyCode = vbKeyF4 Then
       str_orgin = txt_FACT_CD1.Text
       DD.sWitch = "MS"
       DD.sKey = "Q0063"
       DD.nameType = "2"
       txt_FACT_CD1.Text = ""
       DD.rControl.Add Item:=txt_FACT_CD1
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        If Len(Trim(txt_FACT_CD1.Text)) > 0 Then
            sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD1.Text) + "' AND CD_MANA_NO = 'Q0063'"
            txt_FACT_Name1.Text = Gf_FloatFind(M_CN1, sQuery)
            If bClicked = False Then
               ss1.Row = lngCurRow: ss1.Col = 14: ss1.Text = txt_FACT_CD1.Text
               Call Sp_Change(bClicked)
            End If
        Else
            txt_FACT_CD1.Text = str_orgin
        End If
        'Call ss2_Change(ss2.Col, ss2.Row)
    End If
End Sub

Private Sub txt_FACT_CD2_Change()
    Dim sQuery As String

    If Len(Trim(txt_FACT_CD2.Text)) > 0 Then
       sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD2.Text) + "' AND CD_MANA_NO = 'Q0063'"
       txt_FACT_Name2.Text = Gf_FloatFind(M_CN1, sQuery)
       If bClicked = False Then
            If Len(Trim(txt_FACT_Name2.Text)) > 1 Then
                 ss1.Row = lngCurRow: ss1.Col = 15: ss1.Text = txt_FACT_CD2.Text
            Else
                 ss1.Row = lngCurRow: ss1.Col = 15: ss1.Text = ""
            End If
        End If
    Else
        txt_FACT_Name2.Text = ""
    End If
    Call Sp_Change(bClicked)
End Sub

Private Sub txt_FACT_CD2_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim str_orgin As String
    Dim sQuery As String
    
    If KeyCode = vbKeyF4 Then
       str_orgin = txt_FACT_CD2.Text
       DD.sWitch = "MS"
       DD.sKey = "Q0063"
       DD.nameType = "2"
       txt_FACT_CD2.Text = ""
       DD.rControl.Add Item:=txt_FACT_CD2
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        If Len(Trim(txt_FACT_CD2.Text)) > 0 Then
            sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD2.Text) + "' AND CD_MANA_NO = 'Q0063'"
            txt_FACT_Name2.Text = Gf_FloatFind(M_CN1, sQuery)
            ss1.Row = lngCurRow: ss1.Col = 15: ss1.Text = txt_FACT_CD2.Text
            Call Sp_Change(bClicked)
        Else
            txt_FACT_CD2.Text = str_orgin
        End If
        'Call ss2_Change(ss2.Col, ss2.Row)
    End If

End Sub

Private Sub txt_FACT_CD3_Change()
    Dim sQuery As String

    If Len(Trim(txt_FACT_CD3.Text)) > 0 Then
       sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD3.Text) + "' AND CD_MANA_NO = 'Q0063'"
       txt_FACT_Name3.Text = Gf_FloatFind(M_CN1, sQuery)
       If bClicked = False Then
            If Len(Trim(txt_FACT_Name3.Text)) > 1 Then
                 ss1.Row = lngCurRow: ss1.Col = 16: ss1.Text = txt_FACT_CD3.Text
            Else
                 ss1.Row = lngCurRow: ss1.Col = 16: ss1.Text = ""
            End If
       End If
    Else
        txt_FACT_Name3.Text = ""
    End If
    Call Sp_Change(bClicked)
End Sub

Private Sub txt_FACT_CD3_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim str_orgin As String
    Dim sQuery As String
    
    If KeyCode = vbKeyF4 Then
       str_orgin = txt_FACT_CD3.Text
       DD.sWitch = "MS"
       DD.sKey = "Q0063"
       DD.nameType = "2"
       txt_FACT_CD3.Text = ""
       DD.rControl.Add Item:=txt_FACT_CD3
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        If Len(Trim(txt_FACT_CD3.Text)) > 0 Then
            sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD3.Text) + "' AND CD_MANA_NO = 'Q0063'"
            txt_FACT_Name3.Text = Gf_FloatFind(M_CN1, sQuery)
            ss1.Row = lngCurRow: ss1.Col = 16: ss1.Text = txt_FACT_CD3.Text
            Call Sp_Change(bClicked)
        Else
            txt_FACT_CD3.Text = str_orgin
        End If
        'Call ss2_Change(ss2.Col, ss2.Row)
    End If
End Sub

Private Sub txt_FACT_CD4_Change()
    Dim sQuery As String

    If Len(Trim(txt_FACT_CD4.Text)) > 0 Then
       sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD4.Text) + "' AND CD_MANA_NO = 'Q0063'"
       txt_FACT_Name4.Text = Gf_FloatFind(M_CN1, sQuery)
       If bClicked = False Then
       
            If Len(Trim(txt_FACT_Name4.Text)) > 1 Then
                 ss1.Row = lngCurRow: ss1.Col = 17: ss1.Text = txt_FACT_CD4.Text
            Else
                 ss1.Row = lngCurRow: ss1.Col = 17: ss1.Text = ""
            End If
       End If
    Else
        txt_FACT_Name4.Text = ""
    End If
    Call Sp_Change(bClicked)
End Sub

Private Sub txt_FACT_CD4_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim str_orgin As String
    Dim sQuery As String
    
    If KeyCode = vbKeyF4 Then
       str_orgin = txt_FACT_CD4.Text
       DD.sWitch = "MS"
       DD.sKey = "Q0063"
       DD.nameType = "2"
       txt_FACT_CD4.Text = ""
       DD.rControl.Add Item:=txt_FACT_CD4
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        If Len(Trim(txt_FACT_CD4.Text)) > 0 Then
            sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD4.Text) + "' AND CD_MANA_NO = 'Q0063'"
            txt_FACT_Name4.Text = Gf_FloatFind(M_CN1, sQuery)
            ss1.Row = lngCurRow: ss1.Col = 17: ss1.Text = txt_FACT_CD4.Text
            Call Sp_Change(bClicked)
        Else
            txt_FACT_CD4.Text = str_orgin
        End If
        'Call ss2_Change(ss2.Col, ss2.Row)
    End If
End Sub

Private Sub txt_FACT_CD5_Change()
    Dim sQuery As String

    If Len(Trim(txt_FACT_CD5.Text)) > 0 Then
       sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD5.Text) + "' AND CD_MANA_NO = 'Q0063'"
       txt_FACT_Name5.Text = Gf_FloatFind(M_CN1, sQuery)
       If bClicked = False Then
            If Len(Trim(txt_FACT_Name5.Text)) > 1 Then
                 ss1.Row = lngCurRow: ss1.Col = 18: ss1.Text = txt_FACT_CD5.Text
            Else
                 ss1.Row = lngCurRow: ss1.Col = 18: ss1.Text = ""
            End If
       End If
    Else
        txt_FACT_Name5.Text = ""
    End If
    Call Sp_Change(bClicked)
End Sub

Private Sub txt_FACT_CD5_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim str_orgin As String
    Dim sQuery As String
    
    If KeyCode = vbKeyF4 Then
       str_orgin = txt_FACT_CD5.Text
       DD.sWitch = "MS"
       DD.sKey = "Q0063"
       DD.nameType = "2"
       txt_FACT_CD5.Text = ""
       DD.rControl.Add Item:=txt_FACT_CD5
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        If Len(Trim(txt_FACT_CD5.Text)) > 0 Then
            sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD5.Text) + "' AND CD_MANA_NO = 'Q0063'"
            txt_FACT_Name5.Text = Gf_FloatFind(M_CN1, sQuery)
            ss1.Row = lngCurRow: ss1.Col = 18: ss1.Text = txt_FACT_CD5.Text
            Call Sp_Change(bClicked)
        Else
            txt_FACT_CD5.Text = str_orgin
        End If
        'Call ss2_Change(ss2.Col, ss2.Row)
    End If
End Sub

Private Sub txt_FACT_CD6_Change()
    Dim sQuery As String

    If Len(Trim(txt_FACT_CD6.Text)) > 0 Then
       sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD6.Text) + "' AND CD_MANA_NO = 'Q0063'"
       txt_FACT_Name6.Text = Gf_FloatFind(M_CN1, sQuery)
       If bClicked = False Then
            If Len(Trim(txt_FACT_Name6.Text)) > 1 Then
                 ss1.Row = lngCurRow: ss1.Col = 19: ss1.Text = txt_FACT_CD6.Text
            Else
                 ss1.Row = lngCurRow: ss1.Col = 19: ss1.Text = ""
            End If
       End If
    Else
        txt_FACT_Name6.Text = ""
    End If
    Call Sp_Change(bClicked)
End Sub

Private Sub txt_FACT_CD6_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim str_orgin As String
    Dim sQuery As String
    
    If KeyCode = vbKeyF4 Then
       str_orgin = txt_FACT_CD6.Text
       DD.sWitch = "MS"
       DD.sKey = "Q0063"
       DD.nameType = "2"
       txt_FACT_CD6.Text = ""
       DD.rControl.Add Item:=txt_FACT_CD6
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        If Len(Trim(txt_FACT_CD2.Text)) > 0 Then
            sQuery = "SELECT CD_NAME FROM ZP_CD WHERE CD = '" + Trim(txt_FACT_CD6.Text) + "' AND CD_MANA_NO = 'Q0063'"
            txt_FACT_Name6.Text = Gf_FloatFind(M_CN1, sQuery)
            ss1.Row = lngCurRow: ss1.Col = 19: ss1.Text = txt_FACT_CD6.Text
            Call Sp_Change(bClicked)
        Else
            txt_FACT_CD6.Text = str_orgin
        End If
        'Call ss2_Change(ss2.Col, ss2.Row)
    End If
End Sub

Private Sub txt_MACRO_Change()
    If bClicked = True Then Exit Sub
    
    ss1.Row = lngCurRow: ss1.Col = 13
    ss1.Text = txt_macro.Text
    Call Sp_Change(bClicked)
End Sub

Private Sub txt_S_Print_Change()
    If bClicked = True Then Exit Sub
    
    ss1.Row = lngCurRow: ss1.Col = 12
    ss1.Text = txt_s_print.Text
    Call Sp_Change(bClicked)
End Sub

Private Sub CtrlCls()
    SSFrame1.Caption = ""
    txt_Slab_NO.Text = ""
    txt_smp_loc.Text = ""
    txt_loc_name.Text = ""
    txt_s_print.Text = ""
    txt_macro.Text = ""
    txt_DCS_CD.Text = ""
    txt_STLGRD.Text = ""
    txt_STLGRD_DETAIL.Text = ""
    txt_FACT_CD1.Text = "": txt_FACT_Name1.Text = ""
    txt_FACT_CD2.Text = "": txt_FACT_Name2.Text = ""
    txt_FACT_CD3.Text = "": txt_FACT_Name3.Text = ""
    txt_FACT_CD4.Text = "": txt_FACT_Name4.Text = ""
    txt_FACT_CD5.Text = "": txt_FACT_Name5.Text = ""
    txt_FACT_CD6.Text = "": txt_FACT_Name6.Text = ""
End Sub
Private Sub Sp_To_Ctrl(ByVal iRow As Integer)
    With ss1
        .Row = iRow
        .Col = 1: SSFrame1.Caption = Trim(.Text) + "---试样实绩信息"
        .Col = 2: txt_Slab_NO.Text = Trim(.Text)
        .Col = 7: txt_STLGRD.Text = Trim(.Text)
        .Col = 8: txt_STLGRD_DETAIL.Text = Trim(.Text)
        .Col = 10: txt_smp_loc.Text = Trim(.Text)
        'txt_loc_name.Text = ""
        .Col = 11: txt_DCS_CD.Text = Trim(.Text)
        .Col = 12: txt_s_print.Text = Trim(.Text)
        .Col = 13: txt_macro.Text = Trim(.Text)
        
        .Col = 14: txt_FACT_CD1.Text = Trim(.Text)
        'txt_FACT_Name1.Text = ""
        .Col = 15: txt_FACT_CD2.Text = Trim(.Text)
        'txt_FACT_Name2.Text = ""
        .Col = 16: txt_FACT_CD3.Text = Trim(.Text)
        'txt_FACT_Name3.Text = ""
        .Col = 17: txt_FACT_CD4.Text = Trim(.Text)
        'txt_FACT_Name4.Text = ""
        .Col = 18: txt_FACT_CD5.Text = Trim(.Text)
        'txt_FACT_Name5.Text = ""
        .Col = 19: txt_FACT_CD6.Text = Trim(.Text)
        'txt_FACT_Name6.Text = ""
    End With
End Sub
Private Sub CtrlLock(Optional bLocked As Boolean = True)
    SSFrame1.Enabled = Not bLocked
End Sub

Private Sub Sp_Change(Optional bClick As Boolean = False)
    If bClick = False Then
        Call Gp_Sp_UpdateMake(ss1, 2)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), 23)
    End If
End Sub
