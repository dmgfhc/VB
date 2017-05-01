VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQD0090C 
   Caption         =   "炉号工确认_AQD0090C"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8910
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   15716
      _Version        =   196609
      AutoSize        =   1
      BorderStyle     =   0
      PaneTree        =   "AQD0090C.frx":0000
      Begin TabDlg.SSTab SSTab1 
         Height          =   5070
         Left            =   0
         TabIndex        =   7
         Top             =   3840
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   8943
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "坯料管理确认"
         TabPicture(0)   =   "AQD0090C.frx":0072
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ss2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "SSFrame1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "检验人员确认"
         TabPicture(1)   =   "AQD0090C.frx":008E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "SSFrame2"
         Tab(1).Control(1)=   "ss3"
         Tab(1).ControlCount=   2
         Begin Threed.SSFrame SSFrame1 
            Height          =   4395
            Left            =   150
            TabIndex        =   8
            Top             =   570
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   7752
            _Version        =   196609
            Begin VB.CommandButton CMD_WEIGHT_CHECK_OK 
               Caption         =   "确认"
               Height          =   315
               Left            =   120
               TabIndex        =   12
               Top             =   3690
               Width           =   1335
            End
            Begin VB.TextBox txt_ORD_LEN 
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   3090
               Width           =   2505
            End
            Begin VB.TextBox txt_SLAB_SIZE 
               Height          =   315
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   1890
               Width           =   2505
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   2
               Left            =   120
               Top             =   90
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
               ForeColor       =   0
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   3
               Left            =   120
               Top             =   1290
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               Caption         =   "规格"
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
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   4
               Left            =   120
               Top             =   2490
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               Caption         =   "定尺长度"
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
            Begin InDate.ULabel ul_HEAT_NO 
               Height          =   315
               Left            =   120
               Top             =   690
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   556
               Caption         =   ""
               Alignment       =   0
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16384
            End
         End
         Begin FPSpread.vaSpread ss2 
            Height          =   4380
            Left            =   3690
            TabIndex        =   9
            Top             =   570
            Width           =   11460
            _Version        =   393216
            _ExtentX        =   20214
            _ExtentY        =   7726
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
            MaxCols         =   14
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AQD0090C.frx":00AA
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   4395
            Left            =   -74910
            TabIndex        =   13
            Top             =   450
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   7752
            _Version        =   196609
            Begin VB.CheckBox check1 
               Caption         =   "全选"
               Height          =   285
               Left            =   2310
               TabIndex        =   23
               Top             =   225
               Width           =   915
            End
            Begin VB.TextBox txt_REMARK 
               Height          =   2205
               Left            =   240
               MultiLine       =   -1  'True
               TabIndex        =   16
               Top             =   2070
               Width           =   2985
            End
            Begin VB.CommandButton cmd_print_card 
               Caption         =   "打印平衡卡"
               Height          =   315
               Left            =   240
               TabIndex        =   14
               Top             =   240
               Width           =   1335
            End
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   5
               Left            =   240
               Top             =   1590
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               Caption         =   "备注"
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
            Begin InDate.ULabel ULabel1 
               Height          =   315
               Index           =   6
               Left            =   240
               Top             =   630
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
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   0
            End
            Begin InDate.ULabel UL_HEAT_NO_P 
               Height          =   315
               Left            =   240
               Top             =   1110
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   556
               Caption         =   ""
               Alignment       =   0
               BackColor       =   14804173
               BackgroundStyle =   1
               ChiselText      =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   16384
            End
         End
         Begin FPSpread.vaSpread ss3 
            Height          =   4380
            Left            =   -71310
            TabIndex        =   15
            Top             =   510
            Width           =   11460
            _Version        =   393216
            _ExtentX        =   20214
            _ExtentY        =   7726
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
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AQD0090C.frx":1BF5
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   990
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15240
         _ExtentX        =   26882
         _ExtentY        =   1746
         _Version        =   196609
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin VB.TextBox txt_TEST_STS 
            Height          =   315
            Left            =   9720
            TabIndex        =   22
            Text            =   "A"
            Top             =   630
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txt_STL_GRD_NAME 
            Height          =   315
            Left            =   2670
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   600
            Width           =   2925
         End
         Begin VB.TextBox txt_STL_GRD_CD 
            Height          =   315
            Left            =   1560
            MaxLength       =   11
            TabIndex        =   4
            Top             =   600
            Width           =   1095
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Index           =   0
            Left            =   3450
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            Caption         =   "～"
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
            ForeColor       =   16576
         End
         Begin VB.TextBox txt_CHARGE_NO_MAX 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   3
            Top             =   90
            Width           =   1815
         End
         Begin VB.TextBox txt_CHARGE_NO_MIN 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1560
            MaxLength       =   8
            TabIndex        =   2
            Top             =   120
            Width           =   1815
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   0
            Left            =   120
            Top             =   120
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
            ForeColor       =   0
         End
         Begin InDate.ULabel ULabel1 
            Height          =   315
            Index           =   1
            Left            =   120
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "钢种"
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
            ForeColor       =   12582912
         End
         Begin InDate.UDate udt_DATE_MAX 
            Height          =   315
            Left            =   9015
            TabIndex        =   17
            Top             =   120
            Width           =   1515
            _ExtentX        =   2672
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
         Begin InDate.UDate udt_DATE_MIN 
            Height          =   315
            Left            =   7335
            TabIndex        =   18
            Top             =   120
            Width           =   1515
            _ExtentX        =   2672
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
         Begin InDate.ULabel ULabel3 
            Height          =   315
            Left            =   5940
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            Caption         =   "确认日期"
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
         Begin Threed.SSOption opt_TEST_STS 
            Height          =   300
            Index           =   0
            Left            =   6000
            TabIndex        =   19
            Top             =   615
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   196609
            BackColor       =   14737632
            Caption         =   "未打印平衡卡"
            Value           =   -1
         End
         Begin Threed.SSOption opt_TEST_STS 
            Height          =   300
            Index           =   1
            Left            =   7380
            TabIndex        =   20
            Top             =   615
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   196609
            BackColor       =   14737632
            Caption         =   "已打印平衡卡"
         End
         Begin Threed.SSOption opt_TEST_STS 
            Height          =   300
            Index           =   2
            Left            =   8775
            TabIndex        =   21
            Top             =   615
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   529
            _Version        =   196609
            BackColor       =   14737632
            Caption         =   "全部"
         End
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   2670
         Left            =   0
         TabIndex        =   6
         Top             =   1080
         Width           =   15240
         _Version        =   393216
         _ExtentX        =   26882
         _ExtentY        =   4710
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
         MaxCols         =   12
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQD0090C.frx":37FF
      End
   End
End
Attribute VB_Name = "AQD0090C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      质量证明书二次发放
'-- Program ID        AQD0030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Chu Kyo Su
'-- Coder             Chu Kyo Su
'-- Date              2003.07. 25
'-- Description       质量证明书二次发放
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

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

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

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection
Dim bPrintCheck As Boolean

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'------------------------------ Report Variable ----------------------------------------------
'---------------------------------------------------------------------------------------------

Dim sQuery      As String
Dim sErrMsg     As String
Dim sDate       As String
Dim AdoRs       As adodb.Recordset


'---------------------------------------------------------------------------------------------

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_CHARGE_NO_MIN, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
       Call Gp_Ms_Collection(txt_CHARGE_NO_MAX, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
          Call Gp_Ms_Collection(txt_STL_GRD_CD, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
        Call Gp_Ms_Collection(txt_STL_GRD_NAME, " ", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(udt_DATE_MIN, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(udt_DATE_MAX, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
            Call Gp_Ms_Collection(txt_TEST_STS, "p", " ", " ", " ", "r", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)

    
    'MASTER Collection
    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=cControl1, Key:="cControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(ul_HEAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(UL_HEAT_NO_P, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc3.Add Item:=pControl2, Key:="pControl"
    Mc3.Add Item:=nControl2, Key:="nControl"
    Mc3.Add Item:=mControl2, Key:="mControl"
    Mc3.Add Item:=iControl2, Key:="iControl"
    Mc3.Add Item:=rControl2, Key:="rControl"
    Mc3.Add Item:=cControl2, Key:="cControl"
    Mc3.Add Item:=aControl2, Key:="aControl"
    Mc3.Add Item:=lControl2, Key:="lControl"
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
     Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQD0090C.P_REF_1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
'    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQD0090C.P_REF_2", Key:="P-R"
    sc2.Add Item:="AQD0090C.P_MODIFY_1", Key:="P-M"
    sc2.Add Item:="AQD0090C.P_ONEROW", Key:="P-O"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", "i", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 2, "p", "n", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 3, "p", "n", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
     Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 14, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", "i", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 16, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AQD0090C.P_REF_3", Key:="P-R"
    sc3.Add Item:="AQD0090C.P_MODIFY_2", Key:="P-M"
    sc3.Add Item:="AQD0090C.P_ONEROW", Key:="P-O"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=sc3, Key:="Sc3"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub check1_Click()
Dim iRow        As Long

If check1.Value = 1 Then
    With ss3
            For iRow = 1 To .MaxRows
                .Row = iRow
                .Col = 1
                .Value = 1
                .Col = 0
                .Text = "Update"
            Next iRow
        End With
ElseIf check1.Value = 0 Then
    With ss3
            For iRow = 1 To .MaxRows
                .Row = iRow
                .Col = 1
                .Value = 0
                .Col = 0
                .Text = ""
            Next iRow
        End With
End If


End Sub

Private Sub cmd_print_card_Click()
Dim iRow        As Long
Dim sREMARK     As String
Dim iMaxrow     As Long
Dim sHEAT_NO    As String

If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub

sREMARK = Trim(txt_REMARK.Text)
sHEAT_NO = Trim(UL_HEAT_NO_P.Caption)
        With ss3
            For iRow = 1 To .MaxRows
                .Row = iRow
                .Col = 14
                .Text = sREMARK
                .Col = 15
                .Text = sUserID
                .Col = 8
                .Text = sUserID
'                .Col = 0
'                .Text = "Update"
            Next iRow
        End With
      
      With ss3
          For iRow = 1 To .MaxRows
                .Row = iRow
                .Col = 1
                If .Value = 1 Then
                 .Col = 0
                 .Text = "Update"
                Else:
                  .Col = 0
                  .Text = ""
                End If
            Next iRow
      End With
iRow = ss1.Row
iMaxrow = ss1.MaxRows

If Gf_Sp_Process(M_CN1, Proc_Sc("Sc3"), Mc1, True) Then
   Call funGetQuery_D(sHEAT_NO)
   Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl"), Mc1("mControl"))
   Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
   Call Gp_Goto_Row(ss1, iMaxrow, iRow)
   Call subButtonHide
End If
End Sub

Private Sub CMD_WEIGHT_CHECK_OK_Click()
Dim iRow        As Long
Dim sSLAB_SIZE  As String
Dim sORD_LEN    As String
Dim iMaxrow As Long

If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub

sSLAB_SIZE = Trim(txt_SLAB_SIZE.Text)
sORD_LEN = Trim(txt_ORD_LEN.Text)
        With ss2
            For iRow = 1 To .MaxRows
                .Row = iRow
                .Col = 12
                .Text = sORD_LEN
                .Col = 13
                .Text = sSLAB_SIZE
                .Col = 8
                .Text = sUserID
                .Col = 0
                .Text = "Update"
            Next iRow
        End With

iRow = ss1.Row
iMaxrow = ss1.MaxRows

If Gf_Sp_Process(M_CN1, Proc_Sc("Sc2"), Mc1, True) Then
   Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl"), Mc1("mControl"))
   Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
   Call Gp_Goto_Row(ss1, iMaxrow, iRow)

   Call subButtonHide
End If

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
        Case "txt_STL_GRD_CD"           '钢种
            sCode = "STLGRD"
            Set oCodeName = txt_STL_GRD_NAME
            
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    Call subButtonHide

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
    
    Call Gp_Ms_Cls(Mc2("pControl"))
    
    Call Gp_Ms_Cls(Mc3("pControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    txt_TEST_STS.Text = "A"
    
    Call subButtonHide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc2")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc3")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
    
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
    
    Set pControl3 = Nothing
    Set nControl3 = Nothing
    Set iControl3 = Nothing
    Set rControl3 = Nothing
    Set cControl3 = Nothing
    Set aControl3 = Nothing
    Set lControl3 = Nothing
    Set mControl3 = Nothing
    
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
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
    
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set Mc3 = Nothing
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
    Call subButtonHide
    
End Sub



Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc1")) Then
        Call Gf_Sp_Cls(Proc_Sc("Sc2"))
        Call Gf_Sp_Cls(Proc_Sc("Sc3"))
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("pControl"))
        Call Gp_Ms_Cls(Mc3("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc2").Item("Spread")) Then Exit Sub
    If Gf_Sp_ProceExist(Proc_Sc("Sc3").Item("Spread")) Then Exit Sub
    
       
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                ss1.Click
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call subButtonHide
                Exit Sub
            End If
            
    
    Call subButtonHide
        
    Exit Sub

Refer_Err:

End Sub


Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc3")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_TEST_STS_Click(Index As Integer, Value As Integer)
        Select Case Index
        Case 0
            txt_TEST_STS.Text = "A"
'            ULabel5.Visible = True
'            dtp_fr_date.Visible = True
'            dtp_to_date.Visible = True
'            dtp_fr_date.RawData = ""
'            dtp_to_date.RawData = ""

        Case 1
            txt_TEST_STS.Text = "B"
'            ULabel5.Visible = True
'            dtp_fr_date.Visible = True
'            dtp_to_date.Visible = True
        Case 2
            txt_TEST_STS.Text = "C"
'            ULabel5.Visible = True
'            dtp_fr_date.Visible = True
'            dtp_to_date.Visible = True

    End Select
    Call Form_Ref

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    With ss1
        .Row = .ActiveRow
        .Col = 1
        ul_HEAT_NO.Caption = .Text
        UL_HEAT_NO_P.Caption = .Text
        
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2, Mc2("nControl"), Mc2("mControl"), False)
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc3"), Mc3, Mc3("nControl"), Mc3("mControl"), False)
        
        .Col = 4:  txt_SLAB_SIZE.Text = .Text
        .Col = 5:  txt_ORD_LEN.Text = .Text
        .Col = 9:  txt_REMARK = .Text
        
        ss3.OperationMode = OperationModeNormal
    End With
    
    check1.Value = 0
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub subButtonHide()

    MDIMain.MenuTool.Buttons(5).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    
    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
    

End Sub
Public Sub Form_Pro()
'    Dim iMaxrow As Long
'    Dim sMesg As String
'    Dim iRow As Long
'    Dim sHEAT_NO    As String
'    Dim i As Long
'
'    iRow = ss1.Row
'    iMaxrow = ss1.MaxRows
'
'    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc1"), Mc1, True) Then
'        With ss1
'            For i = 1 To .MaxRows
'                .Row = i
'                .Col = 1
'                If .Text = "1" Then
'                    .Col = 3:     sHEAT_NO = Trim(.Text)
'                    sMesg = funGetQuery_D(sHEAT_NO)
'
'                    If sMesg <> "" Then
'                        i = .MaxRows
'                    Else
'                        .Row = i
'                        .Col = 1
'                        .Text = ""
'                    End If
'                End If
'            Next i
'        End With
'        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'        Call Gp_Goto_Row(ss1, iMaxrow, iRow)
'        Call Sp_to_Ms(ss1, iRow)
'        Call subButtonHide
'    End If
        
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Long
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
           
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc2")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc2")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc2"), 8)
    End If
    
End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Long
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
           
End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
  
    
    Call Gp_Sp_Sort(Proc_Sc("Sc3")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc3")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), 8)
    End If
    
End Sub

Private Sub ss3_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc2"))
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc3"))
      
End Sub
