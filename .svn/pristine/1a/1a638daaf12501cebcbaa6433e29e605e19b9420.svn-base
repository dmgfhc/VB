VERSION 5.00
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGB3010C 
   Caption         =   "母板指示界面_AGB3010C"
   ClientHeight    =   9225
   ClientLeft      =   810
   ClientTop       =   2280
   ClientWidth     =   15450
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15450
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8355
      Left            =   60
      TabIndex        =   1
      Top             =   810
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   14737
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "AGB3010C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   7680
         Left            =   0
         TabIndex        =   3
         Top             =   675
         Width           =   15315
         _ExtentX        =   27014
         _ExtentY        =   13547
         _Version        =   196609
         SplitterBarWidth=   4
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   16761087
         PaneTree        =   "AGB3010C.frx":0052
         Begin FPSpread.vaSpread ss1 
            Height          =   7680
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   15315
            _Version        =   393216
            _ExtentX        =   27014
            _ExtentY        =   13547
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
            MaxCols         =   20
            MaxRows         =   2
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AGB3010C.frx":0084
         End
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   645
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   15315
         _ExtentX        =   27014
         _ExtentY        =   1138
         _Version        =   196609
         BackColor       =   14737918
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSPanel SSPsend 
            Height          =   315
            Left            =   12840
            TabIndex        =   26
            Top             =   180
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   556
            _Version        =   196609
            ForeColor       =   16711680
            BackColor       =   8454143
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "指示已下达"
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
         End
         Begin VB.ComboBox cbo_plt 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            ItemData        =   "AGB3010C.frx":0B4C
            Left            =   13290
            List            =   "AGB3010C.frx":0B56
            TabIndex        =   24
            Text            =   "C1"
            Top             =   120
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.TextBox txt_onoff 
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
            Height          =   330
            Left            =   12750
            MaxLength       =   1
            TabIndex        =   23
            Text            =   " "
            Top             =   120
            Visible         =   0   'False
            Width           =   465
         End
         Begin InDate.ULabel ULabel2 
            Height          =   315
            Left            =   4200
            Top             =   180
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            Caption         =   "剪切线"
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
         Begin Threed.SSPanel SSPanel5 
            Height          =   345
            Left            =   5700
            TabIndex        =   15
            Top             =   210
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption opt_line1 
               Height          =   285
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737918
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
               Caption         =   "#1"
            End
            Begin Threed.SSOption opt_line2 
               Height          =   285
               Left            =   720
               TabIndex        =   17
               Top             =   0
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "#2"
               Value           =   -1
            End
         End
         Begin Threed.SSPanel SSPanel4 
            Height          =   345
            Left            =   630
            TabIndex        =   12
            Top             =   180
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   609
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption opt_mo 
               Height          =   285
               Left            =   0
               TabIndex        =   13
               Top             =   30
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "母板指示"
               Value           =   -1
            End
            Begin Threed.SSOption opt_req 
               Height          =   285
               Left            =   2910
               TabIndex        =   14
               Top             =   30
               Visible         =   0   'False
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "信息请求"
            End
            Begin Threed.SSOption opt_mo_can 
               Height          =   285
               Left            =   1350
               TabIndex        =   18
               Top             =   30
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "母板指示取消"
            End
         End
         Begin Threed.SSPanel SSPanel3 
            Height          =   375
            Left            =   9360
            TabIndex        =   8
            Top             =   210
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   661
            _Version        =   196609
            BackColor       =   14737918
            BevelOuter      =   0
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin Threed.SSOption opt_bed1 
               Height          =   285
               Left            =   0
               TabIndex        =   9
               Top             =   0
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737918
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
               Caption         =   "一号"
            End
            Begin Threed.SSOption opt_bed3 
               Height          =   285
               Left            =   1650
               TabIndex        =   10
               Top             =   0
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   255
               BackColor       =   14737918
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "三号"
               Value           =   -1
            End
            Begin Threed.SSOption opt_bed2 
               Height          =   285
               Left            =   810
               TabIndex        =   11
               Top             =   0
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   196609
               Font3D          =   1
               BackColor       =   14737918
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
               Caption         =   "二号"
            End
         End
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
            Height          =   330
            Left            =   14070
            MaxLength       =   1
            TabIndex        =   7
            Text            =   " "
            Top             =   120
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox txt_cbed_indic 
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
            Height          =   330
            Left            =   14580
            MaxLength       =   2
            TabIndex        =   5
            Text            =   " "
            Top             =   120
            Visible         =   0   'False
            Width           =   465
         End
         Begin InDate.ULabel ULabel19 
            Height          =   315
            Left            =   7860
            Top             =   180
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   556
            Caption         =   "冷床"
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
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   720
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1270
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
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
         ItemData        =   "AGB3010C.frx":0B62
         Left            =   14130
         List            =   "AGB3010C.frx":0B6F
         TabIndex        =   25
         Tag             =   "班次"
         Top             =   180
         Width           =   735
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
         Left            =   5520
         MaxLength       =   12
         TabIndex        =   6
         Tag             =   "作业人员"
         Top             =   180
         Width           =   1635
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   4170
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "母板号"
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
         Left            =   7860
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "生产日期"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption opt_on 
         Height          =   285
         Left            =   1950
         TabIndex        =   19
         Top             =   210
         Width           =   705
         _ExtentX        =   1244
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
         Caption         =   "在线"
         Value           =   -1
      End
      Begin Threed.SSOption opt_off 
         Height          =   285
         Left            =   2790
         TabIndex        =   20
         Top             =   210
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
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
         Caption         =   "离线"
      End
      Begin InDate.UDate udt_date_fr 
         Height          =   315
         Left            =   9210
         TabIndex        =   21
         Tag             =   "INS_DATE"
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
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
      Begin InDate.UDate udt_date_to 
         Height          =   315
         Left            =   10650
         TabIndex        =   22
         Tag             =   "INS_DATE"
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
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
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Left            =   450
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "在/离线"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   12780
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "班次"
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
   End
End
Attribute VB_Name = "AGB3010C"
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
'-- Program Name      MOTHER PLATE SEND L2 界面
'-- Program ID        AGB3010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2010.7.12
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

Dim Proc_Sc As New Collection       'Spread Struc Collection
 
Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection

Dim opt_chk As Boolean

Const SPD_MP = 1
Const SPD_SCH_FL = 9


Private Sub Form_Define()

    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(CBO_PLT, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_onoff, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_PRC_LINE, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_cbed_indic, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(udt_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(udt_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
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
    For iCol = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
   
     'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGB3010C.P_SREFER1", Key:="P-R"
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

    Dim sQuery As String
    
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    txt_PRC_LINE.Text = "2"
    txt_cbed_indic.Text = "30"
    CBO_PLT.Text = "C1"
    txt_onoff.Text = "I"
    opt_chk = True
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_ReadOnlySet(ss1)
    Call Gf_Sp_Cls(sc1)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)

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
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
        
        Call Gp_Ms_Cls(Mc1("rControl"))
        
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        
        CBO_PLT.Text = "C1"
        opt_on.Value = True
'        opt_mo.Value = True
        opt_line2.Value = True
        opt_bed3.Value = True
        
    End If
    
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Master_Cpy()

End Sub

Public Sub Master_Pst()

End Sub

Public Sub Form_Ref()

Dim iRow As Integer
Dim iCol As Integer

    If Gf_Sp_ProceExist(ss1) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
    
        ss1.OperationMode = OperationModeNormal
        Call Gp_Sp_EvenRowBackcolor(ss1)
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    
    End If
    
    With ss1
          For iRow = 1 To .MaxRows
             .Row = iRow:       .Col = SPD_SCH_FL
              If .Text <> "" Then
                  Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , SSPsend.BackColor)
              End If

          Next iRow
'         Call .SetActiveCell(1, .MaxRows)
    End With
    
End Sub

Private Sub opt_bed1_Click(Value As Integer)

    If opt_bed1.Value Then
        opt_bed1.ForeColor = &HFF&
        opt_bed2.ForeColor = &H80000012
        opt_bed3.ForeColor = &H80000012
        txt_cbed_indic.Text = "10"
    End If

End Sub

Private Sub opt_bed2_Click(Value As Integer)

    If opt_bed2.Value Then
        opt_bed1.ForeColor = &H80000012
        opt_bed2.ForeColor = &HFF&
        opt_bed3.ForeColor = &H80000012
        txt_cbed_indic.Text = "20"
    End If
    
End Sub

Private Sub opt_bed3_Click(Value As Integer)

    If opt_bed3.Value Then
        opt_bed1.ForeColor = &H80000012
        opt_bed2.ForeColor = &H80000012
        opt_bed3.ForeColor = &HFF&
        txt_cbed_indic.Text = "30"
    End If

End Sub

Private Sub opt_line1_Click(Value As Integer)

    If opt_line1.Value Then
        opt_line1.ForeColor = &HFF&
        opt_line2.ForeColor = &H80000012
        txt_PRC_LINE.Text = "1"
    End If
    
End Sub

Private Sub opt_line2_Click(Value As Integer)

    If opt_line2.Value Then
        opt_line1.ForeColor = &H80000012
        opt_line2.ForeColor = &HFF&
        txt_PRC_LINE.Text = "2"
    End If

End Sub

Private Sub opt_mo_can_Click(Value As Integer)

    Dim iRow As Integer
    
    Call Form_Ref

    If opt_mo_can.Value Then

        opt_mo.ForeColor = &H80000012
        opt_mo_can.ForeColor = &HFF&
        opt_req.ForeColor = &H80000012

'        Call Gp_Sp_EvenRowBackcolor(ss1)
'
'        For iRow = 1 To ss1.MaxRows
'            ss1.Row = iRow
'            ss1.Col = 0
'            ss1.Text = ""
'        Next iRow

    End If
    
End Sub

Private Sub opt_mo_Click(Value As Integer)

    Dim iRow As Integer
    
    Call Form_Ref

    If opt_mo.Value Then

        opt_mo.ForeColor = &HFF&
        opt_mo_can.ForeColor = &H80000012
        opt_req.ForeColor = &H80000012

'        Call Gp_Sp_EvenRowBackcolor(ss1)
'
'        For iRow = 1 To ss1.MaxRows
'            ss1.Row = iRow
'            ss1.Col = 0
'            ss1.Text = ""
'        Next iRow

    End If

End Sub

Private Sub opt_off_Click(Value As Integer)

    If opt_off.Value Then

        opt_off.ForeColor = &HFF&
        opt_on.ForeColor = &H80000012
        txt_onoff.Text = "O"

    End If

End Sub

Private Sub opt_on_Click(Value As Integer)

    If opt_on.Value Then
        
        opt_on.ForeColor = &HFF&
        opt_off.ForeColor = &H80000012
        txt_onoff.Text = "I"
        
    End If

End Sub

Private Sub opt_req_Click(Value As Integer)

    Dim iRow As Integer
    
    If opt_req.Value Then
        
        opt_mo.ForeColor = &H80000012
        opt_req.ForeColor = &HFF&
        opt_mo_can.ForeColor = &H80000012
        
        Call Gp_Sp_EvenRowBackcolor(ss1)
        
        For iRow = 1 To ss1.MaxRows
            ss1.Row = iRow
            ss1.Col = 0
            ss1.Text = ""
        Next iRow
        
    End If

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim sCh_fl As String

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ss1.MaxRows < 1 Or Row <= 0 Then Exit Sub
    
    ss1.Col = SPD_SCH_FL:     ss1.Row = Row
    sCh_fl = ss1.Text
    
    ss1.Col = 0
    ss1.Row = Row
    
    If ss1.Text = "" Then
    
        If opt_mo_can.Value And sCh_fl = "" Then
            Call Gp_MsgBoxDisplay("母板指示未下达，不允许取消", "Q")
            Exit Sub
        End If
    
        ss1.Text = "选择"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.Row, ss1.Row, , CYAN)
        
    Else
    
        ss1.Text = ""

        If sCh_fl <> "" Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , SSPsend.BackColor)
        Else
            If Row Mod 2 <> 0 Then
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HF2F2F2)
            Else
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFFFF)
            End If
        End If
        
    End If
        
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Public Sub Form_Pro()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrCode As String
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iRow As Integer
    Dim iPro_Cnt As Integer
    Dim sCh_fl As String
        
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    If M_CN1.State = 0 Then
        If GF_DbConnect = False Then Exit Sub
    End If
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = M_CN1
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "AGB3010C.P_MODIFY"
    
    M_CN1.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    iPro_Cnt = 0
    
    'MOTHER PLATE SEND L2
    For iRow = 1 To ss1.MaxRows
    
        ss1.Row = iRow:   ss1.Col = SPD_SCH_FL:   sCh_fl = ss1.Text
    
        ss1.Row = iRow
        ss1.Col = 0
        
        If ss1.Text <> "" Then  '选择
            
            If opt_mo.Value Then
                adoCmd.Parameters(0).Value = "A"
                If sCh_fl <> "" Then
                adoCmd.Parameters(0).Value = "U"
                End If
            ElseIf opt_mo_can.Value Then
                adoCmd.Parameters(0).Value = "D"
            Else
                adoCmd.Parameters(0).Value = "A"
            End If
            
            adoCmd.Parameters(1).Value = "MP"
            ss1.Col = 1
            adoCmd.Parameters(2).Value = ss1.Text
            adoCmd.Parameters(3).Value = txt_PRC_LINE.Text
            adoCmd.Parameters(4).Value = txt_cbed_indic.Text
            adoCmd.Parameters(5).Value = sUserID
                
            adoCmd.Execute
                
            'Error Check
            If adoCmd("Error") <> "0" Then
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(ss1, iRow, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                M_CN1.RollbackTrans
                Set adoCmd = Nothing
                Screen.MousePointer = vbDefault
                Exit Sub
            Else
                iPro_Cnt = iPro_Cnt + 1
            End If
        
        End If
        
    Next iRow
    
    M_CN1.CommitTrans
    Set adoCmd = Nothing
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 0
        ss1.Text = ""
    Next iRow
    
    If iPro_Cnt > 0 Then
        Call Gp_MsgBoxDisplay("母板指示下达完毕", "I")
        Call Form_Ref
    End If
        
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_Error : " & Error)
 
End Sub

Public Sub Form_Del()

End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
    
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
            
    End With

End Sub

