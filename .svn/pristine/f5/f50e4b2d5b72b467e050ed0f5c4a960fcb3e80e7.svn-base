VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKN2050C 
   BackColor       =   &H00C0C0C0&
   Caption         =   "指定及解除连铸作业指示锁定界面_AKN2050C"
   ClientHeight    =   9285
   ClientLeft      =   210
   ClientTop       =   2235
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   270
      Top             =   120
   End
   Begin VB.TextBox txt_ccm_line 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   480
      MaxLength       =   1
      TabIndex        =   13
      Tag             =   "工厂"
      Top             =   150
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txt_plt 
      Alignment       =   2  'Center
      Enabled         =   0   'False
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
      Left            =   180
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "工厂"
      Top             =   450
      Width           =   465
   End
   Begin VB.TextBox txt_plt_name 
      Enabled         =   0   'False
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
      Left            =   645
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   450
      Width           =   1290
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8265
      Left            =   45
      TabIndex        =   0
      Top             =   990
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   14579
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AKN2050C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   2820
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   5250
         _Version        =   393216
         _ExtentX        =   9260
         _ExtentY        =   4974
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   14
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2050C.frx":00D2
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   5385
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2880
         Width           =   5250
         _Version        =   393216
         _ExtentX        =   9260
         _ExtentY        =   9499
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   22
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2050C.frx":08FA
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   2820
         Left            =   5310
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   4830
         _Version        =   393216
         _ExtentX        =   8520
         _ExtentY        =   4974
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   14
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2050C.frx":14A7
      End
      Begin FPSpread.vaSpread ss5 
         Height          =   2820
         Left            =   10200
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   4950
         _Version        =   393216
         _ExtentX        =   8731
         _ExtentY        =   4974
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   14
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2050C.frx":1CCF
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   5385
         Left            =   5310
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2880
         Width           =   4830
         _Version        =   393216
         _ExtentX        =   8520
         _ExtentY        =   9499
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   22
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2050C.frx":24F7
      End
      Begin FPSpread.vaSpread ss6 
         Height          =   5385
         Left            =   10200
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2880
         Width           =   4950
         _Version        =   393216
         _ExtentX        =   8731
         _ExtentY        =   9499
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   22
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2050C.frx":307C
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   180
      Top             =   90
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      Caption         =   "工厂"
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
   Begin Threed.SSFrame Frame3 
      Height          =   465
      Left            =   7665
      TabIndex        =   3
      Top             =   90
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox txt_to 
         Enabled         =   0   'False
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
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Tag             =   "终止板坯号"
         Top             =   60
         Width           =   1170
      End
      Begin VB.TextBox txt_charge 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   645
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "炉号"
         Top             =   60
         Width           =   1110
      End
      Begin VB.TextBox txt_from 
         Enabled         =   0   'False
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
         Left            =   3375
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Tag             =   "起始板坯号"
         Top             =   60
         Width           =   1170
      End
      Begin Threed.SSOption opt_from 
         Height          =   315
         Left            =   2010
         TabIndex        =   4
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "起始板坯号"
         Value           =   -1
      End
      Begin Threed.SSOption opt_to 
         Height          =   315
         Left            =   4740
         TabIndex        =   8
         Top             =   75
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "终止板坯号"
      End
      Begin Threed.SSPanel SSPanel2 
         Height          =   300
         Left            =   120
         TabIndex        =   24
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   529
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   0
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "炉号"
         BevelOuter      =   0
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin Threed.SSPanel SSPpdt 
      Height          =   300
      Left            =   13815
      TabIndex        =   11
      Top             =   600
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   529
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "已下达"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPsend 
      Height          =   300
      Left            =   13140
      TabIndex        =   12
      Top             =   600
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   529
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "锁定"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSOption opt_time_on 
      Height          =   285
      Left            =   10710
      TabIndex        =   18
      Top             =   615
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "动态进程"
      Value           =   -1
   End
   Begin Threed.SSOption opt_time_off 
      Height          =   285
      Left            =   11925
      TabIndex        =   19
      Top             =   615
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "静态进程"
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   300
      Left            =   14490
      TabIndex        =   20
      Top             =   600
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   529
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "钢种变"
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSFrame Frame1 
      Height          =   465
      Left            =   2070
      TabIndex        =   21
      Top             =   90
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      Enabled         =   0   'False
      ShadowStyle     =   1
      Begin Threed.SSOption opt_lock 
         Height          =   285
         Left            =   150
         TabIndex        =   22
         Top             =   90
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "解锁 -> 锁定"
         Value           =   -1
      End
      Begin Threed.SSOption opt_unlock 
         Height          =   285
         Left            =   2025
         TabIndex        =   23
         Top             =   90
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16711680
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "锁定 -> 解锁"
      End
   End
   Begin Threed.SSFrame Frame2 
      Height          =   465
      Left            =   5970
      TabIndex        =   25
      Top             =   90
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   820
      _Version        =   196609
      BackColor       =   14737632
      Enabled         =   0   'False
      ShadowStyle     =   1
      Begin Threed.SSOption opt_charge 
         Height          =   285
         Left            =   120
         TabIndex        =   26
         Top             =   90
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "炉"
         Value           =   -1
      End
      Begin Threed.SSOption opt_slab 
         Height          =   285
         Left            =   765
         TabIndex        =   27
         Top             =   90
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   12583104
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "板坯"
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   30
      X2              =   15185
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "AKN2050C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       System
'-- Sub_System Name
'-- Program Name
'-- Program ID        AKN2050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              20011.9.8
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

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection

Dim pColumn5 As New Collection      'Spread Primary Key Collection
Dim nColumn5 As New Collection      'Spread necessary Column Collection
Dim mColumn5 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn5 As New Collection      'Spread Insert Column Collection
Dim aColumn5 As New Collection      'Master -> Spread Column Collection
Dim lColumn5 As New Collection      'Spread Lock Column Collection

Dim pColumn6 As New Collection      'Spread Primary Key Collection
Dim nColumn6 As New Collection      'Spread necessary Column Collection
Dim mColumn6 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn6 As New Collection      'Spread Insert Column Collection
Dim aColumn6 As New Collection      'Master -> Spread Column Collection
Dim lColumn6 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Sc4 As New Collection           'Spread Collection
Dim Sc5 As New Collection           'Spread Collection
Dim Sc6 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim P_Fr_Edt_Seq As Long            'Slab_Edt_Seq (From)
Dim P_To_Edt_Seq As Long            'Slab_Edt_Seq (To)
Dim P_Tr_Edt_Seq As Long            'Slab_Edt_Seq (Target)

Private Sub Form_Define()
        
    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_plt, "p", "n", "m", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_plt_name, " ", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_ccm_line, "p", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        
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
    For iCol = 1 To ss1.MaxCols
     Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AFN2050C.P_REFER1", Key:="P-R"
    
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    
    'Spread_Collection
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="AFN2050C.P_REFER2", Key:="P-R"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"
   
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss3.MaxCols
        Call Gp_Sp_Collection(ss3, iCol, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Next iCol
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AFN2050C.P_REFER1", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss4.MaxCols
        Call Gp_Sp_Collection(ss4, iCol, " ", " ", " ", " ", " ", "l", pColumn4, nColumn4, mColumn4, iColumn4, aColumn4, lColumn4)
    Next iCol
    
    'Spread_Collection
    Sc4.Add Item:=ss4, Key:="Spread"
    Sc4.Add Item:="AFN2050C.P_REFER2", Key:="P-R"
    Sc4.Add Item:=pColumn4, Key:="pColumn"
    Sc4.Add Item:=nColumn4, Key:="nColumn"
    Sc4.Add Item:=aColumn4, Key:="aColumn"
    Sc4.Add Item:=mColumn4, Key:="mColumn"
    Sc4.Add Item:=iColumn4, Key:="iColumn"
    Sc4.Add Item:=lColumn4, Key:="lColumn"
    Sc4.Add Item:=1, Key:="First"
    Sc4.Add Item:=ss4.MaxCols, Key:="Last"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss5.MaxCols
        Call Gp_Sp_Collection(ss5, iCol, " ", " ", " ", " ", " ", "l", pColumn5, nColumn5, mColumn5, iColumn5, aColumn5, lColumn5)
    Next iCol
    
    'Spread_Collection
    Sc5.Add Item:=ss5, Key:="Spread"
    Sc5.Add Item:="AFN2050C.P_REFER1", Key:="P-R"
    Sc5.Add Item:=pColumn5, Key:="pColumn"
    Sc5.Add Item:=nColumn5, Key:="nColumn"
    Sc5.Add Item:=aColumn5, Key:="aColumn"
    Sc5.Add Item:=mColumn5, Key:="mColumn"
    Sc5.Add Item:=iColumn5, Key:="iColumn"
    Sc5.Add Item:=lColumn5, Key:="lColumn"
    Sc5.Add Item:=1, Key:="First"
    Sc5.Add Item:=ss5.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss6.MaxCols
        Call Gp_Sp_Collection(ss6, iCol, " ", " ", " ", " ", " ", "l", pColumn6, nColumn6, mColumn6, iColumn6, aColumn6, lColumn6)
    Next iCol
    
    'Spread_Collection
    Sc6.Add Item:=ss6, Key:="Spread"
    Sc6.Add Item:="AFN2050C.P_REFER2", Key:="P-R"
    Sc6.Add Item:=pColumn6, Key:="pColumn"
    Sc6.Add Item:=nColumn6, Key:="nColumn"
    Sc6.Add Item:=aColumn6, Key:="aColumn"
    Sc6.Add Item:=mColumn6, Key:="mColumn"
    Sc6.Add Item:=iColumn6, Key:="iColumn"
    Sc6.Add Item:=lColumn6, Key:="lColumn"
    Sc6.Add Item:=1, Key:="First"
    Sc6.Add Item:=ss6.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    Proc_Sc.Add Item:=Sc4, Key:="Sc4"
    Proc_Sc.Add Item:=Sc5, Key:="Sc5"
    Proc_Sc.Add Item:=Sc6, Key:="Sc6"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 2, True)
    Call Gp_Sp_ColHidden(ss1, 5, True)
    Call Gp_Sp_ColHidden(ss1, 14, True)
    
    Call Gp_Sp_ColHidden(ss2, 1, True)
    Call Gp_Sp_ColHidden(ss2, 19, True)
    Call Gp_Sp_ColHidden(ss2, 20, True)
    Call Gp_Sp_ColHidden(ss2, 21, True)
    
    Call Gp_Sp_ColHidden(ss3, 2, True)
    Call Gp_Sp_ColHidden(ss3, 5, True)
    Call Gp_Sp_ColHidden(ss3, 14, True)
    
    Call Gp_Sp_ColHidden(ss4, 1, True)
    Call Gp_Sp_ColHidden(ss4, 19, True)
    Call Gp_Sp_ColHidden(ss4, 20, True)
    Call Gp_Sp_ColHidden(ss4, 21, True)
    
    Call Gp_Sp_ColHidden(ss5, 2, True)
    Call Gp_Sp_ColHidden(ss5, 5, True)
    Call Gp_Sp_ColHidden(ss5, 14, True)
    
    Call Gp_Sp_ColHidden(ss6, 1, True)
    Call Gp_Sp_ColHidden(ss6, 19, True)
    Call Gp_Sp_ColHidden(ss6, 20, True)
    Call Gp_Sp_ColHidden(ss6, 21, True)
    
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
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc4.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc5.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc6.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc2.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc4.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc5.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc6.Item("Spread"))
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(Sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(Sc4)
    Call Gf_Sp_Cls(Sc5)
    Call Gf_Sp_Cls(Sc6)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "K-System.INI", Me.Name)

    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc2.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc4.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc5.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc6.Item("Spread"), "K-System.INI", Me.Name)
    
    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    
    Call Form_Ref
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Spl_SizeSet(SSSplitter1, "K-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc2.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc4.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc5.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc6.Item("Spread"), "K-System.INI", Me.Name)
    
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
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing
    
    Set iColumn5 = Nothing
    Set pColumn5 = Nothing
    Set lColumn5 = Nothing
    Set nColumn5 = Nothing
    Set mColumn5 = Nothing
    Set aColumn5 = Nothing
    
    Set iColumn6 = Nothing
    Set pColumn6 = Nothing
    Set lColumn6 = Nothing
    Set nColumn6 = Nothing
    Set mColumn6 = Nothing
    Set aColumn6 = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set Sc3 = Nothing
    Set Sc4 = Nothing
    Set Sc5 = Nothing
    Set Sc6 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(Sc2)
    Call Gf_Sp_Cls(Sc3)
    Call Gf_Sp_Cls(Sc4)
    Call Gf_Sp_Cls(Sc5)
    Call Gf_Sp_Cls(Sc6)
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    txt_plt.Text = "B1"
    Call txt_plt_KeyUp(0, 0)
    
    opt_lock.Value = True
    opt_from.Value = True
    
    txt_charge.Text = ""
    txt_from.Text = ""
    txt_to.Text = ""
    
    P_Fr_Edt_Seq = 0
    P_To_Edt_Seq = 0
    P_Tr_Edt_Seq = 0
    
End Sub

Public Sub Form_Ref()

    Dim Ref_FL As String
    Dim sQuery As String
    Dim Dynamic_Slab As String
    
    txt_ccm_line.Text = "1"
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    Dynamic_Slab = "SC1"
    sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
        
    If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
        Call Gp_Sp_Scolor(ss1, "Y")
        Call Gp_Sp_Scolor(ss2, "Y")
    Else
        Call Gp_Sp_Scolor(ss1, "N")
        Call Gp_Sp_Scolor(ss2, "N")
    End If
    
    txt_ccm_line.Text = "2"
    If Gf_Sp_Refer(M_CN1, Sc3, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc4, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If

    Dynamic_Slab = "SC2"
    sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
        
    If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
        Call Gp_Sp_Scolor(ss3, "Y")
        Call Gp_Sp_Scolor(ss4, "Y")
    Else
        Call Gp_Sp_Scolor(ss3, "N")
        Call Gp_Sp_Scolor(ss4, "N")
    End If
    
    txt_ccm_line.Text = "3"
    If Gf_Sp_Refer(M_CN1, Sc5, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc6, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        Ref_FL = "1"
    End If
    
    Dynamic_Slab = "SC3"
    sQuery = "SELECT GF_SYSTEM_RUN('" & Dynamic_Slab & "') FROM DUAL "
        
    If Gf_CodeFind(M_CN1, sQuery) = "Y" Then
        Call Gp_Sp_Scolor(ss5, "Y")
        Call Gp_Sp_Scolor(ss6, "Y")
    Else
        Call Gp_Sp_Scolor(ss5, "N")
        Call Gp_Sp_Scolor(ss6, "N")
    End If
    
    If Ref_FL = "1" Then
    
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        
        ss1.OperationMode = OperationModeNormal
        ss2.OperationMode = OperationModeNormal
        ss3.OperationMode = OperationModeNormal
        ss4.OperationMode = OperationModeNormal
        ss5.OperationMode = OperationModeNormal
        ss6.OperationMode = OperationModeNormal
        
        Call Spread_Color_Setting(ss1)
        Call Spread_Color_Setting(ss2)
        Call Spread_Color_Setting(ss3)
        Call Spread_Color_Setting(ss4)
        Call Spread_Color_Setting(ss5)
        Call Spread_Color_Setting(ss6)
        
        txt_charge.Text = ""
        txt_from.Text = ""
        txt_to.Text = ""
        
        P_Fr_Edt_Seq = 0
        P_To_Edt_Seq = 0
        P_Tr_Edt_Seq = 0
        
    End If
    
    If opt_time_on Then
    
        Frame1.Enabled = False
        Frame2.Enabled = False
        Frame3.Enabled = False
        
        opt_lock.Value = False
        opt_unlock.Value = False
        
        opt_charge.Value = False
        opt_slab.Value = False
        
        opt_from.Enabled = True
        txt_from.Enabled = True
        opt_from.Value = False
        opt_to.Value = False
        
        txt_from.Left = 3375
        txt_to.Left = 6120
        
        opt_from.Caption = "起始板坯号"
        opt_to.Caption = "终止板坯号"
        
        txt_charge.Text = ""
        txt_from.Text = ""
        txt_to.Text = ""
        
    Else
    
        opt_lock.Value = True
    
    End If
            
End Sub

Public Sub Spread_Forzens_Setting()
    
    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Active_Spread, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Pro()

On Error GoTo Process_Error

    Dim OutParam(1, 4) As Variant
    Dim errMsg As String
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    
    Dim sFrom_No As String
    Dim sTo_No As String
    Dim sTarget_No As String
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
     
    sQuery = "{call AFN2050C.P_MODIFY ('" & txt_plt.Text & "','" & IIf(opt_lock.Value, 1, 0) & "', '" & IIf(opt_charge.Value, 1, 0) & "', '" & txt_charge.Text & "', '" & txt_from.Text & "','" & txt_to.Text & "','" & sUserID & "',?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    M_CN1.BeginTrans
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        errMsg = sErrMessg
        M_CN1.RollbackTrans
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Exit Sub
    End If
    
    M_CN1.CommitTrans
    Set adoCmd = Nothing
    
    Call Form_Ref
    
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Error:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    sErrMessg = Err.Description & sQuery
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub

Public Sub Spread_Del()
    
End Sub

Private Sub opt_charge_Click(Value As Integer)

    If opt_charge.Value = False Then Exit Sub
    
    opt_from.Enabled = True
    opt_from.Value = True
    opt_to.Enabled = False
    
    txt_from.Left = 3180
    txt_to.Left = 5900
    
    opt_from.Caption = "起始炉号"
    opt_to.Caption = "终止炉号"
    
    txt_charge.Text = ""
    txt_from.Text = ""
    txt_to.Text = ""
    
End Sub

Private Sub opt_lock_Click(Value As Integer)

    If opt_lock.Value = False Then Exit Sub
    
    Frame2.Enabled = False
    SSPanel2.Enabled = True
    
    opt_charge.Value = False
    opt_slab.Value = False
    
    opt_from.Enabled = False
    opt_to.Enabled = True
    opt_to.Value = True
    
    txt_from.Left = 3375
    txt_to.Left = 6120
    
    opt_from.Caption = "起始板坯号"
    opt_to.Caption = "终止板坯号"
    
    txt_charge.Text = ""
    txt_from.Text = ""
    txt_to.Text = ""

End Sub

Private Sub opt_slab_Click(Value As Integer)

    If opt_slab.Value = False Then Exit Sub
    
    opt_from.Enabled = True
    opt_from.Value = True
    opt_to.Enabled = False
    
    txt_from.Left = 3375
    txt_to.Left = 6120
    
    opt_from.Caption = "起始板坯号"
    opt_to.Caption = "终止板坯号"
    
    txt_charge.Text = ""
    txt_from.Text = ""
    txt_to.Text = ""

End Sub

Private Sub opt_time_off_Click(Value As Integer)

    If opt_time_off.Value = False Then Exit Sub
    
    Frame1.Enabled = True
    Frame2.Enabled = False
    Frame3.Enabled = True
    
    SSPanel2.Enabled = True
    
    opt_lock.Value = True
    
    opt_from.Enabled = False
    txt_from.Enabled = False
    
    opt_to.Value = True
    
    txt_from.Left = 3375
    txt_to.Left = 6120
    
    opt_from.Caption = "起始板坯号"
    opt_to.Caption = "终止板坯号"
    
    opt_time_on.Value = False
    
    Timer1.Enabled = False
    
End Sub

Private Sub opt_time_on_Click(Value As Integer)

    If opt_time_on.Value = False Then Exit Sub
    
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False
    
    SSPanel2.Enabled = True
    
    opt_lock.Value = False
    opt_unlock.Value = False
    opt_charge.Value = False
    opt_slab.Value = False
    opt_from.Enabled = True
    txt_from.Enabled = True
    opt_from.Value = False
    opt_to.Value = False
    
    opt_time_off.Value = False
    
    txt_from.Left = 3375
    txt_to.Left = 6120
    
    opt_from.Caption = "起始板坯号"
    opt_to.Caption = "终止板坯号"
    
    txt_charge.Text = ""
    txt_from.Text = ""
    txt_to.Text = ""
    
    Timer1.Enabled = True
    
End Sub

Private Sub opt_unlock_Click(Value As Integer)

    If opt_unlock.Value = False Then Exit Sub
    
    Frame2.Enabled = True
    SSPanel2.Enabled = False
    
    opt_charge.Value = True
    opt_from.Value = True
    opt_to.Enabled = False
    
    txt_from.Left = 3180
    txt_to.Left = 5900
    
    opt_from.Caption = "起始炉号"
    opt_to.Caption = "终止炉号"
    
    txt_charge.Text = ""
    txt_from.Text = ""
    txt_to.Text = ""
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss1.MaxCols
        ss3.ColWidth(iCol) = ss1.ColWidth(iCol)
        ss5.ColWidth(iCol) = ss1.ColWidth(iCol)
    Next iCol
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_time_on Then Exit Sub
    
    With ss1
    
        .Row = Row
        
        .Col = 13  'L2
        If .Text <> "Y" Then Exit Sub
        
        .Col = 14  'Lock
        If opt_lock.Value Then
            
            .Col = 8   'Heat_No
            txt_charge.Text = .Text
        
        Else
        
            If .Text = "" Then Exit Sub
            
            If opt_charge.Value Then
            
                .Col = 8   'Heat_No
                
                If opt_from.Value Then
                    txt_from.Text = .Text
                Else
                    txt_to.Text = .Text
                End If
                
            End If
            
        End If
        
    End With
                
End Sub

Private Sub ss2_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss2.MaxCols
        ss4.ColWidth(iCol) = ss2.ColWidth(iCol)
        ss6.ColWidth(iCol) = ss2.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_time_on Then Exit Sub
    If opt_unlock And opt_charge Then Exit Sub
    
    With ss2
    
        .Row = Row
        
        .Col = 18  'L2
        
        If .Text <> "Y" Then Exit Sub
        
        .Col = 19  'Lock
        
        If opt_lock Then
        
            If .Text = "Y" Then Exit Sub
            
            .Col = 8   'Slab_No
            
            If opt_from Then
                txt_from.Text = .Text
                
                .Col = 20   'Slab_Edt_Seq
                P_Fr_Edt_Seq = .Value
                
            Else
                txt_to.Text = .Text
                
                .Col = 20   'Slab_Edt_Seq
                P_To_Edt_Seq = .Value
            End If
        
        Else
        
            If .Text = "" Then Exit Sub
            If opt_charge.Value Then Exit Sub
            
            .Col = 8   'Slab_No
            txt_from.Text = .Text
                
            .Col = 20   'Slab_Edt_Seq
            P_Fr_Edt_Seq = .Value
            
        End If
        
    End With
    
End Sub

Private Sub ss3_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss3.MaxCols
        ss1.ColWidth(iCol) = ss3.ColWidth(iCol)
        ss5.ColWidth(iCol) = ss3.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss3_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_time_on Then Exit Sub
    
    With ss3
    
        .Row = Row
        
        .Col = 13  'L2
        If .Text <> "Y" Then Exit Sub
        
        .Col = 14  'Lock
        If opt_lock.Value Then
            
            .Col = 8   'Heat_No
            txt_charge.Text = .Text
        
        Else
        
            If .Text = "" Then Exit Sub
            
            If opt_charge.Value Then
            
                .Col = 8   'Heat_No
                
                If opt_from.Value Then
                    txt_from.Text = .Text
                Else
                    txt_to.Text = .Text
                End If
                
            End If
            
        End If
    
    End With
    
End Sub

Private Sub ss4_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss4.MaxCols
        ss2.ColWidth(iCol) = ss4.ColWidth(iCol)
        ss6.ColWidth(iCol) = ss4.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss4_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_time_on Then Exit Sub
    If opt_unlock And opt_charge Then Exit Sub
    
    With ss4
    
        .Row = Row
    
        .Col = 18  'L2
        
        If .Text <> "Y" Then Exit Sub
        
        .Col = 19  'Lock
        
        If opt_lock Then
        
            If .Text = "Y" Then Exit Sub
            
            .Col = 8   'Slab_No
            
            If opt_from Then
                txt_from.Text = .Text
                
                .Col = 20   'Slab_Edt_Seq
                P_Fr_Edt_Seq = .Value
                
            Else
                txt_to.Text = .Text
                
                .Col = 20   'Slab_Edt_Seq
                P_To_Edt_Seq = .Value
            End If
        
        Else
        
            If .Text = "" Then Exit Sub
            If opt_charge.Value Then Exit Sub
            
            .Col = 8   'Slab_No
            txt_from.Text = .Text
                
            .Col = 20   'Slab_Edt_Seq
            P_Fr_Edt_Seq = .Value
            
        End If
        
    End With
    
End Sub

Private Sub ss5_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss5.MaxCols
        ss1.ColWidth(iCol) = ss5.ColWidth(iCol)
        ss3.ColWidth(iCol) = ss5.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss5_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_time_on Then Exit Sub
    
    With ss5
    
        .Row = Row
        
        .Col = 13  'L2
        If .Text <> "Y" Then Exit Sub
        
        .Col = 14  'Lock
        If opt_lock.Value Then
            
            .Col = 8   'Heat_No
            txt_charge.Text = .Text
        
        Else
        
            If .Text = "" Then Exit Sub
            
            If opt_charge.Value Then
            
                .Col = 8   'Heat_No
                
                If opt_from.Value Then
                    txt_from.Text = .Text
                Else
                    txt_to.Text = .Text
                End If
                
            End If
            
        End If
    
    End With
    
End Sub

Private Sub ss6_ColWidthChange(ByVal Col1 As Long, ByVal Col2 As Long)

    Dim iCol As Integer
    
    For iCol = 1 To ss6.MaxCols
        ss2.ColWidth(iCol) = ss6.ColWidth(iCol)
        ss4.ColWidth(iCol) = ss6.ColWidth(iCol)
    Next iCol

End Sub

Private Sub ss6_DblClick(ByVal Col As Long, ByVal Row As Long)

    If opt_time_on Then Exit Sub
    If opt_unlock And opt_charge Then Exit Sub
    
    With ss6
    
        .Row = Row
        
        .Col = 18  'L2
        
        If .Text <> "Y" Then Exit Sub
        
        .Col = 19  'Lock
        
        If opt_lock Then
        
            If .Text = "Y" Then Exit Sub
            
            .Col = 8   'Slab_No
            
            If opt_from Then
                txt_from.Text = .Text
                
                .Col = 20   'Slab_Edt_Seq
                P_Fr_Edt_Seq = .Value
                
            Else
                txt_to.Text = .Text
                
                .Col = 20   'Slab_Edt_Seq
                P_To_Edt_Seq = .Value
            End If
        
        Else
        
            If .Text = "" Then Exit Sub
            If opt_charge.Value Then Exit Sub
            
            .Col = 8   'Slab_No
            txt_from.Text = .Text
                
            .Col = 20   'Slab_Edt_Seq
            P_Fr_Edt_Seq = .Value
            
        End If
        
    End With
    
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss4_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss5_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss6_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss1
    
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss2
    
End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss3
    
End Sub

Private Sub ss4_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss4
    
End Sub

Private Sub ss5_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss5
    
End Sub

Private Sub ss6_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    Set Active_Spread = Me.ss6
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss3_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss4_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss5_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss6_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub Timer1_Timer()

    Call Form_Ref
    
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
        
        If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
            txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
        Else
            txt_plt_name.Text = ""
        End If
        
    End If

End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(9).Enabled = False                  'Row Cancel
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

Private Sub Spread_Color_Setting(oSpr As vaSpread)

    Dim iRow As Long
    Dim sPlan_Name As String
    Dim sAct_Stlgrd_Grp As String
    Dim sAct_Stlgrd As String
    
    With oSpr
    
        If oSpr.Name = "ss1" Or oSpr.Name = "ss3" Or oSpr.Name = "ss5" Then
    
            For iRow = 1 To .MaxRows
                
                .Row = iRow
                
                .Col = 7  'PLAN_NAME
                
                If iRow = 1 Then
                
                    sPlan_Name = .Text
                    
                    Call Gp_Sp_Bold(oSpr, "N", .Row)
                
                    .Col = 14  'L2-LCOK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                    Else
                        .Col = 13  'L2-CCM-SEND
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                        End If
                        
                        .Col = 1
                        If .Text <> "" Then
                            sAct_Stlgrd_Grp = .Text
                            .Col = 2
                            sAct_Stlgrd = .Text
                            
'                            .Col = 4
'                            If sAct_Stlgrd_Grp = "Z" Then
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
'                                Else
                                    .Col = 5
                                    If sAct_Stlgrd <> .Text Then
                                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
                                    End If
                               
'                                End If
'
'                            Else
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
'                                End If
'
'                            End If
                            
                        End If
                        
                    End If
                
                ElseIf sPlan_Name <> .Text Then
                    
                    sPlan_Name = .Text
                    
                    Call Gp_Sp_Bold(oSpr, "Y", .Row)
                    
                    .Col = 14  'L2-LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFC0C0)
                    Else
                        .Col = 13  'L2-CCM-SEND
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HC0FFFF)
                        Else
                            Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF)
                        End If
                        
                        .Col = 1
                        If .Text <> "" Then
                            sAct_Stlgrd_Grp = .Text
                            .Col = 2
                            sAct_Stlgrd = .Text
                            
'                            .Col = 4
'                            If sAct_Stlgrd_Grp = "Z" Then
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFFFC0)
'                                Else
                                    .Col = 5
                                    If sAct_Stlgrd <> .Text Then
                                        Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFFFC0)
                                    End If
                               
'                                End If
'
'                            Else
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFFFC0)
'                                End If
'
'                            End If
                            
                        End If
                    
                    End If
                    
                Else
                    
                    Call Gp_Sp_Bold(oSpr, "N", .Row)
                    
                    .Col = 14  'L2-LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                    Else
                        .Col = 13  'L2-CCM-SEND
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                        End If
                        
                        .Col = 1
                        If .Text <> "" Then
                            sAct_Stlgrd_Grp = .Text
                            .Col = 2
                            sAct_Stlgrd = .Text
                            
'                            .Col = 4
'                            If sAct_Stlgrd_Grp = "Z" Then
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
'                                Else
                                    .Col = 5
                                    If sAct_Stlgrd <> .Text Then
                                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
                                    End If
                               
'                                End If
'
'                            Else
'
'                                If sAct_Stlgrd_Grp <> .Text Then
'                                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFFFC0)
'                                End If
'
'                            End If
                            
                        End If
                        
                    End If
                
                End If
                
            Next iRow
          
        Else
        
            For iRow = 1 To .MaxRows
                
                .Row = iRow
                
                .Col = 6  'PLAN_NAME
                
                If iRow = 1 Then
                
                    sPlan_Name = .Text
                
                    Call Gp_Sp_Bold(oSpr, "N", iRow)
                    
                    .Col = 19  'LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                    Else
                        .Col = 18  'L2
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                        End If
                    End If
                
                ElseIf sPlan_Name <> .Text Then
                    
                    sPlan_Name = .Text
                
                    Call Gp_Sp_Bold(oSpr, "Y", .Row)
                    
                    .Col = 19  'LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFC0C0)
                    Else
                        .Col = 18  'L2
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HC0FFFF)
                        Else
                            Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF)
                        End If
                    End If
                    
                Else
                    
                    Call Gp_Sp_Bold(oSpr, "N", .Row)
                    
                    .Col = 19  'LOCK
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                    Else
                        .Col = 18  'L2
                        If .Text = "Y" Then
                            Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                        End If
                    End If
                
                End If
                
                .Row = iRow
                .Col = 21  'insert program-id
                
                If .Text <> "" Then
                    .Col = 8: .Col2 = 8
                    .Row = iRow: .Row2 = iRow
                    .BlockMode = True
                    .ForeColor = vbRed
                    .BlockMode = False
                End If
                
            Next iRow
        
        End If
        
        .RowHeight(-1) = 12.54
          
    End With
    
End Sub

Private Sub Gp_Sp_Scolor(sPname As Variant, sColType As String)

    With sPname
    
        .Row = 0: .Row2 = 0
        .Col = 0: .Col2 = 0
        
        .BlockMode = True
        
        .CellType = SS_CELL_TYPE_STATIC_TEXT
        .TypeHAlign = SS_CELL_H_ALIGN_CENTER
        .TypeVAlign = SS_CELL_V_ALIGN_CENTER
        .TypeTextWordWrap = True
        
        .BackColor = &HE1E4CD
        
        If sColType = "N" Then
            .ForeColor = vbRed
        Else
            .ForeColor = vbBlack
        End If
        
        .BlockMode = False
        
    End With
    
End Sub

Private Sub Gp_Sp_Bold(sPname As Variant, sType As String, iRow As Long)

    With sPname
    
        .Row = iRow: .Row2 = iRow
        .Col = 1: .Col2 = .MaxCols
        
        .BlockMode = True
        
        If sType = "N" Then
            .FontBold = False
        Else
            .FontBold = True
        End If
        
        .BlockMode = False
        
    End With
    
End Sub
