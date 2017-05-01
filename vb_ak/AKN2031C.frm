VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKN2031C 
   Caption         =   "工序流程变更"
   ClientHeight    =   8760
   ClientLeft      =   600
   ClientTop       =   2820
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10380
   Begin Threed.SSCommand cmd_ok 
      Height          =   465
      Left            =   9330
      TabIndex        =   23
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确定"
      BevelWidth      =   3
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   720
      Left            =   5220
      TabIndex        =   15
      Top             =   420
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   1270
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.TextBox txt_ccm_before 
         Alignment       =   2  'Center
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
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   " "
         Top             =   30
         Width           =   480
      End
      Begin VB.TextBox txt_ccm_after 
         Alignment       =   2  'Center
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
         Left            =   3465
         TabIndex        =   17
         Text            =   " "
         Top             =   30
         Width           =   480
      End
      Begin VB.TextBox txt_heat_mana_no 
         Alignment       =   2  'Center
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   " "
         Top             =   360
         Width           =   1365
      End
      Begin Threed.SSCheck chk_CCM_Change 
         Height          =   585
         Left            =   90
         TabIndex        =   19
         Top             =   90
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1032
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
         Caption         =   "CCM号机变更"
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   2895
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         Caption         =   "变更"
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
         Left            =   1545
         Top             =   30
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         Caption         =   "现在"
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
      Begin InDate.ULabel ULabel8 
         Height          =   300
         Left            =   1545
         Top             =   375
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   529
         Caption         =   "变更炉号"
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
         BackColor       =   &H00E0E0E0&
         Caption         =   "->"
         Height          =   210
         Left            =   2685
         TabIndex        =   20
         Top             =   75
         Width           =   255
      End
   End
   Begin VB.TextBox txt_prc_line 
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
      Left            =   8775
      MaxLength       =   2
      TabIndex        =   14
      Tag             =   "工厂"
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt_Stlgrd 
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
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   " "
      Top             =   60
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   9030
      MaxLength       =   2
      TabIndex        =   11
      Tag             =   "工厂"
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7530
      Left            =   75
      TabIndex        =   9
      Top             =   1170
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   13282
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AKN2031C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   4170
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   10260
         _Version        =   393216
         _ExtentX        =   18098
         _ExtentY        =   7355
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
         SpreadDesigner  =   "AKN2031C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3300
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4230
         Width           =   10260
         _Version        =   393216
         _ExtentX        =   18098
         _ExtentY        =   5821
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
         MaxCols         =   12
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2031C.frx":07CE
      End
   End
   Begin VB.TextBox txt_mlt_prc_cd 
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
      Left            =   6990
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   " "
      Top             =   45
      Width           =   1605
   End
   Begin VB.TextBox txt_heat_no_to 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
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
      Left            =   3465
      MaxLength       =   8
      TabIndex        =   7
      Text            =   " "
      Top             =   45
      Width           =   1020
   End
   Begin VB.TextBox txt_heat_no_fr 
      Alignment       =   2  'Center
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
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   " "
      Top             =   45
      Width           =   1020
   End
   Begin VB.TextBox txt_bof_proc 
      Alignment       =   2  'Center
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
      Height          =   345
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   " "
      Top             =   780
      Width           =   480
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   375
      Left            =   1665
      TabIndex        =   1
      Top             =   765
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   661
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin Threed.SSOption opt_lf 
         Height          =   225
         Index           =   2
         Left            =   720
         TabIndex        =   2
         Top             =   90
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
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
         Caption         =   "#2"
      End
      Begin Threed.SSOption opt_lf 
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   90
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
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
         Caption         =   "#1"
      End
      Begin Threed.SSOption opt_lf 
         Height          =   225
         Index           =   3
         Left            =   1320
         TabIndex        =   25
         Top             =   90
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   397
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
         Caption         =   "#3"
      End
   End
   Begin VB.TextBox txt_heat_no 
      Alignment       =   2  'Center
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
      Height          =   345
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   " "
      Top             =   780
      Width           =   1020
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   135
      Top             =   420
      Width           =   1020
      _ExtentX        =   1799
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   3675
      Top             =   420
      Width           =   1490
      _ExtentX        =   2619
      _ExtentY        =   556
      Caption         =   "VD/RH"
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   1665
      Top             =   420
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   556
      Caption         =   "LF"
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
   Begin Threed.SSFrame SSFrame2 
      Height          =   375
      Left            =   3675
      TabIndex        =   4
      Top             =   765
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin Threed.SSCheck chk_VD 
         Height          =   270
         Left            =   165
         TabIndex        =   21
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         Caption         =   "VD"
      End
      Begin Threed.SSCheck chk_RH 
         Height          =   270
         Left            =   855
         TabIndex        =   22
         Top             =   45
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   476
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         Caption         =   "RH"
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   1170
      Top             =   420
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   556
      Caption         =   "BOF"
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
      Left            =   135
      Top             =   45
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Caption         =   "起始炉号"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   2430
      Top             =   45
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Caption         =   "终止炉号"
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
      Left            =   5640
      Top             =   45
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "原工序流程"
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
   Begin Threed.SSCommand cmd_exit 
      Height          =   465
      Left            =   9330
      TabIndex        =   24
      Top             =   660
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "取消"
      BevelWidth      =   3
   End
End
Attribute VB_Name = "AKN2031C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name
'-- Program ID        AKN2031C
'-- Document No
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2006.1.10
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
Public sDateTime As String          'Active Form Authority Setting
Public sQuery_Rt As String          'Active Form Authority Setting

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

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
Dim Mc3 As New Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sProcCd   As String
Dim sCASStatus  As String
Dim sLFStatus   As String
Dim sVDStatus   As String
Dim sRHStatus   As String
Dim sCCMStatus  As String

Private Sub Form_Define()
     
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(txt_ccm_after, "p", "n", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(txt_Stlgrd, "p", "n", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(txt_heat_no_fr, "p", "n", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_heat_mana_no, "p", " ", " ", " ", "r", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
       
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AKN2031C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
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
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AKN2031C.P_REFER2", Key:="P-R"
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
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss2, 1, True)     'SEQ_NO

End Sub
 
Private Sub chk_CCM_Change_Click(Value As Integer)

    txt_heat_no_to.Text = txt_heat_no_fr.Text
    
    If Value = -1 Then
        txt_heat_no_to.Locked = True
        
        If txt_ccm_before.Text = "1" Then
            txt_ccm_after.Text = "2"
        ElseIf txt_ccm_before.Text = "2" Then
            txt_ccm_after.Text = "3"
        Else
            txt_ccm_after.Text = "1"
        End If
        
        sCCMStatus = "BF" & txt_ccm_after.Text
        txt_heat_no_to.Locked = True
'        Call Form_Ref

    Else
        txt_heat_no_to.Locked = False
        txt_ccm_after.Text = ""
        sCCMStatus = "BF" & txt_ccm_before.Text
        txt_heat_no_to.Locked = False
        Call Gf_Sp_Cls(sc1)
        Call Gf_Sp_Cls(sc2)
    End If
End Sub

Private Sub chk_RH_Click(Value As Integer)

    If chk_VD = True And chk_RH = True Then
       Call MsgBox("VD & RH 一起不能做...！", vbInformation, "系统提示信息")
       chk_RH.Value = af_CHK_RH
    End If
    
End Sub

Private Sub chk_VD_Click(Value As Integer)

    If chk_VD = True And chk_RH = True Then
       Call MsgBox("VD & RH 一起不能做...！", vbInformation, "系统提示信息")
       chk_VD.Value = af_CHK_VD
    End If
    
End Sub

Private Sub cmd_exit_Click()
   Call Form_Exit
End Sub

Private Sub Form_Load()
    
    Dim sMltProc    As String
    Dim sPrcLine    As String
    Dim iLoc        As Integer
    
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Gp_FormCenter(Me)
    
    Call Form_Define
  
    Screen.MousePointer = vbDefault
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "K-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "K-System.INI", Me.Name)
    
    Chg_Lf = opt_lf(1).Value
    
    Chg_VD = chk_VD.Value
    Chg_RH = chk_RH.Value
    bf_CHK_VD = chk_VD.Value
    bf_CHK_RH = chk_RH.Value

    sProcCd = AKN2030C.txt_proc_fl.Text
    
    iLoc = 0
    If txt_plt.Text = "" Then txt_plt.Text = "B1"
    
    With AKN2030C.ss1
    
        .Row = .ActiveRow
    
        .Col = 7:     sMltProc = .Text
        
        iLoc = InStr(1, sMltProc, "BD")
        sPrcLine = Trim(Mid(sMltProc, iLoc + 2, 1))
        If iLoc > 0 And IsNumeric(sPrcLine) Then
            opt_lf(sPrcLine).Value = True
        End If
         
        iLoc = InStr(1, sMltProc, "BE")
        If iLoc > 0 Then
            chk_VD.Value = True
        End If
        
        iLoc = InStr(1, sMltProc, "BH")
        If iLoc > 0 Then
            chk_RH.Value = True
        End If
        
        iLoc = InStr(1, sMltProc, "BF")
        sPrcLine = Trim(Mid(sMltProc, iLoc + 2, 1))
        If iLoc > 0 And IsNumeric(sPrcLine) Then
            txt_ccm_before.Text = sPrcLine
            sCCMStatus = "BF" & sPrcLine
        Else
            sCCMStatus = "BF" & txt_bof_proc.Text
        End If
        
    End With
    
    'This step makes that you can't change the prc_line_code of passed line.
    If sProcCd = "B" Then
        Call ProdeuctResult_Search
    End If
    
    bf_OPT_LF1 = opt_lf(1).Value
    af_OPT_LF1 = opt_lf(1).Value
    bf_CHK_VD = chk_VD.Value
    af_CHK_VD = chk_VD.Value
    bf_CHK_RH = chk_RH.Value
    af_CHK_RH = chk_RH.Value
    
End Sub

Public Sub ProdeuctResult_Search()

    Dim sQuery      As String
    
    Set AdoRs = New ADODB.Recordset
        
    sQuery = "         SELECT MAX(LF), MAX(VD), MAX(RH)    " & vbCrLf
    sQuery = sQuery & "FROM (                              " & vbCrLf
    sQuery = sQuery & "SELECT DECODE(PRC,'BD',PRC_LINE,'0') LF  " & vbCrLf
    sQuery = sQuery & "      ,DECODE(PRC,'BE',PRC_LINE,'0') VD  " & vbCrLf
    sQuery = sQuery & "      ,DECODE(PRC,'BH',PRC_LINE,'0') RH  " & vbCrLf
    sQuery = sQuery & "  FROM FP_MSPSTATUS        " & vbCrLf
    sQuery = sQuery & " WHERE HEAT_NO  = '" & txt_heat_no.Text & "' " & vbCrLf
    sQuery = sQuery & " GROUP BY HEAT_NO, PRC, PRC_LINE )"
    
    AdoRs.Open sQuery, M_CN1, adOpenForwardOnly, adLockReadOnly
                   
    If sLFStatus = "1" Or sLFStatus = "2" Or sLFStatus = "3" Then
        opt_lf(1).Enabled = False
        opt_lf(2).Enabled = False
        opt_lf(3).Enabled = False
        sLFStatus = "BD" & sLFStatus
    End If
     
    If sVDStatus = "1" Or sVDStatus = "2" Then
        chk_VD.Enabled = False
        sVDStatus = "BE1"
    End If
     
    If sRHStatus = "1" Or sRHStatus = "2" Then
        chk_RH.Enabled = False
        sRHStatus = "BH2"
    End If
    
    AdoRs.Close
End Sub

Public Sub Form_Ref()
    
    Call Gf_Sp_Cls(sc2)
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Sp_EvenRowBackcolor(ss1)
        
'        ss1.OperationMode = OperationModeNormal
    End If
            
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc2) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gf_Sp_Cls(sc1)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
    End If

End Sub

Private Sub Cmd_Ok_Click()
  
    Dim sMltProc  As String
    Dim iRow      As Integer
    Dim iEndRow   As Integer
    Dim iMaxRow   As Integer
    Dim iForStep  As Integer
    
    iEndRow = 0
    iForStep = 1
    
    If chk_VD = True And chk_RH = True Then
       Call MsgBox("VD & RH 一起不能做...！", vbInformation, "系统提示信息")
       chk_RH.Value = af_CHK_RH
       Exit Sub
    End If
    
    af_OPT_LF1 = opt_lf(1).Value
    af_CHK_VD = chk_VD.Value
    af_CHK_RH = chk_RH.Value
                
    If txt_heat_no_fr.Text <> txt_heat_no_to.Text Then
    
        If Len(txt_heat_no_to) <> 8 Or Mid(txt_heat_no_fr.Text, 3, 1) <> Mid(txt_heat_no_to.Text, 3, 1) Then
            Call MsgBox("终止炉号错了..！", vbInformation, "系统提示信息")
            Exit Sub
        End If
        
        With AKN2030C.ss1
            If txt_heat_no_fr.Text < txt_heat_no_to.Text Then
                iMaxRow = .MaxRows
                iForStep = 1
            Else
                iMaxRow = 1
                iForStep = -1
            End If
            
            For iRow = .ActiveRow To iMaxRow Step iForStep
                .Row = iRow
                .Col = 1
                If .Text = txt_heat_no_to.Text Then
                    .Col = 17
                    If sProcCd = .Text Then
                        iEndRow = iRow
                    End If
                    Exit For
                End If
            Next iRow
        End With
        
        If iEndRow = 0 Then
            Call MsgBox("终止炉号错了..！", vbInformation, "系统提示信息")
            Exit Sub
        End If
    Else
        iEndRow = AKN2030C.ss1.ActiveRow
    End If
    
    If Len(sLFStatus) <> 3 Then
        Call MsgBox("LF工程错了..！", vbInformation, "系统提示信息")
        Exit Sub
    End If
        
    If Len(sCCMStatus) <> 3 Then
        Call MsgBox("CCM工程错了..！", vbInformation, "系统提示信息")
        Exit Sub
    End If
    
    If chk_CCM_Change.Value = -1 And Trim(txt_heat_mana_no.Text) = "" Then
        Call MsgBox("变更炉号错了..！", vbInformation, "系统提示信息")
        Exit Sub
    End If
    
    If Not Gf_MessConfirm("确定要变更作业指示的工序流程吗？", "W", "系统提示信息确认") Then
       Exit Sub
    End If
    
    sCASStatus = InStr(txt_mlt_prc_cd, "G")
    If sCASStatus <> "0" Then
       sCASStatus = "BG1"
    Else
       sCASStatus = ""
    End If
    
    If AKN2031C.chk_VD = True Then
       sVDStatus = "BE1"
    Else
       sVDStatus = ""
    End If
    
    If AKN2031C.chk_RH = True Then
       sRHStatus = "BH2"
    Else
       sRHStatus = ""
    End If
     
    sMltProc = "BC" & Trim(txt_bof_proc.Text) & sCASStatus & sLFStatus & sRHStatus & sVDStatus & sCCMStatus
    
    'If Mid(txt_heat_no_fr.Text, 3, 1) = "1" Then
        With AKN2030C.ss1
            For iRow = .ActiveRow To iEndRow Step iForStep
                .Row = iRow
                .Col = 7
                If .Text <> sMltProc Then
                    .Text = sMltProc
                    .Col = 0:      .Text = "Update"
                End If
            Next iRow
        End With
    'End If
        
    Call AKN2030C.Form_Pro
    
    Unload Me
    
    MDIMain.MenuTool.Buttons(9).Enabled = True
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "K-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "K-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
    
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
    Set Mc3 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
End Sub

Private Sub opt_lf_Click(Index As Integer, Value As Integer)

    sLFStatus = "BD" & Index
    
End Sub

Private Sub opt_vdrh_Click(Index As Integer, Value As Integer)

    If Index = 1 Then
        sVDStatus = "BE1"
        sRHStatus = ""
    Else
        sRHStatus = "BH2"
        sVDStatus = ""
    End If
    
End Sub

Private Sub opt_ccm_Click(Index As Integer, Value As Integer)

    sCCMStatus = "BF" & Index
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor               As String
    
    If Row < 1 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 1
    txt_heat_mana_no.Text = ss1.Text

    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub

    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    ss2.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss2)

    With ss1
    
        For iRow1 = .ActiveRow To .MaxRows
        
            .Col = 1
            .Row = iRow1
            sColor = .BackColor
            sHeat = .Text
            
            With ss2
              .Col = 3
              For iRow2 = 1 To .MaxRows
                  .Row = iRow2
                  
                   If Left(.Text, 8) = sHeat Then
                      For iCol = 1 To .MaxCols
                          .Col = iCol
                          .BackColor = sColor
                      Next iCol
                      sTemp = .Text
                   End If
    
                   If sTemp <> "" And sTemp <> Left(.Text, 8) Then
                      sTemp = ""
                      Exit For
                   End If
                
                  .Col = 3
                  
              Next iRow2
              
            End With

        Next iRow1
        
    End With
    
End Sub

Private Sub txt_ccm_after_Change()

    If chk_CCM_Change.Value = ssCBUnchecked Then
        If txt_ccm_before.Text = "" Then txt_ccm_before.Text = txt_bof_proc.Text
        sCCMStatus = "BF" & txt_ccm_before.Text
        Exit Sub
    End If
    
    If Trim(txt_ccm_after.Text) <> "" And txt_ccm_before.Text = txt_ccm_after.Text Then
        txt_ccm_after.Text = ""
        Call MsgBox("CCM工程错了..！", vbInformation, "系统提示信息")
        Exit Sub
    End If
    
    sCCMStatus = "BF" & txt_ccm_after.Text
    Call Form_Ref
    
End Sub

