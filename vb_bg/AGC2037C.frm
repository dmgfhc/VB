VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2037C 
   Caption         =   "钢板剪切实绩查询及修改界面_AGC2037C"
   ClientHeight    =   9225
   ClientLeft      =   810
   ClientTop       =   2445
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
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
      ItemData        =   "AGC2037C.frx":0000
      Left            =   14010
      List            =   "AGC2037C.frx":000D
      TabIndex        =   21
      Tag             =   "班次"
      Top             =   90
      Width           =   735
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
      Left            =   11430
      MaxLength       =   1
      TabIndex        =   15
      Text            =   " "
      Top             =   450
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox cbo_sUserID 
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
      ItemData        =   "AGC2037C.frx":001A
      Left            =   11580
      List            =   "AGC2037C.frx":002A
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txt_stdspec 
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
      Left            =   5340
      TabIndex        =   6
      Top             =   480
      Width           =   2925
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
      Left            =   8280
      TabIndex        =   5
      Top             =   480
      Width           =   2445
   End
   Begin VB.TextBox txt_plt 
      CausesValidation=   0   'False
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
      Left            =   11940
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "生产工厂"
      Top             =   450
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txt_mplate_no 
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
      Left            =   1770
      MaxLength       =   14
      TabIndex        =   2
      Tag             =   "物料号"
      Top             =   480
      Width           =   1755
   End
   Begin VB.ComboBox cbo_group 
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
      ItemData        =   "AGC2037C.frx":0052
      Left            =   10890
      List            =   "AGC2037C.frx":0062
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   645
   End
   Begin InDate.UDate udt_date_fr 
      Height          =   315
      Left            =   9270
      TabIndex        =   7
      Tag             =   "INS_DATE"
      Top             =   90
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
      Left            =   10740
      TabIndex        =   8
      Tag             =   "INS_DATE"
      Top             =   90
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
   Begin InDate.ULabel ULabel20 
      Height          =   315
      Left            =   240
      Top             =   480
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "母板号"
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
      ForeColor       =   0
   End
   Begin SSSplitter.SSSplitter SSSp1 
      Height          =   8325
      Left            =   60
      TabIndex        =   9
      Top             =   840
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14684
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AGC2037C.frx":0076
      Begin FPSpread.vaSpread ss2 
         Height          =   4110
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   7250
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
         MaxCols         =   18
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AGC2037C.frx":00C8
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   4155
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   4170
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   7329
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ButtonDrawMode  =   4
         ColsFrozen      =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   38
         MaxRows         =   2
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AGC2037C.frx":0BEA
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   7740
      Top             =   90
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   3810
      Top             =   480
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "标准号 / 改判"
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
   Begin CSTextLibCtl.sitxEdit txt_cut_time 
      Height          =   315
      Left            =   13050
      TabIndex        =   10
      Tag             =   "出炉时间"
      Top             =   450
      Visible         =   0   'False
      Width           =   2130
      _Version        =   262145
      _ExtentX        =   3757
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "____-__-__ __:__:__"
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   16
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   "____-__-__ __:__:__"
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
      ValidateMask    =   0   'False
   End
   Begin VB.TextBox txt_rec_sts 
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
      Left            =   12150
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   450
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txt_line 
      Alignment       =   2  'Center
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
      Left            =   12555
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "2"
      Top             =   450
      Visible         =   0   'False
      Width           =   480
   End
   Begin Threed.SSOption opt_on 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   120
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
      Left            =   2610
      TabIndex        =   17
      Top             =   120
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
   Begin Threed.SSPanel SSPanel5 
      Height          =   315
      Left            =   5370
      TabIndex        =   18
      Top             =   90
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_line1 
         Height          =   285
         Left            =   90
         TabIndex        =   19
         Top             =   30
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   0
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
         Caption         =   "一号"
      End
      Begin Threed.SSOption opt_line2 
         Height          =   285
         Left            =   1020
         TabIndex        =   20
         Top             =   30
         Width           =   735
         _ExtentX        =   1296
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
         Caption         =   "二号"
      End
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   3810
      Top             =   90
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "剪切线"
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
   Begin VB.TextBox txt_plate_no 
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
      Left            =   15330
      MaxLength       =   14
      TabIndex        =   14
      Tag             =   "物料号"
      Top             =   1230
      Visible         =   0   'False
      Width           =   1755
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   240
      Top             =   90
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.ULabel ULabel31 
      Height          =   315
      Left            =   12660
      Top             =   90
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
Attribute VB_Name = "AGC2037C"
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
'-- Program Name      钢板剪切实绩查询及修改界面
'-- Program ID        AGC2037C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2008.3.24
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

Dim pControl2 As New Collection     'Master Primary Key Collection
Dim nControl2 As New Collection     'Master Necessary Collection
Dim mControl2 As New Collection     'Master Maxlength check Collection
Dim iControl2 As New Collection     'Master Insert Collection
Dim rControl2 As New Collection     'Master Refer Collection
Dim cControl2 As New Collection     'Master Copy Collection
Dim aControl2 As New Collection     'Master -> Spread Collection
Dim lControl2 As New Collection     'Master Lock Collection

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
Dim lMain_Row As Integer

Const SPD_LINE1 = 1
Const SPD_LINE2 = 2
Const SPD_PLATE_NO = 3
Const SPD_THK = 11
Const SPD_WID = 12
Const SPD_LEN = 13
Const SPD_WGT = 14
Const SPD_LAST_YN = 15
Const SPD_SIZE_KND = 16
Const SPD_TRIM_FL = 17
Const SPD_APLY_STDSPEC = 18
Const SPD_APLY_STDSPEC_NEW = 19
Const SPD_SURF_GRD = 20
Const SPD_MARK_YN = 23
Const SPD_STAMP_YN = 24
Const SPD_BAR_YN = 25
Const SPD_PROD_DATE = 26
Const SPD_EMP_CD = 27
Const SPD_PAINT = 28
Const SPD_LABEL = 29
Const SPD_STDSPEC_YY = 30
Const SPD_STLGRD = 31
Const SPD_REC_STS = 32

Const SS2_PLATE_NO = 1
Const SS2_URGNT_FL = 18


Private Sub Form_Define()
        
    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
       
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_mplate_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_onoff, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(udt_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(udt_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_stdspec, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_rec_sts, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_line, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(TXT_PLATE_NO, "p", "n", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
         
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2037C.P_SREFER2", Key:="P-R"
    sc1.Add Item:="AGC2037C.P_ONEROW2", Key:="P-O"
    sc1.Add Item:="AGC2037C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    For iCol = 2 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGC2037C.P_SREFER1", Key:="P-R"
    sc2.Add Item:="AGC2037C.P_ONEROW1", Key:="P-O"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
'    Call Gp_Sp_ColHidden(ss1, SPD_LINE1, True)
    Call Gp_Sp_ColHidden(ss1, SPD_PAINT, True)
    Call Gp_Sp_ColHidden(ss1, SPD_LABEL, True)
    Call Gp_Sp_ColHidden(ss1, SPD_MARK_YN, True)
    Call Gp_Sp_ColHidden(ss1, SPD_STAMP_YN, True)
    Call Gp_Sp_ColHidden(ss1, SPD_BAR_YN, True)
'    Call Gp_Sp_ColHidden(ss1, SPD_STDSPEC_YY, True)
'    Call Gp_Sp_ColHidden(ss1, SPD_STLGRD, True)
    Call Gp_Sp_ColHidden(ss1, SPD_REC_STS, True)

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
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(ss2)
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    txt_plt.Text = "C1"
    CBO_sUserID.Text = sUserID
    
    txt_onoff.Text = "I"
    txt_rec_sts.Text = "1"
    lMain_Row = 0
    opt_line2.Value = True
    txt_line.Text = "2"
    
    Call Gp_Spl_SizeGet(SSSp1, "G-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Call Gp_Spl_SizeSet(SSSp1, "G-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
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

Public Sub Form_Cls()

    If Gf_Sp_Cls(sc1) Then
    
        Call Gf_Sp_Cls(sc2)
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        
        txt_plt.Text = "C1"
        opt_on.Value = True
        txt_rec_sts.Text = "1"
        opt_line2.Value = True
        txt_line.Text = "2"
        txt_stdspec_chg = ""
        
        lMain_Row = 0
        
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, sc1.Item("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim iCount As Integer
    
    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
            
    Call Gf_Sp_Cls(sc1)
    
    If Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss2.OperationMode = OperationModeNormal
        Call Gp_Sp_EvenRowBackcolor(ss2)
        Call Gp_Ms_Cls(Mc2("rControl"))
        lMain_Row = 0
    End If
    
    '是否紧急订单
    With ss2
        If .MaxRows <= 1 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
             .Row = iCount
             ss2.Row = .Row:       ss2.Col = SS2_URGNT_FL
           If ss2.Text = "Y" Then
                Call Gp_Sp_BlockColor(ss2, SS2_PLATE_NO, SS2_PLATE_NO, .Row, .Row, &HC000&)
                Call Gp_Sp_BlockColor(ss2, SS2_URGNT_FL, SS2_URGNT_FL, .Row, .Row, &HC000&)
           End If
        Next iCount
    End With
        
End Sub

Public Sub Form_Pro()

    Dim iRow As Integer
    Dim sQuery As String
    
    If txt_rec_sts = "1" Then
    
        If Gf_Sp_Process(M_CN1, sc1, Mc2) Then
        
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
            
            With ss1
            
                For iRow = 1 To .MaxRows
                
                    .Row = iRow
                    .Col = SPD_REC_STS
                    
                    If .Text = "1" Then      'EXIST GP_PLATETRK, REC_STS = '1' IN GP_PLATE
                    
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HC0FFFF)
                        
                    ElseIf .Text = "2" Then  'EXIST GP_PLATETRK, REC_STS = '2' IN GP_PLATE
                    
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &H80000005)
                        Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, iRow, iRow, True)
                        
                    ElseIf .Text = "3" Then  'EXIST GP_PLATETRK, REC_STS = '3' IN GP_PLATE
                
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &H80000005)
                        Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, iRow, iRow, True)
                        
                    Else                     'NOT EXIST GP_PLATETRK, REC_STS = '1' IN GP_PLATE
                    
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HC0FFC0)
                        
                    End If
                
                Next iRow
                
            End With
                    
        End If
        
'        Call Form_Ref
        
        'SS2 RE-DISPLAY
        If lMain_Row <> 0 Then
            sQuery = Gf_Sp_MakeQuery(ss2, sc2.Item("P-O"), "O", sc2.Item("pColumn"), lMain_Row)
            Call Gp_Sp_OneRowDisplay(M_CN1, sQuery, ss2, lMain_Row)
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
    Dim sLotNo      As String
    Dim sCutNo      As String
    Dim sClipText   As String
    
    Dim sSize_knd   As Integer
    Dim sTrim_fl    As Integer
    Dim sAply_stdspec  As String
    Dim sEmp_Cd     As String
    Dim sStdspec_yy As String
    Dim sStdspec    As String
    
    Dim iCount As Integer
    
    sPlateNo = ""
    
    With ss1
    
        If .MaxRows = 0 Then
        
           If Len(TXT_PLATE_NO.Text) = 12 Then
               Call Gp_Sp_Ins(Proc_Sc("Sc1"))
              .Row = 1
              .Col = SPD_PLATE_NO
              .Text = TXT_PLATE_NO.Text & "01"
              .Col = SPD_THK:           .Value = 0
              .Col = SPD_WID:           .Value = 0
              .Col = SPD_LEN:           .Value = 0
              .Col = SPD_APLY_STDSPEC:  .Text = "GB-XXX"
           Else
               Call Gp_MsgBoxDisplay("请正确输入母板号 ！")
           End If
           
           Exit Sub
           
        End If
        
        For iCount = .ActiveRow To .MaxRows
            .Row = iCount
            .Col = SPD_PLATE_NO
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
    Call Gp_Sp_Ins(Proc_Sc("Sc1"))

    With ss1
        .ReDraw = False
        If lRow > 0 Then
            .Row = lRow
            .Col = SPD_PLATE_NO:      sPlateNo = .Text
            .Col = SPD_THK:           dThk = Val(.Value) 'Val(.Text & "")
            .Col = SPD_WID:           dWid = Val(.Value) 'Val(.Text & "")
            .Col = SPD_LEN:           dLen = Val(.Value) 'Val(.Text & "")
            .Col = SPD_WGT:           dWgt = Val(.Value) 'Val(.Text & "")
            .Col = SPD_SIZE_KND:      sSize_knd = .Value
            .Col = SPD_TRIM_FL:       sTrim_fl = .Value
            .Col = SPD_APLY_STDSPEC:  sAply_stdspec = .Text
            .Col = SPD_STDSPEC_YY:    sStdspec_yy = .Text
            .Col = SPD_EMP_CD:        sEmp_Cd = .Text
            .Col = SPD_STLGRD:        sStdspec = .Text
        Else
            sPlateNo = TXT_PLATE_NO.Text & "00"
        End If

        .Row = lRow + 1
        .Col = SPD_PLATE_NO:      .Text = sPlateNo
        .Col = SPD_THK:           .Value = dThk
        .Col = SPD_WID:           .Value = dWid
        .Col = SPD_LEN:           .Value = dLen
        .Col = SPD_WGT:           .Value = dWgt
        .Col = SPD_SIZE_KND:      .Value = sSize_knd
        .Col = SPD_TRIM_FL:       .Value = sTrim_fl
        .Col = SPD_APLY_STDSPEC:  .Text = sAply_stdspec
        .Col = SPD_EMP_CD:        .Text = sEmp_Cd
        .Col = SPD_STDSPEC_YY:    .Text = sStdspec_yy
        .Col = SPD_STLGRD:        .Text = sStdspec
        .Col = 0:                 .Text = "Input"
        .Col = SPD_PLATE_NO: .Text = Mid(.Text, 1, 12) & Format(Val(Mid(.Text, 13, 2) & "") + 1, "00")
        .Col = SPD_SURF_GRD:      .Value = 1
        .Col = SPD_MARK_YN:       .Value = 1
        .Col = SPD_STAMP_YN:      .Value = 1
        .Col = SPD_BAR_YN:        .Value = 1
         If opt_line1.Value = True Then
        .Col = SPD_LINE1:         .Value = 1
         Else
        .Col = SPD_LINE2:         .Value = 1
         End If
        .Col = 0:                 .Text = "Input"
        
         Call .SetActiveCell(1, .Row)
        .ReDraw = True
    End With

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

Public Sub Spread_Del()
    ss1.Row = ss1.ActiveRow:        ss1.Col = SPD_EMP_CD:        ss1.Text = sUserID
    Call Gp_Sp_Del(Proc_Sc("Sc1"))
End Sub

Public Sub Spread_Can()

    Dim iCount As Integer
    Dim sPlateNo As String
    
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc1"))
    
    With ss1
                
        For iCount = 1 To .MaxRows
            
            .Row = iCount
            .Col = SPD_PLATE_NO
             sPlateNo = .Text
            
            If Left(.Text, 12) = Left(sPlateNo, 12) Then
            Else
               .Row = iCount - 1
               .Col = SPD_LAST_YN
               .Value = 1
            End If
            
            .Col = .MaxCols
            
            If .Text = "1" Then      'EXIST GP_PLATETRK, REC_STS = '1' IN GP_PLATE
            
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCount, iCount, , &HC0FFFF)
                
            ElseIf .Text = "2" Then  'EXIST GP_PLATETRK, REC_STS = '2' IN GP_PLATE
            
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCount, iCount, , &H80000005)
                Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, iCount, iCount, True)
                
            ElseIf .Text = "3" Then  'EXIST GP_PLATETRK, REC_STS = '3' IN GP_PLATE
        
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCount, iCount, , &H80000005)
                Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, iCount, iCount, True)
                
            Else                     'NOT EXIST GP_PLATETRK, REC_STS = '1' IN GP_PLATE
            
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCount, iCount, , &HC0FFC0)
                
            End If
        
        Next iCount
    
    End With
            
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_line1_Click(Value As Integer)

    If opt_line1.Value Then
        opt_line1.ForeColor = &HFF&
        opt_line2.ForeColor = &H80000012
        txt_line.Text = "1"
        Call Gp_Sp_ColHidden(ss1, SPD_LINE1, False)
        Call Gp_Sp_ColHidden(ss1, SPD_LINE2, True)
        ss1.MaxRows = 0
        ss2.MaxRows = 0
    End If

End Sub

Private Sub opt_line2_Click(Value As Integer)

    If opt_line2.Value Then
        opt_line2.ForeColor = &HFF&
        opt_line1.ForeColor = &H80000012
        txt_line.Text = "2"
        Call Gp_Sp_ColHidden(ss1, SPD_LINE1, True)
        Call Gp_Sp_ColHidden(ss1, SPD_LINE2, False)
        ss1.MaxRows = 0
        ss2.MaxRows = 0
    End If

End Sub

Private Sub opt_off_Click(Value As Integer)

    If opt_off.Value Then
        opt_on.ForeColor = &H80000012
        opt_off.ForeColor = &HFF&
        txt_onoff.Text = "O"
    End If

End Sub

Private Sub opt_on_Click(Value As Integer)

    If opt_on.Value Then
        opt_off.ForeColor = &H80000012
        opt_on.ForeColor = &HFF&
        txt_onoff.Text = "I"
    End If

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    If Row <= 0 Then Exit Sub
  
    'Call Cmd_SEND_SET(Row)
    ss1.Row = Row
     
    If Col = SPD_APLY_STDSPEC_NEW Then
    
        ss1.Col = SPD_REC_STS
        If ss1.Text <> "3" Then
            
            ss1.Col = Col
            If ss1.Text = "" Then
               ss1.Text = txt_stdspec_chg
               If txt_stdspec_chg <> "" Then
                   ss1.Col = SPD_SURF_GRD
                   ss1.Value = 0
               End If
            Else
               ss1.Text = ""
               ss1.Col = SPD_SURF_GRD
               ss1.Value = 1
            End If
        
        End If
        
    End If
    
    If Col = SPD_PROD_DATE Then
        
        ss1.Col = SPD_REC_STS
        If ss1.Text = "1" Or ss1.Text = "4" Or ss1.Text = "" Then
            TXT_CUT_TIME.RawData = Gf_DTSet(M_CN1, , "X")
            ss1.Col = SPD_PROD_DATE
            ss1.Text = TXT_CUT_TIME.Text
        End If
        
    End If
  
End Sub

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)

    Dim iCount As Integer
    Dim sPlateNo As String
    
    If Gf_Sp_Cls(sc1) Then
        
        If lMain_Row <> 0 Then
    
            ss2.Row = lMain_Row
            ss2.Col = 0
            ss2.Text = ""
            
            If lMain_Row Mod 2 <> 0 Then
                Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, lMain_Row, lMain_Row, , &HF2F2F2)
            Else
                Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, lMain_Row, lMain_Row, , &HFFFFFF)
            End If
            
        End If
    
        ss2.Row = Row
        ss2.Col = 0
        
        If ss2.Text = "" Then
            ss2.Text = "选择"
            Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, Row, Row, , CYAN)
            lMain_Row = Row
        End If
        
        ss2.Row = Row
        ss2.Col = 1
        TXT_PLATE_NO.Text = ss2.Text
        
        If Gf_Sp_Refer(M_CN1, sc1, Mc2, Mc2("nControl"), Mc2("mControl"), False) Then
        
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call MenuTool_ReSet
            ss1.OperationMode = OperationModeNormal
            
            With ss1
                
                For iCount = 1 To .MaxRows
                    
                    .Row = iCount
                    .Col = SPD_PLATE_NO
                     sPlateNo = .Text
                    
                    If Left(.Text, 12) = Left(sPlateNo, 12) Then
                    Else
                       .Row = iCount - 1
                       .Col = SPD_LAST_YN
                       .Value = 1
                    End If
                    
                    .Col = SPD_REC_STS
                    
                    If .Text = "1" Then      'EXIST GP_PLATETRK, REC_STS = '1' IN GP_PLATE
                    
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCount, iCount, , &HC0FFFF)
                        
                    ElseIf .Text = "2" Then  'EXIST GP_PLATETRK, REC_STS = '2' IN GP_PLATE
                    
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCount, iCount, , &H80000005)
                        Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, iCount, iCount, True)
                        
                    ElseIf .Text = "3" Then  'EXIST GP_PLATETRK, REC_STS = '3' IN GP_PLATE
                
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCount, iCount, , &H80000005)
                        Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, iCount, iCount, True)
                        
                    Else                     'NOT EXIST GP_PLATETRK, REC_STS = '1' IN GP_PLATE
                    
                        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iCount, iCount, , &HC0FFC0)
                        
                    End If
                
                Next iCount
            
            End With
    
        End If
    
    End If

End Sub

Private Sub txt_mplate_no_Change()
    If ss1.MaxRows < 1 And ss2.MaxRows < 1 Then
        TXT_PLATE_NO.Text = txt_mplate_no.Text
    End If
End Sub

Private Sub txt_stdspec_chg_DblClick()

    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim iCol As Long
    Dim iRow As Long
    Dim iMode As Integer
    
    Dim iRowNum As Long
    Dim iRowfr As Long
    Dim iRowto As Long
    
    iCol = Col
    iRow = Row
    iMode = Mode

    If Row <= 0 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") And Col >= SPD_LINE2 Then
    
         iRowto = iRow - 1
         iRowfr = iRow + 1
        
        If Col = SPD_THK Or Col = SPD_WID Or Col = SPD_LEN Then
            If Mode = 1 Then
               ss1.Col = iCol
               ss1.Row = iRow
               'ss1.Text = 0
            End If
        End If
    
        Call Gp_Sp_UpdateMake(Proc_Sc("SC1")("Spread"), iMode)
        
        ss1.Row = iRow  'ss1.ActiveRow
        ss1.Col = SPD_EMP_CD
        ss1.Text = CBO_sUserID.Text
        
    End If

End Sub

Private Sub ss1_LostFocus()
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
    
    If ss1.MaxRows > 0 Then
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

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    Else

    End If
    
End Sub

Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyF4 Then
  
         DD.sWitch = "MS"
         DD.DataDicType = "C"
         DD.rControl.Add Item:=txt_stdspec_chg
        
         Call Pf_Common_DD(M_CN1, KeyCode)
         
         Exit Sub
    End If

End Sub

Private Sub txt_stdspec_DblClick()

    Call txt_STDSPEC_KeyUp(vbKeyF4, 0)
    
End Sub

Private Function Pf_Common_DD(Conn As ADODB.Connection, KeyCode As Integer) As Boolean

    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sSelect = False
        DD.sWhere = ""
        DD.sKey = ""
        
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "HC"        'Common Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    DD.sQuery = "SELECT CD_SHORT_NAME ""标准代号"", CD_NAME ""标准中文名"" FROM ZP_CD WHERE CD_MANA_NO = 'G0030'"
    
    Call Gf_DD_Display(Conn, DD.sQuery, False)
    
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function

Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        txt_stdspec.Text = ""
        DD.rControl.Add Item:=txt_stdspec

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)

        Exit Sub

    End If
    
End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
'        .Buttons(7).Enabled = False                  'Row Insert
'        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_USER_ComboAdd
'   2.Name         :
'   3.Input  Value : Conn Connection, Cbo Variant,sPRC String,
'                    {sFACT_CD,sPRC_LINE String, sADDNUM As Integer, ClsChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Yang Meng
'   6.Create Date  : 2004. 08 .25
'   7.Modify Date  :
'   8.Comment      : combo Add
'---------------------------------------------------------------------------------------
Public Function Gf_USER_ComboAdd(Conn As ADODB.Connection, Cbo As Variant, sPgmId As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim sQuery As String

    Dim AdoRs As ADODB.Recordset
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_USER_ComboAdd = False: Exit Function
    End If
    
    sQuery = "SELECT EMP_ID FROM ZP_AUTHORITY  "
    sQuery = sQuery + "    WHERE PGMID = '" & sPgmId & "'"
    sQuery = sQuery + "      AND UPD   = '1' AND EMP_ID <> '1JS1005'"
    sQuery = sQuery + "    ORDER BY EMP_ID"

    If ClsChk Then
        Cbo.Clear
    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If VarType(AdoRs.Fields(0)) <> vbNull Then
                Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Gf_USER_ComboAdd = True
    Else
        Gf_USER_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Gf_USER_ComboAdd = False

End Function
