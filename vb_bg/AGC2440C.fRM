VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2440C 
   Caption         =   "剪切前当班取样项目查询界面_AGC2440C"
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
   Begin VB.TextBox txt_HTM_PLT 
      Height          =   285
      Left            =   14400
      TabIndex        =   18
      Text            =   "0"
      Top             =   90
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txt_PLATE_NO 
      Height          =   315
      Left            =   10530
      MaxLength       =   14
      TabIndex        =   15
      Top             =   540
      Width           =   1530
   End
   Begin VB.TextBox txt_OUT_SHEET_NO 
      Height          =   315
      Left            =   13500
      MaxLength       =   14
      TabIndex        =   14
      Top             =   540
      Width           =   1560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   60
      TabIndex        =   11
      Top             =   465
      Width           =   2400
      Begin VB.TextBox txt_LINE 
         Height          =   300
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   13
         Top             =   90
         Width           =   450
      End
      Begin VB.CheckBox txt_DH_FL 
         BackColor       =   &H00E0E0E0&
         Caption         =   "热处理"
         Height          =   240
         Left            =   45
         TabIndex        =   12
         Top             =   150
         Width           =   915
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Index           =   2
         Left            =   1110
         Top             =   90
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         Caption         =   "产线"
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
   End
   Begin VB.TextBox txt_HEAT 
      Height          =   285
      Left            =   13890
      TabIndex        =   7
      Text            =   "0"
      Top             =   90
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txt_STDSPEC 
      Height          =   315
      Left            =   9060
      MaxLength       =   14
      TabIndex        =   4
      Top             =   90
      Width           =   1530
   End
   Begin VB.TextBox txt_PLT 
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
      Left            =   975
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   90
      Width           =   1470
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8295
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   945
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   14631
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   34
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AGC2440C.frx":0000
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   60
      Top             =   90
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      Caption         =   "工厂"
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
   Begin InDate.UDate dtp_DATE_FR 
      Height          =   315
      Left            =   4170
      TabIndex        =   2
      Top             =   90
      Width           =   1500
      _ExtentX        =   2646
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
      MaxLength       =   10
   End
   Begin InDate.UDate dtp_DATE_TO 
      Height          =   315
      Left            =   5640
      TabIndex        =   3
      Top             =   90
      Width           =   1500
      _ExtentX        =   2646
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
      MaxLength       =   10
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   2940
      Top             =   90
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "日期"
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
      Left            =   7710
      Top             =   90
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "标准号"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   1
      Left            =   11070
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "厚度"
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
   Begin CSTextLibCtl.sidbEdit txt_THK_FR 
      Height          =   315
      Left            =   12210
      TabIndex        =   5
      Top             =   90
      Width           =   735
      _Version        =   262145
      _ExtentX        =   1296
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   1
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   3
      MaxValue        =   130
      MinValue        =   1
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit txt_THK_TO 
      Height          =   315
      Left            =   12960
      TabIndex        =   6
      Top             =   90
      Width           =   765
      _Version        =   262145
      _ExtentX        =   1349
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   1
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   3
      MaxValue        =   130
      MinValue        =   1
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   405
      Left            =   3870
      TabIndex        =   8
      Top             =   495
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   714
      _Version        =   196609
      BackStyle       =   1
      Begin VB.OptionButton opt_HTM_PLT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "板卷"
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   10
         Top             =   90
         Width           =   720
      End
      Begin VB.OptionButton opt_HTM_PLT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "中板"
         Height          =   270
         Index           =   1
         Left            =   750
         TabIndex        =   9
         Top             =   75
         Width           =   675
      End
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   9180
      Top             =   540
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "钢板号"
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   5610
      Top             =   540
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "时间"
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
      Left            =   12360
      Top             =   540
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "轧批号"
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
   Begin CSTextLibCtl.sitxEdit dtp_TIME_FR 
      Height          =   315
      Left            =   6840
      TabIndex        =   16
      Tag             =   "时间"
      Top             =   540
      Width           =   960
      _Version        =   262145
      _ExtentX        =   1693
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Text            =   "__:__:__"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   "__:__:__"
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
      ValidateMask    =   0   'False
   End
   Begin CSTextLibCtl.sitxEdit dtp_TIME_TO 
      Height          =   315
      Left            =   7980
      TabIndex        =   17
      Tag             =   "时间"
      Top             =   540
      Width           =   900
      _Version        =   262145
      _ExtentX        =   1587
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Text            =   "__:__:__"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   "__:__:__"
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
      ValidateMask    =   0   'False
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   7800
      Top             =   540
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   556
      Caption         =   "-"
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
   Begin InDate.ULabel ULabel12 
      Height          =   315
      Index           =   3
      Left            =   2700
      Top             =   540
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Caption         =   "热处理工厂"
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
End
Attribute VB_Name = "AGC2440C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       剪切前当班取样项目查询界面
'-- Sub_System Name
'-- Program Name
'-- Program ID        AGC2440C
'-- Document No       Q-00-0010(Specification)
'-- Designer          SUN BIN
'-- Coder             SUN BIN
'-- Date              2009.07.20
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
   ' FormType = "Msheet"
   FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(dtp_DATE_FR, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(dtp_DATE_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(dtp_TIME_FR, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(dtp_TIME_TO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_HEAT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_OUT_SHEET_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_THK_FR, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(TXT_THK_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_HTM_PLT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                            
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
    'Spread_Collection
    'Call Spread_Collection("Column1_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
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
       Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2440C.P_REFER", Key:="P-R"
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

Private Sub dtp_TIME_FR_DblClick()
    
    dtp_TIME_FR.RawData = Gf_DTSet(M_CN1, "T", "X")

End Sub

Private Sub dtp_TIME_TO_DblClick()
    
    dtp_TIME_TO.RawData = Gf_DTSet(M_CN1, "T", "X")

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
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    '    板卷不显示轧制批号
    If App.Title = "BG" Or App.Title = "DG" Then
         Call Gp_Sp_ColHidden(ss1, 2, True)
    End If
    
    Screen.MousePointer = vbDefault
       
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
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

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
    
On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    Exit Sub

Refer_Err:
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

Public Sub Form_Exit()

    Unload Me
    
End Sub

Private Sub opt_HTM_PLT_Click(Index As Integer)
   txt_HTM_PLT.Text = Index
   Call Form_Ref
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
        
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Sub txt_DH_FL_Click()
    
    If txt_DH_FL.Value = 1 Then
       txt_HEAT.Text = 1
    Else
       txt_HEAT.Text = 0
    End If
    
End Sub


Private Sub txt_plt_DblClick()

    Call txt_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Public Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

    
End Sub

Private Sub txt_STDSPEC_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=TXT_STDSPEC

        Call Gf_StdSPEC_DD2(M_CN1, KeyCode)
        
    End If

End Sub
