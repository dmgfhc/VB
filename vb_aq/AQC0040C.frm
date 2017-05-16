VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0040C 
   Caption         =   "材质试验实绩确认_AQC0040C"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000E&
      Caption         =   "PWHT性能"
      Height          =   345
      Left            =   13560
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox txt_COLOR_STROKE 
      BackColor       =   &H00E1E4CD&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   8400
      Width           =   14775
   End
   Begin VB.CheckBox CHK_AllItem 
      BackColor       =   &H00E0E0E0&
      Caption         =   "显示所有项目"
      Height          =   255
      Left            =   12480
      TabIndex        =   22
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox chk_Cond 
      BackColor       =   &H00E0E0E0&
      Caption         =   "已确认"
      Height          =   255
      Left            =   9120
      TabIndex        =   21
      Tag             =   ",SUBSTR(A.SLAB_NO,1,8)"
      Top             =   600
      Width           =   1020
   End
   Begin VB.CommandButton cmd_UnCheck 
      Caption         =   "确认取消"
      Height          =   345
      Left            =   9360
      TabIndex        =   20
      Top             =   135
      Width           =   1155
   End
   Begin VB.CheckBox cbo_loc 
      Caption         =   "B"
      Height          =   255
      Index           =   1
      Left            =   14520
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox cbo_loc 
      Caption         =   "T"
      Height          =   255
      Index           =   0
      Left            =   14040
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消保留标记"
      Height          =   345
      Left            =   13560
      TabIndex        =   17
      Top             =   135
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "判定保留标记"
      Height          =   345
      Left            =   12120
      TabIndex        =   16
      Top             =   135
      Width           =   1275
   End
   Begin VB.TextBox txt_SMP_CUT_LOC_NAME 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6210
      TabIndex        =   13
      Top             =   135
      Width           =   2655
   End
   Begin VB.TextBox txt_PLT_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1710
      TabIndex        =   12
      Top             =   510
      Width           =   2430
   End
   Begin VB.TextBox txt_PLT 
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
      Height          =   300
      Left            =   1215
      MaxLength       =   2
      TabIndex        =   11
      Top             =   510
      Width           =   420
   End
   Begin VB.TextBox txt_STDSPEC 
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
      Left            =   1215
      MaxLength       =   18
      TabIndex        =   6
      Top             =   900
      Width           =   2295
   End
   Begin VB.TextBox txt_STDSPEC_NAME 
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
      Height          =   310
      Left            =   3555
      MaxLength       =   18
      TabIndex        =   5
      Top             =   900
      Width           =   570
   End
   Begin InDate.UDate dtp_date_t 
      Height          =   315
      Left            =   7425
      TabIndex        =   4
      Top             =   510
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
   End
   Begin InDate.UDate dtp_date_f 
      Height          =   315
      Left            =   5640
      TabIndex        =   3
      Top             =   510
      Width           =   1470
      _ExtentX        =   2593
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
   Begin VB.CommandButton cmd_AllCheck 
      Caption         =   "全部确认"
      Height          =   345
      Left            =   10800
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
   Begin VB.TextBox txt_SMP_CUT_LOC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5670
      MaxLength       =   1
      TabIndex        =   1
      Top             =   135
      Width           =   450
   End
   Begin VB.TextBox txt_SMP_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1215
      MaxLength       =   14
      TabIndex        =   0
      Top             =   135
      Width           =   1965
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   529
      Caption         =   "试样编号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   1
      Left            =   4545
      Top             =   135
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      Caption         =   "取样位置"
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
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   2
      Left            =   4545
      Top             =   510
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      Caption         =   "生产日期"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   3
      Left            =   240
      Top             =   900
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Caption         =   "标准编号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   4
      Left            =   4545
      Top             =   900
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "产品厚度"
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
      ForeColor       =   -2147483642
   End
   Begin CSTextLibCtl.sidbEdit sdb_ORD_WID 
      Height          =   300
      Left            =   5670
      TabIndex        =   7
      Top             =   900
      Width           =   810
      _Version        =   262145
      _ExtentX        =   1429
      _ExtentY        =   529
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoScroll      =   0   'False
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.0"
      Text            =   " 0.00"
      StartText.x     =   3
      StartText.y     =   2
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
      NumDecDigits    =   1
      NumIntDigits    =   4
      ShowZero        =   0   'False
      Undo            =   0
      Data            =   0
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7065
      Left            =   -360
      TabIndex        =   8
      Top             =   1290
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   12462
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "AQC0040C.frx":0000
      Begin FPSpread.vaSpread SS2 
         Height          =   7065
         Left            =   6765
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   2715
         _Version        =   393216
         _ExtentX        =   4789
         _ExtentY        =   12462
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
         MaxCols         =   7
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0040C.frx":0072
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   7065
         Left            =   9570
         TabIndex        =   15
         Top             =   0
         Width           =   5715
         _Version        =   393216
         _ExtentX        =   10081
         _ExtentY        =   12462
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
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
         MaxCols         =   5
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0040C.frx":0489
      End
      Begin FPSpread.vaSpread SS1 
         Height          =   7065
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   6675
         _Version        =   393216
         _ExtentX        =   11774
         _ExtentY        =   12462
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
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
         MaxCols         =   13
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0040C.frx":07EC
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   5
      Left            =   6615
      Top             =   900
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   529
      Caption         =   "生产日期"
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
   Begin CSTextLibCtl.sitxEdit txt_TIME 
      Height          =   315
      Left            =   7650
      TabIndex        =   10
      Top             =   900
      Width           =   1185
      _Version        =   262145
      _ExtentX        =   2090
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __:__:__"
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
      Text            =   "____-__-__ __:__:__"
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
      Mask            =   "____-__-__"
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
   End
   Begin InDate.ULabel lab_MATR_FL 
      Height          =   315
      Index           =   0
      Left            =   10170
      Top             =   960
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      Caption         =   " 产品力学性能："
      Alignment       =   0
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
      ForeColor       =   255
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   6
      Left            =   225
      Top             =   510
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   529
      Caption         =   "生产厂"
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
Attribute VB_Name = "AQC0040C"
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
'-- Program Name      材质试验实绩输入
'-- Program ID        AQC0030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HAN.Y.S
'-- Coder             ZENG.W SUN BIN
'-- Date              2005.10. 25
'-- Description       材质试验实绩输入
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
Public sPLT_Authority As String     'Active User Plant Authority Setting

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

Dim pColumn12 As New Collection      'Spread Primary Key Collection
Dim nColumn12 As New Collection      'Spread necessary Column Collection
Dim mColumn12 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn12 As New Collection      'Spread Insert Column Collection
Dim aColumn12 As New Collection      'Master -> Spread Column Collection
Dim lColumn12 As New Collection      'Spread Lock Column Collection

Dim pColumn13 As New Collection      'Spread Primary Key Collection
Dim nColumn13 As New Collection      'Spread necessary Column Collection
Dim mColumn13 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn13 As New Collection      'Spread Insert Column Collection
Dim aColumn13 As New Collection      'Master -> Spread Column Collection
Dim lColumn13 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection
Dim sc3 As New Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim arrChem(6, 35) As String

Dim V_SMP_NO, V_SMP_LOC As String
Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_smp_cut_loc, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(dtp_date_f, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(dtp_date_t, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdb_ORD_WID, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(chk_Cond, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    '是否确认，为了给材质取消确认使用 耿学玉  2011 -5-16
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
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQC0040C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQC0040C.P_MODIFY1", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
'     Call SS1.AddCellSpan(5, 0, 1, 2)

      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     
     'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQC0040C.P_SREFER_1", Key:="P-R"
    sc2.Add Item:=pColumn12, Key:="pColumn"
    sc2.Add Item:=nColumn12, Key:="nColumn"
    sc2.Add Item:=aColumn12, Key:="aColumn"
    sc2.Add Item:=mColumn12, Key:="mColumn"
    sc2.Add Item:=iColumn12, Key:="iColumn"
    sc2.Add Item:=lColumn12, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     
     'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AQC0040C.P_SREFER_2", Key:="P-R"
    sc3.Add Item:=pColumn13, Key:="pColumn"
    sc3.Add Item:=nColumn13, Key:="nColumn"
    sc3.Add Item:=aColumn13, Key:="aColumn"
    sc3.Add Item:=mColumn13, Key:="mColumn"
    sc3.Add Item:=iColumn13, Key:="iColumn"
    sc3.Add Item:=lColumn13, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 5, True)
    Call Gp_Sp_ColHidden(ss1, 7, True)
    Call Gp_Sp_ColHidden(ss2, 0, True)
    Call Gp_Sp_ColHidden(ss3, 0, True)
    Call Gp_Sp_ColHidden(ss3, 5, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, 1, ss1.MaxRows, , &HFFFF&)


End Sub

Private Sub cbo_loc_Click(Index As Integer) 'louyannan 20101215
If cbo_loc(Index).Value = "1" Then
cbo_loc(Abs(Index - 1)) = "0"
Else
cbo_loc(Abs(Index - 1)) = "1"
End If

ss1_Click ss1.ActiveCol, ss1.ActiveRow



End Sub

Private Sub cmd_AllCheck_Click()
    Dim i       As Integer
    Dim sAllChk As String
    
    If ss1.MaxRows < 1 Or ss1.Row = 0 Then Exit Sub
    
    If cmd_AllCheck.Caption = "全部确认" Then
        sAllChk = "ALL"
    Else
        sAllChk = ""
    End If
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        
        For i = 1 To ss1.MaxRows
            ss1.Row = i
            If sAllChk = "ALL" Then
                ss1.Col = 1
                ss1.Text = 1
                ss1.Col = 0
                ss1.Text = "Update"
                cmd_AllCheck.Caption = "全部取消"
            Else
                ss1.Col = 1
                ss1.Text = 0
                ss1.Col = 0
                ss1.Text = ""
                cmd_AllCheck.Caption = "全部确认"
            End If
        Next i
              
    End If

End Sub

Private Sub MenuToolSet()
     
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
'    MDIMain.MenuTool.Buttons(14).Enabled = False
    
End Sub

'材质确认取消
Private Sub cmd_UnCheck_Click()
 If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If

If Gf_Sc_Authority(sAuthority, "U") Then
    Call DataSave_D
Else
    Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
End If

End Sub

Private Sub Command1_Click()

    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If

    If Gf_Sc_Authority(sAuthority, "U") Then
      Call DataSave_H
    Else
      Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
    End If
 

End Sub

Private Sub Command2_Click()

    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If

    If Gf_Sc_Authority(sAuthority, "U") Then
      Call DataSave_C
    Else
      Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
    End If
              

End Sub

Private Sub Command3_Click()
  'AQC0041C.txt_PROD_NO.Text = Gf_Get_Cell_Value(SS1, SS1.ActiveRow, 1)
  AQC0041C.txt_SMP_NO.Text = V_SMP_NO           '"15103465010101"
  AQC0041C.txt_smp_cut_loc.Text = V_SMP_LOC     '"T"
  AQC0041C.Show
  AQC0041C.Form_Ref
  
  
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuToolSet

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        Case "txt_PLT"                     '工厂
            sCode = "C0001"
            Set oCodeName = txt_PLT_NAME
        Case "txt_SMP_CUT_LOC"             '取样位置
            sCode = "Q0042"
            Set oCodeName = txt_SMP_CUT_LOC_NAME
        Case "txt_STDSPEC"                 '标准
            sCode = "STDSPEC"
            Set oCodeName = txt_STDSPEC_NAME
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing

Err_Track:
    
    Set oCodeName = Nothing

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)
'    sAuthority = "1111"
    sPLT_Authority = Gf_PLT_Authority(Me.Name)
    If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
       txt_plt.Text = sPLT_Authority
    Else
       txt_plt.Text = ""
    End If
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuToolSet

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(ss1)
    Call Gp_Sp_Setting(ss2)
    Call Gp_Sp_Setting(ss3)
    Call Gp_Sp_ReadOnlySet(ss2)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss2, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss3, "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    cbo_loc(0).Value = "1"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss2, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss3, "Q-System.INI", Me.Name)
    
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
    
    Set iColumn12 = Nothing
    Set pColumn12 = Nothing
    Set lColumn12 = Nothing
    Set nColumn12 = Nothing
    Set mColumn12 = Nothing
    Set aColumn12 = Nothing
    
    Set iColumn13 = Nothing
    Set pColumn13 = Nothing
    Set lColumn13 = Nothing
    Set nColumn13 = Nothing
    Set mColumn13 = Nothing
    Set aColumn13 = Nothing

    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gf_Sp_Cls(sc2)
        Call Gf_Sp_Cls(sc3)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        sdb_ORD_WID.Value = 0
        txt_TIME.RawData = ""
        
        If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
           txt_plt.Text = sPLT_Authority
        Else
           txt_plt.Text = ""
        End If
        
    End If
    
    txt_SMP_CUT_LOC_NAME = ""
    txt_PLT_NAME = ""
    txt_STDSPEC_NAME.Text = ""

End Sub

Public Sub Form_Ref()
    Dim iRow, iCol  As Integer
    Dim sQuery      As String
    Dim sMesg       As String
    Dim AdoRs       As adodb.Recordset

    On Error GoTo Refer_Err
    
    cmd_AllCheck.Caption = "全部确认"
    
    If dtp_date_f.RawData = "" Then
       'dtp_date_f.RawData = Format(Now, "yyyymm") + "01"
       dtp_date_f.RawData = ""
    End If
    
    If dtp_date_t.RawData = "" Then
       dtp_date_t.RawData = Format(Now, "yyyymmdd")
    End If

    If txt_SMP_NO = "" And txt_smp_cut_loc <> "" Then
       MsgBox "请先输入取样号！", vbCritical, "系统提示信息"
       txt_smp_cut_loc = ""
       Exit Sub
    End If

'    If chk_Cond.Value = 1 And (Mid(dtp_date_t.RawData, 1, 6) <> Mid(dtp_date_f.RawData, 1, 6)) Then
'        MsgBox "查询已经确认的试样号，必须是同一个月的数据"
'        Exit Sub
'    End If
    
       If chk_Cond.Value = 1 And txt_SMP_NO = "" Then
        MsgBox "查询已经确认的试样号信息，必须输入试样号"
        Exit Sub
    End If

    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuToolSet
    End If
    
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(sc3)
    
    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub
    
    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)                   ' YELLOW
            ElseIf .Text = "S" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0E0FF)                  ' PINK
            ElseIf .Text = "R" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFFC0)                  ' GREEN
            ElseIf .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFF80FF)                  ' 淡紫色
            ElseIf .Text = "H" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0C0C0)                  ' 灰色
            Else
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)                ' WHITE
            End If
                  
         Next iRow
    End With
    
Refer_Err:
    
    Screen.MousePointer = vbDefault
    

End Sub

Public Sub Form_Pro()

    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If
 
    Call DataSave("1")
    
End Sub

'Private Sub cmd_PIC_Click()
'
'    If sPLT_Authority <> "**" And sPLT_Authority <> txt_PLT.Text Then
'       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
'       Exit Sub
'    End If
'
'   Call DataSave("2")
'
'End Sub

Public Sub DataSave(SaveFL As String)
    Dim iRow, iCol As Integer
    
    Sc1.Remove ("P-M")
    If SaveFL = "1" Then
        Sc1.Add Item:="AQC0040C.P_MODIFY1", Key:="P-M"
    Else
        Sc1.Add Item:="AQC0040C.P_MODIFY2", Key:="P-M"
    End If
    
    With ss1
       For iRow = 1 To .MaxRows
           .Row = iRow
           .Col = 0
           If .Text = "Update" Then
              .Col = 7
              .Text = sUserID
           End If
       Next iRow
    End With
    
    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    ss1.OperationMode = OperationModeNormal
    Call MenuToolSet
    
    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub
    
    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)                   ' YELLOW
            ElseIf .Text = "S" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0E0FF)                  ' PINK
            ElseIf .Text = "R" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFFC0)                  ' GREEN
            ElseIf .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFF80FF)                  ' 淡紫色
            ElseIf .Text = "H" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0C0C0)                  ' 灰色
            Else
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)                ' WHITE
            End If
         Next iRow
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

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
            
    Dim sQuery          As String
    Dim sMesg           As String
    Dim AdoRs           As adodb.Recordset
    Dim ArrayRecords    As Variant
    Dim arr             As Variant
    Dim SMP_NO, smp_loc As Variant
    Dim s_ORD_NO, s_ORD_ITEM As String
    Dim s_MATR_FL  As String
    Dim V_PWHT_FL  As String
    
    Dim s_COLOR_STROKE  As String
    
    s_MATR_FL = "Y"
    
    s_COLOR_STROKE = ""
    
 On Error GoTo Error_Rtn
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    If ss1.MaxRows < 1 Or Row = 0 Or Col = 1 Then Exit Sub
    
    If Col = 0 Then
    
        Unload AQC0080C
        
        ss1.Row = Row
        ss1.Col = 8
        AQC0080C.txt_ORD_NO = Trim(ss1.Text)
        ss1.Col = 9
        AQC0080C.txt_ORD_ITEM = Trim(ss1.Text)
        
        AQC0080C.Show
        AQC0080C.Form_Ref
        
        Exit Sub
        
    End If
    
    With ss1
        .Col = 2
        .Row = .ActiveRow
        SMP_NO = .Text
        V_SMP_NO = .Text
        .Col = 3
        smp_loc = .Text
        V_SMP_LOC = .Text
        .Col = 4
        TXT_STDSPEC = .Text
        .Col = 6
        sdb_ORD_WID = .Text
        .Col = 11
        txt_TIME.RawData = .Text
        .Col = 8
        s_ORD_NO = .Text
        .Col = 9
        s_ORD_ITEM = .Text
        .Col = 13
        V_PWHT_FL = .Text
    End With
    

'    If V_PWHT_FL = "Y" Then
'      Command3.BackColor = "H00FFFF00&"
'    Else
'      Command3.BackColor = "&H8000000F&"
'    End If

    
'    With ss2
'        For i = 1 To .MaxRows
'            .Row = i
'            .Col = 3: v_chem_rslt_fp = Val(.Text)
'            .Col = 4: v_chem_rslt = Val(.Text)
'            .Col = 5: v_chem_diff = Val(.Text)
'            .Col = 8: v_chem_diff_min = Val(.Text)
'            .Col = 9: v_chem_diff_max = Val(.Text)
'
'
'          If v_chem_rslt_fp < v_chem_rslt And v_chem_diff < v_chem_diff_min Then
'            Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, i, i, , &HFF)
'          End If
'
'          If v_chem_rslt_fp > v_chem_rslt And v_chem_diff > v_chem_diff_max Then
'            Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, i, i, , &HFF)
'          End If
'
'        Next i
'
'    End With
    
    ss2.MaxRows = 0
    ss3.MaxRows = 0
    
    ss1.ReDraw = False
    ss2.ReDraw = False
    ss3.ReDraw = False
    
    Set AdoRs = New adodb.Recordset
    sQuery = "SELECT MATR_FL,COLOR_STROKE,INSP_CD  FROM BP_ORDER_ITEM WHERE ORD_NO = " + "'" + s_ORD_NO + "'"
    sQuery = sQuery + " AND ORD_ITEM = " + "'" + s_ORD_ITEM + "'"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not (AdoRs.BOF And AdoRs.EOF) And IsNull(AdoRs.Fields(0).Value) = False Then
        'ArrayRecords = AdoRs.GetRows
        s_MATR_FL = AdoRs.Fields(0).Value
    Else
        s_MATR_FL = "Y"
    End If
      
    If IsNull(AdoRs.Fields(0).Value) = False Then
'        s_COLOR_STROKE = AdoRs.Fields(1).Value
        If IsNull(AdoRs.Fields(1).Value) = True Then
          txt_COLOR_STROKE.Text = " 订单备注： " + s_ORD_NO + "-" + s_ORD_ITEM + ":"
        Else
          txt_COLOR_STROKE.Text = " 订单备注： " + s_ORD_NO + "-" + s_ORD_ITEM + ":" + AdoRs.Fields(1).Value
        End If
        If IsNull(AdoRs.Fields(2).Value) = False Then
            txt_COLOR_STROKE.Text = txt_COLOR_STROKE.Text + "  见证机关：" + AdoRs.Fields(2).Value
        End If
    End If
    
    sQuery = "{call AQC0040C.P_SREFER_1('" + Trim(SMP_NO) + "','" + Trim(smp_loc) + "')}"
    
    AdoRs.Close
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not (AdoRs.BOF And AdoRs.EOF) Then
        ArrayRecords = AdoRs.GetRows
        Call subSpreadView2(ArrayRecords)
        Erase ArrayRecords
    End If
    
'    Call Gp_Sp_EvenRowBackcolor(ss2)
        
    If s_MATR_FL = "N" Or s_MATR_FL = "n" Then
        lab_MATR_FL.Item(0).Caption = " 产品力学性能： " + " 不保证 "
    Else
        lab_MATR_FL.Item(0).Caption = " 产品力学性能： " + " 保证 "
    End If
    
    If smp_loc = "Y" Then
     cbo_loc(0).Visible = True
     cbo_loc(1).Visible = True
    
     If cbo_loc(0).Value = "1" Then 'louyannan  20101215
     smp_loc = "T"
     Else
     smp_loc = "B"
     End If
   Else
    cbo_loc(0).Visible = False
    cbo_loc(1).Visible = False
    
   End If
   
    
    sQuery = "{call AQC0040C.P_SREFER_2('" + Trim(SMP_NO) + "','" + Trim(smp_loc) + "')}"
                    
    AdoRs.Close
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        ArrayRecords = AdoRs.GetRows
        Call subSpreadView1(ArrayRecords)
        Erase ArrayRecords
    End If
     
    sQuery = "{call AQC0040C.P_SREFER_3('" + Trim(SMP_NO) + "')}"
    
    AdoRs.Close
                    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        ArrayRecords = AdoRs.GetRows
        Call subSpreadView3(ArrayRecords)
        Erase ArrayRecords
    End If
    
    
    '--------------------配置化项目显示  刘翔  2012.11.20-----------------------------------------------------
    
    Erase ArrayRecords
    
    AdoRs.Close
    sQuery = "{call AQC0040C.P_SREFER_CONFIG('" + Trim(SMP_NO) + "','" + Trim(smp_loc) + "')}"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Set AdoRs = M_CN1.Execute(sQuery)
    
    If Not AdoRs.EOF And Not AdoRs.BOF Then
      ArrayRecords = AdoRs.GetRows
      Call subSpreadView_Config(ArrayRecords)
    End If
    
    AdoRs.Close
    Erase ArrayRecords
    
    '-----------------------------------------------------------------------------------------------------------
    

    'Call Gp_Sp_EvenRowBackcolor(ss3)
    
    Set AdoRs = Nothing
    Set ArrayRecords = Nothing
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True
    
    
    If V_PWHT_FL = "Y" Then
      Command3.BackColor = &HFFFF00
    Else
      Command3.BackColor = &H8000000E
    End If
    
       
    Exit Sub
    
Error_Rtn:
    
    Set AdoRs = Nothing
    Set ArrayRecords = Nothing
    Screen.MousePointer = vbDefault
    ss1.ReDraw = True
    ss2.ReDraw = True
    ss3.ReDraw = True

End Sub

Private Sub InputEditCheck()

    If ss1.ActiveCol <> 1 Then
        pControl(1).SetFocus
    End If
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call InputEditCheck
End Sub


Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    Call InputEditCheck
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

'Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    ss1.Row = ss1.ActiveRow + 1
'    Call ss1_Click(ss1.Col, ss1.ActiveRow + 1)
'End Sub

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

Private Sub subSpreadView1(ByVal strArr As Variant)

    Dim i           As Integer
    Dim iRow        As Integer
    Dim sMatr(216)   As String '215
    
    If UBound(strArr, 2) < 0 Then Exit Sub
        
    sMatr(0) = "屈服点实绩                            "
    sMatr(1) = "拉伸规定总伸长应力实绩                "
    sMatr(2) = "抗拉强度实绩                          "
    sMatr(3) = "屈强比实绩                            "
    sMatr(4) = "断后伸长率实绩                        "
    sMatr(5) = "断面收缩率实绩1                       "
    sMatr(6) = "断面收缩率实绩2                       "
    sMatr(7) = "断面收缩率实绩3                       "
    sMatr(8) = "断面收缩率实绩平均                    "
    sMatr(9) = "冷弯试验实绩                          "
    sMatr(10) = "冲击试验温度                         "
    sMatr(11) = "冲击试样尺寸                         "
    sMatr(12) = "冲击试验实绩 1                       "
    sMatr(13) = "冲击试验实绩 2                       "
    sMatr(14) = "冲击试验实绩 3                       "
    sMatr(15) = "冲击试验实绩 4                       "
    sMatr(16) = "冲击试验实绩 5                       "
    sMatr(17) = "冲击试验实绩 6                       "
    sMatr(18) = "冲击试验实绩平均                     "
   
    sMatr(19) = "冲击断面纤维率实绩平均                 "
    sMatr(20) = "冲击断面纤维率实绩 1                   "
    sMatr(21) = "冲击断面纤维率实绩 2                   "
    sMatr(22) = "冲击断面纤维率实绩 3                   "
    sMatr(23) = "冲击断面纤维率实绩 4                   "
    sMatr(24) = "冲击断面纤维率实绩 5                   "
    sMatr(25) = "冲击断面纤维率实绩 6                   "
    sMatr(26) = "时效冲击试验温度                     "
    sMatr(27) = "时效冲击试样尺寸                     "
    sMatr(28) = "时效冲击功实绩1                      "
    sMatr(29) = "时效冲击功实绩2                      "
    sMatr(30) = "时效冲击功实绩3                      "
    sMatr(31) = "时效冲击功实绩4                      "
    sMatr(32) = "时效冲击功实绩5                      "
    sMatr(33) = "时效冲击功实绩6                      "
    sMatr(34) = "时效冲击实绩平均                     "
                  
    sMatr(35) = "时效冲击纤维断面率实绩               "
    sMatr(36) = "重力撕裂温度                         "
    sMatr(37) = "重力撕裂实绩1                        "
    sMatr(38) = "重力撕裂实绩2                        "
    sMatr(39) = "重力撕裂实绩平均                     "
    sMatr(40) = "硬度实绩                             "
    sMatr(41) = "拉伸规定非比例伸长应力实绩           "
    sMatr(42) = "拉伸规定残余伸长应力实绩实绩         "
    sMatr(43) = "高温拉伸屈服强度实绩                 "
    sMatr(44) = "高温拉伸抗拉强度实绩                 "
    sMatr(45) = "高温拉伸断面收缩率实绩1              "
'20090806 SUN BIN
    sMatr(46) = "高温拉伸断面收缩率实绩2              "
    sMatr(47) = "高温拉伸断面收缩率实绩3              "
    sMatr(48) = "高温拉伸断面收缩率实绩平均           "
'20090806 SUN BIN END
    sMatr(49) = "高温拉伸断后伸长率实绩               "
    sMatr(50) = "高温拉伸规定非比例伸长应力实绩       "
    sMatr(51) = "高温拉伸规定残余伸长应力实绩         "
    sMatr(52) = "焊接硬度实绩                         "
    sMatr(53) = "焊缝弯曲实绩                         "
    sMatr(54) = "反复弯曲实绩                         "
    sMatr(55) = "锻平试验实绩                         "
    sMatr(56) = "抗氢裂能力CSR实绩（停用）                   "
    sMatr(57) = "抗氢裂能力CLR实绩（停用）                   "
    sMatr(58) = "抗氢裂能力CTR实绩（停用）                   "
    sMatr(59) = "硫化物腐蚀裂纹实绩                   "
    sMatr(60) = "追加冲击试验温度                     "
    sMatr(61) = "追加击试样尺寸                       "
    sMatr(62) = "追加冲击试验实绩平均                 "
    sMatr(63) = "追加冲击试验实绩 1                   "
    sMatr(64) = "追加冲击试验实绩 2                   "
    sMatr(65) = "追加冲击试验实绩 3                   "
    sMatr(66) = "追加冲击试验实绩 4                   "
    sMatr(67) = "追加冲击试验实绩 5                   "
    sMatr(68) = "追加冲击试验实绩 6                   "
    sMatr(69) = "追加冲击剪切面积实绩平均             "
    sMatr(70) = "追加冲击断面纤维率实绩 1               "
    sMatr(71) = "追加冲击断面纤维率实绩 2               "
    sMatr(72) = "追加冲击断面纤维率实绩 3               "
    sMatr(73) = "追加冲击断面纤维率实绩 4               "
    sMatr(74) = "追加冲击断面纤维率实绩 5               "
    sMatr(75) = "追加冲击断面纤维率实绩 6               "
    sMatr(76) = "追加时效冲击试验温度                 "
    sMatr(77) = "追加时效冲击试样尺寸                 "
    sMatr(78) = "追加时效冲击实绩平均                 "
    sMatr(79) = "追加时效冲击功实绩1                  "
    sMatr(80) = "追加时效冲击功实绩2                  "
    sMatr(81) = "追加时效冲击功实绩3                  "
    sMatr(82) = "追加时效冲击功实绩4                  "
    sMatr(83) = "追加时效冲击功实绩5                  "
    sMatr(84) = "追加时效冲击功实绩6                  "
    sMatr(85) = "追加时效冲击纤维断面率实绩           "
    sMatr(86) = "晶粒度实绩                           "
    sMatr(87) = "脱碳层实绩                           "
    sMatr(88) = "硫印实绩                             "
    sMatr(89) = "断口检验实绩1                        "
    sMatr(90) = "断口检验实绩2                        "
    sMatr(91) = "断口检验实绩3                        "
    sMatr(92) = "断口检验实绩4                        "
    sMatr(93) = "断口检验实绩5                        "
    sMatr(94) = "酸浸检验实绩1                        "
    sMatr(95) = "酸浸检验实绩2                        "
    sMatr(96) = "酸浸检验实绩3                        "
    sMatr(97) = "酸浸检验实绩4                        "
    sMatr(98) = "酸浸检验实绩5                        "
    sMatr(99) = "带状组织实绩                         "
    sMatr(100) = "淬透性试验实绩1                     "
    sMatr(101) = "淬透性试验实绩2                      "
    sMatr(102) = "淬透性试验实绩3                      "
    sMatr(103) = "非金属夹杂物(粗)实绩1                "
    sMatr(104) = "非金属夹杂物(粗)实绩2                "
    sMatr(105) = "非金属夹杂物(粗)实绩3                "
    sMatr(106) = "非金属夹杂物(粗)实绩4                "
    sMatr(107) = "非金属夹杂物(细)实绩1                "
    sMatr(108) = "非金属夹杂物(细)实绩2                "
    sMatr(109) = "非金属夹杂物(细)实绩3                "
    sMatr(110) = "非金属夹杂物(细)实绩4                "
    sMatr(111) = "奥氏体晶粒度实绩                     "
    sMatr(112) = "DS类非金属夹杂实绩                   "
    sMatr(113) = "TIN类非金属夹杂实绩                  "
'20090804 sun bin start
    sMatr(114) = "追加屈服点实绩                           "
    sMatr(115) = "追加拉伸规定总伸长应力实绩               "
    sMatr(116) = "追加抗拉强度实绩                         "
    sMatr(117) = "追加屈强比实绩                           "
    sMatr(118) = "追加断后伸长率实绩                       "
    sMatr(119) = "追加断面收缩率实绩1                      "
    sMatr(120) = "追加断面收缩率实绩2                      "
    sMatr(121) = "追加断面收缩率实绩3                      "
    sMatr(122) = "追加断面收缩率实绩平均                   "
    sMatr(123) = "追加冷弯试验实绩                         "
    sMatr(124) = "追加硬度实绩                             "
    sMatr(125) = "追加拉伸规定非比例伸长应力实绩           "
    sMatr(126) = "追加拉伸规定残余伸长应力实绩实绩         "
    sMatr(127) = "追加高温拉伸屈服强度实绩                 "
    sMatr(128) = "追加高温拉伸抗拉强度实绩                 "
    sMatr(129) = "追加高温拉伸断面收缩率实绩1               "
'20090806 sun bin start
    sMatr(130) = "追加高温拉伸断面收缩率实绩2               "
    sMatr(131) = "追加高温拉伸断面收缩率实绩3               "
    sMatr(132) = "追加高温拉伸断面收缩率实绩平均           "
'louyanan 20101121 start
'20090806 sun bin end
    sMatr(133) = "追加高温拉伸断后伸长率实绩               "
    sMatr(134) = "追加高温拉伸规定非比例伸长应力实绩       "
    sMatr(135) = "追加高温拉伸规定残余伸长应力实绩         "
'20090804 sun bin end
   sMatr(136) = "厚度方向断面收缩率实绩1"
   sMatr(137) = "厚度方向断面收缩率实绩2"
   sMatr(138) = "厚度方向断面收缩率实绩3"
    '2016-11-15 ljn 厚度方向断面收缩实际添加4/5/6 start
    sMatr(139) = "厚度方向断面收缩率实绩4              "
    sMatr(140) = "厚度方向断面收缩率实绩5              "
    sMatr(141) = "厚度方向断面收缩率实绩6              "
 '2016-11-15 ljn 厚度方向断面收缩实际添加4/5/6 end
 
    
    sMatr(142) = "厚度方向断面收缩率实绩平均"
  '2016-11-15 ljn  start
    sMatr(143) = "厚度方向抗拉强度1              "
    sMatr(144) = "厚度方向抗拉强度2              "
    sMatr(145) = "厚度方向抗拉强度3              "
  '2016-12-2  LJN  START
    sMatr(146) = "厚度方向抗拉强度4              "
    sMatr(147) = "厚度方向抗拉强度5              "
    sMatr(148) = "厚度方向抗拉强度6              "
 '2016-12-2  LJN  END
    sMatr(149) = "追加厚度方向面收缩率1              "
    sMatr(150) = "追加厚度方向面收缩率2              "
    sMatr(151) = "追加厚度方向面收缩率3          "
    sMatr(152) = "追加厚度方向面收缩率4              "
    sMatr(153) = "追加厚度方向面收缩率5              "
    sMatr(154) = "追加厚度方向面收缩率6          "
    sMatr(155) = "追加厚度方向面收缩率平均           "
    sMatr(156) = "追加厚度方向抗拉强度1              "
    sMatr(157) = "追加厚度方向抗拉强度2          "
    sMatr(158) = "追加厚度方向抗拉强度3              "
'2016-12-2  LJN  START
    sMatr(159) = "追加厚度方向抗拉强度4              "
    sMatr(160) = "追加厚度方向抗拉强度5              "
    sMatr(161) = "追加厚度方向抗拉强度6              "
 '2016-12-2  LJN  END
    
 '2016-11-15 ljn  end

   
   sMatr(162) = "高温拉伸厚度方向断面收缩率实绩1"
   sMatr(163) = "高温拉伸厚度方向断面收缩率实绩2"
   sMatr(164) = "高温拉伸厚度方向断面收缩率实绩3"
   sMatr(165) = "高温拉伸厚度方向断面收缩率实绩平均"
   sMatr(166) = "冲击侧向膨胀值实绩1"
   sMatr(167) = "冲击侧向膨胀值实绩2"
   sMatr(168) = "冲击侧向膨胀值实绩3"
   sMatr(169) = "冲击侧向膨胀值实绩4"
   sMatr(170) = "冲击侧向膨胀值实绩5"
   sMatr(171) = "冲击侧向膨胀值实绩6"
   sMatr(172) = "冲击侧向膨胀值实绩平均"
   sMatr(173) = "追加冲击侧向膨胀值实绩1"
   sMatr(174) = "追加冲击侧向膨胀值实绩2"
   sMatr(175) = "追加冲击侧向膨胀值实绩3"
   sMatr(176) = "追加冲击侧向膨胀值实绩4"
   sMatr(177) = "追加冲击侧向膨胀值实绩5"
   sMatr(178) = "追加冲击侧向膨胀值实绩6"
   sMatr(179) = "追加冲击侧向膨胀值实绩平均"
   sMatr(180) = "NDT重力撕裂实绩"
'edit by gengxueyu 20110212 for kangda start
   sMatr(181) = "均匀变形伸长率UEL"
   sMatr(182) = "追加均匀变形伸长率UEL"
   sMatr(183) = "追加应力比项目1"
   sMatr(184) = "追加应力比项目2"
   sMatr(185) = "追加应力比项目3"
   sMatr(186) = "追加应力比项目4"
   sMatr(187) = "追加应力比项目5"
'edit by gengxueyu 20110212 for kangda end
    
   'HIC结果由一组3个增加到三组9个，原显示在中间现改到最后  刘翔  2014.8.5
   sMatr(188) = "抗氢裂能力CSR实绩1                   "
   sMatr(189) = "抗氢裂能力CLR实绩1                   "
   sMatr(190) = "抗氢裂能力CTR实绩1                   "
   sMatr(191) = "抗氢裂能力CSR实绩2                   "
   sMatr(192) = "抗氢裂能力CLR实绩2                   "
   sMatr(193) = "抗氢裂能力CTR实绩2                   "
   sMatr(194) = "抗氢裂能力CSR实绩3                   "
   sMatr(195) = "抗氢裂能力CLR实绩3                   "
   sMatr(196) = "抗氢裂能力CTR实绩3                   "
   sMatr(197) = "抗氢裂能力CSR实绩4                   "
   sMatr(198) = "抗氢裂能力CLR实绩4                   "
   sMatr(199) = "抗氢裂能力CTR实绩4                   "
   sMatr(200) = "抗氢裂能力CSR实绩5                   "
   sMatr(201) = "抗氢裂能力CLR实绩5                   "
   sMatr(202) = "抗氢裂能力CTR实绩5                   "
   sMatr(203) = "抗氢裂能力CSR实绩6                   "
   sMatr(204) = "抗氢裂能力CLR实绩6                   "
   sMatr(205) = "抗氢裂能力CTR实绩6                   "
   sMatr(206) = "抗氢裂能力CSR实绩7                   "
   sMatr(207) = "抗氢裂能力CLR实绩7                   "
   sMatr(208) = "抗氢裂能力CTR实绩7                   "
   sMatr(209) = "抗氢裂能力CSR实绩8                   "
   sMatr(210) = "抗氢裂能力CLR实绩8                   "
   sMatr(211) = "抗氢裂能力CTR实绩8                   "
   sMatr(212) = "抗氢裂能力CSR实绩9                   "
   sMatr(213) = "抗氢裂能力CLR实绩9                   "
   sMatr(214) = "抗氢裂能力CTR实绩9                   " '208
   sMatr(215) = "断口                  "
  
    With ss3
        .MaxRows = 216 '166'215
    
        For i = 1 To 216 '166'215
            .Row = i
            .Col = 1: .Text = sMatr(i - 1)   '初始化试验项目名称
        Next i
                
        For i = 1 To UBound(strArr, 1) + 1
        
            .Row = i: .Col = 4
            .Text = NullCheck(strArr(i - 1, 0), "")  '初始化实绩值
            
        Next i
    End With

End Sub


Private Sub subSpreadView3(ByVal strArr As Variant)

    Dim i                     As Integer
    Dim iRow                  As Integer
    Dim sMatr(3, 216)         As Variant '166
    Dim sMatrCON(6, 216)      As Variant '166
    Dim sMin, sMax, sFL, sRE  As Variant
    
    If UBound(strArr, 2) < 0 Then Exit Sub
      
    If UBound(strArr, 2) = 0 Then        '初始化用与显示数据的数组
        For i = 0 To 215 '165
            sMatr(0, i) = NullCheck(strArr(i, 0), "")
        Next i
        
        For i = 0 To 215 '165
            sMatr(1, i) = NullCheck(strArr(i + 216, 0)) '166
        Next i
    
        For i = 0 To 215 '165
            sMatr(2, i) = NullCheck(strArr(i + 432, 0)) '332
        Next i
        
        
        With ss3
                
            For i = 1 To 216 '166
                .Row = i
                .Col = 2: .Text = sMatr(1, i - 1)
                .Col = 3: .Text = sMatr(2, i - 1)
                .Col = 5: .Text = sMatr(0, i - 1)
            Next i
         End With
    End If
     
    If UBound(strArr, 2) = 1 Then
        For i = 0 To 215 '165
            sMatrCON(0, i) = NullCheck(strArr(i, 0), "")
            sMatrCON(3, i) = NullCheck(strArr(i, 1), "")
        Next i
        
        For i = 0 To 215 '165
            sMatrCON(1, i) = NullCheck(strArr(i + 216, 0)) '166
            sMatrCON(4, i) = NullCheck(strArr(i + 216, 1)) '166
        Next i
    
        For i = 0 To 215 '165
            sMatrCON(2, i) = NullCheck(strArr(i + 432, 0)) '332
            sMatrCON(5, i) = NullCheck(strArr(i + 432, 1)) '332
        Next i
        
            
        For i = 1 To 215 '166
            If sMatrCON(0, i - 1) = "A" Or sMatrCON(0, i - 1) = "B" Then
                If sMatrCON(3, i - 1) = "A" Or sMatrCON(3, i - 1) = "B" Then
                   If Val(sMatrCON(1, i - 1)) >= Val(sMatrCON(4, i - 1)) Then
                      sMin = sMatrCON(1, i - 1)
                   Else
                      sMin = sMatrCON(4, i - 1)
                   End If
                   If Val(sMatrCON(2, i - 1)) = 0 Then
                        sMax = sMatrCON(5, i - 1)
                   Else
                        If Val(sMatrCON(2, i - 1)) >= Val(sMatrCON(5, i - 1)) Then
                           sMax = sMatrCON(5, i - 1)
                        Else
                           sMax = sMatrCON(2, i - 1)
                        End If
                   End If
                   sFL = "A"
                Else
                   sFL = "A"
                   sMin = sMatrCON(1, i - 1)
                   sMax = sMatrCON(2, i - 1)
                End If
               
            Else
                  If sMatrCON(3, i - 1) = "A" Or sMatrCON(3, i - 1) = "B" Then
                     sFL = "A"
                     sMin = sMatrCON(4, i - 1)
                     sMax = sMatrCON(5, i - 1)
                  Else
                     sFL = ""
                     sMin = ""
                     sMax = ""
                  End If
                  
            End If
            With ss3
                .Row = i
                .Col = 2: .Text = sMin
                .Col = 3: .Text = sMax
                .Col = 5: .Text = sFL
            End With
            
         Next i
    End If
     
     Call subSpreadCheck1
     Call subSpreadERROR(ss3)
      With ss3
        For i = 1 To .MaxRows
            sRE = Gf_Get_Cell_Value(ss3, i, 4)
            sFL = Gf_Get_Cell_Value(ss3, i, 5)
            If sFL = "A" And sRE = "" Then
             .Col = 4
             .BackColor = RED
            End If
        Next i
      End With
    

End Sub


Private Sub subSpreadView2(ByVal strArr As Variant)

    Dim i As Integer
    Dim iRow As Integer
    Dim sChem(34) As String
    Dim TEMP As String
    
    
    If UBound(strArr) < 104 Then Exit Sub
    
    sChem(0) = "C  "
    sChem(1) = "Mn "
    sChem(2) = "P  "
    sChem(3) = "S  "
    sChem(4) = "Si "
    sChem(5) = "Nb "
    sChem(6) = "Als"
    sChem(7) = "Alt"
    sChem(8) = "Ceq"
    sChem(9) = "Ni "
    sChem(10) = "Cr "
    sChem(11) = "Cu "
    sChem(12) = "Mo "
    sChem(13) = "V  "
    sChem(14) = "Ti "
    sChem(15) = "Pcm"
    sChem(16) = "W  "
    sChem(17) = "B  "
    sChem(18) = "Pb "
    sChem(19) = "Ca "
    sChem(20) = "N  "
    sChem(21) = "O  "
    sChem(22) = "H  "
    sChem(23) = "Zr "
    sChem(24) = "Mg "
    sChem(25) = "Sn "
    sChem(26) = "As "
    sChem(27) = "Co "
    sChem(28) = "Te "
    sChem(29) = "Bi "
    sChem(30) = "Sb "
    sChem(31) = "Zn "
    sChem(32) = "RE "
    sChem(33) = "Se "
    sChem(34) = "Ta "

    For i = 0 To 34
        
        arrChem(0, i) = NullCheck(strArr(i, 0), "")
    
    Next i
    
    For i = 0 To 34
        
        arrChem(1, i) = NullCheck(strArr(i + 35, 0))
    
    Next i

    For i = 0 To 34
        
        arrChem(2, i) = NullCheck(strArr(i + 70, 0))
    
    Next i
    
    '成品成分
    For i = 0 To 34
        
        arrChem(3, i) = NullCheck(strArr(i + 105, 0), "")
    
    Next i
    
    For i = 0 To 34
        
        arrChem(4, i) = NullCheck(strArr(i + 140, 0))
    
    Next i

    For i = 0 To 34
        
        arrChem(5, i) = NullCheck(strArr(i + 175, 0))
    
    Next i
    
    With ss2
    
        .MaxRows = 0
        .MaxRows = 35
    
        For i = 1 To 35
            .Row = i
            .Col = 1: .Text = sChem(i - 1)
            .Col = 2: .Text = arrChem(1, i - 1)
            .Col = 3: .Text = arrChem(0, i - 1)
            .Col = 4: .Text = arrChem(2, i - 1)
            .Col = 5: .Text = arrChem(4, i - 1)
            .Col = 6: .Text = arrChem(3, i - 1)
            .Col = 7: .Text = arrChem(5, i - 1)
        Next i
          
    End With
    
    Call subSpreadCheck2
    Call subSpreadERROR(ss2)
End Sub

Private Sub subSpreadCheck2()
    
    Dim i As Long
    Dim j As Long
    
    j = 15
    With ss2
        
        For i = 16 To 35
                                    
            If (Gf_Get_Cell_Value(ss2, i, 4) = "" Or Gf_Get_Cell_Value(ss2, i, 4) = "0") _
               And (Gf_Get_Cell_Value(ss2, i, 2) = "0" And Gf_Get_Cell_Value(ss2, i, 3) = "0") Then
                .Row = i
                .RowHidden = True
            Else
                .RowHidden = False
                j = j + 1
                .Col = 0: .Text = j
            End If
        Next i
                
    End With
    
End Sub

Private Sub subSpreadCheck1()
    
    Dim i As Long
    Dim j As Long
    
    With ss3
       
       For i = 1 To .MaxRows

           If Gf_Get_Cell_Value(ss3, i, 5) <> "A" And Gf_Get_Cell_Value(ss3, i, 5) <> "B" Then
               .Row = i
               .RowHidden = True
           Else
                .RowHidden = False
                j = j + 1
                .Col = 0: .Text = j

           End If
           
           '旧的HIC项目隐藏，新的9个结果显示在最后
           If i = 57 Or i = 58 Or i = 59 Then
             .Row = i
             .RowHidden = True
           End If
           
''           以前用户提出 只有管线钢 需要这个项目显示 现在修改为 增加9Ni钢的显示 耿学玉  2011-5-10
''           断面纤维率 增加管线管 "QSY-X70M HD1" 刘翔 2012.6.25
''            If Mid(Trim(txt_STDSPEC), 1, 3) <> "API" And Mid(Trim(txt_STDSPEC), 1, 10) <> "GB/T9711.2" And Trim(txt_STDSPEC) <> "70081MR-06Ni9" And Trim(txt_STDSPEC) <> "ASTM A553-9Ni" And Trim(txt_STDSPEC) <> "QSY-X70M HD1"  Then
''                If i = 20 Or i = 21 Or i = 22 Or i = 23 Or i = 24 _
''                   Or i = 25 Or i = 26 Or i = 70 Or i = 71 Or i = 72 _
''                   Or i = 73 Or i = 74 Or i = 75 Or i = 76 Then
''                   .RowHidden = True
''                End If
''            End If
''            增加 测膨胀值 只有9Ni钢才显示，因为用的都是冲击的判定代码 所以显示的内容较多
''            增加70021MR-06Ni9  20110819  liuxiang
''            If Trim(txt_STDSPEC) <> "70081MR-06Ni9" And Trim(txt_STDSPEC) <> "70021MR-06Ni9" And Trim(txt_STDSPEC) <> "70131MR-06Ni9" And Trim(txt_STDSPEC) <> "ASTM A553-9Ni" Then
''                If i = 144 Or i = 145 Or i = 146 Or i = 147 Or i = 148 _
''                   Or i = 149 Or i = 150 Or i = 151 Or i = 152 Or i = 153 _
''                   Or i = 154 Or i = 155 Or i = 156 Or i = 157 Then
''                   .RowHidden = True
''                End If
''            End If
            
            If CHK_AllItem.Value = 1 Then
                '断面纤维率
                If i = 20 Or i = 21 Or i = 22 Or i = 23 Or i = 24 _
                   Or i = 25 Or i = 26 Or i = 70 Or i = 71 Or i = 72 _
                   Or i = 73 Or i = 74 Or i = 75 Or i = 76 Then
                   .RowHidden = False
                End If
                '侧膨胀值
                If i = 167 Or i = 168 Or i = 169 Or i = 170 Or i = 171 _
                   Or i = 172 Or i = 173 Or i = 174 Or i = 175 Or i = 176 _
                   Or i = 177 Or i = 178 Or i = 179 Or i = 180 Then
                   .RowHidden = False
                End If
            Else
                '断面纤维率
                If i = 20 Or i = 21 Or i = 22 Or i = 23 Or i = 24 _
                   Or i = 25 Or i = 26 Or i = 70 Or i = 71 Or i = 72 _
                   Or i = 73 Or i = 74 Or i = 75 Or i = 76 Then
                   .RowHidden = True
                End If
                '侧膨胀值
                If i = 167 Or i = 168 Or i = 169 Or i = 170 Or i = 171 _
                   Or i = 172 Or i = 173 Or i = 174 Or i = 175 Or i = 176 _
                   Or i = 177 Or i = 178 Or i = 179 Or i = 180 Then
                   .RowHidden = True
                End If

            End If
            
            
       Next i
                
    End With
End Sub
'Private Sub txt_SMP_CUT_LOC_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
''        If txt_SMP_NO = "" Then
''           MsgBox "请先输入取样号！", vbCritical, "系统提示信息"
''           txt_SMP_CUT_LOC = ""
''           Exit Sub
''        End If
'
'        DD.sWitch = "MS"
'        DD.sKey = "Q0042"
'        DD.rControl.Add Item:=txt_SMP_CUT_LOC
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'End Sub

Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    If Row < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        With ss1
            .Row = Row
            .Col = 5
            If .BackColor = &HFFFF& Then
                .Col = 1
                If .Text = "1" Then
                    .Col = 5:   .Text = ""
                    .Col = 0:   .Text = "Update"
                Else
                    .Col = 5:   .Text = "Y"
                    .Col = 0:   .Text = ""
                End If
            Else
                .Col = 1
                If .Text = "1" Then
                    .Col = 5:   .Text = "Y"
                    .Col = 0:   .Text = "Update"
                Else
                    .Col = 5:   .Text = ""
                    .Col = 0:   .Text = ""
                End If
            End If
        End With
    End If
    
End Sub



Private Sub txt_STDSPEC_Change()
    If Trim(TXT_STDSPEC.Text) = "" Then
        txt_STDSPEC_NAME.Text = ""
    End If

End Sub
Private Sub subSpreadERROR(sPname As vaSpread)
    
    Dim i As Long
    Dim C_DSC_CD, C_MAX, C_MIN, C_RESULT, C_FL, C_RSLT_VAL As Variant

    With sPname
    
       If .MaxRows < 1 Then Exit Sub
       
       For i = 1 To .MaxRows
           .Row = i
           C_DSC_CD = Gf_Get_Cell_Value(ss3, i, 5) '(Gf_Get_Cell_Value(sPname, i, 5))
           C_MIN = Val(Gf_Get_Cell_Value(sPname, i, 2))
           C_MAX = Val(Gf_Get_Cell_Value(sPname, i, 3))
           C_RESULT = Val(Gf_Get_Cell_Value(sPname, i, 4))
           C_RSLT_VAL = Gf_Get_Cell_Value(ss3, i, 4)
           If C_MIN <> 0 And C_MAX <> 0 Then
              If C_RESULT > C_MAX Or C_RESULT < C_MIN Then
                 Call Gp_Sp_CellColor(sPname, 4, i, RED)
              End If
           Else
              If C_MIN = 0 And C_MAX <> 0 Then
                 If C_RESULT > C_MAX Then
                    Call Gp_Sp_CellColor(sPname, 4, i, RED)
                 End If
              Else
                 If C_MIN <> 0 And C_MAX = 0 Then
                    If C_RESULT < C_MIN Then
                      Call Gp_Sp_CellColor(sPname, 4, i, RED)
                    End If
                 End If
              End If
           End If
           If C_DSC_CD = "A" Or C_DSC_CD = "B" Then
              If C_RSLT_VAL = "N" Then
                 Call Gp_Sp_CellColor(sPname, 4, i, RED)
              End If
           End If
          
       Next i
 
    End With
    
End Sub

Public Sub DataSave_H()
    Dim iRow, iCol As Integer

    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_H", Key:="P-M"

    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = "Update"
           .Col = 7
           .Text = sUserID
    End With

    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

    ss1.OperationMode = OperationModeNormal
    Call MenuToolSet

    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub

    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)                   ' YELLOW
            ElseIf .Text = "S" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0E0FF)                  ' PINK
            ElseIf .Text = "R" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFFC0)                  ' GREEN
            ElseIf .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFF80FF)                  ' 淡紫色
            ElseIf .Text = "H" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0C0C0)                  ' 灰色
            Else
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)                ' WHITE
            End If
         Next iRow
    End With
    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_H", Key:="P-M"
    
    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = ""
    End With
    

End Sub

Public Sub DataSave_C()
    Dim iRow, iCol As Integer

    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_C", Key:="P-M"

    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = "Update"
           .Col = 7
           .Text = sUserID
    End With

    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

    ss1.OperationMode = OperationModeNormal
    Call MenuToolSet

    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub

    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)                   ' YELLOW
            ElseIf .Text = "S" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0E0FF)                  ' PINK
            ElseIf .Text = "R" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFFC0)                  ' GREEN
            ElseIf .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFF80FF)                  ' 淡紫色
            ElseIf .Text = "H" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H808080)                        ' 红色
            Else
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)                ' WHITE
            End If
         Next iRow
    End With
    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_C", Key:="P-M"
    
    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = ""
    End With
    

End Sub

'材质确认取消
Public Sub DataSave_D()
    Dim iRow, iCol As Integer

    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_D", Key:="P-M"

    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = "Update"
           .Col = 7
           .Text = sUserID
    End With

    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

    ss1.OperationMode = OperationModeNormal
    Call MenuToolSet

    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub

    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 5
            If .Text = "Y" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)                   ' YELLOW
            ElseIf .Text = "S" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HC0E0FF)                  ' PINK
            ElseIf .Text = "R" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFFC0)                  ' GREEN
            ElseIf .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFF80FF)                  ' 淡紫色
            ElseIf .Text = "H" Then
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H808080)                        ' 红色
            Else
               Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &H80000005)                ' WHITE
            End If
         Next iRow
    End With
    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0040C.P_MODIFY_D", Key:="P-M"
    
    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = ""
    End With
    

End Sub


Private Sub subSpreadView_Config(ByVal strArr As Variant)

    Dim i As Integer
    Dim OLD_MAXROWS As Integer
    Dim sMin, sMax, sFL, sRE  As Variant
    
    If UBound(strArr, 2) < 0 Then Exit Sub
    
    With ss3
        OLD_MAXROWS = .MaxRows
        .MaxRows = .MaxRows + UBound(strArr, 2) + 1

        For i = 1 To UBound(strArr, 2) + 1
            .Row = OLD_MAXROWS + i
            .Col = 1: .Text = GF_NullChange(strArr(0, i - 1))
            .Col = 2: .Text = GF_NullChange(strArr(1, i - 1)) & ""
            .Col = 3: .Text = GF_NullChange(strArr(2, i - 1)) & ""
            .Col = 4: .Text = GF_NullChange(strArr(3, i - 1)) & ""
            .Col = 5: .Text = GF_NullChange(strArr(4, i - 1)) & ""
        Next i
            
    End With
    
   'subSpreadCheck1 隐藏试验项目或判定代码为空的行，且重写第一列顺序号
    'Call subSpreadCheck1
    
    '检查上下限，超标实绩红字
    Call subSpreadERROR(ss3)
    
    '检查有无实绩，如果有判定代码但没实绩，底色变红
      With ss3
        For i = 1 To .MaxRows
            sRE = Gf_Get_Cell_Value(ss3, i, 4)
            sFL = Gf_Get_Cell_Value(ss3, i, 5)
            If sFL = "A" And sRE = "" Then
             .Col = 4
             .BackColor = RED
            End If
        Next i
      End With

End Sub


