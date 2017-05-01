VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB5050C 
   Caption         =   "板坯卸车实绩录入_ACB5050C"
   ClientHeight    =   8190
   ClientLeft      =   225
   ClientTop       =   2310
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand cmd_Multi_Print 
      Height          =   345
      Left            =   8625
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   90
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "码单多种打印"
   End
   Begin VB.CheckBox chk_Excel_Fl 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Excel下载后打印"
      Height          =   210
      Left            =   8490
      TabIndex        =   15
      Top             =   585
      Width           =   1665
   End
   Begin VB.TextBox text_cur_inv_name 
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
      Left            =   3675
      TabIndex        =   12
      Top             =   105
      Width           =   1500
   End
   Begin VB.TextBox txt_to_inv_name 
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
      Left            =   7005
      TabIndex        =   11
      Tag             =   "目标库"
      Top             =   105
      Width           =   1455
   End
   Begin VB.TextBox txt_mv_lst_no 
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
      Left            =   1395
      MaxLength       =   15
      TabIndex        =   10
      Tag             =   "移拨码单号"
      Top             =   495
      Width           =   1830
   End
   Begin VB.TextBox text_cur_inv_code 
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
      Left            =   3285
      MaxLength       =   2
      TabIndex        =   6
      Tag             =   "出库"
      Top             =   105
      Width           =   375
   End
   Begin VB.TextBox txt_to_inv 
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
      Left            =   6615
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "工 厂"
      Top             =   105
      Width           =   375
   End
   Begin VB.TextBox text_prod_cd 
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
      Left            =   1395
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "产品"
      Top             =   105
      Width           =   375
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   10215
      Top             =   495
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "数量合计"
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
      Height          =   315
      Left            =   12525
      Top             =   495
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "重量合计"
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
   Begin CSTextLibCtl.sidbEdit text_tot_wgt 
      Height          =   315
      Left            =   13635
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   495
      Width           =   1110
      _Version        =   262145
      _ExtentX        =   1958
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   255
      BackColor       =   16777215
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
      DataProperty    =   2
      ReadOnly        =   -1  'True
      Insert          =   0   'False
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.000"
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit text_tot_sheets 
      Height          =   315
      Left            =   11295
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   495
      Width           =   915
      _Version        =   262145
      _ExtentX        =   1614
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   16711680
      BackColor       =   16777215
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
      DataProperty    =   2
      ReadOnly        =   -1  'True
      Insert          =   0   'False
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   10215
      Top             =   105
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "转库日期"
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
   Begin InDate.UDate udate_in_plt_date_a 
      Height          =   315
      Left            =   11295
      TabIndex        =   7
      Tag             =   "转库日期"
      Top             =   105
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
   End
   Begin InDate.UDate udate_in_plt_date_b 
      Height          =   315
      Left            =   12945
      TabIndex        =   8
      Tag             =   "转库日期"
      Top             =   105
      Width           =   1455
      _ExtentX        =   2566
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
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   2010
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "起始库"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   5340
      Tag             =   "目标库"
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "目标库"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   120
      Top             =   105
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "产品"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   120
      Top             =   495
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "移拨码单号"
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
   Begin InDate.ULabel ULabel34 
      Height          =   315
      Left            =   3345
      Top             =   495
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "到达日期"
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
   Begin CSTextLibCtl.sitxEdit txt_input_date 
      Height          =   315
      Left            =   4620
      TabIndex        =   13
      Tag             =   "到达日期"
      Top             =   495
      Width           =   2115
      _Version        =   262145
      _ExtentX        =   3731
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
      Mask            =   "____-__-__ __:__:__"
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
      ValidateMask    =   0   'False
   End
   Begin Threed.SSCommand cmd_input 
      Height          =   345
      Left            =   6780
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   510
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "录入转库实绩"
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8280
      Left            =   90
      TabIndex        =   16
      Top             =   900
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14605
      _Version        =   196609
      SplitterBarWidth=   3
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "ACB5050C.frx":0000
      Begin FPSpread.vaSpread ss2 
         Height          =   3315
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   0
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   5847
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
         MaxCols         =   14
         MaxRows         =   1
         Protect         =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "ACB5050C.frx":0052
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   4920
         Left            =   0
         TabIndex        =   18
         Top             =   3360
         Width           =   15165
         _Version        =   393216
         _ExtentX        =   26749
         _ExtentY        =   8678
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
         MaxCols         =   20
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "ACB5050C.frx":0786
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   105
      X2              =   15165
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   135
      Left            =   12720
      TabIndex        =   9
      Top             =   195
      Width           =   255
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   105
      X2              =   15165
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "吨"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14790
      TabIndex        =   3
      Top             =   585
      Width           =   195
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   12240
      TabIndex        =   2
      Top             =   615
      Width           =   195
   End
End
Attribute VB_Name = "ACB5050C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACB5050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2007.8.12
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public STR1 As String
Public BASE As String
Public AIMNO As String
Dim sQuery As String

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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim SumCnt   As Integer
Dim SumCol   As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
     'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
      FormType = "Msheet"
         
           Call Gp_Ms_Collection(text_prod_cd, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(text_cur_inv_code, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_to_inv, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udate_in_plt_date_a, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udate_in_plt_date_b, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_mv_lst_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                                            
      'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
                                                      
    ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB5050C.P_REFER", Key:="P-R"
    sc1.Add Item:="ACB5050C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="ACB5050C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
                                                  
    ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
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
   
    'Spread_Collection
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=2, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc2, Key:="Sc2"

    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    SumCnt = 2
    
   ' Sum Column Setting
    SumCol.Add Item:=9
    SumCol.Add Item:=10
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 4, True)
    Call Gp_Sp_ColHidden(ss1, 15, True)
    
End Sub

Private Sub cmd_input_Click()

    Dim iDx     As Long
    Dim sMvNo   As String
    
    If ss1.MaxRows = 0 Then Exit Sub
    
    If Not Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_MsgBoxDisplay("你没有权限进行移库删除操作！！！")
        Exit Sub
    End If
    
    If Trim(txt_mv_lst_no.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_mv_lst_no.Tag & "必须输入")
        Exit Sub
    End If
    
    If Not IsDate(txt_input_date.Text) Then
        Call Gp_MsgBoxDisplay(txt_input_date.Tag & "必须输入")
        Exit Sub
    End If
    
    sMvNo = Trim(txt_mv_lst_no.Text)
    
    With ss1
        For iDx = 1 To .MaxRows
            .Row = iDx
            
            .Col = 2
            If sMvNo <> Trim(.Text) Then
                Call Gp_MsgBoxDisplay("移拨码单号不一样! 查询后处理一下..")
                Exit Sub
            End If
            
            .Col = 19:     .Text = Trim(txt_input_date.Text)
            .Col = 20:     .Text = sUserID
            .Col = 0:      .Text = "Update"
        Next iDx
    End With
    
End Sub

Private Sub Text_PROD_CD_Change()
   
    Select Case text_prod_cd.Text
        Case "S", "s", "SL"
            text_prod_cd.Text = "SL"
'        Case "P", "p", "PP"
'            text_prod_cd.Text = "PP"
'        Case "H", "h", "HC"
'            text_prod_cd.Text = "HC"
        Case ""
            text_prod_cd.Text = ""
        Case Else
            text_prod_cd.Text = ""
            Call MsgBox("产品分类代码" & Chr(10) & "不符合规范! 请更正。", vbExclamation + vbOKOnly, "警告")
    End Select

End Sub

Private Sub Text_PROD_CD_DblClick()

    Call Text_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_prod_cd_LostFocus()

    If text_prod_cd.Text <> "" Then
        If (Len(text_prod_cd.Text) < text_prod_cd.MaxLength) Then
            Call Gp_MsgBoxDisplay("产品分类代码输入未完成！")
            text_prod_cd.SetFocus
        End If
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

    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "C-System.INI", Me.Name)

    udate_in_plt_date_a.Text = Format(Date, "YYYY-MM-01")

    udate_in_plt_date_b.RawData = Gf_GetLastDay(udate_in_plt_date_b.RawData)
    
    Screen.MousePointer = vbDefault
    
    text_prod_cd.Text = "SL"
    text_cur_inv_code.Text = "00"
        
    If Gf_Sc_Authority(sAuthority, "U") Then
        cmd_input.Enabled = True
    Else
        cmd_input.Enabled = False
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '查询结束

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc2.Item("Spread"), "C-System.INI", Me.Name)
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
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
        
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
        
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Sc2 = Nothing
    Set Proc_Sc = Nothing
    Set SumCol = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gf_Sp_Cls(Proc_Sc("Sc2"))
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
    End If
 
    udate_in_plt_date_a.Text = Format(Date, "YYYY-MM-01")

    udate_in_plt_date_b.RawData = Gf_GetLastDay(udate_in_plt_date_b.RawData)
    text_tot_sheets.Value = 0
    text_tot_wgt.Value = 0
    txt_input_date.Text = ""
    text_prod_cd.Text = "SL"
    chk_Excel_Fl.Value = 0
    
End Sub

Public Sub Form_Exc()

    Dim iDR As Long

    If Trim(txt_mv_lst_no.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_mv_lst_no.Tag & "必须输入")
        Exit Sub
    End If
    
    If Trim(txt_to_inv_name.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_to_inv.Tag & "必须输入")
        Exit Sub
    End If
    
    Call ExcelPrn

'    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()
    Dim sQuery      As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    ss1.MaxRows = 0
    txt_input_date.Text = ""
    
    sQuery = "          Select   MV_LST_NO,"
    sQuery = sQuery & "          PROD_CD,"
    sQuery = sQuery & "          '',"           'APLY_STDSPEC
    sQuery = sQuery & "          FR_INV,"
    sQuery = sQuery & "          Gf_ComnNameFind('C0013',FR_INV),"
    sQuery = sQuery & "          TO_INV,"
    sQuery = sQuery & "          Gf_ComnNameFind('C0013',TO_INV),"
    sQuery = sQuery & "          MOVE_CAR_NO,"
    sQuery = sQuery & "          COUNT(*),"
    sQuery = sQuery & "          SUM(WGT),"
    sQuery = sQuery & "          DECODE(MAX(MOVE_DATE),NULL,NULL,MAX(SUBSTR(MOVE_DATE,1,4)||'-'||SUBSTR(MOVE_DATE,5,2)||'-'||SUBSTR(MOVE_DATE,7,2))),"
    sQuery = sQuery & "          GF_EMPNAMEFIND(MAX(MOVE_EMP)),"
    sQuery = sQuery & "          DECODE(MAX(RCV_DATE),NULL,NULL,MAX(SUBSTR(RCV_DATE,1,4)||'-'||SUBSTR(RCV_DATE,5,2)||'-'||SUBSTR(RCV_DATE,7,2))),"
    sQuery = sQuery & "          GF_EMPNAMEFIND(MAX(RCV_EMP))"
    sQuery = sQuery & "   FROM   CP_MOVE_SLT "
    sQuery = sQuery & "  WHERE   PROD_CD = '" & text_prod_cd.Text & "'"
    sQuery = sQuery & "    AND   NVL(FR_INV,' ')  LIKE '" & Trim(text_cur_inv_code.Text) + "%' "
    sQuery = sQuery & "    AND   NVL(TO_INV,' ')  LIKE '" & Trim(txt_to_inv.Text) & "%' "
    
    If Trim(txt_mv_lst_no.Text) <> "" Then
        sQuery = sQuery & "   AND NVL(MV_LST_NO,' ')  Like '" & Trim(txt_mv_lst_no.Text) & "%' "
    End If
    
    If IsDate(udate_in_plt_date_a.Text) Then
        sQuery = sQuery & "   AND MOVE_DATE >= " & udate_in_plt_date_a.RawData
    End If
    
    If IsDate(udate_in_plt_date_b.Text) Then
        sQuery = sQuery & "   AND MOVE_DATE <= " & udate_in_plt_date_b.RawData
    End If
    
    sQuery = sQuery & "   Group By MV_LST_NO,PROD_CD,FR_INV,TO_INV,MOVE_CAR_NO"
'    sQuery = sQuery & "   Group By MV_LST_NO,PROD_CD,APLY_STDSPEC,FR_INV,TO_INV,MOVE_CAR_NO"
    sQuery = sQuery & "   Order By MV_LST_NO DESC"
                    
    If Gf_Total_Display(M_CN1, Proc_Sc("Sc2"), sQuery, iDupCnt, SumCnt, SumCol) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss2.OperationMode = OperationModeNormal
    End If
    
    'Gp_Sp_ColHidden
    With ss2
        If .MaxRows = 0 Then
            text_tot_sheets.Text = "0"
            text_tot_wgt.Value = 0
        Else
            .Row = .MaxRows
            .Col = 9:  text_tot_sheets.Text = Val(.Value & "")
            .Col = 10: text_tot_wgt.Text = Val(.Value & "")
            .InsertRows 1, 1
            .Row = 1
            .Col = 1:   .Text = "  合  计 "
            .Col = 9:   .Text = text_tot_sheets.Text
            .Col = 10:  .Text = text_tot_wgt.Text
            Call Gp_Sp_BlockColor(Sc2.Item("Spread"), 1, .MaxCols, 1, 1, BLACK, &HE6E6FF)
        End If
    End With
End Sub

Public Sub Form_Pro()

    Dim iRow  As Long
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 0
        If ss1.Text = "Update" Then
            ss1.Col = 19
            If Not IsDate(ss1.Text) Then
                Call Gp_MsgBoxDisplay("到达日期必须输入")
                Exit Sub
            End If
        End If
    Next iRow
        
    Screen.MousePointer = vbHourglass
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        txt_input_date.Text = ""
    End If
       
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Del()

    Dim iRow  As Long
    
    Call Gp_Sp_Del(sc1)
    
    For iRow = 1 To ss1.MaxRows
        ss1.Row = iRow
        ss1.Col = 0
        If UCase(ss1.Text) = "DELETE" Then
            ss1.Col = 20
            ss1.Text = sUserID
        End If
    Next iRow
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub text_cur_inv_code_DblClick()

    Call text_cur_inv_code_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub text_cur_inv_code_Change()
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv_name.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
    Else
      text_cur_inv_name.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
         DD.sWitch = "MS"
         DD.sKey = "C0013"
    
         DD.rControl.Add Item:=text_cur_inv_code
    
         DD.nameType = "2"
         Call Gf_Common_DD(M_CN1, KeyCode)
    
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

Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    If Row < 1 Then Exit Sub
    
    ss2.Row = Row
    ss2.Col = 1
    If Len(Trim(ss2.Text)) > 10 Then
        txt_mv_lst_no.Text = ss2.Text
        ss2.Col = 4
        text_cur_inv_code.Text = ss2.Text
        ss2.Col = 6
        txt_to_inv.Text = ss2.Text
    Else
        txt_mv_lst_no.Text = ""
    End If
    
    ss2.Col = 9:  text_tot_sheets.Text = Val(ss2.Value & "")
    ss2.Col = 10: text_tot_wgt.Text = Val(ss2.Value & "")
    
    Call Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False)
    ss1.OperationMode = OperationModeNormal
    
End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)

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

Private Sub ss2_LostFocus()

'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
'   Call ss1_row_Click(Col, Row)

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)


    If Row <= 0 Then Exit Sub
    
    If Not Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_MsgBoxDisplay("你没有权限进行移库删除操作！！！")
        Exit Sub
    End If
    
    ss1.Row = Row
    ss1.Col = Col

    If Mode = 1 Then
        ss1.Tag = ss1.Text
    Else
        If Trim(ss1.Tag) <> Trim(ss1.Text) Then
            ss1.Col = 0
            Select Case Trim(ss1.Text)
                Case "Input", "Update", "Delete"
                Case Else
                    ss1.Text = "Update"
                    ss1.Col = 20:   ss1.Text = sUserID
            End Select
        End If
    End If
    
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

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
'        .Buttons(8).Enabled = False                 'Row Delete
'        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

End Sub

Private Function Gf_GetLastDay(Optional DTDay As String = "") As Variant

On Error GoTo DGet_Error

    Dim sQuery As String
    Dim strDay As String
    
    If DTDay = "" Then
        sQuery = "SELECT TO_CHAR(LAST_DAY(SYSDATE),'YYYYMMDD') FROM DUAL"
    Else
       strDay = DTDay
       sQuery = "SELECT TO_CHAR(LAST_DAY(TO_DATE('" + strDay + "','YYYYMMDD')),'YYYYMMDD') FROM DUAL"
    End If
       
    Dim AdoRs As ADODB.Recordset
    
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            If VarType(AdoRs.Fields(0)) = vbNull Then
                Gf_GetLastDay = ""
            Else
                Gf_GetLastDay = AdoRs.Fields(0)
            End If
        End If
        
    Else
        Gf_GetLastDay = "00000000"
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

DGet_Error:

    Set AdoRs = Nothing
    Gf_GetLastDay = "00000000"

End Function

Private Sub txt_input_date_Click()

    txt_input_date.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS') FROM DUAL")
    
End Sub

Private Sub txt_to_inv_DblClick()

    Call txt_to_inv_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_to_inv_Change()
    If Len(Trim(txt_to_inv.Text)) = txt_to_inv.MaxLength Then
        txt_to_inv_name.Text = Gf_ComnNameFind(M_CN1, "C0013", txt_to_inv.Text, 2)
    Else
      txt_to_inv_name.Text = ""
    End If
End Sub

Private Sub txt_to_inv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"
        DD.rControl.Add Item:=txt_to_inv
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
    End If
    
End Sub

Private Sub Text_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=text_prod_cd
        'DD.rControl.Add Item:=Text_PROD_CD_Name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If
    
End Sub

Private Sub udate_in_plt_date_a_LostFocus()
'    UDate_IN_PLT_DATE_b.RawData = Gf_GetLastDay(UDate_IN_PLT_DATE_a.RawData)
End Sub


Private Sub ExcelPrn()
    Dim I               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    Dim Wb              As Object
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If ERR.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    ERR.Clear

    Set Wb = xlApp.Workbooks.Open(App.Path & "\ACB5050C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    xlApp.Rows("5:200").Select
    xlApp.Selection.Delete Shift:=1
    
    xlApp.Sheets("Sheet2").Select
    xlApp.Range("A1:J1").Select
    xlApp.Selection.Copy
    xlApp.Sheets("Sheet1").Select
    sRow = "A" & 5 & ":" & "J" & ss1.MaxRows + 5
    xlApp.Range(sRow).Select
    xlApp.ActiveSheet.Paste
 
'    For I = 2 To ss1.MaxRows
'          xlApp.Rows("5:5").Select
'          xlApp.Selection.Copy
'          xlApp.Selection.Insert Shift:=1
'    Next I
            
    Select Case text_prod_cd.Text
        Case "PP"
            xlApp.Range("B3").Value = "钢板"
        Case "SL"
            xlApp.Range("B3").Value = "板坯"
        Case "HC"
            xlApp.Range("B3").Value = "钢卷"
    End Select
    
    xlApp.Range("D3").Value = Format(Date, "YYYY-MM-DD")
    xlApp.Range("G3").Value = txt_mv_lst_no.Text
    xlApp.Range("J3").Value = txt_to_inv_name.Text
          
    Clipboard.Clear
    ss1.SetSelection 3, 1, 3, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("A5").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 10, 1, 11, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("B5").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 13, 1, 14, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("D5").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 15, 1, 15, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("F5").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 5, 1, 8, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("G5").Select
    xlApp.ActiveSheet.Paste
 
    xlApp.Sheets("Sheet2").Select
    xlApp.Range("A2:J2").Select
    xlApp.Selection.Copy
    xlApp.Sheets("Sheet1").Select
    sRow = "A" & ss1.MaxRows + 5 & ":" & "J" & ss1.MaxRows + 5
    xlApp.Range(sRow).Select
    xlApp.ActiveSheet.Paste
    
'    sRow = ss1.MaxRows + 7 & ":" & 60 + ss1.MaxRows
'    xlApp.Rows(sRow).Select
'    xlApp.Selection.Delete Shift:=1
'
'    Clipboard.Clear
'    sRow = "A" & ss1.MaxRows + 5 & ":" & "J" & ss1.MaxRows + 5
'    xlApp.Range(sRow).Value = ""
'    sRow = "A" & ss1.MaxRows + 6 & ":" & "J" & ss1.MaxRows + 6
'    xlApp.Range(sRow).Value = ""

    Clipboard.Clear
    sRow = "E" & ss1.MaxRows + 5
    xlApp.Range(sRow).Value = "数量合计: " & text_tot_sheets.Text

    Clipboard.Clear
    sRow = "I" & ss1.MaxRows + 5
    xlApp.Range(sRow).Value = "总计:"

    sRow = "J" & ss1.MaxRows + 5
    xlApp.Range(sRow).Value = text_tot_wgt.Text

    sRow = "A" & ss1.MaxRows + 6
    ss2.Row = ss2.ActiveRow: ss2.Col = 12
    xlApp.Range(sRow).Value = "转库发货员姓名:" & ss2.Text

    sRow = "D" & ss1.MaxRows + 6
    xlApp.Range(sRow).Value = "仓库操作员工号:" & sUsername

    sRow = "H" & ss1.MaxRows + 6
    ss2.Row = ss2.ActiveRow: ss2.Col = 8
    xlApp.Range(sRow).Value = "车辆号:" & ss2.Text

    Clipboard.Clear
    xlApp.Range("A2").Select
    xlApp.ActiveSheet.Paste
    
    If chk_Excel_Fl = 0 Then
        xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    End If
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    If chk_Excel_Fl = 0 Then
        xlApp.Application.Visible = False
        Wb.Close False
        xlApp.QuitSet
        Set Wb = Nothing
        Set xlApp = Nothing
    Else
        xlApp.Application.Visible = True
    End If
    
'    Wb.Close
'    xlApp.Quit
    
'    Set Wb = Nothing
'    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set Wb = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Multi_Print_Click()
    Dim iDR     As Long
    Dim sFromNo As String
    Dim sToNo   As String
    
    If lBlkrow1 < 2 And lBlkrow2 < 2 Then Exit Sub
    
    ss2.Col = 1
    ss2.Row = lBlkrow1:     sFromNo = ss2.Text
    ss2.Row = lBlkrow2:     sToNo = ss2.Text
    
    If Not Gf_MessConfirm("您确定要码单多种打印(" & sFromNo & " ~ " & sToNo & ")吗？", "Q") Then Exit Sub
                   
    For iDR = lBlkrow1 To lBlkrow2
        If iDR > 1 Then
            Call ss2_DblClick(1, iDR)
            Call ExcelPrn
        End If
    Next iDR
        
End Sub

