VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Begin VB.Form AAA1070C 
   Caption         =   "板坯计划录入_AAA1070C"
   ClientHeight    =   8130
   ClientLeft      =   555
   ClientTop       =   2550
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   13995
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7395
      Left            =   90
      TabIndex        =   14
      Top             =   1740
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   13044
      _Version        =   196609
      PaneTree        =   "AAA1070C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   3210
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   14985
         _Version        =   393216
         _ExtentX        =   26432
         _ExtentY        =   5662
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
         MaxCols         =   0
         MaxRows         =   0
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AAA1070C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4035
         Left            =   30
         TabIndex        =   16
         Top             =   3330
         Width           =   14985
         _Version        =   393216
         _ExtentX        =   26432
         _ExtentY        =   7117
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
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AAA1070C.frx":026B
      End
   End
   Begin VB.TextBox txt_excel 
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txt_aply_item 
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
      Left            =   1485
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "项目"
      Top             =   495
      Width           =   600
   End
   Begin VB.TextBox txt_aply_item_name 
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
      Left            =   2115
      TabIndex        =   3
      Top             =   495
      Width           =   3300
   End
   Begin VB.TextBox txt_prod_cd 
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
      Height          =   300
      Left            =   5985
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "产品"
      Text            =   "SL"
      Top             =   135
      Width           =   555
   End
   Begin VB.TextBox txt_stlgrd 
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
      Left            =   1485
      MaxLength       =   11
      TabIndex        =   4
      Tag             =   "钢种"
      Top             =   855
      Width           =   1410
   End
   Begin VB.TextBox txt_stlgrd_des 
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
      Left            =   2925
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "钢种"
      Top             =   855
      Width           =   6405
   End
   Begin InDate.ULabel ULabel6 
      Height          =   285
      Left            =   135
      Top             =   855
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      Caption         =   "钢种"
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
   Begin InDate.ULabel ULabel5 
      Height          =   300
      Left            =   4635
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "产品"
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
   Begin InDate.UDate dtp_date_str 
      Height          =   300
      Left            =   1485
      TabIndex        =   0
      Tag             =   "日期"
      Top             =   135
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      Text            =   "____-__"
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
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Left            =   135
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "日期"
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
   Begin CSTextLibCtl.sidbEdit sdb_buy 
      Height          =   330
      Left            =   4365
      TabIndex        =   6
      Top             =   1350
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel4 
      Height          =   330
      Left            =   5850
      Top             =   1350
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "外销坯"
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
   Begin InDate.ULabel ULabel3 
      Height          =   330
      Left            =   3015
      Top             =   1350
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "外购坯"
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
   Begin CSTextLibCtl.sidbEdit sdb_sale 
      Height          =   330
      Left            =   7200
      TabIndex        =   7
      Top             =   1350
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_send 
      Height          =   330
      Left            =   10080
      TabIndex        =   8
      Top             =   1350
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel2 
      Height          =   330
      Left            =   11565
      Top             =   1350
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "合计"
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
   Begin InDate.ULabel ULabel7 
      Height          =   330
      Left            =   8730
      Top             =   1350
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "厂内调拨坯"
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
   Begin CSTextLibCtl.sidbEdit sdb_sum 
      Height          =   330
      Left            =   12915
      TabIndex        =   9
      Top             =   1350
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   16711680
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
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel8 
      Height          =   300
      Left            =   135
      Top             =   495
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "项目"
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
   Begin CSTextLibCtl.sidbEdit sdb_oldsms 
      Height          =   330
      Left            =   1470
      TabIndex        =   10
      Top             =   1350
      Width           =   1275
      _Version        =   262145
      _ExtentX        =   2249
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   255
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
      BorderEffect    =   2
      DataProperty    =   2
      ReadOnly        =   -1  'True
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
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel9 
      Height          =   330
      Left            =   120
      Top             =   1350
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Caption         =   "老炼钢"
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
   Begin Threed.SSCommand SSCommand2 
      Height          =   330
      Left            =   11805
      TabIndex        =   11
      Top             =   135
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      Caption         =   "详细查询"
   End
   Begin Threed.SSCommand SCmd2 
      Height          =   330
      Left            =   13740
      TabIndex        =   12
      Top             =   135
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "上传Excel"
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   90
      X2              =   15100
      Y1              =   1260
      Y2              =   1260
   End
End
Attribute VB_Name = "AAA1070C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       production plan
'-- Sub_System Name
'-- Program Name
'-- Program ID        AAA1070C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
'-- Date              2004.4.14
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim THK_GRP As Collection
Dim WID_GRP As Collection
Dim MIN_VALUE As Collection
Dim MAX_VALUE As Collection

Dim arrValue As Variant

Private Sub Form_Define()
    
    Dim sQuery As String
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(dtp_date_str, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_aply_item, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_aply_item_name, " ", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_stlgrd_des, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     
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
    Sc1.Add Item:="AAA1070C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

    'Spread_Collection
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="AAA1070C.P_SREFER", Key:="P-R"
    Sc2.Add Item:="AAA1070C.P_UPLOAD", Key:="P-M"
    
    Proc_Sc.Add Item:=Sc2, Key:="Sc2"

    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Menu_Setting

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
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Sp_Setting1
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))

    Screen.MousePointer = vbDefault
    
    If Mid(sAuthority, 1, 3) = "111" Then
       SSCommand2.Enabled = True
       SCmd2.Enabled = True
    ElseIf Mid(sAuthority, 1, 1) = "1" Then
       SSCommand2.Enabled = True
       SCmd2.Enabled = False
    Else
       SSCommand2.Enabled = False
       SCmd2.Enabled = False
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set iColumn2 = Nothing
    Set aColumn2 = Nothing
    Set lColumn2 = Nothing
    
    Set THK_GRP = Nothing
    Set WID_GRP = Nothing
    Set MIN_VALUE = Nothing
    Set MAX_VALUE = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    rControl(1).SetFocus
    
    ss1.MaxCols = 0
    ss1.MaxRows = 0
    ss2.MaxRows = 0
    sdb_sum.Value = 0
    sdb_sale.Value = 0
    sdb_send.Value = 0
    sdb_buy.Value = 0
    sdb_oldsms.Value = 0

End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    Dim iCol As Integer
    Dim RowTot, ColTot As Double
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
        
        If Sp_Header_Refer() Then
        
            If Left(dtp_date_str.RawData, 6) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Then
               Call Gp_Sp_BlockLock(ss1, 1, -1, 1, -1, True)
            Else
                Call Gp_Sp_BlockLock(ss1, 1, -1, 1, -1, False)
                
                With ss1
                     For iCol = 1 To .MaxCols - 1 Step 2
                         Call Gp_Sp_BlockLock(ss1, iCol, iCol, 1, .MaxRows, True)
                     Next iCol
                End With
              
            End If
        
            If Sp_Data_Refer() Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Menu_Setting
                Call Gp_Ms_ControlLock(Mc1!lControl, True)
                ss1.SetFocus
                
                ss2.MaxRows = 0
                With ss1
                    .MaxRows = .MaxRows + 1
                    .MaxCols = .MaxCols + 1
                    '列合计
                    .Row = .MaxRows
                    .Col = SpreadHeader
                    .Text = "合计"
                    
                    For iCol = 1 To .MaxCols - 1
                        .Col = iCol
                        ColTot = Gf_Sp_ColSum(ss1, .Col, 1, .MaxRows - 1)
                        .Row = .MaxRows
                        If ColTot > 0 Then
                            .Text = ColTot
                        Else
                            .Text = ""
                        End If
                        .CellType = CellTypeNumber
                        '.TypeNumberDecPlaces = 3
                        .TypeNumberMax = 999999999
                        .TypeNumberMin = 0
                        .TypeNumberShowSep = True
                        .TypeHAlign = TypeHAlignRight
                        .TypeVAlign = TypeVAlignCenter
                    Next iCol
                    
                    '行合计
                    .Col = .MaxCols
                    .Row = SpreadHeader
                    .Text = "合计"
                    .Row = SpreadHeader + 1
                    .Text = "合计"
                    
                    .Col = .MaxCols:       .Col2 = .MaxCols
                    .Row = SpreadHeader:   .Row2 = SpreadHeader + 1
                    .ColMerge = MergeAlways
                    .RowMerge = MergeAlways
                    
                    For iCol = 1 To .MaxRows
                        .Row = iCol
                        RowTot = Gf_Sp_RowSum(ss1, .Row, 1, .MaxCols - 1)
                        .Col = .MaxCols
                        If RowTot > 0 Then
                            .Text = RowTot
                        Else
                            .Text = ""
                        End If
                        .CellType = CellTypeNumber
                        '.TypeNumberDecPlaces = 3
                        .TypeNumberMax = 999999999
                        .TypeNumberMin = 0
                        .TypeNumberShowSep = True
                        .TypeHAlign = TypeHAlignRight
                        .TypeVAlign = TypeVAlignCenter
                
                        .ColWidth(.Col) = 11
                    Next iCol
                    
                End With

            End If
        End If
            
    Else
        sMesg = sMesg + " 必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
End Sub

Public Sub Form_Pro()
Dim sInput As String
    ss2.Col = 0
    ss2.Row = 1
    sInput = ss2.Text

If sInput = "Input" Then '导入
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc2"), Mc1, True) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
        Call Gp_Sp_BlockLock(ss2, 10, 10, 1, ss2.MaxRows, True)
    End If
Else
    If Sp_Process(M_CN1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
        Call Form_Ref
    End If
End If
    
End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

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

Public Sub Form_Exc()
If txt_excel.Text = "1" Then
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
ElseIf txt_excel.Text = "2" Then
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End If
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
End Sub

Private Sub SCmd2_Click()
   Load frm_Excel
   frm_Excel.txt_load_file.Text = "AAA1070C"
   frm_Excel.Show 1
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    txt_excel.Text = "1"
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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
        MDIMain.Mnu_Sorting.Visible = False
        MDIMain.Line1.Visible = False
        
        PopupMenu MDIMain.PopUp_Spread
        
        MDIMain.Mnu_Sorting.Visible = True
        MDIMain.Line1.Visible = True
    End If

End Sub

Public Sub Sp_Setting1()

    With ss1

        .ColHeaderRows = 3
        .RowHeaderCols = 2
        .Col = -1
        .Row = SpreadHeader + 1
        .FontBold = True
        
        .RowHeight(SpreadHeader) = 12
        .RowHeight(SpreadHeader + 1) = 12
        
        .Row = SpreadHeader + 2
        
        .ColWidth(0) = 10
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
          
        .Row = SpreadHeader
        .Col = SpreadHeader
        .Text = "宽度组\厚度组"
        .Row = SpreadHeader + 1
        .Col = SpreadHeader
        .Text = "宽度组\厚度组"
        
        .Row = SpreadHeader + 2
        .RowHidden = True
        
        .Col = SpreadHeader + 1
        .ColHidden = True
        
    End With

End Sub

Public Sub Menu_Setting()

    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Row cancel
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Row cancel
    
End Sub

Public Function Sp_Header_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    Dim sQuery2 As String
    
    Dim AdoRs2 As ADODB.Recordset
    Dim ArrayRecords2 As Variant

    Set adoRs = New ADODB.Recordset
    
    sQuery = "SELECT THK_CD, FR_THK, TO_THK "
    sQuery = sQuery + "   FROM BP_THICK_GRP "
    sQuery = sQuery + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    sQuery = sQuery + "    AND THK_CD <> '*' "
    sQuery = sQuery + "  ORDER BY THK_CD "
    
    With ss1

        Sp_Header_Refer = True
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        adoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If adoRs.BOF Or adoRs.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            adoRs.Close
            Set adoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = adoRs.GetRows
        adoRs.Close
        Set adoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 2
        
            For iCol = 0 To .MaxCols - 1 Step 2
            
               .Col = iCol + 1
               .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                  
                .Col = iCol + 2
                .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                           
                .Col = iCol + 1:  .Row = SpreadHeader + 1:  .Text = "实绩"
                .Col = iCol + 2:  .Row = SpreadHeader + 1:  .Text = "计划"
                
                .Col = iCol + 1
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                .Col = iCol + 2
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                'Column Type Setting
                .Col = iCol + 1: .Col2 = iCol + 1
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                .ColWidth(iCol + 1) = 9
                
                .Col = iCol + 2: .Col2 = iCol + 2
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                .ColWidth(iCol + 2) = 9
                
                iCnt = iCnt + 1
                
            Next iCol
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Set AdoRs2 = New ADODB.Recordset
    
    sQuery2 = "SELECT WID_CD, FR_WID, TO_WID "
    sQuery2 = sQuery2 + "   FROM BP_WIDTH_GRP "
    sQuery2 = sQuery2 + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    sQuery2 = sQuery2 + "    AND WID_CD <> '*' "
    sQuery2 = sQuery2 + "  ORDER BY WID_CD "
    
    With ss1

        Sp_Header_Refer = True
        .ColWidth(0) = 15
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs2.Open sQuery2, M_CN1, adOpenKeyset
        
        If AdoRs2.BOF Or AdoRs2.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            AdoRs2.Close
            Set AdoRs2 = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords2 = AdoRs2.GetRows
        AdoRs2.Close
        Set AdoRs2 = Nothing

        If UBound(ArrayRecords2, 2) + 1 <> 0 Then
        
            .MaxRows = (UBound(ArrayRecords2, 2) + 1)
            iCnt = 0
            
            For iRow = 1 To .MaxRows
            
                .Row = iRow
                .Col = SpreadHeader
                
                If VarType(ArrayRecords2(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords2(1, iCnt)) & " ~ " & Trim(ArrayRecords2(2, iCnt)) & "mm"
                End If
                
                .Col = SpreadHeader + 1
                .Text = Trim(ArrayRecords2(0, iCnt))
                
                .Row = iRow + 2: .Row2 = iRow + 2
                .Col = 1: .Col2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                iCnt = iCnt + 1
            Next iRow
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    With ss1
    
        For iCol = 1 To .MaxCols - 1 Step 2
            Call Gp_Sp_BlockLock(ss1, iCol, iCol, 1, .MaxRows, True)
        Next iCol
        
        For iCol = 2 To .MaxCols Step 2
            .Col = iCol
            .Row = 1
            .Col2 = iCol
            .Row2 = .MaxRows
             If Trim(txt_prod_cd.Text) = "" Or Trim(txt_aply_item.Text) = "" Or Trim(txt_stlgrd.Text) = "" Then
'                .BlockMode = True
'                .Lock = True
'                .BlockMode = False
'                .Protect = True
                 Call Gp_Sp_BlockLock(ss1, iCol, iCol, 1, .MaxRows, True)
             Else
                .BlockMode = True
                .Lock = False
                .BackColor = &HC0FFFF
                .BlockMode = False
                .Protect = True
             End If
        Next iCol
        
    End With
    
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Set AdoRs2 = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
    Dim sWID_GRP As String
    Dim sTHK_GRP As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
          
    Call Slab_Wgt_Find
    sdb_sum.Value = sdb_sale.Value + sdb_send.Value - sdb_buy.Value + sdb_oldsms.Value
    

    Set adoRs = New ADODB.Recordset
    
    sEdate = Left(dtp_date_str.RawData, 6)
  
    sQuery = "SELECT WID_GRP, THK_GRP, sum(RST_WGT),sum(PLN_WGT)"
    sQuery = sQuery + "   FROM AP_SLAB_PLAN "
    sQuery = sQuery + "  WHERE YEAR_MONTH =      '" + sEdate + "' "
    sQuery = sQuery + "    AND APLY_ITEM  LIKE   '" + Trim(txt_aply_item.Text) + "%' "
    sQuery = sQuery + "    AND PROD_CD    LIKE   '" + Trim(txt_prod_cd.Text) + "%' "
    sQuery = sQuery + "    AND STLGRD     LIKE   '" + Trim(txt_stlgrd.Text) + "%' "
    sQuery = sQuery + "  GROUP BY WID_GRP, THK_GRP "
    sQuery = sQuery + "  ORDER BY WID_GRP, THK_GRP "
    
    With ss1

        Sp_Data_Refer = True
        
        .ReDraw = False
       ' .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        adoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If adoRs.BOF Or adoRs.EOF Then
        
            Sp_Data_Refer = False
            .ReDraw = True
            adoRs.Close
            Set adoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = adoRs.GetRows
        adoRs.Close
     '   Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
            iRow = 1
            For iCnt = 0 To UBound(ArrayRecords, 2)
                .Row = iRow
                .Col = SpreadHeader + 1
                 sWID_GRP = .Text
                 Do While iRow <= .MaxRows And sWID_GRP <> Trim(ArrayRecords(0, iCnt))
                    iRow = iRow + 1
                    .Row = iRow
                    sWID_GRP = .Text
                 Loop
                           
                 For iCol = 1 To .MaxCols - 1 Step 2
                    .Col = iCol
                    .Row = SpreadHeader + 2
                    sTHK_GRP = .Text

                    If sTHK_GRP = ArrayRecords(1, iCnt) Then
                        
                        .Row = iRow
                        If VarType(ArrayRecords(2, iCnt)) = vbNull Or ArrayRecords(2, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(2, iCnt))
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Or ArrayRecords(3, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(3, iCnt))
                        End If
                
                    End If

                Next iCol
                
            Next iCnt
            
        End If
  
        MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
        Screen.MousePointer = vbDefault
        
    End With
         
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim sMesg As String
    Dim sTemp As String
    Dim sPara As String
    
    Dim adoCmd As ADODB.Command

    Sp_Process = True
    
    If Trim(txt_prod_cd.Text) = "" Or Trim(txt_aply_item.Text) = "" Or Trim(txt_stlgrd.Text) = "" Then
       Sp_Process = False
       Call Gp_MsgBoxDisplay("can't save ...")
       Exit Function
    End If
    
    With ss1
    
        'MaxRow = 0 is Exit Function Or iCount = 0
        If .MaxRows < 1 Then
            Sp_Process = False
            Exit Function
        End If
        
        Screen.MousePointer = vbHourglass
        
        .ReDraw = False
        
        'Db Connection Check
        If Conn Is Nothing Then
            If GF_DbConnect = False Then Sp_Process = False: Exit Function
        End If
        
        'Ado Setting
        Conn.CursorLocation = adUseServer
        Set adoCmd = New ADODB.Command
        
        Set adoCmd.ActiveConnection = Conn
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = Sc.Item("P-M")
        
        Conn.BeginTrans
        
        'Ceate Parameter (Input) iType + iColumn
        For iCount = 1 To 8
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        For iRow = 1 To .MaxRows
            
            .Row = iRow
            
            'Parameters Setting
            For iCol = 2 To .MaxCols Step 2
            
                .Col = iCol
                If Trim(.Text) <> "" Then
                
                    .Row = SpreadHeader + 2
                    .Col = iCol
                    adoCmd.Parameters(4).Value = .Text     'thk_grp
              
                    .Row = iRow
                    .Col = SpreadHeader + 1
                    adoCmd.Parameters(5).Value = .Text     'wid_grp
                    
                    .Col = iCol
                 
                    If Trim(.Text) = "" Then               'plan_value
                        adoCmd.Parameters(6).Value = 0
                    Else
                        dTempInt = .Text
                        adoCmd.Parameters(6).Value = dTempInt
                    End If
                    
                    adoCmd.Parameters(7).Value = sUserID                            'User-id
                    
                    adoCmd.Parameters(0).Value = Mid(dtp_date_str.Text, 1, 4) + _
                                                 Mid(dtp_date_str.Text, 6, 2)        'YEAR_MONTH
                    adoCmd.Parameters(1).Value = txt_aply_item.Text                  'APLY_ITEM
                    adoCmd.Parameters(2).Value = txt_prod_cd.Text                    'PROD_CD
                    adoCmd.Parameters(3).Value = txt_stlgrd.Text                     'STLGRD
                                   
                    adoCmd.Execute
                    
                    'Error Check
                    If adoCmd("Error") <> "0" Then
               
                        ret_Result_ErrCode = adoCmd("Error")
                        ret_Result_ErrMsg = adoCmd("Messg")
                        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
               
                        Call Gp_MsgBoxDisplay(sErrMessg)
                        Screen.MousePointer = vbDefault
                        Set adoCmd = Nothing
                        Conn.RollbackTrans
                        Sp_Process = False
                        Exit Function
               
                     End If
                
                End If
            
            Next iCol
            
        Next iRow
        
        Conn.CommitTrans
        .ReDraw = True
        MDIMain.StatusBar1.Panels(1) = "提示信息:数据处理完成"
        Screen.MousePointer = vbDefault
        Exit Function
    
    End With

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("SpreadPro_Error : " & Error)

End Function

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
  txt_excel.Text = "2"
End Sub

Private Sub SSCommand2_Click()
    If dtp_date_str.RawData <> "" And Trim(txt_prod_cd.Text) <> "" Then
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           ss2.OperationMode = OperationModeNormal
        End If
    End If
End Sub

Private Sub txt_aply_item_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "A0001"
        
        DD.rControl.Add Item:=txt_aply_item
        DD.rControl.Add Item:=txt_aply_item_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        If txt_aply_item <> "011" And txt_aply_item <> "012" And txt_aply_item <> "013" And txt_aply_item <> "014" Then
            txt_aply_item.Text = ""
            txt_aply_item_name.Text = ""
        End If
        
        Exit Sub

    End If
    
    If Len(Trim(txt_aply_item)) = txt_aply_item.MaxLength Then
        txt_aply_item_name.Text = Gf_ComnNameFind(M_CN1, "A0001", Trim(txt_aply_item.Text), 2)
        If txt_aply_item <> "011" And txt_aply_item <> "012" And txt_aply_item <> "013" And txt_aply_item <> "014" Then
            txt_aply_item.Text = ""
            txt_aply_item_name.Text = ""
        End If
    Else
        txt_aply_item_name.Text = ""
    End If
    
End Sub

Private Sub txt_prod_cd_KeyPress(KeyAscii As Integer)

     KeyAscii = Asc(UCase(Chr(KeyAscii)))
     
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_stlgrd_des

        DD.nameType = "2"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_stlgrd)) = txt_stlgrd.MaxLength Then
        txt_stlgrd_des.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
    Else
        txt_stlgrd_des.Text = ""
    End If

End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)

    End If

End Sub

Public Sub Slab_Wgt_Find()

On Error GoTo Slab_Wgt_Find_Error

    Dim sQuery As String
    Dim adoRs As ADODB.Recordset
    Set adoRs = New ADODB.Recordset
    
    sQuery = "        SELECT SUM(A_011), SUM(A_012), SUM(A_013), SUM(A_014) "
    sQuery = sQuery + " FROM (SELECT DECODE(APLY_ITEM, '011',SUM(PLN_WGT),0) A_011, "
    sQuery = sQuery + "              DECODE(APLY_ITEM, '012',SUM(PLN_WGT),0) A_012, "
    sQuery = sQuery + "              DECODE(APLY_ITEM, '013',SUM(PLN_WGT),0) A_013, "
    sQuery = sQuery + "              DECODE(APLY_ITEM, '014',SUM(PLN_WGT),0) A_014  "
    sQuery = sQuery + "         FROM AP_SLAB_PLAN "
    sQuery = sQuery + "        WHERE YEAR_MONTH = '" + Left(dtp_date_str.RawData, 6) + "' "
    sQuery = sQuery + "          AND PROD_CD    = 'SL'"
    sQuery = sQuery + "  GROUP BY APLY_ITEM) "
      
      
    'Ado Execute
    adoRs.Open sQuery, M_CN1, adOpenKeyset

    If Not adoRs.BOF And Not adoRs.EOF Then
    
        If Not adoRs.EOF Then
            If VarType(adoRs.Fields(0)) = vbNull Then
                sdb_oldsms.Value = 0
            Else
                sdb_oldsms.Value = adoRs.Fields(0)
            End If
        
            If VarType(adoRs.Fields(1)) = vbNull Then
                sdb_buy.Value = 0
            Else
                sdb_buy.Value = adoRs.Fields(1)
            End If
            
            If VarType(adoRs.Fields(2)) = vbNull Then
                sdb_sale.Value = 0
            Else
                sdb_sale.Value = adoRs.Fields(2)
            End If
            
            If VarType(adoRs.Fields(3)) = vbNull Then
                sdb_send.Value = 0
            Else
                sdb_send.Value = adoRs.Fields(3)
            End If
        End If
        
    End If
    
    adoRs.Close
    Set adoRs = Nothing
    Exit Sub

Slab_Wgt_Find_Error:

    Set adoRs = Nothing
    Call Gp_MsgBoxDisplay("Slab_Wgt_Find_Error : " & Error)

End Sub
