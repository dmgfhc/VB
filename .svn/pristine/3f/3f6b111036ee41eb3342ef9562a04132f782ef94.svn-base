VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACE1152C 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "板卷替代余量分配/分配余量_ACE1152C"
   ClientHeight    =   7635
   ClientLeft      =   435
   ClientTop       =   3030
   ClientWidth     =   15180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   15180
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_slab_no 
      Height          =   270
      Left            =   8160
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   180
   End
   Begin Threed.SSCommand cmd_exit 
      Height          =   405
      Left            =   12780
      TabIndex        =   0
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "退出"
      BevelWidth      =   3
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   3
      Left            =   270
      Top             =   135
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "分配对象长度"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_rem_len 
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   135
      Width           =   840
      _Version        =   262145
      _ExtentX        =   1482
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
      FocusSelect     =   -1  'True
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
      NumIntDigits    =   7
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   5
      Left            =   4005
      Top             =   135
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      Caption         =   "分配对象重量"
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
   Begin CSTextLibCtl.sidbEdit sdb_slab_rem_wgt 
      Height          =   315
      Left            =   5625
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   135
      Width           =   840
      _Version        =   262145
      _ExtentX        =   1482
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
      FocusSelect     =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   7
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSCommand cmd_process 
      Height          =   405
      Left            =   8865
      TabIndex        =   3
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "分配处理"
      BevelWidth      =   3
   End
   Begin Threed.SSCommand cmd_cancel 
      Height          =   405
      Left            =   10035
      TabIndex        =   4
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "分配取消"
      BevelWidth      =   3
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_org_rem_wgt 
      Height          =   315
      Left            =   5670
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   45
      Visible         =   0   'False
      Width           =   840
      _Version        =   262145
      _ExtentX        =   1482
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   0
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
      FocusSelect     =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   7
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSCommand cmd_cnf 
      Height          =   405
      Left            =   11205
      TabIndex        =   7
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "分配确定"
      BevelWidth      =   3
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_rem_len1 
      Height          =   315
      Left            =   2730
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   135
      Width           =   840
      _Version        =   262145
      _ExtentX        =   1482
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
      FocusSelect     =   -1  'True
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
      NumIntDigits    =   7
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_org_rem_len 
      Height          =   315
      Left            =   1980
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   45
      Visible         =   0   'False
      Width           =   840
      _Version        =   262145
      _ExtentX        =   1482
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   0
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
      FocusSelect     =   -1  'True
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
      NumIntDigits    =   7
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_rem_wgt1 
      Height          =   315
      Left            =   6465
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   135
      Width           =   840
      _Version        =   262145
      _ExtentX        =   1482
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
      FocusSelect     =   -1  'True
      Modified        =   -1  'True
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
      NumIntDigits    =   7
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_thk 
      Height          =   315
      Left            =   90
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   45
      Visible         =   0   'False
      Width           =   840
      _Version        =   262145
      _ExtentX        =   1482
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
      FocusSelect     =   -1  'True
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
      NumIntDigits    =   7
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit sdb_slab_wid 
      Height          =   315
      Left            =   90
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   225
      Visible         =   0   'False
      Width           =   840
      _Version        =   262145
      _ExtentX        =   1482
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
      FocusSelect     =   -1  'True
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
      NumIntDigits    =   7
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   3525
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   13725
      _Version        =   393216
      _ExtentX        =   24209
      _ExtentY        =   6218
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ButtonDrawMode  =   4
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   11
      MaxRows         =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "ACE1152C.frx":0000
   End
End
Attribute VB_Name = "ACE1152C"
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
'-- Program ID        ACE1152C
'-- Designer          chen xiangxiang
'-- Coder             chen xiangxiang
'-- Date              2014.6.16
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------


Public FormType As String           'Form Type
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

 
Dim dSlab_Len As Double             '平均分配长度
Dim dSlab_Wgt As Double             '平均分配的重量

Dim sStlgrd As String



Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

'    MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"


    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", "i", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    

    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACE1152C.P_REFER", Key:="P-R"
    sc1.Add Item:="ACE1152C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"



    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0


End Sub


Private Sub Cmd_Cancel_Click()
Dim iRow As Integer

sdb_slab_rem_len1.Value = sdb_slab_rem_len.Value
sdb_slab_rem_wgt1.Value = sdb_slab_rem_wgt.Value
If (ss1.MaxRows <> 0) Then
  For iRow = 1 To ss1.MaxRows
  ss1.Col = 1
  ss1.Row = iRow
  ss1.Value = 0

  ss1.Col = 6
  ss1.Row = iRow
  ss1.Value = 0
  
  ss1.Col = 7
  ss1.Row = iRow
  ss1.Value = 0
  
    
  Next iRow
  
  
    dSlab_Len = 0                '平均分配的长度置空
    dSlab_Wgt = 0                '平均分配的重量置空

 
'    dSlab_Wgt_rem = 0            '切块重置空
'    dSlab_wid_rem = 0            '切块宽置空
End If

  

End Sub

'lbl_slab_c(i)的宽度，是根据分配到的长度加上原来lbl_slab_c(i)的宽度，如果lbl_slab_c(i)的宽度为0，则，改为加上板坯的长度。
'重量和宽度的计算方法一致

Private Sub cmd_cnf_Click()

    Dim iRow As Integer
    Dim iRlen As String
    Dim i As Integer
    Dim sdb_len As Double
    Dim SDB_WGT As Double
    
'    Call ss1_LeaveCell                          '直接点击分配确认的情况下，这么调用
    Call cmd_process_Click
    
    iRlen = ACE1150C.Shape5.Width
    
   For iRow = 1 To ss1.MaxRows
            
        ss1.Col = 4
        ss1.Row = iRow
        sdb_len = ss1.Value
        ss1.Col = 5
        ss1.Row = iRow
        SDB_WGT = ss1.Value
        
     If (ACE1150C.lbl_slab_c(iRow).ToolTipText = sdb_len) Then
        ss1.Col = 6
        ss1.Row = iRow
        sdb_len = sdb_len + ss1.Value

        ss1.Col = 7
        ss1.Row = iRow
        SDB_WGT = SDB_WGT + ss1.Value
    
     Else
        sdb_len = ACE1150C.lbl_slab_c(iRow).ToolTipText
        SDB_WGT = ACE1150C.lbl_slab_c(iRow).Tag
        ss1.Col = 6
        ss1.Row = iRow
        sdb_len = sdb_len + ss1.Value

        ss1.Col = 7
        ss1.Row = iRow
        SDB_WGT = SDB_WGT + ss1.Value

     End If

        ACE1150C.lbl_slab_c(iRow).Tag = Str(SDB_WGT)
        ACE1150C.lbl_slab_c(iRow).ToolTipText = Str(sdb_len)

        Call Plate_Slab_Update(iRow, sdb_len, SDB_WGT)
          
        ACE1150C.lbl_slab_c(iRow).Width = (ACE1150C.Shape5.Width) * (sdb_len) / (ACE1150C.sdb_slab_all_len.Value)
        
        If iRow = 1 Then
            ACE1150C.lbl_slab_c(iRow).Left = ACE1150C.Shape5.Left
        Else
            ACE1150C.lbl_slab_c(iRow).Left = ACE1150C.lbl_slab_c(iRow - 1).Left + ACE1150C.lbl_slab_c(iRow - 1).Width
        End If

      
   Next iRow
   
       
   
        ACE1150C.sdb_slab_rem_wgt.Value = sdb_slab_rem_wgt1.Value
        ACE1150C.sdb_slab_rem_len.Value = sdb_slab_rem_len1.Value
        
        cmd_process.Enabled = False
        cmd_cancel.Enabled = False
        cmd_cnf.Enabled = False
        
        



End Sub

Private Sub cmd_exit_Click()
       
   Unload Me

End Sub

'分配处理
'统计勾选的自动分配长度的板坯个数（iChk）
'不做比较，直接用板坯宽度减去轧件宽度，得到需要切掉的宽度(切块宽)，计算得到切块重。余下的就是和轧件相同宽度的板坯了
'手动分配，则板坯剩余长度减去手动分配的长度，将值赋给板坯剩余长度。鼠标离开事件计算手动分配的重量。
'这时，余下的被勾选的板坯，可以自动分配，按照iChk个数，板坯剩余长度除于iChk，平均分配。


Private Sub cmd_process_Click()
Dim iRow As Integer
Dim iChk As Integer                                       '统计有几个被选中
Dim iCol As Integer
Dim sdb_slab_thk As Double                                '可替代板坯的厚度
Dim sdb_slab_wid As Double                                '可替代板坯的宽度
Dim sdb_slab_wid_asoll As Double                          '轧件宽
Dim sdb_set_len As Double                                 '手动分配长度
Dim eql_fl As Boolean

sdb_slab_wid = ACE1150C.sdb_slab_wid1.Value               '可替代板坯的宽度
sdb_slab_thk = ACE1150C.sdb_slab_thk1.Value               '可替代板坯的厚度
'sdb_slab_wid_asoll = ACE1150C.sdb_slab_wid.Value          '轧件宽



iChk = 0

If (ss1.MaxRows <> 0) Then

   For iRow = 1 To ss1.MaxRows

        ss1.Col = 1
        ss1.Row = iRow

        If ss1.Value = 1 Then
           ss1.Col = 6
           ss1.Row = iRow
           If ss1.Value = 0 Then                          '只统计没有手动分配长度的被勾选的行，然后进行自动分配
           iChk = iChk + 1
           End If
        End If
    Next iRow
    
    
    
    
    '自动分配

    If iChk = 0 Then Exit Sub
        
        dSlab_Len = Gf_FloatFind(M_CN1, "SELECT TRUNC(" & sdb_slab_rem_len1.Value & "/" & iChk & ") FROM DUAL")        '每块能分到的长度
        
        
     For iRow = 1 To ss1.MaxRows
        
        ss1.Col = 1
        ss1.Row = iRow
      If ss1.Value = 1 Then
      
        ss1.Col = 10
        sdb_slab_wid_asoll = ss1.Value
        dSlab_Wgt = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & sStlgrd & "'," & sdb_slab_thk & "," & sdb_slab_wid_asoll & "," & dSlab_Len & ",0) FROM DUAL ")  '该轧件重
      
        ss1.Col = 6
        ss1.Row = iRow
        If ss1.Value = 0 Then                             '判断是否已经手动分配，没有则开始自动分配
                                                          '单块轧件取的不是平均值，而是余量，因为可能出现不能整除的情况，而导致丢失长度。最后一块取余量，这避免了长度丢失
          If iChk = 1 Then                                '单块轧件
            ss1.Col = 6
            ss1.Row = iRow
            ss1.Value = sdb_slab_rem_len1.Value
            sdb_slab_rem_len1.Value = 0                   '算剩余长度

            ss1.Col = 7
            ss1.Row = iRow
            ss1.Value = dSlab_Wgt
            sdb_slab_rem_wgt1.Value = sdb_slab_rem_wgt1.Value - dSlab_Wgt  '算剩余重量
            
          Else                                             '多块轧件的时候
           ss1.Col = 6
           ss1.Row = iRow
           ss1.Value = dSlab_Len
           sdb_slab_rem_len1.Value = sdb_slab_rem_len1.Value - dSlab_Len
           
           ss1.Col = 7
           ss1.Row = iRow
           ss1.Value = dSlab_Wgt
           sdb_slab_rem_wgt1.Value = sdb_slab_rem_wgt1.Value - ss1.Value
           
           iChk = iChk - 1


         End If
       End If

    End If

    Next iRow
    End If
    

     
'    cmd_cnf.Enabled = True

End Sub

'鼠标点击事件
Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim iRow As Integer
    Dim dSlab_Len As Double
    Dim dSlab_Wgt As Double

    If Not ChangeMade Then Exit Sub
    If Col <> 1 Then Exit Sub
    If Row <= 0 Then Exit Sub

    ss1.Row = Row
    ss1.Col = 1

    If Col = 1 And ss1.Value = 1 Then
        ss1.Col = 6
        ss1.Value = 0
        ss1.Col = 7
        ss1.Value = 0
        Call Gp_Sp_CellColor(ss1, 6, Row)
    End If

    For iRow = 1 To ss1.MaxRows

        ss1.Row = iRow
        ss1.Col = 1

        If ss1.Value = 1 Then
            ss1.Col = 6
            dSlab_Len = dSlab_Len + IIf(ss1.Value = 0, 0, ss1.Value)
            ss1.Col = 7
            dSlab_Wgt = dSlab_Wgt + IIf(ss1.Value = 0, 0, ss1.Value)
            Call Gp_Sp_BlockLock(ss1, 6, 6, iRow, iRow, False)
            Call Gp_Sp_CellColor(ss1, 6, iRow, , &HC0FFFF)
        Else
            ss1.Col = 6
            ss1.Value = 0
            ss1.Col = 7
            ss1.Value = 0
            Call Gp_Sp_BlockLock(ss1, 6, 6, iRow, iRow, True)
        End If

    Next iRow

    sdb_slab_rem_len1.Value = sdb_slab_rem_len.Value - dSlab_Len
    sdb_slab_rem_wgt1.Value = sdb_slab_rem_wgt.Value - dSlab_Wgt

End Sub

'鼠标离开 自动计算重量
Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim iRow As Integer
    Dim iCnt As Integer


    Dim dSlab_Len_set As Double                                                '自定义输入的长度
    Dim dSlab_Wgt_set As Double                                                '自定义分配的重量
    Dim sdb_slab_wid_asoll As Double                                           '轧件宽度
    Dim sdb_slab_thk As Double                                                 '可替代板坯的厚度

    sdb_slab_thk = ACE1150C.sdb_slab_thk1.Value                                '可替代板坯的厚度

'    sdb_slab_wid_asoll = ACE1150C.sdb_slab_wid.Value                           '轧件宽

    With ss1

        If .CellTag = "False" Then Exit Sub
        .Row = Row
        .Col = Col

        Select Case Col

            Case 6     'SLAB_LEN

                If IIf(.Value = 0, 0, .Value) <> 0 Then
                    dSlab_Len_set = .Value
                    'SLAB_WGT(ADD WGT)
                    .Col = 10
                    sdb_slab_wid_asoll = .Value
                    
                    dSlab_Wgt_set = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('WGT','" & sStlgrd & "'," & sdb_slab_thk & "," & sdb_slab_wid_asoll & "," & dSlab_Len_set & ",0) FROM DUAL ")
                    '这里直接根据长度，计算这块轧件能分配到重量。

                    ss1.Col = 7
                    ss1.Value = dSlab_Wgt_set

'                    End If

                Else
                    .Col = 6
                    .Value = 0
                    .Col = 7
                    .Value = 0
                    Call Gp_Sp_CellColor(ss1, 6, Row)
                End If

                dSlab_Len_set = 0                                                '参数置空，用作去做统计
                dSlab_Wgt_set = 0                                                '参数置空，用作去做统计

                For iRow = 1 To .MaxRows
                    .Row = iRow
                    .Col = Col
                    dSlab_Len_set = dSlab_Len_set + IIf(.Value = 0, 0, .Value)   '累计计算手动分配的总长度
                    .Col = 7
                    dSlab_Wgt_set = dSlab_Wgt_set + IIf(.Value = 0, 0, .Value)   '累计计算手动分配的总重量
                Next iRow

                If dSlab_Len_set = 0 Then
                    sdb_slab_rem_len1.Value = sdb_slab_rem_len.Value             '如果手动分配的长度为0，则剩余长度和剩余重量重置为原始值
                    sdb_slab_rem_wgt1.Value = sdb_slab_rem_wgt.Value

                    Exit Sub
                End If

                If Abs(dSlab_Len_set) > Round(Abs(sdb_slab_rem_len.Value), 3) Then

                    .Col = Col
                    .Row = Row
                    .CellTag = "False"

                    Call Gp_MsgBoxDisplay("已超过分配对象长度...!!")

                    .Col = Col
                    .Row = Row
                    .CellTag = ""
                    .Value = 0
                    .Col = 7
                    .Value = 0
                    .TabStop = True
                    .SetFocus
                    .SetActiveCell Col, Row
                    .Action = SS_ACTION_ACTIVE_CELL
                    .EditMode = True
                    .TabStop = False
'                    Call Gp_Sp_BlockLock(ss1, 11, 11, Row, Row, False)

                Else

                    sdb_slab_rem_len1.Value = sdb_slab_rem_len.Value - dSlab_Len_set     '分配后，剩余长度
                    sdb_slab_rem_wgt1.Value = sdb_slab_rem_wgt.Value - dSlab_Wgt_set     '分配后，剩余重量

                End If

        End Select

    End With
    
'    cmd_cnf.Enabled = True

End Sub

'更新数据
Private Sub Plate_Slab_Update(Current_Row As Variant, slab_chg_len As Double, slab_chg_wgt As Double)

On Error GoTo Process_Exec_ERROR

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    ss1.Row = Current_Row
    
    'SLAB_NO, BLOCK_SEQ, SEQ
    sQuery = "{call ACE1152C.P_MODIFY1 ('" + txt_slab_no.Text + "', "
    
    'SEQ
    ss1.Col = 8
    sQuery = sQuery + "'" + ss1.Text + "',"
    
    '变更的总长度
    sQuery = sQuery + "'" & slab_chg_len & "',"
    
     '变更的总重量
    sQuery = sQuery + "'" & slab_chg_wgt & "'"
    
    
   
    sQuery = sQuery + ",?,?)}"
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)

End Sub




Private Sub Form_Load()
    
    Screen.MousePointer = vbHourglass
    sAuthority = Gf_Pgm_Authority(Me.Name)
    Call Form_Define
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

    sdb_slab_rem_len.Value = ACE1150C.sdb_slab_rem_len.Value                              '将板坯剩余长度传给分配对象长度
    sdb_slab_rem_wgt.Value = ACE1150C.sdb_slab_rem_wgt.Value                              '将板坯剩余重量传给分配对象重量
    sdb_slab_rem_len1.Value = sdb_slab_rem_len.Value                                      '剩下的待分配对象长度
    sdb_slab_rem_wgt1.Value = sdb_slab_rem_wgt.Value                                      '剩下的待分配对象重量
    sStlgrd = ACE1150C.txt_ord_stlgrd1.Text                                               '钢种
    txt_slab_no.Text = ACE1150C.txt_slab_no.Text                                          '取得板坯号
  
   
    Call Form_Ref
    Call Cmd_Cancel_Click
    
    

End Sub


Private Sub Form_Ref()

    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl")) Then
        ss1.OperationMode = OperationModeNormal
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

    End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
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

    
End Sub



'Private Sub ss1_Advance(ByVal AdvanceNext As Boolean)
'
'    lBlkcol1 = BlockCol
'    lBlkcol2 = BlockCol2
'    lBlkrow1 = BlockRow
'    lBlkrow2 = BlockRow2
'End Sub


Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
