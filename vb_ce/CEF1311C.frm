VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form CEF1311C 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "替代产品长度变更/分配余量_CEF1311C"
   ClientHeight    =   4125
   ClientLeft      =   615
   ClientTop       =   5565
   ClientWidth     =   13980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   13980
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread ss1 
      Height          =   3525
      Left            =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   540
      Width           =   13890
      _Version        =   393216
      _ExtentX        =   24500
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
      MaxCols         =   20
      MaxRows         =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "CEF1311C.frx":0000
   End
   Begin Threed.SSCommand cmd_exit 
      Height          =   405
      Left            =   12780
      TabIndex        =   1
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
      Left            =   4365
      Top             =   135
      Visible         =   0   'False
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
         Size            =   9.75
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
      Left            =   5985
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   135
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   5
      Left            =   360
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
         Size            =   9.75
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
      Left            =   1980
      TabIndex        =   3
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
      TabIndex        =   4
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
      TabIndex        =   5
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
      Left            =   2025
      TabIndex        =   7
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
      TabIndex        =   8
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
      Left            =   6825
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   135
      Visible         =   0   'False
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
      Left            =   6075
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
      Left            =   2820
      TabIndex        =   10
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
      Left            =   4185
      TabIndex        =   11
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
      Left            =   4185
      TabIndex        =   12
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
End
Attribute VB_Name = "CEF1311C"
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
'-- Program ID        CEF1311C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2010.12.16
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
Dim Select_Slab_Stlgrd As String

Dim Clear_Fl As Boolean
Dim ADD_Fl As Boolean

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Chg_Slab_Len As New Collection
Dim Add_Slab_Len As New Collection
Dim Chg_Slab_Wgt As New Collection
Dim Add_Slab_Wgt As New Collection

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    'Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
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
    sc1.Add Item:=ss1, Key:="Spread"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    'Call Gp_Sp_ColHidden(ss1, 1, True)
    
End Sub

Private Sub Cmd_Cancel_Click()

    Dim iRow As Integer
    
    If ADD_Fl Then Exit Sub
    
    sdb_slab_rem_len.Value = 0
    sdb_slab_rem_len1.Value = 0
    
    sdb_slab_rem_wgt.Value = sdb_slab_org_rem_wgt.Value
    sdb_slab_rem_wgt1.Value = sdb_slab_org_rem_wgt.Value
    
    For iRow = 1 To ss1.MaxRows
    
        ss1.Row = iRow
        
        If Clear_Fl Then
            ss1.Col = 1
            ss1.Value = 0
            Call Gp_Sp_BlockLock(ss1, 12, 12, iRow, iRow, True)
        Else
            ss1.Col = 1
            
            If ss1.Value = 0 Then
                Call Gp_Sp_BlockLock(ss1, 12, 12, iRow, iRow, True)
            Else
                Call Gp_Sp_BlockLock(ss1, 12, 12, iRow, iRow, False)
            End If
        End If
        
        ss1.Col = 11
        ss1.Text = ""
        ss1.Col = 12
        ss1.Text = ""
        ss1.Col = 14
        ss1.Text = ""
        ss1.Col = 17
        ss1.Text = ""
    
    Next iRow
    
End Sub

Private Sub cmd_cnf_Click()

    Dim iRow As Integer
    Dim dSlab_Len As Double
    Dim dSlab_Wgt As Double
    
    For iRow = 1 To ss1.MaxRows
    
        CEF1310C.Chg_Slab_Len.Remove (ss1.MaxRows - iRow + 1)
        CEF1310C.Chg_Slab_Wgt.Remove (ss1.MaxRows - iRow + 1)
        CEF1310C.Add_Slab_Len.Remove (ss1.MaxRows - iRow + 1)
        CEF1310C.Add_Slab_Wgt.Remove (ss1.MaxRows - iRow + 1)
        
    Next iRow
    
    For iRow = 1 To ss1.MaxRows
    
        ss1.Row = iRow
        ss1.Col = 1
        
        If ss1.Value = 1 Then
        
            ss1.Col = 11
            If IIf(ss1.Text = "", 0, ss1.Value) <> 0 Then
                
                ADD_Fl = True
                cmd_process.Enabled = False
                cmd_cancel.Enabled = False
                cmd_cnf.Enabled = False
                
                ss1.Col = 9
                dSlab_Len = ss1.Value
                
                ss1.Col = 10
                dSlab_Wgt = ss1.Value
                
                ss1.Col = 11
                dSlab_Len = dSlab_Len + ss1.Value
                Add_Slab_Len.Add Item:=ss1.Value
                
                ss1.Col = 12
                dSlab_Wgt = dSlab_Wgt + ss1.Value
                Add_Slab_Wgt.Add Item:=ss1.Value
                
                Chg_Slab_Len.Add Item:=dSlab_Len
                Chg_Slab_Wgt.Add Item:=dSlab_Wgt
                
                'CEF1066C.sdb_slab_rem_len.Value = sdb_slab_rem_len1.Value
                CEF1310C.sdb_div_wgt.Value = sdb_slab_rem_wgt1.Value
                CEF1310C.sdb_ord_wgt.Value = CEF1310C.sdb_mat_wgt.Value - sdb_slab_rem_wgt1.Value
                
            Else
                
                ss1.Col = 9
                Chg_Slab_Len.Add Item:=ss1.Value
                Add_Slab_Len.Add Item:=0
                
                ss1.Col = 10
                Chg_Slab_Wgt.Add Item:=ss1.Value
                Add_Slab_Wgt.Add Item:=0
                
            End If
            
        Else
        
            ss1.Col = 9
            Chg_Slab_Len.Add Item:=ss1.Value
            Add_Slab_Len.Add Item:=0
            
            ss1.Col = 10
            Chg_Slab_Wgt.Add Item:=ss1.Value
            Add_Slab_Wgt.Add Item:=0
            
        End If
        
        CEF1310C.Chg_Slab_Len.Add Item:=Chg_Slab_Len(iRow)
        CEF1310C.Add_Slab_Len.Add Item:=Add_Slab_Len(iRow)
        CEF1310C.Chg_Slab_Wgt.Add Item:=Chg_Slab_Wgt(iRow)
        CEF1310C.Add_Slab_Wgt.Add Item:=Add_Slab_Wgt(iRow)
        
        CEF1310C.lbl_slab(iRow).Width = (CEF1310C.Shape1.Width / CEF1310C.sdb_mat_wgt.Value) * (Chg_Slab_Wgt.Item(iRow) + IIf(iRow > 0, CEF1310C.dSlab_Cut_Spare_Wgt, 0))
        
        If iRow = 1 Then
            CEF1310C.lbl_slab(iRow).Left = CEF1310C.Shape1.Left
        Else
            CEF1310C.lbl_slab(iRow).Left = CEF1310C.lbl_slab(iRow - 1).Left + CEF1310C.lbl_slab(iRow - 1).Width
        End If
        
        'ToolTipText Setting
        CEF1310C.lbl_slab(iRow).ToolTipText = CEF1310C.Ord_No.Item(iRow) & "-" & CEF1310C.Ord_Item.Item(iRow) & "(" & CEF1310C.Slab_Len.Item(iRow) & ", " & CEF1310C.Slab_Wgt.Item(iRow) & ")"
    
    Next iRow
    
End Sub

Private Sub cmd_exit_Click()
    Unload Me
End Sub

Private Sub cmd_process_Click()

    Dim iChk As Integer
    Dim iRow As Integer
    Dim iLastRow As Integer
    Dim dSlab_Len As Double
    Dim dSlab_Wgt As Double
    Dim dSlab_uWgt As Double
    Dim dAsroll_Thk As Double
    Dim dAsroll_Wid As Double
    Dim dAsRoll_Len As Double
    Dim dProd_Cnt As Double
    Dim dProd_Wgt As Double
    Dim sStlgrd As String
    Dim sOrd_No As String
    Dim sOrd_item As String
    Dim sQuery As String
    
    If ADD_Fl Then Exit Sub
    
    Clear_Fl = False
    Call Cmd_Cancel_Click
    Clear_Fl = True
    
    For iRow = 1 To ss1.MaxRows
    
        ss1.Row = iRow
        ss1.Col = 1
        
        If ss1.Value = 1 Then
            iChk = iChk + 1
        End If
    
    Next iRow
    
    If iChk = 0 Then Exit Sub
    
    dSlab_uWgt = Gf_FloatFind(M_CN1, "SELECT TRUNC((" & Round(sdb_slab_rem_wgt.Value, 3) & "/" & iChk & "), 3) FROM DUAL")
    
    For iRow = 1 To ss1.MaxRows
    
        ss1.Row = iRow
        ss1.Col = 1
        
        If ss1.Value = 1 Then
        
            ss1.Col = 2
            sOrd_No = ss1.Text
            
            ss1.Col = 3
            sOrd_item = ss1.Text
            
            ss1.Col = 8
            dProd_Cnt = ss1.Text
            
            ss1.Col = 12
            ss1.Value = dSlab_uWgt
            sdb_slab_rem_wgt1.Value = sdb_slab_rem_wgt1.Value - dSlab_uWgt
            
            ss1.Col = 15
            sStlgrd = ss1.Text
            dSlab_Len = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('LEN','" & sStlgrd & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & ", 0, " & dSlab_uWgt & ") FROM DUAL ")
            
            ss1.Col = 11
            ss1.Value = dSlab_Len
            'sdb_slab_rem_LEN1.Value = sdb_slab_rem_LEN1.Value - dSlab_LEN
            
            'AsRoll Thk
            ss1.Col = 19
            dAsroll_Thk = ss1.Value
            
            'AsRoll Wid
            ss1.Col = 20
            dAsroll_Wid = ss1.Value

            ss1.Col = 10
            dSlab_Wgt = ss1.Value
            ss1.Col = 12
            dSlab_Wgt = dSlab_Wgt + ss1.Value
            dAsRoll_Len = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('LEN','" & sStlgrd & "'," & dAsroll_Thk & "," & dAsroll_Wid & ", 0, " & dSlab_Wgt & ") FROM DUAL ")
            ss1.Col = 14
            ss1.Text = dAsRoll_Len
            
            ss1.Col = 18
            dProd_Wgt = ss1.Text
            
            ss1.Col = 17
            ss1.Text = ((dProd_Cnt * dProd_Wgt) / dSlab_Wgt) * 100
            
            ss1.Col = 13
            If ss1.Value < dAsRoll_Len Then
                ss1.Col = 1
                ss1.Value = 0
                Call Gp_Sp_CellColor(ss1, 14, iRow, vbRed)
            Else
                Call Gp_Sp_CellColor(ss1, 14, iRow)
            End If
            
            iLastRow = iRow
            
        End If
    
    Next iRow
    
    If sdb_slab_rem_wgt1.Value <> 0 And iLastRow <> 0 Then
    
        ss1.Row = iLastRow
        ss1.Col = 12
        ss1.Value = ss1.Value + sdb_slab_rem_wgt1.Value
        dSlab_Wgt = ss1.Value
        dSlab_Len = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('LEN','" & sStlgrd & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & ", 0, " & dSlab_Wgt & ") FROM DUAL ")
        ss1.Col = 11
        ss1.Value = dSlab_Len
        
        sdb_slab_rem_len1.Value = 0
        sdb_slab_rem_wgt1.Value = 0
        
        ss1.Col = 10
        dSlab_Wgt = ss1.Value
        ss1.Col = 12
        dSlab_Wgt = dSlab_Wgt + ss1.Value
        dAsRoll_Len = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('LEN','" & sStlgrd & "'," & dAsroll_Thk & "," & dAsroll_Wid & ", 0, " & dSlab_Wgt & ") FROM DUAL ")
        ss1.Col = 14
        ss1.Text = dAsRoll_Len
        
        ss1.Col = 13
        If ss1.Value < dAsRoll_Len Then
            ss1.Col = 1
            ss1.Value = 0
            Call Gp_Sp_CellColor(ss1, 14, iRow, vbRed)
        Else
            Call Gp_Sp_CellColor(ss1, 14, iRow)
        End If
    
    End If
    
End Sub

Private Sub Form_Activate()
     
    Screen.MousePointer = vbHourglass
    
    Call Form_Define
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    ADD_Fl = False
    
    Call Gp_FormCenter(Me)
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

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
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    
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
    
    Set Chg_Slab_Len = Nothing
    Set Chg_Slab_Wgt = Nothing
    Set Add_Slab_Len = Nothing
    Set Add_Slab_Wgt = Nothing
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Ref()

End Sub

Public Sub Form_Pro()
    
End Sub

Public Sub Spread_ColumnsSort()

End Sub

Public Sub Spread_Forzens_Setting()
    
End Sub

Public Sub Spread_Forzens_Cancel()

End Sub

Public Sub Form_Exc()
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

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
        ss1.Col = 11
        ss1.Text = ""
        ss1.Col = 12
        ss1.Text = ""
        ss1.Col = 14
        ss1.Text = ""
        ss1.Col = 17
        ss1.Text = ""
        Call Gp_Sp_CellColor(ss1, 12, Row)
    End If
    
    For iRow = 1 To ss1.MaxRows
    
        ss1.Row = iRow
        ss1.Col = 1
        
        If ss1.Value = 1 Then
            ss1.Col = 11
            dSlab_Len = dSlab_Len + IIf(ss1.Text = "", 0, ss1.Value)
            ss1.Col = 12
            dSlab_Wgt = dSlab_Wgt + IIf(ss1.Text = "", 0, ss1.Value)
            Call Gp_Sp_BlockLock(ss1, 12, 12, iRow, iRow, False)
            Call Gp_Sp_CellColor(ss1, 12, iRow, , &HC0FFFF)
        Else
            ss1.Col = 11
            ss1.Text = ""
            ss1.Col = 12
            ss1.Text = ""
            ss1.Col = 14
            ss1.Text = ""
            ss1.Col = 17
            ss1.Text = ""
            Call Gp_Sp_BlockLock(ss1, 12, 12, iRow, iRow, True)
        End If
        
    Next iRow
    
    sdb_slab_rem_len1.Value = sdb_slab_org_rem_len.Value - dSlab_Len
    sdb_slab_rem_wgt1.Value = sdb_slab_org_rem_wgt.Value - dSlab_Wgt
    
End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim iRow As Integer
    Dim iCnt As Integer
    Dim dAsroll_Thk As Double
    Dim dAsroll_Wid As Double
    Dim dAsRoll_Len As Double
    Dim dProd_Wgt As Double
    Dim dSlab_Len As Double
    Dim dSlab_Wgt As Double
    Dim sStlgrd As String
    
    With ss1
    
        If .CellTag = "False" Then Exit Sub
              
        .Row = Row
        .Col = Col
        
        Select Case Col
        
            Case 12     'SLAB_WGT
                
                If IIf(.Text = "", 0, .Value) <> 0 Then
                    dSlab_Wgt = .Value
                    'SLAB_WGT(ADD WGT)
                    .Col = 15
                    sStlgrd = ss1.Text
                    dSlab_Len = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('LEN','" & sStlgrd & "'," & sdb_slab_thk.Value & "," & sdb_slab_wid.Value & ",0, " & dSlab_Wgt & ") FROM DUAL ")
                    .Col = 11
                    .Value = dSlab_Len
                    
                    'SLAB_WGT(TOTAL WGT)
                    .Col = 10
                    dSlab_Wgt = .Value + dSlab_Wgt
                    
                    'PROD_WGT
                    .Col = 18
                    dProd_Wgt = .Value
                    
                    'PROD_CNT
                    .Col = 8
                    iCnt = .Value
                    
                    'RATE
                    .Col = 17
                    .Value = ((iCnt * dProd_Wgt) / dSlab_Wgt) * 100
                    
                    'ASROLL_THK
                    .Col = 19
                    dAsroll_Thk = .Value
                    'ASROLL_WID
                    .Col = 20
                    dAsroll_Wid = .Value
                    
                    'ASROLL_LEN
                    dAsRoll_Len = Gf_FloatFind(M_CN1, "SELECT GF_JP_WGT('LEN','" & sStlgrd & "'," & dAsroll_Thk & "," & dAsroll_Wid & ", 0, " & dSlab_Wgt & ") FROM DUAL ")
                    ss1.Col = 14
                    .Value = dAsRoll_Len
                    
                    .Col = 13
                    If .Value < dAsRoll_Len Then
                        .Col = 1
                        .Value = 0
                        Call Gp_Sp_CellColor(ss1, 14, Row, vbRed)
                    Else
                        Call Gp_Sp_CellColor(ss1, 14, Row)
                    End If
                    
                Else
                    .Col = 11
                    .Text = ""
                    .Col = 12
                    .Text = ""
                    .Col = 14
                    .Text = ""
                    .Col = 17
                    .Text = ""
                    Call Gp_Sp_CellColor(ss1, 14, Row)
                End If
                
                dSlab_Len = 0
                dSlab_Wgt = 0
                
                For iRow = 1 To .MaxRows
                    .Row = iRow
                    .Col = 11
                    dSlab_Len = dSlab_Len + IIf(.Text = "", 0, .Value)
                    .Col = Col
                    dSlab_Wgt = dSlab_Wgt + IIf(.Text = "", 0, .Value)
                Next iRow
                    
                If dSlab_Wgt = 0 Then
                    sdb_slab_rem_len.Value = 0
                    sdb_slab_rem_len1.Value = 0
                    sdb_slab_rem_wgt.Value = sdb_slab_org_rem_wgt.Value
                    sdb_slab_rem_wgt1.Value = sdb_slab_org_rem_wgt.Value
                    Exit Sub
                End If
                
                If Abs(dSlab_Wgt) > Round(Abs(sdb_slab_rem_wgt.Value), 3) Then
                
                    .Col = Col
                    .Row = Row
                    .CellTag = "False"
                    
                    Call Gp_MsgBoxDisplay("已超过分配对象重量...!!")
                    
                    .Col = Col
                    .Row = Row
                    .CellTag = ""
                    .Text = ""
                    .Col = 11
                    .Text = ""
                    .Col = 14
                    .Text = ""
                    .Col = 17
                    .Text = ""
                    .TabStop = True
                    .SetFocus
                    .SetActiveCell Col, Row
                    .Action = SS_ACTION_ACTIVE_CELL
                    .EditMode = True
                    .TabStop = False
                    Call Gp_Sp_BlockLock(ss1, 12, 12, Row, Row, False)
                
                Else
            
                    sdb_slab_rem_len1.Value = sdb_slab_org_rem_len.Value - dSlab_Len
                    sdb_slab_rem_wgt1.Value = sdb_slab_org_rem_wgt.Value - dSlab_Wgt
                
                End If
                        
        End Select
            
    End With

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
