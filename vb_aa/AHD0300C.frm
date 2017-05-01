VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1090C 
   Caption         =   "月生产计划详细查询_AAA1090C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   630
      Left            =   120
      TabIndex        =   2
      Top             =   2355
      Width           =   15060
      Begin CSTextLibCtl.sidbEdit sidbEdit4 
         Height          =   345
         Left            =   11835
         TabIndex        =   11
         Top             =   195
         Width           =   1695
         _Version        =   262145
         _ExtentX        =   2990
         _ExtentY        =   609
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   2
         StartText.y     =   5
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   13
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sidbEdit3 
         Height          =   360
         Left            =   8430
         TabIndex        =   9
         Top             =   195
         Width           =   1575
         _Version        =   262145
         _ExtentX        =   2778
         _ExtentY        =   635
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   2
         StartText.y     =   5
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   13
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sidbEdit2 
         Height          =   345
         Left            =   5325
         TabIndex        =   7
         Top             =   195
         Width           =   1605
         _Version        =   262145
         _ExtentX        =   2831
         _ExtentY        =   609
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   2
         StartText.y     =   5
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   13
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         Undo            =   0
         Data            =   0
      End
      Begin CSTextLibCtl.sidbEdit sidbEdit1 
         Height          =   360
         Left            =   2370
         TabIndex        =   5
         Top             =   180
         Width           =   1485
         _Version        =   262145
         _ExtentX        =   2619
         _ExtentY        =   635
         _StockProps     =   125
         Text            =   " 0.00"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DataProperty    =   2
         Modified        =   0   'False
         HideSelection   =   -1  'True
         RawData         =   "0.00"
         Text            =   " 0.00"
         StartText.x     =   2
         StartText.y     =   5
         FirstVisPos     =   0
         HiAnchor        =   0
         HiNew           =   0
         CaretHeight     =   13
         CurNumDataChars =   0
         MaxDataChars    =   0
         FirstDataPos    =   0
         CurPos          =   0
         MaxLen          =   0
         DataReadOnly    =   0   'False
         Mask            =   ""
         Justification   =   2
         Undo            =   0
         Data            =   0
      End
      Begin VB.Label Label5 
         Caption         =   "外卖材"
         Height          =   255
         Left            =   11100
         TabIndex        =   10
         Top             =   285
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "外购材"
         Height          =   330
         Left            =   7665
         TabIndex        =   8
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "转库"
         Height          =   240
         Left            =   4605
         TabIndex        =   6
         Top             =   255
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "合计"
         Height          =   240
         Left            =   1905
         TabIndex        =   4
         Top             =   255
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "计划板坯使用量"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   270
         Width           =   1470
      End
   End
   Begin InDate.UDate dtp_yy_mm 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Tag             =   "年份月报"
      Top             =   120
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Text            =   "____-__"
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
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   135
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "年份月报"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   1845
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   495
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   3254
      _StockProps     =   64
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
      MaxCols         =   11
      MaxRows         =   3
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AHD0300C.frx":0000
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   1890
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3060
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   3334
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
      MaxCols         =   11
      MaxRows         =   3
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AHD0300C.frx":075F
   End
   Begin FPSpread.vaSpread ss3 
      Height          =   4335
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5025
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   7646
      _StockProps     =   64
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
      MaxCols         =   11
      MaxRows         =   9
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AHD0300C.frx":0F42
   End
End
Attribute VB_Name = "AAA1090C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Order Management System
'-- Sub_System Name
'-- Program Name
'-- Program ID        ABY1010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          ZHANGLIN
'-- Coder             ZHANGLIN
'-- Date              2005.8.19
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
    Dim sQuery As String

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(dtp_yy_mm, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'         Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'    Call Gp_Ms_Collection(txt_prod_cd_name, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 8, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", " ", " ", " ", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ABY1010C.P_MODIFY", Key:="P-M"
  '  Sc1.Add Item:="ABX1090C.P_REFER", Key:="P-R"
   ' Sc1.Add Item:="ABX1090C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    Proc_Sc.Add Item:=sc1, Key:="Sc1"

    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="ABY1010C.P_MODIFY", Key:="P-M"
  '  Sc2.Add Item:="ABX1090C.P_REFER", Key:="P-R"
   ' Sc2.Add Item:="ABX1090C.P_ONEROW", Key:="P-O"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"

    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="ABY1010C.P_MODIFY", Key:="P-M"
  '  Sc3.Add Item:="ABX1090C.P_REFER", Key:="P-R"
   ' Sc3.Add Item:="ABX1090C.P_ONEROW", Key:="P-O"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"


    Proc_Sc.Add Item:=sc3, Key:="Sc3"


    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.MenuTool.Buttons(4).Enabled = False
    
'    Dim iCol As Integer
'
'    With ss1
''        .Protect = True
''        .BlockMode = False
'        .Row = 1
'        For iCol = 1 To 9
'            .Col = iCol
'            .Lock = True
'        Next iCol
'
'        .Row = 2
'        For iCol = 1 To 9
'            .Col = iCol
'            .Lock = True
'        Next iCol
'
'        .Row = 5
'        For iCol = 1 To 9
'            .Col = iCol
'            .Lock = True
'        Next iCol
'
'    End With
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

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "B-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "B-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "B-System.INI", Me.Name)
    
    Call Sp_Setting1
'    Call Sp_Setting2
'    Call Sp_Setting3
    
    Screen.MousePointer = vbDefault
    MDIMain.MenuTool.Buttons(4).Enabled = False
'    userid.Text = sUserID


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
'        Cancel = 1
'        Exit Sub
'    End If
'
'    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)

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

    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("SC1")) And Gf_Sp_Cls(Proc_Sc("SC2")) And Gf_Sp_Cls(Proc_Sc("SC3")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), True)
'        rControl(1).SetFocus
        MDIMain.MenuTool.Buttons(4).Enabled = False
    End If
    
    Call Sp_Setting1
'    Call Sp_Setting2
'    Call Sp_Setting3
    
    dtp_yy_mm.Enabled = True
'    txt_prod_cd.Enabled = True
    dtp_yy_mm.RawData = ""
'    txt_prod_cd_name.Text = ""
    
'    Dim iCol As Integer
    
'    With ss1
''        .Protect = True
''        .BlockMode = False
'        .Row = 1
'        For iCol = 1 To 9
'            .Col = iCol
'            .Lock = True
'        Next iCol
'
'        .Row = 2
'        For iCol = 1 To 9
'            .Col = iCol
'            .Lock = True
'        Next iCol
'
'        .Row = 5
'        For iCol = 1 To 9
'            .Col = iCol
'            .Lock = True
'        Next iCol
'
'    End With
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    Dim sEdate1, sEdate2 As String
    Dim iRow, iCol As Integer
    Dim i As Integer
    Dim DATA1 As Double
    Dim DATA2 As Double
    Dim DATA3 As Double


'    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    sMesg = Gf_Ms_NeceCheck(nControl)
    
    sEdate1 = Mid(dtp_YEAR_MONTH.Text, 1, 4) + Mid(dtp_YEAR_MONTH.Text, 6, 2)
'    sEdate2 = Mid(dtp_TODAY.Text, 1, 4) + Mid(dtp_TODAY.Text, 6, 2)
'    If Val(sEdate1) < Val(sEdate2) Then
'
'        With ss1
''            .Protect = True
''            .BlockMode = False
'            For iRow = 1 To 7
'               .Row = iRow
'                For iCol = 1 To 9
'                    .Col = iCol
'                    .Lock = True
'                Next iCol
'            Next iRow
'        End With
'
'        MDIMain.MenuTool.Buttons(4).Enabled = False
'
'    End If
    
    If sMesg = "OK" Then

        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then

'            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1) Then
              If Sp_Data_Refer1() Then
                 If Sp_Data_Refer2() Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
                    dtp_YEAR_MONTH.Enabled = False
'                    txt_prod_cd.Enabled = False
                    MDIMain.MenuTool.Buttons(4).Enabled = True
                    
                   
                    For i = 1 To 9            '总计
                        ss1.Col = i
                        ss1.Row = 2
                        If ss1.Value = 0 Then
                           DATA2 = 0
                        Else
                           DATA2 = ss1.Value
                        End If
                        
                        ss1.Row = 5
                        If ss1.Value = 0 Then
                           DATA3 = 0
                        Else
                           DATA3 = ss1.Value
                        End If
                        
                        DATA1 = DATA2 + DATA3
                        ss1.Row = 1: ss1.Value = DATA1
                    Next i

                    Exit Sub
                  End If
              End If
        Else
            sMesg = sMesg + " 必须按项目长度输入"
            Call Gp_MsgBoxDisplay(sMesg)
        End If

    Else
        sMesg = sMesg + " 必须输入"
        Call Gp_MsgBoxDisplay(sMesg)

    End If
    
    MDIMain.MenuTool.Buttons(4).Enabled = True
    dtp_YEAR_MONTH.Enabled = False
'    txt_prod_cd.Enabled = False
    Exit Sub
Refer_Err:

End Sub
Public Sub Form_Pro()

    Dim iRow  As Integer
    Dim iTEXT As String

    If Sp_Process1(M_CN1, Proc_Sc("SC")) Then
       If Sp_Process2(M_CN1, Proc_Sc("SC")) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                 For iRow = 1 To 7
                     ss1.Row = iRow
                     ss1.Col = SpreadHeader
                     iTEXT = ss1.Text
                         If iTEXT = "Update" Then
                         ss1.Text = ""
                         End If
                 Next iRow
        End If
    End If

'    Call Sp_Data_Refer1
'    Call Sp_Data_Refer2
    Call Form_Ref
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

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

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), SpreadHeader + 1, 9, SpreadHeader, 3)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub


Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub

    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub

    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
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
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub


Public Function Sp_Data_Refer2() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sTdate As String
    Dim sQuery As String
    Dim sEdate1 As String
    Dim iEdate As Integer
    Dim sUnit_kind As String
    Dim sTHK_GRP As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim IWGT As Double
    Dim i As Integer
    Dim DATA1 As Double
    Dim DATA2 As Double
    Dim DATA3 As Double

    Set AdoRs = New ADODB.Recordset

'    sEdate = Mid(dtp_yy_mm.Text, 1, 4)
    sEdate1 = Mid(dtp_YEAR_MONTH.Text, 1, 4) + Mid(dtp_YEAR_MONTH.Text, 6, 2)

'    sTdate = Mid(dtp_yy_mm.Text, 6, 2)
'    iEdate = Val(sTdate)

    sQuery = "SELECT "
    sQuery = sQuery + " PROD_CLS_CD , PROD_CD , SALE_WAY , PLN_WGT "
    sQuery = sQuery + " FROM BP_SALE_PLN_WGT"
    sQuery = sQuery + " WHERE YYYYMM     = '" + sEdate1 + "' "
    sQuery = sQuery + "   AND PROD_CD = 'HC' "
'    sQuery = sQuery + "   AND STDSPEC = '" + Trim(txt_stdspec.Text) + "' "

    With ss1

        Sp_Data_Refer2 = True

        .ReDraw = False

        Screen.MousePointer = vbHourglass

        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset

        If AdoRs.BOF Or AdoRs.EOF Then

            Sp_Data_Refer2 = False
            .ReDraw = True

            AdoRs.Close
            Set AdoRs = Nothing

            Screen.MousePointer = vbDefault
            Exit Function
        End If

        ArrayRecords = AdoRs.GetRows

        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 1) + 1 <> 0 Then
            For iCnt = 0 To UBound(ArrayRecords, 2)

                If Trim(ArrayRecords(0, iCnt)) = "PB" Then
                   iRow = 3
                Else
                   iRow = 4
                End If



                 If Trim(ArrayRecords(2, iCnt)) = "ZL" Then
                    iCol = 1
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "CZ" Then
                    iCol = 2
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "CK" Then
                    iCol = 3
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GX" Then
                    iCol = 4
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GH" Then
                    iCol = 5
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GJ" Then
                    iCol = 6
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GN" Then
                    iCol = 7
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GB" Then
                    iCol = 8
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GZ" Then
                    iCol = 9
                 End If

                If Val((ArrayRecords(3, iCnt))) = 0 Then
'                If VarType(ArrayRecords(3, iCnt)) = vbNull Or Val((ArrayRecords(3, iCnt))) = 0 Then
                    IWGT = 0
                Else
                     If Val((ArrayRecords(3, iCnt))) <> 0 Then
                        IWGT = ArrayRecords(3, iCnt)
                    End If
                End If

                .Row = iRow
                .Col = iCol
                .Value = IWGT

            Next iCnt

        End If

        Screen.MousePointer = vbDefault

    End With

'合计
For i = 1 To 9
    ss1.Col = i
    ss1.Row = 3

    If Val(ss1.Value) = 0 Then
'    If ss1.Value = "" Or Val(ss1.Value) = 0 Then
       DATA2 = 0
    Else
       DATA2 = ss1.Value
    End If

    ss1.Row = 4
    If Val(ss1.Value) = 0 Then
'    If ss1.Value = "" Or Val(ss1.Value) = 0 Then
       DATA3 = 0
    Else
       DATA3 = ss1.Value
    End If

    DATA1 = DATA2 + DATA3
    ss1.Row = 2: ss1.Value = DATA1
Next i



Exit Function

SpreadDisplay_Error:

    Set AdoRs = Nothing
    Sp_Data_Refer2 = False
    Screen.MousePointer = vbDefault

End Function

Public Function Sp_Data_Refer1() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sTdate As String
    Dim sQuery As String
    Dim sEdate1 As String
    Dim iEdate As Integer
    Dim sUnit_kind As String
    Dim sTHK_GRP As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim IWGT As Double
    Dim i As Integer
    Dim DATA1 As Double
    Dim DATA2 As Double
    Dim DATA3 As Double

    Set AdoRs = New ADODB.Recordset

'    sEdate = Mid(dtp_yy_mm.Text, 1, 4)
    sEdate1 = Mid(dtp_YEAR_MONTH.Text, 1, 4) + Mid(dtp_YEAR_MONTH.Text, 6, 2)

'    sTdate = Mid(dtp_yy_mm.Text, 6, 2)
'    iEdate = Val(sTdate)

    sQuery = "SELECT "
    sQuery = sQuery + " PROD_CLS_CD , PROD_CD , SALE_WAY , PLN_WGT "
    sQuery = sQuery + " FROM BP_SALE_PLN_WGT"
    sQuery = sQuery + " WHERE YYYYMM     = '" + sEdate1 + "' "
    sQuery = sQuery + "   AND PROD_CD = 'PP' "
'    sQuery = sQuery + "   AND STDSPEC = '" + Trim(txt_stdspec.Text) + "' "

    With ss1

        Sp_Data_Refer1 = True

        .ReDraw = False

        Screen.MousePointer = vbHourglass

        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset

        If AdoRs.BOF Or AdoRs.EOF Then

            Sp_Data_Refer1 = False
            .ReDraw = True

            AdoRs.Close
            Set AdoRs = Nothing

            Screen.MousePointer = vbDefault
            Exit Function
        End If

        ArrayRecords = AdoRs.GetRows

        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 1) + 1 <> 0 Then
            For iCnt = 0 To UBound(ArrayRecords, 2)

                If Trim(ArrayRecords(0, iCnt)) = "PB" Then
                   iRow = 6
                Else
                   iRow = 7
                End If



                 If Trim(ArrayRecords(2, iCnt)) = "ZL" Then
                    iCol = 1
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "CZ" Then
                    iCol = 2
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "CK" Then
                    iCol = 3
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GX" Then
                    iCol = 4
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GH" Then
                    iCol = 5
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GJ" Then
                    iCol = 6
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GN" Then
                    iCol = 7
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GB" Then
                    iCol = 8
                 ElseIf Trim(ArrayRecords(2, iCnt)) = "GZ" Then
                    iCol = 9
                 End If

                If Val((ArrayRecords(3, iCnt))) = 0 Then
'                If VarType(ArrayRecords(3, iCnt)) = vbNull Or Val((ArrayRecords(3, iCnt))) = 0 Then
                    IWGT = 0
                Else
                     If Val((ArrayRecords(3, iCnt))) <> 0 Then
                        IWGT = ArrayRecords(3, iCnt)
                    End If
                End If

                .Row = iRow
                .Col = iCol
                .Value = IWGT

            Next iCnt

        End If

        Screen.MousePointer = vbDefault

    End With

'合计
For i = 1 To 9
    ss1.Col = i
    ss1.Row = 6

    If Val(ss1.Value) = 0 Then
'    If ss1.Value = "" Or Val(ss1.Value) = 0 Then
       DATA2 = 0
    Else
       DATA2 = ss1.Value
    End If

    ss1.Row = 7
    If Val(ss1.Value) = 0 Then
'    If ss1.Value = "" Or Val(ss1.Value) = 0 Then
       DATA3 = 0
    Else
       DATA3 = ss1.Value
    End If

    DATA1 = DATA2 + DATA3
    ss1.Row = 5: ss1.Value = DATA1
Next i



Exit Function

SpreadDisplay_Error:

    Set AdoRs = Nothing
    Sp_Data_Refer1 = False
    Screen.MousePointer = vbDefault

End Function
Public Sub Sp_Setting1()

    With ss1

        .MaxRows = 3
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        ss1.Col = SpreadHeader + 1
        
        ss1.Row = 1
        ss1.Text = "钢卷"
        
        ss1.Row = 2
        ss1.Text = "钢板"
        
        ss1.Row = 3
        ss1.Text = "合计"
        
    End With
    
    
    With ss2
        
        .MaxRows = 3
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        ss2.Col = SpreadHeader + 2
        ss2.Row = SpreadHeader
        ss2.Text = ""
        ss2.Col = SpreadHeader + 1
        ss2.Text = ""
        Call ss2.AddCellSpan(SpreadHeader + 1, SpreadHeader, 2, 1)

        
        ss2.Col = SpreadHeader + 1
        ss2.Row = 1
        ss2.Text = "连铸机"
        ss2.Row = 2
        ss2.Text = "连铸机"
        Call ss2.AddCellSpan(SpreadHeader + 1, 1, 1, 2)

        ss2.Col = SpreadHeader + 2
        
        ss2.Row = 1
        ss2.Text = "1#"
        
        ss2.Row = 2
        ss2.Text = "2#"
        
        ss2.Col = SpreadHeader + 2
        ss2.Row = 3
        ss2.Text = "合计"
        ss2.Col = SpreadHeader + 1
        ss2.Text = "合计"
        Call ss2.AddCellSpan(SpreadHeader + 1, 3, 2, 1)
    
    End With
        
    With ss3
        
        .MaxRows = 9
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        ss3.Col = SpreadHeader + 2
        ss3.Row = SpreadHeader
        ss3.Text = ""
        ss3.Col = SpreadHeader + 1
        ss3.Text = ""
        Call ss3.AddCellSpan(SpreadHeader + 1, SpreadHeader, 2, 1)

        
        ss3.Col = SpreadHeader + 1
        ss3.Row = 1
        ss3.Text = "转炉"
        ss3.Row = 2
        ss3.Text = "转炉"
        Call ss3.AddCellSpan(SpreadHeader + 1, 1, 1, 2)

        ss3.Col = SpreadHeader + 2
        
        ss3.Row = 1
        ss3.Text = "1#"
        
        ss3.Row = 2
        ss3.Text = "2#"
        
        ss3.Col = SpreadHeader + 2
        ss3.Row = 3
        ss3.Text = "合计"
        ss3.Col = SpreadHeader + 1
        ss3.Text = "合计"
        Call ss3.AddCellSpan(SpreadHeader + 1, 3, 2, 1)
    
        ss3.Col = SpreadHeader + 1
        ss3.Row = 4
        ss3.Text = "转炉"
        ss3.Row = 5
        ss3.Text = "转炉"
        Call ss3.AddCellSpan(SpreadHeader + 1, 4, 1, 2)

        ss3.Col = SpreadHeader + 2
        
        ss3.Row = 4
        ss3.Text = "1#"
        
        ss3.Row = 5
        ss3.Text = "2#"
        
        ss3.Col = SpreadHeader + 2
        ss3.Row = 6
        ss3.Text = "合计"
        ss3.Col = SpreadHeader + 1
        ss3.Text = "合计"
        Call ss3.AddCellSpan(SpreadHeader + 1, 6, 2, 1)
        
        ss3.Col = SpreadHeader + 2
        ss3.Row = 7
        ss3.Text = "VD"
        ss3.Col = SpreadHeader + 1
        ss3.Text = "VD"
        Call ss3.AddCellSpan(SpreadHeader + 1, 7, 2, 1)
    
        ss3.Col = SpreadHeader + 2
        ss3.Row = 8
        ss3.Text = "RH"
        ss3.Col = SpreadHeader + 1
        ss3.Text = "RH"
        Call ss3.AddCellSpan(SpreadHeader + 1, 8, 2, 1)
    
        ss3.Col = SpreadHeader + 2
        ss3.Row = 9
        ss3.Text = "合计"
        ss3.Col = SpreadHeader + 1
        ss3.Text = "合计"
        Call ss3.AddCellSpan(SpreadHeader + 1, 9, 2, 1)
    
    End With
    

End Sub

Public Function Sp_Process1(Conn As ADODB.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim dTempInt As Double
    Dim sMesg As String
    Dim sTemp As String
    Dim sPara As String
    Dim iTEXT As String
    Dim PZ As String
    Dim sEdate1 As String
    
    
        sEdate1 = Mid(dtp_YEAR_MONTH.Text, 1, 4) + Mid(dtp_YEAR_MONTH.Text, 6, 2)

    

    Dim adoCmd As ADODB.Command

    Sp_Process1 = True

    With ss1

        If .MaxRows < 1 Then
            Sp_Process1 = False
            Exit Function
        End If

        Screen.MousePointer = vbHourglass

        .ReDraw = False

        'Db Connection Check
        If Conn Is Nothing Then
            If GF_DbConnect = False Then Sp_Process1 = False: Exit Function
        End If

        'Ado Setting
        Conn.CursorLocation = adUseServer
        Set adoCmd = New ADODB.Command

        Set adoCmd.ActiveConnection = Conn
        adoCmd.CommandType = adCmdStoredProc
        
        adoCmd.CommandText = Sc.Item("P-M")

        Conn.BeginTrans

        'Ceate Parameter (Input) iType + iColumn
        For iCount = 0 To 5
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount

        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)

        For iRow = 6 To 7
'            For iCol = 1 To 9

            .Row = iRow
            .Col = SpreadHeader
             iTEXT = .Text

             If iTEXT = "Update" Then

                .Col = SpreadHeader + 2
                If .Text = "普板" Then
                   PZ = "PB"
                Else
                   PZ = "XP"
                End If
                
                For iCol = 1 To 9
                    
                    .Col = iCol
                    If Trim(.Text) = "" Then              '重量
                        adoCmd.Parameters(4).Value = 0
                        
                    Else
                        dTempInt = .Value
                        adoCmd.Parameters(4).Value = dTempInt
                    End If
                    
'                    If adoCmd.Parameters(4).Value <> 0 Then
                        If .Col = 1 Then
                           adoCmd.Parameters(5).Value = "ZL"
                        ElseIf .Col = 2 Then
                           adoCmd.Parameters(5).Value = "CZ"
                        ElseIf .Col = 3 Then
                           adoCmd.Parameters(5).Value = "CK"
                        ElseIf .Col = 4 Then
                           adoCmd.Parameters(5).Value = "GX"
                        ElseIf .Col = 5 Then
                           adoCmd.Parameters(5).Value = "GH"
                        ElseIf .Col = 6 Then
                           adoCmd.Parameters(5).Value = "GJ"
                        ElseIf .Col = 7 Then
                           adoCmd.Parameters(5).Value = "GN"
                        ElseIf .Col = 8 Then
                           adoCmd.Parameters(5).Value = "GB"
                        ElseIf .Col = 9 Then
                           adoCmd.Parameters(5).Value = "GZ"
                        End If

                        adoCmd.Parameters(0).Value = "1"                         'iTable
                        adoCmd.Parameters(1).Value = sEdate1
                        adoCmd.Parameters(2).Value = "PP"
                        adoCmd.Parameters(3).Value = PZ     '品种
                        
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
                            Sp_Process1 = False
                            Exit Function
                       End If
                       
'                    End If
  
                Next iCol
                
             End If
             
         Next iRow

         Conn.CommitTrans

        .ReDraw = True

        Screen.MousePointer = vbDefault

        Exit Function

    End With

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans

    Sp_Process1 = False

    Err.Raise Err.Number, Err.Description

End Function
Public Function Sp_Process2(Conn As ADODB.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim dTempInt As Double
    Dim sMesg As String
    Dim sTemp As String
    Dim sPara As String
    Dim iTEXT As String
    Dim PZ As String
    Dim sEdate1 As String
    
    
        sEdate1 = Mid(dtp_YEAR_MONTH.Text, 1, 4) + Mid(dtp_YEAR_MONTH.Text, 6, 2)

    

    Dim adoCmd As ADODB.Command

    Sp_Process2 = True

    With ss1

        If .MaxRows < 1 Then
            Sp_Process2 = False
            Exit Function
        End If

        Screen.MousePointer = vbHourglass

        .ReDraw = False

        'Db Connection Check
        If Conn Is Nothing Then
            If GF_DbConnect = False Then Sp_Process2 = False: Exit Function
        End If

        'Ado Setting
        Conn.CursorLocation = adUseServer
        Set adoCmd = New ADODB.Command

        Set adoCmd.ActiveConnection = Conn
        adoCmd.CommandType = adCmdStoredProc
        
        adoCmd.CommandText = Sc.Item("P-M")

        Conn.BeginTrans

        'Ceate Parameter (Input) iType + iColumn
        For iCount = 0 To 5
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount

        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)

        For iRow = 3 To 4
'            For iCol = 1 To 9

            .Row = iRow
            .Col = SpreadHeader
             iTEXT = .Text

             If iTEXT = "Update" Then

                .Col = SpreadHeader + 2
                If .Text = "普板" Then
                   PZ = "PB"
                Else
                   PZ = "XP"
                End If
                
                For iCol = 1 To 9
                    
                    .Col = iCol
                    If Trim(.Text) = "" Or Val(Trim(.Text)) = 0 Then               '重量
'                    If .Value = 0 Then
                        adoCmd.Parameters(4).Value = 0
                        
                    Else
                        dTempInt = .Value
                        adoCmd.Parameters(4).Value = dTempInt
                    End If
                    
'                    If adoCmd.Parameters(4).Value <> 0 Then
                    If .Col = 1 Then
                           adoCmd.Parameters(5).Value = "ZL"
                        ElseIf .Col = 2 Then
                           adoCmd.Parameters(5).Value = "CZ"
                        ElseIf .Col = 3 Then
                           adoCmd.Parameters(5).Value = "CK"
                        ElseIf .Col = 4 Then
                           adoCmd.Parameters(5).Value = "GX"
                        ElseIf .Col = 5 Then
                           adoCmd.Parameters(5).Value = "GH"
                        ElseIf .Col = 6 Then
                           adoCmd.Parameters(5).Value = "GJ"
                        ElseIf .Col = 7 Then
                           adoCmd.Parameters(5).Value = "GN"
                        ElseIf .Col = 8 Then
                           adoCmd.Parameters(5).Value = "GB"
                        ElseIf .Col = 9 Then
                           adoCmd.Parameters(5).Value = "GZ"
                    End If

                        adoCmd.Parameters(0).Value = "1"                         'iTable
                        adoCmd.Parameters(1).Value = sEdate1
                        adoCmd.Parameters(2).Value = "HC"
                        adoCmd.Parameters(3).Value = PZ     '品种
                        
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
                            Sp_Process2 = False
                            Exit Function
                       End If
                       
'                    End If
  
                Next iCol
                
             End If
             
         Next iRow

         Conn.CommitTrans

        .ReDraw = True

        Screen.MousePointer = vbDefault

        Exit Function

    End With

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans

    Sp_Process2 = False

    Err.Raise Err.Number, Err.Description

End Function


