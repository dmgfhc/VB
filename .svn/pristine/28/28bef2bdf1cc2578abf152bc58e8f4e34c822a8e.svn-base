VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGH2020C 
   Caption         =   "轧钢生产线停机实绩查询及修改界面_CGH2020C"
   ClientHeight    =   7605
   ClientLeft      =   570
   ClientTop       =   1890
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_GROUP 
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
      ItemData        =   "CGH2020C.frx":0000
      Left            =   14475
      List            =   "CGH2020C.frx":0010
      TabIndex        =   6
      Top             =   90
      Width           =   735
   End
   Begin VB.ComboBox CBO_SHIFT 
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
      ItemData        =   "CGH2020C.frx":0020
      Left            =   12660
      List            =   "CGH2020C.frx":002D
      TabIndex        =   5
      Top             =   90
      Width           =   735
   End
   Begin VB.ComboBox CBO_PLT 
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
      ItemData        =   "CGH2020C.frx":003A
      Left            =   1200
      List            =   "CGH2020C.frx":003C
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   735
   End
   Begin VB.ComboBox CBO_PRC 
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
      ItemData        =   "CGH2020C.frx":003E
      Left            =   3540
      List            =   "CGH2020C.frx":0063
      TabIndex        =   1
      Top             =   90
      Width           =   735
   End
   Begin VB.TextBox TXT_DEL_RES_CD 
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
      Left            =   10745
      TabIndex        =   3
      Top             =   90
      Width           =   825
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8655
      Left            =   90
      TabIndex        =   4
      Top             =   510
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   15266
      _StockProps     =   64
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
      MaxCols         =   15
      MaxRows         =   10
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGH2020C.frx":0093
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   90
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "工厂代码"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   2430
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "工序代码"
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
   Begin CSTextLibCtl.sitxEdit TXT_OCCR_TIME 
      Height          =   315
      Left            =   5910
      TabIndex        =   2
      Top             =   90
      Width           =   1245
      _Version        =   262145
      _ExtentX        =   2196
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
      Text            =   "____-__-__"
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
      ValidateMask    =   0   'False
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   4770
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "发生时间"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   9630
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "停机代码"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   11790
      Top             =   90
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "班次"
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
      Left            =   13605
      Top             =   90
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      Caption         =   "班别"
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
   Begin InDate.UDate SDT_PROD_DATE_FROM 
      Height          =   315
      Left            =   15750
      TabIndex        =   7
      Tag             =   "起始日期"
      Top             =   4170
      Visible         =   0   'False
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
      BackColor       =   16777215
   End
   Begin InDate.UDate SDT_PROD_DATE_TO 
      Height          =   315
      Left            =   17445
      TabIndex        =   8
      Tag             =   "起始日期"
      Top             =   4170
      Visible         =   0   'False
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
      BackColor       =   16777215
   End
   Begin CSTextLibCtl.sitxEdit TXT_OCCR_TIME_TO 
      Height          =   330
      Left            =   7365
      TabIndex        =   10
      Top             =   90
      Width           =   1245
      _Version        =   262145
      _ExtentX        =   2196
      _ExtentY        =   582
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
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
      BorderEffect    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "____-__-__"
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
      ValidateMask    =   0   'False
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   7185
      TabIndex        =   11
      Top             =   210
      Width           =   210
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   17265
      TabIndex        =   9
      Top             =   4290
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "CGH2020C"
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
'-- Program Name      轧钢生产线停机实绩查询及修改界面
'-- Program ID        CGH2020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2007.10.10
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
'Public sDateTime As String          'Active Form Time Setting

Dim sDateTime_str As String         'Active Form Time Setting
Dim sDateTime_end As String         'Active Form Time Setting
Dim sDateTime_cnt As Double         'Active Form Time Setting

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
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(CBO_PRC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_OCCR_TIME, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(TXT_OCCR_TIME_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'  Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_DEL_RES_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGH2020C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="CGH2020C.P_REFER", Key:="P-R"
    sc1.Add Item:="CGH2020C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 1, True)
    'Call Gp_Sp_ColHidden(ss1, 4, True)
    Call Gp_Sp_ColHidden(ss1, 13, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

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

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 2)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 5)
    
    CBO_PLT.Text = "C3"

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        pControl(1).SetFocus
        CBO_PLT.Text = "C3"
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    Dim iDateCheck As String
    Dim iCount As Integer

'    For iCount = 1 To ss1.MaxRows
'
'        Select Case Trim(Gf_Sp_RcvData(ss1, 0, iCount))
'
'            Case "Update", "Input"
'
'                  With ss1
''                      .Col = 3
''                      If Not Gp_DateCheck(.Text, "S") Then
''                         Call Gp_MsgBoxDisplay("请正确输入发生时间")
''                         Exit Sub
''                      End If
'''                      .Col = 7
'''                      If Not Gp_DateCheck(.Text, "S") Then
'''                         Call Gp_MsgBoxDisplay("请正确输入停机开始时间")
'''                         Exit Sub
'''                      End If
''                      sDateTime_str = Mid(.Text, 1, 4) & Mid(.Text, 6, 2) & Mid(.Text, 9, 2) & Mid(.Text, 12, 2) & Mid(.Text, 15, 2) & Mid(.Text, 18, 2)
'''                      .Col = 8
'''                      iDateCheck = Replace(Trim(.Text), "-", "")
'''                      iDateCheck = Replace(iDateCheck, ":", "")
'''                      iDateCheck = Trim(iDateCheck)
'''                      If iDateCheck <> "" Then
'''                            If Not Gp_DateCheck(.Text, "S") Then
'''                               Call Gp_MsgBoxDisplay("请正确输入停机结束时间")
'''                               Exit Sub
'''                            End If
''                            sDateTime_end = Mid(.Text, 1, 4) & Mid(.Text, 6, 2) & Mid(.Text, 9, 2) & Mid(.Text, 12, 2) & Mid(.Text, 15, 2) & Mid(.Text, 18, 2)
''                            If Val(sDateTime_end) - Val(sDateTime_str) < 0 Then
''                               Call Gp_MsgBoxDisplay("停机结束时间应大于停机开始时间")
''                               Exit Sub
''                            End If
'                      End If
'                  End With
'
'        End Select
'
'    Next iCount

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

End Sub

Public Sub Form_Ins()

    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    ss1.ROW = ss1.ActiveRow
    ss1.Col = 1
    ss1.Text = IIf(CBO_PLT.Text = "", "C3", CBO_PLT.Text)
    ss1.Col = 2
    If CBO_PRC.Text = "" Then
       ss1.Text = "CB"
    Else
       ss1.Text = CBO_PRC.Text
    End If
    ss1.Col = 3
    ss1.Text = "1"
    ss1.Col = 10
    ss1.Text = sShiftSet
    ss1.Col = 12
    ss1.Text = sUserID
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

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

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()

    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()
     If SDT_PROD_DATE_FROM.RawData = "" Then
        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_DATE_TO_GotFocus()
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

' Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
'    ss1.Row = Row
'    ss1.Col = Col
'
'    If ss1.Col = 7 Then
'           ss1.Col = 5
'           sDateTime_str = Mid(ss1.Text, 6, 2) & "/" & Mid(ss1.Text, 9, 2) & "/" & Mid(ss1.Text, 1, 4) & " " & Mid(ss1.Text, 12, 2) & ":" & Mid(ss1.Text, 15, 2) & ":" & Mid(ss1.Text, 18, 2)
'          'sDateTime_str = "#" & Mid(ss1.Text, 6, 2) & "/" & Mid(ss1.Text, 9, 2) & "/" & Mid(ss1.Text, 1, 4) & " " & Mid(ss1.Text, 12, 2) & ":" & Mid(ss1.Text, 15, 2) & ":" & Mid(ss1.Text, 18, 2) & "#"
'           ss1.Col = 6
'           sDateTime_end = Mid(ss1.Text, 6, 2) & "/" & Mid(ss1.Text, 9, 2) & "/" & Mid(ss1.Text, 1, 4) & " " & Mid(ss1.Text, 12, 2) & ":" & Mid(ss1.Text, 15, 2) & ":" & Mid(ss1.Text, 18, 2)
'           sDateTime_cnt = Round(Mid((CDate(sDateTime_end) - CDate(sDateTime_str)) * 1440, 1, 4), 0)
'               'sDateTime_cnt = DateDiff("n", (CDate(sDateTime_end) - CDate(sDateTime_str)))
'               'sDateTime_cnt = Round(Mid((CDate(Mid(sDateTime_end, 2, 19)) - CDate(Mid(sDateTime_str, 2, 19))) * 1440, 1, 4), 0)
'               ss1.Col = 7
'                   ss1.Text = sDateTime_cnt
'                   ss1.Col = 0
'                   Select Case Trim(ss1.Text)
'                          Case "Input", "Update", "Delete"
'                          Case Else
'                          ss1.Text = "Update"
'                   End Select
'    End If

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)

    If ROW = 0 Then Exit Sub
    ss1.ROW = ROW
    ss1.Col = Col
     
    If ss1.Lock = False Then
        If ss1.Col = 4 Then

'        ss1.Text = Mid(Gf_DTSet(M_CN1), 1, 4) + " " + Mid(Gf_DTSet(M_CN1), 5, 2) + " " + Mid(Gf_DTSet(M_CN1), 7, 2) + " " + _
'                   Mid(Gf_DTSet(M_CN1), 9, 2) + " " + Mid(Gf_DTSet(M_CN1), 11, 2) + " " + Mid(Gf_DTSet(M_CN1), 13, 2)

         ss1.Text = Format(Now, "YYYY-MM-DD")

         ss1.Col = 0
         Select Case Trim(ss1.Text)
                Case "Input", "Update", "Delete"
                Case Else
                ss1.Text = "Update"
         End Select
       End If
    End If

    If ss1.Col = 7 Then
         
'        ss1.Text = Mid(Gf_DTSet(M_CN1, "I"), 1, 4) + " " + Mid(Gf_DTSet(M_CN1, "I"), 5, 2) + " " + Mid(Gf_DTSet(M_CN1, "I"), 7, 2) + " " + _
'                   Mid(Gf_DTSet(M_CN1, "I"), 9, 2) + " " + Mid(Gf_DTSet(M_CN1, "I"), 11, 2)
        If ss1.Lock = False Then
            ss1.Text = Format(Now, "HH:MM")
                    
            ss1.Col = 0
            Select Case Trim(ss1.Text)
                   Case "Input", "Update", "Delete"
                   Case Else
                        ss1.Text = "Update"
            End Select
        End If
    End If
    
    If ss1.Col = 8 Then

'        ss1.Text = Mid(Gf_DTSet(M_CN1, "I"), 1, 4) + " " + Mid(Gf_DTSet(M_CN1, "I"), 5, 2) + " " + Mid(Gf_DTSet(M_CN1, "I"), 7, 2) + " " + _
'                   Mid(Gf_DTSet(M_CN1, "I"), 9, 2) + " " + Mid(Gf_DTSet(M_CN1, "I"), 11, 2)
                   
        ss1.Text = Format(Now, "HH:MM")

        ss1.Col = 0
        Select Case Trim(ss1.Text)
               Case "Input", "Update", "Delete"
               Case Else
                    ss1.Text = "Update"
        End Select
    End If
'    ss1.Col = 1
'    ss1.Text = CBO_PLT.Text
    
End Sub

Private Sub ss1_EditChange(ByVal Col As Long, ByVal ROW As Long)
Dim sTemp_Mana_Code As String
Dim sTemp_Code As String
    ss1.ROW = ss1.ActiveRow
    ss1.Col = Col
    If ss1.Col = 5 Then
       If ss1.Text = "" Then
          ss1.Col = 6
          ss1.Text = ""
       ElseIf Len(Trim(ss1.Text)) = 5 Then
              sTemp_Mana_Code = "G0013"
              sTemp_Code = ss1.Text
              ss1.Col = 6
              ss1.Text = Gf_ComnNameFind(M_CN1, sTemp_Mana_Code, Trim(sTemp_Code), 1)
       End If
    End If
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If

End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub

    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub

    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
      '  Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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

Private Sub TXT_DEL_RES_CD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0013"
        DD.rControl.Add Item:=TXT_DEL_RES_CD
       
        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub

Private Sub TXT_OCCR_TIME_DblClick()

    TXT_OCCR_TIME.RawData = Gf_DTSet(M_CN1, "D")

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If ss1.Col = 5 Then
          
         If KeyCode = vbKeyF4 Then
              
            Set DD.sPname = Me.ss1
                  
            DD.sWitch = "SP"
            DD.sKey = "G0013"
            DD.rControl.Add Item:=5
            DD.rControl.Add Item:=6
                  
            DD.nameType = "1"
                  
            Call Gf_Common_DD(M_CN1, KeyCode)
              
        End If
        
    End If

    If ss1.Col = 2 Then
          
         If KeyCode = vbKeyF4 Then
              
            Set DD.sPname = Me.ss1
                  
            DD.sWitch = "SP"
            DD.sKey = "C0002"
            DD.rControl.Add Item:=2
            
                  
            DD.nameType = "1"
                  
            Call Gf_Common_DD(M_CN1, KeyCode)
              
        End If
        
    End If

End Sub


Private Sub TXT_OCCR_TIME_TO_DblClick()

     TXT_OCCR_TIME_TO.RawData = Gf_DTSet(M_CN1, "D")

End Sub
