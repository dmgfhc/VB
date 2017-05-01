VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC3020C 
   Caption         =   "产品退判查询_AGC3020C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_KND 
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
      Height          =   315
      ItemData        =   "AGC3020C.frx":0000
      Left            =   12315
      List            =   "AGC3020C.frx":0010
      TabIndex        =   7
      Top             =   90
      Width           =   1680
   End
   Begin VB.ComboBox CBO_PLT 
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
      Height          =   315
      ItemData        =   "AGC3020C.frx":0042
      Left            =   1455
      List            =   "AGC3020C.frx":0055
      TabIndex        =   6
      Top             =   90
      Width           =   750
   End
   Begin VB.TextBox TXT_PLATE_NO 
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
      Left            =   9090
      TabIndex        =   1
      Top             =   90
      Width           =   1695
   End
   Begin VB.TextBox TXT_KND 
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
      Left            =   14400
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   435
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8730
      Left            =   90
      TabIndex        =   2
      Top             =   510
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   15399
      _StockProps     =   64
      ColsFrozen      =   4
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
      MaxRows         =   10
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AGC3020C.frx":006D
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   2685
      Top             =   90
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Caption         =   "退判时间"
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
      Left            =   7935
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "查询号"
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
   Begin InDate.UDate SDT_PROD_DATE 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Tag             =   "起始日期"
      Top             =   90
      Width           =   1485
      _ExtentX        =   2619
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
   Begin InDate.UDate SDT_PROD_TO_DATE 
      Height          =   315
      Left            =   6120
      TabIndex        =   5
      Tag             =   "起始日期"
      Top             =   90
      Width           =   1485
      _ExtentX        =   2619
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   300
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "生产厂"
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
      Left            =   11160
      Top             =   90
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      Caption         =   "分类"
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
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   5910
      TabIndex        =   3
      Top             =   210
      Width           =   195
   End
End
Attribute VB_Name = "AGC3020C"
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
'-- Program Name      钢板实绩查询界面
'-- Program ID        AGC2200C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2003.7.23
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
Public sDateTime As String          'Active Form Time Setting
Public sQuery_load As String        'Active Form sQuery Setting

Dim pControl As New Collection      'Master Primary Key Collection
Dim nControl As New Collection      'Master Necessary Collection
Dim mControl As New Collection      'Master Maxlength check Collection
Dim iControl As New Collection      'Master Insert Collection
Dim rControl As New Collection      'Master Refer Collection
Dim cControl As New Collection      'Master Copy Collection
Dim aControl As New Collection      'Master -> Spread Collection
Dim lControl As New Collection      'Master Lock Collection

'Dim pControl1 As New Collection      'Master Primary Key Collection
'Dim nControl1 As New Collection      'Master Necessary Collection
'Dim mControl1 As New Collection      'Master Maxlength check Collection
'Dim iControl1 As New Collection      'Master Insert Collection
'Dim rControl1 As New Collection      'Master Refer Collection
'Dim cControl1 As New Collection      'Master Copy Collection
'Dim aControl1 As New Collection      'Master -> Spread Collection
'Dim lControl1 As New Collection      'Master Lock Collection

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
'Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Const SPD_1 = 1
Const SPD_PLATE_NO = 2
Const SPD_SQ_REASON = 4
Const SPD_SQ_DATE = 5
Const SPD_CL_DATE = 8
Const SPD_SQ_ID = 20
Const SPD_USER_ID = 21

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
          Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_PLATE_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDT_PROD_DATE, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(SDT_PROD_TO_DATE, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(TXT_KND, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 
        Mc1.Add Item:=pControl, Key:="pControl"
        Mc1.Add Item:=nControl, Key:="nControl"
        Mc1.Add Item:=mControl, Key:="mControl"
        Mc1.Add Item:=iControl, Key:="iControl"
        Mc1.Add Item:=rControl, Key:="rControl"
        Mc1.Add Item:=cControl, Key:="cControl"
        Mc1.Add Item:=aControl, Key:="aControl"
        Mc1.Add Item:=lControl, Key:="lControl"
        
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC3020C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGC3020C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AGC3020C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
     
End Sub

Private Sub CBO_KND_Click()

     Call Form_Cls

     TXT_KND.Text = Mid(CBO_KND.Text, 1, 1)
     If TXT_KND.Text = "1" Then
        ULabel5.Caption = "退判时间"
        ss1.Row = 0: ss1.Col = SPD_SQ_REASON: ss1.Text = "退判原因"
        ss1.Row = 0: ss1.Col = SPD_SQ_DATE: ss1.Text = "退判时间"
        ss1.Row = 0: ss1.Col = SPD_CL_DATE: ss1.Text = "删除时间"
        ss1.Row = 0: ss1.Col = SPD_SQ_ID: ss1.Text = "退判人员"
        ss1.Row = 0: ss1.Col = SPD_USER_ID: ss1.Text = "删除人员"
        Call Gp_Sp_ColHidden(ss1, SPD_1, True)
        Call Gp_Sp_ColHidden(ss1, SPD_CL_DATE, True)
        Call Gp_Sp_ColHidden(ss1, SPD_USER_ID, True)
     ElseIf TXT_KND.Text = "4" Then
        ULabel5.Caption = "钢板删除时间"
        ss1.Row = 0: ss1.Col = SPD_SQ_REASON: ss1.Text = "申请原因"
        ss1.Row = 0: ss1.Col = SPD_SQ_DATE: ss1.Text = "申请时间"
        ss1.Row = 0: ss1.Col = SPD_CL_DATE: ss1.Text = "删除时间"
        ss1.Row = 0: ss1.Col = SPD_SQ_ID: ss1.Text = "申请人员"
        ss1.Row = 0: ss1.Col = SPD_USER_ID: ss1.Text = "删除人员"
        Call Gp_Sp_ColHidden(ss1, SPD_1, False)
        Call Gp_Sp_ColHidden(ss1, SPD_CL_DATE, False)
        Call Gp_Sp_ColHidden(ss1, SPD_USER_ID, False)
     ElseIf TXT_KND.Text = "6" Then
        ULabel5.Caption = "退入库时间"
        ss1.Row = 0: ss1.Col = SPD_SQ_REASON: ss1.Text = "退入库原因"
        ss1.Row = 0: ss1.Col = SPD_SQ_DATE: ss1.Text = "退入库时间"
        ss1.Row = 0: ss1.Col = SPD_CL_DATE: ss1.Text = "删除时间"
        ss1.Row = 0: ss1.Col = SPD_SQ_ID: ss1.Text = "退入库人员"
        ss1.Row = 0: ss1.Col = SPD_USER_ID: ss1.Text = "删除人员"
        Call Gp_Sp_ColHidden(ss1, SPD_1, True)
        Call Gp_Sp_ColHidden(ss1, SPD_CL_DATE, True)
        Call Gp_Sp_ColHidden(ss1, SPD_USER_ID, True)
     ElseIf TXT_KND.Text = "8" Then
        ULabel5.Caption = "强制排产时间"
        ss1.Row = 0: ss1.Col = SPD_SQ_REASON: ss1.Text = "强制排产原因"
        ss1.Row = 0: ss1.Col = SPD_SQ_DATE: ss1.Text = "强制排产时间"
        ss1.Row = 0: ss1.Col = SPD_CL_DATE: ss1.Text = "删除时间"
        ss1.Row = 0: ss1.Col = SPD_SQ_ID: ss1.Text = "强制排产人员"
        ss1.Row = 0: ss1.Col = SPD_USER_ID: ss1.Text = "删除人员"
        Call Gp_Sp_ColHidden(ss1, SPD_1, True)
        Call Gp_Sp_ColHidden(ss1, SPD_CL_DATE, True)
        Call Gp_Sp_ColHidden(ss1, SPD_USER_ID, True)
     End If
     
End Sub

Private Sub SDT_PROD_DATE_GotFocus()
     SDT_PROD_DATE.RawData = Gf_DTSet(M_CN1, "D")
     SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub SDT_PROD_TO_DATE_GotFocus()
     SDT_PROD_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub TXT_PLATE_NO_Change()
   Dim SMESG As String
      If Len(TXT_PLATE_NO.Text) > 14 Then
      SMESG = "板坯号长度不能超过10位，请确认板坯号 ！！！"
      Call Gp_MsgBoxDisplay(SMESG)
   End If
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        If Len(TXT_PLATE_NO.Text) >= 8 Then
           Call Form_Ref
        End If
'        KeyAscii = 0
'        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_ControlLock(Mc1("lControl"), True)

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    If App.Title = "BG" Then
        CBO_PLT.Text = "C1"
    ElseIf App.Title = "CG" Then
        CBO_PLT.Text = "C3"
    Else
        CBO_PLT.Text = "C1"
    End If
    CBO_KND.Text = "1 钢板退判"
    TXT_KND.Text = "1"
    
    Call Gp_Sp_ColHidden(ss1, SPD_1, True)
    Call Gp_Sp_ColHidden(ss1, SPD_CL_DATE, True)
    Call Gp_Sp_ColHidden(ss1, SPD_USER_ID, True)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call Gp_Ms_Cls(Mc1("rControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If

End Sub

Public Sub Master_Cpy()

'    Call Gf_Ms_Copy(Mc1)

End Sub

Public Sub Master_Pst()

'     If Gf_Ms_Paste(M_CN1, Mc1) Then
'        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
'     End If

End Sub

Public Sub Form_Ref()
    
    Dim SMESG As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Val(SDT_PROD_DATE.RawData) - Val(SDT_PROD_TO_DATE.RawData) > 0 Then
         SMESG = " 时间范围输入错误，请重新输入时间信息 ！！！"
         Call Gp_MsgBoxDisplay(SMESG)
         Exit Sub
    End If
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

End Sub
Public Sub Form_Pro()

    Dim iCount As Integer

    For iCount = 1 To ss1.MaxRows

        Select Case Trim(Gf_Sp_RcvData(ss1, 0, iCount))
        
            Case "Insert"

            ss1.Col = SPD_USER_ID
            ss1.Text = sUserID

            Case "Update"

            ss1.Col = SPD_USER_ID
            ss1.Text = sUserID
            
            Case "Delete"

            ss1.Col = SPD_USER_ID
            ss1.Text = sUserID

        End Select

    Next iCount
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
End Sub
Public Sub Form_Ins()

    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    ss1.Row = ss1.ActiveRow
    ss1.Col = SPD_USER_ID
    ss1.Text = sUserID

End Sub
Public Sub Spread_Del()

    Call Gp_Sp_Del(Proc_Sc("SC"))
    ss1.Row = ss1.ActiveRow
    ss1.Col = SPD_USER_ID   '12 -> SPD_USER_ID
    ss1.Text = sUserID

End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub
Public Sub Spread_Cpy()
'
'    Call Gp_Sp_Copy(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Pst()

'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If TXT_KND.Text = "4" Then
        If Gf_Sc_Authority(sAuthority, "U") Then
            Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        End If
    End If

End Sub
Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    If ss1.MaxRows < 1 Then Exit Sub
    
    If Row = 0 Then
    
        Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0

    End If
    
End Sub
'Private Sub txt_stdspec_chg_DblClick()
'    Call txt_stdspec_chg_KeyUp(vbKeyF4, 0)
'End Sub
'
'Private Sub txt_stdspec_chg_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.rControl.Add Item:=txt_stdspec_chg
'
'        Call Gf_StdSPEC_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'End Sub



