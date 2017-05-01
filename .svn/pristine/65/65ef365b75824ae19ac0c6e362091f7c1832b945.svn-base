VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0020C 
   Caption         =   "标准成分信息查询 - AQA0020C"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   11400
   Begin VB.TextBox txt_THK_MAX 
      Height          =   270
      Left            =   7350
      TabIndex        =   6
      Tag             =   "厚度组-最大"
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt_THK_MIN 
      Height          =   270
      Left            =   6540
      TabIndex        =   5
      Tag             =   "厚度组-最小"
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmd_ListView 
      Caption         =   "<"
      Height          =   270
      Left            =   3315
      TabIndex        =   4
      Top             =   540
      Width           =   435
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   270
      Left            =   1290
      TabIndex        =   3
      Top             =   540
      Width           =   1965
      _Version        =   393216
      _ExtentX        =   3466
      _ExtentY        =   476
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   2
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "AQA0020C.frx":0000
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
      Left            =   1320
      MaxLength       =   18
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txt_STDSPEC_YY 
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
      Left            =   5010
      MaxLength       =   4
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "标准号"
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
      Left            =   3840
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "发布年度"
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
      Left            =   120
      Top             =   510
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "厚度组"
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
      Height          =   8160
      Left            =   120
      TabIndex        =   2
      Top             =   990
      Width           =   15105
      _Version        =   393216
      _ExtentX        =   26644
      _ExtentY        =   14393
      _StockProps     =   64
      EditEnterAction =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   22
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0020C.frx":0378
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   90
      X2              =   15210
      Y1              =   900
      Y2              =   900
   End
End
Attribute VB_Name = "AQA0020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量标准管理
'-- Program Name      标准成分信息输入
'-- Program ID        AQA0020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       标准成分信息输入
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long

Dim bChk As Boolean
Dim btChk As Boolean

Dim ArrayRecords As Variant

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Hsheet"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_STDSPEC, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STDSPEC_YY, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_THK_MIN, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_THK_MAX, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        
    'MASTER Collection
     Mc1.Add Item:="AQA0020C.P_DELETE_ALL", Key:="P-M"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0020C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQA0020C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQA0020C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=2, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
 
End Sub


Private Sub cmd_ListView_Click()
Dim sQuery As String

    sQuery = "Select Distinct THK_MIN,THK_MAX From QP_STD_CHEM Where STDSPEC = "
    btChk = Not btChk

    If btChk = False Then
            
        With ss2
        
            .MaxRows = 1
            .Height = 255
        
            btChk = False
    
        End With

    Else
        
        If txt_STDSPEC.Text = "" Or Trim(txt_STDSPEC.Text) = "" Then
            Exit Sub
        ElseIf txt_STDSPEC_YY.Text = "" Or Trim(txt_STDSPEC_YY.Text) = "" Then
            Exit Sub
        End If
        sQuery = sQuery + " '" + txt_STDSPEC.Text + "' And"
        sQuery = sQuery + " STDSPEC_YY = '" + txt_STDSPEC_YY.Text + "'"
        
        Call GS_Combo_SS_ADD(sQuery, ss2)
        'Call GS_Combo_THK_MAX2(Me)
        
        Call subBackColor
    
    End If
    
    If Gf_GetCellNullCheck(ss2, 1, 1) <> "" And Gf_GetCellNullCheck(ss2, 1, 2) <> "" Then
            txt_THK_MIN.Text = Gf_GetCellNullCheck(ss2, 1, 1)
            txt_THK_MAX.Text = Gf_GetCellNullCheck(ss2, 1, 2)
    End If
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    Call subMenuHide

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        Case "txt_STDSPEC"
            sCode = "STDSPEC"
            Set oCodeName = txt_STDSPEC_YY
        Case Else
            Exit Sub
    End Select
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub


Private Sub Form_Load()

    Dim x As Boolean

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
        
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call subMenuHide
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
        
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 5)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 10)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 15)
    
    Call GP_ROW_BACKCOLOR(ss2)
    
    Call subBackColor
    
    Screen.MousePointer = vbDefault
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    ArrayRecords = GF_GetChemicalCode
    
    ss1.TextTip = SS_TEXTTIP_FLOATINGFOCUSONLY
    ss1.TextTipDelay = 250
    x = ss1.SetTextTipAppearance("宋体", "11", False, False, &HFFFF&, &H800000)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
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
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    Call subMenuHide
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
    
    Call GS_SetChemicalSpreadLineColor(ss1, "0914")
      
End Sub


Public Sub Spread_Del()
    
'    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub form_Cpy()
  Call Gf_Ms_Copy(Mc1)
End Sub

Public Sub form_Pst()

    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19, "P")
    End If
    
End Sub


Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call subMenuHide
        ss2.MaxRows = 1
        ss2.Height = 255
        btChk = False
        Call GP_SET_CELL_VALUE(ss2, 1, 1, "")
        Call GP_SET_CELL_VALUE(ss2, 1, 2, "")
        txt_THK_MIN.Text = ""
        txt_THK_MAX.Text = ""
    End If
        
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    Dim i As Integer

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call subMenuHide
        Call GS_SetChemicalLength(ss1, ArrayRecords, "051015", "1")
        ss2.MaxRows = 1
        ss2.Height = 255
        btChk = False
    End If
    
    Call GS_SetChemicalSpreadLineColor(ss1, "0914")
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    Dim sMin As String
    Dim sMax As String
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        Call Form_Ref
    End If
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)

End Sub

Public Sub Form_Del()

    If Gf_Ms_Del(M_CN1, Mc1) Then
        Call Gf_Sp_Cls(Proc_Sc("Sc"))
        Call GS_Combo_THK_MAX(Me)
        Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
        Call subMenuHide
    End If

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub
Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub
Public Sub Master_Pst()

    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call subMenuHide
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19, "P")
    End If
   
End Sub


Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
   
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub


Private Sub ss1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    Dim ChemCode As String
    
    With ss1
     .Col = Col - 3: ChemCode = .Text
        If ChemCode <> "Ceq" Then
            .Col = Col
            .Row = Row
            .Text = ""
'        Else
'            Call GF_GetCeqValue(ss1, Me.Name, "1")
        End If
    End With
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)
    End If
    
End Sub


Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sTemp_Code As String
    Dim iCol As Integer

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol

        Case 5, 10, 15

            If KeyCode = vbKeyF4 Then
                
                ss1.Row = ss1.ActiveRow
                
                Set DD.sPname = Me.ss1
        
                DD.sWitch = "SP"
                DD.rControl.Add Item:=ss1.ActiveCol
                DD.nameType = "2"
                
                Call GF_CHEM_SEQ(M_CN1, KeyCode)
                
                Call GS_SetChemicalLength(ss1, ArrayRecords, "051015", "1")
        
            End If

    End Select

End Sub


Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub


'----------------------------------------------------
'ss1_TextTipFetch
'----------------------------------------------------
Private Sub ss1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    ShowTip = True
    TipText = GF_GetCellMaxLength(ss1, ss1.ActiveRow, ss1.ActiveCol)
End Sub


Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
            
    With ss2
    
        If Gf_GetCellNullCheck(ss2, Row, 1) <> "" And Gf_GetCellNullCheck(ss2, Row, 2) <> "" Then
            Call GP_SET_CELL_VALUE(ss2, 1, 1, Gf_GetCellNullCheck(ss2, Row, 1))
            Call GP_SET_CELL_VALUE(ss2, 1, 2, Gf_GetCellNullCheck(ss2, Row, 2))
        End If
        
        .MaxRows = 1
        .Height = 255
        
        txt_THK_MIN.Text = Gf_GetCellNullCheck(ss2, 1, 1)
        txt_THK_MAX.Text = Gf_GetCellNullCheck(ss2, 1, 2)
        
        btChk = False
    
    End With
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
        txt_THK_MIN.Text = Gf_GetCellNullCheck(ss2, 1, 1)
        txt_THK_MAX.Text = Gf_GetCellNullCheck(ss2, 1, 2)
End Sub


Private Sub subMenuHide()
    
    With MDIMain.MenuTool

        '.Buttons(12).Enabled = True                    'Copy
        .Buttons(11).ButtonMenus(1).Enabled = True      'All Copy
        .Buttons(11).ButtonMenus(2).Enabled = False     'Master Copy
        .Buttons(11).ButtonMenus(3).Enabled = False     'Spread Copy
                    
        '.Buttons(12).Enabled = True                    'Paste
        .Buttons(12).ButtonMenus(1).Enabled = True      'All Paste
        .Buttons(12).ButtonMenus(2).Enabled = False     'Master Paste
        .Buttons(12).ButtonMenus(3).Enabled = False     'Spread Paste
        
    End With

End Sub

Private Sub subBackColor()

    Dim i As Integer
    
    For i = 1 To ss2.MaxRows

        Call Gp_Sp_RowColor(ss2, i, vbBlack, &HC0FFFF)
        
    Next i

End Sub


