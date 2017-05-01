VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0090C 
   Caption         =   "客户特殊要求成分输入 - AQA0090C"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15240
   Begin VB.TextBox txt_CUST_STD_NAME 
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
      Left            =   4500
      TabIndex        =   3
      Top             =   120
      Width           =   3165
   End
   Begin VB.TextBox txt_INS_EMP 
      Height          =   315
      Left            =   7740
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   825
   End
   Begin InDate.ULabel uLab_CUST_DETAIL 
      Height          =   315
      Left            =   1950
      Top             =   525
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   556
      Caption         =   ""
      Alignment       =   0
      BackColor       =   -2147483639
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txt_CUST_SPEC_NO 
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
      Left            =   1950
      MaxLength       =   18
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   150
      Top             =   120
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      Caption         =   "客户特殊要求编号"
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
      Left            =   150
      Top             =   525
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      Caption         =   "说明"
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
      Height          =   8025
      Left            =   120
      TabIndex        =   1
      Top             =   1050
      Width           =   15090
      _Version        =   393216
      _ExtentX        =   26617
      _ExtentY        =   14155
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
      MaxCols         =   19
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0090C.frx":0000
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   135
      X2              =   15165
      Y1              =   930
      Y2              =   945
   End
End
Attribute VB_Name = "AQA0090C"
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
'-- Program Name      客户特殊要求成分输入
'-- Program ID        AQA0020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       客户特殊要求成分输入
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

Dim ArrayRecords As Variant

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Hsheet"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_CUST_SPEC_NO, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    
    'MASTER Collection
    Mc1.Add Item:="AQA0090C.P_DELETE_ALL", Key:="P-M"
'     Mc1.Add Item:="PKG_MASTER.P_REFER", Key:="P-R"
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
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0090C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQA0090C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQA0090C.P_ONEROW", Key:="P-O"
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


     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
 
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        
        Case "txt_CUST_SPEC_NO"          '客户特殊要求编号
            sCode = "CUST_SPEC_NO"
            Set oCodeName = txt_CUST_STD_NAME
                        
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call subMenuHide

    
    If Len(Trim(txt_CUST_SPEC_NO.Text)) <> 0 Then
        Call Form_Ref
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

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
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 2)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 7)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 12)
    
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
    
    Call GS_SetChemicalSpreadLineColor(ss1, "0611")
      
End Sub


Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub


Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call subMenuHide
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        txt_CUST_STD_NAME.Text = ""
        uLab_CUST_DETAIL.Caption = ""
        pControl(1).SetFocus
    End If
    
   
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

'    sMesg = Gf_Ms_NeceCheck(pControl)
'    If sMesg = "OK" Then
'
'        sMesg = Gf_Ms_NeceCheck2(mControl)
'        If sMesg = "OK" Then

            'If Gf_Sp_Display(M_CN1, Proc_Sc("Sc").Item("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), Proc_Sc("Sc").Item("pColumn")) Then
           If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                Call Gp_Ms_ControlLock(Mc1("pControl"), True)
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call subMenuHide
                Call GS_SetChemicalLength(ss1, ArrayRecords, "020712", "2")
'                Call GP_ChemCode_RowHeader_Clear("AQA0090C", ss1, 1, 1)
                ss1.SetFocus
            End If

'        Else
'            sMesg = sMesg + " Must input according to length of item"
'            Call Gp_MsgBoxDisplay(sMesg)
'        End If
'
'    Else
'        sMesg = sMesg + " Must input necessarily"
'        Call Gp_MsgBoxDisplay(sMesg)
'
'    End If
    
    Dim i As Integer
    
    Call GS_SetChemicalSpreadLineColor(ss1, "0611")
    
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
            
    If Gf_Mc_Authority(sAuthority, Mc1, Proc_Sc("Sc")) Then
         txt_ins_emp.Text = sUserID
         If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
             Call subMenuHide
             Call GS_SetChemicalLength(ss1, ArrayRecords, "020712", "2")
         End If
     End If
           
         
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub form_Cpy()
  Call Gf_Ms_Copy(Mc1)
End Sub

Public Sub form_Pst()

    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call subMenuHide
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 16, "P")
    End If
    
End Sub
Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub
Public Sub Master_Pst()

    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 16, "P")
    End If
   
End Sub

Public Sub Spread_Pst()
   txt_CUST_SPEC_NO.Enabled = True
End Sub
Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 16)

End Sub

Public Sub Form_Del()

    If Gf_Ms_Del(M_CN1, Mc1) Then
        Call Gf_Sp_Cls(Proc_Sc("Sc"))
        txt_CUST_SPEC_NO.Text = ""
        txt_CUST_STD_NAME.Text = ""
        uLab_CUST_DETAIL.Caption = ""
        Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
        Call subMenuHide
    End If
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
'            .Col = Col
'            .Row = Row
'            Call GF_GetCeqValue(ss1, Me.Name, "2")
        End If
    End With
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 16)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 16)
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

        Case 2, 7, 12

            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1

                DD.sWitch = "SP"
              '  DD.sKey = "Q0001"
                DD.rControl.Add Item:=ss1.ActiveCol
                DD.nameType = "2"

                Call GF_CHEM_SEQ(M_CN1, KeyCode)
                
                Call GS_SetChemicalLength(ss1, ArrayRecords, "020712", "2")
        
            End If

    End Select
    
End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Call GP_ChemCode_RowHeader_Clear("AQA0090C", ss1, NewRow, NewCol)
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
    'TipText = ss1.CellTag
    'ShowTip = True
    Dim sTip As String
    ShowTip = True
    'With ss1
        
    sTip = GF_GetCellMaxLength(ss1, ss1.ActiveRow, ss1.ActiveCol)
    TipText = sTip
    'End With
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

