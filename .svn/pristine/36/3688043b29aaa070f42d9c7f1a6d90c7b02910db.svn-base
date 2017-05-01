VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQA0133C 
   Caption         =   "企标成分信息输入"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_INS_EMP 
      Height          =   285
      Left            =   3780
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_STLGRD_GRP 
      Height          =   300
      Left            =   870
      TabIndex        =   3
      Top             =   510
      Width           =   2025
   End
   Begin VB.TextBox txt_STLGRD_DETAIL 
      Height          =   300
      Left            =   3750
      TabIndex        =   1
      Top             =   120
      Width           =   6465
   End
   Begin VB.TextBox txt_STLGRD 
      Height          =   300
      Left            =   870
      MaxLength       =   11
      TabIndex        =   0
      Top             =   120
      Width           =   2025
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Left            =   120
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      Caption         =   "钢种"
      Alignment       =   0
      BackColor       =   14804173
      BackgroundStyle =   1
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
   Begin InDate.ULabel ULabel2 
      Height          =   300
      Index           =   0
      Left            =   3000
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      Caption         =   "说明"
      Alignment       =   0
      BackColor       =   14804173
      BackgroundStyle =   1
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
   Begin FPSpread.vaSpread ss1 
      Height          =   7530
      Left            =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   990
      Width           =   15255
      _Version        =   393216
      _ExtentX        =   26908
      _ExtentY        =   13282
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      SpreadDesigner  =   "AQA0133C.frx":0000
   End
   Begin InDate.ULabel ULabel2 
      Height          =   300
      Index           =   1
      Left            =   120
      Top             =   510
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      Caption         =   "钢种组"
      Alignment       =   0
      BackColor       =   14804173
      BackgroundStyle =   1
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   15240
      Y1              =   930
      Y2              =   930
   End
End
Attribute VB_Name = "AQA0133C"
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
'-- Program Name      企标成分信息输入（钢种说明）
'-- Program ID        AQA0133C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
'-- Date              2003.5.19
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long


Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Hsheet"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(txt_STLGRD, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STLGRD_DETAIL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STLGRD_GRP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_INS_EMP, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
     Mc1.Add Item:="AQA0133C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQA0133C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

    Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0133C.P_SMODIFY", Key:="P-M"
    Sc1.Add Item:="AQA0133C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AQA0133C.P_SONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=2, Key:="First"
    Sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0


     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
 
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
     Call subButtonHide

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
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
     Call subButtonHide
     
      Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
   

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)
    
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
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
    
    Call subLineColorSet
      
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
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        pControl(1).SetFocus
    End If
    
     Call subButtonHide
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg  As String
    Dim sQuery As String
    

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    sMesg = Gf_Ms_NeceCheck(pControl)
    If sMesg = "OK" Then

        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then
            'Make Query
            sQuery = Gf_Ms_MakeQuery(Mc1.Item("P-R"), "R", Mc1.Item("pControl"))
    
            'Query Excete and Display
             If Gf_Ms_Display(M_CN1, sQuery, Mc1.Item("rControl"), Mc1.Item("lControl")) = "OK" Then
                Call Gf_Sp_Display(M_CN1, Proc_Sc("Sc").Item("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), Proc_Sc("Sc").Item("pColumn"))
                Call Gp_Ms_ControlLock(Mc1("pControl"), True)
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Else
                Call Gf_Sp_Display(M_CN1, Proc_Sc("Sc").Item("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), Proc_Sc("Sc").Item("pColumn"))
                Call Gp_Ms_ControlLock(Mc1("pControl"), False)
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            
            End If
        Else
            sMesg = sMesg + " Must input according to length of item"
            Call Gp_MsgBoxDisplay(sMesg)
        End If

    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)

    End If
    
    Dim i As Integer
    
    Call subLineColorSet
        
   Call subButtonHide
    
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
         
If Len(Trim(txt_STLGRD_DETAIL.Text)) = 0 Then
 Call Gp_MsgBoxDisplay("请输入钢种说明")
 Exit Sub
End If

If Len(Trim(txt_STLGRD_GRP.Text)) = 0 Then
 Call Gp_MsgBoxDisplay("请输入钢种说明")
 Exit Sub
End If

    If Gf_Mc_Authority(sAuthority, Mc1, Proc_Sc("Sc")) Then
        txt_INS_EMP.Text = sUserID
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Then
            Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1)
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        End If
    End If
         
         Call subButtonHide
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)

End Sub

Public Sub Form_Del()

    If Not Gf_Ms_AllDel(M_CN1, Proc_Sc("Sc"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub Form_Resize()
 ss1.Left = 0
 If Me.Width > 198 Then
 ss1.Width = Me.Width - 198
 End If
 If Me.Height > 1485 Then
 ss1.Height = Me.Height - 1485
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
     .Col = Col - 4: ChemCode = .Text
        If ChemCode <> "Ceq" Then
            .Col = Col
            .Row = Row
            .Text = ""
        Else
            .Col = Col
            .Row = Row
            Call subGetCeqValue(Col, Row, .Text)
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
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc")("Spread"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub subButtonHide()
    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
   ' MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
   
    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
    

End Sub

Private Sub subGetCeqValue(ByVal iCol As Long, ByVal iRow As Long, ByVal FOMULA_CD As String)
    
    Dim i As Integer
    
    Dim CHEM_C(3) As Double
    Dim CHEM_SI(3) As Double
    Dim CHEM_MN(3) As Double
    Dim CHEM_P(3) As Double
    Dim CHEM_S(3) As Double
    Dim CHEM_CR(3) As Double
    Dim CHEM_V(3) As Double
    Dim CHEM_MO(3) As Double
    Dim CHEM_CU(3) As Double
    Dim CHEM_NI(3) As Double
    Dim CHEM_B(3) As Double
    
    Dim Chem_Val(3) As Double 'Ceq Values

    
    For i = 1 To 3
    
        CHEM_C(i - 1) = subGetChemValue("C", i)
        CHEM_SI(i - 1) = subGetChemValue("Si", i)
        CHEM_MN(i - 1) = subGetChemValue("Mn", i)
        CHEM_P(i - 1) = subGetChemValue("P", i)
        CHEM_S(i - 1) = subGetChemValue("S", i)
        CHEM_CR(i - 1) = subGetChemValue("Cr", i)
        CHEM_V(i - 1) = subGetChemValue("V", i)
        CHEM_MO(i - 1) = subGetChemValue("Mo", i)
        CHEM_CU(i - 1) = subGetChemValue("Cu", i)
        CHEM_NI(i - 1) = subGetChemValue("Ni", i)
        CHEM_B(i - 1) = subGetChemValue("B", i)
        
    Next i
    
    For i = 0 To 2
        
        Chem_Val(i) = get_CHEM_CEQ(FOMULA_CD, CHEM_C(i), CHEM_SI(i), CHEM_MN(i), CHEM_P(i), CHEM_S(i), CHEM_CR(i), CHEM_V(i), _
                       CHEM_MO(i), CHEM_CU(i), CHEM_NI(i), CHEM_B(i))
        
    Next i
    
    ss1.Row = iRow
    ss1.Col = iCol - 3: ss1.Text = Chem_Val(0) 'Chem_min
    ss1.Col = iCol - 2: ss1.Text = Chem_Val(1) 'Chem_max
    ss1.Col = iCol - 1: ss1.Text = Chem_Val(2) 'Chem_tgt
    
End Sub


Private Function subGetChemValue(sChem As String, iNo As Integer) As Double
    
    Dim i As Integer
    Dim iCol_1 As Integer
    Dim iCol_2 As Integer
    Dim iCol_3 As Integer

    
    Select Case iNo
    
        Case 1              'MIN Value
            iCol_1 = 3      'First Chem
            iCol_2 = 9     'Second Chem
            iCol_3 = 15     'Third Chem
        Case 2              'MAX Value
            iCol_1 = 4      'First Chem
            iCol_2 = 10     'Second Chem
            iCol_3 = 16     'Third Chem
        Case 3              'TGT Value
            iCol_1 = 5
            iCol_2 = 11
            iCol_3 = 17
    End Select
    
    
    With ss1
    
        .Col = 2 'CHEM_CD ColAddress
    
        For i = 1 To .MaxRows
                    
            .Row = i
            
            If .Text = sChem Then
                .Col = iCol_1
                    If .Value = "" Then
                        subGetChemValue = 0
                    Else
                        subGetChemValue = .Value
                    End If
                Exit Function
            End If
        
        Next i
        
        .Col = 8 '12
        
        For i = 1 To .MaxRows
                    
            .Row = i
            
            If .Text = sChem Then
                .Col = iCol_2
                    If .Value = "" Then
                        subGetChemValue = 0
                    Else
                        subGetChemValue = .Value
                    End If
                Exit Function
            End If
        
        Next i
        
        .Col = 14
        
        For i = 1 To .MaxRows
            
            .Row = i
            
            If .Text = sChem Then
                .Col = iCol_3
                    If .Value = "" Then
                        subGetChemValue = 0
                    Else
                        subGetChemValue = .Value
                    End If
                Exit Function
            End If
        
        Next i
        
    
    End With
    
    
End Function

Private Function get_CHEM_CEQ(ByVal FOMULA_CD As String, ByVal CHEM_C As Double, ByVal CHEM_SI As Double, ByVal CHEM_MN As Double, _
                          ByVal CHEM_P As Double, ByVal CHEM_S As Double, ByVal CHEM_CR As Double, ByVal CHEM_V As Double, _
                          ByVal CHEM_MO As Double, ByVal CHEM_CU As Double, ByVal CHEM_NI As Double, ByVal CHEM_B As Double) As Double
                          
    Dim V_CEQ As Double
           
    If FOMULA_CD = "A" Then
          V_CEQ = CHEM_C + CHEM_MN / 6
    
    ElseIf FOMULA_CD = "B" Then
          V_CEQ = CHEM_C + CHEM_MN / 6 + (CHEM_CR + CHEM_V + CHEM_MO) / 5 + (CHEM_CU + CHEM_NI) / 15
                                
    
    ElseIf FOMULA_CD = "C" Then
          V_CEQ = CHEM_C + CHEM_MN / 6 + (CHEM_SI / 24) + (CHEM_CR / 5) + (CHEM_MO / 4) + (CHEM_V / 14)
    
    
    ElseIf FOMULA_CD = "D" Then
          V_CEQ = CHEM_C + CHEM_MN / 6 + (CHEM_CU / 40) + (CHEM_NI / 20) + (CHEM_CR / 10) + (CHEM_MO / 50) - (CHEM_V / 10)
    
    
    ElseIf FOMULA_CD = "E" Then
          V_CEQ = CHEM_C + CHEM_SI / 30 + (CHEM_MN / 20) + (CHEM_CU / 20) + (CHEM_CR / 20) + (CHEM_NI / 60) + (CHEM_MO / 15) + (CHEM_V / 10) + (CHEM_B * 5)
    
    Else
          get_CHEM_CEQ = 0
    End If
               
    get_CHEM_CEQ = Round(V_CEQ, 2)

End Function


Private Sub subLineColorSet()
    
    Dim i As Integer
    With ss1
        
        .Col = 7
        
        For i = 1 To .MaxRows
            .Row = i
            .BackColor = &HE1E4CD
        Next i
        
        .Col = 13
        For i = 1 To .MaxRows
            .Row = i
            .BackColor = &HE1E4CD
        Next i
    
    End With
    
End Sub

Private Sub txt_STLGRD_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_STLGRD
        
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
    End If
End Sub
