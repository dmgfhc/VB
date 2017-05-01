VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ABX1010C 
   Caption         =   "订单主要标准录入_ABX1010C"
   ClientHeight    =   7620
   ClientLeft      =   210
   ClientTop       =   2535
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_CLASS_FL 
      Height          =   345
      Left            =   10710
      TabIndex        =   5
      Top             =   45
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.ComboBox Cbo_plt 
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
      Left            =   1275
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   105
      Width           =   1260
   End
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "订单接受条件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8340
      TabIndex        =   4
      Top             =   90
      Width           =   1605
   End
   Begin VB.OptionButton Opt1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "设备限制条件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6375
      TabIndex        =   2
      Top             =   90
      Width           =   1605
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   225
      Top             =   105
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "工厂"
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
      Left            =   3135
      Top             =   105
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin VB.ComboBox Cbo_prod_cd 
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
      Left            =   4185
      TabIndex        =   1
      Tag             =   "产品"
      Top             =   105
      Width           =   1215
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8610
      Left            =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   510
      Width           =   15045
      _Version        =   393216
      _ExtentX        =   26538
      _ExtentY        =   15187
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
      MaxCols         =   24
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ABX1010C.frx":0000
   End
End
Attribute VB_Name = "ABX1010C"
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
'-- Program ID        ABX1010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
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
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    Dim sQuery   As String

   'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(Cbo_plt, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(Cbo_prod_cd, "p", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(TXT_CLASS_FL, "P", " ", " ", "i", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ABX1010C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="ABX1010C.P_REFER", Key:="P-R"
    sc1.Add Item:="ABX1010C.P_ONEROW", Key:="P-O"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  EDIT  End      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sQuery = "SELECT CD FROM ZP_CD WHERE CD_MANA_NO='B0005'"
    Call Gf_ComboAdd(M_CN1, Cbo_prod_cd, sQuery)
    
    sQuery = "SELECT CD FROM ZP_CD WHERE CD_MANA_NO='C0001'"
    Call Gf_ComboAdd(M_CN1, Cbo_plt, sQuery)
    
    ss1.Col = 19
    ss1.ColHidden = True
    ss1.Col = 23
    ss1.ColHidden = True
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Cbo_plt_Change()
   
   If Mid(Cbo_plt.Text, 1, 1) = "B" Then
      Cbo_prod_cd.Clear
      Cbo_prod_cd.Text = "SL"
   Else
      If Cbo_prod_cd.Text <> "PP" And Cbo_prod_cd.Text <> "HC" Then
         Cbo_prod_cd.Clear
         Cbo_prod_cd.AddItem ("PP")
         Cbo_prod_cd.AddItem ("HC")
      End If
   End If
   
End Sub

Private Sub Cbo_plt_Click()

   If Mid(Cbo_plt.Text, 1, 1) = "B" Then
      Cbo_prod_cd.Clear
      Cbo_prod_cd.Text = "SL"
   Else
      If Cbo_prod_cd.Text <> "PP" And Cbo_prod_cd.Text <> "HC" Then
         Cbo_prod_cd.Clear
         Cbo_prod_cd.AddItem ("PP")
         Cbo_prod_cd.AddItem ("HC")
      End If
   End If
   
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "B-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    Opt1.Value = True
    TXT_CLASS_FL.Text = "FC"
    
'    Call Gp_Sp_HdColColor(ss1, 1)
'    Call Gp_Sp_HdColColor(ss1, 2)
    Call Gp_Sp_HdColColor(ss1, 5)


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "B-System.INI", Me.Name)
    
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

    Dim sQuery As String
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If
    
    sQuery = "SELECT CD FROM ZP_CD WHERE CD_MANA_NO='B0005'"
    Call Gf_ComboAdd(M_CN1, Cbo_prod_cd, sQuery)
    
    sQuery = "SELECT CD FROM ZP_CD WHERE CD_MANA_NO='C0001'"
    Call Gf_ComboAdd(M_CN1, Cbo_plt, sQuery)
    
    Opt1.Value = True
    Opt1.Enabled = True
    Opt2.Enabled = True
    
    TXT_CLASS_FL.Text = "FC"
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim SMESG As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Opt1.Enabled = False
        Opt2.Enabled = False
    End If
    
     If Trim(Cbo_plt.Text) = "" Or Trim(Cbo_prod_cd.Text) = "" Then
        Call Gp_Sp_BlockLock(ss1, 5, 16, 1, -1, True)
        MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
        MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
     Else
        Call Gp_Sp_BlockLock(ss1, 5, 16, 1, -1, False)
        MDIMain.MenuTool.Buttons(7).Enabled = True    'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = True    'Row delete
        MDIMain.MenuTool.Buttons(9).Enabled = True    'Row cancel
         
     End If

            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    Dim sProd_cd As String
    Dim iRow As Integer
    
    If Trim(Cbo_plt.Text) = "" Or Trim(Cbo_prod_cd.Text) = "" Then Exit Sub

    With ss1
    
         For iRow = 1 To .MaxRows
         
             .Row = iRow
             .Col = 2
             sProd_cd = .Text
    
             If sProd_cd = "HC" Then
      
                .Col = 13
                .Text = 0
                .Lock = True
                .Col = 14
                .Text = 0
                .Lock = True
             Else
                .Col = 13
                If .Value = "" Then
                   Call Gp_MsgBoxDisplay("最小长度必须输入...")
                   Exit Sub
                End If
                .Col = 14
                If .Value = "" Then
                   Call Gp_MsgBoxDisplay("最大长度必须输入...")
                   Exit Sub
                End If
             
             
             End If
             
         Next iRow
         
     End With
     
     Call Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1)
      
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    
    If Cbo_plt.Text <> "" Then
       ss1.Col = 1
       ss1.Text = Cbo_plt
       ss1.Lock = True
    End If
    
    If Cbo_prod_cd.Text <> "" Then
       ss1.Col = 2
       ss1.Text = Cbo_prod_cd
       ss1.Lock = True
    End If

    If TXT_CLASS_FL.Text = "FC" Then
       ss1.Col = 3
       ss1.Text = "FC"
    Else
       ss1.Col = 3
       ss1.Text = "OC"
    End If
    
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
    
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

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

Private Sub Opt1_Click()

    Opt2.Value = False
    Opt1.Value = True

    TXT_CLASS_FL.Text = "FC"

End Sub

Private Sub Opt2_Click()

    Opt1.Value = False
    Opt2.Value = True
    
    TXT_CLASS_FL.Text = "OC"

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

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
    
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 23)
    
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

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    Dim sfac_cd As String

    ss1.Col = 2
    sfac_cd = ss1.Text
    If sfac_cd = "HC" Then
        If Cbo_plt.Text <> "" And Cbo_prod_cd.Text <> "" Then
           ss1.Col = 13
           ss1.Text = 0
           ss1.Lock = True
    
           ss1.Col = 14
           ss1.Text = 0
           ss1.Lock = True
        End If
       
    Else
       If Cbo_plt.Text <> "" And Cbo_prod_cd.Text <> "" Then
            ss1.Col = 13
            ss1.Lock = False
            ss1.Col = 14
            ss1.Lock = False
       End If
    End If
    
'   Min<=Max check

    Dim iCol As Integer
    Dim iRow As Integer
    Dim dMin As Double
    Dim dMax As Double
    
    If Row < 0 Or Row = 0 Then Exit Sub
    
    With ss1
    
        If .CellTag = "False" Then Exit Sub
            
        .Row = Row
        Select Case Col
        
            Case 8, 10, 12, 14, 16  'MAX
            
                .Col = Col - 1
                If .Value = "" Then
                    dMin = 0
                Else
                    dMin = Val(.Value)
                End If
                
                .Col = Col
                If .Value = "" Then
                    dMax = 0
                Else
                    dMax = Val(.Value)
                End If
                                
                If dMin = 0 Then Exit Sub
                
                If dMax <> 0 Then
                
                    If dMax < dMin Then
                        .Col = Col
                        .Row = Row
                        .CellTag = "False"
                     
                        Call Gp_MsgBoxDisplay("最大值应大于最小值...")
                      
                        .Col = Col
                        .Row = Row
                        .CellTag = ""
                        .Value = 0
                        .TabStop = True
                        .SetFocus
                        .SetActiveCell Col, Row
                        .Action = SS_ACTION_ACTIVE_CELL
                        .EditMode = True
                        .TabStop = False
                    End If
                    
                 End If
           
            Case 7, 9, 11, 13, 15 'MIN
                
                .Col = Col
                If .Value = "" Then
                    dMin = 0
                Else
                    dMin = Val(.Value)
                End If
                
                .Col = Col + 1
                
                If .Value = "" Then
                    dMax = 0
                Else
                    dMax = Val(.Value)
                End If
                                
                If dMax = 0 Then Exit Sub
                
                If dMin <> 0 Then
                
                    If dMax < dMin Then
                     
                      .Col = Col
                        .Row = Row
                        .CellTag = "False"
                        Call Gp_MsgBoxDisplay("最大值应大于最小值...")
                        .Col = Col
                        .Row = Row
                        .CellTag = ""
                        .Value = 0
                        .TabStop = True
                        .SetFocus
                        .SetActiveCell Col, Row
                        .Action = SS_ACTION_ACTIVE_CELL
                        .EditMode = True
                        .TabStop = False
                    End If
                    
                End If
                
        End Select
            
   End With

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol
    
'        Case 1
'
'            If KeyCode = vbKeyF4 Then
'
'                Set DD.sPname = Me.ss1
'
'                DD.sWitch = "SP"
'                DD.sKey = "C0001"
'                DD.rControl.Add Item:=1
'
'                DD.nameType = "2"
'                Call Gf_Common_DD(M_CN1, KeyCode)
'
'            End If
'
'        Case 2
'
'            If KeyCode = vbKeyF4 Then
'
'                Set DD.sPname = Me.ss1
'
'                DD.sWitch = "SP"
'                DD.sKey = "B0005"
'                DD.rControl.Add Item:=2
'
'                DD.nameType = "2"
'                Call Gf_Common_DD(M_CN1, KeyCode)
'
'            End If
            
         Case 5
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.rControl.Add Item:=5
                
                Call Gf_StdSPEC_DD(M_CN1, KeyCode)
                
            End If
            
    End Select
    
End Sub

