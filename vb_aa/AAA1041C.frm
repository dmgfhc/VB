VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1041C 
   Caption         =   "生产计划查询/坯料计划录入_AAA1041C"
   ClientHeight    =   8550
   ClientLeft      =   180
   ClientTop       =   1740
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   15000
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_plt 
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
      Tag             =   "工厂"
      Top             =   135
      Width           =   870
   End
   Begin VB.ComboBox cbo_prc 
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
      Left            =   6660
      TabIndex        =   2
      Tag             =   "工序"
      Top             =   135
      Width           =   825
   End
   Begin VB.ComboBox cbo_line 
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
      TabIndex        =   3
      Tag             =   "PRC_LINE"
      Top             =   135
      Width           =   600
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8610
      Left            =   90
      TabIndex        =   4
      Top             =   495
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
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
      MaxCols         =   17
      MaxRows         =   11
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AAA1041C.frx":0000
   End
   Begin InDate.UDate dtp_yy_mm 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Tag             =   "日期"
      Top             =   135
      Width           =   1095
      _ExtentX        =   1931
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
      Left            =   90
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
   End
   Begin InDate.ULabel ULabel4 
      Height          =   300
      Left            =   5310
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "工序"
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
      Height          =   300
      Left            =   2835
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "工厂"
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
   Begin InDate.ULabel ULabel2 
      Height          =   300
      Left            =   7740
      Top             =   135
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "PRC_LINE"
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
End
Attribute VB_Name = "AAA1041C"
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
'-- Program ID        AAA1040C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
'-- Date              2003.7.9
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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
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
    Call Gp_Ms_Collection(dtp_yy_mm, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(cbo_plt, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(cbo_prc, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(cbo_line, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     
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
    Sc1.Add Item:="AAA1040C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    sQuery = "SELECT DISTINCT SUBSTR(CD,1,2) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' "
    Call Gf_ComboAdd(M_CN1, cbo_plt, sQuery)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cbo_plt_Change()

    Dim sQuery As String
    
    sQuery = "SELECT DISTINCT SUBSTR(CD,3,2) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND SUBSTR(CD,3,1) = '" + Mid(cbo_plt.Text, 1, 1) + "' "
    Call Gf_ComboAdd(M_CN1, cbo_prc, sQuery)

End Sub

Private Sub cbo_plt_Click()

    Dim sQuery As String

    If Trim(cbo_plt.Text) = "**" Then
        sQuery = "SELECT DISTINCT SUBSTR(CD,3,2) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND (SUBSTR(CD,3,1)='B' OR SUBSTR(CD,3,1)='*') "
        cbo_line.Clear
        cbo_line.Text = "*"
    Else
        sQuery = "SELECT DISTINCT SUBSTR(CD,3,2) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND SUBSTR(CD,3,1)='" + Mid(cbo_plt.Text, 1, 1) + "' "
    End If
    
    Call Gf_ComboAdd(M_CN1, cbo_prc, sQuery)

End Sub

Private Sub cbo_prc_Change()

    Dim sQuery As String
    
    sQuery = "SELECT DISTINCT SUBSTR(CD,5,1) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND SUBSTR(CD,3,2) = SUBSTR('" + cbo_prc.Text + "', 1, 2) "
    Call Gf_ComboAdd(M_CN1, cbo_line, sQuery)

End Sub

Private Sub cbo_prc_Click()

   Dim sQuery As String
   
   If Trim(cbo_plt.Text) = "**" Then
      cbo_line.Clear
      cbo_line.Text = "*"
   Else
       sQuery = "SELECT DISTINCT SUBSTR(cd,5,1) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND SUBSTR(CD,3,2) = SUBSTR('" + cbo_prc.Text + "', 1, 2) "
       Call Gf_ComboAdd(M_CN1, cbo_line, sQuery)
   End If
   
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
    Call Sp_Setting
    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Screen.MousePointer = vbDefault

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
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Cls()
    
   ' ss1.MaxCols = 0
    ss1.MaxRows = 0
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    rControl(1).SetFocus

End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    
    If dtp_yy_mm.Enabled = False Then
       Exit Sub
    End If
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
        Call Sp_Header_Set
        If Sp_Data_Refer() Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Menu_Setting
            Call Gp_Ms_ControlLock(Mc1("lControl"), True)
        End If
            
    Else
        sMesg = sMesg + " Must input necessarily"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
End Sub

Public Sub Form_Pro()

    If Sp_Process(M_CN1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
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
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

'    Dim dTemp, DCURR As Double
'
'    With ss1
'
'         dTemp = 0
'        .Col = Col
'        .Row = 2
'
'        dTemp = dTemp + IIf(.Value = "", 0, .Value)
'        .Row = 3
'        dTemp = dTemp + IIf(.Value = "", 0, .Value)
'
'        .Row = 4
'        If dTemp <> 0 Then
'           .Text = dTemp
'        End If
'
'        .Row = 5
'
'         DCURR = IIf(.Value = "", 0, .Value)
'         If DCURR <> 0 Then
'            dTemp = dTemp / DCURR
'         End If
'
'        .Row = 6
'        If dTemp <> 0 Then
'           .Text = dTemp
'        End If
'
'         .Row = 7
'        dTemp = dTemp - IIf(.Value = "", 0, .Value)
'
'        .Row = 8
'        dTemp = dTemp + IIf(.Value = "", 0, .Value)
'
'        .Row = 9
'        dTemp = dTemp + IIf(.Value = "", 0, .Value)
'
'        .Row = 10
'        If dTemp <> 0 Then
'           .Text = dTemp
'        End If
'
'        .Row = 11
'
'        DCURR = IIf(.Value = "", 0, .Value)
'        If DCURR <> 0 Then
'            dTemp = dTemp / DCURR
'         End If
'
'        .Row = 12
'        If dTemp <> 0 Then
'           .Text = dTemp
'        End If
'
'        .Row = 13
'        DCURR = IIf(.Value = "", 0, .Value)
'        If DCURR <> 0 Then
'            dTemp = dTemp / DCURR
'        End If
'
'        .Row = 14
'        If dTemp <> 0 Then
'           .Text = dTemp
'        End If
'
'        .Row = 15
'        DCURR = IIf(.Value = "", 0, .Value)
'         If DCURR <> 0 Then
'            dTemp = dTemp / DCURR
'         End If
'
'        .Row = 16
'        If dTemp <> 0 Then
'           .Text = dTemp
'        End If
'
'    End With

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

Public Sub Sp_Setting()

    With ss1

        .ColHeaderRows = 2
        .RowHeaderCols = 2
        
        .RowHeight(SpreadHeader) = 16
        
        .Row = SpreadHeader + 1
     '   .RowHidden = True
        
        .ColWidth(0) = 20
        .ColWidth(SpreadHeader + 1) = 5
     '   .ColWidth(1) = 10
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        .Row = SpreadHeader
        .Col = SpreadHeader
        .Text = "项目\月份"
        .Row = SpreadHeader + 1
        .Col = SpreadHeader
        .Text = "项目\月份"
        
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

Public Sub Sp_Header_Set()

    Dim iCol As Integer
    Dim iRow As Integer
    Dim sMonth As String
    
    With ss1
    
         .Row = SpreadHeader + 1
         
         For iCol = 1 To 12
    '       sMonth = Year(DateAdd("M", iCol - 1, dtp_yy_mm.Text)) & "年" & Month(DateAdd("M", iCol - 1, dtp_yy_mm.Text)) & "月"
            sMonth = Month(DateAdd("M", iCol - 1, dtp_yy_mm.Text)) & "月"
            .Col = iCol
            .Text = sMonth
                  
            'Column Type Setting
            .Col = iCol: .Col2 = iCol
            .Row = 1: .Row2 = -1
            .BlockMode = True
            .CellType = 13      'SS_CELL_TYPE_NUMBER
            .TypeNumberDecPlaces = 2
            .TypeNumberMax = 9999999
            .TypeNumberMin = 0
            .TypeNumberShowSep = True
            .TypeNumberLeadingZero = TypeLeadingZeroNo
            .TypeHAlign = TypeHAlignRight
            .BlockMode = False
            
            .ColWidth(iCol) = 10
            
            .Col = iCol + 1: .Col2 = iCol + 1
            .Row = 1: .Row2 = -1
            .BlockMode = True
            .CellType = 13      'SS_CELL_TYPE_NUMBER
            .TypeNumberDecPlaces = 2
            .TypeNumberMax = 9999999
            .TypeNumberMin = 0
            .TypeNumberShowSep = True
            .TypeNumberLeadingZero = TypeLeadingZeroNo
            .TypeHAlign = TypeHAlignRight
            .BlockMode = False
            
            .ColWidth(iCol + 1) = 10
          Next iCol
          
    End With
    
    ss1.MaxRows = 11
    
    ss1.Row = 1
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '007'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "007"
    
    ss1.Row = 2
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '008'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "008"
    
    ss1.Row = 3
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '009'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "009"
    
    ss1.Row = 4
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '001'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "001"
    
    ss1.Row = 5
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '002'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "002"
    
    ss1.Row = 6
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '003'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "003"
    
    ss1.Row = 7
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '005'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "005"
    
    ss1.Row = 8
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '006'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "006"
    
    ss1.Row = 9
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '012'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "012"
    
    ss1.Row = 10
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '013'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "013"
    
    ss1.Row = 11
    ss1.Col = SpreadHeader
    ss1.Text = Gf_CodeFind(M_CN1, "SELECT CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO = 'A0001' AND CD = '014'")
    ss1.Col = SpreadHeader + 1
    ss1.Text = "014"
               
    ss1.Col = 1
    ss1.Row = 1
    ss1.Col2 = ss1.MaxCols
    ss1.Row2 = 8
    ss1.BlockMode = True
    ss1.Lock = True
    ss1.BlockMode = False
    ss1.Protect = True
    
    ss1.Col = 1
    ss1.Row = 9
    ss1.Col2 = 12
    ss1.Row2 = 11
    ss1.BlockMode = True
    ss1.BackColor = &HC0FFFF
    ss1.BlockMode = False
    ss1.Protect = True
     
    
End Sub

'Public Function Sp_Header_Refer() As Boolean
'
'On Error GoTo SpreadDisplay_Error
'
'    Dim iCol As Integer
'    Dim iRow As Integer
'    Dim iCnt As Integer
'    Dim sQuery As String
'    Dim sEdate As String
'    Dim AdoRs As adodb.Recordset
'    Dim ArrayRecords As Variant
'
'    Dim sQuery2 As String
'
'    Dim AdoRs2 As adodb.Recordset
'    Dim ArrayRecords2 As Variant
'
'    Set AdoRs = New adodb.Recordset
'
'    sQuery = "SELECT THK_CD, FR_THK, TO_THK "
'    sQuery = sQuery + "   FROM BP_THICK_GRP "
'    sQuery = sQuery + "  WHERE PROD_CD = '" + cbo_prod_cd.Text + "' "
'    sQuery = sQuery + "    AND THK_CD <> '*' "
'    sQuery = sQuery + "  ORDER BY THK_CD "
'
'    With ss1
'
'        Sp_Header_Refer = True
'
'        .ReDraw = False
'        .MaxRows = 0:  .MaxCols = 0
'
'        Screen.MousePointer = vbHourglass
'
'        'Ado Execute
'        AdoRs.Open sQuery, M_CN1, adOpenKeyset
'
'        If AdoRs.BOF Or AdoRs.EOF Then
'
'            Sp_Header_Refer = False
'            '.ReDraw = True
'
'            AdoRs.Close
'            Set AdoRs = Nothing
'
'            Screen.MousePointer = vbDefault
'
'            Exit Function
'
'        End If
'
'        ArrayRecords = AdoRs.GetRows
'
'        AdoRs.Close
'        Set AdoRs = Nothing
'
'        If UBound(ArrayRecords, 2) + 1 <> 0 Then
'
'            .MaxCols = (UBound(ArrayRecords, 2) + 1)
'
'            For iCol = 1 To .MaxCols
'
'                .Col = iCol
'                .Row = SpreadHeader
'
'                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
'                    .Text = ""
'                Else
'                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
'                End If
'
'                .Row = SpreadHeader + 1
'                .Text = Trim(ArrayRecords(0, iCnt))
'
'                .ColWidth(iCol) = 15
'
'                .Col = iCol + 2: .Col2 = iCol + 2
'                .Row = 1: .Row2 = -1
'                .BlockMode = True
'                .CellType = 13      'SS_CELL_TYPE_NUMBER
'                .TypeNumberDecPlaces = 2
'                .TypeNumberMax = 9999999
'                .TypeNumberMin = 0
'                .TypeNumberShowSep = True
'                .TypeNumberLeadingZero = TypeLeadingZeroNo
'                .TypeHAlign = TypeHAlignRight
'                .BlockMode = False
'
'                iCnt = iCnt + 1
'            Next iCol
'
'        End If
'
'        .ReDraw = True
'        .Refresh
'
'        Screen.MousePointer = vbDefault
'
'    End With
'
'
'    Set AdoRs2 = New adodb.Recordset
'
'    sQuery2 = "SELECT WID_CD, FR_WID, TO_WID "
'    sQuery2 = sQuery2 + "   FROM BP_WIDTH_GRP "
'    sQuery2 = sQuery2 + "  WHERE PROD_CD = '" + cbo_prod_cd.Text + "' "
' '   sQuery2 = sQuery2 + "    AND WID_CD <> '*' "
'    sQuery2 = sQuery2 + "  ORDER BY WID_CD "
'
'    With ss1
'
'        Sp_Header_Refer = True
'
'     '   .ReDraw = False
'     '   .MaxRows = 0:  .MaxCols = 0
'         .ColWidth(0) = 20
'      '  .ColWidth(1) = 20
'        Screen.MousePointer = vbHourglass
'
'        'Ado Execute
'        AdoRs2.Open sQuery2, M_CN1, adOpenKeyset
'
'        If AdoRs2.BOF Or AdoRs2.EOF Then
'
'            Sp_Header_Refer = False
'            '.ReDraw = True
'
'            AdoRs2.Close
'            Set AdoRs2 = Nothing
'
'            Screen.MousePointer = vbDefault
'
'            Exit Function
'
'        End If
'
'        ArrayRecords2 = AdoRs2.GetRows
'
'        AdoRs2.Close
'        Set AdoRs2 = Nothing
'
'        If UBound(ArrayRecords2, 2) + 1 <> 0 Then
'
'            .MaxRows = (UBound(ArrayRecords2, 2) + 1)
'
'            iCnt = 0
'
'            For iRow = 1 To .MaxRows
'
'                .Row = iRow
'                .Col = SpreadHeader
'
'                If VarType(ArrayRecords2(0, iCnt)) = vbNull Then
'                    .Text = ""
'                Else
'                    .Text = Trim(ArrayRecords2(1, iCnt)) & " ~ " & Trim(ArrayRecords2(2, iCnt)) & "mm"
'                End If
'
'                .Col = SpreadHeader + 1
'                .Text = Trim(ArrayRecords2(0, iCnt))
'
'
'
'                .Row = iRow + 2: .Row2 = iRow + 2
'                .Col = 1: .Col2 = -1
'                .BlockMode = True
'                .CellType = 13      'SS_CELL_TYPE_NUMBER
'                .TypeNumberDecPlaces = 2
'                .TypeNumberMax = 9999999
'                .TypeNumberMin = 0
'                .TypeNumberShowSep = True
'                .TypeNumberLeadingZero = TypeLeadingZeroNo
'                .TypeHAlign = TypeHAlignRight
'                .BlockMode = False
'
'                iCnt = iCnt + 1
'            Next iRow
'
'        End If
'
'        .ReDraw = True
'        .Refresh
'
'        Screen.MousePointer = vbDefault
'
'    End With
'
'Exit Function
'
'SpreadDisplay_Error:
'
'    Set AdoRs = Nothing
'    Set AdoRs2 = Nothing
'    ss1.ReDraw = True
'    Sp_Header_Refer = False
'    Screen.MousePointer = vbDefault
'
'End Function

Public Function Sp_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sTdate As String
    Dim sYear As String
    Dim sQuery As String
    Dim sQuery1 As String
    Dim sEdate, sEdate2 As String
    Dim iEdate As Integer
    Dim iMonth As Integer
    Dim sWID_GRP As String
    Dim sTHK_GRP As String
   ' Dim SPARA As String
    Dim AdoRs As adodb.Recordset
    Dim AdoRs1 As adodb.Recordset
    Dim ArrayRecords As Variant
    Dim ArrayMonth(11) As String
    
    Dim dTemp, DCURR As Double

'    Set AdoRs = New adodb.Recordset
    sYear = Mid(dtp_yy_mm.Text, 1, 4)
    sTdate = Mid(dtp_yy_mm.Text, 6, 2)
    iEdate = Val(sTdate)
    iEdate = iEdate - 1
   
    sEdate2 = dtp_yy_mm.Text & "-01"
 '   sEdate2 = Format(sEdate, "yyyy-mm-dd")
    
'    sEdate = Mid(Format(DateAdd("M", 1, sEdate), "yyyymmdd"), 1, 8) & "01"
    
    For iCnt = 1 To 12
'        sEdate = Format(DateAdd("M", iCol - 1, Format(sEdate2, "yyyy-mm-dd")), "yyyy-mm-dd")
        sEdate = Format(DateAdd("M", iCnt - 1, sEdate2), "yyyy-mm-dd")
        ArrayMonth(iCnt - 1) = Mid(sEdate, 1, 4) & Mid(sEdate, 6, 2)

    Next iCnt
    
'  select plt,
'  sum(case when substr(faci_manage_str,7,2)='01' then faci_manage_tme else 0 end ) as "1",
'  sum(case when substr(faci_manage_str,7,2)='02' then faci_manage_tme else 0 end ) as "2",
'  sum(case when substr(faci_manage_str,7,2)='03' then faci_manage_tme else 0 end ) as "3"
'  From ap_faci_plan
'  where plt='B1'
'  group by plt
  
  
'    sQuery = "SELECT "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(0) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(1) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(2) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(3) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(4) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(5) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(6) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(7) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(8) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(9) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(10) + "' then faci_manage_tme else 0 end ), "
'    sQuery = sQuery + "  sum(case when substr(faci_manage_str,1,6)='" + ArrayMonth(11) + "' then faci_manage_tme else 0 end ) "
'    sQuery = sQuery + " FROM ap_faci_plan  "
'    sQuery = sQuery + " WHERE  "
'    sQuery = sQuery + "   substr(faci_manage_str,1,4)='" + sYear + "'"
      
'    Debug.Print sQuery
    
'    With ss1
'
'        Sp_Data_Refer = True
'
'        .ReDraw = False
'       ' .MaxRows = 0
'
'        Screen.MousePointer = vbHourglass
'
'        'Ado Execute
'        AdoRs.Open sQuery, M_CN1, adOpenKeyset
'
'        If AdoRs.BOF Or AdoRs.EOF Then
'
'            Sp_Data_Refer = False
'            .ReDraw = True
'
'            AdoRs.Close
'            Set AdoRs = Nothing
'
'            Screen.MousePointer = vbDefault
'
'            Exit Function
'
'        End If
'
'        ArrayRecords = AdoRs.GetRows
'
'        AdoRs.Close
'        Set AdoRs = Nothing
'
'        If UBound(ArrayRecords, 2) + 1 <> 0 Then
'
'                .Row = 1
'
'                For iCol = 1 To 12
'                    .Col = iCol
'
'                        If VarType(ArrayRecords(iCol - 1, 0)) = vbNull Then
'                            .Text = ""
'                        Else
'                             If Val((ArrayRecords(iCol - 1, 0))) <> 0 Then
'                                .Text = Trim(ArrayRecords(iCol - 1, 0))
'                             End If
'                        End If
'
'                Next iCol
'
'        End If
'
'     '   .ReDraw = True
'
'        Screen.MousePointer = vbDefault
'
'    End With
    
    Set AdoRs1 = New adodb.Recordset
    
    sQuery1 = "SELECT    APLY_ITEM, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(0) + "' then plan_value else 0 end ) as m1, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(1) + "' then plan_value else 0 end ) as m2, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(2) + "' then plan_value else 0 end ) as m3, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(3) + "' then plan_value else 0 end ) as m4, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(4) + "' then plan_value else 0 end ) as m5, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(5) + "' then plan_value else 0 end ) as m6, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(6) + "' then plan_value else 0 end ) as m7, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(7) + "' then plan_value else 0 end ) as m8, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(8) + "' then plan_value else 0 end ) as m9, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(9) + "' then plan_value else 0 end ) as m10, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(10) + "' then plan_value else 0 end ) as m11, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(11) + "' then plan_value else 0 end ) as m12 "
    sQuery1 = sQuery1 + " FROM AP_PROD_PLAN "
    sQuery1 = sQuery1 + " WHERE PLT      = '" + cbo_plt.Text + "' "
    sQuery1 = sQuery1 + "   AND PRC      = '" + cbo_prc.Text + "' "
    sQuery1 = sQuery1 + "   AND PRC_LINE = '" + cbo_line.Text + "' "
    sQuery1 = sQuery1 + "   AND APLY_ITEM IN ('001','002','003','005','006','007','008','009','012','012','013','014') "
    sQuery1 = sQuery1 + " GROUP BY APLY_ITEM  "
    sQuery1 = sQuery1 + " ORDER BY APLY_ITEM  "
    
    With ss1

        Sp_Data_Refer = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs1.Open sQuery1, M_CN1, adOpenKeyset
        
        If AdoRs1.BOF Or AdoRs1.EOF Then
        
            Sp_Data_Refer = False
            .ReDraw = True
            AdoRs1.Close
            Set AdoRs1 = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs1.GetRows
        AdoRs1.Close
        
        sEdate = Mid(dtp_yy_mm.Text, 6, 2)
        iMonth = Val(sEdate) - 1
        
        If UBound(ArrayRecords, 2) + 1 <> 0 Then
            
            For iCnt = 0 To UBound(ArrayRecords, 2)
            
                Select Case Trim(ArrayRecords(0, iCnt))
                    Case "001"
                        .Row = 4
                    Case "002"
                        .Row = 5
                    Case "003"
                        .Row = 6
                    Case "005"
                        .Row = 7
                    Case "006"
                        .Row = 8
                    Case "007"
                        .Row = 1
                    Case "008"
                        .Row = 2
                    Case "009"
                        .Row = 3
                    Case "012"
                        .Row = 9
                    Case "013"
                        .Row = 10
                    Case "014"
                        .Row = 11
                End Select
                
                For iCol = 1 To 12
                    .Col = iCol
                    If VarType(ArrayRecords(iCol, iCnt)) = vbNull Or Val((ArrayRecords(iCol, iCnt))) = 0 Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iCol, iCnt))
                    End If
                Next iCol
                
            Next iCnt
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
'    sQuery1 = "SELECT "
'    sQuery1 = sQuery1 + " aply_item,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(0) + "' then plan_value else 0 end ) as m1, "
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(1) + "' then plan_value else 0 end ) as m2,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(2) + "' then plan_value else 0 end ) as m3,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(3) + "' then plan_value else 0 end ) as m4,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(4) + "' then plan_value else 0 end ) as m5,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(5) + "' then plan_value else 0 end ) as m6,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(6) + "' then plan_value else 0 end ) as m7,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(7) + "' then plan_value else 0 end ) as m8,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(8) + "' then plan_value else 0 end ) as m9,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(9) + "' then plan_value else 0 end ) as m10,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(10) + "' then plan_value else 0 end ) as m11,"
'    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(11) + "' then plan_value else 0 end ) as m12 "
'    sQuery1 = sQuery1 + " FROM ap_prod_plan "
'    sQuery1 = sQuery1 + " WHERE "
'    sQuery1 = sQuery1 + " aply_item in ('015','016') "
'    sQuery1 = sQuery1 + " GROUP BY aply_item  "
'    sQuery1 = sQuery1 + " ORDER BY aply_item  "
'
'      With ss1
'
'        Sp_Data_Refer = True
'
'        .ReDraw = False
'       ' .MaxRows = 0
'
'        Screen.MousePointer = vbHourglass
'
'        'Ado Execute
'        AdoRs1.Open sQuery1, M_CN1, adOpenKeyset
'
'        If AdoRs1.BOF Or AdoRs1.EOF Then
'
'            Sp_Data_Refer = False
'            .ReDraw = True
'
'            AdoRs1.Close
'            Set AdoRs1 = Nothing
'
'            Screen.MousePointer = vbDefault
'
'            Exit Function
'
'        End If
'
'        ArrayRecords = AdoRs1.GetRows
'
'        AdoRs1.Close
'      '  Set AdoRs1 = Nothing
'
'        sEdate = Mid(dtp_yy_mm.Text, 6, 2)
'        iMonth = Val(sEdate) - 1
'      '  iEdate = Val(sEdate) - 1
'
'        If UBound(ArrayRecords, 2) + 1 <> 0 Then
'
'            iRow = 2
'            For iCnt = 0 To UBound(ArrayRecords, 2)
'                .Row = iRow
'                For iCol = 1 To 12
'                    .Col = iCol
'                    If VarType(ArrayRecords(iCol, iCnt)) = vbNull Or Val((ArrayRecords(iCol, iCnt))) = 0 Then
'                        .Text = ""
'                    Else
'                        .Text = Trim(ArrayRecords(iCol, iCnt))
'                    End If
'
'                Next iCol
'                iRow = iRow + 1
'
'            Next iCnt
'
'        End If
'
'
'     '   .ReDraw = True
'
'        Screen.MousePointer = vbDefault
'
'    End With
'
'
'
'    sQuery1 = "SELECT "
'    sQuery1 = sQuery1 + " prc, "
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(0) + "' then plan_value else 0 end  as m1, "
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(1) + "' then plan_value else 0 end  as m2,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(2) + "' then plan_value else 0 end  as m3,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(3) + "' then plan_value else 0 end  as m4,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(4) + "' then plan_value else 0 end  as m5,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(5) + "' then plan_value else 0 end  as m6,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(6) + "' then plan_value else 0 end  as m7,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(7) + "' then plan_value else 0 end  as m8,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(8) + "' then plan_value else 0 end  as m9,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(9) + "' then plan_value else 0 end  as m10,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(10) + "' then plan_value else 0 end  as m11,"
'    sQuery1 = sQuery1 + "case when year_month='" + ArrayMonth(11) + "' then plan_value else 0 end  as m12 "
'    sQuery1 = sQuery1 + " FROM ap_prod_plan "
'    sQuery1 = sQuery1 + " WHERE "
'    sQuery1 = sQuery1 + " aply_item='002' "
'    sQuery1 = sQuery1 + " AND PLT='**'"
'    sQuery1 = sQuery1 + " AND PRC IN ('BC','BD','BE','BF')"
'    sQuery1 = sQuery1 + " AND PRC_LINE='*'"
'    sQuery1 = sQuery1 + " AND PROD_CD='**'"
'    sQuery1 = sQuery1 + " AND STLGRD='***********'"
'    sQuery1 = sQuery1 + " AND THK_GRP='*'"
'    sQuery1 = sQuery1 + " AND WID_GRP='*'"
'      With ss1
'
'        Sp_Data_Refer = True
'
'        .ReDraw = False
'       ' .MaxRows = 0
'
'        Screen.MousePointer = vbHourglass
'
'        'Ado Execute
'        AdoRs1.Open sQuery1, M_CN1, adOpenKeyset
'
'        If AdoRs1.BOF Or AdoRs1.EOF Then
'
'            Sp_Data_Refer = False
'            .ReDraw = True
'
'            AdoRs1.Close
'            Set AdoRs1 = Nothing
'
'            Screen.MousePointer = vbDefault
'
'            Exit Function
'
'        End If
'
'        ArrayRecords = AdoRs1.GetRows
'
'        AdoRs1.Close
'        Set AdoRs1 = Nothing
'
'        sEdate = Mid(dtp_yy_mm.Text, 6, 2)
'        iMonth = Val(sEdate) - 1
'      '  iEdate = Val(sEdate) - 1
'
'        If UBound(ArrayRecords, 2) + 1 <> 0 Then
'
'           ' iRow = 2
'            For iCnt = 0 To UBound(ArrayRecords, 2)
'                Select Case ArrayRecords(0, iCnt)
'                       Case "BF"
'                            .Row = 5
'                       Case "BE"
'                            .Row = 11
'                       Case "BD"
'                            .Row = 13
'                       Case "BC"
'                            .Row = 15
'                End Select
'
'              '  .Row = iRow
'
'                iEdate = iMonth
'                For iCol = 1 To 12
'                    .Col = iCol
'                    If VarType(ArrayRecords(iCol, iCnt)) = vbNull Or Val((ArrayRecords(iCol, iCnt))) = 0 Then
'                        .Text = ""
'                    Else
'                        .Text = Trim(ArrayRecords(iCol, iCnt))
'                    End If
'
'
'                Next iCol
'             '   iRow = iRow + 1
'
'            Next iCnt
'
'        End If
'
'
'     '   .ReDraw = True
'
'        Screen.MousePointer = vbDefault
'
'    End With
   
    
'    With ss1
'         For iCol = 1 To 12
'                dTemp = 0
'               .Col = iCol
'               .Row = 2
'
'               dTemp = dTemp + IIf(.Value = "", 0, .Value)
'               .Row = 3
'               dTemp = dTemp + IIf(.Value = "", 0, .Value)
'               .Row = 4
'               If dTemp <> 0 Then
'                  .Text = dTemp
'               End If
'               .Row = 5
'
'                DCURR = IIf(.Value = "", 0, .Value)
'                If DCURR <> 0 Then
'                   dTemp = dTemp / DCURR
'                End If
'               .Row = 6
'               If dTemp <> 0 Then
'                  .Text = dTemp
'               End If
'                .Row = 7
'               dTemp = dTemp - IIf(.Value = "", 0, .Value)
'               .Row = 8
'               dTemp = dTemp + IIf(.Value = "", 0, .Value)
'               .Row = 9
'               dTemp = dTemp + IIf(.Value = "", 0, .Value)
'               .Row = 10
'               If dTemp <> 0 Then
'                  .Text = dTemp
'               End If
'
'               .Row = 11
'
'               DCURR = IIf(.Value = "", 0, .Value)
'               If DCURR <> 0 Then
'                   dTemp = dTemp / DCURR
'                End If
'               .Row = 12
'               If dTemp <> 0 Then
'                  .Text = dTemp
'               End If
'
'               .Row = 13
'               DCURR = IIf(.Value = "", 0, .Value)
'               If DCURR <> 0 Then
'                   dTemp = dTemp / DCURR
'                End If
'               .Row = 14
'               If dTemp <> 0 Then
'                  .Text = dTemp
'               End If
'
'               .Row = 15
'               DCURR = IIf(.Value = "", 0, .Value)
'                If DCURR <> 0 Then
'                   dTemp = dTemp / DCURR
'                End If
'               .Row = 16
'               If dTemp <> 0 Then
'                  .Text = dTemp
'               End If
'         Next iCol
'
'    End With
    
    MDIMain.StatusBar1.Panels(1) = "Message : Data inquiry completed"
    Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Process(Conn As adodb.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim sTdate As String
    Dim sEdate As String
    Dim sYear As String
    Dim sMesg As String
    Dim sTemp As String
    Dim sPara As String
    
    Dim iEdate As Integer
    Dim iMonth As Integer
    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    
    Dim adoCmd As adodb.Command

    sEdate = Mid(dtp_yy_mm.Text, 1, 4)
    sYear = sEdate
    sTdate = Mid(dtp_yy_mm.Text, 6, 2)
    iMonth = Val(sTdate)
    iEdate = Val(sTdate)

    Sp_Process = True
    
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
        Set adoCmd = New adodb.Command
        
        Set adoCmd.ActiveConnection = Conn
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = Sc.Item("P-M")
        
        Conn.BeginTrans
        
        'Ceate Parameter (Input) iType + iColumn
        For iCount = 1 To 11
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        For iRow = 9 To 11                'input ->row 9, 10, 11
            
            .Row = iRow
            iEdate = iMonth
            sEdate = sYear
            
            'Parameters Setting
            For iCol = 1 To 12           'month 1->12
                 .Col = SpreadHeader + 1
                 adoCmd.Parameters(8).Value = .Text                          'aply_item
                 
                 .Col = iCol
                 If Trim(.Text) <> "" Then
               
                     If Trim(.Text) = "" Then                                'plan_value
                         adoCmd.Parameters(9).Value = 0
                     Else
                         dTempInt = .Value
                         adoCmd.Parameters(9).Value = dTempInt
                     End If
                     
                     adoCmd.Parameters(10).Value = sUserID                    'User-id
                     
                     If iEdate < 10 Then                                      'year_month
                        adoCmd.Parameters(0).Value = sEdate + "0" + LTrim(Str(iEdate))
                     Else
                        adoCmd.Parameters(0).Value = sEdate + LTrim(Str(iEdate))
                     End If
                     
                     adoCmd.Parameters(1).Value = cbo_plt.Text                'plt
                     adoCmd.Parameters(2).Value = cbo_prc.Text                'prc
                     adoCmd.Parameters(3).Value = cbo_line.Text               'prc_line
                     adoCmd.Parameters(4).Value = "**"                        'prod_cd
                     adoCmd.Parameters(5).Value = "***********"               'stlgrd
                     adoCmd.Parameters(6).Value = "*"                         'thk_grp
                     adoCmd.Parameters(7).Value = "*"                         'wid_grp
                     
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
                
                iEdate = iEdate + 1
                If iEdate = 13 Then
                   iEdate = 1
                   sEdate = LTrim(Str(Val(sEdate) + 1))
                End If
            
            Next iCol
            
        Next iRow
        
        Conn.CommitTrans
        .ReDraw = True
        MDIMain.StatusBar1.Panels(1) = "Message : Data update completed"
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

