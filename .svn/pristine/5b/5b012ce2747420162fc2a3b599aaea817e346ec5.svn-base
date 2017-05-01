VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AHD0110C 
   Caption         =   "中厚板卷厂板材产量_AHD0110C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_plt_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   6930
      MaxLength       =   40
      TabIndex        =   5
      Tag             =   "mill_plt"
      Top             =   645
      Width           =   2505
   End
   Begin VB.TextBox txt_plt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   6420
      MaxLength       =   2
      TabIndex        =   4
      Tag             =   "plt"
      Top             =   645
      Width           =   495
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8355
      Left            =   180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   990
      Width           =   15000
      _Version        =   393216
      _ExtentX        =   26458
      _ExtentY        =   14737
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
      MaxCols         =   8
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AHD0110C.frx":0000
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   180
      Top             =   630
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
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
   Begin InDate.UDate dtp_yy_mm 
      Height          =   315
      Left            =   1425
      TabIndex        =   2
      Tag             =   "日期"
      Top             =   645
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
      MaxLength       =   10
   End
   Begin InDate.UDate dtp_yy_mm2 
      Height          =   315
      Left            =   3150
      TabIndex        =   3
      Tag             =   "日期"
      Top             =   645
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
      MaxLength       =   10
   End
   Begin InDate.ULabel ULabel8 
      Height          =   300
      Left            =   2850
      Top             =   645
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      Caption         =   "至"
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
   Begin InDate.ULabel ULabel01 
      Height          =   315
      Index           =   14
      Left            =   5025
      Top             =   645
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "生产厂"
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
   Begin VB.Label Label1 
      Caption         =   "中厚板卷厂板材产量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5760
      TabIndex        =   0
      Top             =   105
      Width           =   3525
   End
End
Attribute VB_Name = "AHD0110C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name
'-- Sub_System Name
'-- Program Name
'-- Program ID        ABY1020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          ZHANGLIN
'-- Coder             ZHANGLIN
'-- Date              2005.9.23
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

Dim iSumCol As New Collection       'Sum Column

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
      Call Gp_Ms_Collection(dtp_yy_mm, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(dtp_yy_mm2, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_plt_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AHD0110C.P_REFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 8, True)
     
    'Duplicate Count
'    iDupCnt = 3
    iDupCnt = 1
    'Sum Column Count
    iSumCnt = 1
    
    'Sum Column Setting
    iSumCol.Add Item:=7
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub
Private Sub dtp_YEAR_MONTH_Change()

'    Label1.Caption = Mid(dtp_YEAR_MONTH.Value, 1, 7)
'    dtp_YEAR_MONTH1.Text = Mid(dtp_YEAR_MONTH.Value, 1, 7)
'    Call Form_Ref
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
    
'    sAuthority = Gf_Pgm_Authority(Me.Name, True)
'
    Call Form_Define
'
'    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "i-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
    
'    dtp_yy_mm.RawData = Gf_DTSet(M_CN1, "M")
'    Label1.Caption = dtp_YEAR_MONTH1.Text
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "H-System.INI", Me.Name)
    
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
    
'    Set iSumCol = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    Set iSumCol = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call subButtonHide
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
  '      rControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    Dim sQuery As String
    sQuery = "{ CALL " + "AHD0110C.P_REFER" + "("
    sQuery = sQuery + " '" + dtp_yy_mm.RawData + "','" + dtp_yy_mm2.RawData + "','" + txt_plt.Text + "'"
    sQuery = sQuery + ")"
    sQuery = sQuery + "}"
'    sQuery = "SELECT 'PP'                                                                          " & vbCrLf
'    sQuery = sQuery & ",GF_ENDUSE_NAME('P',T.ENDUSE_CD)                                            " & vbCrLf
'    sQuery = sQuery & ",GF_STLGRD_DETAIL(T.STLGRD)                                                 " & vbCrLf
'    sQuery = sQuery & ",T.THK||'*'|| T.WID || '*' ||T.LEN                                          " & vbCrLf
'    sQuery = sQuery & ",DECODE(T.SIZE_KND,'01','双定尺','02','单定尺','毛边')                      " & vbCrLf
'    sQuery = sQuery & ",GET_PROD_GRD_NAME(T.PROD_GRD)                                              " & vbCrLf
'    sQuery = sQuery & ",SUM(T.WGT)                                                                 " & vbCrLf
'    sQuery = sQuery & "FROM GP_PLATE T                                                             " & vbCrLf
'    sQuery = sQuery & "WHERE T.IN_PLT_DATE LIKE '" + dtp_yy_mm.RawData + "'|| '%'                  " & vbCrLf
'    sQuery = sQuery & "GROUP BY T.ENDUSE_CD,T.STLGRD,T.THK,T.WID,T.LEN,T.SIZE_KND,T.PROD_GRD"
'
'    sQuery = sQuery & " UNION ALL                                                                  " & vbCrLf
'    sQuery = sQuery & "SELECT 'CC'                                                                 " & vbCrLf
'    sQuery = sQuery & ",GF_ENDUSE_NAME('P',T.ENDUSE_CD)                                            " & vbCrLf
'    sQuery = sQuery & ",GF_STLGRD_DETAIL(T.STLGRD)                                                 " & vbCrLf
'    sQuery = sQuery & ",T.THK||'*'|| T.WID || '*' ||T.LEN                                          " & vbCrLf
'    sQuery = sQuery & ",DECODE(T.SIZE_KND,'01','双定尺','02','单定尺','毛边')                      " & vbCrLf
'    sQuery = sQuery & ",GET_PROD_GRD_NAME(T.PROD_GRD)                                              " & vbCrLf
'    sQuery = sQuery & ",SUM(T.WGT)                                                                 " & vbCrLf
'    sQuery = sQuery & "FROM GP_PLATE T                                                             " & vbCrLf
'    sQuery = sQuery & "WHERE T.IN_PLT_DATE LIKE '" + dtp_yy_mm.RawData + "'|| '%'                  " & vbCrLf
'    sQuery = sQuery & "GROUP BY T.ENDUSE_CD,T.STLGRD,T.THK,T.WID,T.LEN,T.SIZE_KND,T.PROD_GRD"

    
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If dtp_yy_mm.RawData = "" Or dtp_yy_mm2.RawData = "" Then
       Call Gp_MsgBoxDisplay("请输入日期", "I")
       Exit Sub
    End If
    
    If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, iDupCnt, iSumCnt, iSumCol) Then
'    If Gf_Sp_Display(M_CN1, ss1, sQuery) Then
'    If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, 1, 4, iSumCnt, iSumCol, False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call subButtonHide
'        Call Sp_AutoInsertSum
'        Call Sp_AutoInsertSumGroup
    End If
    

    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
'    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    
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
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, ss1.MaxCols - 1, lBlkrow1, lBlkrow1)

End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub


Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0

End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
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

    MDIMain.MenuTool.Buttons(4).Enabled = False    'Save
    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    
    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_name

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
        txt_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_name.Text = ""
    End If

End Sub




'Private Sub Sp_AutoInsertSum()
'    Dim dValue As Double
'    Dim iCount As Integer
'    Dim strProdTag As String
'    Dim strProdGrd1 As String
'    Dim strProdGrd2 As String
'    Dim iRow As Integer
'    Dim iCurRow As Integer
'    Dim irow2 As Integer
'    Dim bLoop As Boolean
'    bLoop = True
'    With ss1
'        irow2 = 1
'        If .MaxRows < 2 Then Exit Sub
'        While bLoop
'
'            If irow2 >= .MaxRows Then
'                bLoop = False
'                Exit Sub
'            End If
'            'bLoop = False
'            iRow = irow2
'            dValue = 0
'            For iCurRow = iRow To .MaxRows
'                .Col = 1: .Row = iCurRow: strProdTag = .Text
'                .Col = 2: strProdGrd1 = .Text
'                .Row = iCurRow + 1: strProdGrd2 = .Text
'                .Col = 1
'                If .Text = strProdTag And strProdGrd1 = strProdGrd2 Then
'                    irow2 = irow2 + 1
'                Else
'                    .MaxRows = .MaxRows + 1
'
'                    .Row = iCurRow + 1
'                    .Action = SS_ACTION_INSERT_ROW
'                    .Col = 0: .Text = "∑"
'                    Call .AddCellSpan(1, .Row, 3, 1)
'                    .Col = 1: .Text = strProdTag & strProdGrd1 & "总计"
'                    For iCount = 1 To .MaxCols
'                        .Col = iCount
'                        If .CellType = SS_CELL_TYPE_COMBOBOX Then .Value = 0
'                    Next iCount
'                    Call Gp_Sp_RowColor(ss1, .Row, vbBlue)
'
'                    dValue = Sp_SumAbove(ss1, 7, iRow, irow2)
'                    .Row = iCurRow + 1
'                    .Col = 7: .Value = CStr(IIf(dValue > 0, dValue, ""))
'                    irow2 = irow2 + 2
'                    Exit For
'                End If
'            Next iCurRow
'        Wend
'    End With
'
'End Sub

'Private Sub Sp_AutoInsertSum()
'    Dim dValue As Double
'    Dim iCount As Integer
'    Dim x As Integer
'    Dim strProdTag As String
'    Dim iRow As Integer
'    Dim iCurRow As Integer
'    Dim irow2 As Integer
'    Dim bLoop As Boolean
'    bLoop = True
'
'    With ss1
'        irow2 = 1
'        If .MaxRows < 2 Then Exit Sub
'        While bLoop
'
'            If irow2 >= .MaxRows Then
'                bLoop = False
'                Exit Sub
'            End If
'            'bLoop = False
'            iRow = irow2
'            dValue = 0
'            For iCurRow = iRow To .MaxRows
'                .Col = 1: .Row = iCurRow: strProdTag = .Text
'                .Row = iCurRow + 1
'                If .Text = strProdTag Then
'                    irow2 = irow2 + 1
'                Else
'                    .MaxRows = .MaxRows + 1
'
'                    .Row = iCurRow + 1
'                    .Action = SS_ACTION_INSERT_ROW
'                    .Col = 0: .Text = "∑"
'                    .Col = 1: .Text = strProdTag & " 合计"
'                    For iCount = 1 To .MaxCols
'                        .Col = iCount
'                        If .CellType = SS_CELL_TYPE_COMBOBOX Then .Value = 0
'                    Next iCount
'                    Call Gp_Sp_RowColor(ss1, .Row, vbBlue)
'
'                    dValue = Sp_SumAbove(ss1, 7, iRow, irow2)
'                    .Row = iCurRow + 1
'                    .Col = 7: .Value = CStr(IIf(dValue > 0, dValue, ""))
'                    irow2 = irow2 + 2
'                    Exit For
'                End If
'            Next iCurRow
'        Wend
'    End With
'
'End Sub
'
'
'Private Function Sp_SumAbove(ByVal SS As Variant, ByVal iCol As Long, ByVal irow1, ByVal irow2) As Double
'    Dim dSum As Double
'    Dim iCount As Integer
'
'    dSum = 0
'
'    With SS
'        If irow1 > irow2 Then
'            Sp_SumAbove = 0
'            Exit Function
'        End If
'        If irow2 > .MaxRows Then irow2 = .MaxRows
'        If irow2 < 2 Then
'            Sp_SumAbove = 0
'            Exit Function
'        End If
'        .Col = iCol
'        For iCount = irow1 To irow2
'            .Row = iCount
'            If .CellType = SS_CELL_TYPE_NUMBER And .Text <> "" Then
'                dSum = dSum + .Value
'            End If
'        Next iCount
'
'    End With
'    Sp_SumAbove = dSum
'End Function
'
'Private Function Sp_SumGroup(ByVal SS As Variant, ByVal iRow As Long, ByVal iCol As Long) As Double
'    Dim dSum As Double
'
'    dSum = 0
'
'    With SS
'        If .MaxRows < 2 Then
'            Sp_SumGroup = 0
'            Exit Function
'        End If
'        .Col = iCol
'        .Row = iRow
'        If .CellType = SS_CELL_TYPE_NUMBER And .Text <> "" Then
'            dSum = dSum + .Value
'        End If
'
'    End With
'    Sp_SumGroup = dSum
'End Function
'
'Private Sub Sp_AutoInsertSumGroup()
'    Dim dValue As Double
'    Dim dValue2 As Double
'    Dim dValue3 As Double
'    Dim dValue4 As Double
'    Dim iCount As Integer
'    Dim iRow As Integer
'    Dim bLoop As Boolean
'    Dim curRow As Integer
'    Dim strProdTag As String
'    Dim strTag As String
'    Dim strTag2 As String
'    Dim strTag3 As String
'    Dim strTag4 As String
'    Dim strTag21 As String
'    Dim strTag31 As String
'    Dim strTag41 As String
'    iRow = 1
'    bLoop = True
'    dValue2 = 0
'    dValue3 = 0
'    dValue4 = 0
'
'    With ss1
'
'        If .MaxRows < 2 Then Exit Sub
'
'        While bLoop
'            If iRow >= .MaxRows Then
'                bLoop = False
'                Exit Sub
'            End If
'
'            .Col = 0: .Row = iRow
'            If .Text = "∑" Then
'                If (iRow + 1) < .MaxRows Then
'                    iRow = iRow + 1
'                Else
'                    Exit Sub
'                End If
'            End If
'
'            For iCount = 7 To 7
'                .Row = iRow
'                .Col = 2: strTag2 = .Text
'                .Col = 3: strTag3 = .Text
'                .Col = 4: strTag4 = .Text
'
'                dValue = 0
'
'                For curRow = iRow To .MaxRows
'                    .Row = curRow:
'
'                    .Col = 2: strTag21 = .Text
'                    .Col = 3: strTag31 = .Text
'                    .Col = 4: strTag41 = .Text
'
'                    .Col = 2
'                    If .Text <> "" And strTag2 = strTag21 Then
'
'                        If strTag3 = strTag31 Then
'
'                            If strTag4 = strTag41 Then
'                                 dValue = dValue + Sp_SumGroup(ss1, curRow, iCount)
'                            Else
'                                 dValue2 = dValue2 + dValue
'                                 dValue3 = dValue3 + dValue
'
'                                 .MaxRows = .MaxRows + 1
'                                 .Row = curRow
'                                 .Action = SS_ACTION_INSERT_ROW
'                                 iRow = .Row + 1
'                                 .Col = 0: .Text = "∑"
'                                 .Col = 4: .Text = strTag4 + " 小计"
'                                 strTag4 = strTag41
'
'                                 Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'                                 .Col = iCount: .Value = IIf(dValue > 0, dValue, "")
'
'                                 dValue = 0
'                                 Exit For
'                             End If
'
'                        Else
'
'                             dValue2 = dValue2 + dValue
'                             dValue3 = dValue3 + dValue
'
'                             .MaxRows = .MaxRows + 1
'                             .Row = curRow
'                             .Action = SS_ACTION_INSERT_ROW
'                             iRow = .Row + 1
'                             .Col = 0: .Text = "∑"
'                             .Col = 4: .Text = strTag4 + " 小计"
'                             strTag4 = strTag41
'
'                             Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'                             .Col = iCount: .Value = IIf(dValue > 0, dValue, "")
'
'                             .MaxRows = .MaxRows + 1
'                             .Row = iRow
'                             .Action = SS_ACTION_INSERT_ROW
'                             iRow = .Row + 1
'                             .Col = 0: .Text = "∑"
'                             .Col = 3: .Text = strTag3 + " 小计"
'                             strTag3 = strTag31
'
'                             Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'                             .Col = iCount: .Value = IIf(dValue3 > 0, dValue3, "")
'
'                             dValue3 = 0
'                             dValue = 0
'
'                             Exit For
'
'                        End If
'                    Else
'
'                         dValue2 = dValue2 + dValue
'                         dValue3 = dValue3 + dValue
'
'                         .MaxRows = .MaxRows + 1
'                         .Row = curRow
'                         .Action = SS_ACTION_INSERT_ROW
'                         iRow = .Row + 1
'                         .Col = 0: .Text = "∑"
'                         .Col = 4: .Text = strTag4 + " 小计"
'
'                         Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'                         .Col = iCount: .Value = IIf(dValue > 0, dValue, "")
'                         strTag4 = strTag41
'
'                         .MaxRows = .MaxRows + 1
'                         .Row = iRow
'                         .Action = SS_ACTION_INSERT_ROW
'                         iRow = .Row + 1
'                         .Col = 0: .Text = "∑"
'                         .Col = 3: .Text = strTag3 + " 小计"
'                         strTag3 = strTag31
'
'                         Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'
'                         .Col = iCount: .Value = IIf(dValue3 > 0, dValue3, "")
'
'                         .MaxRows = .MaxRows + 1
'                         .Row = iRow
'                         .Action = SS_ACTION_INSERT_ROW
'                         iRow = .Row + 1
'                         .Col = 0: .Text = "∑"
'                         .Col = 2: .Text = strTag2 + " 小计"
'                         strTag2 = strTag21
'
'                         Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'
'                         .Col = iCount: .Value = IIf(dValue2 > 0, dValue2, "")
'
'                         dValue2 = 0
'                         dValue3 = 0
'                         dValue = 0
'
'                         Exit For
'
'                    End If
'
'                Next curRow
'
'            Next iCount
'
'            'iRow = curRow
'            'dValue2 = dValue2 + dValue
'       Wend
'    End With
'
'End Sub

