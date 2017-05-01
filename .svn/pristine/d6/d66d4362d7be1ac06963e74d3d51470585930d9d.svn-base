VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1050C 
   Caption         =   "原料计划使用量_AAA1050C"
   ClientHeight    =   6570
   ClientLeft      =   465
   ClientTop       =   2520
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   13755
   WindowState     =   2  'Maximized
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
      Left            =   6885
      TabIndex        =   5
      Tag             =   "炼钢工序"
      Top             =   90
      Width           =   1050
   End
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
      Left            =   4275
      TabIndex        =   4
      Tag             =   "工厂"
      Top             =   90
      Width           =   870
   End
   Begin VB.TextBox txt_stlgrd 
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
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   1
      Tag             =   "钢种"
      Top             =   450
      Width           =   1320
   End
   Begin VB.TextBox txt_stlgrd_des 
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
      Left            =   2790
      TabIndex        =   2
      Top             =   450
      Width           =   5820
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   6540
      Left            =   90
      TabIndex        =   3
      Top             =   2565
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   11536
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
      MaxCols         =   1
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AAA1050C.frx":0000
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Left            =   5535
      Top             =   90
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "炼钢工序"
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
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Tag             =   "日期"
      Top             =   90
      Width           =   1140
      _ExtentX        =   2011
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
   Begin InDate.ULabel ULabel2 
      Height          =   300
      Left            =   90
      Top             =   90
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
   Begin InDate.ULabel ULabel6 
      Height          =   285
      Left            =   90
      Top             =   450
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      Caption         =   "钢种"
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
   Begin InDate.ULabel ULabel3 
      Height          =   300
      Left            =   2970
      Top             =   90
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
   Begin FPSpread.vaSpread ss2 
      Height          =   1725
      Left            =   90
      TabIndex        =   6
      Top             =   810
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   3043
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
      MaxCols         =   12
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      SpreadDesigner  =   "AAA1050C.frx":02C8
   End
End
Attribute VB_Name = "AAA1050C"
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
'-- Program ID        AAA1050C
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
Dim Sc2 As New Collection           'Spread Collection
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
         Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_stlgrd_des, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         
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
    Sc1.Add Item:="AAA1050C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
      
    Sc2.Add Item:=ss2, Key:="Spread"
    
    sQuery = "SELECT DISTINCT SUBSTR(CD,1,2) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND SUBSTR(CD,1,1) = 'B' "
    Call Gf_ComboAdd(M_CN1, cbo_plt, sQuery)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cbo_plt_Change()
 
   Dim sQuery As String
   
   sQuery = "SELECT DISTINCT SUBSTR(CD,3,2) FROM ZP_CD WHERE SUBSTR(CD,3,2) in ('BC','BD','BE') AND CD_MANA_NO = 'A0002' AND SUBSTR(CD,1,2) = '" + cbo_plt.Text + "' "
   Call Gf_ComboAdd(M_CN1, cbo_prc, sQuery)
 
End Sub

Private Sub cbo_plt_Click()
   
   Dim sQuery As String
   
   sQuery = "SELECT DISTINCT SUBSTR(CD,3,2) FROM ZP_CD WHERE SUBSTR(CD,3,2) IN ('BC','BD','BE') AND CD_MANA_NO = 'A0002' and SUBSTR(CD,1,2) = '" + cbo_plt.Text + "' "
   Call Gf_ComboAdd(M_CN1, cbo_prc, sQuery)

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
    Call Gp_Sp_Setting(Sc2.Item("Spread"))
    Call Sp_Setting
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    'Call Sp_Header_Set
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
    Set Sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Cls()
    
    ss1.MaxRows = 0
    ss1.MaxCols = 1
'    Sp_Header_Set
    ss2.MaxRows = 0
    ss2.MaxCols = 0
'    If Sp_Header_Refer() Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Menu_Setting
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
 '   End If
    
End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
                
        sMesg = Gf_Ms_NeceCheck2(mControl)
        If sMesg = "OK" Then
        
            Call Sp2_Header_Set
            Call Sp2_Data_Refer
            
            Call Sp_Header_Set
            If Sp_Header_Refer() Then
                If Sp_Data_Refer() Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                    Call Menu_Setting
                    Call Gp_Ms_ControlLock(Mc1!lControl, True)
                End If
            End If

            If Left(dtp_yy_mm.RawData, 6) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Or Trim(txt_stlgrd.Text) = "***********" Or Trim(txt_stlgrd.Text) = "" Then
               Call Gp_Sp_BlockLock(ss1, 1, 1, 1, ss1.MaxRows, True)
            Else
               Call Gp_Sp_BlockLock(ss1, 1, 1, 1, ss1.MaxRows, False)
               Call Gp_Sp_ColColor(ss1, 1, &H80000008, &HC0FFFF)
            End If
        
            
        Else
            sMesg = sMesg + " 必须按项目长度输入"
            Call Gp_MsgBoxDisplay(sMesg)
        End If
    Else
        sMesg = sMesg + " 必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
        
    End If
    
End Sub

Public Sub Form_Pro()

    If Trim(txt_stlgrd.Text) = "***********" Then Exit Sub
    
    If Left(dtp_yy_mm.RawData, 6) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Then
        Call Gp_MsgBoxDisplay("Can't Process...")
        Exit Sub
    End If
    
    If Sp_Process(M_CN1, Proc_Sc("Sc")) Then
'        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'        Call Menu_Setting
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
    
 '       Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
   
'   ratio check

    Dim iCol As Integer
    Dim iRow As Integer
    Dim DCURR As Double
    Dim dRatio As Double
    
    If Trim(txt_stlgrd.Text) = "***********" Then Exit Sub
    
    With ss1
    
        If .CellTag = "False" Then Exit Sub
        
        dRatio = 0
        .Col = 1
        For iRow = 1 To .MaxRows
            .Row = iRow
            
            If .Value = "" Then
                DCURR = 0
            Else
                DCURR = .Value
            End If
            
            dRatio = dRatio + DCURR
            If dRatio >= 1.01 Then
                .Col = Col
                .Row = Row
                .CellTag = "False"
                Call Gp_MsgBoxDisplay("Ratio 之和不能大于1(100%)...")
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
                
                Exit Sub
            End If
            
        Next iRow
            
    End With
   
    If Col = 1 And Row > 0 Then
   
        With ss1
                
             .Col = Col
             .Row = Row
             If .Value = "" Then
                 dRatio = 0
             Else
                 dRatio = .Value
             End If
             
             For iCol = 2 To 24 Step 2
                 .Col = iCol
                 If .Value = "" Then
                    DCURR = 0
                 Else
                    DCURR = .Value
                 End If
                 .Col = iCol + 1
                 If dRatio * DCURR = 0 Then
                    .Text = ""
                 Else
                    .Text = dRatio * DCURR / 100
                 End If
             Next iCol
        End With
    End If
    
   If Col = ss1.MaxCols And Row = ss1.MaxRows Then
      ss1.Col = 1
      ss1.Row = 1
   End If
   
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

      '  .ColHeaderRows = 2
        .RowHeaderCols = 2
        
        .RowHeight(SpreadHeader) = 16
        
     '   .Row = SpreadHeader + 1
     '   .RowHidden = True
        .Col = SpreadHeader + 1
    '    .ColHidden = True

        .ColWidth(0) = 12
      '  .ColWidth(1) = 20
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        
    
'        .BlockMode = True
'        .RowMerge = MergeAlways
'        .ColMerge = MergeAlways
'        .BlockMode = False
        
    End With
    
    ss2.Col = SpreadHeader + 1
    ss2.ColHidden = True

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
    Dim iMonth As Integer
    Dim sMonth As String
    
    With ss1
        .MaxCols = 1
        .MaxRows = 0
        .MaxCols = 25
        .Col = SpreadHeader
        .Row = SpreadHeader
        .Text = "原料单耗项目"
        .ColWidth(SpreadHeader + 1) = 6
        
        .Col = SpreadHeader + 1
   '     .ColHidden = True
        
        .Col = 1
        .Text = "单耗(%)"
        .RowHeight(SpreadHeader) = 24
        .Row = SpreadHeader
        iMonth = 0
        For iCol = 2 To 24 Step 2
            sMonth = Month(DateAdd("M", iMonth, dtp_yy_mm.Text)) & "月"
            .Col = iCol
            .Text = sMonth
            
            '----
            .ColHidden = True
            
            .Col = iCol + 1
            .Text = sMonth
            iMonth = iMonth + 1
                  
            'Column Type Setting
            .Col = iCol: .Col2 = iCol
            .Row = 1: .Row2 = -1
            .BlockMode = True
            .CellType = 13      'SS_CELL_TYPE_NUMBER
            .TypeNumberDecPlaces = 3
            .TypeNumberMax = 999999999
            .TypeNumberMin = 0
            .TypeNumberShowSep = True
            .TypeNumberLeadingZero = TypeLeadingZeroNo
            .TypeHAlign = TypeHAlignRight
            .BlockMode = False

            .ColWidth(iCol + 1) = 8

            .Col = iCol + 1: .Col2 = iCol + 1
            .Row = 1: .Row2 = -1
            .BlockMode = True
            .CellType = 13      'SS_CELL_TYPE_NUMBER
            .TypeNumberDecPlaces = 3
            .TypeNumberMax = 999999999
            .TypeNumberMin = 0
            .TypeNumberShowSep = True
            .TypeNumberLeadingZero = TypeLeadingZeroNo
            .TypeHAlign = TypeHAlignRight
            .BlockMode = False

            .ColWidth(iCol + 2) = 8
             
        Next iCol
        .ReDraw = True
        .Refresh
    
    End With
    
End Sub

Public Function Sp_Header_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim DCURR As Double
    Dim sQuery As String
    Dim sEdate As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    Dim sQuery2 As String
    
    Dim AdoRs2 As ADODB.Recordset
    Dim ArrayRecords2 As Variant

    Set AdoRs2 = New ADODB.Recordset
    
    sQuery2 = "SELECT cd_short_name,cd  "
    sQuery2 = sQuery2 + "   FROM zp_cd "
    sQuery2 = sQuery2 + "  WHERE CD_mana_no in ( 'F0001','F0017','F0018') "
 
    With ss1

        Sp_Header_Refer = True
        .MaxRows = 0
        .ColWidth(0) = 15
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs2.Open sQuery2, M_CN1, adOpenKeyset
        
        If AdoRs2.BOF Or AdoRs2.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            AdoRs2.Close
            Set AdoRs2 = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords2 = AdoRs2.GetRows
        AdoRs2.Close
        Set AdoRs2 = Nothing

        If UBound(ArrayRecords2, 2) + 1 <> 0 Then
        
            .MaxRows = (UBound(ArrayRecords2, 2) + 1)
            If Trim(txt_stlgrd.Text) <> "***********" Then
                .Col = 2
            Else
                .Col = 1
            End If
            .Row = 1
            .Col2 = .MaxCols
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = True
            .BlockMode = False
            .Protect = True
                  
            iCnt = 0
            
            For iRow = 1 To .MaxRows
            
                .Row = iRow
                .Col = SpreadHeader
                
                If VarType(ArrayRecords2(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords2(0, iCnt))
                End If
                
                .Col = SpreadHeader + 1
                .Text = Trim(ArrayRecords2(1, iCnt))
                
                .Row = iRow + 2: .Row2 = iRow + 2
                .Col = 1: .Col2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                iCnt = iCnt + 1
            Next iRow
                
        End If
        
        .ReDraw = True
        .Refresh
        
        Screen.MousePointer = vbDefault
        
    End With
    
    If ss2.MaxRows < 1 Then
       Exit Function
    End If
    
    If Trim(txt_stlgrd.Text) <> "***********" Then
        Select Case cbo_prc.Text
            Case "BE"
                  iRow = 1
            Case "BD"
                  iRow = 2
            Case "BC"
                  iRow = 3
        End Select
        
        For iCol = 1 To 12
           ss2.Row = iRow
           ss2.Col = iCol
           DCURR = IIf(ss2.Value = "", 0, ss2.Value)
           ss1.Col = iCol * 2
           If DCURR <> 0 Then
              For iCnt = 1 To ss1.MaxRows
                  ss1.Row = iCnt
                  ss1.Text = DCURR
              Next iCnt
           End If
        Next iCol
    End If
    
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Set AdoRs2 = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim iArr As Integer
    Dim dRatio, DCURR As Double

    Dim sTdate As String
    Dim sQuery As String
    Dim sEdate2 As String
    Dim sEdate As String
    Dim iEdate As Integer
    Dim sUnit_kind As String
    Dim sTHK_GRP As String
   ' Dim SPARA As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim ArrayMonth(11) As String

    Set adoRs = New ADODB.Recordset
    
    sEdate2 = dtp_yy_mm.Text & "-01"
    
    For iCnt = 1 To 12
        sEdate = Format(DateAdd("M", iCnt - 1, sEdate2), "yyyy-mm-dd")
        ArrayMonth(iCnt - 1) = Mid(sEdate, 1, 4) & Mid(sEdate, 6, 2)
    Next iCnt
    
    If Trim(txt_stlgrd.Text) = "" Then
       txt_stlgrd.Text = "***********"
    End If
    
    If Trim(txt_stlgrd.Text) = "***********" Then
        sQuery = "        SELECT A.MAT_KND, 0, SUM(A.UNIT_RATIO * B.M1/100),  SUM(A.UNIT_RATIO * B.M2/100),  SUM(A.UNIT_RATIO * B.M3/100), "
        sQuery = sQuery + "                    SUM(A.UNIT_RATIO * B.M4/100),  SUM(A.UNIT_RATIO * B.M5/100),  SUM(A.UNIT_RATIO * B.M6/100), "
        sQuery = sQuery + "                    SUM(A.UNIT_RATIO * B.M7/100),  SUM(A.UNIT_RATIO * B.M8/100),  SUM(A.UNIT_RATIO * B.M9/100), "
        sQuery = sQuery + "                    SUM(A.UNIT_RATIO * B.M10/100), SUM(A.UNIT_RATIO * B.M11/100), SUM(A.UNIT_RATIO * B.M12/100) "
        sQuery = sQuery + " FROM AP_UNIT A, "
        sQuery = sQuery + "      (SELECT PRC, STLGRD, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(0) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M1, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(1) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M2, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(2) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M3, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(3) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M4, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(4) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M5, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(5) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M6, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(6) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M7, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(7) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M8, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(8) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M9, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(9) + "'  THEN PLAN_VALUE ELSE 0 END ) AS M10, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(10) + "' THEN PLAN_VALUE ELSE 0 END ) AS M11, "
        sQuery = sQuery + "              SUM(CASE WHEN YEAR_MONTH='" + ArrayMonth(11) + "' THEN PLAN_VALUE ELSE 0 END ) AS M12 "
        sQuery = sQuery + "         FROM AP_PROD_PLAN "
        sQuery = sQuery + "        WHERE PLT       =  '" + Trim(cbo_plt.Text) + "'"
        sQuery = sQuery + "          AND PRC       =  '" + Trim(cbo_prc.Text) + "'"
        sQuery = sQuery + "          AND APLY_ITEM =  '006' "
        sQuery = sQuery + "        GROUP BY PRC, STLGRD) B "
        sQuery = sQuery + " WHERE A.PRC    = B.PRC "
        sQuery = sQuery + "   AND A.STLGRD = B.STLGRD "
        sQuery = sQuery + " GROUP BY MAT_KND "
    Else
        sQuery = "SELECT MAT_KND, UNIT_RATIO, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0"
        sQuery = sQuery + " FROM AP_UNIT "
        sQuery = sQuery + " WHERE PRC    = '" + Trim(cbo_prc.Text) + "'"
        sQuery = sQuery + "   AND STLGRD = '" + Trim(txt_stlgrd.Text) + "' "
    End If
    
    Debug.Print sQuery
    
    With ss1

        Sp_Data_Refer = True
        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        adoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If adoRs.BOF Or adoRs.EOF Then
        
            Sp_Data_Refer = False
            .ReDraw = True
            adoRs.Close
            Set adoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = adoRs.GetRows
        adoRs.Close
        Set adoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
         '   iRow = 1
            For iCnt = 0 To UBound(ArrayRecords, 2)
                iRow = 1
                .Row = iRow
                .Col = SpreadHeader + 1
                 sUnit_kind = .Text
                 Do While iRow <= .MaxRows And sUnit_kind <> Trim(ArrayRecords(0, iCnt))
                    iRow = iRow + 1
                    .Row = iRow
                    sUnit_kind = .Text
                 Loop
                
                 If iRow <= .MaxRows Then
                    .Col = 1
                     
                    If VarType(ArrayRecords(1, iCnt)) = vbNull Then
                        .Text = ""
                    Else
                         If Val((ArrayRecords(1, iCnt))) <> 0 Then
                        .Text = Trim(ArrayRecords(1, iCnt))
                        End If
                    End If
                    
                    iArr = 2
                    For iRow = 3 To 25 Step 2
                        .Col = iRow
                        If VarType(ArrayRecords(iArr, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            If Val((ArrayRecords(iArr, iCnt))) <> 0 Then
                                .Text = Trim(ArrayRecords(iArr, iCnt))
                            End If
                        End If
                        iArr = iArr + 1
                    Next iRow
                    
                 End If
                 
            Next iCnt
            
        End If
        
     '   .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    If Trim(txt_stlgrd.Text) <> "***********" Then
        With ss1
             For iRow = 1 To .MaxRows
                 .Col = 1
                 .Row = iRow
                 If .Value = "" Then
                     dRatio = 0
                 Else
                     dRatio = .Value
                 End If
                 For iCol = 2 To 24 Step 2
                     .Col = iCol
                     If .Value = "" Then
                        DCURR = 0
                     Else
                        DCURR = .Value
                     End If
                     .Col = iCol + 1
                     If dRatio * DCURR = 0 Then
                        .Text = ""
                     Else
                        .Text = dRatio * DCURR / 100
                     End If
                 Next iCol
             Next iRow
        
            .Col = 1
            .Row = 1
            .Col2 = 1
            .Row2 = ss1.MaxRows
            .BlockMode = True
            .BackColor = &HC0FFFF
            .BlockMode = False
            .Protect = True
        End With
    End If
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim sTdate As String
    Dim sEdate As String
    Dim sYear As String
    Dim iEdate As Integer
    
    sEdate = Mid(dtp_yy_mm.Text, 1, 4)
    sYear = sEdate
    sTdate = Mid(dtp_yy_mm.Text, 6, 2)
    iEdate = Val(sTdate)

    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim sMesg As String
    Dim sTemp As String
    Dim sPara As String
    
    Dim adoCmd As ADODB.Command
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
        Set adoCmd = New ADODB.Command
        
        Set adoCmd.ActiveConnection = Conn
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = Sc.Item("P-M")
        
        Conn.BeginTrans
        
        'Ceate Parameter (Input) iType + iColumn
        For iCount = 1 To 9
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        For iRow = 1 To .MaxRows
            
            .Row = iRow
            
            'Parameters Setting
            'For iCol = 1 To .MaxCols
            
            '    .Col = iCol
            '    If Trim(.Text) <> "" Then
                
            '        .Row = SpreadHeader + 1
            '        .Col = iCol
            '        adoCmd.Parameters(6).Value = .Text     'thk_grp
                   
              
            '        .Row = iRow
                    .Col = SpreadHeader + 1
                    adoCmd.Parameters(5).Value = .Text     'unit_knd
                    
                    .Col = 1
                 
                    If Trim(.Text) = "" Then                'unit_ratio
                        adoCmd.Parameters(6).Value = 0
                    Else
                        dTempInt = .Text
                        adoCmd.Parameters(6).Value = dTempInt
                    End If
                    
                    adoCmd.Parameters(8).Value = sUserID                     'User-id
           
                    adoCmd.Parameters(0).Value = "1"                         'iTable
           
                    adoCmd.Parameters(3).Value = cbo_prc.Text                'prc
                    adoCmd.Parameters(4).Value = txt_stlgrd.Text             'stlgrd
                    
                    adoCmd.Parameters(1).Value = "  "                        '
                    adoCmd.Parameters(2).Value = "  "
                    adoCmd.Parameters(7).Value = 0
                                   
                    adoCmd.Execute
                    sEdate = sYear
                    
                    For iCol = 3 To 25 Step 2
                    
                         .Col = SpreadHeader + 1
                         adoCmd.Parameters(5).Value = .Text                      'mat_knd
                         
                         .Col = iCol
                      
                         If Trim(.Text) = "" Then                                'pln_wgt
                             adoCmd.Parameters(7).Value = 0
                         Else
                             dTempInt = .Text
                             adoCmd.Parameters(7).Value = dTempInt
                         End If
                         
                         adoCmd.Parameters(8).Value = sUserID                    'User-id
                
                         adoCmd.Parameters(0).Value = "2"                         'iTable
                         If iEdate < 10 Then                                      'year_month
                            adoCmd.Parameters(1).Value = sEdate + "0" + LTrim(Str(iEdate))
                         Else
                            adoCmd.Parameters(1).Value = sEdate + LTrim(Str(iEdate))
                         End If
                         adoCmd.Parameters(2).Value = cbo_plt.Text                'plt
                         adoCmd.Parameters(3).Value = cbo_prc.Text                'prc
                         adoCmd.Parameters(4).Value = txt_stlgrd.Text             'stlgrd
                    
                         iEdate = iEdate + 1
                         If iEdate = 13 Then
                            iEdate = 1
                            sEdate = LTrim(Str(Val(sEdate) + 1))
                         End If
                         
                         adoCmd.Execute
                    
                    Next iCol
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
                
               ' End If
            
          '  Next iCol
            
        Next iRow
        
        Conn.CommitTrans
        .ReDraw = True
        MDIMain.StatusBar1.Panels(1) = "提示信息: 数据处理完成"
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

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_stlgrd_des

        DD.nameType = "2"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_stlgrd)) = txt_stlgrd.MaxLength Then
        txt_stlgrd_des.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
    Else
        txt_stlgrd_des.Text = ""
    End If

End Sub

Public Sub Sp2_Header_Set()

    Dim iCol As Integer
    Dim iRow As Integer
    
    Dim sMonth As String
    
    With ss2
    
        .ColWidth(0) = 15
        .Col = SpreadHeader
        .Row = SpreadHeader
        .Text = "生产计划"
        .Row = SpreadHeader
        .MaxCols = 12
        
        For iCol = 1 To 12
            .Row = SpreadHeader
            sMonth = Month(DateAdd("M", iCol - 1, dtp_yy_mm.Text)) & "月"
            .Col = iCol
            .Text = sMonth

            'Column Type Setting
            .Col = iCol: .Col2 = iCol
            .Row = 1: .Row2 = -1
            .BlockMode = True
            .CellType = 13      'SS_CELL_TYPE_NUMBER
            .TypeNumberDecPlaces = 3
            .TypeNumberMax = 99999999
            .TypeNumberMin = 0
            .TypeNumberShowSep = True
            .TypeNumberLeadingZero = TypeLeadingZeroNo
            .TypeHAlign = TypeHAlignRight
            .BlockMode = False

            .ColWidth(iCol) = 12

        Next iCol
        .ReDraw = True
        .Refresh
        
    End With
    
    ss2.MaxRows = 3
    ss2.Col = SpreadHeader
    ss2.Row = 1
    ss2.Col = SpreadHeader
    ss2.Row2 = 3
    ss2.Clip = "VD 处理量" + Chr(13) + "LF 处理量" + Chr(13) + "BOF 投入量"
   
    ss2.Col = SpreadHeader + 1
    ss2.Row = 1
    ss2.Col = SpreadHeader + 1
    ss2.Row2 = 3
    ss2.Clip = "BE " + Chr(13) + "BD " + Chr(13) + "BC "
               
    ss2.OperationMode = OperationModeRead

End Sub

Public Function Sp2_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sTdate As String
    Dim sYear As String
    Dim sQuery1 As String
    Dim sEdate, sEdate2 As String
    Dim iEdate As Integer
    Dim iMonth As Integer
    Dim AdoRs1 As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim ArrayMonth(11) As String
    
    sYear = Mid(dtp_yy_mm.Text, 1, 4)
    sTdate = Mid(dtp_yy_mm.Text, 6, 2)
    iEdate = Val(sTdate)
    iEdate = iEdate - 1
    
    sEdate2 = dtp_yy_mm.Text & "-01"
    
    For iCnt = 1 To 12
        sEdate = Format(DateAdd("M", iCnt - 1, Format(sEdate2, "yyyy-mm-dd")), "yyyy-mm-dd")
        ArrayMonth(iCnt - 1) = Mid(sEdate, 1, 4) & Mid(sEdate, 6, 2)
    Next iCnt
    
    'VD计划产量/LF计划产量/BOF计划产量
    sQuery1 = "SELECT     PRC, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(0) + "' then plan_value else 0 end ) as m1, "
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(1) + "' then plan_value else 0 end ) as m2,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(2) + "' then plan_value else 0 end ) as m3,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(3) + "' then plan_value else 0 end ) as m4,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(4) + "' then plan_value else 0 end ) as m5,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(5) + "' then plan_value else 0 end ) as m6,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(6) + "' then plan_value else 0 end ) as m7,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(7) + "' then plan_value else 0 end ) as m8,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(8) + "' then plan_value else 0 end ) as m9,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(9) + "' then plan_value else 0 end ) as m10,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(10) + "' then plan_value else 0 end ) as m11,"
    sQuery1 = sQuery1 + "sum(case when year_month='" + ArrayMonth(11) + "' then plan_value else 0 end ) as m12 "
    sQuery1 = sQuery1 + " FROM  AP_PROD_PLAN "
    sQuery1 = sQuery1 + " WHERE YEAR_MONTH BETWEEN '" + ArrayMonth(0) + "' AND '" + ArrayMonth(11) + "' "
    sQuery1 = sQuery1 + "   AND APLY_ITEM = '006' "
    sQuery1 = sQuery1 + "   AND PLT       = 'B1' "
    sQuery1 = sQuery1 + "   AND PRC       IN ('BC','BD','BE') "
    
    If Trim(txt_stlgrd.Text) <> "***********" Then
        sQuery1 = sQuery1 + "   AND STLGRD    LIKE '" + Trim(txt_stlgrd.Text) + "%' "
    End If
    
    sQuery1 = sQuery1 + " GROUP BY PRC "
    sQuery1 = sQuery1 + " ORDER BY PRC "
    
    Set AdoRs1 = New ADODB.Recordset
    
    With ss2

        Sp2_Data_Refer = True
        .ReDraw = False
        
        'Ado Execute
        AdoRs1.Open sQuery1, M_CN1, adOpenKeyset
        
        If AdoRs1.BOF Or AdoRs1.EOF Then
        
            AdoRs1.Close
            Set AdoRs1 = Nothing
            
        Else
        
            ArrayRecords = AdoRs1.GetRows
            AdoRs1.Close
            Set AdoRs1 = Nothing
            
            If UBound(ArrayRecords, 2) + 1 <> 0 Then
                
                For iCnt = 0 To UBound(ArrayRecords, 2)
                    Select Case ArrayRecords(0, iCnt)
                        Case "BE"
                            .Row = 1
                        Case "BD"
                            .Row = 2
                         Case "BC"
                            .Row = 3
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
        End If
        
        .ReDraw = True
        
    End With
    
    
    Screen.MousePointer = vbDefault
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Function

SpreadDisplay_Error:
    
    Set AdoRs1 = Nothing
    Sp2_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function
