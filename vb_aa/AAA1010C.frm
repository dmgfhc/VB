VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1010C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "设备检修计划录入_AAA1010C"
   ClientHeight    =   8145
   ClientLeft      =   315
   ClientTop       =   2130
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   14100
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txt_today 
      Height          =   300
      Left            =   8880
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_plt_NAME 
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
      Left            =   6480
      MaxLength       =   80
      TabIndex        =   2
      Top             =   90
      Width           =   2265
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   5775
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   90
      Width           =   645
   End
   Begin InDate.UDate dtp_faci_manage_str 
      Height          =   315
      Left            =   1905
      TabIndex        =   0
      Tag             =   "FACI_MANAGE_STR"
      Top             =   90
      Width           =   1410
      _ExtentX        =   2487
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
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   3975
      Top             =   90
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   90
      Top             =   90
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      Caption         =   "检修开始日期"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8640
      Left            =   90
      TabIndex        =   3
      Top             =   495
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   15240
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
      MaxCols         =   0
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AAA1010C.frx":0000
   End
End
Attribute VB_Name = "AAA1010C"
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
'-- Program ID        AAA1010C
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
                Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dtp_faci_manage_str, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Sc1.Add Item:="AAA1010C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    iSumCnt = 2
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
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
    
'    txt_plt.Text = "B1"
'    Call txt_plt_KeyUp(0, 0)
    
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
    
    ss1.MaxCols = 0
    ss1.MaxRows = 0
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    rControl(1).SetFocus
'    txt_plt.Text = "B1"
'    Call txt_plt_KeyUp(0, 0)

End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
   
        If Sp_Header_Refer() Then
            If Sp_Data_Refer() Then
  '              Txt_today.Text = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYY-MM-dd') FROM DUAL")
                Call SubSpreadSum
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Menu_Setting
   '             Call Gp_Ms_ControlLock(Mc1!lControl, True)
            End If
        End If
        Call Gp_Ms_ControlLock(Mc1!lControl, True)
    Else
        sMesg = sMesg + " 必须输入 ..."
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
End Sub

Public Sub Form_Pro()

    If Sp_Process(M_CN1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    Else
     '     Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    End If
    
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
    
 '   Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
   '     Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        'Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

    Dim iCol As Integer
    Dim iRow As Integer
    Dim sCode As String
    Dim dTime As Double
    Dim dCode As Double
    Dim sQuery As String
    
    With ss1
    
        If .ActiveRow = .MaxRows Then Exit Sub
        If .CellTag = "False" Then Exit Sub
            
        .Row = Row
                  
        Select Case Col Mod 2
        
            Case 0      'time
            
                .Col = Col - 1
                sCode = Trim(.Text)
                .Col = Col
                If .Value = "" Then
                    dTime = 0
                Else
                    dTime = .Value
                End If
          
                If dTime = 0 Then Exit Sub
                
                If sCode = "" Then
                  
                    .Col = Col
                    .Row = Row
                    .CellTag = "False"
                    
                    Call Gp_MsgBoxDisplay("必须先输入代码...")
                    
                    .Col = Col
                    .Row = Row
                    .CellTag = ""
                    
                    .Text = ""
                    .TabStop = True
                    .SetFocus
                    .SetActiveCell Col, Row
                    .Action = SS_ACTION_ACTIVE_CELL
                    .EditMode = True
                    .TabStop = False
        
                End If
           
            Case 1      'code
                
                .Col = Col
                sCode = Trim(.Text)
                If sCode = "" Then Exit Sub
                
                sQuery = "SELECT cd  FROM ZP_cd WHERE cd_mana_no = 'A0005' and cd='" + sCode + "'"
                
                If Gf_CodeFind(M_CN1, sQuery) = "" Then
                     
                    .Col = Col
                    .Row = Row
                    .CellTag = "False"
                    
                    Call Gp_MsgBoxDisplay("必须输入正确代码...")
                    
                    .Col = Col
                    .Row = Row
                    .CellTag = ""
                    
                    .Text = ""
                    .TabStop = True
                    .SetFocus
                    .SetActiveCell Col, Row
                    .Action = SS_ACTION_ACTIVE_CELL
                    .EditMode = True
                    .TabStop = False
        
                End If
            
        End Select
            
   End With

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

        .ColHeaderRows = 3
        .RowHeaderCols = 2
        
        .Col = -1
        .Row = SpreadHeader + 1
        .FontBold = True
        
        .RowHeight(SpreadHeader) = 15
        .RowHeight(SpreadHeader + 1) = 15
        
        .Row = SpreadHeader + 2
        .RowHidden = True
        
        .ColWidth(0) = 6
        
        .Col = 0
        .ColHidden = True
        
        .ColWidth(SpreadHeader + 1) = 10
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        .Col = SpreadHeader + 1: .Col2 = -1
        .Row = 0: .Row2 = SpreadHeader + 1
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        .Row = SpreadHeader
        .Col = SpreadHeader + 1
        .Text = "日期\工序"
        .Row = SpreadHeader + 1
        .Col = SpreadHeader + 1
        .Text = "日期\工序"
        
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

Public Function Sp_Header_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iDays As Integer
    Dim sQuery As String
    Dim sEdate, sEdate1, sEdate2 As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    
    sEdate = dtp_faci_manage_str.Text
    sEdate2 = Mid(Format(DateAdd("M", 1, sEdate), "yyyy-mm-dd"), 1, 8) & "01"
    iDays = DateDiff("D", sEdate, sEdate2)
   
    sQuery = "SELECT  SUBSTR(CD,1,4) AS PRC,SUBSTR(CD,5,1) AS LINE, CD_SHORT_NAME "
    sQuery = sQuery + "   FROM  ZP_CD "
  '  sQuery = sQuery + "  WHERE SUBSTR(CD,1,2)= '" + Trim(txt_plt.Text) + "' "
  '  sQuery = sQuery + "    AND CD_MANA_NO = 'A0002' "
  
    If txt_plt.Text = "" Then
       sQuery = sQuery + "  WHERE  CD_MANA_NO = 'A0002' "
    Else
       sQuery = sQuery + "  WHERE SUBSTR(CD,1,2)= '" + Trim(txt_plt.Text) + "' "
      sQuery = sQuery + "    AND CD_MANA_NO = 'A0002' "
    
    End If
  
    sQuery = sQuery + "  ORDER BY PRC||LINE "

    With ss1

        Sp_Header_Refer = True
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        adoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If adoRs.BOF Or adoRs.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            adoRs.Close
            Set adoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = adoRs.GetRows
        adoRs.Close
        Set adoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 2
        
            For iCol = 0 To .MaxCols - 1 Step 2
            
                .Col = iCol + 1
                .Row = SpreadHeader
                
                If VarType(ArrayRecords(2, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(2, iCnt))
                End If
                  
                .Col = iCol + 2
                .Row = SpreadHeader
                If VarType(ArrayRecords(2, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(2, iCnt))
                End If
                
                .Col = iCol + 1:  .Row = SpreadHeader + 1:  .Text = "代码"
                .Col = iCol + 2:  .Row = SpreadHeader + 1:  .Text = "时间"
                
                .Col = iCol + 1
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                .Col = iCol + 2
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(1, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt))
                End If
                
                'Column Type Setting
                .Col = iCol + 1: .Col2 = iCol + 1
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 1       'SS_CELL_TYPE_EDIT
                .TypeMaxEditLen = 6
                .TypeHAlign = TypeHAlignLeft
                .BlockMode = False
                
                .ColWidth(iCol + 1) = 7
                
                .Col = iCol + 2: .Col2 = iCol + 2
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 24
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
            
                .BlockMode = False
                
                .ColWidth(iCol + 2) = 5
                iCnt = iCnt + 1
                
            Next iCol
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    With ss1
        .Col = SpreadHeader + 1
        For iCnt = 1 To iDays
           .MaxRows = .MaxRows + 1
           .Row = .MaxRows
           .Text = Format(DateAdd("D", iCnt - 1, sEdate), "yyyy-mm-dd")
        Next iCnt
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Text = "合  计"
        
        Call Gp_Sp_BlockLock(ss1, 1, -1, ss1.MaxRows, ss1.MaxRows, True)
        Call Gp_Sp_RowColor(ss1, ss1.MaxRows, vbRed)
        Call Gp_Sp_EvenRowBackcolor(ss1)
        
        For iCol = 0 To .MaxCols - 2 Step 2
            .Col = iCol + 2: .Col2 = iCol + 2
            .Row = .MaxRows: .Row2 = .MaxRows
            .BlockMode = True
            .CellType = 13      'SS_CELL_TYPE_NUMBER
            .TypeNumberDecPlaces = 0
            .TypeNumberMax = 999
            .TypeNumberMin = 0
            .TypeNumberShowSep = True
            .TypeNumberLeadingZero = TypeLeadingZeroNo
            .TypeHAlign = TypeHAlignRight
            .BlockMode = False
        Next iCol
    
    End With
    
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
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
    Dim sTdate As String
    Dim sQuery As String
    Dim sEdate, sEdate1, sEdate2 As String
    Dim sTplt_prc As String
    Dim sTprc_line As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    
    '3 Month After
    sEdate = Mid(dtp_faci_manage_str.Text, 1, 4) + _
             Mid(dtp_faci_manage_str.Text, 6, 2) + _
             Mid(dtp_faci_manage_str.Text, 9, 2)
    
    sEdate1 = Mid(Format(DateAdd("M", 1, dtp_faci_manage_str.Text), "yyyy-mm-dd"), 1, 8) & "01"
    sEdate2 = Mid(sEdate1, 1, 4) + _
              Mid(sEdate1, 6, 2) + _
              Mid(sEdate1, 9, 2)

    sQuery = "SELECT TO_DATE(FACI_MANAGE_STR,'YYYY-MM-DD'), PLT || PRC, PRC_LINE, FACI_MANAGE_CD, FACI_MANAGE_TME"
    sQuery = sQuery + "   FROM AP_FACI_PLAN "
'    sQuery = sQuery + "  WHERE FACI_MANAGE_STR BETWEEN '" + sEdate + "' AND '" + sEdate2 + "' "
    sQuery = sQuery + "  WHERE FACI_MANAGE_STR >='" + sEdate + "' AND FACI_MANAGE_STR <'" + sEdate2 + "' "
'    sQuery = sQuery + "    AND PLT  = '" + Trim(txt_plt.Text) + "' "

    If txt_plt.Text = "" Then
    
    Else
        sQuery = sQuery + "    AND PLT  = '" + Trim(txt_plt.Text) + "' "
    End If
    
    sQuery = sQuery + "  ORDER BY FACI_MANAGE_STR, PLT, PRC, PRC_LINE "
    
 '   Debug.Print sQuery
    With ss1

        Sp_Data_Refer = True
        .ReDraw = False
       ' .MaxRows = 0
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
            iRow = 1
            For iCnt = 0 To UBound(ArrayRecords, 2)
                .Row = iRow
                .Col = SpreadHeader + 1
                sTdate = .Text
                sEdate1 = Trim(ArrayRecords(0, iCnt))
                sEdate1 = Format(sEdate1, "yyyy-mm-dd")
                Do While iRow <= .MaxRows And sTdate <> Format(ArrayRecords(0, iCnt), "yyyy-mm-dd")
                   iRow = iRow + 1
                   .Row = iRow
                   sTdate = .Text
                Loop
                    
                For iCol = 1 To .MaxCols Step 2
                    .Row = SpreadHeader + 2
                    .Col = iCol:     sTplt_prc = .Text
                    .Col = iCol + 1: sTprc_line = .Text

                    If sTplt_prc = ArrayRecords(1, iCnt) And sTprc_line = ArrayRecords(2, iCnt) Then

                        .Row = iRow
                        .Col = iCol     'Code
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(3, iCnt))
                        End If
                        
                        .Col = iCol + 1 'Time
                        If VarType(ArrayRecords(4, iCnt)) = vbNull Or Trim(ArrayRecords(4, iCnt)) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(4, iCnt))
                        End If
                
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    Txt_today.Text = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYY-MM-dd') FROM DUAL")
    
    With ss1
        For iRow = 1 To .MaxRows
            .Row = iRow
            If Gf_Get_Cell_Text(iRow, SpreadHeader + 1) < Trim(Txt_today.Text) Or Trim(txt_plt.Text) = "" Then
               Call Gp_Sp_BlockLock(ss1, 1, -1, iRow, iRow, True)
            End If
            
            For iCol = 2 To .MaxCols Step 2
   '             .Row = iRow
                .Col = iCol
                If .Text = "" Then
                   .Col = iCol - 1
                   .Text = ""
            
                End If
            Next iCol
        Next iRow
     End With
     
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

    If txt_plt.Text = "" Then
       Sp_Process = False
       Call Form_Ref
       Exit Function
    End If
    
    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    
    Dim sMesg As String
    Dim sTemp As String
    
    Dim adoCmd As ADODB.Command

    Sp_Process = True
     
    With ss1
        For iRow = 1 To .MaxRows
            For iCol = 2 To .MaxCols Step 2
                .Row = iRow
                .Col = iCol
                If .Text = "" Then
                   .Col = iCol - 1
                   .Text = ""
                End If
            Next iCol
        Next iRow
    End With
    
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
        For iCount = 1 To 7
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        For iRow = 1 To .MaxRows
            
            .Row = iRow
            
            'Parameters Setting
            For iCol = 1 To .MaxCols Step 2
            
                .Col = iCol
                If Trim(.Text) <> "" Then
                
                    .Row = SpreadHeader + 2
                    .Col = iCol
                    adoCmd.Parameters(1).Value = Mid(Trim(.Text), 1, 2)     'plt
                    adoCmd.Parameters(2).Value = Mid(Trim(.Text), 3, 2)     'prc
                    
                    .Col = iCol + 1
                    adoCmd.Parameters(3).Value = Trim(.Text)                'prc_line
                    
                    .Row = iRow
                    .Col = SpreadHeader + 1                                 'faci_manage_str
                    adoCmd.Parameters(0).Value = Mid(Trim(.Text), 1, 4) + _
                                                 Mid(Trim(.Text), 6, 2) + _
                                                 Mid(Trim(.Text), 9, 2)
                    
                    .Col = iCol
                    adoCmd.Parameters(4).Value = Trim(.Text)                'faci_manage_cd
                    
                    .Col = iCol + 1                                         'faci_manage_tme
                    If Trim(.Text) = "" Then
                        adoCmd.Parameters(5).Value = 0
                    Else
                        dTempInt = .Text
                        adoCmd.Parameters(5).Value = dTempInt
                    End If
                    
                    adoCmd.Parameters(6).Value = sUserID                    'User-id
                                   
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
            
            Next iCol
            
        Next iRow
        
        Conn.CommitTrans
        MDIMain.StatusBar1.Panels(1) = "提示信息: 数据处理完成"
        
        For iCount = 1 To .MaxRows
            .Row = iCount
            .Col = SpreadHeader
            .Text = ""
        Next iCount
    
        .ReDraw = True
        Screen.MousePointer = vbDefault
        Call Form_Ref
    
    End With
    
    Exit Function

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("SpreadPro_Error : " & Error)

End Function

Private Sub txt_plt_Change()

    If Trim(txt_plt.Text) = "" Then txt_plt_NAME.Text = ""
    
End Sub

Private Sub txt_plt_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=txt_plt_NAME

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_plt)) = txt_plt.MaxLength Then
        txt_plt_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        txt_plt_NAME.Text = ""
    End If

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim iCol As Integer
    
    If Gf_Get_Cell_Text(ss1.ActiveRow, SpreadHeader + 1) < Trim(Txt_today.Text) Then Exit Sub
    
    iCol = ss1.ActiveCol
    
    If KeyCode = vbKeyF4 Then
    
       Select Case iCol Mod 2
       
           Case 1
                Set DD.sPname = Me.ss1
                DD.sWitch = "SP"
                DD.sKey = "A0005"
                DD.rControl.Add Item:=iCol
                
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
        End Select
        
    End If

End Sub

Private Sub SubSpreadSum()

    Dim i As Long
    Dim j As Long
    
    Dim iTot() As Long
    Dim iCnt  As Integer
    
    With ss1
        
        iCnt = .MaxCols / 2
        
        ReDim iTot(iCnt)
                
        For i = 1 To .MaxRows
             iCnt = 0
            For j = 2 To .MaxCols Step 2
                 iCnt = iCnt + 1
                 iTot(iCnt) = iTot(iCnt) + GF_GET_CELL_VALUE(ss1, i, j)
            Next j
        
        Next i
                
        For i = 1 To .MaxRows
            iCnt = 0
            For j = 2 To .MaxCols Step 2
                iCnt = iCnt + 1
                Call GP_SET_CELL_VALUE(ss1, .MaxRows, j, iTot(iCnt))
            Next j
        Next i
    
    End With
    
End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : GP_SET_CELL_VALUE
'   2.Name         : Set Spread Text
'   3.Input  Value : Spread Name , Row , Col
'   4.Return Value : None
'   5.Writer       :
'   6.Create Date  : 2003. 09 .11
'   7.Modify Date  :
'   8.Comment      : Set Spread Text
'---------------------------------------------------------------------------------------
Public Sub GP_SET_CELL_VALUE(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long, sText As Variant)
    
    If iRow <= 0 Then Exit Sub
    
    With ss1
        .Row = iRow
        .Col = iCol
        '.Text = Val(sText)
        .Value = Val(sText)
    End With
    
End Sub

Public Function GF_GET_CELL_VALUE(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant

    With ss1
        .Row = iRow
        .Col = iCol
        GF_GET_CELL_VALUE = Val(.Value)
    End With
    
End Function

Private Function Gf_Get_Cell_Text(ByVal iRow As Long, ByVal iCol As Long) As Variant
    
    With ss1
        .Row = iRow
        .Col = iCol
        Gf_Get_Cell_Text = .Text
    End With
    
End Function
