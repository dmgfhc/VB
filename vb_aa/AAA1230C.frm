VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1230C 
   Caption         =   "编制板卷连铸计划_AAA1230C"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10875
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_excel 
      Height          =   315
      Left            =   3630
      TabIndex        =   0
      Text            =   "1"
      Top             =   180
      Visible         =   0   'False
      Width           =   675
   End
   Begin InDate.UDate dtp_yy_mm 
      Height          =   300
      Left            =   1455
      TabIndex        =   1
      Tag             =   "年月"
      Top             =   120
      Width           =   1185
      _ExtentX        =   2090
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
      Left            =   120
      Top             =   120
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "年月"
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
   Begin Threed.SSCommand plan_cmd 
      Height          =   330
      Left            =   13635
      TabIndex        =   2
      Top             =   120
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "编制转炉计划"
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   1320
      Left            =   120
      TabIndex        =   3
      Top             =   525
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   2328
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
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
      MaxCols         =   9
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      SpreadDesigner  =   "AAA1230C.frx":0000
   End
   Begin Threed.SSCommand ccm_plan_cmd 
      Height          =   330
      Left            =   11850
      TabIndex        =   4
      Top             =   120
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "编制连铸计划"
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   2355
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   4154
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
      MaxCols         =   10
      MaxRows         =   4
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AAA1230C.frx":0781
   End
   Begin FPSpread.vaSpread ss3 
      Height          =   4770
      Left            =   120
      TabIndex        =   6
      Top             =   4365
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   8414
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
      MaxCols         =   2
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AAA1230C.frx":0EDF
   End
End
Attribute VB_Name = "AAA1230C"
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
'-- Program ID        AAA1230C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder             GUOLI
'-- Date              2009.6.16
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
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
                     
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
   Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows, True)
   Call Gp_Sp_BlockLock(ss2, 1, ss2.MaxCols, 1, ss2.MaxRows, True)
   Call Gp_Sp_BlockLock(ss3, 1, ss3.MaxCols, 1, ss3.MaxRows, True)
   
   Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, 4, 4, BLACK, &HE6E6FF)
       
End Sub

Private Sub ccm_plan_cmd_Click()
On Error GoTo ccm_plan_cmd_Error

    Dim sQuery As String
    Dim iCount As Integer
    
    'If dtp_date_str.Enabled Then Exit Sub
    
    Dim adoCmd As ADODB.Command
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandType = adCmdText
    
    'Ceate Parameter (Output)
    For iCount = 1 To 7
        adoCmd.Parameters.Append adoCmd.CreateParameter(Str(iCount), adVariant, adParamOutput)
    Next iCount
    
    'CAST
    sQuery = "{call AAA4020P ('" + dtp_yy_mm.RawData + "','SL', 'B1','BF',?,?,?,?,?,?,? )}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
    
    If adoCmd(6) <> "" Then
        Call Gp_MsgBoxDisplay(adoCmd(6))
        Set adoCmd = Nothing
        Exit Sub
    Else
        MsgBox "连铸计划编制完了!"
        Set adoCmd = Nothing
        Call Form_Ref
        Exit Sub
    End If

ccm_plan_cmd_Error:

    Call Gp_MsgBoxDisplay("编制计划错误 : " & Error)
End Sub

Private Sub plan_cmd_Click()

On Error GoTo plan_cmd_Error

    Dim sQuery As String
    Dim iCount As Integer
    
    'If dtp_date_str.Enabled Then Exit Sub
    
    Dim adoCmd As ADODB.Command
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandType = adCmdText
    
    'Ceate Parameter (Output)
    For iCount = 1 To 7
        adoCmd.Parameters.Append adoCmd.CreateParameter(Str(iCount), adVariant, adParamOutput)
    Next iCount
    
    'CAST
    sQuery = "{call AAA4030P ('" + dtp_yy_mm.RawData + "','**', 'B1', 'BC',?,?,?,?,?,?,? )}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
    
    If adoCmd(6) <> "" Then
        Call Gp_MsgBoxDisplay(adoCmd(6))
        Set adoCmd = Nothing
        Exit Sub
    Else
        MsgBox "转炉计划编制完了!"
        Set adoCmd = Nothing
        Call Form_Ref
        Exit Sub
    End If
    
plan_cmd_Error:

    Call Gp_MsgBoxDisplay("编制计划错误 : " & Error)

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

    Call Sp_Setting2(ss1)
    Call Sp_Setting2(ss2)
    Call Sp_Setting2(ss3)
    
    Call Gp_Sp_ColGet(ss1, "A-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss2, "A-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
    If Mid(sAuthority, 1, 3) = "111" Then
       ccm_plan_cmd.Enabled = True
       plan_cmd.Enabled = True
    Else
       ccm_plan_cmd.Enabled = False
       plan_cmd.Enabled = False
    End If

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
    
    Call Gp_Sp_ColSet(ss1, "A-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss2, "A-System.INI", Me.Name)
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    dtp_yy_mm.SetFocus
    
    ss1.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, False
    ss2.ClearRange 1, 1, ss2.MaxCols, ss2.MaxRows, False
    ss3.MaxRows = 0
    ss3.MaxCols = 0
End Sub

Public Sub Form_Ref()
    Dim iCol As Integer
    Dim RowTot, ColTot As Double

    If Sp_Refer(ss1) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           txt_excel = "1"
           Call Sp_Refer(ss2)
           Call Sp_Header_Refer(ss3, "SL")
           Call Sp_Data_Refer(ss3)
           
           Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows, True)
           Call Gp_Sp_BlockLock(ss2, 1, ss2.MaxCols, 1, ss2.MaxRows, True)
           Call Gp_Sp_BlockLock(ss3, 1, ss3.MaxCols, 1, ss3.MaxRows, True)

            With ss3
                .MaxRows = .MaxRows + 1
                .MaxCols = .MaxCols + 1
                '列合计
                .Row = .MaxRows
                .Col = SpreadHeader
                .Text = "合计"
                .Col = SpreadHeader + 1
                .Text = "合计"
                
                .Row = .MaxRows:       .Row2 = .MaxRows
                .Col = SpreadHeader:   .Col2 = SpreadHeader + 1
                .ColMerge = MergeAlways
                .RowMerge = MergeAlways
                
                For iCol = 1 To .MaxCols - 1
                    .Col = iCol
                    ColTot = Gf_Sp_ColSum(ss3, .Col, 1, .MaxRows - 1)
                    .Row = .MaxRows
                    If ColTot > 0 Then
                        .Text = ColTot
                    Else
                        .Text = ""
                    End If
                    .CellType = CellTypeNumber
                    .TypeNumberDecPlaces = 3
                    .TypeNumberMax = 999999999
                    .TypeNumberMin = 0
                    .TypeNumberShowSep = True
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
                Next iCol
                
                '行合计
                .Col = .MaxCols
                .Row = SpreadHeader
                .Text = "合计"
                .Row = SpreadHeader + 1
                .Text = "合计"
                
                .Col = .MaxCols:       .Col2 = .MaxCols
                .Row = SpreadHeader:   .Row2 = SpreadHeader + 1
                .ColMerge = MergeAlways
                .RowMerge = MergeAlways
                
                For iCol = 1 To .MaxRows
                    .Row = iCol
                    RowTot = Gf_Sp_RowSum(ss3, .Row, 1, .MaxCols - 1)
                    .Col = .MaxCols
                    If RowTot > 0 Then
                        .Text = RowTot
                    Else
                        .Text = ""
                    End If
                    .CellType = CellTypeNumber
                    .TypeNumberDecPlaces = 3
                    .TypeNumberMax = 999999999
                    .TypeNumberMin = 0
                    .TypeNumberShowSep = True
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
            
                    .ColWidth(.Col) = 11
                Next iCol
                
            End With

    End If
End Sub

Public Sub Form_Exc()
If txt_excel.Text = "1" Then
    Call Gp_Sp_Excel(Me, ss1, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
ElseIf txt_excel.Text = "2" Then
    Call Gp_Sp_Excel(Me, ss2, 0, lBlkcol2, lBlkrow1, lBlkrow2)
ElseIf txt_excel.Text = "3" Then
    Call Gp_Sp_Excel(Me, ss3, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
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
txt_excel = "1"
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
txt_excel = "2"
End Sub

Private Sub ss3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
txt_excel = "3"
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Public Sub Menu_Setting()

    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Row cancel
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Row cancel
    
End Sub

Public Function Sp_Header_Refer(ByVal sPname As Variant, ByVal sProdCd As String) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    Dim sQuery2 As String
    
    Dim AdoRs2 As ADODB.Recordset
    Dim ArrayRecords2 As Variant

    Set adoRs = New ADODB.Recordset
    
    sQuery = "SELECT THK_CD, FR_THK, TO_THK "
    sQuery = sQuery + "   FROM BP_THICK_GRP "
    sQuery = sQuery + "  WHERE PROD_CD = '" + sProdCd + "'"
    sQuery = sQuery + "    AND THK_CD <> '*' "
    sQuery = sQuery + "  ORDER BY THK_CD "
    
    With sPname

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
                
            If sPname Is ss2 Then
               .MaxCols = UBound(ArrayRecords, 2) + 1
                For iCol = 0 To .MaxCols - 1
                
                   .Col = iCol + 1
                   .Row = SpreadHeader + 1
                    If VarType(ArrayRecords(0, iCol)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(1, iCol)) & " ~ " & Trim(ArrayRecords(2, iCol))
                        .Row = SpreadHeader
                        If Trim(ArrayRecords(0, iCol)) = "A" Then
                        .Text = "1#"
                        Else
                        .Text = "2#"
                        End If
                    End If
                               
                    'Column Type Setting
                    .Col = iCol + 1: .Col2 = iCol + 1
                    .Row = 1: .Row2 = -1
                    .BlockMode = True
                    .CellType = 13      'SS_CELL_TYPE_NUMBER
                    .TypeNumberDecPlaces = 3
                    .TypeNumberMax = 99999999999.999
                    .TypeNumberMin = 0
                    .TypeNumberShowSep = True
                    .TypeNumberLeadingZero = TypeLeadingZeroNo
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
                    
                    .BlockMode = False
                    
                    .ColWidth(iCol + 1) = 11
                    
                Next iCol
                .Col = SpreadHeader
                .ColWidth(iCol) = 3
                .Col = SpreadHeader + 1
                .ColWidth(iCol) = 13
                
            ElseIf sPname Is ss3 Then
               .MaxRows = UBound(ArrayRecords, 2) + 1
                For iRow = 0 To .MaxRows - 1
                
                   .Row = iRow + 1
                   .Col = SpreadHeader + 1
                    If VarType(ArrayRecords(0, iRow)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(1, iRow)) & " ~ " & Trim(ArrayRecords(2, iRow))
                        .ColWidth(.Col) = 11
                        .Col = SpreadHeader
                        .Text = Trim(ArrayRecords(0, iRow))
                        .ColWidth(.Col) = 3
                    End If
                               
                    'Column Type Setting
                    .Col = 1: .Col2 = -1
                    .Row = .Row: .Row2 = .Row
                    .BlockMode = True
                    .CellType = 13      'SS_CELL_TYPE_NUMBER
                    .TypeNumberDecPlaces = 3
                    .TypeNumberMax = 99999999999.999
                    .TypeNumberMin = 0
                    .TypeNumberShowSep = True
                    .TypeNumberLeadingZero = TypeLeadingZeroNo
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
                    .BlockMode = False
                Next iRow
            End If
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Set AdoRs2 = New ADODB.Recordset
    
    sQuery2 = "SELECT WID_CD, FR_WID, TO_WID "
    sQuery2 = sQuery2 + "   FROM BP_WIDTH_GRP "
    sQuery2 = sQuery2 + "  WHERE PROD_CD = '" + sProdCd + "' "
    sQuery2 = sQuery2 + "    AND WID_CD <> '*' "
    sQuery2 = sQuery2 + "  ORDER BY WID_CD "
    
    With sPname

        Sp_Header_Refer = True
        .ReDraw = False
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
            If sPname Is ss2 Then
                .MaxRows = UBound(ArrayRecords2, 2) + 1
                
                For iRow = 0 To .MaxRows - 1
                
                    .Row = iRow + 1
                    .Col = SpreadHeader + 1
                    If VarType(ArrayRecords2(0, iRow)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords2(1, iRow)) & " ~ " & Trim(ArrayRecords2(2, iRow))
                        .ColWidth(.Col) = 11
                        .Col = SpreadHeader
                        .Text = Trim(ArrayRecords2(0, iRow))
                        .ColWidth(.Col) = 3
                    End If
                                    
                    .Row = iRow + 1: .Row2 = iRow + 1
                    .Col = 1: .Col2 = -1
                    .BlockMode = True
                    .CellType = 13      'SS_CELL_TYPE_NUMBER
                    .TypeNumberDecPlaces = 3
                    .TypeNumberMax = 99999999999.999
                    .TypeNumberMin = 0
                    .TypeNumberShowSep = True
                    .TypeNumberLeadingZero = TypeLeadingZeroNo
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
                    .BlockMode = False
                Next iRow
            ElseIf sPname Is ss3 Then
                .MaxCols = UBound(ArrayRecords2, 2) + 1
                
                For iCol = 0 To .MaxCols - 1
                
                    .Col = iCol + 1
                    .Row = SpreadHeader + 1
                    
                    If VarType(ArrayRecords2(0, iCol)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords2(1, iCol)) & " ~ " & Trim(ArrayRecords2(2, iCol))
                        .Row = SpreadHeader
                        .Text = Trim(ArrayRecords2(0, iCol))
                    End If
                    .ColWidth(.Col) = 11
                                    
                    .Row = iRow: .Row2 = iRow
                    .Col = 1: .Col2 = -1
                    .BlockMode = True
                    .CellType = 13      'SS_CELL_TYPE_NUMBER
                    .TypeNumberDecPlaces = 3
                    .TypeNumberMax = 99999999999.999
                    .TypeNumberMin = 0
                    .TypeNumberShowSep = True
                    .TypeNumberLeadingZero = TypeLeadingZeroNo
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
                    .BlockMode = False
                Next iCol
            End If
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Set AdoRs2 = Nothing
    Sp_Header_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function


Public Sub Sp_Setting2(ByVal sPname As Variant)

    With sPname
    
        .RowHeight(-1) = 12
        .RowHeight(0) = 16
        
'        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        .OperationMode = OperationModeNormal
        '.RetainSelBlock = True

        '.UserResize = UserResizeNone
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .Row = 0
        .FontBold = True
        
'        .Col = 0
'        .Row = -1
'        .FontBold = True
        
'        .LockBackColor = RGB(255, 255, 255)
        
'        If .Name = "ss3" Then Call Gp_Sp_RowColor(ss3, 3, vbRed)
'        If .Name = "ss4" Then .RowHeadersShow = False
        
    End With
    
End Sub


Public Function Sp_Data_Refer(ByVal sPname As Variant) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
   ' Dim SPARA As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    
    sQuery = "{CALL AAA1230C.P_SREFER3('" + dtp_yy_mm.RawData + "')}"
    
    'Ado Execute
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    
    With sPname

        Sp_Data_Refer = True
        .ReDraw = False
       ' .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
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
            
            For iCnt = 0 To UBound(ArrayRecords, 2)
                If Not (VarType(ArrayRecords(0, iCnt)) = vbNull) Then
                    
                    .Row = Asc(ArrayRecords(0, iCnt)) - 64
                    
                    For iCol = 1 To .MaxCols
                        .Col = iCol
                         
                            If VarType(ArrayRecords(iCol, iCnt)) = vbNull Or ArrayRecords(iCol, iCnt) = 0 Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(iCol, iCnt))
                            End If

                    Next iCol

                End If
            Next iCnt
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function


Public Function Sp_Refer(ByVal sPname As Variant) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    If sPname Is ss2 Then
        sQuery = "{CALL AAA1230C.P_SREFER2('" + dtp_yy_mm.RawData + "')}"
    ElseIf sPname Is ss1 Then
        sQuery = "{CALL AAA1230C.P_SREFER1('" + dtp_yy_mm.RawData + "')}"
    End If
    
    'Ado Execute
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    
    With sPname

        Sp_Refer = True
        .ReDraw = False
       ' .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        If adoRs.BOF Or adoRs.EOF Then
        
            Sp_Refer = False
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
            
            For iCnt = 0 To UBound(ArrayRecords, 2)
                .Row = iCnt + 1
                For iCol = 1 To .MaxCols
                    .Col = iCol
                    If VarType(ArrayRecords(iCol - 1, iCnt)) = vbNull Then
                       .Text = ""
                    Else
                       .Text = Trim(ArrayRecords(iCol - 1, iCnt))
                    End If
                Next iCol
            Next iCnt
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Sp_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function
