VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form DGA1180C 
   Caption         =   "热处理车间接收/转出总计查询_DGA1180C"
   ClientHeight    =   5610
   ClientLeft      =   360
   ClientTop       =   2220
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   13710
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_prc_line 
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
      Left            =   30
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   210
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3420
      Top             =   90
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   556
      Caption         =   "作业日期"
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
   Begin InDate.UDate sdt_wrk_date_fr 
      Height          =   315
      Left            =   4545
      TabIndex        =   2
      Tag             =   "作业日期"
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
      MaxLength       =   10
   End
   Begin InDate.UDate sdt_wrk_date_to 
      Height          =   315
      Left            =   6300
      TabIndex        =   3
      Tag             =   "作业日期"
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
      MaxLength       =   10
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8715
      Left            =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   15210
      _Version        =   393216
      _ExtentX        =   26829
      _ExtentY        =   15372
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
      MaxCols         =   1
      MaxRows         =   20
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "DGA1180C.frx":0000
   End
   Begin Threed.SSOption opt_htn_plt1 
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1号热处理线"
      Value           =   -1
   End
   Begin Threed.SSOption opt_htn_plt2 
      Height          =   285
      Left            =   1740
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2号热处理线"
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   120
      Left            =   6060
      TabIndex        =   4
      Top             =   195
      Width           =   195
   End
End
Attribute VB_Name = "DGA1180C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   HTM System
'-- Program Name      热处理实绩总计查询
'-- Program ID        DGA1180C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim.Sung.Ho
'-- Coder             Kim.Sung.Ho
'-- Date              2008.3.26
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

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Refer"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_PRC_LINE, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdt_wrk_date_fr, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(sdt_wrk_date_to, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
               
     'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
   
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="DGA1180C.P_REFER", Key:="P-R"
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

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    txt_PRC_LINE.Text = "1"
    sdt_wrk_date_fr.RawData = Mid(sdt_wrk_date_fr.RawData, 1, 6) & "01"
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "DG-System.INI", Me.Name)
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "DG-System.INI", Me.Name)
    
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
        txt_PRC_LINE.Text = "1"
        sdt_wrk_date_fr.RawData = Mid(sdt_wrk_date_fr.RawData, 1, 6) & "01"
    End If

End Sub

Public Sub Form_Ref()

    Dim sQuery As String    'Header Display
    Dim sQuery1 As String   'STDSPEC TOTAL Display
    Dim sQuery2 As String   'IN OUT PRC TOTAL Display
    Dim lCol As Integer
    
    'Header Display
    sQuery = "          SELECT  CD_NAME "
    sQuery = sQuery + "  FROM  ZP_CD "
    sQuery = sQuery + " WHERE  CD_MANA_NO = 'G0029' "
    sQuery = sQuery + " ORDER  BY CD "
    
    'Header Display
    Call Sp_Header_Refer1(ss1, sQuery)      'Header Display
    
    'STDSPEC TOTAL Display
    sQuery1 = " {call DGA1180C.P_TOTAL_STDSPEC ('" & txt_PRC_LINE.Text & "','" & sdt_wrk_date_fr.RawData & "','" & _
                                                                                 sdt_wrk_date_to.RawData & "')} "
    
    'IN_OUT PRC TOTAL Display
    sQuery2 = " {call DGA1180C.P_TOTAL_PRC     ('" & txt_PRC_LINE.Text & "','" & sdt_wrk_date_fr.RawData & "','" & _
                                                                                 sdt_wrk_date_to.RawData & "')} "
    
    
    If Sp_Display(M_CN1, sc1.Item("Spread"), Gf_Ms_MakeQuery(sc1.Item("P-R"), "R", Mc1("pControl")), _
                                    sc1.Item("pColumn"), False) Then
                                    
        Call StdSpec_Total(M_CN1, ss1, sQuery1)
        Call In_Out_Prc_Total(M_CN1, ss1, sQuery2)
        ss1.OperationMode = OperationModeNormal
        
        For lCol = 1 To ss1.MaxCols
            ss1.Col = lCol
            ss1.Row = SpreadHeader + (ss1.ColHeaderRows - 1)
            If ss1.Text = "合计" Then
                Call Gp_Sp_BlockColor(ss1, lCol, lCol, 1, ss1.MaxRows, vbRed)
            End If
        Next lCol
        
        Call Gp_Sp_EvenRowBackcolor(ss1, 1)
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.MaxRows, ss1.MaxRows, vbRed, &HE6E6FF)
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    End If
    
    
'    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
'        ss1.OperationMode = OperationModeNormal
'        ss1.Row = ss1.MaxRows
'        ss1.Col = 1
'        ss1.Text = "合  计"
'        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.MaxRows, ss1.MaxRows, BLACK, &HE6E6FF)
'        Call Gp_Sp_EvenRowBackcolor(ss1, 1)
'        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'    End If
    
End Sub

Public Sub Form_Pro()

End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

End Sub

Public Sub Spread_Del()
    
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

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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

Private Function Sp_Header_Refer1(sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay1_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Header_Refer1 = True
        
        .ReDraw = False
        .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer1 = False
            '.ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing
        
        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .ColWidth(1) = 19
            
            iColCnt = UBound(ArrayRecords, 2) + 1
            .MaxCols = ((UBound(ArrayRecords, 2) + 1) * 2) + 3
            
            For iCol = 1 To iColCnt
            
                .Col = iCol + 1
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                .Text = "接收"
                
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                If VarType(ArrayRecords(0, iCol - 1)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCol - 1))
                End If
                
                .ColWidth(iCol + 1) = 12
                
                .Col = iCol + 2 + iColCnt
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                .Text = "转出"
                
                .Row = SpreadHeader + (.ColHeaderRows - 1)
                If VarType(ArrayRecords(0, iCol - 1)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCol - 1))
                End If
                
                .ColWidth(iCol + 2 + iColCnt) = 12
                
            Next iCol
            
            .Col = iColCnt + 2
            .Row = SpreadHeader + (.ColHeaderRows - 2)
            .Text = "接收"
            .Row = SpreadHeader + (.ColHeaderRows - 1)
            .Text = "合计"
            .ColWidth(.Col) = 10
            
            .Col = .MaxCols
            .Row = SpreadHeader + (.ColHeaderRows - 2)
            .Text = "转出"
            .Row = SpreadHeader + (.ColHeaderRows - 1)
            .Text = "合计"
            .ColWidth(.Col) = 10
            
            .Col = 2: .Col2 = -1
            .Row = 1: .Row2 = -1
            .BlockMode = True
            .TypeHAlign = TypeHAlignCenter
            .TypeVAlign = TypeVAlignCenter
            .BlockMode = False
            
            .Col = 1: .Col2 = 1
            .Row = 1: .Row2 = -1
            .BlockMode = True
            .TypeHAlign = TypeHAlignLeft
            .BlockMode = False
            
        End If
        
        .BlockMode = True
        .Row = SpreadHeader + (.ColHeaderRows - 2)
        .Col = 2
        .Row2 = SpreadHeader + (.ColHeaderRows - 2)
        .Col2 = .MaxCols
        .RowMerge = MergeAlways
        '.ColMerge = MergeAlways
        .BlockMode = False
        
        .ReDraw = True
        .Refresh
        
        Screen.MousePointer = vbDefault
        
    End With
        
    Exit Function

SpreadDisplay1_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer1 = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay1_Error : " & Error)
    
End Function

Private Function Sp_Display(Conn As ADODB.Connection, sPname As vaSpread, sQuery As String, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim icount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim Stdspec As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Display = True
        
        .ReDraw = False
        .MaxRows = 0: icount = 0
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            Sp_Display = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        For iRowCount = 0 To UBound(ArrayRecords, 2)
            
            If Trim(ArrayRecords(0, iRowCount)) <> Stdspec Then
                Stdspec = Trim(ArrayRecords(0, iRowCount))
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .Col = 1
                .Text = Stdspec
            End If
            
            For iColcount = 2 To .MaxCols
                .Col = iColcount
                
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                If .Text = Trim(ArrayRecords(2, iRowCount)) Then
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 1)
                    If .Text = Trim(ArrayRecords(1, iRowCount)) Then
                        .Row = .MaxRows
                        .Text = Trim(ArrayRecords(3, iRowCount))
                    End If
                    
                End If
                
            Next iColcount
            
        Next iRowCount
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Display = False
    Call Gp_MsgBoxDisplay("Sp_Display Error : " & sQuery)
    Screen.MousePointer = vbDefault
    
End Function

Private Function StdSpec_Total(Conn As ADODB.Connection, sPname As vaSpread, sQuery As String, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo Stdspec_Total_Error

    Dim icount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim Stdspec As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then StdSpec_Total = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        StdSpec_Total = True
        
        .ReDraw = False
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            StdSpec_Total = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        icount = 0
        For iRowCount = 0 To UBound(ArrayRecords, 2)
            
            For icount = 1 To .MaxRows
                .Row = icount
                .Col = 1
                If Trim(ArrayRecords(0, iRowCount)) = .Text Then
                    Exit For
                End If
            Next icount
            
            For iColcount = 2 To .MaxCols
                .Col = iColcount
                
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                If .Text = Trim(ArrayRecords(1, iRowCount)) Then
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 1)
                    If .Text = "合计" Then
                        .Row = icount
                        .Text = Trim(ArrayRecords(2, iRowCount))
                    End If
                    
                End If
                
            Next iColcount
            
        Next iRowCount
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

Stdspec_Total_Error:
    
    Set AdoRs = Nothing
    StdSpec_Total = False
    Call Gp_MsgBoxDisplay("StdSpec_Total Error : " & sQuery)
    Screen.MousePointer = vbDefault
    
End Function

Private Function In_Out_Prc_Total(Conn As ADODB.Connection, sPname As vaSpread, sQuery As String, _
                                  Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

On Error GoTo In_Out_Prc_Total_Error

    Dim icount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim Stdspec As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then In_Out_Prc_Total = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        In_Out_Prc_Total = True
        
        .ReDraw = False
        
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            In_Out_Prc_Total = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        icount = 0
        
        .MaxRows = .MaxRows + 1
        .Col = 1
        .Row = .MaxRows
        .Text = "合  计"
        
        For iRowCount = 0 To UBound(ArrayRecords, 2)
            
            .Row = .MaxRows
            For iColcount = 2 To .MaxCols
                .Col = iColcount
                
                .Row = SpreadHeader + (.ColHeaderRows - 2)
                If .Text = Trim(ArrayRecords(1, iRowCount)) Then
                    
                    .Row = SpreadHeader + (.ColHeaderRows - 1)
                    If .Text = Trim(ArrayRecords(0, iRowCount)) Then
                        .Row = .MaxRows
                        .Text = Trim(ArrayRecords(2, iRowCount))
                    End If
                    
                End If
                
            Next iColcount
            
        Next iRowCount
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

In_Out_Prc_Total_Error:
    
    Set AdoRs = Nothing
    In_Out_Prc_Total = False
    Call Gp_MsgBoxDisplay("In_Out_Prc_Total Error : " & sQuery)
    Screen.MousePointer = vbDefault
    
End Function

