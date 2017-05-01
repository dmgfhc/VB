VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACF2090C 
   Caption         =   "轧钢生产月报查询_ACF2090C"
   ClientHeight    =   3375
   ClientLeft      =   510
   ClientTop       =   3390
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3375
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   8595
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   15165
      _Version        =   393216
      _ExtentX        =   26749
      _ExtentY        =   15161
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
      MaxCols         =   17
      MaxRows         =   300
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACF2090C.frx":0000
   End
   Begin InDate.UDate dtp_date 
      Height          =   315
      Left            =   1665
      TabIndex        =   1
      Tag             =   "指示日期"
      Top             =   105
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.74
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
      MaxLength       =   10
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   60
      Top             =   90
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "年月日"
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
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   15165
      Y1              =   570
      Y2              =   570
   End
End
Attribute VB_Name = "ACF2090C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       DAILY SCHEDULE
'-- Sub_System Name
'-- Program Name
'-- Program ID        AEA1100C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.6.20
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
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

  
    Call Gp_Ms_Collection(dtp_date, "p", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    sc1.Add Item:=ss1, Key:="Spread"
  
    sc1.Add Item:="ACF2090C.P_REFER", Key:="P-R"
   
    sc1.Add Item:=pColumn1, Key:="pColumn"
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
   
        
End Sub
Public Sub Sp_Header_display(sPname As Variant)

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim sQuery As String

    
    With sPname

       
        Screen.MousePointer = vbHourglass
        
         .Row = 0
         .Col = 0
         .Text = "日期\项目"
         .ColMerge = 2
         
         
         .Row = SpreadHeader
         .Col = SpreadHeader + 1
         .Text = "班别"
        
         .MaxRows = 0
         
         
         .Col = 0
         .Row = 1
          .Text = "月总计"
         .Col = 0
         .Row = 2
          .Text = "月总计"
         .Col = 0
         .Row = 3
          .Text = "月总计"
         .Col = 0
         .Row = 4

         .Text = "月总计"
         .ReDraw = False
         .Refresh
        
          Screen.MousePointer = vbDefault
        
    End With
    
Exit Sub

SpreadDisplay_Error:
    
 '   Set AdoRs = Nothing
   ' ss1.ReDraw = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste

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
 
    
    dtp_date.Text = DateAdd("d", -1, Date)
    
    
    Call Gp_Ms_Cls(Mc1("rControl"))
  
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Sp_Header_display(Proc_Sc("Sc1")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc1")("Spread"))
 
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "C-System.INI", Me.Name)
   

    Screen.MousePointer = vbDefault
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    

End Sub
Public Sub Sp_Setting(ByVal sPname As Variant)

    Dim iRow As Integer

    With sPname
    
        .RowHeight(-1) = 13
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 13
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 13
        Else
            .RowHeight(0) = 24
        End If
        
       ' .RowHeadersShow = False
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        .OperationMode = OperationModeNormal
        .RetainSelBlock = True
        .UserResize = UserResizeColumns
        
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
'
'        .Col = -1
'        .Row = 1
'        .FontBold = True
       
   
'        '-------------第一列
'        .Col = 1
'        .Row = -1
'        .FontSize = 15

        
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "C-System.INI", Me.Name)
    
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
    

 ' ss1.ClearRange 0, 1, ss1.MaxCols, ss1.MaxRows, True
 ' Call Gp_Sp_BlockColor(Proc_Sc("Sc1")("Spread"), 3, ss1.MaxCols, 1, ss1.MaxRows)
 ' Call Gp_Ms_Cls(Mc1("rControl"))
   Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
  Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    

 
  Call Sp_Header_display(Proc_Sc("Sc1")("Spread"))



End Sub

Public Sub Form_Ref()


Dim P_dtp_date As String
P_dtp_date = dtp_date.RawData
Dim sQuery As String
On Error GoTo Refer_Err

 sQuery = "{ CALL " + "ACF2090C.P_REFER_M" + "("
 sQuery = sQuery + " '" + P_dtp_date + "'"
 sQuery = sQuery + ")"
 sQuery = sQuery + "}"
 If Gf_Sp_ProceExist(Proc_Sc("Sc1").Item("Spread")) Then Exit Sub

 Call Sp_Display_M(M_CN1, ss1, sQuery, , True)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1").Item("Spread")) Then Exit Sub

    If Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
        MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
        MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
    End If
            
    Exit Sub

Refer_Err:

End Sub

'-----------------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Refer
'   2.Name         : Spread Refer
'   3.Input  Value : Conn Connection, Sc Collection, Mc Collection, {nCheckControl Collection},
'                                        {mCheckControl Collection},{MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Refer
'-----------------------------------------------------------------------------------------------
Public Function Sp_Refer(Conn As ADODB.Connection, Sc As Collection, Optional MC As Collection, _
                            Optional nCheckControl As Collection, Optional mCheckControl As Collection, Optional MsgChk As Boolean = True) As Boolean
On Error GoTo SpreadRef_Error

    Dim sQuery As String
    Dim sMsg As String



    If Not MC Is Nothing Then
    
        If Not nCheckControl Is Nothing Then
            sMsg = Gf_Ms_NeceCheck(nCheckControl)
            If sMsg <> "OK" Then
                sMsg = sMsg + "必须输入"
                Call Gp_MsgBoxDisplay(sMsg, "", "错误提示")
                Sp_Refer = False
                Exit Function
            End If
        End If
        
        If Not mCheckControl Is Nothing Then
            sMsg = Gf_Ms_NeceCheck2(mCheckControl)
            If sMsg <> "OK" Then
                sMsg = sMsg + "长度不正确"
                Call Gp_MsgBoxDisplay(sMsg, "", "错误提示")
                Sp_Refer = False
                Exit Function
            End If
        End If
        
    End If

    Sc.Item("Spread").OperationMode = OperationModeNormal
    
    If Not MC Is Nothing Then
        Sp_Refer = Sp_Display(Conn, Sc.Item("Spread"), Gf_Ms_MakeQuery(Sc.Item("P-R"), "R", MC("pControl")), _
                                    Sc.Item("pColumn"), MsgChk)
        If Sp_Refer Then
           Call Gp_Ms_ControlLock(MC!lControl, True)
           MDIMain.StatusBar1.Panels(1) = "提示信息：查询成功"
        End If
    Else
        Sp_Refer = Sp_Display(Conn, Sc.Item("Spread"), Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-R"), _
                                    "R", Sc.Item("aColumn"), 1), Sc.Item("pColumn"), MsgChk)
    End If
    
    If Sp_Refer Then
     '   Sc.Item("Spread").OperationMode = OperationModeRow
        MDIMain.StatusBar1.Panels(1) = "提示信息：查询成功"
        'Sc!Spread.SetFocus
    End If
        
    Exit Function
    
SpreadRef_Error:

    Call Gp_MsgBoxDisplay("Gf_Sp_Refer Error : " & Error)
    Sp_Refer = False

End Function
Public Sub Form_Pro()

    Dim lRow As Long
    Dim lCnt As Long
    Dim lSeq1 As Long
    Dim lSeq2 As Long
    Dim SMESG As String
    Dim sStlgrd1 As String
    Dim sStlgrd2 As String
    
    For lRow = 1 To ss1.MaxRows
        
        ss1.Row = lRow
        
        ss1.Col = 1
        sStlgrd1 = ss1.Text
        ss1.Col = 2
        lSeq1 = ss1.Text
        
        For lCnt = lRow + 1 To ss1.MaxRows
        
            ss1.Row = lCnt
            ss1.Col = 1
            sStlgrd2 = ss1.Text
            ss1.Col = 2
            lSeq2 = ss1.Text
            
            If sStlgrd1 = sStlgrd2 And lSeq1 = lSeq2 Then
                Call Gp_Sp_RowColor(ss1, lCnt, , vbYellow)
                Call Gp_MsgBoxDisplay(" Invalid 优先顺序 ")
                Exit Sub
            End If
            
        Next lCnt
    
    Next lRow
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
        MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
        MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
        MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
    End If
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
    
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
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

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

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
   If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC1")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), 7)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)
'
   If Proc_Sc("Sc1")("Spread").MaxRows < 1 Then Exit Sub
    
   If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
End If

   If Shift = 0 Then Proc_Sc("Sc1")("Spread").EditMode = True
End Sub

Private Sub ss1_LostFocus()

 lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    
    MDIMain.Mnu_Sorting.Enabled = False

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

    MDIMain.Mnu_Sorting.Enabled = True
    
    
End Sub

Private Sub txt_stlgrd_grp_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Q0048"
      '  DD.rControl.Add Item:=txt_stlgrd_grp
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

End Sub


'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Sp_Display
'   2.Name         : Spread Row Display
'   3.Input  Value : Conn Connection, sPname vaSpread, sQuery String, {lColumn Variant}, {MsgChk Boolean}
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 05 .06
'   7.Modify Date  :
'   8.Comment      : Spread Row Display
'---------------------------------------------------------------------------------------
Public Function Sp_Display(Conn As ADODB.Connection, sPname As vaSpread, sQuery As String, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

 On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If
   
 
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Display = True
        
        .ReDraw = False
        '.MaxRows = 0: iCount = 0     不刷新 0列,0行
        
      '  .ClearRange 3, 1, .MaxCols, .MaxRows, True
        
         Screen.MousePointer = vbHourglass
        
       ' Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            Sp_Display = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
        Else
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

     .MaxRows = UBound(ArrayRecords, 2) + 6
       
'        .MaxRows = 8
        For iRowCount = 0 To .MaxRows - 6
            .Row = iRowCount + 6
            .Col = 0

            For iColcount = 0 To 1 '(0,0)时,
    
    
    
    
        
                Select Case .CellType
                
                    Case SS_CELL_TYPE_CHECKBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) <> vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "1" Then
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "" Then
                            .Value = 0
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Mid(Trim(ArrayRecords(iColcount, iRowCount)), 1, 4) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 5, 2) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 7, 2)
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Value = ""
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case Else
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                        
                        If .Text = "A1" Then
                        .Text = "A组"
                        End If
                        If .Text = "B1" Then
                        .Text = "B组"
                        End If
                        If .Text = "C1" Then
                        .Text = "C组"
                        End If
                        If .Text = "D1" Then
                        .Text = "D组"
                        End If
                         If .Text = "T1" Then
                        .Text = "总计"
                        End If
                End Select
                .Col = SpreadHeader + 1
            Next iColcount
            
            .Col = 1
            For iColcount = 2 To .MaxCols + 1 '(0,0)时,

                Select Case .CellType
                
                    Case SS_CELL_TYPE_CHECKBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) <> vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "1" Then
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "" Then
                            .Value = 0
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Mid(Trim(ArrayRecords(iColcount, iRowCount)), 1, 4) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 5, 2) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 7, 2)
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Value = ""
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case Else
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                End Select
                
                .Col = .Col + 1
                
            Next iColcount
            
            
            
        Next iRowCount
            
        If Not lColumn Is Nothing Then

            'lControl Lock
            For iCount = 3 To lColumn.Count

                .Protect = True
                .Col = lColumn(iCount): .Col2 = lColumn(iCount)
                .Row = 1:               .Row2 = .MaxRows
                .BlockMode = True: .Lock = True
                .BlockMode = False

            Next iCount

        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Display = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Display Error : " & sQuery)
    Screen.MousePointer = vbDefault

End Function


Public Function Sp_Display_M(Conn As ADODB.Connection, sPname As vaSpread, sQuery As String, _
                              Optional lColumn As Variant = Nothing, Optional MsgChk As Boolean = True) As Boolean

 On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display_M = False: Exit Function
    End If
   
 
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        Sp_Display_M = True
        
        .ReDraw = False
     
        
         Screen.MousePointer = vbHourglass
        
       ' Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            If MsgChk Then Call Gp_MsgBoxDisplay("无相关记录", "I")
                
            Sp_Display_M = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            Exit Function
        Else
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing
        
        
        .MaxRows = 6
       For iRowCount = 0 To .MaxRows - 2
    
        .Row = iRowCount + 1     '数组为(0,0)时，显示在(1,2)上
        
                                    '在一个行下，每次都让列从1开始，因为界面上的列从2开始
        .Col = SpreadHeader
    
        
        
          
            For iColcount = 0 To 1 '(0,0)时,
        
                Select Case .CellType
                
                    Case SS_CELL_TYPE_CHECKBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) <> vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "1" Then
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "" Then
                            .Value = 0
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Mid(Trim(ArrayRecords(iColcount, iRowCount)), 1, 4) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 5, 2) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 7, 2)
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Value = ""
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case Else
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                        If .Text = "A2" Then .Text = "A组"
                        If .Text = "B2" Then .Text = "B组"
                        If .Text = "C2" Then .Text = "C组"
                        If .Text = "D2" Then .Text = "D组"
                        If .Text = "T2" Then .Text = "总计"
                        
                End Select
                .Col = SpreadHeader + 1
            Next iColcount
            
            .Col = 1
            For iColcount = 2 To .MaxCols + 1 '(0,0)时,

                Select Case .CellType
                
                    Case SS_CELL_TYPE_CHECKBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) <> vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "1" Then
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Or _
                           Trim(ArrayRecords(iColcount, iRowCount)) = "" Then
                            .Value = 0
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Mid(Trim(ArrayRecords(iColcount, iRowCount)), 1, 4) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 5, 2) & "-" & _
                                    Mid(Trim(ArrayRecords(iColcount, iRowCount)), 7, 2)
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Value = ""
                        Else
                            .Value = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                    Case Else
                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
                        End If
                        
                End Select
                
                .Col = .Col + 1
                
            Next iColcount
            
            
            
        Next iRowCount
            
                            
                        .Col = 0
                        .Row = 1
                         .Text = "月总计"
                        .Col = 0
                        .Row = 2
                         .Text = "月总计"
                        .Col = 0
                        .Row = 3
                         .Text = "月总计"
                        .Col = 0
                        .Row = 4
                        .Text = "月总计"
                        .Col = 0
                        .Row = 5
                        .Text = "月总计"
     
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Display_M = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Display Error : " & sQuery)
    Screen.MousePointer = vbDefault

End Function

