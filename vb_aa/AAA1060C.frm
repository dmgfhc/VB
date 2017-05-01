VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1060C 
   Caption         =   "工厂制约条件录入_AAA1060C"
   ClientHeight    =   8640
   ClientLeft      =   1005
   ClientTop       =   4725
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_plan_wk 
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
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "日期"
      Top             =   90
      Width           =   825
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
      Left            =   10485
      TabIndex        =   5
      Top             =   90
      Width           =   4290
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
      Left            =   9135
      MaxLength       =   11
      TabIndex        =   4
      Tag             =   "钢种"
      Top             =   90
      Width           =   1320
   End
   Begin VB.TextBox txt_prod_NAME 
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
      Left            =   5745
      MaxLength       =   80
      TabIndex        =   3
      Tag             =   "产品"
      Top             =   90
      Width           =   1680
   End
   Begin VB.TextBox txt_prod_cd 
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
      Left            =   5265
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "产品"
      Top             =   90
      Width           =   465
   End
   Begin InDate.ULabel ULabel6 
      Height          =   285
      Left            =   7785
      Top             =   90
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
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel5 
      Height          =   300
      Left            =   3915
      Top             =   90
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      Caption         =   "产品"
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
   Begin InDate.UDate dtp_date_str 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Tag             =   "日期"
      Top             =   90
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8685
      Left            =   90
      TabIndex        =   6
      Top             =   450
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   15319
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
      SpreadDesigner  =   "AAA1060C.frx":0000
   End
End
Attribute VB_Name = "AAA1060C"
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
'-- Program ID        AAA1060C
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
      Call Gp_Ms_Collection(dtp_date_str, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(cbo_plan_wk, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
    Sc1.Add Item:="AAA1060C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
 

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub dtp_date_str_Validate(Cancel As Boolean)
    
    Dim dType As Variant
    
    dType = Gf_FloatFind(M_CN1, "SELECT PERIOD FROM AP_PLAN_UNIT WHERE YEAR_MONTH = '" & Mid(dtp_date_str.RawData, 1, 6) & "'")
    
    cbo_plan_wk.Clear
    cbo_plan_wk.AddItem " "
    
    Select Case dType
        
        Case 5
            cbo_plan_wk.AddItem "A1"
            cbo_plan_wk.AddItem "A2"
            cbo_plan_wk.AddItem "A3"
            cbo_plan_wk.AddItem "A4"
            cbo_plan_wk.AddItem "A5"
            cbo_plan_wk.AddItem "A6"
            
        Case 10
            cbo_plan_wk.AddItem "B1"
            cbo_plan_wk.AddItem "B2"
            cbo_plan_wk.AddItem "B3"
            
        Case 15
            cbo_plan_wk.AddItem "C1"
            cbo_plan_wk.AddItem "C2"
        Case 30
            cbo_plan_wk.AddItem "D0"
            
        Case Else
            cbo_plan_wk.AddItem "A1"
            cbo_plan_wk.AddItem "A2"
            cbo_plan_wk.AddItem "A3"
            cbo_plan_wk.AddItem "A4"
            cbo_plan_wk.AddItem "A5"
            cbo_plan_wk.AddItem "A6"
            
    End Select
    
    cbo_plan_wk.ListIndex = 0
    
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
    
    dtp_date_str.RawData = Format(Now, "YYYYMM")
    
    Call dtp_date_str_Validate(True)
    
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
    
    ss1.MaxCols = 0
    ss1.MaxRows = 0
        
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call dtp_date_str_Validate(True)
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    rControl(1).SetFocus

End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
       
        If Sp_Header_Refer() Then
        
            If Left(dtp_date_str.RawData, 6) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Or txt_stlgrd.Text = "" Then
               Call Gp_Sp_BlockLock(ss1, 1, -1, 1, -1, True)
            Else
                Call Gp_Sp_BlockLock(ss1, 1, -1, 1, -1, False)
              
            End If
            
            If Sp_Data_Refer() Then
          
                Call SubSpreadSum
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Menu_Setting
        '        Call Gp_Ms_ControlLock(Mc1!lControl, True)

            End If
            
            Call Gp_Ms_ControlLock(Mc1!lControl, True)
             
        End If
            
    Else
        sMesg = sMesg + " 必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
End Sub

Public Sub Form_Pro()
    
    Dim sMesg As String
    Dim sCurrentDay As String
    Dim sCurrentType As String
    Dim dType As Variant
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg <> "OK" Then
        sMesg = sMesg + " 必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If Trim(txt_stlgrd.Text) = "" Then
        sMesg = " 钢种 必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    sCurrentDay = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    dType = Gf_FloatFind(M_CN1, "SELECT PERIOD FROM AP_PLAN_UNIT WHERE YEAR_MONTH = '" & Mid(sCurrentDay, 1, 6) & "'")
    
    Select Case dType
        
        Case 5
            
            If Mid(sCurrentDay, 7, 2) >= "01" And Mid(sCurrentDay, 7, 2) <= "05" Then
                sCurrentType = "A1"
            ElseIf Mid(sCurrentDay, 7, 2) >= "06" And Mid(sCurrentDay, 7, 2) <= "10" Then
                sCurrentType = "A2"
            ElseIf Mid(sCurrentDay, 7, 2) >= "11" And Mid(sCurrentDay, 7, 2) <= "15" Then
                sCurrentType = "A3"
            ElseIf Mid(sCurrentDay, 7, 2) >= "16" And Mid(sCurrentDay, 7, 2) <= "20" Then
                sCurrentType = "A4"
            ElseIf Mid(sCurrentDay, 7, 2) >= "21" And Mid(sCurrentDay, 7, 2) <= "25" Then
                sCurrentType = "A5"
            Else
                sCurrentType = "A6"
            End If
            
        Case 10
        
            If Mid(sCurrentDay, 7, 2) >= "01" And Mid(sCurrentDay, 7, 2) <= "10" Then
                sCurrentType = "B1"
            ElseIf Mid(sCurrentDay, 7, 2) >= "11" And Mid(sCurrentDay, 7, 2) <= "20" Then
                sCurrentType = "B2"
            Else
                sCurrentType = "B3"
            End If
            
        Case 15
        
            If Mid(sCurrentDay, 7, 2) >= "01" And Mid(sCurrentDay, 7, 2) <= "15" Then
                sCurrentType = "C1"
            Else
                sCurrentType = "C2"
            End If
            
        Case 30
            
            sCurrentType = "C2"
            
    End Select
    
    'BEFORE CURRENT DAY PROCESS IMPOSSIBLE
    If Mid(dtp_date_str.RawData, 1, 6) & cbo_plan_wk.Text < Mid(sCurrentDay, 1, 6) & sCurrentType Then
        sMesg = " 日期 输入错误"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If Sp_Process(M_CN1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
        Call Form_Ref
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
    
    If ChangeMade Then
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
   
'   Min<=Max check
    Dim iCol As Integer
    Dim iRow As Integer
    Dim dMin As Double
    Dim dMax As Double
    
    If Row < 0 Or Row = 0 Then Exit Sub
    
    With ss1
    
        If .CellTag = "False" Then Exit Sub
            
        .Row = Row
        Select Case Col Mod 2
        
            Case 0      'MAX
            
                .Col = Col - 1
                If .Value = "" Then
                    dMin = 0
                Else
                    dMin = .Value
                End If
                
                .Col = Col
                If .Value = "" Then
                    dMax = 0
                Else
                    dMax = .Value
                End If
                                
                If dMin = 0 Then Exit Sub
                
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
           
            Case 1      'MIN
                
                .Col = Col
                If .Value = "" Then
                    dMin = 0
                Else
                    dMin = .Value
                End If
                
                .Col = Col + 1
                
                If .Value = "" Then
                    dMax = 0
                Else
                    dMax = .Value
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
     ''   .RowHidden = True
        
        .ColWidth(SpreadHeader) = 8
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        
        .BlockMode = True
        .ColMerge = MergeAlways
        .BlockMode = False
        
        .Row = SpreadHeader
        .Col = SpreadHeader
        .Text = "宽度\厚度"
        .Row = SpreadHeader + 1
        .Col = SpreadHeader
        .Text = "宽度\厚度"
        
        .Row = SpreadHeader + 2
        .RowHidden = True
        
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

Public Function Sp_Header_Refer() As Boolean

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
    sQuery = sQuery + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    sQuery = sQuery + "    AND THK_CD <> '*' "
    sQuery = sQuery + "  ORDER BY THK_CD "
    
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
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
    
                .Col = iCol + 2
                .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                           
                .Col = iCol + 1:  .Row = SpreadHeader + 1:  .Text = "最小量"
                .Col = iCol + 2:  .Row = SpreadHeader + 1:  .Text = "最大量"
                
                .Col = iCol + 1
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                .Col = iCol + 2
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                'Column Type Setting
                .Col = iCol + 1: .Col2 = iCol + 1
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                .ColWidth(iCol + 1) = 9
                
                .Col = iCol + 2: .Col2 = iCol + 2
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 0
                .TypeNumberMax = 9999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                .ColWidth(iCol + 2) = 9
                iCnt = iCnt + 1
                
            Next iCol
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Set AdoRs2 = New ADODB.Recordset
    
    sQuery2 = "SELECT WID_CD, FR_WID, TO_WID "
    sQuery2 = sQuery2 + "   FROM BP_WIDTH_GRP "
    sQuery2 = sQuery2 + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    sQuery2 = sQuery2 + "    AND WID_CD <> '*' "
    sQuery2 = sQuery2 + "  ORDER BY WID_CD "
    
    With ss1

        Sp_Header_Refer = True
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
            
            iCnt = 0
            
            For iRow = 1 To .MaxRows
            
                .Row = iRow
                .Col = SpreadHeader
                
                If VarType(ArrayRecords2(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords2(1, iCnt)) & " ~ " & Trim(ArrayRecords2(2, iCnt)) & "mm"
                End If
                
                .Col = SpreadHeader + 1
                .Text = Trim(ArrayRecords2(0, iCnt))
                
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
    
    Call Gp_Sp_EvenRowBackcolor(Sc1.Item("Spread"), 0)
    Call Gp_Sp_ColHidden(ss1, SpreadHeader + ss1.ColHeaderRows - 3, False)
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
    Dim sTdate As String
    Dim sQuery As String
    Dim sEdate As String
    Dim sWID_GRP As String
    Dim sTHK_GRP As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    
    sEdate = Mid(dtp_date_str.Text, 1, 4) + Mid(dtp_date_str.Text, 6, 2)
  
    sQuery = "SELECT WID_GRP, THK_GRP, MIN,MAX"
    sQuery = sQuery + "   FROM  AP_LIMIT_CON "
    sQuery = sQuery + "  WHERE  YEAR_MONTH =   '" + sEdate + "' "
    sQuery = sQuery + "    AND  PLAN_WK    =   '" + Trim(cbo_plan_wk.Text) + "' "
    sQuery = sQuery + "    AND  PROD_CD LIKE   '" + Trim(txt_prod_cd.Text) + "%' "
    sQuery = sQuery + "    AND  STLGRD  LIKE   '" + Trim(txt_stlgrd.Text) + "%' "
    'sQuery = sQuery + "    GROUP BY WID_GRP, THK_GRP "
    sQuery = sQuery + "    ORDER BY WID_GRP, THK_GRP "
    
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
            iRow = 1
            For iCnt = 0 To UBound(ArrayRecords, 2)
                .Row = iRow
                .Col = SpreadHeader + 1
                 sWID_GRP = .Text
                 Do While iRow <= .MaxRows And sWID_GRP <> Trim(ArrayRecords(0, iCnt))
                    iRow = iRow + 1
                    .Row = iRow
                    sWID_GRP = .Text
                 Loop
                           
                 For iCol = 1 To .MaxCols Step 2
                    .Col = iCol
                    .Row = SpreadHeader + 2
                    sTHK_GRP = .Text

                    If sTHK_GRP = ArrayRecords(1, iCnt) Then
                        .Row = iRow
                     
                        If VarType(ArrayRecords(2, iCnt)) = vbNull Or ArrayRecords(2, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(2, iCnt))
                        End If
                        
                        .Col = iCol + 1
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Or ArrayRecords(3, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(3, iCnt))
                        End If
                
                    End If

                Next iCol
                
            Next iCnt
            
        End If
        
     '   .ReDraw = True
        
        MDIMain.StatusBar1.Panels(1) = "提示信息 : 数据查询完成"
        Screen.MousePointer = vbDefault
        
    End With
    
Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

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
        
        For iRow = 1 To .MaxRows - 1
            
            .Row = iRow
            
            'Parameters Setting
            For iCol = 1 To .MaxCols - 1 Step 2
            
                .Col = iCol
                If Trim(.Text) <> "" Then
                
                    adoCmd.Parameters(0).Value = Mid(dtp_date_str.Text, 1, 4) + _
                                                  Mid(dtp_date_str.Text, 6, 2)       'YEAR_MONTH
                    adoCmd.Parameters(1).Value = cbo_plan_wk.Text                    'PLAN_WK
                    adoCmd.Parameters(2).Value = txt_prod_cd.Text                    'PROD_CD
                    adoCmd.Parameters(3).Value = txt_stlgrd.Text                     'STLGRD
                    
                    .Row = SpreadHeader + 2
                    .Col = iCol
                    adoCmd.Parameters(4).Value = .Text     'THK_GRP
                   
                    .Row = iRow
                    .Col = SpreadHeader + 1
                    adoCmd.Parameters(5).Value = .Text     'WID_GRP
                    
                    .Col = iCol
                 
                    If Trim(.Text) = "" Then                'MIN
                        adoCmd.Parameters(6).Value = 0
                    Else
                        dTempInt = .Text
                        adoCmd.Parameters(6).Value = dTempInt
                    End If
                    
                    .Col = iCol + 1
                 
                    If Trim(.Text) = "" Then                'MAX
                        adoCmd.Parameters(7).Value = 0
                    Else
                        dTempInt = .Text
                        adoCmd.Parameters(7).Value = dTempInt
                    End If
                    
                    adoCmd.Parameters(8).Value = sUserID                             'User-id
                    
                    adoCmd.Execute
                    
                    'Error Check
                    If adoCmd("Error") <> "0" Then
                        
                        ret_Result_ErrCode = adoCmd("Error")
                        ret_Result_ErrMsg = adoCmd("Messg")
                        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                        
                        Call Gp_Sp_CellColor(ss1, iCol, iRow, , vbYellow)
                        Call Gp_Sp_CellColor(ss1, iCol + 1, iRow, , vbYellow)
                        
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

Private Sub txt_prod_cd_Change()
    If Trim(txt_prod_cd.Text) = "" Then
        txt_prod_NAME.Text = ""
    End If
End Sub

Private Sub txt_prod_cd_KeyPress(KeyAscii As Integer)

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
   
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        DD.rControl.Add Item:=txt_prod_cd
        DD.rControl.Add Item:=txt_prod_NAME

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_prod_cd)) = txt_prod_cd.MaxLength Then
        txt_prod_NAME.Text = Gf_ComnNameFind(M_CN1, "B0005", Trim(txt_prod_cd.Text), 2)
    Else
        txt_prod_NAME.Text = ""
    End If

End Sub

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

Private Sub SubSpreadSum()

    Dim i As Long
    Dim j As Long
    
    Dim iTot() As Long
    Dim iRowSum() As Long
    Dim iCnt  As Integer
    
    With ss1
    
        .ReDraw = False
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        Call GP_SET_CELL_TEXT(ss1, .MaxRows, 0, "合  计")
        Call .AddCellSpan(0, .MaxRows, 2, 1)
        
        .MaxCols = .MaxCols + 1
        Call GP_SET_CELL_TEXT(ss1, 0, .MaxCols, "合  计")
        Call GP_SET_CELL_TEXT(ss1, SpreadHeader + 1, .MaxCols, " ")
        Call GP_SET_CELL_TEXT(ss1, SpreadHeader + 2, .MaxCols, " ")
        Call .AddCellSpan(.MaxCols, 0, 1, 3)
        
        iCnt = .MaxCols
        ReDim iTot(iCnt)
                
        'Col Sum
        For i = 1 To .MaxRows
            For j = 1 To .MaxCols
                 iTot(j) = iTot(j) + GF_GET_CELL_VALUE(ss1, i, j)
            Next j
        Next i
                
        'Col Sum Display
        For i = 1 To .MaxRows
            For j = 2 To .MaxCols Step 2
                Call GP_SET_CELL_VALUE(ss1, .MaxRows, j, iTot(j))
            Next j
        Next i
        
        
'        ReDim iRowSum(.MaxCols)
        ReDim iRowSum(.MaxRows)

        'Row Sum
        For i = 1 To .MaxRows
            For j = 2 To .MaxCols - 1 Step 2
                 iRowSum(i) = iRowSum(i) + GF_GET_CELL_VALUE(ss1, i, j)
            Next j
        Next i

        'Row Sum Display
        For i = 1 To .MaxRows
            For j = 2 To .MaxCols - 1 Step 2
                Call GP_SET_CELL_VALUE(ss1, i, .MaxCols, iRowSum(i))
                .CellType = CellTypeNumber
                .TypeHAlign = TypeHAlignRight
            Next j
        Next i
        
        .ReDraw = True
    End With
    
    Call Gp_Sp_BlockLock(ss1, ss1.MaxCols, ss1.MaxCols, 1, -1, True)
    Call Gp_Sp_BlockLock(ss1, 1, -1, ss1.MaxRows, ss1.MaxRows, True)
    
    Call Gp_Sp_ColColor(ss1, ss1.MaxCols, vbRed)
    Call Gp_Sp_RowColor(ss1, ss1.MaxRows, vbRed)
    
End Sub

Private Sub GP_SET_CELL_VALUE(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long, sText As Variant)
    
    If iRow <= 0 Then Exit Sub
    
    With ss1
        .Row = iRow
        .Col = iCol
        '.Text = Val(sText)
        .Value = Val(sText)
    End With
    
End Sub

Private Function GF_GET_CELL_VALUE(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long) As Variant
    
    If iRow <= 0 Then Exit Function
    
    With ss1
        .Row = iRow
        .Col = iCol
        GF_GET_CELL_VALUE = Val(.Value)
    End With
    
End Function

Private Sub GP_SET_CELL_TEXT(ss1 As vaSpread, ByVal iRow As Long, ByVal iCol As Long, sText As Variant)
    
  '  If iRow <= 0 Then Exit Sub
    
    With ss1
        .Row = iRow
        .Col = iCol
        .Text = sText
    End With
    
End Sub
