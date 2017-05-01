VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1240C 
   Caption         =   "编制板卷炼钢计划_AAA1240C"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
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
      Top             =   240
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
      Top             =   240
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
   Begin Threed.SSCommand MSP_plan_cmd 
      Height          =   330
      Left            =   13590
      TabIndex        =   2
      Top             =   240
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
      Caption         =   "编制精炼计划"
   End
   Begin Threed.SSCommand BOF_plan_cmd 
      Height          =   330
      Left            =   11850
      TabIndex        =   3
      Top             =   240
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
   Begin FPSpread.vaSpread ss2 
      Height          =   2790
      Left            =   120
      TabIndex        =   4
      Top             =   3465
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   4921
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
      SpreadDesigner  =   "AAA1240C.frx":0000
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   2640
      Left            =   120
      TabIndex        =   5
      Top             =   705
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   4657
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
      SpreadDesigner  =   "AAA1240C.frx":075C
   End
   Begin FPSpread.vaSpread ss3 
      Height          =   2760
      Left            =   120
      TabIndex        =   6
      Top             =   6375
      Width           =   15015
      _Version        =   393216
      _ExtentX        =   26485
      _ExtentY        =   4868
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
      MaxRows         =   3
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AAA1240C.frx":0EBA
   End
End
Attribute VB_Name = "AAA1240C"
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
'-- Program ID        AAA1240C
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

Dim Mc1 As New Collection           'Master Collection

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
   
   Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 4, 4, BLACK, &HE6E6FF)
   Call Gp_Sp_BlockColor(ss2, 1, ss2.MaxCols, 4, 4, BLACK, &HE6E6FF)
   Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, 3, 3, BLACK, &HE6E6FF)
   
End Sub

Private Sub BOF_plan_cmd_Click()
On Error GoTo bof_plan_cmd_Error

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

bof_plan_cmd_Error:

    Call Gp_MsgBoxDisplay("编制计划错误 : " & Error)
End Sub

Private Sub MSP_plan_cmd_Click()

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
    For iCount = 1 To 2
        adoCmd.Parameters.Append adoCmd.CreateParameter(Str(iCount), adVariant, adParamOutput)
    Next iCount
    
    'CAST
    sQuery = "{call AAA4040P ('" + dtp_yy_mm.RawData + "',?,?)}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords
    
    If adoCmd(0) <> "YY" Then
        Call Gp_MsgBoxDisplay(adoCmd(1))
        Set adoCmd = Nothing
        Exit Sub
    Else
        MsgBox "精炼计划编制完了!"
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
    Call Gp_Sp_ColGet(ss3, "A-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    If Mid(sAuthority, 1, 3) = "111" Then
       BOF_plan_cmd.Enabled = True
       MSP_plan_cmd.Enabled = True
    Else
       BOF_plan_cmd.Enabled = False
       MSP_plan_cmd.Enabled = False
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
    Call Gp_Sp_ColSet(ss3, "A-System.INI", Me.Name)
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    
    dtp_yy_mm.SetFocus
    
    ss1.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, False
    ss2.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, False
    ss3.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, False
End Sub

Public Sub Form_Ref()
    If Sp_Refer(ss1) Then
       Call Sp_Refer(ss2)
       Call Sp_Refer(ss3)
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       txt_excel = "1"
       Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows, True)
       Call Gp_Sp_BlockLock(ss2, 1, ss2.MaxCols, 1, ss2.MaxRows, True)
       Call Gp_Sp_BlockLock(ss3, 1, ss3.MaxCols, 1, ss3.MaxRows, True)
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
        
        '.LockBackColor = RGB(255, 255, 255)
                
    End With
    
End Sub

Public Function Sp_Refer(ByVal sPname As Variant) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    If sPname Is ss1 Then
       sQuery = "{CALL AAA1240C.P_SREFER1('" + dtp_yy_mm.RawData + "')}"
    ElseIf sPname Is ss2 Then
       sQuery = "{CALL AAA1240C.P_SREFER2('" + dtp_yy_mm.RawData + "')}"
    ElseIf sPname Is ss3 Then
       sQuery = "{CALL AAA1240C.P_SREFER3('" + dtp_yy_mm.RawData + "')}"
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
                    .Text = Trim(ArrayRecords(iCol - 1, iCnt))
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

