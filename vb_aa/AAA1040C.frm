VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1040C 
   Caption         =   "生产计划查询_AAA1040C"
   ClientHeight    =   9870
   ClientLeft      =   795
   ClientTop       =   2550
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   12195
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   8610
      Left            =   90
      TabIndex        =   0
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
      MaxCols         =   12
      MaxRows         =   23
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AAA1040C.frx":0000
   End
   Begin InDate.UDate dtp_yy_mm 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
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
End
Attribute VB_Name = "AAA1040C"
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
    FormType = "Refer"

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
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=Sc1, Key:="Sc"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
   Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 3, 3, BLACK, &HE6E6FF)
   Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 5, 5, BLACK, &HE6E6FF)
   Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 9, 9, BLACK, &HE6E6FF)
   Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 13, 13, BLACK, &HE6E6FF)
   Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 17, 17, BLACK, &HE6E6FF)
   Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, 23, 23, BLACK, &HE6E6FF)

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

    'Call Gp_Sp_ColGet(ss1, "A-System.INI", Me.Name)
    Call Sp_Setting(ss1)
    
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
    
    Call Gp_Sp_ColSet(ss1, "A-System.INI", Me.Name)
    
End Sub

Public Sub Spread_Can()

End Sub

Public Sub Form_Cls()
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    rControl(1).SetFocus
    ss1.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, False

End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
    
        If Sp_Data_Refer() Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), True)
            Call Gp_Sp_ColGet(ss1, "A-System.INI", Me.Name)
        End If
            
    Else
        sMesg = sMesg + " 必须输入"
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
End Sub

Public Sub Form_Pro()

    'If Sp_Process(M_CN1, Proc_Sc("Sc")) Then
    '    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    'End If
    
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
Public Sub Sp_Setting(ByVal sPname As Variant)

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

Public Function Sp_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sDate As String
    
    Dim sYear As String
    Dim sMonth As String
    Dim sQuery As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    Dim dTemp, DCURR As Double
    
    sDate = dtp_yy_mm.RawData
    
    sYear = Mid(sDate, 1, 4)
    sMonth = Mid(sDate, 5, 2)
    
    For iCnt = 1 To 12
        If iCnt > 1 Then
           sMonth = Format(sMonth - 1, "00")
           If sMonth = "00" Then
              sMonth = "12"
              sYear = Format(sYear - 1, "0000")
           End If
        End If
        ss1.Col = iCnt
        ss1.Row = SpreadHeader + 1
        If Mid(sMonth, 1, 1) = "0" Then
           sMonth = Mid(sMonth, 2, 1)
        End If
        ss1.Text = sMonth + "月"
    Next iCnt

'Show Data
    Set adoRs = New ADODB.Recordset
    
    sQuery = "{CALL AAA1040C.P_REFER('" + dtp_yy_mm.RawData + "')}"
    
    'Ado Execute
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    
    With ss1

        Sp_Data_Refer = True
        .ReDraw = False
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
                .Row = iCnt + 1
                For iCol = 1 To .MaxCols
                    .Col = iCol
                         
                    If VarType(ArrayRecords(iCol - 1, iCnt)) = vbNull Or ArrayRecords(iCol - 1, iCnt) = 0 Then
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
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function
