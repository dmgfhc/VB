VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKP3070C 
   Caption         =   "能源消耗日报表_AKP3070C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   180
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
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
   Begin InDate.UDate txt_DATE 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Tag             =   "日期"
      Top             =   120
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
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8610
      Left            =   90
      TabIndex        =   1
      Top             =   510
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   15187
      _StockProps     =   64
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
      MaxCols         =   12
      MaxRows         =   16
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   2
      SpreadDesigner  =   "AKP3070C.frx":0000
   End
   Begin Threed.SSCommand Cmd_Edit 
      Height          =   360
      Left            =   8910
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "更新数据"
   End
End
Attribute VB_Name = "AKP3070C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      中厚板卷厂能源消耗日报表
'-- Program ID        AKP3070C
'-- Document No       Q-00-0010(Specification)
'-- Designer          WANGYU
'-- Coder             WANGYU
'-- Date              2009.6.30
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
Public sDateTime As String          'Active Form Time Setting

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
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_DATE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, "P", " ", " ", " ", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    'sc1.Add Item:="AKP3051C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AKP3070C.P_SREFER", Key:="P-R"
    'sc1.Add Item:="AKP3051C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Call Gp_Sp_ColHidden(ss1, 5, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub



Private Sub Cmd_Edit_Click()
    Dim adoCmd              As ADODB.Command
    Dim Response            As Variant

    On Error GoTo Process_Exec_ERROR
    
    Response = MsgBox("重新生成" + Mid(txt_DATE.RawData, 1, 4) + "年" + Mid(txt_DATE.RawData, 5, 2) + "月" + Mid(txt_DATE.RawData, 7, 2) + "日  " + "的报表吗?", vbYesNo, "系统提示信息")
    If Response = vbNo Then
        Exit Sub
    End If
             
    Screen.MousePointer = vbHourglass
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = "{call AKP3070C.p_ini_data('" + txt_DATE.RawData + "')}"

    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    Call MsgBox("报表已重新生成！", vbInformation, "系统提示信息")
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Call Form_Ref
    Exit Sub
    
Process_Exec_ERROR:
    
    Set adoCmd = Nothing
    Call Gp_MsgBoxDisplay(Err.Description & "{call AKP3070C.p_ini_data('" + txt_DATE.RawData + "')}")
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    With MDIMain.MenuTool
        .Buttons(11).Enabled = False                'Spread Copy
        .Buttons(12).Enabled = False                'Paste
    End With

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

    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Sp_Setting(Proc_Sc("Sc")("Spread"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
       Cmd_Edit.Enabled = True
    End If

    Screen.MousePointer = vbDefault
    txt_DATE.RawData = Format(Date - 1, "yyyymmdd")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)

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


Public Sub Form_Cls()

    Dim iRow  As Long
    Dim iCol  As Long
    ss1.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, True
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
End Sub

Public Sub Form_Ref()

'On Error GoTo Refer_Err
'
'    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
'
'    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
'        ss1.OperationMode = OperationModeNormal
'        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'    With MDIMain.MenuTool
'        .Buttons(11).Enabled = False                'Spread Copy
'        .Buttons(12).Enabled = False                'Paste
'    End With
'    End If
'
'    Exit Sub
'
'Refer_Err:
    ss1.ReDraw = False
    Call Form_Cls
    Screen.MousePointer = vbHourglass
        
    Call ENERGY_Sp_Display(M_CN1, ss1, Gf_Ms_MakeQuery(sc1.Item("P-R"), "R", Mc1("pControl")))
           
    ss1.Col = 2
    ss1.Row = 13
    ss1.Text = "--"
    ss1.Row = 14
    ss1.Text = "--"
    ss1.Row = 15
    ss1.Text = "--"
    ss1.Row = 16
    ss1.Text = "--"
    
    ss1.Col = 3
    ss1.Row = 7
    ss1.Text = "--"
    ss1.Row = 8
    ss1.Text = "--"
    ss1.Row = 15
    ss1.Text = "--"
    ss1.Row = 16
    ss1.Text = "--"
    
    ss1.Col = 4
    ss1.Row = 9
    ss1.Text = "--"
    ss1.Row = 10
    ss1.Text = "--"
    ss1.Row = 11
    ss1.Text = "--"
    ss1.Row = 12
    ss1.Text = "--"
    ss1.Row = 13
    ss1.Text = "--"
    ss1.Row = 14
    ss1.Text = "--"
    ss1.Row = 15
    ss1.Text = "--"
    ss1.Row = 16
    ss1.Text = "--"
    
    ss1.Col = 6
    ss1.Row = 15
    ss1.Text = "--"
    ss1.Row = 16
    ss1.Text = "--"
    
    ss1.Col = 7
    ss1.Row = 15
    ss1.Text = "--"
    ss1.Row = 16
    ss1.Text = "--"
    
    ss1.Col = 8
    ss1.Row = 7
    ss1.Text = "--"
    ss1.Row = 8
    ss1.Text = "--"
    ss1.Row = 13
    ss1.Text = "--"
    ss1.Row = 14
    ss1.Text = "--"
    ss1.Row = 15
    ss1.Text = "--"
    ss1.Row = 16
    ss1.Text = "--"
    
    ss1.Col = 9
    ss1.Row = 5
    ss1.Text = "--"
    ss1.Row = 6
    ss1.Text = "--"
    ss1.Row = 7
    ss1.Text = "--"
    ss1.Row = 8
    ss1.Text = "--"
    ss1.Row = 9
    ss1.Text = "--"
    ss1.Row = 10
    ss1.Text = "--"
    ss1.Row = 11
    ss1.Text = "--"
    ss1.Row = 12
    ss1.Text = "--"
    ss1.Row = 13
    ss1.Text = "--"
    ss1.Row = 14
    ss1.Text = "--"
    ss1.Row = 15
    ss1.Text = "--"
    ss1.Row = 16
    ss1.Text = "--"
    
    ss1.Col = 10
    ss1.Row = 5
    ss1.Text = "--"
    ss1.Row = 6
    ss1.Text = "--"
    ss1.Row = 7
    ss1.Text = "--"
    ss1.Row = 8
    ss1.Text = "--"
    ss1.Row = 9
    ss1.Text = "--"
    ss1.Row = 10
    ss1.Text = "--"
    ss1.Row = 11
    ss1.Text = "--"
    ss1.Row = 12
    ss1.Text = "--"
    ss1.Row = 13
    ss1.Text = "--"
    ss1.Row = 14
    ss1.Text = "--"
    ss1.Row = 15
    ss1.Text = "--"
    ss1.Row = 16
    ss1.Text = "--"
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    ss1.ReDraw = True
     With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    
    Screen.MousePointer = vbDefault

End Sub

Public Sub Form_Pro()

End Sub

Public Sub Form_Ins()

End Sub

Public Sub Spread_Cpy()

'    Call Gp_Sp_Copy(Proc_Sc("Sc"))
'
End Sub

Public Sub Spread_Pst()

'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_ColumnsSort()

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

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Public Sub Sp_Setting(ByVal sPname As Variant, Optional MsgChk As Boolean = True)
    With sPname
    
        .RowHeight(-1) = 12.54
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 12
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 12
        Else
            .RowHeight(0) = 24
        End If
        
        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
' 115,80,195
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
        
        
        If MsgChk Then
            .LockBackColor = RGB(255, 255, 255)
        End If

    End With
    
End Sub



Public Function ENERGY_Sp_Display(Conn As ADODB.Connection, sPname As Variant, sQuery As String) As Boolean

    On Error Resume Next

    Dim icount          As Integer
    Dim iRowCount       As Long
    Dim iColcount       As Long
    Dim AdoRs           As ADODB.Recordset
    Dim ArrayRecords    As Variant

    ENERGY_Sp_Display = True

    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then ENERGY_Sp_Display = False: Exit Function
    End If

    Set AdoRs = New ADODB.Recordset

    With sPname

        .ReDraw = False
        icount = 0

'        .ClearRange 1, 1, .MaxCols, .MaxRows, True

        Screen.MousePointer = vbHourglass

        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset

        If AdoRs.BOF Or AdoRs.EOF Then

            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            ENERGY_Sp_Display = False
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Call Form_Cls
            Screen.MousePointer = vbDefault
            Exit Function

        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) <> 0 Then

            For iRowCount = 0 To .MaxRows - 1
            
                .Row = iRowCount + 1

                For iColcount = 1 To .MaxCols
    
                    .Col = iColcount
    
                    If VarType(ArrayRecords(iColcount - 1, iRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iColcount - 1, iRowCount))
                    End If

                Next iColcount

            Next iRowCount

        End If

        .ReDraw = True
        Screen.MousePointer = vbDefault

    End With

End Function
