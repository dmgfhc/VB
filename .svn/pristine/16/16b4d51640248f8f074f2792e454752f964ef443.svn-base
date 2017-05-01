VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AHD0140C 
   Caption         =   "厂库别 板库存收发存报表（综判）_AHD0140C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
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
      Left            =   7935
      MaxLength       =   40
      TabIndex        =   6
      Tag             =   "mill_plt"
      Top             =   420
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
      Left            =   7425
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "plt"
      Top             =   420
      Width           =   495
   End
   Begin VB.TextBox text_cur_inv 
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
      Left            =   4530
      TabIndex        =   2
      Top             =   420
      Width           =   1305
   End
   Begin VB.TextBox text_cur_inv_code 
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
      Left            =   4155
      MaxLength       =   2
      TabIndex        =   1
      Top             =   420
      Width           =   345
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8595
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   15030
      _Version        =   393216
      _ExtentX        =   26511
      _ExtentY        =   15161
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
      MaxCols         =   21
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AHD0140C.frx":0000
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   165
      Top             =   405
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
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel14 
      Height          =   300
      Left            =   3180
      Top             =   420
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   529
      Caption         =   "仓库"
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
   Begin InDate.UDate dtp_yy_mm 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Tag             =   "日期"
      Top             =   405
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
   Begin InDate.ULabel ULabel01 
      Height          =   315
      Index           =   14
      Left            =   6030
      Top             =   420
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
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5490
      TabIndex        =   3
      Top             =   15
      Width           =   165
   End
End
Attribute VB_Name = "AHD0140C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Order Management System
'-- Sub_System Name
'-- Program Name
'-- Program ID        ABX1090C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2003.5.19
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2
Dim iSumCol As New Collection       'Sum Column


Private Sub Form_Define()
    Dim sQuery As String
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(dtp_yy_mm, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
Call Gp_Ms_Collection(text_cur_inv_code, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(text_cur_inv, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
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
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AHD0140C.P_REFER", Key:="P-R"

    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, ss1.MaxCols, True)
    
    'Duplicate Count
    iDupCnt = 3
    
    'Sum Column Count
    iSumCnt = 15
    
    'Sum Column Setting
    iSumCol.Add Item:=6
    iSumCol.Add Item:=7
    iSumCol.Add Item:=8
    iSumCol.Add Item:=9
    iSumCol.Add Item:=10
    iSumCol.Add Item:=11
    iSumCol.Add Item:=12
    iSumCol.Add Item:=13
    iSumCol.Add Item:=14
    iSumCol.Add Item:=15
    iSumCol.Add Item:=16
    iSumCol.Add Item:=17
    iSumCol.Add Item:=18
    iSumCol.Add Item:=19
    iSumCol.Add Item:=20
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "B-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
'        Cancel = 1
'        Exit Sub
'    End If
'
'    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Z-System.INI", Me.Name)
    
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
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Set iSumCol = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call subButtonHide
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If
    
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim SMESG As String
    Dim sQuery As String
    sQuery = "{ CALL " + "AHD0140C.P_REFER" + "("
    sQuery = sQuery + "'" + dtp_yy_mm.RawData + "',"
    sQuery = sQuery + " '" + text_cur_inv_code + "','" + txt_plt.Text + "'"
    sQuery = sQuery + ")"
    sQuery = sQuery + "}"

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If dtp_yy_mm.RawData = "" Then
       Call Gp_MsgBoxDisplay("请输入日期", "I")
       Exit Sub
    End If
   
    
    If text_cur_inv_code.Text = "" Then
       Call Gp_MsgBoxDisplay("请输入仓库", "I")
       Exit Sub
    End If

'    If Gf_Sp_Display(M_CN1, ss1, sQuery) Then
    If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, 1, 4, iSumCnt, iSumCol, False) Then

        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call subButtonHide
'        Call Sp_AutoInsertSum
'        Call Sp_AutoInsertSumGroup
    End If
    ss1.ScrollBarExtMode = True
    ss1.ReDraw = True
    

    Exit Sub

Refer_Err:

End Sub
Private Sub subButtonHide()

    MDIMain.MenuTool.Buttons(4).Enabled = False    'Save
    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    
    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
    MDIMain.MenuTool.Buttons(14).Enabled = True     'EXCLE
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
    
'    Call Sp_Excel(Me, ss1, 1, ss1.MaxCols, SpreadHeader, ss1.MaxRows)
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, ss1.MaxCols - 1, lBlkrow1, lBlkrow1)
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

'Public Sub Sp_Excel(Fm As Form, sPname As Variant, bLkcol1 As Long, bLkcol2 As Long, bLkrow1 As Long, bLkrow2 As Long)
'
'On Error GoTo Excel_Error
'
'    Dim ret As Boolean
'    Dim xlApp As Object
'    Dim xlBpp As Object
'    Dim xlBook As Object
'    Dim xlSheet As Object
'
'
'
'    With sPname
'
''        If .MaxRows = 0 Then Exit Sub
'
'        If bLkcol1 = 0 Then
'           bLkcol1 = 1
'        End If
'
'        If bLkcol2 = 0 Then
'            bLkcol2 = -1
'        End If
'
'        If bLkrow2 = 0 Then
'            bLkrow2 = -1
'        End If
'
'        Clipboard.Clear
'
'        .Col = bLkcol1: .Col2 = bLkcol2
'        .Row = bLkrow1: .Row2 = bLkrow2
'        Clipboard.SetText .Clip
'
'        'Call Excel
'        Set xlApp = CreateObject("Excel.Application")
'        xlApp.Workbooks.Open (App.Path & "\AHD0120.xls")
'
'      '  Set xlBook = xlApp.Workbooks.Add
'
''        Set xlSheet = xlBook.Worksheets(1)
'
' Set xlSheet = xlApp.Worksheets("Sheet1")
'    xlApp.Sheets("Sheet1").Select
'
'        xlApp.Visible = True
'
'        xlSheet.cells.NumberFormatLocal = "G/通用格式"
'        xlSheet.Range("A2").Select
'        xlSheet.Range("A2").Clear
'
'        xlSheet.Paste
'        xlSheet.cells.EntireColumn.AutoFit       'Column AutoFit
'
'        Set xlSheet = Nothing
'        Set xlBook = Nothing
'        Set xlApp = Nothing
'
'    End With
'
'    Exit Sub
'
'Excel_Error:
'    Call Gp_MsgBoxDisplay("您的机器尚未安装Excel", "W")
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
'                    .Col = 1: .Text = strProdTag + " 合计"
'                    For iCount = 1 To .MaxCols
'                        .Col = iCount
'                        If .CellType = SS_CELL_TYPE_COMBOBOX Then .Value = 0
'                    Next iCount
'                    Call Gp_Sp_RowColor(ss1, .Row, vbBlue, vbYellow)
''                    Call Gp_Sp_RowColor(ss1, .Row, vbBlue)
'
'                    For x = 6 To 13
'                        dValue = Sp_SumAbove(ss1, x, iRow, irow2)
'                        .Row = iCurRow + 1
'                        .Col = x: .Value = CStr(IIf(dValue > 0, dValue, 0))
'                        irow2 = irow2 + 2
'                    Next x
'                    Exit For
'                End If
'            Next iCurRow
'        Wend
'    End With
'
'End Sub
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
'    Dim dValue106 As Double
'    Dim dValue206 As Double
'    Dim dValue306 As Double
'    Dim dValue107 As Double
'    Dim dValue207 As Double
'    Dim dValue307 As Double
'    Dim dValue108 As Double
'    Dim dValue208 As Double
'    Dim dValue308 As Double
'    Dim dValue109 As Double
'    Dim dValue209 As Double
'    Dim dValue309 As Double
'    Dim dValue110 As Double
'    Dim dValue210 As Double
'    Dim dValue310 As Double
'    Dim dValue111 As Double
'    Dim dValue211 As Double
'    Dim dValue311 As Double
'    Dim dValue112 As Double
'    Dim dValue212 As Double
'    Dim dValue312 As Double
'    Dim dValue113 As Double
'    Dim dValue213 As Double
'    Dim dValue313 As Double
'
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
'
'    iRow = 1
'    bLoop = True
'    dValue106 = 0
'    dValue107 = 0
'    dValue108 = 0
'    dValue109 = 0
'    dValue110 = 0
'    dValue111 = 0
'    dValue112 = 0
'    dValue113 = 0
'
'    dValue206 = 0
'    dValue207 = 0
'    dValue208 = 0
'    dValue209 = 0
'    dValue210 = 0
'    dValue211 = 0
'    dValue212 = 0
'    dValue213 = 0
'
'    dValue306 = 0
'    dValue307 = 0
'    dValue308 = 0
'    dValue309 = 0
'    dValue310 = 0
'    dValue311 = 0
'    dValue312 = 0
'    dValue313 = 0
'
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
'                .Row = iRow
'                .Col = 2: strTag2 = .Text
'                .Col = 3: strTag3 = .Text
'                .Col = 4: strTag4 = .Text
'
'                dValue106 = 0
'                dValue107 = 0
'                dValue108 = 0
'                dValue109 = 0
'                dValue110 = 0
'                dValue111 = 0
'                dValue112 = 0
'                dValue113 = 0
'
'                For curRow = iRow To .MaxRows
'                    .Row = curRow:
'                    .Col = 2: strTag21 = .Text
'                    .Col = 3: strTag31 = .Text
'                    .Col = 4: strTag41 = .Text
'                    .Col = 2
'                    If .Text <> "" And strTag2 = strTag21 Then
'                        If strTag3 = strTag31 Then
'                            If strTag4 = strTag41 Then
'                                 dValue106 = dValue106 + Sp_SumGroup(ss1, curRow, 6)
'                                 dValue107 = dValue107 + Sp_SumGroup(ss1, curRow, 7)
'                                 dValue108 = dValue108 + Sp_SumGroup(ss1, curRow, 8)
'                                 dValue109 = dValue109 + Sp_SumGroup(ss1, curRow, 9)
'                                 dValue110 = dValue110 + Sp_SumGroup(ss1, curRow, 10)
'                                 dValue111 = dValue111 + Sp_SumGroup(ss1, curRow, 11)
'                                 dValue112 = dValue112 + Sp_SumGroup(ss1, curRow, 12)
'                                 dValue113 = dValue113 + Sp_SumGroup(ss1, curRow, 13)
'                            Else
'                                 dValue206 = dValue206 + dValue106
'                                 dValue207 = dValue207 + dValue107
'                                 dValue208 = dValue208 + dValue108
'                                 dValue209 = dValue209 + dValue109
'                                 dValue210 = dValue210 + dValue110
'                                 dValue211 = dValue211 + dValue111
'                                 dValue212 = dValue212 + dValue112
'                                 dValue213 = dValue213 + dValue113
'
'                                 dValue306 = dValue306 + dValue106
'                                 dValue307 = dValue307 + dValue107
'                                 dValue308 = dValue308 + dValue108
'                                 dValue309 = dValue309 + dValue109
'                                 dValue310 = dValue310 + dValue110
'                                 dValue311 = dValue311 + dValue111
'                                 dValue312 = dValue312 + dValue112
'                                 dValue313 = dValue313 + dValue113
'
'                                 .MaxRows = .MaxRows + 1
'                                 .Row = curRow
'                                 .Action = SS_ACTION_INSERT_ROW
'                                 iRow = .Row + 1
'                                 .Col = 0: .Text = "∑"
'                                 .Col = 4: .Text = strTag4 + " 小计"
'
'                                 Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'
'                                 .Col = 6: .Value = IIf(dValue106 > 0, dValue106, 0)
'                                 .Col = 7: .Value = IIf(dValue107 > 0, dValue107, 0)
'                                 .Col = 8: .Value = IIf(dValue108 > 0, dValue108, 0)
'                                 .Col = 9: .Value = IIf(dValue109 > 0, dValue109, 0)
'                                 .Col = 10: .Value = IIf(dValue110 > 0, dValue110, 0)
'                                 .Col = 11: .Value = IIf(dValue111 > 0, dValue111, 0)
'                                 .Col = 12: .Value = IIf(dValue112 > 0, dValue112, 0)
'                                 .Col = 13: .Value = IIf(dValue113 > 0, dValue113, 0)
'                                Exit For
'                            End If
'                        Else
'                            dValue206 = dValue206 + dValue106
'                            dValue207 = dValue207 + dValue107
'                            dValue208 = dValue208 + dValue108
'                            dValue209 = dValue209 + dValue109
'                            dValue210 = dValue210 + dValue110
'                            dValue211 = dValue211 + dValue111
'                            dValue212 = dValue212 + dValue112
'                            dValue213 = dValue213 + dValue113
'
'                            dValue306 = dValue306 + dValue106
'                            dValue307 = dValue307 + dValue107
'                            dValue308 = dValue308 + dValue108
'                            dValue309 = dValue309 + dValue109
'                            dValue310 = dValue310 + dValue110
'                            dValue311 = dValue311 + dValue111
'                            dValue312 = dValue312 + dValue112
'                            dValue313 = dValue313 + dValue113
'
'                             .MaxRows = .MaxRows + 1
'                             .Row = curRow
'                             .Action = SS_ACTION_INSERT_ROW
'                             iRow = .Row + 1
'                             .Col = 0: .Text = "∑"
'                             .Col = 4: .Text = strTag4 + " 小计"
'
'                             Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'
'                             .Col = 6: .Value = IIf(dValue106 > 0, dValue106, 0)
'                             .Col = 7: .Value = IIf(dValue107 > 0, dValue107, 0)
'                             .Col = 8: .Value = IIf(dValue108 > 0, dValue108, 0)
'                             .Col = 9: .Value = IIf(dValue109 > 0, dValue109, 0)
'                             .Col = 10: .Value = IIf(dValue110 > 0, dValue110, 0)
'                             .Col = 11: .Value = IIf(dValue111 > 0, dValue111, 0)
'                             .Col = 12: .Value = IIf(dValue112 > 0, dValue112, 0)
'                             .Col = 13: .Value = IIf(dValue113 > 0, dValue113, 0)
'
'                             .MaxRows = .MaxRows + 1
'                             .Row = iRow
'                             .Action = SS_ACTION_INSERT_ROW
'                             iRow = .Row + 1
'                             .Col = 0: .Text = "∑"
'                             .Col = 3: .Text = strTag3 + " 小计"
'
'                             Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'
'                             .Col = 6: .Value = IIf(dValue206 > 0, dValue206, 0)
'                             .Col = 7: .Value = IIf(dValue207 > 0, dValue207, 0)
'                             .Col = 8: .Value = IIf(dValue208 > 0, dValue208, 0)
'                             .Col = 9: .Value = IIf(dValue209 > 0, dValue209, 0)
'                             .Col = 10: .Value = IIf(dValue210 > 0, dValue210, 0)
'                             .Col = 11: .Value = IIf(dValue211 > 0, dValue211, 0)
'                             .Col = 12: .Value = IIf(dValue212 > 0, dValue212, 0)
'                             .Col = 13: .Value = IIf(dValue213 > 0, dValue213, 0)
'
'                            dValue106 = 0
'                            dValue107 = 0
'                            dValue108 = 0
'                            dValue109 = 0
'                            dValue110 = 0
'                            dValue111 = 0
'                            dValue112 = 0
'                            dValue113 = 0
'
'                            dValue206 = 0
'                            dValue207 = 0
'                            dValue208 = 0
'                            dValue209 = 0
'                            dValue210 = 0
'                            dValue211 = 0
'                            dValue212 = 0
'                            dValue213 = 0
'
'                            Exit For
'
'                        End If
'                    Else
'                         dValue206 = dValue206 + dValue106
'                         dValue207 = dValue207 + dValue107
'                         dValue208 = dValue208 + dValue108
'                         dValue209 = dValue209 + dValue109
'                         dValue210 = dValue210 + dValue110
'                         dValue211 = dValue211 + dValue111
'                         dValue212 = dValue212 + dValue112
'                         dValue213 = dValue213 + dValue113
'
'                         dValue306 = dValue306 + dValue106
'                         dValue307 = dValue307 + dValue107
'                         dValue308 = dValue308 + dValue108
'                         dValue309 = dValue309 + dValue109
'                         dValue310 = dValue310 + dValue110
'                         dValue311 = dValue311 + dValue111
'                         dValue312 = dValue312 + dValue112
'                         dValue313 = dValue313 + dValue113
'
'                         .MaxRows = .MaxRows + 1
'                         .Row = curRow
'                         .Action = SS_ACTION_INSERT_ROW
'                         iRow = .Row + 1
'                         .Col = 0: .Text = "∑"
'                         .Col = 4: .Text = strTag4 + " 小计"
'
'                         Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'                         .Col = 6: .Value = IIf(dValue106 > 0, dValue106, 0)
'                         .Col = 7: .Value = IIf(dValue107 > 0, dValue107, 0)
'                         .Col = 8: .Value = IIf(dValue108 > 0, dValue108, 0)
'                         .Col = 9: .Value = IIf(dValue109 > 0, dValue109, 0)
'                         .Col = 10: .Value = IIf(dValue110 > 0, dValue110, 0)
'                         .Col = 11: .Value = IIf(dValue111 > 0, dValue111, 0)
'                         .Col = 12: .Value = IIf(dValue112 > 0, dValue112, 0)
'                         .Col = 13: .Value = IIf(dValue113 > 0, dValue113, 0)
'
'                         .MaxRows = .MaxRows + 1
'                         .Row = iRow
'                         .Action = SS_ACTION_INSERT_ROW
'                         iRow = .Row + 1
'                         .Col = 0: .Text = "∑"
'                         .Col = 3: .Text = strTag3 + " 小计"
'
'                         Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'
'                         .Col = 6: .Value = IIf(dValue206 > 0, dValue206, 0)
'                         .Col = 7: .Value = IIf(dValue207 > 0, dValue207, 0)
'                         .Col = 8: .Value = IIf(dValue208 > 0, dValue208, 0)
'                         .Col = 9: .Value = IIf(dValue209 > 0, dValue209, 0)
'                         .Col = 10: .Value = IIf(dValue210 > 0, dValue210, 0)
'                         .Col = 11: .Value = IIf(dValue211 > 0, dValue211, 0)
'                         .Col = 12: .Value = IIf(dValue212 > 0, dValue212, 0)
'                         .Col = 13: .Value = IIf(dValue213 > 0, dValue213, 0)
'
'                         .MaxRows = .MaxRows + 1
'                         .Row = iRow
'                         .Action = SS_ACTION_INSERT_ROW
'                         iRow = .Row + 1
'                         .Col = 0: .Text = "∑"
'                         .Col = 2: .Text = strTag2 + " 小计"
'
'                         Call Gp_Sp_RowColor(ss1, .Row, vbRed, vbYellow)
'
'                         .Col = 6: .Value = IIf(dValue306 > 0, dValue306, 0)
'                         .Col = 7: .Value = IIf(dValue307 > 0, dValue307, 0)
'                         .Col = 8: .Value = IIf(dValue308 > 0, dValue308, 0)
'                         .Col = 9: .Value = IIf(dValue309 > 0, dValue309, 0)
'                         .Col = 10: .Value = IIf(dValue310 > 0, dValue310, 0)
'                         .Col = 11: .Value = IIf(dValue311 > 0, dValue311, 0)
'                         .Col = 12: .Value = IIf(dValue312 > 0, dValue312, 0)
'                         .Col = 13: .Value = IIf(dValue313 > 0, dValue313, 0)
'
'                         dValue106 = 0
'                         dValue107 = 0
'                         dValue108 = 0
'                         dValue109 = 0
'                         dValue110 = 0
'                         dValue111 = 0
'                         dValue112 = 0
'                         dValue113 = 0
'
'                         dValue206 = 0
'                         dValue207 = 0
'                         dValue208 = 0
'                         dValue209 = 0
'                         dValue210 = 0
'                         dValue211 = 0
'                         dValue212 = 0
'                         dValue213 = 0
'
'                         dValue306 = 0
'                         dValue307 = 0
'                         dValue308 = 0
'                         dValue309 = 0
'                         dValue310 = 0
'                         dValue311 = 0
'                         dValue312 = 0
'                         dValue313 = 0
'
'                        Exit For
'                    End If
'                Next curRow
'            'iRow = curRow
'            'dValue2 = dValue2 + dValue
'       Wend
'    End With
'
'End Sub

Private Sub text_cur_inv_code_Change()
    If Len(Trim(text_cur_inv_code.Text)) = text_cur_inv_code.MaxLength Then
        text_cur_inv.Text = Gf_ComnNameFind(M_CN1, "C0013", text_cur_inv_code.Text, 2)
        Label2.Caption = text_cur_inv.Text + "  板/卷 库存收发存报表"
    Else
      text_cur_inv.Text = ""
    End If
End Sub

Private Sub text_cur_inv_code_DblClick()
   Call text_cur_inv_code_KeyUp(vbKeyF4, 0)

End Sub

Private Sub text_cur_inv_code_KeyUp(KeyCode As Integer, Shift As Integer)

     If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0013"

        DD.rControl.Add Item:=text_cur_inv_code
        DD.rControl.Add Item:=text_cur_inv
        

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    End If
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



