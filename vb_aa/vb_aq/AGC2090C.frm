VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2090C 
   Caption         =   "轧钢生产线停机实绩查询及修改界面-AGC2090C"
   ClientHeight    =   7305
   ClientLeft      =   570
   ClientTop       =   1890
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   14160
   WindowState     =   2  '???
   Begin VB.ComboBox CBO_PLT 
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
      ItemData        =   "AGC2090C.frx":0000
      Left            =   1200
      List            =   "AGC2090C.frx":000A
      TabIndex        =   0
      Top             =   90
      Width           =   735
   End
   Begin VB.ComboBox CBO_PRC 
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
      ItemData        =   "AGC2090C.frx":0016
      Left            =   3540
      List            =   "AGC2090C.frx":003B
      TabIndex        =   1
      Top             =   90
      Width           =   735
   End
   Begin VB.TextBox TXT_DEL_RES_CD 
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
      Left            =   8865
      TabIndex        =   3
      Top             =   90
      Width           =   825
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8655
      Left            =   90
      TabIndex        =   4
      Top             =   495
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   15266
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
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AGC2090C.frx":006B
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   90
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "工厂代码"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   2430
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "工序代码"
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
   Begin CSTextLibCtl.sitxEdit TXT_OCCR_TIME 
      Height          =   315
      Left            =   5895
      TabIndex        =   2
      Top             =   90
      Width           =   1245
      _Version        =   262145
      _ExtentX        =   2196
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   "____-__-__ __-__-__"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      Modified        =   -1  'True
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   "____-__-__ "
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   "____-__-__ "
      CharacterTable  =   ""
      BorderStyle     =   0
      MaxLength       =   0
      ValidateMask    =   0   'False
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   4770
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "发生时间"
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
   Begin InDate.ULabel ULabel10 
      Height          =   315
      Left            =   7740
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "停机代码"
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
End
Attribute VB_Name = "AGC2090C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name        System
'-- Sub_System Name
'-- Program Name      AGC2090C
'-- Program ID        AGC2090C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang Meng
'-- Coder             Yang Meng
'-- Date              2003.7.24
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
Dim sDateTime_str As String         'Active Form Time Setting
Dim sDateTime_end As String         'Active Form Time Setting
Dim sDateTime_cnt As Double         'Active Form Time Setting

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
         Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(CBO_PRC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(TXT_OCCR_TIME, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(TXT_DEL_RES_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2090C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AGC2090C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGC2090C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
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

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 2)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 4)
    
    CBO_PLT.ListIndex = 0

    Screen.MousePointer = vbDefault

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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

Public Sub form_cls()

    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        pControl(1).SetFocus
        CBO_PLT.Text = "C1"
    End If

End Sub

Public Sub form_ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    Dim iCount As Integer

    For iCount = 1 To ss1.MaxRows

        Select Case Trim(Gf_Sp_RcvData(ss1, 0, iCount))

            Case "Update" ', "Input"

                  With ss1
                      .Col = 6
                      If Val(Mid(.Text, 1, 4)) < 2000 Or Val(Mid(.Text, 1, 4)) > 2050 Or Val(Mid(.Text, 6, 2)) > 12 Or Val(Mid(.Text, 9, 2)) > 31 Or _
                         Val(Mid(.Text, 12, 2)) > 24 Or Val(Mid(.Text, 15, 2)) > 60 Then
                         Call Gp_MsgBoxDisplay("请正确输入日期时间")
                         Exit Sub
                      End If
                      sDateTime_str = Mid(.Text, 1, 4) & Mid(.Text, 6, 2) & Mid(.Text, 9, 2) & Mid(.Text, 12, 2) & Mid(.Text, 15, 2) & Mid(.Text, 18, 2)
                      .Col = 7
                      If Val(Mid(.Text, 1, 4)) < 2000 Or Val(Mid(.Text, 1, 4)) > 2050 Or Val(Mid(.Text, 6, 2)) > 12 Or Val(Mid(.Text, 9, 2)) > 31 Or _
                         Val(Mid(.Text, 12, 2)) > 24 Or Val(Mid(.Text, 15, 2)) > 60 Then
                         Call Gp_MsgBoxDisplay("请正确输入日期时间")
                         Exit Sub
                      End If
                      sDateTime_end = Mid(.Text, 1, 4) & Mid(.Text, 6, 2) & Mid(.Text, 9, 2) & Mid(.Text, 12, 2) & Mid(.Text, 15, 2) & Mid(.Text, 18, 2)
                      If Val(sDateTime_end) - Val(sDateTime_str) < 0 Then
                         Call Gp_MsgBoxDisplay("结束时间应大于开始时间")
                         Exit Sub
                      End If
                  End With

        End Select

    Next iCount

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

End Sub

Public Sub Form_Ins()

    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)

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

Public Sub form_exit()
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

' Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
'    ss1.Row = Row
'    ss1.Col = Col
'
'    If ss1.Col = 7 Then
'           ss1.Col = 5
'           sDateTime_str = Mid(ss1.Text, 6, 2) & "/" & Mid(ss1.Text, 9, 2) & "/" & Mid(ss1.Text, 1, 4) & " " & Mid(ss1.Text, 12, 2) & ":" & Mid(ss1.Text, 15, 2) & ":" & Mid(ss1.Text, 18, 2)
'          'sDateTime_str = "#" & Mid(ss1.Text, 6, 2) & "/" & Mid(ss1.Text, 9, 2) & "/" & Mid(ss1.Text, 1, 4) & " " & Mid(ss1.Text, 12, 2) & ":" & Mid(ss1.Text, 15, 2) & ":" & Mid(ss1.Text, 18, 2) & "#"
'           ss1.Col = 6
'           sDateTime_end = Mid(ss1.Text, 6, 2) & "/" & Mid(ss1.Text, 9, 2) & "/" & Mid(ss1.Text, 1, 4) & " " & Mid(ss1.Text, 12, 2) & ":" & Mid(ss1.Text, 15, 2) & ":" & Mid(ss1.Text, 18, 2)
'           sDateTime_cnt = Round(Mid((CDate(sDateTime_end) - CDate(sDateTime_str)) * 1440, 1, 4), 0)
'               'sDateTime_cnt = DateDiff("n", (CDate(sDateTime_end) - CDate(sDateTime_str)))
'               'sDateTime_cnt = Round(Mid((CDate(Mid(sDateTime_end, 2, 19)) - CDate(Mid(sDateTime_str, 2, 19))) * 1440, 1, 4), 0)
'               ss1.Col = 7
'                   ss1.Text = sDateTime_cnt
'                   ss1.Col = 0
'                   Select Case Trim(ss1.Text)
'                          Case "Input", "Update", "Delete"
'                          Case Else
'                          ss1.Text = "Update"
'                   End Select
'    End If

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If Row = 0 Then Exit Sub
    ss1.Row = Row
    ss1.Col = Col
     
    If ss1.Lock = False Then
        If ss1.Col = 3 Then

         ss1.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
         
         ss1.Col = 0
         Select Case Trim(ss1.Text)
                Case "Input", "Update", "Delete"
                Case Else
                ss1.Text = "Update"
         End Select
       End If
    End If

    If ss1.Col = 6 Then
         
        ss1.Text = Format(Now, "YYYY-MM-DD HH:MM")
        
        ss1.Col = 0
        Select Case Trim(ss1.Text)
               Case "Input", "Update", "Delete"
               Case Else
                    ss1.Text = "Update"
        End Select
    End If
    
    If ss1.Col = 7 Then

        ss1.Text = Format(Now, "YYYY-MM-DD HH:MM")
        
        ss1.Col = 0
        Select Case Trim(ss1.Text)
               Case "Input", "Update", "Delete"
               Case Else
                    ss1.Text = "Update"
        End Select
    End If
    ss1.Col = 1
    ss1.Text = CBO_PLT.Text
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
       ' Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If

End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub

    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub

    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
      '  Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
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
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub TXT_DEL_RES_CD_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "G0013"
        DD.rControl.Add Item:=TXT_DEL_RES_CD
       
        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub

Private Sub TXT_OCCR_TIME_DblClick()

    TXT_OCCR_TIME.RawData = Format(Now, "YYYYMMDD")

End Sub

Private Sub SS1_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If ss1.Col = 4 Then
          
         If KeyCode = vbKeyF4 Then
              
            Set DD.sPname = Me.ss1
                  
            DD.sWitch = "SP"
            DD.sKey = "G0013"
            DD.rControl.Add Item:=4
            DD.rControl.Add Item:=5
                  
            DD.nameType = "1"
                  
            Call Gf_Common_DD(M_CN1, KeyCode)
              
        End If
        
    End If

    If ss1.Col = 2 Then
          
         If KeyCode = vbKeyF4 Then
              
            Set DD.sPname = Me.ss1
                  
            DD.sWitch = "SP"
            DD.sKey = "C0002"
            DD.rControl.Add Item:=2
            
                  
            DD.nameType = "1"
                  
            Call Gf_Common_DD(M_CN1, KeyCode)
              
        End If
        
    End If

End Sub

