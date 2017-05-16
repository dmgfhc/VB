VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0430C 
   Caption         =   "热处理方法及条件查询_AQA0430C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_THK_MIN 
      Height          =   315
      Left            =   7860
      TabIndex        =   4
      Tag             =   "厚度组-最小"
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txt_THK_MAX 
      Height          =   315
      Left            =   8670
      TabIndex        =   3
      Tag             =   "厚度组-最大"
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmd_ListView 
      Caption         =   "<"
      Height          =   315
      Left            =   6975
      TabIndex        =   1
      Top             =   30
      Width           =   435
   End
   Begin VB.TextBox txt_STDSPEC 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Tag             =   "标准号"
      Top             =   30
      Width           =   1755
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   0
      Top             =   45
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "标准编号"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   3900
      Top             =   30
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "厚度组"
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
   Begin FPSpread.vaSpread ss2 
      Height          =   270
      Left            =   4950
      TabIndex        =   2
      Top             =   30
      Width           =   1965
      _Version        =   393216
      _ExtentX        =   3466
      _ExtentY        =   476
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      MaxCols         =   2
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "AQA0430C.frx":0000
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8685
      Left            =   0
      TabIndex        =   5
      Top             =   540
      Width           =   15240
      _Version        =   393216
      _ExtentX        =   26882
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
      MaxCols         =   20
      MaxRows         =   5
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0430C.frx":03E4
   End
End
Attribute VB_Name = "AQA0430C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   质量标准管理
'-- Program Name      录入热处理方法及条件
'-- Program ID        AQA0430C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Sun Bin
'-- Coder             Sun Bin
'-- Date              2007.8.29
'-- Description       录入热处理方法及条件
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

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_THK_MIN, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_THK_MAX, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0430C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQA0430C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:="AQA0430C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String

    Select Case Me.ActiveControl.Name

        Case "txt_STDSPEC"
            sCode = "STDSPEC"

    End Select

    If sCode = "" Then Exit Sub

    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)

    Set oCodeName = Nothing
Err_Track:
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

    sAuthority = Gf_Pgm_Authority(Me.Name, True)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)

    Screen.MousePointer = vbDefault

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
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
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
        ss2.MaxRows = 1
        ss2.Height = 255
        Call GP_SET_CELL_VALUE(ss2, 1, 1, "")
        Call GP_SET_CELL_VALUE(ss2, 1, 2, "")
        txt_THK_MIN.Text = ""
        txt_THK_MAX.Text = ""
    End If
 
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
       If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Exit Sub
        End If
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

         If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
         End If
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 15)

End Sub
Private Sub cmd_ListView_Click()
Dim sQuery As String

    sQuery = "Select Distinct THK_MIN,THK_MAX From QP_HEAT_STD Where STDSPEC = "
            
        With ss2
        
            .MaxRows = 1
            .Height = 255
    
        End With

        
        If txt_STDSPEC.Text = "" Or Trim(txt_STDSPEC.Text) = "" Then
            Exit Sub
        End If
        sQuery = sQuery + " '" + txt_STDSPEC.Text + "'"
        
        Call GS_Combo_SS_ADD(sQuery, ss2)

    
    If Gf_GetCellNullCheck(ss2, 1, 1) <> "" And Gf_GetCellNullCheck(ss2, 1, 2) <> "" Then
            txt_THK_MIN.Text = Gf_GetCellNullCheck(ss2, 1, 1)
            txt_THK_MAX.Text = Gf_GetCellNullCheck(ss2, 1, 2)
    End If
    
      ss2.ZOrder
End Sub
Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 15)
    
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
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

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
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)
        
    End If
    
End Sub


Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim sTemp_Code As String
    Dim iCol As Long
    Dim iRow As Long

    iCol = ss1.ActiveCol
    iRow = ss1.ActiveRow

    If ss1.MaxRows < 1 Then Exit Sub

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol

        Case 1

            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1
                DD.sWitch = "SP"
                DD.rControl.Add Item:=1

                 
                  DD.nameType = "2"
                
                Call Gf_StdSPEC_DD(M_CN1, KeyCode)
            End If
        Case 6
            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1
                DD.sWitch = "SP"
                DD.sKey = "Q0074"
                DD.rControl.Add Item:=6


                DD.nameType = "2"

                Call Gf_Common_DD(M_CN1, KeyCode)
'            Else
'
'                If Gf_GetCellText(ss1, irow, iCol) = "" Then
'                    Call GP_SET_CELL_VALUE(ss1, irow, iCol + 1, "")
'                End If
            End If
        Case 7, 9, 11
            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1
                DD.sWitch = "SP"
                DD.sKey = "Q0073"
                DD.rControl.Add Item:=ss1.ActiveCol


                DD.nameType = "2"

                Call Gf_Common_DD(M_CN1, KeyCode)
'            Else
'
'                If Gf_GetCellText(ss1, irow, iCol) = "" Then
'                    Call GP_SET_CELL_VALUE(ss1, irow, iCol + 1, "")
'                End If
            End If
       Case 8, 10, 12
            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1
                DD.sWitch = "SP"
                DD.rControl.Add Item:=ss1.ActiveCol

                DD.nameType = "2"

                Call Gf_HEAT_COND_DD(M_CN1, KeyCode)
     
'            Else
'
'                If Gf_GetCellText(ss1, irow, iCol) = "" Then
'                    Call GP_SET_CELL_VALUE(ss1, irow, iCol + 1, "")
'                End If
            End If


    End Select
     Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 15)
  
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
Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
            
    With ss2
    
        If Gf_GetCellNullCheck(ss2, Row, 1) <> "" And Gf_GetCellNullCheck(ss2, Row, 2) <> "" Then
            Call GP_SET_CELL_VALUE(ss2, 1, 1, Gf_GetCellNullCheck(ss2, Row, 1))
            Call GP_SET_CELL_VALUE(ss2, 1, 2, Gf_GetCellNullCheck(ss2, Row, 2))
        End If
        
        .MaxRows = 1
        .Height = 255
        
        txt_THK_MIN.Text = Gf_GetCellNullCheck(ss2, 1, 1)
        txt_THK_MAX.Text = Gf_GetCellNullCheck(ss2, 1, 2)
        
    
    End With
End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
        txt_THK_MIN.Text = Gf_GetCellNullCheck(ss2, 1, 1)
        txt_THK_MAX.Text = Gf_GetCellNullCheck(ss2, 1, 2)
End Sub

Private Sub txt_STDSPEC_Change()

   If Len(Trim(txt_STDSPEC.Text)) = 0 Then
        ss2.MaxRows = 1
        ss2.Height = 255
        Call GP_SET_CELL_VALUE(ss2, 1, 1, "")
        Call GP_SET_CELL_VALUE(ss2, 1, 2, "")
        txt_THK_MIN.Text = ""
        txt_THK_MAX.Text = ""
   End If
   
End Sub
