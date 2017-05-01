VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form DGA1190C 
   Caption         =   "热处理产品日报表(按订单)_DGA1190C"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10590
   ScaleWidth      =   16515
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_PLT_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "工厂"
      Top             =   210
      Width           =   1680
   End
   Begin FPSpread.vaSpread SS1 
      Height          =   8415
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   14895
      _Version        =   393216
      _ExtentX        =   26273
      _ExtentY        =   14843
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
      MaxCols         =   19
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "DGA1190C.frx":0000
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   4710
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "生产厂"
      Top             =   210
      Width           =   960
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   120
      Top             =   210
      Width           =   1080
      _ExtentX        =   1905
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
   Begin InDate.UDate SDT_PROD_DATE 
      Height          =   315
      Left            =   1245
      TabIndex        =   0
      Tag             =   "日期"
      Top             =   210
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
   Begin InDate.ULabel ULabel17 
      Height          =   315
      Left            =   3690
      Top             =   210
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "生产厂"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin Threed.SSCommand cmd_upd 
      Height          =   345
      Left            =   12810
      TabIndex        =   4
      Top             =   210
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
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
      Caption         =   "更新最新数据"
   End
End
Attribute VB_Name = "DGA1190C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   HTM System
'-- Program Name      热处理产品日报表(按订单)
'-- Program ID        DGA1190C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2008.7.8
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
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(SDT_PROD_DATE, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_plt, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(SS1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 18, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 19, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=SS1, Key:="Spread"
    sc1.Add Item:="DGA1190C.P_REFER", Key:="P-R"
    sc1.Add Item:="DGA1190C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="DGA1190C.P_SONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=SS1.MaxCols, Key:="Last"
        
    Call Gp_Sp_ColHidden(SS1, 18, True)
    Call Gp_Sp_ColHidden(SS1, 19, True)
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cmd_upd_Click()
    If SDT_PROD_DATE.RawData = "" Then
       MsgBox "请先输入日期", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    Call UPD_DATA
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
    If Mid(sAuthority, 3, 1) = "1" Then
       cmd_upd.Enabled = True
    End If
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(sc1.Item("Spread"), False)

    Call Gp_Sp_ColGet(sc1.Item("Spread"), "DG-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


    Call Gp_Sp_ColSet(sc1.Item("Spread"), "DG-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
        
    Set pColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set iColumn1 = Nothing
    Set aColumn1 = Nothing
    Set lColumn1 = Nothing
    
    Set Mc1 = Nothing
    
    Set sc1 = Nothing
    
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)

End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
        
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ins()

    If SDT_PROD_DATE.RawData = "" Then
       MsgBox "请先输入日期", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    If txt_plt.Text = "" Then
       MsgBox "请先输入生产厂", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    SS1.Row = SS1.ActiveRow
    SS1.Col = 18
    SS1.Text = SDT_PROD_DATE.RawData
    SS1.Col = 19
    SS1.Text = txt_plt.Text
End Sub

Public Sub Form_Ref()

Dim iRow  As Long
Dim iCol  As Long

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
        For iRow = 1 To SS1.MaxRows
            SS1.Row = iRow
            For iCol = 7 To 14
                SS1.Col = iCol
                If Val(SS1.Text & "") = 0 Then
                    SS1.Text = ""
                End If
            Next iCol
        Next iRow
        
        SS1.OperationMode = OperationModeNormal
    End If
                
    Exit Sub

Refer_Err:

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

Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
Dim ORD_NO As String
Dim ORD_ITEM As String
Dim sQuery As String
    SS1.Row = SS1.ActiveRow
    SS1.Col = 1
    ORD_NO = SS1.Text
    SS1.Col = 2
    ORD_ITEM = SS1.Text

    If SS1.ActiveCol = 2 And ORD_NO <> "" And ORD_ITEM <> "" Then
       sQuery = "SELECT GF_CustNameFind(CUST_CD), STDSPEC, ORD_SIZE, TOT_WGT from bp_order_item WHERE ORD_NO = '" & ORD_NO & "' AND ORD_ITEM = '" & ORD_ITEM & "'"

       Set AdoRs = New ADODB.Recordset
       AdoRs.Open sQuery, M_CN1, adOpenKeyset

       If (Not AdoRs.BOF) Or (Not AdoRs.EOF) Then
          SS1.Col = 3
          SS1.Text = AdoRs.Fields(0).Value
          SS1.Col = 4
          SS1.Text = AdoRs.Fields(1).Value
          SS1.Col = 5
          SS1.Text = AdoRs.Fields(2).Value
          SS1.Col = 6
          SS1.Text = CStr(AdoRs.Fields(3).Value)
       End If

    End If
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
    End If
        
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

'Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'
'    If Row > 0 Then
'        Set Active_Spread = Me.ss1
'        PopupMenu MDIMain.PopUp_Spread
'    End If
'
'End Sub
Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call Form_Ref
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt
        DD.rControl.Add Item:=TXT_PLT_NAME
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_plt.Text)) = txt_plt.MaxLength Then
        TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(txt_plt.Text), 2)
    Else
        TXT_PLT_NAME.Text = ""
    End If
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    If SDT_PROD_DATE.RawData = "" Then
       MsgBox "请先输入日期", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    If txt_plt.Text = "" Then
       MsgBox "请先输入生产厂", vbCritical, "系统提示信息"
       Exit Sub
    End If

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
    SS1.Row = SS1.ActiveRow
    SS1.Col = 18
    SS1.Text = SDT_PROD_DATE.RawData
    SS1.Col = 19
    SS1.Text = txt_plt.Text

End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub UPD_DATA()

    Dim OutParam(1, 4)      As Variant
    Dim adoCmd              As ADODB.Command
    Dim Response            As Variant
    Dim OUTPUT              As String

    On Error GoTo Process_Exec_ERROR
    
    Response = MsgBox("确定要更新 " + Mid(SDT_PROD_DATE.RawData, 1, 4) + "年" + Mid(SDT_PROD_DATE.RawData, 5, 2) + "月" + Mid(SDT_PROD_DATE.RawData, 7, 2) + "日  " + "的数据吗?", vbYesNo, "系统提示信息")
    If Response = vbNo Then
        Exit Sub
    End If
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
             
    Screen.MousePointer = vbHourglass
    
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.CommandText = "{call DGA1190P ('" + SDT_PROD_DATE.RawData + "','" + txt_plt.Text + "', ?)}"
    
    adoCmd.Execute , , adExecuteNoRecords
    
    If adoCmd("arg_e_msg") <> "" Then
        Call MsgBox(adoCmd("arg_e_msg"), vbInformation, "系统提示信息")
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        'Process Error Check
        Call MsgBox("数据更新成功！", vbInformation, "系统提示信息")
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:
    
    Set adoCmd = Nothing
    Call Gp_MsgBoxDisplay(Err.Description & "{call DGA1190P ('" + SDT_PROD_DATE.RawData + "','" + txt_plt.Text + "',?)}")
    
End Sub

