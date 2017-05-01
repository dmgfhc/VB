VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AKP1019C 
   Caption         =   "炼钢日成本核算(按钢种)_AKP1019C"
   ClientHeight    =   9225
   ClientLeft      =   540
   ClientTop       =   1695
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_prc_line 
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
      ItemData        =   "AKP1019C.frx":0000
      Left            =   6810
      List            =   "AKP1019C.frx":000D
      TabIndex        =   1
      Tag             =   "连铸机号"
      Top             =   90
      Width           =   750
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8715
      Left            =   45
      TabIndex        =   0
      Top             =   480
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   15372
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
      MaxCols         =   45
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AKP1019C.frx":001A
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   210
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "查询日期"
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
   Begin InDate.UDate txt_from_DATE 
      Height          =   315
      Left            =   1620
      TabIndex        =   2
      Tag             =   "起始日期"
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
   End
   Begin InDate.UDate txt_to_date 
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Tag             =   "起始日期"
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
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   5400
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "连铸机号"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3150
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "AKP1019C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      炼钢日成本核算(按钢种平均)
'-- Program ID        AKP1019C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2006.4.30
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
Public sProc_cd As String

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
    
    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
  Call Gp_Ms_Collection(txt_from_DATE, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_to_date, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(cbo_prc_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Next iCol
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AKP1019C.P_REFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    
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
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
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
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If

End Sub

Public Sub Form_Ref()

    Dim sShow As String
    Dim sLink As String
    Dim sTemp_Mana_Code As String
    Dim sTemp_Code As String
    Dim i, j As Integer
    Dim slab_wgt, hm_wgt, scr_out, scr_self, pig_hm, cost1 As Double
    Dim rec_mu, rec_mu_no, cost2 As Double
    Dim alloy_wgt, alloy_cost, sub_wgt, sub_cost, nh_cb_cost, nh_qt_cost As Double
    Dim oxy_wgt, oxy_cost, n_wgt, n_cost, ar_wgt, ar_cost, water_wgt, water_cost, elec_wgt, elec_cost, gas_wgt, gas_cost, rec_gas_wgt, rec_gas_cost, cost3 As Double
    Dim sal, reward, control, fix_fee, tot_cost As Double

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
    
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        slab_wgt = Pf_Sp_ColSum(ss1, 4, 1, ss1.MaxRows)
        hm_wgt = Pf_Sp_ColSum(ss1, 7, 1, ss1.MaxRows)
        scr_out = Pf_Sp_ColSum(ss1, 8, 1, ss1.MaxRows)
        scr_self = Pf_Sp_ColSum(ss1, 9, 1, ss1.MaxRows)
        pig_hm = Pf_Sp_ColSum(ss1, 10, 1, ss1.MaxRows)
        cost1 = Pf_Sp_ColSum(ss1, 11, 1, ss1.MaxRows)
        
        rec_mu = Pf_Sp_ColSum(ss1, 12, 1, ss1.MaxRows)
        rec_mu_no = Pf_Sp_ColSum(ss1, 13, 1, ss1.MaxRows)
        cost2 = Pf_Sp_ColSum(ss1, 15, 1, ss1.MaxRows)
        
        alloy_wgt = Pf_Sp_ColSum(ss1, 16, 1, ss1.MaxRows)
        alloy_cost = Pf_Sp_ColSum(ss1, 17, 1, ss1.MaxRows)
        sub_wgt = Pf_Sp_ColSum(ss1, 18, 1, ss1.MaxRows)
        sub_cost = Pf_Sp_ColSum(ss1, 19, 1, ss1.MaxRows)
        nh_cb_cost = Pf_Sp_ColSum(ss1, 20, 1, ss1.MaxRows)
        nh_qt_cost = Pf_Sp_ColSum(ss1, 21, 1, ss1.MaxRows)
        
        oxy_wgt = Pf_Sp_ColSum(ss1, 22, 1, ss1.MaxRows)
        oxy_cost = Pf_Sp_ColSum(ss1, 23, 1, ss1.MaxRows)
        n_wgt = Pf_Sp_ColSum(ss1, 24, 1, ss1.MaxRows)
        n_cost = Pf_Sp_ColSum(ss1, 25, 1, ss1.MaxRows)
        ar_wgt = Pf_Sp_ColSum(ss1, 26, 1, ss1.MaxRows)
        ar_cost = Pf_Sp_ColSum(ss1, 27, 1, ss1.MaxRows)
        water_wgt = Pf_Sp_ColSum(ss1, 28, 1, ss1.MaxRows)
        water_cost = Pf_Sp_ColSum(ss1, 29, 1, ss1.MaxRows)
        elec_wgt = Pf_Sp_ColSum(ss1, 30, 1, ss1.MaxRows)
        elec_cost = Pf_Sp_ColSum(ss1, 31, 1, ss1.MaxRows)
        gas_wgt = Pf_Sp_ColSum(ss1, 32, 1, ss1.MaxRows)
        gas_cost = Pf_Sp_ColSum(ss1, 33, 1, ss1.MaxRows)
        rec_gas_wgt = Pf_Sp_ColSum(ss1, 34, 1, ss1.MaxRows)
        rec_gas_cost = Pf_Sp_ColSum(ss1, 35, 1, ss1.MaxRows)
        cost3 = Pf_Sp_ColSum(ss1, 36, 1, ss1.MaxRows)
        
        sal = Pf_Sp_ColSum(ss1, 37, 1, ss1.MaxRows)
        reward = Pf_Sp_ColSum(ss1, 38, 1, ss1.MaxRows)
        control = Pf_Sp_ColSum(ss1, 39, 1, ss1.MaxRows)
        fix_fee = Pf_Sp_ColSum(ss1, 40, 1, ss1.MaxRows)
        tot_cost = Pf_Sp_ColSum(ss1, 41, 1, ss1.MaxRows)
        
        ss1.MaxRows = ss1.MaxRows + 1
        ss1.Row = ss1.MaxRows
        
        ss1.Col = 1: ss1.Text = "总计"
        ss1.Col = 4: ss1.Text = Str(slab_wgt)
        ss1.Col = 7: ss1.Text = Str(hm_wgt)
        ss1.Col = 8: ss1.Text = Str(scr_out)
        ss1.Col = 9: ss1.Text = Str(scr_self)
        ss1.Col = 10: ss1.Text = Str(pig_hm)
        ss1.Col = 11: ss1.Text = Str(cost1)
        ss1.Col = 12: ss1.Text = Str(rec_mu)
        ss1.Col = 13: ss1.Text = Str(rec_mu_no)
        ss1.Col = 15: ss1.Text = Str(cost2)
        ss1.Col = 16: ss1.Text = Str(alloy_wgt)
        ss1.Col = 17: ss1.Text = Str(alloy_cost)
        ss1.Col = 18: ss1.Text = Str(sub_wgt)
        ss1.Col = 19: ss1.Text = Str(sub_cost)
        ss1.Col = 20: ss1.Text = Str(nh_cb_cost)
        ss1.Col = 21: ss1.Text = Str(nh_qt_cost)
        ss1.Col = 22: ss1.Text = Str(oxy_wgt)
        ss1.Col = 23: ss1.Text = Str(oxy_cost)
        ss1.Col = 24: ss1.Text = Str(n_wgt)
        ss1.Col = 25: ss1.Text = Str(n_cost)
        ss1.Col = 26: ss1.Text = Str(ar_wgt)
        ss1.Col = 27: ss1.Text = Str(ar_cost)
        ss1.Col = 28: ss1.Text = Str(water_wgt)
        ss1.Col = 29: ss1.Text = Str(water_cost)
        ss1.Col = 30: ss1.Text = Str(elec_wgt)
        ss1.Col = 31: ss1.Text = Str(elec_cost)
        ss1.Col = 32: ss1.Text = Str(gas_wgt)
        ss1.Col = 33: ss1.Text = Str(gas_cost)
        ss1.Col = 34: ss1.Text = Str(rec_gas_wgt)
        ss1.Col = 35: ss1.Text = Str(rec_gas_cost)
        ss1.Col = 36: ss1.Text = Str(cost3)
        ss1.Col = 37: ss1.Text = Str(sal)
        ss1.Col = 38: ss1.Text = Str(reward)
        ss1.Col = 39: ss1.Text = Str(control)
        ss1.Col = 40: ss1.Text = Str(fix_fee)
        ss1.Col = 41: ss1.Text = tot_cost
        If tot_cost <> 0 And slab_wgt <> 0 Then
            ss1.Col = 42: ss1.Text = Round(tot_cost / slab_wgt, 2)
        End If
        
        With ss1
        For i = 1 To .MaxRows
            .Row = i
            For j = 4 To .MaxCols
                .Col = j
                If Val(.Text) = 0 Then
                   .Text = ""
                End If
            Next j
        Next i
        End With
        
        ss1.OperationMode = OperationModeNormal
        Call Gp_Sp_EvenRowBackcolor(ss1, 1)
        Call Gp_Sp_BlockColor(sc1.Item("Spread"), 1, ss1.MaxCols, ss1.MaxRows, ss1.MaxRows, BLACK, &HE6E6FF)
        
    Else
        MsgBox "请先生成查询日期范围内每一天的核算数据!", vbInformation, "系统提示信息"
    End If
            
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

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub
'
'Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
'
'If Col = 2 Then
'    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0
'
'End If
'
'End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Function Pf_Sp_ColSum(ByVal sPname As Variant, iCol As Long, Optional Start_Row As Long = 1, _
                                                                    Optional End_Row As Long = 0) As Double
        
    Dim lCount As Long
    Dim dSum As Double
    
    With sPname
    
        If End_Row > .MaxRows Or End_Row = 0 Then
            End_Row = .MaxRows
        End If
        
        .Col = iCol
        
        For lCount = Start_Row To End_Row
            .Row = lCount
            If .Text <> "" Then
                dSum = dSum + .VALUE
            End If
        Next lCount
    
    End With
    
    Pf_Sp_ColSum = dSum
    
End Function

