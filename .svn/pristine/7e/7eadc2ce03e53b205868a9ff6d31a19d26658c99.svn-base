VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGA2060C 
   Caption         =   "库情况查询界面_CGA2060C"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   1380
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9900
   ScaleWidth      =   15420
   WindowState     =   2  'Maximized
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   60
      Top             =   180
      Width           =   1230
      _ExtentX        =   2170
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
   Begin InDate.UDate sdt_in_plt_date 
      Height          =   315
      Left            =   1350
      TabIndex        =   0
      Tag             =   "起始日期"
      Top             =   180
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
   Begin InDate.UDate sdt_out_plt_date 
      Height          =   315
      Left            =   3105
      TabIndex        =   1
      Tag             =   "起始日期"
      Top             =   180
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8565
      Left            =   60
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   15150
      _Version        =   393216
      _ExtentX        =   26723
      _ExtentY        =   15108
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ButtonDrawMode  =   4
      ColsFrozen      =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   28
      MaxRows         =   20
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "CGA2060C.frx":0000
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   2910
      TabIndex        =   2
      Top             =   300
      Width           =   195
   End
End
Attribute VB_Name = "CGA2060C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name
'-- Program Name
'-- Program ID        CGA2060
'-- Designer          SHIN.C.S
'-- Coder             SHIN.C.S
'-- Date              2007.07.26
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection

Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim iSumCol As New Collection       'Sum Column

Dim S As String

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    '===============< 入库出库情况查询 Collection define  Start>======================================
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(sdt_in_plt_date, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(sdt_out_plt_date, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGA2060C.P_REFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
   '===============< 入库出库情况查询 Collection define  End >======================================
    
    'Duplicate Count
    iDupCnt = 3

    
    'Sum Column Count
    iSumCnt = 24
    
    'Sum Column Setting
    iSumCol.Add Item:=4
    iSumCol.Add Item:=5
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
    iSumCol.Add Item:=21
    iSumCol.Add Item:=22
    iSumCol.Add Item:=23
    iSumCol.Add Item:=24
    iSumCol.Add Item:=25
    iSumCol.Add Item:=26
    iSumCol.Add Item:=27

    Call Gp_Sp_ColHidden(ss1, 28, True)
    
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
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)

    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
        
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "F-System.INI", Me.Name)

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    

    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "F-System.INI", Me.Name)

    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    Set iSumCol = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing

    
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    rControl(1).SetFocus
    
End Sub

Public Sub Form_Ref()

Dim sMesg As String
Dim sQuery As String
    
    If sdt_in_plt_date = "" Or sdt_out_plt_date = "" Then
       Call Gp_MsgBoxDisplay("请输入日期", "I")
       Exit Sub
    End If

    sQuery = "{ CALL " & "CGA2060C.P_REFER" & "("
    sQuery = sQuery & " '" & sdt_in_plt_date.RawData & "','" & sdt_out_plt_date.RawData & "'"
    sQuery = sQuery & ")"
    sQuery = sQuery & "}"

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

    If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), sQuery, 1, 2, iSumCnt, iSumCol, False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

    'If Gf_Sp_Refer(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

End Sub
Public Sub Sheet_Ref()

    'dddddd

End Sub
Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
  Dim iRowCount As Long
  Dim MaxRow    As Long
  Dim iRow      As Integer
  Dim grd       As String
  Dim sStlgrd   As String
  Dim sThk      As Long
  Dim sWid      As Long
  
    If ss1.MaxRows < 1 Or ROW <= 0 Then Exit Sub
        If ss1.ActiveCol <= 3 Then Exit Sub
        
            ss1.ROW = ss1.ActiveRow
            ss1.Col = 28
            
            If Len(Trim(ss1.Text)) = 0 Then Exit Sub
            
            CGA2061C.sdt_in_plt_date = sdt_in_plt_date
            CGA2061C.sdt_out_plt_date = sdt_out_plt_date
            
            CGA2061C.txt_act_stlgrd = Trim(ss1.Text)
            
            ss1.Col = 1
            sStlgrd = Trim(ss1.Text)
            CGA2061C.txt_act_stlgrd_dec = Trim(ss1.Text)
            
                    
            ss1.Col = 2
            sThk = CLng(Trim(ss1.Text))
            CGA2061C.txt_Thk.Value = Trim(ss1.Value)
            CGA2061C.txt_thk_to.Value = Trim(ss1.Value)
            
            If sThk = 0 Then Exit Sub
            
            ss1.Col = 3
            sWid = CLng(Trim(ss1.Text))
            CGA2061C.txt_Wid.Value = Trim(ss1.Value)
            CGA2061C.txt_wid_to.Value = Trim(ss1.Value)
            
            If sWid = 0 Then Exit Sub
            
            CGA2061C.txt_cur_inv_code.Text = ""
            CGA2061C.txt_cur_inv.Text = ""
            
            CGA2061C.Show
            CGA2061C.ss1.MaxRows = 0
            CGA2061C.ss2.MaxRows = 0
            CGA2061C.ss3.MaxRows = 0
            CGA2061C.ss4.MaxRows = 0
            CGA2061C.SetFocus
            
    '        ss1.ActiveCol = Col
    
            Select Case ss1.ActiveCol
            
                Case 4, 5, 10, 11, 16, 17, 22, 23
                    
                    CGA2061C.txt_plt.Text = "B1"
                    Call CGA2061C.txt_plt_KeyUp(0, 0)
                    
                Case 6, 7, 12, 13, 18, 19, 24, 25
                    
                    CGA2061C.txt_plt.Text = "B3"
                    Call CGA2061C.txt_plt_KeyUp(0, 0)
                    
                Case 8, 9, 14, 15, 20, 21, 26, 27
                    
                    CGA2061C.txt_plt.Text = "ZZ"
                    CGA2061C.txt_plt_name.Text = "其他"
                
            End Select
    
            Select Case ss1.ActiveCol
            
                Case 4, 5, 6, 7, 8, 9
                    
                    CGA2061C.SSTab1.Tab = 0
                    
                Case 10, 11, 12, 13, 14, 15
                    
                    CGA2061C.SSTab1.Tab = 1
                    
                Case 16, 17, 18, 19, 20, 21
                    
                    CGA2061C.SSTab1.Tab = 2
                
                Case 22, 23, 24, 25, 26, 27
                    
                    CGA2061C.SSTab1.Tab = 3
            
            End Select
            
                
'            If ss1.ActiveCol = 4 Or ss1.ActiveCol = 5 Or ss1.ActiveCol = 6 Or ss1.ActiveCol = 7 Or ss1.ActiveCol = 8 Or ss1.ActiveCol = 9 Then
'               CGA2061C.SSTab1.Tab = 0
'            End If
'
'            If ss1.ActiveCol = 10 Or ss1.ActiveCol = 11 Or ss1.ActiveCol = 12 Or ss1.ActiveCol = 13 Or ss1.ActiveCol = 14 Or ss1.ActiveCol = 15 Then
'               CGA2061C.SSTab1.Tab = 1
'            End If
'
'            If ss1.ActiveCol = 16 Or ss1.ActiveCol = 17 Or ss1.ActiveCol = 18 Or ss1.ActiveCol = 19 Or ss1.ActiveCol = 20 Or ss1.ActiveCol = 21 Then
'              CGA2061C.SSTab1.Tab = 2
'            End If
'
'            If ss1.ActiveCol = 22 Or ss1.ActiveCol = 23 Or ss1.ActiveCol = 24 Or ss1.ActiveCol = 25 Or ss1.ActiveCol = 26 Or ss1.ActiveCol = 27 Then
'              CGA2061C.SSTab1.Tab = 3
'            End If
            
            CGA2061C.Form_Ref
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal ROW As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'    If Row > 0 Then
'        Set Active_Spread = Me.ss1
'        PopupMenu MDIMain.PopUp_Spread
'    End If

End Sub



