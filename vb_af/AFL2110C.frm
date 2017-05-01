VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AFL2110C 
   Caption         =   "外卖板坯号录入_AFL2110C"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   13410
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_ord_no 
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
      ItemData        =   "AFL2110C.frx":0000
      Left            =   8160
      List            =   "AFL2110C.frx":0002
      TabIndex        =   5
      Top             =   120
      Width           =   5205
   End
   Begin VB.TextBox txt_cust_cd 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1695
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "客户"
      Top             =   120
      Width           =   870
   End
   Begin VB.TextBox txt_cust_cd_name 
      BackColor       =   &H00FFFFFF&
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
      Left            =   2565
      MaxLength       =   40
      TabIndex        =   3
      Tag             =   "客户"
      Top             =   120
      Width           =   4095
   End
   Begin VB.TextBox txt_in_slab_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   1695
      MaxLength       =   10
      TabIndex        =   1
      Top             =   600
      Width           =   1200
   End
   Begin VB.TextBox txt_mat_no 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   8160
      MaxLength       =   10
      TabIndex        =   0
      Top             =   600
      Width           =   1200
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   6840
      Top             =   600
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "物料号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   120
      Top             =   600
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      Caption         =   "外卖坯号"
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
      Height          =   7815
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   14610
      _Version        =   393216
      _ExtentX        =   25770
      _ExtentY        =   13785
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
      MaxCols         =   14
      MaxRows         =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AFL2110C.frx":0004
   End
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      Caption         =   "客户"
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
   Begin InDate.ULabel ULabel26 
      Height          =   315
      Left            =   6840
      Top             =   120
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "订单号"
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
End
Attribute VB_Name = "AFL2110C"
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
'-- Program Name      板坯外卖信息录入
'-- Program ID        AFL2110C
'-- Designer          WUTAO
'-- Coder             WUTAO
'-- Date              2006.10.26
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

'----> THIS variable USE IS AFL2110C PROGRAM <--------------
Public islab_no As Long
Public chkNo As Integer
Public sslab_no As String
Public botIntLen As Long
'----------------------------------------------------------

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
       Call Gp_Ms_Collection(cbo_ord_no, "p", "n", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_in_slab_no, "p", " ", " ", "", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFL2110C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AFL2110C.P_REFER", Key:="P-R"
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
    
    Call Gp_Sp_ColHidden(ss1, 11, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    txt_cust_cd.BackColor = &HC0FFFF
    txt_cust_cd_name.BackColor = &HC0FFFF
    'txt_in_slab_no.BackColor = &HC0FFFF
    
End Sub

Private Sub cbo_ord_no_DropDown()
  Dim sQuery As String
   
   'sQuery = "SELECT  distinct ORD_NO||ORD_ITEM||'  '||THK||'*'||WID||'*'||LEN  From NISCO.FP_SLAB  WHERE  CUST_CD  LIKE  '" + txt_cust_cd.Text + "'|| '%'  ORDER BY ORD_NO"
    If txt_cust_cd <> "" And txt_cust_cd_name <> "" Then
        sQuery = "        SELECT  ORD_NO||ORD_ITEM||'  '||ORD_THK||'*'||ORD_WID||'*'||ORD_LEN||'  '||STDSPEC "
        sQuery = sQuery & " From  NISCO.BP_ORDER_ITEM  "
        sQuery = sQuery & "WHERE  CUST_CD  LIKE '" & txt_cust_cd & "%'  "
        sQuery = sQuery & "  AND PROD_CD = 'SL'"
        sQuery = sQuery & "  AND ORD_STS = 'E'"
        sQuery = sQuery & "  AND NVL(HOLD_FL,'N') <> 'Y' "
        sQuery = sQuery & "GROUP BY ORD_NO||ORD_ITEM||'  '||ORD_THK||'*'||ORD_WID||'*'||ORD_LEN||'  '||STDSPEC "
        
        Call Gf_ComboAdd(M_CN1, cbo_ord_no, sQuery)
    End If
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

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

    Call Form_Activate
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "F-System.INI", Me.Name)
    
    txt_in_slab_no.Enabled = True
    'cmd_add_rec.Enabled = False
    
'----> THIS variable USE IS AFL2011C PROGRAM <--------------
    chkNo = 0
    botIntLen = 0
    sslab_no = ""
'-----------------------------------------------------------
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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        'Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        'txt_in_slab_no.Enabled = False
        'cmd_add_rec.Enabled = False
    End If

    txt_cust_cd = ""
    txt_cust_cd_name = ""
    cbo_ord_no.Clear
    txt_mat_no = ""
    txt_in_slab_no = ""
    
    chkNo = 0
    botIntLen = 0
    sslab_no = ""
End Sub

Public Sub Form_Ref()
Dim i As Integer
On Error GoTo Refer_Err


    'ERROR CHECK
    If txt_cust_cd = "" And cbo_ord_no = "" And txt_in_slab_no = "" Then
       Call Gp_MsgBoxDisplay("请输入相应的查询条件...!", "Q", "")
       txt_cust_cd.SetFocus
       Exit Sub
    End If
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Nothing, Mc1("mControl")) Then
        Call Form_Activate
        'Call Gp_Ms_ControlLock(Mc1("pControl"), True)
        ss1.SetFocus
        If ss1.MaxRows >= 1 Then
           txt_in_slab_no.Enabled = True
           'cmd_add_rec.Enabled = True
        End If
    
    End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then Call Form_Activate
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    
    With ss1
        .Row = .ActiveRow
        .Col = 8
        .Text = sUserID
    End With

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
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim icount, icount1 As Integer
    Dim iText As Long
    Dim k, i, len_z As Long
    Dim sCol As Long
    Dim zeroSlab As String
    
    
    If chkNo = 0 And Len(Trim(txt_in_slab_no.Text)) <> 0 Then
       ''''''''''''''''''''''''''''''
       For i = Len(Trim(txt_in_slab_no.Text)) To 1 Step -1
           If Asc(Mid(txt_in_slab_no.Text, i, 1)) < 48 Or Asc(Mid(txt_in_slab_no.Text, i, 1)) > 57 Then
              k = i
              Exit For
           End If
       Next i
       islab_no = CDbl(Mid(txt_in_slab_no.Text, k + 1, Len(Trim(txt_in_slab_no.Text)) - k))
       botIntLen = Len(Trim(txt_in_slab_no.Text)) - k
       sslab_no = Mid(txt_in_slab_no.Text, 1, k)
       ''''''''''''''''''''''''''''''
    End If

    If Len(Trim(txt_in_slab_no.Text)) <> 0 And Row > 0 And Col = 2 Then
       sCol = Col
       ss1.Col = 0
       ss1.Row = Row
       ss1.Text = "Update"
       ss1.Col = 14
       ss1.Text = sUserID
       ss1.Col = sCol
       zeroSlab = String(botIntLen - Len(CStr(islab_no)), "0")
       ss1.Text = UCase(sslab_no & zeroSlab & CStr(islab_no))
       islab_no = islab_no + 1
       chkNo = chkNo + 1
    End If


    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)

End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
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

Private Sub cmd_add_rec_Click()

    Dim icount As Integer
    Dim icount1 As Integer
    Dim islab_no As Long
    Dim sslab_no As String
    Dim strTem As String
    Dim len_z As String
    Dim i, k, j As Integer
    Dim A As Double
    
    If ss1.MaxRows < 1 Then
       Call Gp_MsgBoxDisplay("没有所要追加的纪录", "Q", "")
       Exit Sub
    End If
    
    If txt_in_slab_no.Text = "" Then
       Call Gp_MsgBoxDisplay("请输入首个外卖坯号", "Q", "")
       Exit Sub
    End If
    
    For i = Len(Trim(txt_in_slab_no.Text)) To 1 Step -1
        If Asc(Mid(txt_in_slab_no.Text, i, 1)) < 48 Or Asc(Mid(txt_in_slab_no.Text, i, 1)) > 57 Then
          k = i
          Exit For
        End If
    Next i
    
    islab_no = CDbl(Mid(txt_in_slab_no.Text, k + 1, Len(Trim(txt_in_slab_no.Text)) - k))
    'strTem = Mid(txt_in_slab_no.Text, 1, Len(txt_in_slab_no.Text) - k)
   ' islab_no = txt_in_slab_no.Text
    sslab_no = CStr(islab_no)
    For icount = 1 To ss1.MaxRows
        ss1.Col = 2
        ss1.Row = icount
        len_z = Len(txt_in_slab_no.Text) - k - Len(Trim(islab_no))
        If len_z > 0 Then
           For icount1 = 1 To len_z
               sslab_no = "0" + sslab_no
           Next icount1
        End If
        ss1.Text = Mid(txt_in_slab_no.Text, 1, k) + sslab_no
        ss1.Col = 0
        ss1.Text = "Update"
        ss1.Col = 11
        ss1.Text = sUserID
        islab_no = islab_no + 1
        sslab_no = CStr(islab_no)
    Next icount
    
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_cd_name

        DD.nameType = "1"

        Call Gf_Customer_DD(M_CN1, KeyCode)

        Exit Sub

    End If

    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_cd_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_cd_name.Text = ""
    End If
    
End Sub

