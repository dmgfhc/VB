VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQD0050C 
   Caption         =   "船板质量证明书编制"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9255
   ScaleWidth      =   15240
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_SMP_NO 
      Height          =   345
      Left            =   6990
      TabIndex        =   8
      Top             =   420
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox txt_END_CHECK 
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      Top             =   450
      Visible         =   0   'False
      Width           =   510
   End
   Begin Threed.SSCheck sck_TEST_END 
      Height          =   315
      Left            =   3690
      TabIndex        =   6
      Top             =   510
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      _Version        =   196609
      BackColor       =   14804173
      Caption         =   "船检完毕    "
      Alignment       =   1
      MaskColor       =   255
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   7605
      Left            =   10920
      TabIndex        =   5
      Top             =   1470
      Width           =   4245
      _Version        =   393216
      _ExtentX        =   7488
      _ExtentY        =   13414
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
      MaxCols         =   2
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQD0040C.frx":0000
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   7605
      Left            =   90
      TabIndex        =   4
      Top             =   1470
      Width           =   10635
      _Version        =   393216
      _ExtentX        =   18759
      _ExtentY        =   13414
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
      MaxCols         =   6
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQD0040C.frx":02DB
   End
   Begin VB.TextBox txt_INSP_CD 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1200
      TabIndex        =   3
      Top             =   60
      Width           =   645
   End
   Begin VB.TextBox txt_CONTROL_NO 
      Height          =   310
      Left            =   1170
      TabIndex        =   2
      Top             =   510
      Width           =   2265
   End
   Begin VB.TextBox txt_STD_ORGAN_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1860
      MaxLength       =   14
      TabIndex        =   1
      Top             =   60
      Width           =   1575
   End
   Begin VB.ComboBox cbo_STDSPEC 
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
      Left            =   4800
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   3690
      Top             =   60
      Width           =   1035
      _ExtentX        =   1826
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   90
      Top             =   60
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "检查机关"
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
      Left            =   90
      Top             =   510
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      Caption         =   "控制号"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   90
      Top             =   1050
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      Caption         =   "船检取样选择"
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
      ForeColor       =   255
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   11040
      Top             =   1050
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      Caption         =   "              号包含产品"
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
      ForeColor       =   255
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   10815
      X2              =   10815
      Y1              =   960
      Y2              =   9855
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   15405
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "AQD0050C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      质量证明书二次发放
'-- Program ID        AQD0030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Chu Kyo Su
'-- Coder             Chu Kyo Su
'-- Date              2003.07. 25
'-- Description       质量证明书二次发放
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

Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

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
Dim Mc2 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection
Dim bPrintCheck As Boolean

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

'---------------------------------------------------------------------------------------------
'------------------------------ Report Variable ----------------------------------------------
'---------------------------------------------------------------------------------------------
Dim crxApplication As New CRAXDRT.Application

Public WithEvents Report As CRAXDRT.Report
Attribute Report.VB_VarHelpID = -1

Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxSubreport As CRAXDRT.Report
'Dim CPProperties As CRAXDRT.ConnectionProperties
Dim cVal As New Collection
Dim sVal As New Collection
Dim sQueryHeadC As String        'QP_CERT_HEAD   -C  QUERY
Dim sQueryDetailC As String      'QP_CERT_DETAIL - C QUERY
Dim sQueryHeadS As String        'QP_CERT_HEAD   -S  QUERY
Dim sQueryDetailS As String      'QP_CERT_DETAIL - S QUERY
Dim AlreadyNoData As Integer                    ' Used to ensure the "No Data" message only appears once

'---------------------------------------------------------------------------------------------

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_INSP_CD, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_STD_ORGAN_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(cbo_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_CONTROL_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_END_CHECK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'       Call Gp_Ms_Collection(txt_TRNS_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
   
    'MASTER Collection
    Mc2.Add Item:=pControl1, Key:="pControl"
    Mc2.Add Item:=nControl1, Key:="nControl"
    Mc2.Add Item:=mControl1, Key:="mControl"
    Mc2.Add Item:=iControl1, Key:="iControl"
    Mc2.Add Item:=rControl1, Key:="rControl"
    Mc2.Add Item:=cControl1, Key:="cControl"
    Mc2.Add Item:=aControl1, Key:="aControl"
    Mc2.Add Item:=lControl1, Key:="lControl"
    
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQD0040C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQD0040C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQD0040C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    Sc2.Add Item:=ss1, Key:="Spread"
    Sc2.Add Item:="AQD0040C.P_REFER_D", Key:="P-R"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
     
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
    
    AlreadyNoData = 0
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
    Call subButtonHide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
    Set iColumn1 = Nothing
    Set pColumn1 = Nothing
    Set lColumn1 = Nothing
    Set nColumn1 = Nothing
    Set mColumn1 = Nothing
    Set aColumn1 = Nothing
    
    Set iColumn2 = Nothing
    Set pColumn2 = Nothing
    Set lColumn2 = Nothing
    Set nColumn2 = Nothing
    Set mColumn2 = Nothing
    Set aColumn2 = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Mc2 = Nothing
    Set Sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
    Call subButtonHide
    
End Sub



Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gf_Sp_Cls(Proc_Sc("SC2"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
     '  rControl(1).SetFocus
    End If
    
    Call subMasterClear

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
     If subCheck = True Then
        
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call subButtonHide
                Exit Sub
            End If
            
    Else
                
        GoTo Refer_Err
        
    End If
    
    Call subButtonHide
    
    bPrintCheck = False
    
    Exit Sub

Refer_Err:

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

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
  '  If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 12)
   ' End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 12)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
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


Private Sub subButtonHide()

    MDIMain.MenuTool.Buttons(4).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(5).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    
    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
    

End Sub




'########################################################################################################################
'####################################################### REPORT #########################################################
'########################################################################################################################

Private Sub Report_BeforeFormatPage(ByVal PageNumber As Long)
  '  MsgBox "Report_BeforeFormatPage " + PageNumber
 End Sub

Private Sub Report_AfterFormatPage(ByVal PageNumber As Long)
 '
End Sub

Private Sub Report_FieldMapping(reportFieldArray As Variant, ByVal databaseFieldArray As Variant, useDefault As Boolean)
'
End Sub

Private Sub Report_NoData(pCancel As Boolean)
    
    Dim Response As Integer
    
    Response = MsgBox("There are no records to display!  " & vbNewLine & _
           "Do you still want to display the empty report?", vbYesNo, _
           "No Records to Display")
    
    If Response = vbYes Then
        pCancel = False
    Else
        pCancel = True
    End If
    
'    If AlreadyNoData = 0 Then MsgBox "There are no Data Found!", vbExclamation, "No Data Found"
'    AlreadyNoData = AlreadyNoData + 1
'    pCancel = True
End Sub


'-----------------------------------------------------------------------
'---------------------------- Report Main ------------------------------
'-----------------------------------------------------------------------
Private Sub cmdReport_Click()
    
    If funGetSpreadData = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) = False Then Exit Sub
    
    Call funGetQuery
    
    Call subGetOracleData
    
    Call subReportStartC
    
    Call subReportStartS
        
    Call subButtonHide
    
    Call Form_Ref
    
    Screen.MousePointer = vbDefault
        
End Sub


'------------------------------------------------------------------------
'----------------- From Sppread -> Get Cert_No --------------------------
'------------------------------------------------------------------------
Private Function funGetSpreadData() As Boolean

    Dim i As Long
    Dim sParam As String
    
    On Error GoTo Err_Track
    
    Call subRemoveCollection
                        
    With ss1
        
        For i = 1 To .MaxRows
            
            If Gf_Get_Cell_Value(ss1, i, 1) = 1 Then
                
                Select Case Gf_Get_Cell_Value(ss1, i, 11)
                    Case "C"
                        cVal.Add Gf_Get_Cell_Value(ss1, i, 2)
                    Case "S"
                        sVal.Add Gf_Get_Cell_Value(ss1, i, 2)
                End Select
                                    
            End If
            
        Next i
                            
    End With
        
    If cVal.Count = 0 And sVal.Count = 0 Then Exit Function
           
    funGetSpreadData = True
    Exit Function

Err_Track:
    funGetSpreadData = False
End Function


'------------------------------------------------------------------------
'------------------------ Report Initialize    --------------------------
'------------------------------------------------------------------------
Private Sub subReportStartC()

    On Error GoTo Err_Track

    If sQueryHeadC <> "" Then

         Set Report = crxApplication.OpenReport(App.Path & "/AQD0040C.rpt", 1)
        
         For Each crxDatabaseTable In Report.Database.Tables
             
             crxDatabaseTable.Location = App.Path & "/Q_Report.mdb"
         
         Next crxDatabaseTable
         
         Set crxSubreport = Report.OpenSubreport("sub1")
         
         For Each crxDatabaseTable In crxSubreport.Database.Tables
             
             crxDatabaseTable.Location = App.Path & "/Q_Report.mdb"
         
         Next crxDatabaseTable
             
        Call frmReport.form_init(Me)
        frmReport.Show
         
        'Report.PrintOut
        
        'Report.PrintOut False
        
         
        Set Report = Nothing
        Set crxSubreport = Nothing
        
        bPrintCheck = True
        
    Else
        bPrintCheck = False
            
    End If
    
    Exit Sub
    
Err_Track:
            
End Sub

'------------------------------------------------------------------------
'------------------------ Report Initialize    --------------------------
'------------------------------------------------------------------------
Private Sub subReportStartS()

    On Error GoTo Err_Track

    ' Later Delete ------------------------------------------------------------
        If bPrintCheck = True Then Exit Sub

    If sQueryHeadS <> "" Then

         Set Report = crxApplication.OpenReport(App.Path & "/AQD0050C.rpt", 1)
        
         For Each crxDatabaseTable In Report.Database.Tables
             
             crxDatabaseTable.Location = App.Path & "/Q_Report.mdb"
         
         Next crxDatabaseTable
         
         Set crxSubreport = Report.OpenSubreport("SUB1")
         
         For Each crxDatabaseTable In crxSubreport.Database.Tables
             
             crxDatabaseTable.Location = App.Path & "/Q_Report.mdb"
         
         Next crxDatabaseTable
             
' Later Delete ------------------------------------------------------------
        Call frmReport.form_init(Me)
        frmReport.Show
         
         
         'Report.PrintOut False
         
         Set Report = Nothing
         Set crxSubreport = Nothing
            
    End If
    
        Exit Sub
    
Err_Track:
            
End Sub

'Remove Collection
Private Sub subRemoveCollection()

    Dim Num As Integer
    
    For Num = 1 To cVal.Count
        cVal.Remove 1
    Next Num
    
    For Num = 1 To sVal.Count
        sVal.Remove 1
    Next Num

End Sub


'------------------------------------------------------------------------------------------
'----------------------------- Oracle Data Select (To MDB ) -------------------------------
'------------------------------------------------------------------------------------------
Private Sub subGetOracleData()
    
    Dim sQuery As String
    Dim adoRs As ADODB.Recordset
    Dim arrRecords1 As Variant      'sQueryHeadC
    Dim arrRecords2 As Variant      'sQueryDetailC
    Dim arrRecords3 As Variant      'sQueryHeadS
    Dim arrRecords4 As Variant      'sQueryDetailS
    
    On Error GoTo Err_Track
                
    Set adoRs = New ADODB.Recordset
    
'-----------------------------------------------------------------------------
        
    If sQueryHeadC <> "" Then           'QP_CERT_HEAD TABLE - CERT_KND = "C"
        
        adoRs.Open sQueryHeadC, M_CN1, adOpenKeyset
    
        If Not adoRs.EOF Then
           
            arrRecords1 = adoRs.GetRows
            adoRs.Close
                        
            If sQueryHeadC <> "" Then   'QP_CERT_DETAIL TABLE - CERT_KND = "C"
            
                adoRs.Open sQueryDetailC, M_CN1, adOpenKeyset
            
                If Not adoRs.EOF Then
            
                    arrRecords2 = adoRs.GetRows
                    adoRs.Close
                
                End If
                
            End If
                                   
        End If
    
    End If
    
'-----------------------------------------------------------------------------

    If sQueryHeadS <> "" Then           'QP_CERT_HEAD TABLE - CERT_KND = "S"
        
        adoRs.Open sQueryHeadS, M_CN1, adOpenKeyset
    
        If Not adoRs.EOF Then
           
            arrRecords3 = adoRs.GetRows
            adoRs.Close
            
            If sQueryHeadS <> "" Then
                
                adoRs.Open sQueryDetailS, M_CN1, adOpenKeyset
            
                If Not adoRs.EOF Then   'QP_CERT_DETAIL TABLE - CERT_KND = "S"
            
                    arrRecords4 = adoRs.GetRows
                    adoRs.Close
                
                End If
            
            End If
                                   
        End If
    
    End If
    
'-----------------------------------------------------------------------------
    
    Set adoRs = Nothing
                
    
    If (sQueryHeadC = "" Or sQueryDetailC = "") And (sQueryHeadS = "" Or sQueryDetailS = "") Then Exit Sub
    
    Call subMdbUpdate(arrRecords1, arrRecords2, arrRecords3, arrRecords4)
    
    Exit Sub
    
Err_Track:
        
    If IsObject(adoRs) = True Then
        Set adoRs = Nothing
    End If
        
End Sub


'------------------------------------------------------------------------------------------------------------------------------------
'--------------------------- MDB Update --------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------------------
Private Sub subMdbUpdate(ByVal arrRecords1 As Variant, ByVal arrRecords2 As Variant, ByVal arrRecords3 As Variant, ByVal arrRecords4 As Variant)
    
    Dim i As Integer
    Dim iBlank As Integer
    
    Dim sQuery As String
    Dim oConn As ADODB.Connection
    Dim oRS As ADODB.Recordset
    Dim iMod As Integer
    Dim sCertNo As String
    Dim iPageRecCnt As Integer
    Dim PageSkip As Boolean
    Dim RstMode As Boolean
    Dim vMyPath As Variant
    Dim sMyConn As String
    
    vMyPath = App.Path
    sMyConn = "Provider=Microsoft.Jet.OLEDB.4.0;"
    sMyConn = sMyConn + "Data Source=" + vMyPath + "\Q_Report.MDB;"
    sMyConn = sMyConn + "User ID=admin; Password=;"
    
    On Error GoTo Err_Track
            
    Set oConn = New ADODB.Connection
'    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\Q_Report.MDB;User ID=admin; Password=;"
'     oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;"
'     Data Source=\Q_Report.mdb;User ID=admin; Password=;"
    'oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\Q_Report.MDB;Persist Security Info=False;"
     oConn.Open sMyConn

    Set oRS = New ADODB.Recordset
    oRS.CursorLocation = adUseClient
    oRS.Properties("Update Resync") = adResyncAutoIncrement

'-----------------------------------------------------------------------------
'AQD0040C_HEAD
'-----------------------------------------------------------------------------
    
    sQuery = "Select * From AQD0040C_HEAD"
    oRS.Open sQuery, oConn, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    If Not oRS.EOF Then
        sQuery = "DELETE FROM AQD0040C_HEAD"
        oConn.Execute sQuery, , adCmdText + adExecuteNoRecords
    End If
            
    If IsEmpty(arrRecords1) = False Then
    
        
            
        For i = 0 To UBound(arrRecords1, 2)
        
      
            oRS.AddNew
            
                oRS("CERT_NO").Value = arrRecords1(0, i)
                oRS("PROD_NAME").Value = arrRecords1(1, i)
                oRS("STDSPEC_NAME").Value = arrRecords1(2, i)
                oRS("PROD_SPEC_NO").Value = arrRecords1(3, i)
                oRS("CUST_NAME").Value = arrRecords1(4, i)
                oRS("COND_SUPPLY").Value = arrRecords1(5, i)
                oRS("PROD_SIZE").Value = arrRecords1(6, i)
                oRS("IMPACT_SMP_SIZE").Value = arrRecords1(7, i)
                oRS("QLTY_REC_NO").Value = arrRecords1(8, i)
                oRS("PONO").Value = arrRecords1(9, i)
                oRS("TRNS_NO").Value = arrRecords1(10, i)
                oRS("TRAIN_LINE_NAME").Value = arrRecords1(11, i)
                oRS("DEST_DETAIL").Value = arrRecords1(12, i)
                oRS("CERT_RPT_DATE").Value = arrRecords1(13, i)
                oRS("BEND_DIA").Value = "d=" + str(arrRecords1(14, i)) + "a"
                oRS("TEST_EMP").Value = arrRecords1(15, i)
                oRS("SHP_EMP").Value = arrRecords1(16, i)
                oRS("SUM_CNT").Value = arrRecords1(17, i)
                oRS("SUM_WGT").Value = arrRecords1(18, i)
            
            oRS.Update   ' Update local Recordset (since adLockBatchOptimistic)
        
        Next i
                
        oRS.MarshalOptions = adMarshalModifiedOnly
        oRS.UpdateBatch
            
    End If
    
    oRS.Close
            
'-----------------------------------------------------------------------------
'AQD0040C_DETAIL
'-----------------------------------------------------------------------------
    
    sQuery = "Select * From AQD0040C_DETAIL"
    oRS.Open sQuery, oConn, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    If Not oRS.EOF Then
        sQuery = "DELETE FROM AQD0040C_DETAIL"
        oConn.Execute sQuery, , adCmdText + adExecuteNoRecords
    End If
    
    iPageRecCnt = 0
        
    If IsEmpty(arrRecords2) = False Then
                
        sCertNo = arrRecords2(0, 0)     'First sCertNo
            
        For i = 0 To UBound(arrRecords2, 2)
            
            iPageRecCnt = iPageRecCnt + 1
                               
            oRS.AddNew
            
                oRS("CERT_NO").Value = arrRecords2(0, i)
                oRS("PROD_NO").Value = arrRecords2(1, i)
                oRS("STLGRD").Value = arrRecords2(2, i)
                oRS("PRDT_QNTY").Value = arrRecords2(3, i)
                oRS("PRDT_WGT").Value = arrRecords2(4, i)
                oRS("C_RST").Value = arrRecords2(5, i) * 100
                oRS("SI_RST").Value = arrRecords2(6, i) * 100
                oRS("MN_RST").Value = arrRecords2(7, i) * 100
                oRS("P_RST").Value = arrRecords2(8, i) * 1000
                oRS("S_RST").Value = arrRecords2(9, i) * 1000
                oRS("NB_RST").Value = arrRecords2(10, i) * 1000
                oRS("AL_RST").Value = arrRecords2(11, i) * 1000
                oRS("MO_RST").Value = arrRecords2(12, i) * 1000
                oRS("CU_RST").Value = arrRecords2(13, i) * 1000
                oRS("NI_RST").Value = arrRecords2(14, i) * 1000
                oRS("CR_RST").Value = arrRecords2(15, i) * 1000
                oRS("V_RST").Value = arrRecords2(16, i) * 1000
                oRS("TI_RST").Value = arrRecords2(17, i) * 1000
                oRS("CEQ_RST").Value = arrRecords2(18, i) * 100
                oRS("YP_RST").Value = arrRecords2(19, i)
                oRS("TS_RST").Value = arrRecords2(20, i)
                oRS("EL_RST").Value = arrRecords2(21, i)
                oRS("BEND_RST").Value = arrRecords2(22, i)
                oRS("UST_GRD_RST").Value = arrRecords2(23, i)
                oRS("IMPACT_TMP").Value = arrRecords2(24, i)
                oRS("IMPACT_RST_AVE").Value = arrRecords2(25, i)
                oRS("TIM_IMPACT_TMP").Value = arrRecords2(26, i)
                oRS("TIM_IMPACT_RST_AVE").Value = arrRecords2(27, i)
            
            oRS.Update   ' Update local Recordset (since adLockBatchOptimistic)
                        
            If i = UBound(arrRecords2, 2) Then              'LastRecord
                PageSkip = True
            ElseIf arrRecords2(0, i + 1) <> sCertNo Then    'Next Record -> CertNo Change -> Last Cert No
                PageSkip = True
                sCertNo = arrRecords2(0, i + 1)   'First sCertNo
            ElseIf iPageRecCnt >= 11 Then
                PageSkip = True
            Else
                PageSkip = False
            End If
        
            If PageSkip = True Then
                    
                iMod = 11 - iPageRecCnt
                                    
                For iBlank = 1 To iMod
                
                    oRS.AddNew
                    
                        oRS("CERT_NO").Value = arrRecords2(0, i)    'First sCertNo
                        oRS("PROD_NO").Value = Null
                        oRS("STLGRD").Value = Null
                        oRS("PRDT_QNTY").Value = Null
                        oRS("PRDT_WGT").Value = Null
                        oRS("C_RST").Value = Null
                        oRS("SI_RST").Value = Null
                        oRS("MN_RST").Value = Null
                        oRS("P_RST").Value = Null
                        oRS("S_RST").Value = Null
                        oRS("NB_RST").Value = Null
                        oRS("AL_RST").Value = Null
                        oRS("MO_RST").Value = Null
                        oRS("CU_RST").Value = Null
                        oRS("NI_RST").Value = Null
                        oRS("CR_RST").Value = Null
                        oRS("V_RST").Value = Null
                        oRS("TI_RST").Value = Null
                        oRS("CEQ_RST").Value = Null
                        oRS("YP_RST").Value = Null
                        oRS("TS_RST").Value = Null
                        oRS("EL_RST").Value = Null
                        oRS("BEND_RST").Value = Null
                        oRS("UST_GRD_RST").Value = Null
                        oRS("IMPACT_TMP").Value = Null
                        oRS("IMPACT_RST_AVE").Value = Null
                        oRS("TIM_IMPACT_TMP").Value = Null
                        oRS("TIM_IMPACT_RST_AVE").Value = Null
                    
                    oRS.Update   ' Update local Recordset (since adLockBatchOptimistic)
                        
                Next iBlank
                
                iPageRecCnt = 0
                
            End If
            
        Next i
                        
        oRS.MarshalOptions = adMarshalModifiedOnly
        oRS.UpdateBatch
        
    End If
    
    oRS.Close
    
'-----------------------------------------------------------------------------
'AQD0050C_HEAD
'-----------------------------------------------------------------------------
    
    sQuery = "Select * From AQD0050C_HEAD"
    oRS.Open sQuery, oConn, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    If Not oRS.EOF Then
        sQuery = "DELETE FROM AQD0050C_HEAD"
        oConn.Execute sQuery, , adCmdText + adExecuteNoRecords
    End If
            
    If IsEmpty(arrRecords3) = False Then
            
        For i = 0 To UBound(arrRecords3, 2)
        
            oRS.AddNew
            
                oRS("CERT_NO").Value = arrRecords3(0, i)
                oRS("PROD_SPEC_NO").Value = arrRecords3(1, i)
                oRS("STDSPEC_NAME").Value = arrRecords3(2, i)
                oRS("COND_SUPPLY").Value = arrRecords3(3, i)
                oRS("STDSPEC").Value = arrRecords3(4, i)
                oRS("IMPACT_SMP_SIZE").Value = arrRecords3(5, i)
                oRS("PROD_NAME1").Value = arrRecords3(4, i) + " 试样尺寸（mm):" + arrRecords3(5, i) + ")"
                oRS("PROD_NAME2").Value = arrRecords3(4, i) + " Dimensions of test spec:" + arrRecords3(5, i) + ")"
                oRS("SHIP_CMPY_NO").Value = arrRecords3(8, i) + " 控制号：" + arrRecords3(6, i)
                oRS("SHIP_CMPY_NO2").Value = arrRecords3(8, i) + " 控制号：" + arrRecords3(6, i)
                oRS("CERT_RPT_DATE1").Value = Left(arrRecords3(7, i), 4) + "年 " + Mid(arrRecords3(7, i), 5, 2) + "月 " + Right(arrRecords3(7, i), 2) + "日"
                oRS("CERT_RPT_DATE2").Value = Mid(arrRecords3(7, i), 5, 2) + "/" + Right(arrRecords3(7, i), 2) + "/" + Left(arrRecords3(7, i), 4)
                oRS("STD_ORGAN").Value = arrRecords3(8, i)
                oRS("TEST_EMP").Value = arrRecords3(9, i)
                oRS("SUM_CNT").Value = arrRecords3(10, i)
                oRS("SUM_WGT").Value = arrRecords3(11, i)
            
            oRS.Update   ' Update local Recordset (since adLockBatchOptimistic)
        
        Next i
        
        oRS.MarshalOptions = adMarshalModifiedOnly
        oRS.UpdateBatch
        
    End If
    
    oRS.Close
        
'-----------------------------------------------------------------------------
'AQD0050C_DETAIL
'-----------------------------------------------------------------------------
        
    sQuery = "Select * From AQD0050C_DETAIL"
    oRS.Open sQuery, oConn, adOpenStatic, adLockBatchOptimistic, adCmdText
    
    If Not oRS.EOF Then
        sQuery = "DELETE FROM AQD0050C_DETAIL"
        oConn.Execute sQuery, , adCmdText + adExecuteNoRecords
    End If
    
    iPageRecCnt = 0
        
    If IsEmpty(arrRecords4) = False Then
    
        sCertNo = arrRecords4(0, 0)     'First sCertNo
        
        For i = 0 To UBound(arrRecords4, 2)
        
            iPageRecCnt = iPageRecCnt + 1
        
            oRS.AddNew
            
                oRS("CERT_NO").Value = arrRecords4(0, i)
                oRS("PROD_NO").Value = arrRecords4(1, i)
                oRS("PROD_SIZE").Value = arrRecords4(2, i)
                oRS("PRDT_QNTY").Value = arrRecords4(3, i)
                oRS("PRDT_WGT").Value = arrRecords4(4, i)
                oRS("C_RST").Value = arrRecords4(5, i) * 100
                oRS("SI_RST").Value = arrRecords4(6, i) * 100
                oRS("MN_RST").Value = arrRecords4(7, i) * 100
                oRS("P_RST").Value = arrRecords4(8, i) * 1000
                oRS("S_RST").Value = arrRecords4(9, i) * 1000
                oRS("NB_RST").Value = arrRecords4(10, i) * 1000
                oRS("AL_RST").Value = arrRecords4(11, i) * 1000
                oRS("MO_RST").Value = arrRecords4(12, i) * 1000
                oRS("CU_RST").Value = arrRecords4(13, i) * 1000
                oRS("NI_RST").Value = arrRecords4(14, i) * 1000
                oRS("CR_RST").Value = arrRecords4(15, i) * 1000
                oRS("TI_RST").Value = arrRecords4(16, i) * 1000
                oRS("CEQ_RST").Value = arrRecords4(17, i) * 100
                oRS("YP_RST").Value = arrRecords4(18, i)
                oRS("TS_RST").Value = arrRecords4(19, i)
                oRS("EL_RST").Value = arrRecords4(20, i)
                oRS("IMPACT_RST_AVE").Value = arrRecords4(21, i)
                                        
            oRS.Update   ' Update local Recordset (since adLockBatchOptimistic)
            
            RstMode = False
            
            If i = UBound(arrRecords4, 2) Then              'LastRecord
                PageSkip = True
            ElseIf arrRecords4(0, i + 1) <> sCertNo Then    'Next Record -> CertNo Change -> Last Cert No
                PageSkip = True
                sCertNo = arrRecords4(0, i + 1)   'First sCertNo
            Else
                PageSkip = False
            End If
        
            
            If PageSkip = True Then
                                  
                iMod = 12 - iPageRecCnt
                                                  
                For iBlank = 1 To iMod
                
                    oRS.AddNew
                        
                        oRS("CERT_NO").Value = arrRecords4(0, i)
                        oRS("PROD_NO").Value = Null
                        oRS("PROD_SIZE").Value = Null
                        oRS("PRDT_QNTY").Value = Null
                        oRS("PRDT_WGT").Value = Null
                        oRS("C_RST").Value = Null
                        oRS("MN_RST").Value = Null
                        oRS("P_RST").Value = Null
                        oRS("S_RST").Value = Null
                        oRS("SI_RST").Value = Null
                        oRS("NB_RST").Value = Null
                        oRS("AL_RST").Value = Null
                        oRS("MO_RST").Value = Null
                        oRS("CU_RST").Value = Null
                        oRS("NI_RST").Value = Null
                        oRS("CR_RST").Value = Null
                        oRS("TI_RST").Value = Null
                        oRS("CEQ_RST").Value = Null
                        oRS("YP_RST").Value = Null
                        oRS("TS_RST").Value = Null
                        oRS("EL_RST").Value = Null
                        oRS("IMPACT_RST_AVE").Value = Null
                        
                    oRS.Update   ' Update local Recordset (since adLockBatchOptimistic)
                        
                Next iBlank
                                
                iPageRecCnt = 0
                                                
            End If
                    
        Next i
        
        oRS.MarshalOptions = adMarshalModifiedOnly
        oRS.UpdateBatch
    
    End If
    
    oRS.Close
    
    Set oRS = Nothing
    Set oConn = Nothing

Err_Track:
    
End Sub

'------------------------------------------------------------------------------------------
'--------------------------------------- Query Make ---------------------------------------
'------------------------------------------------------------------------------------------
Private Sub funGetQuery()
    
    Dim i As Integer
    Dim sParamC As String
    Dim sParamS As String
    Dim sQuery1 As String
    Dim sQuery2 As String
    
    If cVal.Count = 0 Then
        sQueryHeadC = ""
        sQueryDetailC = ""
    End If
                        
    If sVal.Count = 0 Then
        sQueryHeadS = ""
        sQueryDetailC = ""
    End If
            
    For i = 1 To cVal.Count
         
         If i = 1 And i = cVal.Count Then
            sParamC = "('" + cVal.Item(i) + "')"
         ElseIf i = 1 Then
            sParamC = "('" + cVal.Item(i) + "'"
         ElseIf i = cVal.Count Then
            sParamC = sParamC + ",'" + cVal.Item(i) + "')"
         Else
            sParamC = sParamC + ",'" + cVal.Item(i) + "'"
         End If
    
    Next i
    
    For i = 1 To sVal.Count
         
         If i = 1 And i = sVal.Count Then
            sParamS = "('" + sVal.Item(i) + "')"
         ElseIf i = 1 Then
            sParamS = "('" + sVal.Item(i) + "'"
         ElseIf i = sVal.Count Then
            sParamS = sParamS + ",'" + sVal.Item(i) + "')"
         Else
            sParamS = sParamS + ",'" + sVal.Item(i) + "'"
         End If
    
    Next i
        
    sQueryHeadC = "SELECT CERT_NO , PROD_NAME ,STDSPEC_NAME , PROD_SPEC_NO , GF_CUST_NAME(CUST_CD,'') ,COND_SUPPLY , PROD_SIZE"
    sQueryHeadC = sQueryHeadC + ",IMPACT_SMP_SIZE , QLTY_REC_NO , GF_PONO_FIND(ORD_NO), TRNS_NO , TRAIN_LINE_NAME"
    sQueryHeadC = sQueryHeadC + ",DEST_DETAIL , CERT_RPT_DATE , BEND_DIA , GF_EMPNAMEFIND(TEST_EMP) AS TEST_EMP , GF_EMPNAMEFIND(SHP_EMP) AS SHP_EMP "
    sQueryHeadC = sQueryHeadC + ",AQD0050C.F_SUM_CNT(CERT_NO) AS SUM_CNT, AQD0050C.F_SUM_WGT(CERT_NO) AS SUM_WGT"
    sQueryHeadC = sQueryHeadC + " FROM QP_CERT_HEAD WHERE CERT_NO IN "
    
    sQueryDetailC = "SELECT CERT_NO ,PROD_NO , GF_STLGRD_DETAIL(STLGRD) AS STLGRDNAME , PRDT_QNTY , PRDT_WGT"
    sQueryDetailC = sQueryDetailC + ", C_RST , SI_RST , MN_RST , P_RST , S_RST"
    sQueryDetailC = sQueryDetailC + ", NB_RST , AL_RST , MO_RST , CU_RST , NI_RST "
    sQueryDetailC = sQueryDetailC + ", CR_RST , V_RST , TI_RST , CEQ_RST , YP_RST "
    sQueryDetailC = sQueryDetailC + ", TS_RST , EL_RST , BEND_RST , UST_GRD_RST , IMPACT_TMP "
    sQueryDetailC = sQueryDetailC + ", IMPACT_RST_AVE , TIM_IMPACT_TMP , TIM_IMPACT_RST_AVE "
    sQueryDetailC = sQueryDetailC + " FROM QP_CERT_DETAIL WHERE CERT_NO IN "
    
    sQueryHeadS = "SELECT CERT_NO , PROD_SPEC_NO , STDSPEC_NAME , COND_SUPPLY ,STDSPEC , IMPACT_SMP_SIZE "
    sQueryHeadS = sQueryHeadS + ", SHIP_CMPY_NO , CERT_RPT_DATE , STD_ORGAN , GF_EMPNAMEFIND(TEST_EMP) AS TEST_EMP "
    sQueryHeadS = sQueryHeadS + ", AQD0050C.F_SUM_CNT(CERT_NO) AS SUM_CNT, AQD0050C.F_SUM_WGT(CERT_NO) AS SUM_WGT "
    sQueryHeadS = sQueryHeadS + " FROM QP_CERT_HEAD WHERE CERT_NO IN "
    
    sQueryDetailS = "SELECT CERT_NO ,PROD_NO , PROD_SIZE , PRDT_QNTY , PRDT_WGT"
    sQueryDetailS = sQueryDetailS + ", C_RST , SI_RST , MN_RST , P_RST , S_RST "
    sQueryDetailS = sQueryDetailS + ", NB_RST , AL_RST , MO_RST , CU_RST , NI_RST "
    sQueryDetailS = sQueryDetailS + ", CR_RST , TI_RST , CEQ_RST , YP_RST , TS_RST , EL_RST , IMPACT_RST_AVE "
    sQueryDetailS = sQueryDetailS + " FROM QP_CERT_DETAIL WHERE CERT_NO IN "
    
    sQueryHeadC = sQueryHeadC + sParamC
    sQueryDetailC = sQueryDetailC + sParamC
    sQueryHeadS = sQueryHeadS + sParamS
    sQueryDetailS = sQueryDetailS + sParamS
    
    If cVal.Count = 0 Then
        sQueryHeadC = ""
        sQueryDetailC = ""
    End If
                        
    If sVal.Count = 0 Then
        sQueryHeadS = ""
        sQueryDetailS = ""
    End If
                        
End Sub


'########################################################################################################################
'################################################### REPORT END #########################################################
'########################################################################################################################


'--------------------------------------------------------------------------------------------------------
'------------------------------------------- Local Procedure --------------------------------------------
'--------------------------------------------------------------------------------------------------------

Private Sub subMasterClear()
    txt_CERT_NO.Text = ""
    dtp_fr_date.Text = ""
    dtp_to_date.Text = ""
    txt_CUST_CD.Text = ""
    txt_CUST_NAME.Text = ""
    txt_ORD_NO.Text = ""
    txt_PROD_CD.Text = ""
End Sub

Private Function subCheck() As Boolean

    Dim sMesg As String
    Dim sFrDate As String
    Dim sToDate As String
    Dim sProdCd As String
    Dim sCertNo As String
    Dim sOrdNo As String
    Dim sTrnsNo As String
    Dim sCustCD As String
    
    sProdCd = Trim(txt_PROD_CD.Text)
    sCertNo = Trim(txt_CERT_NO.Text)
    sOrdNo = Trim(txt_ORD_NO.Text)
    sTrnsNo = Trim(txt_TRNS_NO.Text)
    sCustCD = Trim(txt_CUST_CD.Text)
    
    sFrDate = Trim(dtp_fr_date.Text)
    sToDate = Trim(dtp_to_date.Text)
    
    sFrDate = Replace(sFrDate, "_", "")
    sToDate = Replace(sToDate, "_", "")
    
    sFrDate = Replace(sFrDate, "-", "")
    sToDate = Replace(sToDate, "-", "")
    
    If sCertNo = "" Then
        If sFrDate = "" Or sToDate = "" Then
            sMesg = "请完整输入发放日期（开始和结束日期）"
            Call Gp_MsgBoxDisplay(sMesg)
            subCheck = False
            Exit Function
        Else
            If sCustCD = "" And sOrdNo = "" And sProdCd = "" And sTrnsNo = "" Then
                sMesg = "请输入“产品代码”或“提货单号”或“客户代码”或“订单号”中的任意一项"
                Call Gp_MsgBoxDisplay(sMesg)
                subCheck = False
                Exit Function
            End If
        End If
    End If
    
'    If sProdCd = "" And sFrDate <> "" Then
'        sMesg = " 产品代码 , 发放日期 Must input necessarily"
'        Call Gp_MsgBoxDisplay(sMesg)
'        Exit Function
'    End If
'
'    If sProdCd = "" And sToDate <> "" Then
'        sMesg = " 产品代码 , 发放日期 Must input necessarily"
'        Call Gp_MsgBoxDisplay(sMesg)
'        Exit Function
'    End If
'
'    If sProdCd <> "" And sFrDate = "" Then
'        sMesg = " 产品代码 , 发放日期 Must input necessarily"
'        Call Gp_MsgBoxDisplay(sMesg)
'        Exit Function
'    End If

'    If sProdCd <> "" And sToDate = "" Then
'        sMesg = " 产品代码 , 发放日期 Must input necessarily"
'        Call Gp_MsgBoxDisplay(sMesg)
'        Exit Function
'    End If
        
    
    subCheck = True

End Function


Private Sub txt_CERT_NO_Change()
    Call Gf_Control_text_Up(txt_CERT_NO)
End Sub

'########################################################################################################################
'################################################### REPORT END #########################################################
'########################################################################################################################
Private Sub txt_PROD_CD_Change()
    Call Gf_Control_text_Up(txt_PROD_CD)
End Sub





