VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQD0012C 
   Caption         =   "特钢质量证明书_AQD0012C"
   ClientHeight    =   9090
   ClientLeft      =   165
   ClientTop       =   1605
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15420
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_PROCCD 
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
      Left            =   5160
      MaxLength       =   14
      TabIndex        =   12
      Top             =   720
      Width           =   1635
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删  除"
      Height          =   375
      Left            =   11400
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txt_packet_no 
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
      Left            =   1665
      MaxLength       =   14
      TabIndex        =   10
      Top             =   720
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "导  出"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txt_SAVE_DIR 
      Enabled         =   0   'False
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
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   3435
   End
   Begin VB.TextBox txt_STDSPEC 
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
      Left            =   5150
      MaxLength       =   14
      TabIndex        =   4
      Top             =   240
      Width           =   1635
   End
   Begin VB.TextBox txt_PLT_NAME 
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
      Left            =   2050
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1100
   End
   Begin VB.TextBox txt_PLT 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin InDate.UDate dtp_DATE_1 
      Height          =   315
      Left            =   8880
      TabIndex        =   0
      Top             =   255
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
   Begin InDate.UDate dtp_DATE_2 
      Height          =   315
      Left            =   10680
      TabIndex        =   1
      Top             =   240
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   7320
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "生产日期"
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
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   120
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "生产工厂"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   3600
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "产品标准"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   6525
      Left            =   -240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   15120
      _Version        =   393216
      _ExtentX        =   26670
      _ExtentY        =   11509
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
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
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQD0012C.frx":0000
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   120
      Top             =   1200
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "保存目录"
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
      ForeColor       =   0
   End
   Begin Threed.SSCommand ssc_DIR_FIND 
      Height          =   345
      Left            =   5160
      TabIndex        =   8
      Top             =   1200
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      _Version        =   196609
      PictureFrames   =   1
      Picture         =   "AQD0012C.frx":05B8
      ButtonStyle     =   1
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   2
      Left            =   120
      Top             =   720
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "包  号"
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
      Index           =   3
      Left            =   3600
      Top             =   720
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "进程代码"
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
   Begin VB.Label Label1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10440
      TabIndex        =   5
      Top             =   240
      Width           =   135
   End
End
Attribute VB_Name = "AQD0012C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public sPLT_Authority As String     'Active User Plant Authority Setting

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


Dim sQuery      As String
Dim sErrMsg     As String
Dim sDate       As String

Dim AdoRs       As adodb.Recordset


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_STDSPEC, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(dtp_DATE_1, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(dtp_DATE_2, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_packet_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_PROCCD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
    
    
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
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQD0012C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQD0012C.P_MODIFY", Key:="P-M"
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
    
    Call Gp_Sp_ColHidden(ss1, 1, True)

End Sub

Private Sub Command1_Click()
  Dim sSave_Path  As String
  Dim i           As Integer
  Dim sPlate_no   As String
  Dim sHEAT_NO    As String
  Dim sPacket_No   As String
  
   
  sSave_Path = Trim(txt_SAVE_DIR)

  Set AdoRs = New adodb.Recordset

    sQuery = "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset

    sDate = AdoRs.Fields(0)

    AdoRs.Close
    Set AdoRs = Nothing

    sPacket_No = Trim(txt_packet_no.Text)
    
    If sPacket_No <> "" Then
      sErrMsg = basCertPrn.funGetQuery_SPCL(sPacket_No, sSave_Path)
      If sErrMsg = "" Then
        MsgBox ("质保书导出完成！")
      End If
    Else
      MsgBox ("请选择待产制质保书包号！")
      Exit Sub
    End If
      
    
End Sub

Private Sub Command2_Click()
    Dim sPacket_No
    Dim sQuery As String
    Dim AdoRs As adodb.Recordset

    Set AdoRs = New adodb.Recordset
    sQuery = "{CALL AQD0012C.P_CERT_DEL('" + sPacket_No + "')}"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    AdoRs.Close
    Set AdoRs = Nothing
    
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call GP_ROW_BACKCOLOR(ss1)
    
    'Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    

    Screen.MousePointer = vbDefault
    
    dtp_DATE_1.Text = Now
    dtp_DATE_2.Text = Now
    
    sAuthority = "1111"
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)

        
    End If
    txt_PLT.SetFocus
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    
    Exit Sub

Refer_Err:

End Sub
Public Sub Form_Pro()
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
      Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
      Call Form_Ref
    End If

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
    
    Proc_Sc("Sc")("Spread").EditMode = True
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 3
    
    txt_packet_no.Text = ss1.Text


End Sub


Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)
  Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
  ss1.Row = Row
  ss1.Col = 0
  ss1.Text = "Update"
  ss1.Col = 1
  ss1.Text = 1
  
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

Private Sub ssc_DIR_FIND_Click()
    Load Form_DIR_SELECT
    Form_DIR_SELECT.Show 1
    txt_SAVE_DIR.Text = sEXLSavePATH
End Sub

Private Sub txt_PLT_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_PLT

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        Case "txt_PLT"                  '工厂
            sCode = "C0001"
            Set oCodeName = txt_PLT_NAME
            
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub
