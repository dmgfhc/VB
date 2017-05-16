VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQA0130C 
   Caption         =   "企标成分信息输入 - AQA0130C"
   ClientHeight    =   9090
   ClientLeft      =   195
   ClientTop       =   1050
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_STL_LQUID_TEMP 
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
      Left            =   12330
      MaxLength       =   11
      TabIndex        =   8
      Tag             =   "钢种"
      Top             =   120
      Width           =   1305
   End
   Begin VB.TextBox txt_STLGRD_FL 
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
      Left            =   9600
      MaxLength       =   11
      TabIndex        =   7
      Tag             =   "钢种"
      Top             =   120
      Width           =   1305
   End
   Begin Threed.SSCommand SSBtn_AQA0410C 
      Height          =   315
      Left            =   8280
      TabIndex        =   6
      Top             =   510
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   3
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "输入成分修约标准"
      ButtonStyle     =   1
   End
   Begin VB.TextBox txt_STLGRD_GRP_NAME 
      Enabled         =   0   'False
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
      Left            =   3555
      TabIndex        =   5
      Top             =   510
      Width           =   4635
   End
   Begin VB.TextBox txt_INS_EMP 
      Height          =   285
      Left            =   11175
      TabIndex        =   3
      Top             =   495
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_STLGRD_GRP 
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
      Left            =   1515
      TabIndex        =   2
      Top             =   510
      Width           =   2025
   End
   Begin VB.TextBox txt_STLGRD_DETAIL 
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
      Left            =   4950
      TabIndex        =   1
      Top             =   120
      Width           =   3225
   End
   Begin VB.TextBox txt_STLGRD 
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
      Left            =   1515
      MaxLength       =   11
      TabIndex        =   0
      Tag             =   "钢种"
      Top             =   120
      Width           =   2025
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   165
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "钢种"
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
      Index           =   0
      Left            =   3600
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "说明"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Index           =   1
      Left            =   165
      Top             =   510
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "钢种组"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8115
      Left            =   135
      TabIndex        =   4
      Top             =   1035
      Width           =   15060
      _Version        =   393216
      _ExtentX        =   26564
      _ExtentY        =   14314
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   22
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0130C.frx":0000
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   8250
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "钢种状态"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   10980
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "液相线温度"
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   135
      X2              =   15165
      Y1              =   945
      Y2              =   945
   End
End
Attribute VB_Name = "AQA0130C"
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
'-- Program Name      企标成分信息输入（钢种说明）
'-- Program ID        AQA0130C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       企标成分信息输入
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
Dim lBlkrow2 As Long

Dim ArrayRecords As Variant

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Hsheet"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_STLGRD, "p", "n", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STLGRD_DETAIL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_STLGRD_GRP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_STLGRD_GRP_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_INS_EMP, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_STLGRD_FL, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_STL_LQUID_TEMP, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
     Mc1.Add Item:="AQA0130C.P_MODIFY", Key:="P-M"
     Mc1.Add Item:="AQA0130C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0130C.P_SMODIFY", Key:="P-M"
    Sc1.Add Item:="AQA0130C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AQA0130C.P_SONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=2, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0


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
        
        Case "txt_STLGRD"
            sCode = "STLGRD"
            Set oCodeName = txt_STLGRD_DETAIL
        
        Case "txt_STLGRD_GRP"
            sCode = "Q0048"
            Set oCodeName = txt_STLGRD_GRP_NAME
    
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call subMenuHide
     

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Dim x As Boolean

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)
    
        
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call subMenuHide
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 2)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 8)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 14)
    
    Screen.MousePointer = vbDefault
    
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    ArrayRecords = GF_GetChemicalCode
    
    ss1.TextTip = SS_TEXTTIP_FLOATINGFOCUSONLY
    ss1.TextTipDelay = 250
    x = ss1.SetTextTipAppearance("宋体", "11", False, False, &HFFFF&, &H800000)
   

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
    Call subMenuHide
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
    
    Call GS_SetChemicalSpreadLineColor(ss1, "0713")
      
End Sub

Public Sub Spread_Forzens_Setting()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Active_Spread.SetFocus
    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call subMenuHide
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        txt_STLGRD_DETAIL.Enabled = True
        txt_STLGRD_GRP.Enabled = True
        txt_STLGRD_GRP_NAME.Enabled = True
        pControl(1).SetFocus
    End If
    
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg  As String
    Dim sQuery As String
    Dim i As Integer

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        If Gf_Ms_Refer(M_CN1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
         Call Gp_Ms_ControlLock(Mc1("pControl"), True)
'         txt_STLGRD_DETAIL.Enabled = False
'         txt_STLGRD_GRP.Enabled = False
'         txt_STLGRD_GRP_NAME.Enabled = False
        End If
    Else
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        txt_STLGRD_DETAIL.Enabled = True
        txt_STLGRD_GRP.Enabled = True
        txt_STLGRD_GRP_NAME.Enabled = True
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call subMenuHide
    Call GS_SetChemicalLength(ss1, ArrayRecords, "020814", "3")
    Call GS_SetChemicalSpreadLineColor(ss1, "0713")
    Exit Sub
                    
Refer_Err:

End Sub

Public Sub Form_Pro()
         
If Len(Trim(txt_STLGRD_DETAIL.Text)) = 0 Then
 Call Gp_MsgBoxDisplay("请输入钢种说明")
 Exit Sub
End If

If Len(Trim(txt_STLGRD_GRP.Text)) = 0 Then
 Call Gp_MsgBoxDisplay("请输入钢种组")
 Exit Sub
End If

    If Gf_Mc_Authority(sAuthority, Mc1, Proc_Sc("Sc")) Then
        txt_INS_EMP.Text = sUserID
            
        If Gf_Ms_Process(M_CN1, Mc1, sAuthority) Or Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call subMenuHide
            Call GS_SetChemicalLength(ss1, ArrayRecords, "020814", "3")
        End If
    End If
         
         
    
End Sub
Public Sub form_Cpy()
  Call Gf_Ms_Copy(Mc1)
End Sub
Public Sub form_Pst()

    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19, "P")
    End If
    
End Sub
Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub
Public Sub Master_Pst()

    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19, "P")
    End If
   
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)

End Sub

Public Sub Form_Del()

    If Gf_Ms_Del(M_CN1, Mc1) Then
        Call Gf_Sp_Cls(Proc_Sc("Sc"))
        txt_STLGRD.Text = ""
        txt_STLGRD_DETAIL.Text = ""
        txt_STLGRD_GRP.Text = ""
        txt_STLGRD_GRP_NAME.Text = ""
        Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
        Call subMenuHide
    End If
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
   
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
End Sub

Private Sub ss1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    Dim ChemCode As String
    
    With ss1
     .Col = Col - 4: ChemCode = .Text
        If ChemCode <> "Ceq" Then
            .Col = Col
            .Row = Row
            .Text = ""
'        Else
'            .Col = Col
'            .Row = Row
'            Call GF_GetCeqValue(ss1, Me.Name, "3")
        End If
    End With
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 19)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol

        Case 2, 8, 14

            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss1

                DD.sWitch = "SP"
                DD.rControl.Add Item:=ss1.ActiveCol
                DD.nameType = "2"

                Call GF_CHEM_SEQ(M_CN1, KeyCode)
                
                Call GS_SetChemicalLength(ss1, ArrayRecords, "020814", "3")
        
            End If

    End Select
    
End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Call GP_ChemCode_RowHeader_Clear("AQA0130C", ss1, NewRow, NewCol)
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

'----------------------------------------------------
'ss1_TextTipFetch
'----------------------------------------------------
Private Sub ss1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
    'TipText = ss1.CellTag
    'ShowTip = True
    Dim sTip As String
    ShowTip = True
    'With ss1
        
    sTip = GF_GetCellMaxLength(ss1, ss1.ActiveRow, ss1.ActiveCol)
    TipText = sTip
    'End With
End Sub

Private Sub subMenuHide()
    
    With MDIMain.MenuTool

        '.Buttons(12).Enabled = True                    'Copy
        .Buttons(11).ButtonMenus(1).Enabled = True      'All Copy
        .Buttons(11).ButtonMenus(2).Enabled = False     'Master Copy
        .Buttons(11).ButtonMenus(3).Enabled = False     'Spread Copy
                    
        '.Buttons(12).Enabled = True                    'Paste
        .Buttons(12).ButtonMenus(1).Enabled = True      'All Paste
        .Buttons(12).ButtonMenus(2).Enabled = False     'Master Paste
        .Buttons(12).ButtonMenus(3).Enabled = False     'Spread Paste
        
    End With

End Sub

Private Sub SSBtn_AQA0410C_Click()
    AQA0410C.txt_STLGRD.Text = Me.txt_STLGRD.Text
    AQA0410C.txt_STLGRD_DETAIL.Text = Me.txt_STLGRD_DETAIL.Text
    AQA0410C.Show
    AQA0410C.SetFocus
End Sub

Private Sub txt_STLGRD_GRP_Change()
    If Trim(txt_STLGRD_GRP.Text) = "" Then
        txt_STLGRD_GRP_NAME.Text = ""
    End If
End Sub
