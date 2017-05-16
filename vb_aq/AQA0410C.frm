VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQA0410C 
   Caption         =   "产品成份修约标准管理_AQA0410C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_STLGRD_DETAIL 
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
      Left            =   5055
      TabIndex        =   2
      Top             =   95
      Width           =   3435
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
      Left            =   1440
      MaxLength       =   11
      TabIndex        =   1
      Tag             =   "钢种"
      Top             =   90
      Width           =   1425
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9060
      Left            =   135
      TabIndex        =   0
      Top             =   555
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   15981
      _Version        =   196609
      PaneTree        =   "AQA0410C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   5280
         Left            =   30
         TabIndex        =   3
         Top             =   3750
         Width           =   15030
         _Version        =   393216
         _ExtentX        =   26511
         _ExtentY        =   9313
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
         MaxCols         =   17
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQA0410C.frx":0072
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3630
         Left            =   3885
         TabIndex        =   4
         Top             =   30
         Width           =   11175
         _Version        =   393216
         _ExtentX        =   19711
         _ExtentY        =   6403
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
         MaxCols         =   12
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQA0410C.frx":0615
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   3630
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   3765
         _Version        =   393216
         _ExtentX        =   6641
         _ExtentY        =   6403
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
         MaxCols         =   3
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQA0410C.frx":0AD8
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   135
      Top             =   90
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
      Left            =   3780
      Top             =   90
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
End
Attribute VB_Name = "AQA0410C"
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
'-- Program Name      产品成份修约标准管理_AQA0410C）
'-- Program ID        AQA0410C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2008.1.7
'-- Description       产品成份修约标准管理_AQA0410C
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

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long

Dim ArrayRecords As Variant

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_STLGRD, "p", " ", " ", "i", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STLGRD_DETAIL, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'        Call Gp_Ms_Collection(txt_STLGRD_GRP, " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'   Call Gp_Ms_Collection(txt_STLGRD_GRP_NAME, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'           Call Gp_Ms_Collection(txt_INS_EMP, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
'     Mc1.Add Item:="AQA0410C.P_MODIFY", Key:="P-M"
'     Mc1.Add Item:="AQA0410C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, "P", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, "P", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0410C.P_SMODIFY", Key:="P-M"
    Sc1.Add Item:="AQA0410C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AQA0410C.P_SONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"


    
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", "", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", "", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", "", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
'    Call Gp_Sp_Collection(ss2, 13, " ", " ", "", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
'    Call Gp_Sp_Collection(ss2, 14, " ", " ", "", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQA0410C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=2, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    
    Call Gp_Sp_Collection(ss3, 1, "p", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
'    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
'    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AQA0410C.P_SREFER3", Key:="P-R"
    sc3.Add Item:=pColumn3, Key:="pColumn"
    sc3.Add Item:=nColumn3, Key:="nColumn"
    sc3.Add Item:=aColumn3, Key:="aColumn"
    sc3.Add Item:=mColumn3, Key:="mColumn"
    sc3.Add Item:=iColumn3, Key:="iColumn"
    sc3.Add Item:=lColumn3, Key:="lColumn"
    sc3.Add Item:=3, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=sc3, Key:="Sc3"
     
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name, False)
    
        
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call subMenuHide
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
       
    Screen.MousePointer = vbDefault
    
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    ArrayRecords = GF_GetChemicalCode
    
    ss1.TextTip = SS_TEXTTIP_FLOATINGFOCUSONLY
    ss1.TextTipDelay = 250
    x = ss1.SetTextTipAppearance("宋体", "11", False, False, &HFFFF&, &H800000)
   

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

     If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc2")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc3")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
    
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
    Set sc2 = Nothing
    Set sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    Call subMenuHide
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc1"))
    
    Call GS_SetChemicalSpreadLineColor(ss1, "0407")
      
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
    
    If Gf_Sp_Cls(Proc_Sc("Sc1")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call subMenuHide
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        pControl(1).SetFocus
    End If
    
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg  As String
    Dim sQuery As String
    Dim i As Integer

    If Gf_Sp_ProceExist(Proc_Sc("Sc1").Item("Spread")) Or Gf_Sp_ProceExist(Proc_Sc("Sc2").Item("Spread")) Or Gf_Sp_ProceExist(Proc_Sc("Sc3").Item("Spread")) Then Exit Sub

            
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
        If Gf_Sp_Refer(M_CN1, sc2, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
           If Gf_Sp_Refer(M_CN1, sc3, Mc1, Mc1("nControl"), Mc1("mControl"), False) Then
            Call Gp_Ms_ControlLock(Mc1("pControl"), True)
            End If
        End If
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call subMenuHide
    Call GS_SetChemicalSpreadLineColor(ss1, "0407")
    Call GS_SetChemicalSpreadLineColor(ss2, "0407")
    Call GS_SetChemicalSpreadLineColor(ss3, "0407")
    Exit Sub
                    
Refer_Err:

End Sub

Public Sub Form_Pro()

'    If Gf_Mc_Authority(sAuthority, Mc1, Proc_Sc("Sc1")) Then
        Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), 10, "I")
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc1"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call subMenuHide
            Call GS_SetChemicalLength(ss1, ArrayRecords, "020814", "3")
        End If
'    End If
    
End Sub

Public Sub form_Cpy()
  Call Gf_Ms_Copy(Mc1)
End Sub

Public Sub form_Pst()

    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc1")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), 10, "P")
    End If
    
End Sub

Public Sub Master_Cpy()

    Call Gf_Ms_Copy(Mc1)
    
End Sub

Public Sub Master_Pst()

    If Gf_Ms_FormPaste(Mc1, Proc_Sc("Sc1")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), 10, "P")
    End If
   
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

'Public Sub Form_Ins()
'
''    Call Gp_Sp_Ins(Proc_Sc("Sc"))
''    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
'
'End Sub

Public Sub Form_Del()

    If Gf_Ms_Del(M_CN1, Mc1) Then
        Call Gf_Sp_Cls(Proc_Sc("Sc1"))
        txt_STLGRD.Text = ""
        txt_STLGRD_DETAIL.Text = ""
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

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc1")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), 10)
    End If
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

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



'Private Sub txt_STLGRD_GRP_Change()
'    If Trim(txt_STLGRD_GRP.Text) = "" Then
'        txt_STLGRD_GRP_NAME.Text = ""
'    End If
'End Sub

Private Sub ss3_DblClick(ByVal Col As Long, ByVal Row As Long)

        If Row > 0 Then
        
                    ss3.Row = Row
                    ss3.Col = 1
                    txt_STLGRD.Text = ss3.Text
                    
                    ss3.Col = 2
                    txt_STLGRD_DETAIL.Text = ss3.Text
        End If
        
        Call Form_Ref

End Sub


