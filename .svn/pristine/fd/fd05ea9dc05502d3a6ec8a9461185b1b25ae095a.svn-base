VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQF0010C 
   Caption         =   "板坯取样标准录入 - AQF0010C"
   ClientHeight    =   9090
   ClientLeft      =   195
   ClientTop       =   1050
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
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
      Left            =   1830
      TabIndex        =   4
      Top             =   510
      Width           =   4845
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
      MaxLength       =   1
      TabIndex        =   2
      Top             =   510
      Width           =   285
   End
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
      Left            =   2865
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   6195
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
      Width           =   1290
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
      TabIndex        =   3
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
      MaxCols         =   18
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQF0010C.frx":0000
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   135
      X2              =   15165
      Y1              =   945
      Y2              =   945
   End
End
Attribute VB_Name = "AQF0010C"
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
'-- Program Name      板坯取样标准录入
'-- Program ID        AQF0010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2006.01.10
'-- Description       板坯取样及试验标准录入
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
Dim lBlkrow2 As Long

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Hsheet"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_STLGRD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STLGRD_DETAIL, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_STLGRD_GRP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_STLGRD_GRP_NAME, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"

     Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AQF0010C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:="AQF0010C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AQF0010C.P_SONEROW", Key:="P-O"
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

    'Call Gp_Sp_ColHidden(ss1, 11, True)
    Call Gp_Sp_ColHidden(ss1, 13, True)
    Call Gp_Sp_ColHidden(ss1, 16, True)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 1)
    'Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 3)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 4)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 5)

    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"

 
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    If KeyCode = vbKeyF4 Then
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
    End If
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
        
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call subMenuHide
    
    Call Gp_Ms_Cls(Mc1("pControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault
    
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
  
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
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    Call subMenuHide
    
End Sub
Public Sub Spread_Del()
    
    'Call GP_SET_CELL_VALUE(ss1, ss1.Row, 0, "Delete")
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
          
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
        Call Gp_Ms_Cls(Mc1("pControl"))
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


    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
            
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
         Call Gp_Ms_ControlLock(Mc1("rControl"), True)
    Else
        Call Gp_Ms_ControlLock(Mc1("rControl"), False)

    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Call subMenuHide
    Exit Sub
                    
Refer_Err:

End Sub

Public Sub Form_Pro()
         
    'If Gf_Mc_Authority(sAuthority, Mc1, Proc_Sc("Sc")) Then
            
        If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call subMenuHide
        End If
    'End If
End Sub


Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 13)

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


Private Sub ss1_EditChange(ByVal Col As Long, ByVal Row As Long)
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), True)
                
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 16)
         
        If (Col = 4 Or Col = 5) And Row > 0 Then
            ss1.Col = Col: ss1.Row = Row
            ss1.Text = UCase(ss1.Text)
        End If

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    
'    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
                
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 16)
         
        If (Col = 4 Or Col = 5) And Row > 0 Then
            ss1.Col = Col: ss1.Row = Row
            ss1.Text = UCase(ss1.Text)
        End If
'    End If

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim str_orgin As String
    Dim sQuery As String
    Dim strSteel_GRD As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyF4 Then
        With ss1
                .Col = .ActiveCol
                .Row = .ActiveRow
                If .ActiveCol = 1 Then
                
                    str_orgin = .Text
                    DD.nameType = "1"
                    DD.sWitch = "MS"
                    .Text = ""
                    DD.rControl.Add Item:=ss1
                    
                    Call Gf_Stlgrd_DD(M_CN1, KeyCode)
                    
                    If Len(Trim(.Text)) > 0 Then
                       strSteel_GRD = .Text
                       sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(.Text) + "'"
                       .Col = .Col + 1
                       .Text = Gf_FloatFind(M_CN1, sQuery)
                       
                       
                        .Col = 3
                        .Text = ""
                        sQuery = "Select STLGRD_GRP From QP_NISCO_CHMC Where STLGRD =  '" + Trim(strSteel_GRD) + "'"
                    
                        .Text = Gf_FloatFind(M_CN1, sQuery)
                        .Col = .Col - 2
                    Else
                        .Text = str_orgin
                    End If
                    Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), True)
                ElseIf .ActiveCol = 4 Then
                
                    str_orgin = .Text
                    .Text = ""
                    
                    DD.sWitch = "MS"
                    DD.sKey = "Q0062"
                    DD.rControl.Add Item:=ss1
                    DD.nameType = "2"
                    
                    Call Gf_Common_DD(M_CN1, KeyCode)
                    Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), True)
                ElseIf .ActiveCol = 5 Then
                    str_orgin = .Text
                    .Text = ""
                    
                    DD.sWitch = "MS"
                    DD.sKey = "Q0042"
                    DD.rControl.Add Item:=ss1
                    DD.nameType = "2"
                    
                    Call Gf_Common_DD(M_CN1, KeyCode)
                    Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), True)
                End If
        End With
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
        .Buttons(11).Enabled = False                    'Copy
        .Buttons(12).Enabled = False                    'Paste
    End With

End Sub

Private Sub txt_STLGRD_Change()
    Dim sQuery As String
    If Len(Trim(txt_STLGRD.Text)) = 11 Then
        sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(txt_STLGRD.Text) + "'"
        txt_STLGRD_DETAIL.Text = Gf_FloatFind(M_CN1, sQuery)
    Else
        txt_STLGRD_DETAIL.Text = ""
    End If
End Sub


Private Sub txt_STLGRD_GRP_Change()
    Dim sQuery As String
    If Trim(txt_STLGRD_GRP.Text) = "" Then
        txt_STLGRD_GRP_NAME.Text = ""
    Else
        txt_STLGRD_GRP.Text = UCase(txt_STLGRD_GRP.Text)
        sQuery = "SELECT CD_SHORT_NAME FROM ZP_CD WHERE  CD_MANA_NO ='Q0048' AND CD = '" + Trim(txt_STLGRD_GRP.Text) + "'"
        txt_STLGRD_GRP_NAME.Text = Gf_FloatFind(M_CN1, sQuery)
    End If
End Sub
