VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQA0200C 
   Caption         =   "标准订单用途查询_AQA0200C"
   ClientHeight    =   9090
   ClientLeft      =   -15
   ClientTop       =   1605
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_PROD_KND_NAME 
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
      Left            =   2160
      TabIndex        =   9
      Tag             =   "PLT"
      Top             =   165
      Width           =   1125
   End
   Begin VB.TextBox txt_STDSPEC2 
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
      Left            =   11805
      TabIndex        =   8
      Tag             =   "PLT"
      Top             =   180
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_PROD_KND2 
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
      Left            =   11235
      TabIndex        =   7
      Tag             =   "PLT"
      Top             =   180
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txt_PROD_KND 
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
      Left            =   1560
      TabIndex        =   6
      Tag             =   "PLT"
      Top             =   165
      Width           =   585
   End
   Begin VB.TextBox txt_STDSPEC_NAME 
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
      Left            =   7260
      TabIndex        =   5
      Tag             =   "PLT"
      Top             =   165
      Width           =   945
   End
   Begin VB.TextBox txt_STDSPEC 
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
      Left            =   4800
      MaxLength       =   18
      TabIndex        =   4
      Tag             =   "PLT"
      Top             =   165
      Width           =   2445
   End
   Begin VB.TextBox txt_THK_MIN 
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
      Left            =   12915
      TabIndex        =   3
      Tag             =   "PLT"
      Top             =   180
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox txt_THK_MAX 
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
      Left            =   13620
      TabIndex        =   2
      Tag             =   "PLT"
      Top             =   180
      Visible         =   0   'False
      Width           =   660
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8160
      Left            =   165
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   8880
      _Version        =   393216
      _ExtentX        =   15663
      _ExtentY        =   14393
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0200C.frx":0000
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   225
      Top             =   165
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "品种"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   3585
      Top             =   165
      Width           =   1140
      _ExtentX        =   2011
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
   Begin FPSpread.vaSpread ss2 
      Height          =   8160
      Left            =   9105
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   945
      Width           =   6105
      _Version        =   393216
      _ExtentX        =   10769
      _ExtentY        =   14393
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQA0200C.frx":0432
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   285
      Left            =   210
      TabIndex        =   10
      Top             =   615
      Width           =   810
      _ExtentX        =   1429
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   255
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "标准"
      Value           =   1
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   285
      Left            =   9150
      TabIndex        =   11
      Top             =   615
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   8421504
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "订单用途"
   End
End
Attribute VB_Name = "AQA0200C"
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
'-- Program Name      质量设计键输入
'-- Program ID        AQA0200C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CHU KYO SU
'-- Coder             CHU KYO SU
'-- Date              2003.8.20
'-- Description       质量设计键输入
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

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

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
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim lCopyRow As Long                'Copy Row

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

'----------------------------------------------------------------------------------------------------------------------------------------------------

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_PROD_KND, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                  
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'----------------------------------------------------------------------------------------------------------------------------------------------------
    
   Call Gp_Ms_Collection(txt_PROD_KND2, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    Call Gp_Ms_Collection(txt_STDSPEC2, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_THK_MIN, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_THK_MAX, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER2 Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
'----------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQA0200C.P_DELETE", Key:="P-M"
    Sc1.Add Item:="AQA0200C.P_REFER1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxRows, Key:="Last"
    
'----------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, "p", "n", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQA0200C.P_MODIFY", Key:="P-M"
    sc2.Add Item:="AQA0200C.P_REFER2", Key:="P-R"
    sc2.Add Item:="AQA0200C.P_ONEROW", Key:="P-O"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxRows, Key:="Last"
    
'----------------------------------------------------------------------------------------------------------------------------------------------------
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    Proc_Sc.Add Item:=sc2, Key:="Sc"
    
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

        Case "txt_PROD_KND"             '品种
            sCode = "Q0001"
            Set oCodeName = txt_PROD_KND_NAME

        Case "txt_STDSPEC"              '标准编号
            sCode = "STDSPEC"
            Set oCodeName = txt_STDSPEC_NAME
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
    
    'Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    
    Call Gp_Sp_Setting(ss1, False)
    Call Gp_Sp_Setting(ss2, False)
    
    Call GP_ROW_BACKCOLOR(ss1)
    Call GP_ROW_BACKCOLOR(ss2)
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(ss1, 1)
    Call Gp_Sp_HdColColor(ss1, 2)
    Call Gp_Sp_HdColColor(ss2, 5)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
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
    Set Mc2 = Nothing
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    If Chk_ss1.Value = -1 Then
        Call GP_SELECT_ROW(ss1, ss1.ActiveRow)
        Call GP_ROW_CANCEL(Sc1)
    Else
        Call GP_SELECT_ROW(ss2, ss2.ActiveRow)
        Call GP_ROW_CANCEL(sc2)
    End If
          
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
        If Gf_Sp_Cls(Sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
          ' rControl(1).SetFocus
        End If
    End If
    
    txt_STDSPEC.Text = ""
    txt_STDSPEC_NAME.Text = ""
    txt_PROD_KND.Text = ""
    txt_PROD_KND_NAME.Text = ""
    txt_THK_MIN.Text = ""
    txt_THK_MAX.Text = ""
    txt_STDSPEC2.Text = ""
    txt_PROD_KND2.Text = ""
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub

   If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False) = True Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call ss1_Click(1, 1)
   End If
    
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
Dim OldRow As Long
Dim OldMaxRow As Long
    
OldRow = ss1.ActiveRow
OldMaxRow = ss1.MaxRows
    
    If Chk_ss1.Value = -1 Then
        If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then
            Call Gf_Sp_Process(M_CN1, sc2, Mc2)
            Call Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False)
            Call ss1_Click(1, ss1.ActiveRow)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    Else
        If Gf_Sp_Process(M_CN1, sc2, Mc2) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           Call Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl"), False)
           With ss1
            If .ActiveRow <> OldRow Then
                If .MaxRows <> OldMaxRow Then
                    If .MaxRows > OldRow + 1 Then
                       .Row = OldRow + 1
                    Else
                        If .MaxRows > OldRow - 1 Then
                            .Row = OldRow - 1
                        Else
                            .Row = .MaxRows - 1
                       End If
                    End If
                Else
                    .Row = OldRow
                End If
            End If
            Call Spread_to_Master(ss1, ss1.Row)
            Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc1("nControl"), Mc1("mControl"), False)
            Call GP_SELECT_ROW(ss1, .Row)
           End With
        End If
    End If
    
End Sub

Public Sub Form_Ins()
        
    If Chk_ss1.Value = -1 Then

        Call Gp_Sp_Ins(Sc1)
        Call Spread_to_Master(ss1, ss1.ActiveRow)
        Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc1("nControl"), Mc1("mControl"), False)
    Else
        If txt_PROD_KND2.Text = "" Then Exit Sub
        Call Gp_Sp_Ins(sc2)
        Call Gp_Sp_InAuthority(sc2, 8)
        
        With ss2
        
            .Row = .ActiveRow
            
            .Col = 1: .Text = txt_PROD_KND2.Text
            .Col = 2: .Text = txt_STDSPEC2.Text
            .Col = 3: .Text = txt_THK_MIN.Text
            .Col = 4: .Text = txt_THK_MAX.Text
            
        End With
    
    End If
        
End Sub

Public Sub Spread_Cpy()
    
    If Chk_ss1.Value = -1 Then
        lCopyRow = ss1.ActiveRow
    Else
        lCopyRow = ss2.ActiveRow
    End If
    
End Sub

Public Sub Spread_Pst()
    
    If Chk_ss1.Value = -1 Then
        Call GP_ROW_PASTE(Sc1, lCopyRow)
        Call Gp_Sp_InAuthority(Sc1, 7)
    Else
        Call GP_ROW_PASTE(sc2, lCopyRow)
        Call Gp_Sp_InAuthority(sc2, 8)
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

Public Sub Spread_Del()
        
    If Chk_ss1.Value = -1 Then
        Call GP_SET_CELL_VALUE(ss1, ss1.ActiveRow, 0, "Delete")
    Else
        Call GP_SET_CELL_VALUE(ss2, ss2.ActiveRow, 0, "Delete")
    End If

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Call Gp_Sp_Sort(ss1, Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ss1.MaxRows < 1 Or Row = 0 Then
        Exit Sub
    End If
    
    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
'     With ss1
'
'        .Row = Row
'
'        .Col = 1: txt_PROD_KND2.Text = .Text
'        .Col = 2: txt_STDSPEC2.Text = .Text
'        .Col = 3: txt_THK_MIN.Text = .Text
'        .Col = 4: txt_THK_MAX.Text = .Text
'
'    End With
    
    Call Spread_to_Master(ss1, ss1.ActiveRow)
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc1("nControl"), Mc1("mControl"), False)
    Call GP_SELECT_ROW(ss1, Row)

End Sub



Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol
    
        Case 1
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "Q0001"
                DD.rControl.Add Item:=1
                
                ss1.Row = ss1.ActiveRow
                
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            End If
            
         Case 2
         
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
               ' DD.sKey = "Q0001"
                DD.rControl.Add Item:=2
                
                ss1.Row = ss1.ActiveRow
                
                DD.nameType = "2"
                
                Call Gf_StdSPEC_DD(M_CN1, KeyCode)
                
            End If
            
    End Select
End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Call Spread_to_Master(ss1, ss1.ActiveRow)
End Sub

Private Sub ss1_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'    With ss1
        
'        .Row = NewRow
'       .Row = .ActiveRow
'
'        .Col = 1: txt_PROD_KND2.Text = .Text
'        .Col = 2: txt_STDSPEC2.Text = .Text
'        .Col = 3: txt_THK_MIN.Text = .Text
'        .Col = 4: txt_THK_MAX.Text = .Text
'
'    End With
    
    Call Spread_to_Master(ss1, ss1.ActiveRow)

    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc1("nControl"), Mc1("mControl"), False)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
End Sub

Private Sub ss1_LostFocus()
    With ss1
        
        .Row = ss1.ActiveRow
        
        .Col = 1: txt_PROD_KND2.Text = .Text
        .Col = 2: txt_STDSPEC2.Text = .Text
        .Col = 3: txt_THK_MIN.Text = .Text
        .Col = 4: txt_THK_MAX.Text = .Text
    
    End With
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(ss2, Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 8)
    End If
    
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 8)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code As String

    If ss2.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss2.ActiveCol
    
        Case 5
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss2
                
                DD.sWitch = "SP"
                DD.sKey = txt_PROD_KND2.Text
                DD.rControl.Add Item:=5
                DD.rControl.Add Item:=6
                ss2.Row = ss2.ActiveRow
                
                Call Gf_Usage_DD(M_CN1, KeyCode)
                
            End If
            
    End Select
    
End Sub

Private Sub ss2_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss2, NewRow)
End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub Chk_ss1_Click(Value As Integer)
    
    If Chk_ss1.Value = ssCBUnchecked Then
       If Chk_ss2.Value = ssCBUnchecked Then
            Chk_ss1.Value = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Gf_Sp_Change(Proc_Sc, Sc1) Then
        Chk_ss1.ForeColor = &HFF&
        Chk_ss2.ForeColor = &H808080
        Chk_ss2.Value = ssCBUnchecked
    Else
        Chk_ss1.Value = ssCBUnchecked
        Chk_ss2.Value = ssCBChecked
    End If
        
End Sub

Private Sub Chk_ss2_Click(Value As Integer)
    
    If Chk_ss2.Value = ssCBUnchecked Then
        If Chk_ss1.Value = ssCBUnchecked Then
            Chk_ss2.Value = ssCBChecked
        End If
        Exit Sub
    End If
    
'    If Gf_Sp_Change(Proc_Sc, Sc2) Then
        Chk_ss1.ForeColor = &H808080
        Chk_ss2.ForeColor = &HFF&
        Chk_ss1.Value = ssCBUnchecked
'    Else
'        Chk_ss2.Value = ssCBUnchecked
'        Chk_ss1.Value = ssCBChecked
'    End If
        
End Sub


Private Function funSpreadInsCheck() As Boolean
    
    Dim i As Integer
    Dim iCnt As Integer
    
    With ss1
    
        For i = 1 To .MaxRows
            .Row = i: .Col = 0
            
            If .Text = "Input" Then
                iCnt = iCnt + 1
            End If
        Next i
        
    End With
    
    If iCnt = 0 Then
        funSpreadInsCheck = True
    End If
    
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Spread_to_Master ------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Spread_to_Master(ByVal sp As vaSpread, ByVal iRow As Long)
    
    With sp
    
        If iRow > 0 Then
            .Row = iRow
                 
            .Col = 1: txt_PROD_KND2.Text = .Text
            .Col = 2: txt_STDSPEC2.Text = .Text
            .Col = 3: txt_THK_MIN.Text = .Text
            .Col = 4: txt_THK_MAX.Text = .Text
        Else
            Exit Sub
        End If
    
    End With

End Sub

