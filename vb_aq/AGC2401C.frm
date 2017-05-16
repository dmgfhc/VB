VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2401C 
   Caption         =   "理化检验委托单_AGC2401C"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_CUT_PLT 
      Height          =   330
      Left            =   7305
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txt_smp_sent_no 
      Height          =   300
      Left            =   9795
      MaxLength       =   13
      TabIndex        =   11
      Top             =   525
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   525
      Left            =   4845
      TabIndex        =   8
      Top             =   360
      Width           =   2400
      Begin VB.TextBox txt_line 
         Height          =   300
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   10
         Top             =   135
         Width           =   450
      End
      Begin VB.CheckBox txt_DH_FL 
         Caption         =   "热处理"
         Height          =   240
         Left            =   45
         TabIndex        =   9
         Top             =   195
         Width           =   915
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   1110
         Top             =   135
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   556
         Caption         =   "产线"
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
   Begin VB.TextBox txt_plt 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6210
      MaxLength       =   2
      TabIndex        =   7
      Tag             =   "plt"
      Top             =   90
      Width           =   600
   End
   Begin VB.CheckBox txt_check 
      Caption         =   "已处理对象"
      Height          =   390
      Left            =   8445
      TabIndex        =   6
      Top             =   75
      Width           =   1275
   End
   Begin VB.ComboBox COB_SHIFT 
      Height          =   300
      ItemData        =   "AGC2401C.frx":0000
      Left            =   3720
      List            =   "AGC2401C.frx":000D
      TabIndex        =   4
      Top             =   555
      Width           =   810
   End
   Begin VB.ComboBox COB_GROUP 
      Height          =   300
      ItemData        =   "AGC2401C.frx":001A
      Left            =   1350
      List            =   "AGC2401C.frx":002A
      TabIndex        =   2
      Top             =   555
      Width           =   960
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "打印委托单"
      Height          =   330
      Left            =   13470
      TabIndex        =   1
      Top             =   540
      Width           =   1260
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8370
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   945
      Width           =   15000
      _Version        =   393216
      _ExtentX        =   26458
      _ExtentY        =   14764
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
      MaxCols         =   27
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "AGC2401C.frx":003A
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   120
      Top             =   90
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "日期"
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
   Begin InDate.UDate dtp_prod_fr 
      Height          =   315
      Left            =   1335
      TabIndex        =   3
      Tag             =   "日期"
      Top             =   90
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
      MaxLength       =   10
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   2490
      Top             =   555
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "班次"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   120
      Top             =   555
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "班别"
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
   Begin InDate.UDate dtp_prod_to 
      Height          =   315
      Left            =   3225
      TabIndex        =   5
      Top             =   90
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   2820
      Top             =   90
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Caption         =   "至"
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
      Height          =   300
      Left            =   4845
      Top             =   90
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "工厂"
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
   Begin InDate.ULabel ULabel6 
      Height          =   300
      Left            =   8415
      Top             =   510
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "委托单号"
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
      ForeColor       =   0
   End
End
Attribute VB_Name = "AGC2401C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Template System
'-- Sub_System Name   Common
'-- Program Name      Refer Template
'-- Program ID        Refer
'-- Document No       Q-00-0010(Specification)
'-- Designer          zhang lin
'-- Coder             zhang lin
'-- Date              2007.3.22
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

'Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(dtp_prod_fr, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(dtp_prod_to, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'     Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_line, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
 Call Gp_Ms_Collection(txt_smp_sent_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_check, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_DH_FL, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2401C.P_ONEROW", Key:="P-O"
    sc1.Add Item:="AGC2401C.P_REFER", Key:="P-R"
    sc1.Add Item:="AGC2401C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxRows, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Call Gp_Sp_ColHidden(ss1, 22, True)
    Call Gp_Sp_ColHidden(ss1, 24, True)
    Call Gp_Sp_ColHidden(ss1, 25, True)
    Call Gp_Sp_ColHidden(ss1, 26, True)
    Call Gp_Sp_ColHidden(ss1, 27, True)
    
    Me.KeyPreview = True

End Sub



Private Sub cmdReport_Click()
'    If dtp_yy_mm.RawData = "" Then
'       MsgBox "请先输入您要查询的日期！", vbCritical, "系统提示信息"
'       Exit Sub
'    End If
Dim iROW As Integer

    If ss1.MaxRows < 1 Then Exit Sub
    
    ss1.Col = 2
    For iROW = 1 To ss1.MaxRows
        ss1.Row = iROW
        
        If ss1.Text <> txt_smp_sent_no.Text Then
           Call MsgBox("请先查询后，再打印！", vbCritical, "系统提示信息")
           Exit Sub
        End If
    Next iROW
        
    
    Screen.MousePointer = vbHourglass
    
    Call ExcelPrn
    Screen.MousePointer = vbDefault


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
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_HdColColor(ss1, 1)
    
    dtp_prod_fr.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    dtp_prod_to.RawData = Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMMDD') FROM DUAL")
    
    Screen.MousePointer = vbDefault
    
    txt_check.Value = 0
    Call subButtonHide
    txt_smp_sent_no.Text = ""
    TXT_CUT_PLT.Text = ""
    
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


End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
    
    txt_check.Value = 0
    Call subButtonHide
    txt_smp_sent_no.Text = ""
    TXT_CUT_PLT.Text = ""
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
                
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            Call subButtonHide
            
            If ss1.MaxRows >= 1 Then
               ss1.Row = 1
               ss1.Col = 26
               TXT_CUT_PLT.Text = ss1.Text
            End If
            
            Exit Sub
        End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    Dim AdoRs       As ADODB.Recordset
    Dim sQuery      As String
    Dim I           As Integer
    
    If txt_check.Value = "1" Then Exit Sub
    
    Set AdoRs = New ADODB.Recordset
       
    sQuery = "SELECT Gf_SMP_SEND_NO( "
    sQuery = sQuery & "'" & txt_plt & "') "
    sQuery = sQuery & "FROM DUAL"
    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    If Not AdoRs.BOF And Not AdoRs.EOF Then
       For I = 1 To ss1.MaxRows
           ss1.Row = I
           ss1.Col = 0
           If ss1.Text = "Update" Then
              ss1.Col = 2
              ss1.Text = AdoRs.Fields(0) & ""
           End If
       Next I
    End If
    AdoRs.Close
    Set AdoRs = Nothing
 
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
      Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
      Call subButtonHide
    End If
    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 22)
    ss1.SetFocus
    

End Sub

Public Sub Spread_Can()

    
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 22)
    
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
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    ss1.Row = ss1.ActiveRow
    ss1.Col = 2
    txt_smp_sent_no.Text = ss1.Text
    ss1.Col = 26
    TXT_CUT_PLT.Text = ss1.Text
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 22)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 22)
    
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

Private Sub ExcelPrn()
    Dim I               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sRow            As String
    
    If ss1.MaxRows < 1 Then
       MsgBox "请先查询数据再打印！", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
'    If Trim(TXT_CHECK) <> dtp_yy_mm.RawData Then
'       MsgBox "选择的日期没有进行查询，请先查询数据再打印！", vbCritical, "系统提示信息"
'       Exit Sub
'
'    End If

    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AGC2401C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
            
    xlApp.Range("B3").Value = txt_smp_sent_no.Text
    xlApp.Range("B4").Value = TXT_CUT_PLT.Text
    xlApp.Range("C4").Value = Mid(Now, 1, 4) + "年" + _
                              Mid(Now, 6, 2) + "月" + _
                              Mid(Now, 9, 2) + "日"
          
    Clipboard.Clear
    ss1.Row = 1: ss1.Col = 3: ss1.Row2 = ss1.MaxRows: ss1.Col2 = 21
    Clipboard.SetText ss1.Clip
    xlApp.Range("A7").Select
    xlApp.ActiveSheet.Paste

   
    Clipboard.Clear
    
'    sRow = "A" & ss1.MaxRows + 9 & ":B" & ss1.MaxRows + 6
'    xlApp.Range(sRow).MERGECELLS = True
    sRow = "A" & ss1.MaxRows + 7
    xlApp.Range(sRow).Value = "委托人：" & sUsername
    xlApp.Range(sRow).Font.Size = 10
    sRow = "B" & ss1.MaxRows + 7
    xlApp.Range(sRow).Value = "委托时间：" & Format(Now, "YYYY-MM-DD HH:MM:SS")
    xlApp.Range(sRow).Font.Size = 10
'    xlApp.Range("O5").Value = Format(Now, "YYYY-MM-DD HH:MM:SS")
    sRow = "F" & ss1.MaxRows + 7
    xlApp.Range(sRow).Value = "送样人："
    xlApp.Range(sRow).Font.Size = 10
'    sRow = "K" & ss1.MaxRows + 7
'    xlApp.Range(sRow).Value = "收样人："
'    xlApp.Range(sRow).Font.Size = 10
    sRow = "M" & ss1.MaxRows + 7
    xlApp.Range(sRow).Value = "送样时间："
    xlApp.Range(sRow).Font.Size = 10
'    xlApp.Range(sRow).Font.Bold = True
    xlApp.ActiveSheet.Paste
    
                    
    ss1.ClearSelection
    With xlApp.Application.FindFormat.Borders
        .LineStyle = 1
    End With

    sRow = "A7:S" & ss1.MaxRows + 6
    xlApp.Range(sRow).Select
    With xlApp.Selection.Borders
        .LineStyle = 1
    End With
'    xlApp.Columns("C:E").AutoFit
'    xlApp.Columns("J").AutoFit
    Screen.MousePointer = vbDefault
    xlApp.Application.Visible = True
'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
'    xlApp.DisplayAlerts = False
'    xlSheet.Close
   
'    Set xlSheet = Nothing
'    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub subButtonHide()

Dim iROW, iCol As Integer

    If txt_check.Value = 1 Then
        MDIMain.MenuTool.Buttons(8).Enabled = True    'Row delete
  '      MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
        cmdReport.Visible = True
        Call Gp_Sp_ColHidden(ss1, 2, False)
        Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows, True)
        
    Else
        MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
 '       MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
        cmdReport.Visible = False
        Call Gp_Sp_ColHidden(ss1, 2, True)
    End If

End Sub

Private Sub txt_check_Click()

    Call subButtonHide
    Call Form_Ref
    txt_smp_sent_no.Text = ""
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub


