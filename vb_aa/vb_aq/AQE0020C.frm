VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQE0020C 
   Caption         =   "录入成材率目标值_AQE0020C"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   1995
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   9525
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_SIZE_KND 
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
      Left            =   1110
      TabIndex        =   20
      Top             =   540
      Width           =   1890
   End
   Begin VB.TextBox txt_CPY_DATE_TO 
      Height          =   330
      Left            =   11430
      TabIndex        =   16
      Top             =   495
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox txt_CPY_DATE_FR 
      Height          =   330
      Left            =   10665
      TabIndex        =   15
      Top             =   495
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox txt_REF_DATE 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3570
      TabIndex        =   8
      Top             =   510
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txt_SIZE_NUM 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4245
      TabIndex        =   7
      Top             =   510
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.TextBox txt_FACTORY 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   90
      Width           =   525
   End
   Begin VB.TextBox txt_FACTORY_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   1365
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8235
      Left            =   45
      TabIndex        =   2
      Top             =   945
      Width           =   15195
      _Version        =   393216
      _ExtentX        =   26802
      _ExtentY        =   14526
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
      MaxCols         =   18
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE0020C.frx":0000
   End
   Begin InDate.ULabel ULabel2 
      Height          =   330
      Index           =   1
      Left            =   45
      Top             =   90
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
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
   Begin InDate.ULabel ULabel3 
      Height          =   330
      Left            =   4905
      Top             =   90
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
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
   Begin Threed.SSCommand SSCom_Add 
      Height          =   345
      Left            =   10620
      TabIndex        =   3
      Top             =   60
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   609
      _Version        =   196609
      Caption         =   "复制各年/月数据"
   End
   Begin InDate.UDate dtp_copy_to 
      Height          =   330
      Left            =   14160
      TabIndex        =   4
      Top             =   435
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel9 
      Height          =   330
      Left            =   12240
      Top             =   450
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   582
      Caption         =   "TO"
      Alignment       =   1
      BackColor       =   16777088
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
   Begin InDate.UDate dtp_copy_from 
      Height          =   330
      Left            =   14145
      TabIndex        =   5
      Top             =   75
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel7 
      Height          =   330
      Left            =   12240
      Top             =   75
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   582
      Caption         =   "FROM"
      Alignment       =   1
      BackColor       =   16777088
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
   Begin InDate.UDate dtp_YEAR_MONTH 
      Height          =   330
      Left            =   5970
      TabIndex        =   6
      Tag             =   "日期"
      Top             =   90
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel1 
      Height          =   330
      Left            =   45
      Top             =   525
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
      Caption         =   "定尺分类"
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
   Begin Threed.SSFrame SSFrame1 
      Height          =   345
      Left            =   3600
      TabIndex        =   9
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      _Version        =   196609
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption opt_REF_YEAR 
         Height          =   255
         Left            =   45
         TabIndex        =   10
         Top             =   45
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16448
         Caption         =   "年"
      End
      Begin Threed.SSOption opt_REF_YMONTH 
         Height          =   255
         Left            =   585
         TabIndex        =   11
         Top             =   45
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   64
         Caption         =   "年月"
         Value           =   -1
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   330
      Left            =   12915
      TabIndex        =   12
      Top             =   90
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption opt_CPY_YEAR_FROM 
         Height          =   255
         Left            =   45
         TabIndex        =   13
         Top             =   45
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16448
         Caption         =   "年"
      End
      Begin Threed.SSOption opt_CPY_YMONTH_FROM 
         Height          =   255
         Left            =   540
         TabIndex        =   14
         Top             =   45
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   64
         Caption         =   "年月"
         Value           =   -1
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   330
      Left            =   12915
      TabIndex        =   17
      Top             =   450
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption opt_CPY_YEAR_TO 
         Height          =   255
         Left            =   45
         TabIndex        =   18
         Top             =   45
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16448
         Caption         =   "年"
      End
      Begin Threed.SSOption opt_CPY_YMONTH_TO 
         Height          =   255
         Left            =   540
         TabIndex        =   19
         Top             =   45
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   64
         Caption         =   "年月"
         Value           =   -1
      End
   End
End
Attribute VB_Name = "AQE0020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       QUALITY MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        AQE0020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Sun
'-- Coder             Sun
'-- Date              2007.06.28
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
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2
Dim lCopyRow As Long                'Copy Row


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
'   FormType = "Refer"
         
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_SIZE_NUM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_FACTORY, "p", "n", " ", " ", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_REF_DATE, "p", " ", " ", " ", "r", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_CPY_DATE_FR, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_CPY_DATE_TO, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(dtp_YEAR_MONTH, " ", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
'         Call Gp_Ms_Collection(txt_STLGRD, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
         
        'Spread_Collection
    'Call Spread_Collection("Column1_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
        Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 5, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 6, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 17, "p", " ", " ", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       Call Gp_Sp_Collection(ss1, 18, "p", " ", " ", "i", "a", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
       

    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQE0020C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQE0020C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:="AQE0020C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQE0020C.P_COPY", Key:="P-C"
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

End Sub

Private Sub CBO_ADD()
Dim sSQL As String
    
    sSQL = "SELECT DISTINCT  CD||'-'||CD_SHORT_NAME FROM ZP_CD WHERE CD_MANA_NO ='B0043'"
    
    Call Gf_ComboAdd(M_CN1, cbo_SIZE_KND, sSQL)
    
    
End Sub

Private Sub Form_Activate()
    
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

        dtp_YEAR_MONTH.Text = Date

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass                '设置屏幕的鼠标样式为“箭头加小沙漏”
    
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
'    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call CBO_ADD
    

    

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

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

End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        txt_FACTORY_NAME.Text = ""
        dtp_copy_from.RawData = ""
        dtp_copy_to.RawData = ""
        
    End If
    
   
    
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Pro()

'         Call Spread_Cheack
         If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           Call Form_Ref
         End If
         
    
    

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    ss1.SetFocus
    

End Sub


Public Sub Form_Ref()
    
On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
        dtp_copy_from.RawData = ""
        dtp_copy_to.RawData = ""
        txt_SIZE_NUM.Text = Mid(Trim(cbo_SIZE_KND.Text), 1, 2)
        Call REF_DATE_CHECK
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
    End If
    
    Exit Sub


Refer_Err:
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
Public Sub Spread_Del()
    
    Call GP_SET_CELL_VALUE(ss1, ss1.Row, 0, "Delete")

End Sub

Private Sub opt_CPY_YEAR_FROM_Click(Value As Integer)
    If opt_CPY_YEAR_FROM.Value = True Then
        dtp_copy_from.Mask = "%%%%"
        dtp_copy_from.Text = "____"
    End If
End Sub

Private Sub opt_CPY_YEAR_TO_Click(Value As Integer)
    If opt_CPY_YEAR_TO.Value = True Then
        dtp_copy_to.Mask = "%%%%"
        dtp_copy_to.Text = "____"
    End If
End Sub

Private Sub opt_CPY_YMONTH_FROM_Click(Value As Integer)
    If opt_CPY_YMONTH_FROM.Value = True Then
        dtp_copy_from.Mask = "%%%%-%%"
        dtp_copy_from.Text = "____-__"
    End If
End Sub

Private Sub opt_CPY_YMONTH_TO_Click(Value As Integer)
    If opt_CPY_YMONTH_TO.Value = True Then
        dtp_copy_to.Mask = "%%%%-%%"
        dtp_copy_to.Text = "____-__"
    End If
End Sub

Private Sub opt_REF_YEAR_Click(Value As Integer)
     If opt_REF_YEAR.Value = True Then
        dtp_YEAR_MONTH.Mask = "%%%%"
        dtp_YEAR_MONTH.Text = "____"
    End If
    dtp_YEAR_MONTH.Text = Mid(Date, 1, 4)

End Sub

Private Sub opt_REF_YMONTH_Click(Value As Integer)
    If opt_REF_YMONTH.Value = True Then
        dtp_YEAR_MONTH.Mask = "%%%%-%%"
        dtp_YEAR_MONTH.Text = "____-__"
    End If
End Sub

Public Sub Spread_Can()

   Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub
Public Sub Spread_Cpy()

    lCopyRow = ss1.ActiveRow

End Sub
Public Sub Spread_Pst()

    Call GP_ROW_PASTE(Proc_Sc("Sc"), lCopyRow)
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
    

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 15)
    End If
    
    With ss1
        .Row = .ActiveRow
        .Col = 1
         If .Text = "" Then
            .Col = 2
            .Text = ""
        End If
    End With
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

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Dim sTemp_Code As String
    Dim iCol As Long
    Dim iRow As Long

    iCol = ss1.ActiveCol
    iRow = ss1.ActiveRow

    If ss1.MaxRows < 1 Then Exit Sub

    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol
                
        Case 1
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "Q0075"
                DD.rControl.Add Item:=1
                DD.rControl.Add Item:=2
                ss1.Row = ss1.ActiveRow
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            Else
                If Gf_GetCellText(ss1, iRow, iCol) = "" Then
                    Call GP_SET_CELL_VALUE(ss1, iRow, iCol + 1, "")
                End If
            
            End If
            
            
        Case 3
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "B0043"
                DD.rControl.Add Item:=3
                DD.rControl.Add Item:=4
                ss1.Row = ss1.ActiveRow
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            Else
                If Gf_GetCellText(ss1, iRow, iCol) = "" Then
                    Call GP_SET_CELL_VALUE(ss1, iRow, iCol + 1, "")
                End If
            
            End If


    End Select
  
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 11)
   
End Sub

Private Sub SSCom_Add_Click()

On Error GoTo Cmd1_Error

    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim sMesg As String
    Dim icount As Integer
    
    Dim adoCmd As adodb.Command

    Call CPY_DATE_CHECK
    
      If dtp_copy_from.RawData = "" Or dtp_copy_to.RawData = "" Or dtp_copy_to.RawData <= dtp_copy_from.RawData Then
          Call Gp_MsgBoxDisplay("必须输入正确的日期...")
          Exit Sub
      End If
    
    
    Screen.MousePointer = vbHourglass
        
    'Db Connection Check
'    If GF_DbConnect = False Then
'       Exit Sub
'    End If
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    Set adoCmd.ActiveConnection = M_CN1
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "AQE0020C.P_COPY"
    
    M_CN1.BeginTrans
    
'    ss1.Col = 3
'    ss1.Row = ss1.ActiveRow
'    txt_SIZE_KND_VAULE.Text = ss1.Text
   txt_SIZE_NUM.Text = Mid(Trim(cbo_SIZE_KND.Text), 1, 2)

    'Ceate Parameter (Input) iType + iColumn
    For icount = 1 To 5
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next icount
    
    'Ceate Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
    adoCmd.Parameters(0).Value = txt_CPY_DATE_FR.Text
    adoCmd.Parameters(1).Value = txt_CPY_DATE_TO.Text
    adoCmd.Parameters(2).Value = txt_FACTORY.Text
    adoCmd.Parameters(3).Value = txt_SIZE_NUM.Text
    adoCmd.Parameters(4).Value = sUserID                            'User-id
    adoCmd.Execute
     

     'Error Check
     If adoCmd("Error") <> "0" Then

         ret_Result_ErrCode = adoCmd("Error")
         ret_Result_ErrMsg = adoCmd("Messg")
         sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg

         Call Gp_MsgBoxDisplay(sErrMessg)
         Screen.MousePointer = vbDefault
         Set adoCmd = Nothing
         M_CN1.RollbackTrans
         Exit Sub

      End If
        
      M_CN1.CommitTrans
      Screen.MousePointer = vbDefault
      Exit Sub
    
Cmd1_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    M_CN1.RollbackTrans
    Call Gp_MsgBoxDisplay("Cmd1_Error : " & Error)


End Sub




Private Sub txt_FACTORY_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0001"

        DD.rControl.Add Item:=txt_FACTORY
        DD.rControl.Add Item:=txt_FACTORY_NAME
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
   ElseIf Len(Trim(txt_FACTORY.Text)) = 2 Then
      Select Case txt_FACTORY.Text
             Case "C1", "c1"
                  txt_FACTORY.Text = "C1"
             Case "C3", "c3"
                  txt_FACTORY.Text = "C3"
             Case "B1", "b1"
                  txt_FACTORY.Text = "B1"
             Case "*", "**"
                  txt_FACTORY.Text = "**"
             Case ""
                  txt_FACTORY.Text = ""
                  txt_FACTORY_NAME.Text = ""
          End Select
        
    End If

    If Len(Trim(txt_FACTORY.Text)) = txt_FACTORY.MaxLength Then
       txt_FACTORY_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", txt_FACTORY.Text, 2)
    Else
        txt_FACTORY_NAME.Text = ""
    End If
    
End Sub

Private Sub REF_DATE_CHECK()

 If opt_REF_YMONTH.Value = True Then
    txt_REF_DATE.Text = dtp_YEAR_MONTH.RawData
    
 ElseIf opt_REF_YEAR.Value = True Then
    txt_REF_DATE.Text = Mid(Trim(dtp_YEAR_MONTH.RawData), 1, 4)

 End If
  
End Sub
Private Sub CPY_DATE_CHECK()

If opt_CPY_YMONTH_FROM.Value = True Then
   txt_CPY_DATE_FR.Text = dtp_copy_from.RawData
ElseIf opt_CPY_YEAR_FROM.Value = True Then
   txt_CPY_DATE_FR.Text = Mid(Trim(dtp_copy_from.RawData), 1, 4)
End If

If opt_CPY_YMONTH_TO.Value = True Then
   txt_CPY_DATE_TO.Text = dtp_copy_to.RawData
ElseIf opt_CPY_YEAR_TO.Value = True Then
   txt_CPY_DATE_TO.Text = Mid(Trim(dtp_copy_to.RawData), 1, 4)
End If

End Sub


