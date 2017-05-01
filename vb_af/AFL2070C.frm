VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AFL2070C 
   Caption         =   "库存状态查询界面_AFL2070C"
   ClientHeight    =   9225
   ClientLeft      =   375
   ClientTop       =   2295
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_PROD_TYPE 
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
      ItemData        =   "AFL2070C.frx":0000
      Left            =   1860
      List            =   "AFL2070C.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "产品类型"
      Top             =   135
      Width           =   765
   End
   Begin VB.ComboBox cbo_YARD_TYPE 
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
      Height          =   315
      ItemData        =   "AFL2070C.frx":0004
      Left            =   4800
      List            =   "AFL2070C.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "跨号"
      Top             =   135
      Width           =   765
   End
   Begin VB.ComboBox cbo_ZONE_TYPE 
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
      Height          =   315
      ItemData        =   "AFL2070C.frx":0008
      Left            =   7785
      List            =   "AFL2070C.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "区号"
      Top             =   135
      Width           =   765
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8655
      Left            =   75
      TabIndex        =   6
      Top             =   525
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   15266
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AFL2070C.frx":000C
      Begin FPSpread.vaSpread ss2 
         Height          =   4965
         Left            =   0
         TabIndex        =   4
         Top             =   3690
         Width           =   15105
         _Version        =   393216
         _ExtentX        =   26644
         _ExtentY        =   8758
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
         MaxCols         =   17
         MaxRows         =   2
         OperationMode   =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFL2070C.frx":005E
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3630
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   15105
         _Version        =   393216
         _ExtentX        =   26644
         _ExtentY        =   6403
         _StockProps     =   64
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
         MaxCols         =   1
         MaxRows         =   1
         OperationMode   =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AFL2070C.frx":08E0
      End
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1125
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4275
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   180
      Top             =   1665
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   180
      Top             =   2250
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   315
      Top             =   135
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "库种类"
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
      Left            =   3240
      Top             =   135
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "跨号"
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
      Left            =   6225
      Top             =   135
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      Caption         =   "区号"
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
Attribute VB_Name = "AFL2070C"
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
'-- Program Name      YARD STOCK STATUS
'-- Program ID        AFL2070C
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2003.9.9
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

Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

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
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

        'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(cbo_PROD_TYPE, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(cbo_YARD_TYPE, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
     Call Gp_Ms_Collection(cbo_ZONE_TYPE, "p", "n", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)
    
    'MASTER Collection
    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=cControl1, Key:="cControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(Text1, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFL2070C.P_SREFER", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
     
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"

    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub cbo_PROD_TYPE_Click()
    cbo_YARD_TYPE.Enabled = True
    Call Gf_ComboAdd(M_CN1, cbo_YARD_TYPE, "SELECT DISTINCT YARD_TYPE FROM FP_STDYARD WHERE PROD_TYPE = '" + cbo_PROD_TYPE + "'")
End Sub

Private Sub cbo_yard_type_Click()
    cbo_ZONE_TYPE.Enabled = True
    Call Gf_ComboAdd(M_CN1, cbo_ZONE_TYPE, "SELECT DISTINCT ZONE_TYPE FROM FP_STDYARD WHERE PROD_TYPE = '" + cbo_PROD_TYPE + "' AND YARD_TYPE = '" + cbo_YARD_TYPE + "' AND YARD_KND = '00' ")
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

    Call Gf_ComboAdd(M_CN1, cbo_PROD_TYPE, "SELECT DISTINCT PROD_TYPE FROM FP_STDYARD")
    
    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc2")("Spread"))

    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    
    Call Gp_Spl_SizeGet(SSSplitter1, "F-System.INI", Me.Name, "H")
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "F-System.INI", Me.Name)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) And Gf_Sp_ProceExist(Proc_Sc("Sc2")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "F-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "F-System.INI", Me.Name)
    
    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
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
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()

    Timer1.Enabled = False
    Timer2.Enabled = False
    If Gf_Sp_Cls(Proc_Sc("SC1")) And Gf_Sp_Cls(Proc_Sc("SC2")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("pControl"), False)
        
        Call Gp_Sp_ColHidden(ss2, 2, False)
        Call Gp_Sp_ColHidden(ss2, 15, False)
        Call Gp_Sp_ColHidden(ss2, 16, False)
        
        ss2.Col = 3
        ss2.Row = 0
        ss2.Text = "物料号"
        
        cbo_YARD_TYPE.Clear
        cbo_ZONE_TYPE.Clear
        cbo_ZONE_TYPE.Enabled = False
        cbo_YARD_TYPE.Enabled = False
        cbo_PROD_TYPE.SetFocus
        cbo_PROD_TYPE.ListIndex = -1
    End If
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Timer1.Enabled = True

    Dim i, j As Integer
    Dim sMesg As String

    If Gf_Sp_ProceExist(Proc_Sc("Sc1").Item("Spread")) Then Exit Sub

    If cbo_PROD_TYPE.Text = "" Then
       MsgBox "产品类型必须输入！", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    If cbo_YARD_TYPE.Text = "" Then
       MsgBox "跨号必须输入！", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
    If cbo_ZONE_TYPE.Text = "" Then
       MsgBox "区号必须输入！", vbCritical, "系统提示信息"
       Exit Sub
    End If
    
        sMesg = Gf_Ms_NeceCheck2(mControl1)
        If sMesg = "OK" Then

            If Sp_Header_Refer() Then
               If Sp_Data_Refer() Then
                    Call Gp_Ms_ControlLock(Mc1("pControl"), True)
                    cbo_YARD_TYPE.Enabled = False
                    cbo_ZONE_TYPE.Enabled = False
                
                    If Mid(LTrim(cbo_PROD_TYPE.Text), 1, 1) = "S" Then
                          ss2.Col = 3
                          ss2.Row = 0
                          ss2.Text = "板坯号"
                          Call Gp_Sp_ColHidden(ss2, 2, True)
                          Call Gp_Sp_ColHidden(ss2, 15, False)
                          Call Gp_Sp_ColHidden(ss2, 16, False)
                    ElseIf Mid(LTrim(cbo_PROD_TYPE.Text), 1, 1) = "P" Then
                          ss2.Col = 3
                          ss2.Row = 0
                          ss2.Text = "钢板号"
                          Call Gp_Sp_ColHidden(ss2, 2, True)
                          Call Gp_Sp_ColHidden(ss2, 15, True)
                          Call Gp_Sp_ColHidden(ss2, 16, True)
                    ElseIf Mid(LTrim(cbo_PROD_TYPE.Text), 1, 1) = "C" Then
                          ss2.Col = 3
                          ss2.Row = 0
                          ss2.Text = "钢卷号"
                          Call Gp_Sp_ColHidden(ss2, 2, False)
                          Call Gp_Sp_ColHidden(ss2, 15, True)
                          Call Gp_Sp_ColHidden(ss2, 16, True)
                    End If
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                    Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc1")("Spread"))
    
                    ss2.MaxRows = 0
                    Exit Sub
               End If
            End If
            
        Else
            sMesg = sMesg + "长度不正确"
            Call Gp_MsgBoxDisplay(sMesg)
        End If

    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, Proc_Sc("SC1"), Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

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

Public Function Sp_Header_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iMaxRow As Integer
    Dim iMaxCol As Integer
    Dim sQuery As String
    
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    sQuery = "SELECT MAX(SUBSTR(LOCATION,4,2)),MAX(SUBSTR(LOCATION,6,2))"
    sQuery = sQuery + " FROM FP_STDYARD "
    sQuery = sQuery + " WHERE PROD_TYPE = '" & LTrim(cbo_PROD_TYPE.Text) & "'"
    sQuery = sQuery + " AND YARD_TYPE = '" & cbo_YARD_TYPE.Text & "'"
    sQuery = sQuery + " AND ZONE_TYPE = '" & cbo_ZONE_TYPE.Text & "'"
    
    With ss1

        Sp_Header_Refer = True
        
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        
        Screen.MousePointer = vbHourglass
        
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Header_Refer = False
            
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            
            Exit Function
            
        End If
        
        iMaxRow = Val(AdoRs(0))
        iMaxCol = Val(AdoRs(1))
        
        AdoRs.Close
        .MaxRows = iMaxRow
        .MaxCols = iMaxCol
        
        For iRow = 0 To .MaxRows - 1
            .Row = iRow + 1
            .Col = 0
            If .Row <= 9 Then
               .Text = "0" + LTrim(STR(.Row))
            Else
               .Text = STR(.Row)
            End If
        Next
        
        For iCol = 0 To .MaxCols - 1
            .Col = iCol + 1
            .Row = 0
            If .Col <= 9 Then
               .Text = "0" + LTrim(STR(.Col))
            Else
               .Text = STR(.Col)
            End If
        Next
        .ReDraw = True
        .Refresh

        Screen.MousePointer = vbDefault

  End With
        
Exit Function

SpreadDisplay_Error:

    Set AdoRs = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer = False
    Screen.MousePointer = vbDefault

End Function

Public Function Sp_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim iMaxRecord As Integer
    Dim sQuery As String
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    
    If LTrim(cbo_PROD_TYPE.Text) = "S" Then
        sQuery = "SELECT A.LOCATION, A.MAX_CNT, B.CNT"
        sQuery = sQuery + " FROM FP_STDYARD A,"
        sQuery = sQuery + " (SELECT YARD_ADDR,COUNT(*) AS CNT FROM FP_SLABYARD WHERE SLAB_NO IS NOT NULL GROUP BY YARD_ADDR) B"
        sQuery = sQuery + " WHERE A.PROD_TYPE = '" & LTrim(cbo_PROD_TYPE.Text) & "'"
        sQuery = sQuery + " AND A.YARD_TYPE ='" & cbo_YARD_TYPE.Text & "'"
        sQuery = sQuery + " AND A.ZONE_TYPE ='" & cbo_ZONE_TYPE.Text & "'"
        sQuery = sQuery + " AND A.LOCATION = B.YARD_ADDR(+)"
    ElseIf LTrim(cbo_PROD_TYPE.Text) = "P" Then
        sQuery = "SELECT A.LOCATION, A.MAX_CNT, B.CNT"
        sQuery = sQuery + " FROM FP_STDYARD A,"
        sQuery = sQuery + " (SELECT YARD_ADDR,COUNT(*) AS CNT FROM GP_PLATEYARD WHERE PLATE_NO IS NOT NULL GROUP BY YARD_ADDR) B"
        sQuery = sQuery + " WHERE A.PROD_TYPE = '" & LTrim(cbo_PROD_TYPE.Text) & "'"
        sQuery = sQuery + " AND A.YARD_TYPE ='" & cbo_YARD_TYPE.Text & "'"
        sQuery = sQuery + " AND A.ZONE_TYPE ='" & cbo_ZONE_TYPE.Text & "'"
        sQuery = sQuery + " AND A.LOCATION = B.YARD_ADDR(+)"
    ElseIf LTrim(cbo_PROD_TYPE.Text) = "C" Then
        sQuery = "SELECT A.LOCATION, A.MAX_CNT, B.CNT"
        sQuery = sQuery + " FROM FP_STDYARD A,"
        sQuery = sQuery + " (SELECT YARD_ADDR,COUNT(*) AS CNT FROM GP_COILYARD WHERE COIL_NO IS NOT NULL AND COIL_LAYER = 1 GROUP BY YARD_ADDR) B"
        sQuery = sQuery + " WHERE A.PROD_TYPE = '" & LTrim(cbo_PROD_TYPE.Text) & "'"
        sQuery = sQuery + " AND A.YARD_TYPE ='" & cbo_YARD_TYPE.Text & "'"
        sQuery = sQuery + " AND A.ZONE_TYPE ='" & cbo_ZONE_TYPE.Text & "'"
        sQuery = sQuery + " AND A.LOCATION = B.YARD_ADDR(+)"
    End If
        
  With ss1

        Sp_Data_Refer = True
        
        .ReDraw = False
        
        Screen.MousePointer = vbHourglass
        
        
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer = False
            .ReDraw = True
            
            AdoRs.Close
            Set AdoRs = Nothing
        
            Screen.MousePointer = vbDefault
            
            Exit Function
            
        End If
        
        'iMaxRecord = AdoRs.RecordCount
        ArrayRecords = AdoRs.GetRows
        iMaxRecord = UBound(ArrayRecords, 2) + 1
        AdoRs.Close
 
            For iRow = 1 To .MaxRows
                For iCol = 1 To .MaxCols
                .Row = iRow
                .Col = iCol
                   For iCnt = 0 To iMaxRecord - 1
                       If Val(Mid(ArrayRecords(0, iCnt), 4, 2)) = .Row Then
                          If Val(Mid(ArrayRecords(0, iCnt), 6, 2)) = .Col Then
                             If ArrayRecords(2, iCnt) <> 0 Then
                               .Text = LTrim(STR(ArrayRecords(2, iCnt))) + " / " + LTrim(STR(ArrayRecords(1, iCnt)))
                                Exit For
                             Else
                               .Text = "0 / " + LTrim(STR(ArrayRecords(1, iCnt)))
                                Exit For
                             End If
                          End If
                
                       End If
                
                   Next
                Next
            Next
        
        .ReDraw = True

        Screen.MousePointer = vbDefault

  End With
  ss1.OperationMode = OperationModeRead
Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    
End Function

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    Dim i, j As String
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    With ss1
      .Row = Row
      .Col = 0
       i = Trim(.Text)
      .Col = Col
      .Row = 0
       j = Trim(.Text)
    End With
    
    If Row > 0 And Col > 0 Then
        Text1.Text = cbo_PROD_TYPE.Text + cbo_YARD_TYPE.Text + cbo_ZONE_TYPE.Text + i + j
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2) Then
           Call Gp_Sp_EvenRowBackcolor(Proc_Sc("Sc2")("Spread"))
        End If
    End If
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 9)
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

Private Sub Timer1_Timer()

    Dim spl, i, j As Integer
    
   With ss1
        For i = 1 To .MaxRows
            For j = 1 To .MaxCols
                .Row = i
                .Col = j
                If .Text <> "" Then
                   spl = Val(InStr(1, LTrim(.Text), "/"))
                   If RTrim(Mid(LTrim(.Text), 1, spl - 1)) = LTrim(Mid(LTrim(.Text), spl + 1, Len(LTrim(.Text)) - spl)) Then
                     .BackColor = &HFFFF&
                   End If
                End If
            Next j
        Next i
   End With
   
   Timer1.Enabled = False
   Timer2.Enabled = True
   
End Sub

Private Sub Timer2_Timer()

    Dim i, j As Integer
    
    With ss1
        For i = 1 To .MaxRows
            For j = 1 To .MaxCols
                .Row = i
                .Col = j
                .BackColor = &H8000000B
            Next j
        Next i
    End With
    
    Timer1.Enabled = True
    Timer2.Enabled = False
    
End Sub
