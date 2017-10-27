VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACB1023C 
   Caption         =   "板坯钢种改判查询及修改_ACB1023C"
   ClientHeight    =   9225
   ClientLeft      =   510
   ClientTop       =   1515
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14115
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8670
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   15293
      _Version        =   196609
      SplitterBarWidth=   2
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "ACB1023C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   6690
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   15240
         _Version        =   393216
         _ExtentX        =   26882
         _ExtentY        =   11800
         _StockProps     =   64
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
         MaxCols         =   37
         MaxRows         =   1
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "ACB1023C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   1950
         Left            =   0
         TabIndex        =   3
         Top             =   6720
         Width           =   15240
         _Version        =   393216
         _ExtentX        =   26882
         _ExtentY        =   3440
         _StockProps     =   64
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
         MaxCols         =   0
         MaxRows         =   3
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "ACB1023C.frx":0E32
         ScrollBarTrack  =   3
      End
   End
   Begin VB.TextBox txt_charge_no 
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
      Left            =   1440
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "炉号信息"
      Top             =   90
      Width           =   975
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   135
      Top             =   90
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "炉    号"
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   2880
      Top             =   90
      Visible         =   0   'False
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   556
      Caption         =   "生产日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.UDate udt_prod_date_fr 
      Height          =   315
      Left            =   3915
      TabIndex        =   4
      Tag             =   "INS_DATE"
      Top             =   90
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
   Begin InDate.UDate udt_prod_date_to 
      Height          =   315
      Left            =   5550
      TabIndex        =   5
      Tag             =   "INS_DATE"
      Top             =   90
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   120
      Left            =   5385
      TabIndex        =   6
      Top             =   195
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   15120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   15105
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "ACB1023C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACB1023C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2007.6.27
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

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

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

Dim sWgtLenFlag As String
Dim sQuery  As String

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(txt_charge_no, "p", "n", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
   'Call Gp_Ms_Collection(udt_prod_date_fr, "p", " ", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
   'Call Gp_Ms_Collection(udt_prod_date_to, "p", " ", " ", " ", "r", " ", "l", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB1023C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="ACB1023C.P_REFER", Key:="P-R"
    sc1.Add Item:="ACB1023C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc1"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").ROW = 0
    sc1.Item("Spread").Text = "◎"
    
    Call Gp_Sp_ColHidden(ss1, 36, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste

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
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Sp1_Setting(SS2)
    
    Call Gf_Sp_Cls(sc1)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 11)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc1")("Spread"), 13)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
      
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

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC1"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        SS2.MaxCols = 0
        rContro1(1).SetFocus
    End If
    
End Sub

Public Sub Form_Ref()

    If Gf_Sp_ProceExist(sc1.Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
        MDIMain.MenuTool.Buttons(4).Enabled = True
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
        SS2.MaxCols = 0
        
    End If
    
End Sub

Public Sub Form_Pro()
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC1"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        MDIMain.MenuTool.Buttons(4).Enabled = True
        MDIMain.MenuTool.Buttons(7).Enabled = False
        MDIMain.MenuTool.Buttons(8).Enabled = False
        MDIMain.MenuTool.Buttons(11).Enabled = False
        MDIMain.MenuTool.Buttons(12).Enabled = False
    End If
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)

    Dim sQuery1 As String   'Chemistry Header, Min, Max, Result Value Display
    Dim sStlgrd As String   'STLGRD
    
    If ROW <= 0 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 6
    sStlgrd = ss1.Text
    
    'Chemistry DATA Display
    sQuery1 = "           SELECT  A.CHEM_COMP_CD, B.CHEM_COMP_MIN, B.CHEM_COMP_MAX, C.CHEM_RSLT,"
    sQuery1 = sQuery1 + "         NVL(B.CHEM_COMP_MIN,0)+NVL(B.CHEM_COMP_MAX,0)+NVL(C.CHEM_RSLT,0), A.CHEM_COMP_SEQ "
    sQuery1 = sQuery1 + "   FROM  (SELECT  CHEM_COMP_CD, CHEM_COMP_SEQ "
    sQuery1 = sQuery1 + "            FROM  QP_CHEM_SEQ) A, "
    sQuery1 = sQuery1 + "         (SELECT  CHEM_COMP_CD, CHEM_COMP_MIN, CHEM_COMP_MAX "
    sQuery1 = sQuery1 + "           FROM   QP_NISCO_CHEM "
    sQuery1 = sQuery1 + "          WHERE   STLGRD = '" & sStlgrd & "') B, "
    sQuery1 = sQuery1 + "         (SELECT  CHEM_COMP_CD, CHEM_RSLT "
    sQuery1 = sQuery1 + "            FROM  QP_CHEM_RSLT "
    sQuery1 = sQuery1 + "           WHERE  HEAT_NO = '" & txt_charge_no.Text & "') C "
    sQuery1 = sQuery1 + "   WHERE  A.CHEM_COMP_CD = B.CHEM_COMP_CD(+) "
    sQuery1 = sQuery1 + "     AND  A.CHEM_COMP_CD = C.CHEM_COMP_CD(+) "
    sQuery1 = sQuery1 + "   ORDER  BY A.CHEM_COMP_SEQ ASC "
    
    'Data Display
    Call Sp_Data_Refer(SS2, sQuery1)
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc1")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc1"), 36)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

'    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
'
'    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
'
'    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 36)
'    End If

    If Shift = 0 Then Proc_Sc("Sc1")("Spread").EditMode = True
    
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim str_orgin As String
    
    With ss1
    
        .Col = .ActiveCol
        .ROW = .ActiveRow
            
        Select Case .ActiveCol
        
            Case 11
            
                If KeyCode = vbKeyF4 Then
                
                    str_orgin = .Text
                    
                    Set DD.sPname = Me.ss1
                    DD.nameType = "1"
                    DD.sWitch = "SP"
                    DD.rControl.Add Item:=11
                    DD.rControl.Add Item:=12
                    
                    Call Gf_Stlgrd_DD(M_CN1, KeyCode)
                    
                    If .Text <> str_orgin Then
                        Call ss1_EditMode(.ActiveCol, .ActiveRow, 2, True)
                    End If
                    
                Else
                
                    str_orgin = .Text
                    sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(.Text) + "'"
                    .Col = .Col + 1
                    .Text = Gf_CodeFind(M_CN1, sQuery)
                    .Col = .Col - 1
                    
                    If .Text <> str_orgin Then
                        Call ss1_EditMode(.ActiveCol, .ActiveRow, 2, True)
                    End If
                    
                End If
                
            Case 13
                
                If KeyCode = vbKeyF4 Then
                
                    str_orgin = .Text
                    
                    Set DD.sPname = Me.ss1
                    DD.sWitch = "SP"
                    DD.sKey = "F0032"
                    DD.rControl.Add Item:=13
                    DD.rControl.Add Item:=14
                    
                    DD.nameType = "2"
                    Call Gf_Common_DD(M_CN1, KeyCode)
                    
                    If .Text <> str_orgin Then
                        Call ss1_EditMode(.ActiveCol, .ActiveRow, 2, True)
                    End If
                    
                Else
                    
                    If Len(Trim(.Text)) = .TypeMaxEditLen Then
                        str_orgin = .Text
                        .Col = 14
                        .Text = Gf_ComnNameFind(M_CN1, "F0032", Trim(str_orgin), 2)
                    Else
                        .Col = 14
                        .Text = ""
                    End If
                    
                End If
                
            End Select
        
    End With
    
End Sub

Private Sub Sp1_Setting(ByVal sPname As Variant, Optional MsgChk As Boolean = True)

    With sPname
    
        .RowHeight(-1) = 12.54
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 12
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 12
        Else
            .RowHeight(0) = 24
        End If
        
        .ColWidth(0) = 6
        
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        '.OperationMode = OperationModeRow
        .RetainSelBlock = True

        .UserResize = UserResizeColumns
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .ROW = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .ROW = 0
        .FontBold = True
        
        'If .ColHeaderRows > 1 Then
        '    .Row = SpreadHeader + 1
        '    .FontBold = True
        'End If
        
        If MsgChk Then
            .LockBackColor = RGB(255, 255, 255)
        End If
        
    End With
    
End Sub

Private Sub Sp_Data_Refer(sPname As Variant, sQuery As String)

On Error GoTo Data_Display_Error

    Dim iCol As Integer
    Dim iCol2 As Integer
    Dim dMaxValue As Double
    Dim dMinValue As Double
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    
    With sPname

        .ReDraw = False
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
            AdoRs.Close
            Set AdoRs = Nothing
            Exit Sub
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1)
            .Col = 1: .Col2 = .MaxCols
            .ROW = 1: .Row2 = -1
            .BlockMode = True
            .TypeHAlign = TypeHAlignRight
            .TypeVAlign = TypeHAlignCenter
            .BlockMode = False
            
            
                    
            For iCol = 0 To .MaxCols - 1
            
                .Col = iCol + 1
                dMinValue = 0
                dMaxValue = 0
                
                If Trim(ArrayRecords(4, iCol)) <> 0 Then    'Min + Max + Result
                
                    .ROW = 0
                    If VarType(ArrayRecords(0, iCol)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(0, iCol))
                    End If
                    
                    .ColWidth(.Col) = 6
                    
                    .ROW = 1  'Min
                    If VarType(ArrayRecords(1, iCol)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(1, iCol))
                        dMinValue = .Value
                    End If
                    
                    .ROW = 2   'Max
                    If VarType(ArrayRecords(2, iCol)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(2, iCol))
                        dMaxValue = .Value
                    End If
                    
                    .ROW = 3   'Result
                    If VarType(ArrayRecords(3, iCol)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(3, iCol))
                    End If
                        
                    If dMinValue + dMaxValue = 0 And Trim(ArrayRecords(3, iCol)) <> 0 Then
                       Call Gp_Sp_CellColor(sPname, .Col, 3, BLUE, WHITE)
                    Else
                        If dMinValue > .Value Or dMaxValue < .Value Then
                            Call Gp_Sp_CellColor(sPname, .Col, 3, RED, WHITE)
                        Else
                            Call Gp_Sp_CellColor(sPname, .Col, 3, BLUE, WHITE)
                        End If
                    End If
                    
                Else
                    Call Gp_Sp_ColHidden(sPname, .Col, True)
                End If
                    
            Next iCol
            
            .ReDraw = True
        End If
        
    End With
        
    Exit Sub

Data_Display_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Call Gp_MsgBoxDisplay("Data_Display_Error : " & Error)
    
End Sub
