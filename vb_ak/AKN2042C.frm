VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKN2042C 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "替换炉"
   ClientHeight    =   8745
   ClientLeft      =   1275
   ClientTop       =   1665
   ClientWidth     =   14070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   14070
   Begin VB.TextBox txt_heat_mana_no 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   10575
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   " "
      Top             =   135
      Width           =   1155
   End
   Begin Threed.SSCommand cmd_ok 
      Height          =   465
      Left            =   11835
      TabIndex        =   7
      Top             =   45
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "确定"
   End
   Begin VB.TextBox txt_ccm_prc_line 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   310
      Left            =   1280
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "连浇号"
      Top             =   135
      Width           =   375
   End
   Begin VB.TextBox txt_Stlgrd_grp 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   8295
      Locked          =   -1  'True
      TabIndex        =   3
      Tag             =   "钢种组"
      Text            =   " "
      Top             =   135
      Width           =   375
   End
   Begin VB.TextBox txt_plt 
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
      Left            =   11460
      MaxLength       =   2
      TabIndex        =   6
      Tag             =   "工厂"
      Top             =   150
      Visible         =   0   'False
      Width           =   285
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8070
      Left            =   75
      TabIndex        =   5
      Top             =   630
      Width           =   13950
      _ExtentX        =   24606
      _ExtentY        =   14235
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AKN2042C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   5055
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   8565
         _Version        =   393216
         _ExtentX        =   15108
         _ExtentY        =   8916
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   15
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2042C.frx":0072
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   8070
         Left            =   8625
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   5325
         _Version        =   393216
         _ExtentX        =   9393
         _ExtentY        =   14235
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
         MaxCols         =   22
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2042C.frx":09E2
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   2955
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   5115
         Width           =   8565
         _Version        =   393216
         _ExtentX        =   15108
         _ExtentY        =   5212
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
         MaxCols         =   16
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKN2042C.frx":158D
      End
   End
   Begin VB.TextBox txt_org_heat_no 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "原炉号"
      Top             =   135
      Width           =   1155
   End
   Begin VB.TextBox txt_plan_no 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   5685
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "计划名"
      Top             =   135
      Width           =   1245
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   4545
      Top             =   135
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "计划名"
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
      Left            =   -315
      Top             =   -315
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "原炉号"
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
   Begin Threed.SSCommand cmd_exit 
      Height          =   465
      Left            =   12915
      TabIndex        =   8
      Top             =   45
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   820
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "取消"
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   1980
      Top             =   135
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "原炉号"
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
      Left            =   135
      Top             =   135
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      Caption         =   "连浇号"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   7155
      Top             =   135
      Width           =   1095
      _ExtentX        =   1931
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
   End
   Begin InDate.ULabel ULabel8 
      Height          =   300
      Left            =   9495
      Top             =   135
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   529
      Caption         =   "变更炉号"
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
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   90
      X2              =   13995
      Y1              =   585
      Y2              =   585
   End
End
Attribute VB_Name = "AKN2042C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      CHANGE CHARGE
'-- Program ID        AKN2042C
'-- Document No
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2011.10.12
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
Public sDateTime As String          'Active Form Authority Setting
Public sQuery_Rt As String          'Active Form Authority Setting

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

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
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim P_Heat_Edt_Seq As Long          'Heat_Edt_Seq
Dim P_Slab_Edt_Seq As Long          'Slab_Edt_Seq

Dim iSelect_ss1_Row As Integer      'SS1 Select Row
Dim iSelect_ss3_Row As Integer      'SS1 Select Row

Private Sub Form_Define()
     
    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(txt_ccm_prc_line, "p", "n", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
      Call Gp_Ms_Collection(txt_Stlgrd_grp, "p", "n", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(txt_plan_no, "p", "n", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_heat_mana_no, "p", "n", " ", " ", " ", " ", " ", pContro2, nContro2, mContro2, iContro2, rContro2, aContro2, lContro2)
    
    'MASTER Collection
    Mc2.Add Item:=pContro2, Key:="pControl"
    Mc2.Add Item:=nContro2, Key:="nControl"
    Mc2.Add Item:=mContro2, Key:="mControl"
    Mc2.Add Item:=iContro2, Key:="iControl"
    Mc2.Add Item:=rContro2, Key:="rControl"
    Mc2.Add Item:=cContro2, Key:="cControl"
    Mc2.Add Item:=aContro2, Key:="aControl"
    Mc2.Add Item:=lContro2, Key:="lControl"
       
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, iCol, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next iCol
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AFN2042C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss2.MaxCols
        Call Gp_Sp_Collection(ss2, iCol, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Next iCol
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AFN2042C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For iCol = 1 To ss3.MaxCols
        Call Gp_Sp_Collection(ss3, iCol, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Next iCol
    
    'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AFN2042C.P_REFER3", Key:="P-R"
    Sc3.Add Item:=pColumn3, Key:="pColumn"
    Sc3.Add Item:=nColumn3, Key:="nColumn"
    Sc3.Add Item:=aColumn3, Key:="aColumn"
    Sc3.Add Item:=mColumn3, Key:="mColumn"
    Sc3.Add Item:=iColumn3, Key:="iColumn"
    Sc3.Add Item:=lColumn3, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
   
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    Call Gp_Sp_ColHidden(ss1, 15, True)    'SEQ_NO
    
    Call Gp_Sp_ColHidden(ss2, 1, True)     'SEQ_NO
    
'    Call Gp_Sp_ColHidden(ss3, 1, True)
'    Call Gp_Sp_ColHidden(ss3, 19, True)
'    Call Gp_Sp_ColHidden(ss3, 20, True)
'    Call Gp_Sp_ColHidden(ss3, 21, True)

End Sub
 
Private Sub cmd_exit_Click()
   
    Call Form_Exit
    
End Sub

Private Sub Form_Load()
    
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Gp_FormCenter(Me)
    
    Call Form_Define
  
    Screen.MousePointer = vbDefault
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_Setting(sc2.Item("Spread"), False)
    Call Gp_Sp_Setting(Sc3.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(sc2.Item("Spread"))
    Call Gp_Sp_ReadOnlySet(Sc3.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "K-System.INI", Me.Name, "W")
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc3.Item("Spread"), "K-System.INI", Me.Name)
    
    txt_plt.Text = "B1"
    
    If AKN2040C.Select_Spread.Name = "ss1" Then
       
        txt_ccm_prc_line.Text = "1"
    
    ElseIf AKN2040C.Select_Spread.Name = "ss3" Then
    
        txt_ccm_prc_line.Text = "2"

    Else
        txt_ccm_prc_line.Text = "3"

    End If
    
    AKN2040C.Select_Spread.Row = AKN2040C.Select_Spread_Row
    
    AKN2040C.Select_Spread.Col = 1
    txt_Stlgrd_grp.Text = AKN2040C.Select_Spread.Text
    
    AKN2040C.Select_Spread.Col = 7
    txt_plan_no.Text = AKN2040C.Select_Spread.Text
    
    AKN2040C.Select_Spread.Col = 8
    txt_org_heat_no.Text = AKN2040C.Select_Spread.Text

    If Gf_Sp_Refer(M_CN1, sc1, Mc1, , , False) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gp_Sp_EvenRowBackcolor(ss1)
        ss1.OperationMode = OperationModeNormal
    Else
        cmd_ok.Enabled = False
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc3, Mc1, , , False) Then
        ss3.OperationMode = OperationModeNormal
        Call Gp_Sp_EvenRowBackcolor(ss3)
        Call Spread_Color_Setting(ss3)
    End If
    
    P_Heat_Edt_Seq = 0
    P_Slab_Edt_Seq = 0
    
End Sub

Public Sub Form_Ref()
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Form_Cls()

End Sub

Private Sub Cmd_Ok_Click()
    
    Dim OutParam(1, 4)      As Variant
    Dim ret_Result_ErrMsg   As String
    Dim sQuery              As String
    Dim sMess               As String
    
    Dim adoCmd As ADODB.Command

    On Error GoTo Process_Exec_ERROR
    
    If txt_org_heat_no.Text = "" Or txt_heat_mana_no.Text = "" Then
        Call Gp_MsgBoxDisplay("原炉号与变更炉号必须要选", "W")
        Exit Sub
    End If
    
'    If ss3.MaxRows > 0 Then
'        If P_Slab_Edt_Seq = 0 Then
'            Call Gp_MsgBoxDisplay("请在画面左侧选择移动的位置", "W")
'            Exit Sub
'        End If
'    End If

    sMess = "是否将" & txt_heat_mana_no.Text & "改为" & txt_org_heat_no.Text

    If Not Gf_MessConfirm(sMess, "Q") Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
                                 
    sQuery = "{call AFZ4000P ('B1','C', 'E', '','" & txt_org_heat_no.Text & _
                              "','" & P_Heat_Edt_Seq & "','" & P_Slab_Edt_Seq & _
                              "','" & txt_ccm_prc_line.Text & "','" & sUserID & "',?)}"
                                 
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    M_CN1.BeginTrans
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        M_CN1.RollbackTrans
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay("Error Mesg : " & ret_Result_ErrMsg)
        Set adoCmd = Nothing
        Exit Sub
    End If
    
    Set adoCmd = Nothing
    M_CN1.CommitTrans
    Screen.MousePointer = vbDefault
    
    Call Gp_MsgBoxDisplay("替换炉完了..!!", "I")
    Call AKN2040C.Form_Ref
    Unload Me
    Exit Sub

Process_Exec_ERROR:
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "K-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc3.Item("Spread"), "K-System.INI", Me.Name)
    
    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
    
    Set pContro2 = Nothing
    Set nContro2 = Nothing
    Set iContro2 = Nothing
    Set rContro2 = Nothing
    Set cContro2 = Nothing
    Set aContro2 = Nothing
    Set lContro2 = Nothing
    Set mContro2 = Nothing
    
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
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing
        
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim iRow1, iRow2, iCol   As Integer
    Dim sColor               As String
    
    If Row < 1 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 0
    
    If ss1.Text = "" Then
        
        If iSelect_ss1_Row <> 0 Then
            ss1.Row = iSelect_ss1_Row
            ss1.Col = 0
            ss1.Text = ""
            
            If iSelect_ss1_Row Mod 2 = 0 Then
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iSelect_ss1_Row, iSelect_ss1_Row, , &HFFFFFF)
            Else
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iSelect_ss1_Row, iSelect_ss1_Row, , &HF2F2F2)
            End If
            
        End If
        
        ss1.Row = Row
        ss1.Col = 0
        ss1.Text = "选择"
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
    
    End If
    
    iSelect_ss1_Row = Row
    ss1.Row = Row
    ss1.Col = 1
    txt_heat_mana_no.Text = ss1.Text
    ss1.Col = 15
    P_Heat_Edt_Seq = ss1.Text

    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, , , False)
    ss2.OperationMode = OperationModeNormal
    Call Gp_Sp_EvenRowBackcolor(ss2)

End Sub

Private Sub Spread_Color_Setting(oSpr As vaSpread)

    Dim iRow As Long
    Dim sPlan_Name As String
    Dim sAct_Stlgrd_Grp As String
    Dim sAct_Stlgrd As String
    
    With oSpr
    
        For iRow = 1 To .MaxRows
            
            .Row = iRow
            
            .Col = 6  'PLAN_NAME
            
            If iRow = 1 Then
            
                sPlan_Name = .Text
            
                Call Gp_Sp_Bold(oSpr, "N", iRow)
                
                .Col = 19  'LOCK
                If .Text = "Y" Then
                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                Else
                    .Col = 18  'L2
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                    End If
                End If
            
            ElseIf sPlan_Name <> .Text Then
                
                sPlan_Name = .Text
                
                Call Gp_Sp_Bold(oSpr, "Y", .Row)
            
                .Col = 19  'LOCK
                If .Text = "Y" Then
                    Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HFFC0C0)
                Else
                    .Col = 18  'L2
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF, &HC0FFFF)
                    Else
                        Call Gp_Sp_RowColor(oSpr, iRow, &HFF00FF)
                    End If
                End If
                
            Else
                
                Call Gp_Sp_Bold(oSpr, "N", .Row)
                
                .Col = 19  'LOCK
                If .Text = "Y" Then
                    Call Gp_Sp_RowColor(oSpr, iRow, , &HFFC0C0)
                Else
                    .Col = 18  'L2
                    If .Text = "Y" Then
                        Call Gp_Sp_RowColor(oSpr, iRow, , &HC0FFFF)
                    End If
                End If
            
            End If
            
            .Row = iRow
            .Col = 21  'insert program-id
            
            If .Text <> "" Then
                .Col = 8: .Col2 = 8
                .Row = iRow: .Row2 = iRow
                
                .BlockMode = True
                .ForeColor = vbRed
                .BlockMode = False
            End If
            
        Next iRow
        
        .RowHeight(-1) = 12.54
          
    End With
    
End Sub

Private Sub Gp_Sp_Bold(sPname As Variant, sType As String, iRow As Long)

    With sPname
    
        .Row = iRow: .Row2 = iRow
        .Col = 1: .Col2 = .MaxCols
        
        .BlockMode = True
        
        If sType = "N" Then
            .FontBold = False
        Else
            .FontBold = True
        End If
        
        .BlockMode = False
        
    End With
    
End Sub
'
'Private Sub ss3_Click(ByVal Col As Long, ByVal Row As Long)
'
'    ss3.Row = Row
'
'    ss3.Col = 19   'Lock
'
'    If ss3.Text = "Y" Then Exit Sub
'
'    ss3.Col = 0
'    If ss3.Text = "" Then
'
'        If iSelect_ss3_Row <> 0 Then
'
'            ss3.Row = iSelect_ss3_Row
'
'            ss3.Col = 0
'            ss3.Text = ""
'
'            ss3.Col = 19  'LOCK
'            If ss3.Text = "Y" Then
'                Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, iSelect_ss3_Row, iSelect_ss3_Row, , &HFFC0C0)
'            Else
'                ss3.Col = 18  'L2
'                If ss3.Text = "Y" Then
'                    Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, iSelect_ss3_Row, iSelect_ss3_Row, , &HC0FFFF)
'                Else
'                    If iSelect_ss3_Row Mod 2 = 0 Then
'                        Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, iSelect_ss3_Row, iSelect_ss3_Row, , &HFFFFFF)
'                    Else
'                        Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, iSelect_ss3_Row, iSelect_ss3_Row, , &HF2F2F2)
'                    End If
'                End If
'            End If
'
'        End If
'
'        ss3.Row = Row
'        ss3.Col = 0
'        ss3.Text = "选择"
'        Call Gp_Sp_BlockColor(ss3, 1, ss3.MaxCols, Row, Row, , &HFFFF80)
'
'    End If
'
'    iSelect_ss3_Row = Row
'    ss3.Row = Row
'    ss3.Col = 20
'    P_Slab_Edt_Seq = ss3.Text
'
'End Sub
