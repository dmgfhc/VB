VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AZD1010C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "数据库/项目_AZD1010C"
   ClientHeight    =   7320
   ClientLeft      =   555
   ClientTop       =   3030
   ClientWidth     =   12270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   12270
   WindowState     =   2  'Maximized
   Begin VB.Frame framProcess 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Upload Processing "
      Height          =   885
      Left            =   2055
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   6735
      Begin MSComctlLib.ProgressBar ProgBar 
         Height          =   255
         Left            =   210
         TabIndex        =   5
         Top             =   525
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProcess 
         BackColor       =   &H00E0E0E0&
         Caption         =   "aaaaaaaa"
         Height          =   180
         Left            =   225
         TabIndex        =   6
         Top             =   300
         Width           =   6255
      End
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   150
      Top             =   105
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "数据库隶属"
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
   Begin VB.TextBox txt_biz_area 
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
      Left            =   1545
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "项目隶属"
      Top             =   105
      Width           =   465
   End
   Begin VB.TextBox txt_biz_area_name 
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
      Left            =   2010
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "项目隶属"
      Top             =   105
      Width           =   4830
   End
   Begin VB.TextBox txt_table_id 
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
      Left            =   8415
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "项目代码"
      Top             =   105
      Width           =   1950
   End
   Begin VB.TextBox txt_table_name 
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
      Left            =   10380
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "项目代码"
      Top             =   105
      Width           =   4725
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   7020
      Top             =   105
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "数据库代码"
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
   Begin Threed.SSCommand cmd_UploadFile 
      Height          =   330
      Left            =   3435
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   510
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   3
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "上载 File名:"
   End
   Begin Threed.SSCommand cmd_Upload 
      Height          =   330
      Left            =   2040
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   510
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   3
      ForeColor       =   0
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "上载 File"
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8280
      Left            =   135
      TabIndex        =   10
      Top             =   900
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   14605
      _Version        =   196609
      SplitterBarWidth=   4
      SplitterBarJoinStyle=   0
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   16761087
      PaneTree        =   "AZD1010C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   8280
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   7725
         _Version        =   393216
         _ExtentX        =   13626
         _ExtentY        =   14605
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
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AZD1010C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   8280
         Left            =   7785
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   0
         Width           =   7185
         _Version        =   393216
         _ExtentX        =   12674
         _ExtentY        =   14605
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
         MaxRows         =   2
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AZD1010C.frx":03D7
      End
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   285
      Left            =   14160
      TabIndex        =   11
      Top             =   555
      Width           =   990
      _ExtentX        =   1746
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
      Caption         =   "项目表"
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   285
      Left            =   165
      TabIndex        =   12
      Top             =   555
      Width           =   1320
      _ExtentX        =   2328
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
      Caption         =   "数据库表"
      Value           =   1
   End
   Begin VB.Label lblFileName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   4935
      TabIndex        =   9
      Top             =   540
      Width           =   5640
   End
End
Attribute VB_Name = "AZD1010C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       System Management
'-- Sub_System Name   TABLE Management
'-- Program Name      TABLE/COLUMN UPDATE
'-- Program ID        AZD1010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SOO HEON
'-- Coder             KIM SOO HEON
'-- Date              2005.12.2
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

Dim pControl2 As New Collection     'Master Primary Key Collection
Dim nControl2 As New Collection     'Master Necessary Collection
Dim mControl2 As New Collection     'Master Maxlength check Collection
Dim iControl2 As New Collection     'Master Insert Collection
Dim rControl2 As New Collection     'Master Refer Collection
Dim cControl2 As New Collection     'Master Copy Collection
Dim aControl2 As New Collection     'Master -> Spread Collection
Dim lControl2 As New Collection     'Master Lock Collection

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
Dim Sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_biz_area, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_biz_area_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   
       Call Gp_Ms_Collection(txt_table_id, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_table_name, " ", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
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
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AZD1010C.P_MODIFY1", Key:="P-M"
    Sc1.Add Item:="AZD1010C.P_REFER1", Key:="P-R"
    Sc1.Add Item:="AZD1010C.P_ONEROW1", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Call Gp_Sp_Collection(ss2, 1, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, "p", "n", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="AZD1010C.P_MODIFY2", Key:="P-M"
    Sc2.Add Item:="AZD1010C.P_ONEROW2", Key:="P-O"
    Sc2.Add Item:="AZD1010C.P_REFER2", Key:="P-R"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=2, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = "◎"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
       
End Sub

Public Sub MenuToolSet()

    MDIMain.MenuTool.Buttons(7).Enabled = False
    MDIMain.MenuTool.Buttons(8).Enabled = False
'    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

End Sub

Private Sub cmd_UploadFile_Click()
    If Chk_ss1.Value = -1 Then
        ss1.MaxRows = 0
    Else
        ss2.MaxRows = 0
    End If
    
    frmFileList.Tag = Me.Name
    frmFileList.Show 1
    
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
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"))
    Call Gp_Sp_Setting(Sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(Sc2)
    
    Call Gp_Spl_SizeGet(SSSplitter1, "Z-System.INI", Me.Name, "W")
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "Z-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Sc2.Item("Spread"), "Z-System.INI", Me.Name)
    Call MenuToolSet
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Spl_SizeSet(SSSplitter1, "Z-System.INI", Me.Name)
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "Z-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Sc2.Item("Spread"), "Z-System.INI", Me.Name)
    
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
    Set Sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    If Chk_ss1.Value = -1 Then
        Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
    Else
        Call Gp_Sp_Cancel(M_CN1, Sc2)
    End If
    
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Sc2) Then
        If Gf_Sp_Cls(Sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call MenuToolSet
        End If
    End If

End Sub

Public Sub Form_Ref()

On Error Resume Next

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
        Call Gf_Sp_Cls(Sc2)
        Call MenuToolSet
    End If
            
End Sub

Public Sub Form_Pro()
        
    If Chk_ss1.Value = -1 Then
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call Gf_Sp_Cls(Sc2)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    Else
        If Gf_Sp_Process(M_CN1, Sc2, , True) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    End If
    Call MenuToolSet
    
End Sub

Public Sub Form_Ins()
    
'
End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("Sc"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    ss1.Row = Row
    
    If Gf_Sp_ProceExist(Sc2.Item("Spread")) Then Exit Sub
    
    ss1.Col = 1
    txt_table_id.Text = ss1.Text
    ss1.Col = 2
    txt_table_name.Text = ss1.Text
    
    Call Gf_Sp_Refer(M_CN1, Sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
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


Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(ss2, Mode)
    End If
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

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

Private Sub Chk_ss1_Click(Value As Integer)
    Dim iDr As Long
    
    If Chk_ss1.Value = ssCBUnchecked Then
       If Chk_ss2.Value = ssCBUnchecked Then
            Chk_ss1.Value = ssCBChecked
       End If
       Exit Sub
    End If

    If Not Gf_Sp_ProceExist(Sc2.Item("Spread")) Then
        Chk_ss1.ForeColor = &HFF&
        Chk_ss2.ForeColor = &H808080
        Chk_ss2.Value = ssCBUnchecked
        For iDr = 1 To ss2.MaxRows
            ss2.Row = iDr:   ss2.Col = 0:    ss2.Text = ""
        Next iDr
    Else
        Chk_ss1.Value = ssCBUnchecked
        Chk_ss2.Value = ssCBChecked
    End If
        
End Sub

Private Sub Chk_ss2_Click(Value As Integer)
    Dim iDr As Long
    
    If Chk_ss2.Value = ssCBUnchecked Then
        If Chk_ss1.Value = ssCBUnchecked Then
            Chk_ss2.Value = ssCBChecked
        End If
        Exit Sub
    End If
    
    If Not Gf_Sp_ProceExist(Sc1.Item("Spread")) Then
        Chk_ss1.ForeColor = &H808080
        Chk_ss2.ForeColor = &HFF&
        Chk_ss1.Value = ssCBUnchecked
        For iDr = 1 To ss1.MaxRows
            ss1.Row = iDr:   ss1.Col = 0:    ss1.Text = ""
        Next iDr
    Else
        Chk_ss2.Value = ssCBUnchecked
        Chk_ss1.Value = ssCBChecked
    End If
        
End Sub

Private Sub txt_biz_area_DblClick()

    Call txt_biz_area_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_biz_area_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "Z0001"
        DD.rControl.Add Item:=txt_biz_area
        DD.rControl.Add Item:=txt_biz_area_name
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(txt_biz_area)) = txt_biz_area.MaxLength Then
        txt_biz_area_name.Text = Gf_ComnNameFind(M_CN1, "Z0001", Trim(txt_biz_area.Text), 2)
    Else
        txt_biz_area_name.Text = ""
    End If

End Sub


Private Sub cmd_Upload_Click()
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim iCount          As Integer
    Dim sMsg            As String
    Dim iRow            As Integer
    Dim iCol            As Integer
    Dim iCnt            As Integer
    Dim iXrow           As Integer
    Dim iXCol           As Integer
    Dim iXCnt           As Integer
    Dim iXStart_Row     As Integer
    Dim iDr             As Integer
    Dim sTableName      As String
    
    If Trim(lblFileName.Caption) = "" Then
        MsgBox "Upload File not Selected ", vbCritical + vbOKOnly
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    lblProcess.Caption = ""
    framProcess.Visible = True
    
    sMsg = ""

    On Error GoTo ErrProc
    
    Set xlApp = GetObject("", "Excel.Application")
    If Err.Number = 429 Then
        Set xlApp = CreateObject("", "Excel.Application")
    End If
    
    xlApp.Workbooks.Open (Trim(lblFileName.Caption))
    
'    If Chk_ss1.Value = -1 Then
'        Set xlSheet = xlApp.Worksheets(1)
'    Else
        Set xlSheet = xlApp.Worksheets(2)
'    End If
    
    iDr = 0
    iXCnt = 0
    iXStart_Row = 3
    
    iXrow = iXStart_Row + iXCnt
    
    While CStr(xlSheet.cells(iXrow, 1)) > " "
        iXCnt = iXCnt + 1
        iXrow = iXrow + 1
    Wend
       
    ProgBar.Min = 0
    ProgBar.Max = iXCnt
            
    For iRow = 1 To iXCnt
        iXrow = iXStart_Row + iRow - 1
        
        ProgBar.Value = iRow
        lblProcess.Caption = "Excel Reading.....  " & CStr(iRow) & " / " & CStr(iXCnt)
        DoEvents
        
        If Chk_ss1.Value = -1 Then
            If sTableName <> xlSheet.cells(iXrow, 2) Then
                iDr = iDr + 1
                ss1.MaxRows = iDr
                ss1.Row = iDr
                ss1.Col = 0:       ss1.Text = "Input"
                ss1.Col = 1:       ss1.Text = xlSheet.cells(iXrow, 2)           ' Table ID
                ss1.Col = 2:       ss1.Text = Trim(xlSheet.cells(iXrow, 1))     ' Table NAME
                
                sTableName = xlSheet.cells(iXrow, 2)
            End If
        Else
            iDr = iDr + 1
            ss2.MaxRows = iDr
            ss2.Row = iDr
            ss2.Col = 0:       ss2.Text = "Input"
            ss2.Col = 1:       ss2.Text = xlSheet.cells(iXrow, 2)           ' Table ID
            ss2.Col = 2:       ss2.Text = xlSheet.cells(iXrow, 4)           ' COLUMN ID
            ss2.Col = 3:       ss2.Text = Trim(xlSheet.cells(iXrow, 3))     ' COLUMN NAME
        End If
            
    Next iRow
        
    framProcess.Visible = False
    
    If iDr > 0 Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    xlApp.ActiveWorkbook.Close Trim(lblFileName.Caption)
    
    xlApp.Quit
    Set xlSheet = Nothing
    Set xlApp = Nothing
    
    Screen.MousePointer = vbDefault
    
    If Len(Trim(sMsg)) > 0 Then MsgBox sMsg, vbCritical

    Exit Sub

ErrProc:
    If Err.Number = 429 Then
        MsgBox "Microsoft Excel Program Not Installed"
    Else
        MsgBox Err.Number & Err.Description
    End If
    
    Set xlSheet = Nothing
    xlApp.ActiveWorkbook.Close False
    xlApp.Quit
    Set xlApp = Nothing
    
    Screen.MousePointer = vbDefault

End Sub

