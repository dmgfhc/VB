VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB4160C 
   BackColor       =   &H00E0E0E0&
   Caption         =   "确定评审处理_ACB4160C"
   ClientHeight    =   9225
   ClientLeft      =   300
   ClientTop       =   2370
   ClientWidth     =   15315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15315
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9135
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   16113
      _Version        =   196609
      SplitterBarWidth=   1
      SplitterBarAppearance=   0
      BorderStyle     =   0
      BackColor       =   12632319
      PaneTree        =   "ACB4160C.frx":0000
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   9135
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   16113
         _Version        =   196609
         SplitterBarWidth=   2
         SplitterBarJoinStyle=   0
         SplitterBarAppearance=   0
         BorderStyle     =   0
         BackColor       =   14737632
         PaneTree        =   "ACB4160C.frx":0032
         Begin Threed.SSPanel SSPanel1 
            Height          =   915
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   1614
            _Version        =   196609
            BackColor       =   14737632
            BevelOuter      =   1
            RoundedCorners  =   0   'False
            FloodShowPct    =   -1  'True
            Begin VB.TextBox txt_rec_sts 
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
               Left            =   14670
               MaxLength       =   1
               TabIndex        =   16
               Tag             =   "处理代码"
               Top             =   90
               Visible         =   0   'False
               Width           =   315
            End
            Begin VB.TextBox txt_est_nm 
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
               Left            =   6795
               MaxLength       =   60
               TabIndex        =   5
               Tag             =   "处理代码"
               Top             =   480
               Width           =   3480
            End
            Begin VB.TextBox txt_est_cd 
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
               Left            =   6240
               MaxLength       =   4
               TabIndex        =   4
               Tag             =   "处理代码"
               Top             =   480
               Width           =   555
            End
            Begin VB.TextBox txt_reason_nm 
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
               Left            =   6795
               MaxLength       =   60
               TabIndex        =   1
               Tag             =   "原因代码"
               Top             =   90
               Width           =   3480
            End
            Begin VB.TextBox txt_reason_cd 
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
               Left            =   6240
               MaxLength       =   4
               TabIndex        =   0
               Tag             =   "原因代码"
               Top             =   90
               Width           =   555
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E0E0E0&
               Height          =   645
               Left            =   10500
               TabIndex        =   7
               Top             =   90
               Width           =   3795
               Begin Threed.SSOption opt_all 
                  Height          =   285
                  Left            =   210
                  TabIndex        =   12
                  Top             =   210
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   503
                  _Version        =   196609
                  Font3D          =   1
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
                  Caption         =   "全部"
               End
               Begin Threed.SSOption opt_wait 
                  Height          =   285
                  Left            =   1140
                  TabIndex        =   13
                  Top             =   210
                  Width           =   1095
                  _ExtentX        =   1931
                  _ExtentY        =   503
                  _Version        =   196609
                  Font3D          =   1
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
                  Caption         =   "等待确定"
                  Value           =   -1
               End
               Begin Threed.SSOption opt_complete 
                  Height          =   285
                  Left            =   2490
                  TabIndex        =   14
                  Top             =   210
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   503
                  _Version        =   196609
                  Font3D          =   1
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
                  Caption         =   "确定完成"
               End
            End
            Begin VB.TextBox txt_slab_no1 
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
               Left            =   1590
               MaxLength       =   10
               TabIndex        =   6
               Tag             =   "板坯号"
               Top             =   95
               Width           =   1395
            End
            Begin InDate.ULabel ULabel4 
               Height          =   315
               Left            =   180
               Top             =   480
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "处理日期"
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
            Begin InDate.ULabel ULabel3 
               Height          =   315
               Left            =   180
               Top             =   90
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "板坯号"
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
            Begin InDate.ULabel ULabel2 
               Height          =   315
               Left            =   4830
               Top             =   90
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "原因代码"
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
               Height          =   315
               Left            =   4830
               Top             =   480
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               Caption         =   "处理代码"
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
            Begin InDate.UDate dpt_est_date_fr 
               Height          =   315
               Left            =   1590
               TabIndex        =   2
               Tag             =   "处理日期"
               Top             =   480
               Width           =   1410
               _ExtentX        =   2487
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
            Begin InDate.UDate dpt_est_date_to 
               Height          =   315
               Left            =   3180
               TabIndex        =   3
               Tag             =   "处理日期"
               Top             =   480
               Width           =   1410
               _ExtentX        =   2487
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "~"
               Height          =   120
               Left            =   3045
               TabIndex        =   15
               Top             =   540
               Width           =   90
            End
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   8190
            Left            =   0
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   945
            Width           =   15225
            _Version        =   393216
            _ExtentX        =   26855
            _ExtentY        =   14446
            _StockProps     =   64
            ColsFrozen      =   1
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
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "ACB4160C.frx":0084
         End
      End
   End
End
Attribute VB_Name = "ACB4160C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   PROCESS MANAGEMENT
'-- Program Name      SLAB DELIBERATION PROCESS CONFIRM EVENT
'-- Program ID        ACB4160C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Kim Sung Ho
'-- Coder             Kim Sung Ho
'-- Date              2009.9.29
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

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sOptFl As Boolean

Private Sub Form_Define()
        
    Dim I As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_slab_no1, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_reason_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_reason_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dpt_est_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dpt_est_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_est_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_est_nm, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_rec_sts, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  
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
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    
    For I = 2 To ss1.MaxCols - 2
        Call Gp_Sp_Collection(ss1, I, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next I
    
    Call Gp_Sp_Collection(ss1, ss1.MaxCols - 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
        Call Gp_Sp_Collection(ss1, ss1.MaxCols, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="ACB4160C.P_REFER", Key:="P-R"
    sc1.Add Item:="ACB4160C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    
    Call Gp_Sp_ColHidden(ss1, ss1.MaxCols - 1, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuTool_ReSet

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
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
    txt_rec_sts.Text = "2"
    
    dpt_est_date_fr.RawData = ""
    dpt_est_date_to.RawData = ""
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "C-System.INI", Me.Name)
    
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

Public Sub Spread_Can()

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call MenuTool_ReSet
        rControl(1).SetFocus
    End If
    
    dpt_est_date_fr.RawData = ""
    dpt_est_date_to.RawData = ""
    txt_rec_sts.Text = "2"
    opt_wait.Value = True

End Sub

Public Sub Form_Ref()

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call Gp_Sp_EvenRowBackcolor(ss1)
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
    End If
    
End Sub

Public Sub Form_Pro()
        
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
        Call Gp_Sp_EvenRowBackcolor(ss1)
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
        ss1.OperationMode = OperationModeNormal
    End If
        
End Sub

Public Sub Form_Ins()
    
End Sub

Public Sub Spread_Cpy()

End Sub

Public Sub Spread_Pst()

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
    
End Sub

Private Sub opt_all_Click(Value As Integer)

    If sOptFl Then
        sOptFl = False
        Exit Sub
    End If
    
    If Not Gf_Sp_Cls(sc1) Then
        sOptFl = True
        If opt_all.ForeColor = &HFF& Then
            opt_all.Value = True
        ElseIf opt_wait.ForeColor = &HFF& Then
            opt_wait.Value = True
        ElseIf opt_complete.ForeColor = &HFF& Then
            opt_complete.Value = True
        End If
        Exit Sub
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet
    rControl(1).SetFocus
    
    If opt_all.Value Then
        opt_all.ForeColor = &HFF&
        opt_wait.ForeColor = &H80000012
        opt_complete.ForeColor = &H80000012
        txt_rec_sts.Text = "A"
    End If
    
End Sub

Private Sub opt_complete_Click(Value As Integer)

    If sOptFl Then
        sOptFl = False
        Exit Sub
    End If
    
    If Not Gf_Sp_Cls(sc1) Then
        sOptFl = True
        If opt_all.ForeColor = &HFF& Then
            opt_all.Value = True
        ElseIf opt_wait.ForeColor = &HFF& Then
            opt_wait.Value = True
        ElseIf opt_complete.ForeColor = &HFF& Then
            opt_complete.Value = True
        End If
        Exit Sub
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet
    rControl(1).SetFocus
    
    If opt_complete.Value Then
        opt_complete.ForeColor = &HFF&
        opt_wait.ForeColor = &H80000012
        opt_all.ForeColor = &H80000012
        txt_rec_sts.Text = "3"
    End If
    
End Sub

Private Sub opt_wait_Click(Value As Integer)

    If sOptFl Then
        sOptFl = False
        Exit Sub
    End If
    
    If Not Gf_Sp_Cls(sc1) Then
        sOptFl = True
        If opt_all.ForeColor = &HFF& Then
            opt_all.Value = True
        ElseIf opt_wait.ForeColor = &HFF& Then
            opt_wait.Value = True
        ElseIf opt_complete.ForeColor = &HFF& Then
            opt_complete.Value = True
        End If
        Exit Sub
    End If
    
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet
    rControl(1).SetFocus
    
    If opt_wait.Value Then
        opt_wait.ForeColor = &HFF&
        opt_complete.ForeColor = &H80000012
        opt_all.ForeColor = &H80000012
        txt_rec_sts.Text = "2"
    End If
    
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Public Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Dim I As Integer
    Dim sReasonCD, sReasonCOMM, sEstCD, sEstCOMM As String
    
    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 0
        
    If ss1.Text <> "Update" Then
        ss1.Col = 4
        If ss1.Text <> "确定完成" Then
            
            ss1.Col = 5
            sReasonCD = ss1.Text
            ss1.Col = 7
            sReasonCOMM = ss1.Text
            ss1.Col = 8
            sEstCD = ss1.Text
            ss1.Col = 10
            sEstCOMM = ss1.Text
            
            If sReasonCD <> "" And sReasonCOMM <> "" And sEstCD <> "" And sEstCOMM <> "" Then
                ss1.Col = 0
                ss1.Text = "Update"
                ss1.Col = ss1.MaxCols - 1
                ss1.Text = sUserID
                ss1.Col = ss1.MaxCols
                ss1.Text = sUsername
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFF80)
            End If
            
        End If
    Else
        ss1.Col = 0
        ss1.Text = ""
        ss1.Col = ss1.MaxCols - 1
        ss1.Text = ""
        ss1.Col = ss1.MaxCols
        ss1.Text = ""
        If Row Mod 2 <> 0 Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HF2F2F2)
        Else
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row)
        End If
    End If
    
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 1
    
    ACB4140C.txt_slab_no1.Text = ss1.Text
    ACB4140C.opt_all.Value = True
    ACB4140C.opt_all.ForeColor = &HFF&
    ACB4140C.opt_in_wait.ForeColor = &H80000012
    ACB4140C.opt_wait.ForeColor = &H80000012
    ACB4140C.opt_complete.ForeColor = &H80000012
    ACB4140C.txt_rec_sts.Text = "A"

    Call ACB4140C.Form_Ref
    Call ACB4140C.ss1_Click(1, 1)

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

Private Sub txt_est_cd_DblClick()

    Call txt_est_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_est_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0018"
        DD.rControl.Add Item:=txt_est_cd
        DD.rControl.Add Item:=txt_est_nm
        
        DD.nameType = "2"
'        DD.sWhere = "AND CD  <>  '9090' "
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_est_cd)) = txt_est_cd.MaxLength Then
            txt_est_nm.Text = Gf_ComnNameFind(M_CN1, "C0018", Trim(txt_est_cd.Text), 2)
        Else
            txt_est_nm.Text = ""
        End If
        
    End If
    
'    If txt_est_cd.Text = "9090" Then
'        txt_est_cd.Text = ""
'        txt_est_nm.Text = ""
'    End If

End Sub

Private Sub txt_reason_cd_DblClick()

    Call txt_reason_cd_KeyUp(vbKeyF4, 0)

End Sub

Private Sub txt_reason_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0017"
        DD.rControl.Add Item:=txt_reason_cd
        DD.rControl.Add Item:=txt_reason_nm
        
        DD.nameType = "2"
'        DD.sWhere = "AND CD  <>  '9090' "
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
    Else
    
        If Len(Trim(txt_reason_cd)) = txt_reason_cd.MaxLength Then
            txt_reason_nm.Text = Gf_ComnNameFind(M_CN1, "C0017", Trim(txt_reason_cd.Text), 2)
        Else
            txt_reason_nm.Text = ""
        End If
        
    End If
    
'    If txt_reason_cd.Text = "9090" Then
'        txt_reason_cd.Text = ""
'        txt_reason_nm.Text = ""
'    End If

End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                  'Row Insert
        .Buttons(8).Enabled = False                  'Row Delete
        .Buttons(9).Enabled = False                  'Row Cancel
        .Buttons(11).Enabled = False                 'Spread Copy
        .Buttons(12).Enabled = False                 'Paste
    End With

End Sub

