VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGC2060C 
   Caption         =   "上下线实绩查询与修改界面_AGC2060C"
   ClientHeight    =   9240
   ClientLeft      =   615
   ClientTop       =   2355
   ClientWidth     =   15405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   15405
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   8430
      Left            =   60
      TabIndex        =   1
      Top             =   735
      Width           =   15345
      _Version        =   393216
      _ExtentX        =   27067
      _ExtentY        =   14870
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
      MaxCols         =   23
      MaxRows         =   2
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AGC2060C.frx":0000
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   660
      Left            =   60
      TabIndex        =   2
      Top             =   45
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   1164
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox cbo_shift 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "AGC2060C.frx":0CDF
         Left            =   14490
         List            =   "AGC2060C.frx":0CEC
         TabIndex        =   14
         Tag             =   "班次"
         Top             =   180
         Width           =   735
      End
      Begin VB.TextBox txt_moda 
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
         Left            =   13560
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   540
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txt_plt 
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
         Left            =   14700
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   4
         Top             =   540
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txt_onoff 
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
         Left            =   14130
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   540
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox txt_mat_no 
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
         Left            =   6930
         MaxLength       =   14
         TabIndex        =   0
         Tag             =   "作业人员"
         Top             =   180
         Width           =   1635
      End
      Begin InDate.ULabel ULabel16 
         Height          =   315
         Left            =   5580
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "物料号"
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
      Begin Threed.SSPanel SSPanel5 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   150
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   196609
         BackColor       =   14737632
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSOption opt_on 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   60
            Width           =   1545
            _ExtentX        =   2725
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
            Caption         =   "在线 -> 离线"
            Value           =   -1
         End
         Begin Threed.SSOption opt_off 
            Height          =   285
            Left            =   1710
            TabIndex        =   7
            Top             =   60
            Width           =   1515
            _ExtentX        =   2672
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
            Caption         =   "离线 -> 在线"
         End
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   150
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   661
         _Version        =   196609
         BackColor       =   14737632
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
         Begin Threed.SSOption opt_mo 
            Height          =   285
            Left            =   150
            TabIndex        =   9
            Top             =   60
            Width           =   705
            _ExtentX        =   1244
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
            Caption         =   "母板"
            Value           =   -1
         End
         Begin Threed.SSOption opt_da 
            Height          =   285
            Left            =   930
            TabIndex        =   10
            Top             =   60
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
            Caption         =   "钢板"
         End
      End
      Begin InDate.ULabel ULabel4 
         Height          =   315
         Left            =   8760
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         Caption         =   "生产日期"
         Alignment       =   1
         BackColor       =   14804173
         BackgroundStyle =   1
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.76
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin InDate.UDate udt_date_fr 
         Height          =   315
         Left            =   10110
         TabIndex        =   12
         Tag             =   "INS_DATE"
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
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
      Begin InDate.UDate udt_date_to 
         Height          =   315
         Left            =   11550
         TabIndex        =   13
         Tag             =   "INS_DATE"
         Top             =   180
         Width           =   1440
         _ExtentX        =   2540
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
      Begin InDate.ULabel ULabel30 
         Height          =   315
         Left            =   13140
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
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
   End
End
Attribute VB_Name = "AGC2060C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      ON/OFF LINE 界面
'-- Program ID        AGC2060C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM SUNG HO
'-- Coder             KIM SUNG HO
'-- Date              2010.7.12
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
 
Dim Proc_Sc As New Collection       'Spread Struc Collection
 
Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection

Dim opt_chk As Boolean

Private Sub Form_Define()

    Dim iCol As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
        Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_moda, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(txt_onoff, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_date_fr, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udt_date_to, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
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
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGC2060C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AGC2060C.P_SMODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
     
    Call Gp_Sp_ColHidden(ss1, 10, True)
    Call Gp_Sp_ColHidden(ss1, 11, True)
    Call Gp_Sp_ColHidden(ss1, 12, True)
    Call Gp_Sp_ColHidden(ss1, 13, True)
    Call Gp_Sp_ColHidden(ss1, 14, True)
    Call Gp_Sp_ColHidden(ss1, ss1.MaxCols - 1, True)
    Call Gp_Sp_ColHidden(ss1, ss1.MaxCols, True)
    
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

    Dim sQuery As String
    
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuTool_ReSet
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    txt_moda.Text = "MP"
    txt_onoff.Text = "I"
    txt_plt.Text = "C1"
    opt_chk = True
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_ReadOnlySet(ss1)
    Call Gf_Sp_Cls(sc1)
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)

    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)

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

Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    Call Gf_Sp_Cls(sc1)
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call MenuTool_ReSet
    TXT_MAT_NO.Enabled = True
    txt_plt.Text = "C1"
    
    If opt_chk Then
        opt_chk = False
        opt_mo.Value = True
        opt_on.Value = True
        opt_chk = True
    Else
        Exit Sub
    End If
    
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Master_Cpy()

End Sub

Public Sub Master_Pst()

End Sub

Public Sub Form_Ref()

    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call Gp_Sp_EvenRowBackcolor(ss1)
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuTool_ReSet
    End If
    
End Sub

Private Sub opt_off_Click(Value As Integer)

    Dim sTxt As String
    Dim sDate_Fr As String
    Dim sDate_To As String
    
    If opt_off.Value Then
        
        If opt_chk Then
        
            opt_chk = False
            sTxt = txt_moda.Text
            sDate_Fr = udt_date_fr.RawData
            sDate_To = udt_date_to.RawData
            
            Call Form_Cls
            
            txt_moda.Text = sTxt
            udt_date_fr.RawData = sDate_Fr
            udt_date_to.RawData = sDate_To
            opt_chk = True
            
        End If
        
        opt_on.ForeColor = &H80000012
        opt_off.ForeColor = &HFF&
        txt_onoff.Text = "O"
    
    End If

End Sub

Private Sub opt_on_Click(Value As Integer)

    Dim sTxt As String
    Dim sDate_Fr As String
    Dim sDate_To As String
    
    If opt_on.Value Then
        
        If opt_chk Then
            
            opt_chk = False
            sTxt = txt_moda.Text
            sDate_Fr = udt_date_fr.RawData
            sDate_To = udt_date_to.RawData
            
            Call Form_Cls
            
            txt_moda.Text = sTxt
            udt_date_fr.RawData = sDate_Fr
            udt_date_to.RawData = sDate_To
            opt_chk = True
            
        End If
        
        opt_on.ForeColor = &HFF&
        opt_off.ForeColor = &H80000012
        txt_onoff.Text = "I"
        
    End If
    
End Sub

Private Sub opt_mo_Click(Value As Integer)

    Dim sTxt As String
    Dim sDate_Fr As String
    Dim sDate_To As String
    
    If opt_mo.Value Then
        
        If opt_chk Then
        
            opt_chk = False
            sTxt = txt_onoff.Text
            sDate_Fr = udt_date_fr.RawData
            sDate_To = udt_date_to.RawData
            
            Call Form_Cls
            
            txt_onoff.Text = sTxt
            udt_date_fr.RawData = sDate_Fr
            udt_date_to.RawData = sDate_To
            opt_chk = True
        
        End If
        
        opt_mo.ForeColor = &HFF&
        opt_da.ForeColor = &H80000012
        txt_moda.Text = "MP"
        
        Call Gp_Sp_ColHidden(ss1, 10, True)
        Call Gp_Sp_ColHidden(ss1, 11, True)
        Call Gp_Sp_ColHidden(ss1, 12, True)
        Call Gp_Sp_ColHidden(ss1, 13, True)
        Call Gp_Sp_ColHidden(ss1, 14, True)
        Call Gp_Sp_ColHidden(ss1, 15, False)
        Call Gp_Sp_ColHidden(ss1, 16, False)
        Call Gp_Sp_ColHidden(ss1, 17, False)
        Call Gp_Sp_ColHidden(ss1, 18, False)
        Call Gp_Sp_ColHidden(ss1, 19, False)
        Call Gp_Sp_ColHidden(ss1, 20, False)
        Call Gp_Sp_ColHidden(ss1, 21, False)

    End If

End Sub

Private Sub opt_da_Click(Value As Integer)

    Dim sTxt As String
    Dim sDate_Fr As String
    Dim sDate_To As String
    
    If opt_da.Value Then
        
        If opt_chk Then
            
            opt_chk = False
            sTxt = txt_onoff.Text
            sDate_Fr = udt_date_fr.RawData
            sDate_To = udt_date_to.RawData
            
            Call Form_Cls
            
            txt_onoff.Text = sTxt
            udt_date_fr.RawData = sDate_Fr
            udt_date_to.RawData = sDate_To
            opt_chk = True
        
        End If
        
        opt_da.ForeColor = &HFF&
        opt_mo.ForeColor = &H80000012
        txt_moda.Text = "PP"
        
        Call Gp_Sp_ColHidden(ss1, 10, False)
        Call Gp_Sp_ColHidden(ss1, 11, False)
        Call Gp_Sp_ColHidden(ss1, 12, False)
        Call Gp_Sp_ColHidden(ss1, 13, False)
        Call Gp_Sp_ColHidden(ss1, 14, False)
        Call Gp_Sp_ColHidden(ss1, 15, True)
        Call Gp_Sp_ColHidden(ss1, 16, True)
        Call Gp_Sp_ColHidden(ss1, 17, True)
        Call Gp_Sp_ColHidden(ss1, 18, True)
        Call Gp_Sp_ColHidden(ss1, 19, True)
        Call Gp_Sp_ColHidden(ss1, 20, True)
        Call Gp_Sp_ColHidden(ss1, 21, True)
        
    End If
    
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

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 0
    
    If ss1.Text <> "" Then
        
        ss1.Text = ""
        
        If Row Mod 2 <> 0 Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HF2F2F2)
        Else
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, Row, Row, , &HFFFFFF)
        End If
        
    Else
        ss1.Text = "Update"
        
        Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.Row, ss1.Row, , CYAN)
        
    End If
    
End Sub

Public Sub Form_Pro()

    Dim iRow As Integer
    
    For iRow = 1 To ss1.MaxRows
    
        ss1.Row = iRow
        ss1.Col = 0
        
        If ss1.Text <> "" Then
        
            ss1.Col = ss1.MaxCols - 1
            ss1.Text = sUserID
            
            ss1.Col = ss1.MaxCols
            If txt_onoff.Text = "I" Then
                ss1.Text = "O"
            Else
                ss1.Text = "I"
            End If
            
        End If
    
    Next iRow
    
    If Gf_Sp_Process(M_CN1, sc1, Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
        Call MenuTool_ReSet
        Call Gp_Sp_EvenRowBackcolor(ss1)
        ss1.OperationMode = OperationModeNormal
    End If
 
End Sub

Public Sub Form_Del()

End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub MenuTool_ReSet()

    With MDIMain.MenuTool
    
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
            
    End With

End Sub
