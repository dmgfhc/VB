VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQD0080C 
   Caption         =   "提货单LOT号编辑_AQD0080C"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel SSPanel3 
      Height          =   6375
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   11245
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin FPSpread.vaSpread ss1 
         Height          =   5715
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   15135
         _Version        =   393216
         _ExtentX        =   26696
         _ExtentY        =   10081
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
         MaxCols         =   11
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQD0080C.frx":0000
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   795
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   1402
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   60
         Top             =   30
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "查询条件"
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
         ForeColor       =   16576
      End
      Begin VB.TextBox txt_PONO 
         Height          =   315
         Left            =   6120
         TabIndex        =   8
         Top             =   420
         Width           =   2235
      End
      Begin VB.TextBox txt_ISP_SHP_NO 
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Top             =   420
         Width           =   2235
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   0
         Left            =   60
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "提货单号"
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
         Index           =   3
         Left            =   4380
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "合同号"
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   4577
      _Version        =   196609
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.TextBox txt_REMARK 
         Height          =   315
         Left            =   1860
         MaxLength       =   100
         TabIndex        =   15
         Top             =   1800
         Width           =   6555
      End
      Begin VB.TextBox txt_PONO_LOT 
         Height          =   315
         Left            =   1860
         TabIndex        =   14
         Top             =   1020
         Width           =   2235
      End
      Begin VB.TextBox txt_COND_SUPPLY_CN 
         Height          =   315
         Left            =   6180
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1410
         Width           =   2235
      End
      Begin VB.TextBox txt_COND_SUPPLY_EN 
         Height          =   315
         Left            =   1860
         MaxLength       =   20
         TabIndex        =   12
         Top             =   1410
         Width           =   2235
      End
      Begin Threed.SSCommand SSCom_Add 
         Height          =   345
         Left            =   3930
         TabIndex        =   11
         Top             =   150
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
         _Version        =   196609
         Caption         =   "自动录入开始"
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   60
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "自动录入"
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
         ForeColor       =   16576
      End
      Begin VB.TextBox txt_Last 
         Height          =   315
         Left            =   6180
         MaxLength       =   11
         TabIndex        =   4
         Top             =   630
         Width           =   2235
      End
      Begin VB.TextBox txt_First 
         Height          =   315
         Left            =   1860
         MaxLength       =   11
         TabIndex        =   3
         Top             =   630
         Width           =   2235
      End
      Begin VB.TextBox txt_Add 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3180
         TabIndex        =   1
         Text            =   "1"
         Top             =   150
         Width           =   540
      End
      Begin MSComCtl2.UpDown UD_Add 
         Height          =   315
         Left            =   1860
         TabIndex        =   2
         Top             =   150
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txt_Add"
         BuddyDispid     =   196617
         OrigLeft        =   5850
         OrigTop         =   750
         OrigRight       =   6090
         OrigBottom      =   1095
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   0   'False
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   1
         Left            =   60
         Top             =   630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "起始提货单号"
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
      Begin Threed.SSRibbon ssRib_Add 
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   150
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   255
         Caption         =   "累加步长"
         ButtonStyle     =   1
      End
      Begin InDate.ULabel ULabel1 
         Height          =   315
         Index           =   2
         Left            =   4380
         Top             =   630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "终止提货单号"
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
         Index           =   10
         Left            =   60
         Top             =   1410
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "交货状态(英文)"
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
         Index           =   12
         Left            =   4380
         Top             =   1410
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "交货状态(中文)"
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
         Index           =   4
         Left            =   60
         Top             =   1020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "合同号"
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
         Index           =   11
         Left            =   60
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         Caption         =   "备注"
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
         Height          =   255
         Index           =   5
         Left            =   120
         Top             =   2160
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   450
         Caption         =   "注意：备注栏最多可输入50个汉字或100个字母和数字"
         Alignment       =   0
         BackgroundStyle =   1
         BorderEffect    =   0
         ChiselText      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   7
      Left            =   6240
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "合同号"
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
Attribute VB_Name = "AQD0080C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      提单对应合同LOT号编辑
'-- Program ID        AQD0080C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Li Qing Yu
'-- Coder             Li Qing Yu
'-- Date              2006.10.16
'-- Description       提单对应合同LOT号编辑
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
Dim bPrintCheck As Boolean

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_ISP_SHP_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_PONO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_COND_SUPPLY_EN, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_COND_SUPPLY_CN, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_REMARK, " ", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       
    
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
    
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQD0080C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQD0080C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQD0080C.P_ONEROW", Key:="P-O"
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
        
    Screen.MousePointer = vbDefault

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
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        pControl(1).SetFocus
    End If
    txt_First.Text = ""
    txt_Last.Text = ""
    txt_PONO_LOT.Text = ""
    txt_COND_SUPPLY_EN.Text = ""
    txt_COND_SUPPLY_CN.Text = ""
    txt_REMARK.Text = ""

End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
        If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

            Exit Sub
        End If
            
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
 
         If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
         End If

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
    ss1.SetFocus
    

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
    
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
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim i As Long
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

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
    
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


Private Sub SSCom_Add_Click()
    Dim lFirst      As Long
    Dim lLast       As Long
    Dim iMin        As Integer
    Dim iMax        As Integer
    Dim iAdd        As Integer
    Dim sISP_NO     As String
    Dim sPONO_LOT   As String
    Dim i           As Integer
    
    If Trim(txt_First.Text) = "" Or Len(Trim(txt_First.Text)) = 0 Then
        Call MsgBox("请输入起始提货单号！", vbOKOnly)
    ElseIf Len(Trim(txt_First.Text)) > 0 And Len(Trim(txt_First.Text)) < 11 Then
        Call MsgBox("起始提货单号输入错误！", vbOKOnly)
        Exit Sub
    Else
        iMin = Val(Mid(Trim(txt_First.Text), 9, 3))
        sISP_NO = Mid(Trim(txt_First.Text), 1, 8)
    End If
    
    If Trim(txt_Last.Text) = "" Or Len(Trim(txt_Last.Text)) = 0 Then
        Call MsgBox("请输入终止提货单号！", vbOKOnly)
    ElseIf Len(Trim(txt_Last.Text)) > 0 And Len(Trim(txt_Last.Text)) < 11 Then
        Call MsgBox("终止提货单号输入错误！", vbOKOnly)
        Exit Sub
    Else
        iMax = Val(Mid(Trim(txt_Last.Text), 9, 3))
    End If
    
'    If Trim(txt_PONO_LOT.Text) = "" Or Len(Trim(txt_PONO_LOT.Text)) = 0 Then
'        Call MsgBox("请输入合同及LOT号！", vbOKOnly)
'        Exit Sub
'    Else
'        sPONO_LOT = Trim(txt_PONO_LOT.Text)
'    End If
        
    
    If txt_Add.Text = 0 Or Trim(txt_Add.Text) = "" _
        Or Len(Trim(txt_Add.Text)) = 0 Then
        iAdd = 1
    Else
        iAdd = Val(Trim(txt_Add.Text))
    End If
    
    lFirst = ss1.MaxRows
    
    For i = iMin To iMax Step iAdd
        With ss1
'            .Row = .MaxRows
            Call Gp_Sp_Ins(Proc_Sc("Sc"))
            Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 7)
'            lFirst = lFirst + 1
            .Row = .ActiveRow 'lFirst
            .Col = 2
                .Text = Trim(sISP_NO + Format(str(i), "000"))
            .Col = 3
                .Text = sPONO_LOT
            .Col = 4
                .Text = txt_COND_SUPPLY_EN.Text
            .Col = 5
            
                .Text = txt_COND_SUPPLY_CN.Text
            .Col = 6
                .Text = txt_REMARK.Text
        
        End With
    Next

End Sub

Private Sub ssRib_Add_Click(Value As Integer)
    
    If Value = False Then
        txt_Add.Enabled = False
        UD_Add.Enabled = False
    Else
        txt_Add.Enabled = True
        UD_Add.Enabled = True
    End If
    
End Sub



Private Sub txt_First_Change()
    Dim iSelStar As Integer
    
    iSelStar = Len(txt_First.Text)
    txt_First.Text = UCase(txt_First.Text)
    txt_First.SelStart = iSelStar
    
End Sub

Private Sub txt_Last_Change()
    Dim iSelStar As Integer
    
    iSelStar = Len(txt_Last.Text)
    txt_Last.Text = UCase(txt_Last.Text)
    txt_Last.SelStart = iSelStar

End Sub

Private Sub txt_PONO_LOT_Change()
    Dim iSelStar As Integer
    
    iSelStar = Len(txt_PONO_LOT.Text)
    txt_PONO_LOT.Text = UCase(txt_PONO_LOT.Text)
    txt_PONO_LOT.SelStart = iSelStar

End Sub
