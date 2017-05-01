VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CKG2060C 
   Caption         =   "中板厂生产工作记录表_CKG2060C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   12990
   ScaleWidth      =   21480
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSP 
      Height          =   9210
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   16245
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      BackColor       =   14737632
      PaneTree        =   "CKG2060C.frx":0000
      Begin VB.TextBox txt_commect 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Index           =   4
         Left            =   7770
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         ToolTipText     =   "备注4： 其它"
         Top             =   7785
         Width           =   7350
      End
      Begin VB.TextBox txt_commect 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Index           =   3
         Left            =   7770
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         ToolTipText     =   "备注3： 安全检查情况"
         Top             =   6435
         Width           =   7350
      End
      Begin VB.TextBox txt_commect 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Index           =   2
         Left            =   7770
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         ToolTipText     =   "备注2： 质量关键设备运行情况及质量情况"
         Top             =   5070
         Width           =   7350
      End
      Begin VB.TextBox txt_commect 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Index           =   1
         Left            =   0
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         ToolTipText     =   "备注1： 新投用设备使用情况"
         Top             =   7245
         Width           =   7710
      End
      Begin VB.TextBox txt_commect 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Index           =   0
         Left            =   0
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         ToolTipText     =   "停时"
         Top             =   5070
         Width           =   7710
      End
      Begin Threed.SSFrame Single 
         Height          =   570
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   15120
         _ExtentX        =   26670
         _ExtentY        =   1005
         _Version        =   196609
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox TXT_USER 
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
            Left            =   7890
            TabIndex        =   7
            Top             =   150
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.ComboBox CBO_SHIFT 
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
            ItemData        =   "CKG2060C.frx":0112
            Left            =   6345
            List            =   "CKG2060C.frx":011F
            TabIndex        =   4
            Tag             =   "班次"
            Top             =   120
            Width           =   735
         End
         Begin Threed.SSCommand Cmd_Edit 
            Height          =   360
            Left            =   10335
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   90
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   635
            _Version        =   196609
            Font3D          =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "更新数据"
         End
         Begin InDate.UDate txt_DATE 
            Height          =   315
            Left            =   2595
            TabIndex        =   3
            Tag             =   "记录日期"
            Top             =   120
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
            Left            =   480
            Top             =   120
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            Caption         =   "记录日期"
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
            Left            =   5130
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
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
      Begin FPSpread.vaSpread ss1 
         Height          =   4380
         Left            =   0
         TabIndex        =   5
         Top             =   630
         Width           =   11205
         _Version        =   393216
         _ExtentX        =   19764
         _ExtentY        =   7726
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
         MaxCols         =   25
         MaxRows         =   24
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKG2060C.frx":012F
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4380
         Left            =   11265
         TabIndex        =   6
         Top             =   630
         Width           =   3855
         _Version        =   393216
         _ExtentX        =   6800
         _ExtentY        =   7726
         _StockProps     =   64
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BackColorStyle  =   1
         ColHeaderDisplay=   1
         DisplayColHeaders=   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         MaxCols         =   6
         MaxRows         =   12
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "CKG2060C.frx":10DF
      End
   End
End
Attribute VB_Name = "CKG2060C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Mill System
'-- Program Name      PROD REPORT
'-- Program ID        CKG2060C
'-- Designer          YANGMENG
'-- Coder             YANGMENG
'-- Date              2008.05.23
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
Public QueryYN As Boolean

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

Dim pColumn4 As New Collection      'Spread Primary Key Collection
Dim nColumn4 As New Collection      'Spread necessary Column Collection
Dim mColumn4 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn4 As New Collection      'Spread Insert Column Collection
Dim aColumn4 As New Collection      'Master -> Spread Column Collection
Dim lColumn4 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
Dim sc4 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(txt_DATE, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(CBO_SHIFT, "p", "n", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_commect(0), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_commect(1), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_commect(2), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_commect(3), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_commect(4), " ", " ", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_USER, " ", " ", " ", "i", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    'MASTER Collection
    Mc1.Add Item:="CKG2060C.P_MODIFY", Key:="P-M"
    Mc1.Add Item:="CKG2060C.P_REFER", Key:="P-R"
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
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="CKG2060C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="CKG2060C.P_SMODIFY", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc1"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

     
    'Spread_Collection
    Sc2.Add Item:=ss2, Key:="Spread"
'    Sc2.Add Item:="CKG2060C.P_SREFER2", Key:="P-R"
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
        
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "Z-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Z-System.INI", Me.Name)
    
    Call Gp_Sp_ColHidden(ss1, 24, True)
    Call Gp_Sp_ColHidden(ss1, 25, True)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
       Cmd_Edit.Enabled = True
    End If

    txt_DATE.RawData = Format(Date - 1, "yyyymmdd")
    CBO_SHIFT.ListIndex = 0
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "K-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "K-System.INI", Me.Name)
    
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
    
    Set iColumn4 = Nothing
    Set pColumn4 = Nothing
    Set lColumn4 = Nothing
    Set nColumn4 = Nothing
    Set mColumn4 = Nothing
    Set aColumn4 = Nothing
   
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set Sc3 = Nothing
    Set sc4 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
End Sub

Public Sub Form_Cls()

    Dim iRow  As Long
    Dim iCol  As Long

    Call Gf_Sp_Cls(Sc1)
    Call ss2_clear
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    txt_DATE.Enabled = True
    CBO_SHIFT.Enabled = True
    
    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    
End Sub

Public Sub Form_Ref()
    
    If Trim(txt_DATE.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_DATE.Tag + "必须正确输入")
        Exit Sub
    End If
    
    If Trim(CBO_SHIFT.Text) <> "1" And Trim(CBO_SHIFT.Text) <> "2" And Trim(CBO_SHIFT.Text) <> "3" Then
        Call Gp_MsgBoxDisplay(CBO_SHIFT.Tag + "必须正确输入")
        Exit Sub
    End If

    Call Form_Cls
        
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl")) Then
        Call Ss2_Data_Refer
        Call Gf_Ms_Refer(M_CN1, Mc1, , , False)
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
             
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
    With MDIMain.MenuTool
        .Buttons(5).Enabled = False                 'Delete
        .Buttons(6).Enabled = False                 'Separator
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
    End With
    
End Sub

Public Sub Form_Exc()

'    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call ExcelPrn
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("SC1")("Spread"), Mode)
        ss1.Col = 25
        ss1.Text = sUserID
    End If

End Sub

Public Sub Form_Pro()
Dim a As String
    TXT_USER = sUserID
    a = MsgBox("您确定要修改当班工作记录吗？", vbQuestion + vbYesNo, "系统提示信息")
    If a = vbNo Then
       Exit Sub
    End If
    
    If Gf_Ms_Process(M_CN1, Mc1, sAuthority) And Gf_Sp_Process(M_CN1, Proc_Sc("SC1"), Mc1) Then
        txt_DATE.Enabled = True
        CBO_SHIFT.Enabled = True
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        With MDIMain.MenuTool
            .Buttons(5).Enabled = False                 'Delete
            .Buttons(6).Enabled = False                 'Separator
            .Buttons(7).Enabled = False                 'Row Insert
            .Buttons(8).Enabled = False                 'Row Delete
            .Buttons(9).Enabled = False                 'Row Cancel
            .Buttons(10).Enabled = False                'Separator
            .Buttons(11).Enabled = False                'Copy
            .Buttons(12).Enabled = False                'Paste
        End With
    End If
    
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

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Public Sub Zero_Cls()
    Dim iRow  As Long
    Dim iCol  As Long
    
    For iRow = 1 To ss1.MaxRows
        ss1.ROW = iRow
        For iCol = 1 To ss1.MaxCols
            ss1.Col = iCol
            If Val(ss1.Text & "") = 0 Then
                ss1.Text = ""
            End If
        Next iCol
    Next iRow

End Sub
Private Sub Cmd_Edit_Click()
    'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    Dim sQuery As String
          
    If Trim(txt_DATE.Text) = "" Then
        Call Gp_MsgBoxDisplay(txt_DATE.Tag + "必须输入")
        Exit Sub
    End If

    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sQuery = "{call CKG2060P ('" + Trim(Format(txt_DATE.Text, "YYYYMMDD")) + "','" + Trim(CBO_SHIFT.Text) + "',?)}"

    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
            
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        strRet_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & strRet_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        
        Call Gp_MsgBoxDisplay("更新成功..!!", "I")
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("更新失败！！")

End Sub


Private Sub ExcelPrn()
    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDate           As String
    
    Dim sShift          As String
    Dim sgroup          As String
    
    Dim sExlRange       As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\CKG2060C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    If Trim(CBO_SHIFT.Text) = "1" Then
       sShift = "大"
    ElseIf Trim(CBO_SHIFT.Text) = "2" Then
       sShift = "白"
    ElseIf Trim(CBO_SHIFT.Text) = "3" Then
       sShift = "小"
    End If
    
    ss1.Col = 1
    ss1.ROW = 1
    sgroup = ss1.Text
    
    sDate = Format(txt_DATE.Text, "YYYYMMDD")
    xlApp.Range("A2").Value = "报表日期：" + Left(sDate, 4) + "年" + Mid(sDate, 5, 2) + "月" + Mid(sDate, 7, 2) + "日"
    xlApp.Range("D2").Value = "班次：" + sShift
    xlApp.Range("F2").Value = "班别：" + sgroup
    xlApp.Range("X30").Value = "制表日期：" + Format(Now, "YYYY-MM-DD HH:MM:SS")
    xlApp.Range("U30").Value = "制表人：" + sUserID

    Clipboard.Clear
    ss1.SetSelection 2, 1, ss1.MaxCols - 2, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("A6").Select
    xlApp.ActiveSheet.Paste
    
    xlApp.Range("A14").Value = "合计"
    
    xlApp.Range("B15").Value = txt_commect(0).Text
    xlApp.Range("B23").Value = txt_commect(1).Text
    xlApp.Range("P15").Value = txt_commect(2).Text
    xlApp.Range("P23").Value = txt_commect(3).Text
    xlApp.Range("P27").Value = txt_commect(4).Text

    Clipboard.Clear
    ss2.SetSelection 3, 3, 6, 6
    ss2.ClipboardCopy
    xlApp.Range("Y5").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss2.SetSelection 3, 9, 6, 12
    ss2.ClipboardCopy
    xlApp.Range("Y11").Select
    xlApp.ActiveSheet.Paste

    ss1.ClearSelection
    ss2.ClearSelection

    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Public Sub Sp_ColLock(sPname As Variant, ColNum As Variant, RowNum As Variant, LockType As Boolean)

    With sPname
        .Protect = True
        .Col = ColNum: .Col2 = ColNum
        .ROW = RowNum: .Row2 = RowNum
        
        .BlockMode = True
        .Lock = LockType
        .BlockMode = False
    End With
    
End Sub
Public Function ss2_clear() As Boolean
    
       ss2.ClearRange 3, 3, 6, 6, True
       ss2.ClearRange 3, 9, 6, 12, True
    
End Function
Public Sub Ss2_Data_Refer()

On Error GoTo Ss2_Display_Error

    Dim sTdate      As String
    Dim sBfdate     As String
    Dim sQuery      As String
    Dim IDc         As Integer

    Dim dNewDate    As Date
    Dim dEndDate    As Date
    Dim lDiff       As Long
    Dim dPlanWgt    As Double
    Dim dActWgt     As Double

    Dim AdoRs As ADODB.Recordset

    Set AdoRs = New ADODB.Recordset
  
    sQuery = "SELECT            *                                       " & vbCrLf
    sQuery = sQuery & "   FROM  gp_rpt_shift_sum3_c                     " & vbCrLf
    sQuery = sQuery & "  WHERE  MILL_DATE      =  '" & txt_DATE.RawData & "'" & vbCrLf
    sQuery = sQuery & "    AND  MILL_SHIFT     =  '" & Trim(CBO_SHIFT.Text) & "'" & vbCrLf

    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Do Until AdoRs.EOF
       
        With ss2
    
            .Col = 3:   .ROW = 3:    .Text = IIf(IsNull(AdoRs.Fields(28)), "", AdoRs.Fields(28))
                        .ROW = 4:    .Text = Val(AdoRs.Fields(5) & "")
                        .ROW = 5:    .Text = Val(AdoRs.Fields(6) & "")
                        .ROW = 6:    .Text = Val(AdoRs.Fields(7) & "")

                        .ROW = 9:    .Text = IIf(IsNull(AdoRs.Fields(32)), "", AdoRs.Fields(32))
                        .ROW = 10:   .Text = Val(AdoRs.Fields(15) & "")
                        .ROW = 11:   .Text = Val(AdoRs.Fields(16) & "")
                        .ROW = 12:   .Text = Val(AdoRs.Fields(17) & "")
            .Col = 4:   .ROW = 3:    .Text = IIf(IsNull(AdoRs.Fields(29)), "", AdoRs.Fields(29))
                        .ROW = 4:    .Text = Val(AdoRs.Fields(5) & "")
                        .ROW = 5:    .Text = Val(AdoRs.Fields(6) & "")
                        .ROW = 6:    .Text = Val(AdoRs.Fields(7) & "")

                        .ROW = 9:    .Text = IIf(IsNull(AdoRs.Fields(33)), "", AdoRs.Fields(33))
                        .ROW = 10:   .Text = Val(AdoRs.Fields(15) & "")
                        .ROW = 11:   .Text = Val(AdoRs.Fields(16) & "")
                        .ROW = 12:   .Text = Val(AdoRs.Fields(17) & "")
            .Col = 5:   .ROW = 3:    .Text = IIf(IsNull(AdoRs.Fields(30)), "", AdoRs.Fields(30))
                        .ROW = 4:    .Text = Val(AdoRs.Fields(10) & "")
                        .ROW = 5:    .Text = Val(AdoRs.Fields(11) & "")
                        .ROW = 6:    .Text = Val(AdoRs.Fields(12) & "")

                        .ROW = 9:    .Text = IIf(IsNull(AdoRs.Fields(34)), "", AdoRs.Fields(34))
                        .ROW = 10:   .Text = Val(AdoRs.Fields(20) & "")
                        .ROW = 11:   .Text = Val(AdoRs.Fields(21) & "")
                        .ROW = 12:   .Text = Val(AdoRs.Fields(22) & "")
            .Col = 6:   .ROW = 3:    .Text = IIf(IsNull(AdoRs.Fields(31)), "", AdoRs.Fields(31))
                        .ROW = 4:    .Text = Val(AdoRs.Fields(10) & "")
                        .ROW = 5:    .Text = Val(AdoRs.Fields(11) & "")
                        .ROW = 6:    .Text = Val(AdoRs.Fields(12) & "")

                        .ROW = 9:    .Text = IIf(IsNull(AdoRs.Fields(35)), "", AdoRs.Fields(35))
                        .ROW = 10:   .Text = Val(AdoRs.Fields(20) & "")
                        .ROW = 11:   .Text = Val(AdoRs.Fields(21) & "")
                        .ROW = 12:   .Text = Val(AdoRs.Fields(22) & "")

        End With
    
        AdoRs.MoveNext
    Loop
    
    AdoRs.Close
    
    Exit Sub

Ss2_Display_Error:
    
    Set AdoRs = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Ss2_Display_Error : " & Error)
    
End Sub





