VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACZ1011C 
   BackColor       =   &H8000000A&
   Caption         =   "综合查询界面_ACZ1011C"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   14400
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter sFrameUserCond 
      Height          =   8790
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   15505
      _Version        =   196609
      SplitterBarWidth=   3
      BorderStyle     =   0
      PaneTree        =   "ACZ1011C.frx":0000
      Begin VB.TextBox txt_TabName 
         Height          =   3510
         Left            =   11355
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   0
         Width           =   3825
      End
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   5220
         Left            =   6045
         TabIndex        =   8
         Top             =   3570
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   9208
         _Version        =   196609
         SplitterBarWidth=   3
         BackColor       =   -2147483638
         PaneTree        =   "ACZ1011C.frx":0092
         Begin VB.TextBox txt_User_Cond 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4215
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   975
            Width           =   7605
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   885
            Left            =   30
            TabIndex        =   9
            Top             =   30
            Width           =   9075
            _ExtentX        =   16007
            _ExtentY        =   1561
            _Version        =   196609
            BackColor       =   14737632
            Begin VB.ComboBox cbo_SQL_Name 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   150
               TabIndex        =   12
               Top             =   420
               Width           =   7515
            End
            Begin Threed.SSCommand cmd_Sql_Delete 
               Height          =   630
               Left            =   7800
               TabIndex        =   26
               Top             =   120
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   1111
               _Version        =   196609
               BackColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "删除 SQL"
            End
            Begin VB.Label lblTableID 
               BackColor       =   &H00E0E0E0&
               Caption         =   "数据表 ID: "
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   270
               Left            =   240
               TabIndex        =   11
               Top             =   120
               Width           =   1950
            End
            Begin VB.Label lblTableName 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Left            =   2460
               TabIndex        =   10
               Top             =   120
               Width           =   5040
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   4215
            Left            =   7695
            TabIndex        =   14
            Top             =   975
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   7435
            _Version        =   196609
            BackColor       =   14737632
            Begin Threed.SSFrame SSFrame3 
               Height          =   555
               Left            =   150
               TabIndex        =   15
               Top             =   1995
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   979
               _Version        =   196609
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Begin VB.Label lblCond 
                  Caption         =   "AND"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   0
                  Left            =   555
                  TabIndex        =   19
                  Top             =   75
                  Width           =   300
               End
               Begin VB.Label lblCond 
                  Caption         =   "OR"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   180
                  Index           =   1
                  Left            =   555
                  TabIndex        =   18
                  Top             =   270
                  Width           =   300
               End
               Begin VB.Label lblCond 
                  Caption         =   "NOT"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Index           =   2
                  Left            =   120
                  TabIndex        =   17
                  Top             =   270
                  Width           =   300
               End
               Begin VB.Label lblCond 
                  Caption         =   "="
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9.75
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   210
                  Index           =   3
                  Left            =   210
                  TabIndex        =   16
                  Top             =   60
                  Width           =   255
               End
            End
            Begin Threed.SSCommand ssCond 
               Height          =   360
               Index           =   0
               Left            =   150
               TabIndex        =   20
               Top             =   810
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   635
               _Version        =   196609
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "SELECT"
            End
            Begin Threed.SSCommand ssCond 
               Height          =   360
               Index           =   1
               Left            =   150
               TabIndex        =   21
               Top             =   1185
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   635
               _Version        =   196609
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "FROM"
            End
            Begin Threed.SSCommand ssCond 
               Height          =   360
               Index           =   2
               Left            =   150
               TabIndex        =   22
               Top             =   1560
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   635
               _Version        =   196609
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "WHERE"
            End
            Begin Threed.SSCommand cmdUserOk 
               Height          =   630
               Left            =   150
               TabIndex        =   23
               Top             =   150
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   1111
               _Version        =   196609
               ForeColor       =   0
               BackColor       =   65280
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "查询"
            End
            Begin Threed.SSCommand cmdUserCancel 
               Height          =   360
               Left            =   150
               TabIndex        =   24
               Top             =   3660
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   635
               _Version        =   196609
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "退出"
            End
            Begin Threed.SSCommand ssCond 
               Height          =   360
               Index           =   3
               Left            =   150
               TabIndex        =   25
               Top             =   2610
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   635
               _Version        =   196609
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "CLEAR"
            End
            Begin Threed.SSCommand cmd_Sql_Insert 
               Height          =   630
               Left            =   150
               TabIndex        =   27
               Top             =   3000
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   1111
               _Version        =   196609
               BackColor       =   65280
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "保存 SQL"
            End
         End
      End
      Begin FPSpread.vaSpread ss3 
         Height          =   8790
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   5985
         _Version        =   393216
         _ExtentX        =   10557
         _ExtentY        =   15505
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
         MaxCols         =   2
         MaxRows         =   20
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "ACZ1011C.frx":0104
      End
      Begin FPSpread.vaSpread ss4 
         Height          =   3510
         Left            =   6045
         TabIndex        =   6
         Top             =   0
         Width           =   5250
         _Version        =   393216
         _ExtentX        =   9260
         _ExtentY        =   6191
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
         MaxCols         =   4
         MaxRows         =   30
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "ACZ1011C.frx":05CA
      End
   End
   Begin Threed.SSCommand cmd_User_Select 
      Height          =   405
      Left            =   11850
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   714
      _Version        =   196609
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "选择条件直接输入"
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   210
      Top             =   90
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   556
      Caption         =   "选择件数"
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
   Begin CSTextLibCtl.sidbEdit txt_Sel_Count 
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   90
      Width           =   1155
      _Version        =   262145
      _ExtentX        =   2037
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   0
      NumIntDigits    =   5
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8640
      Left            =   60
      TabIndex        =   4
      Top             =   570
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   15240
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
      MaxCols         =   50
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACZ1011C.frx":0B88
   End
   Begin VB.Label lbl_count 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3990
      TabIndex        =   2
      Top             =   165
      Visible         =   0   'False
      Width           =   3915
   End
End
Attribute VB_Name = "ACZ1011C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Production Management System
'-- Sub_System Name
'-- Program Name      Table Data Selection
'-- Program ID        ACZ1010C
'-- Designer          KIM SOO HEON
'-- Coder             KIM SOO HEON
'-- Date              2005.12.02
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

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim SQL             As String
Dim iSelectCnt      As Integer
Dim iSelectUserCnt  As Integer
Dim CondBetween     As Long
Dim sUser_ID        As String

Const STAND_ID = "0000000"


Private Sub Form_Define()
    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")

    Mc1.Add Item:=pControl1, Key:="pControl"
    Mc1.Add Item:=nControl1, Key:="nControl"
    Mc1.Add Item:=mControl1, Key:="mControl"
    Mc1.Add Item:=iControl1, Key:="iControl"
    Mc1.Add Item:=rControl1, Key:="rControl"
    Mc1.Add Item:=aControl1, Key:="aControl"
    Mc1.Add Item:=lControl1, Key:="lControl"

    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    For i = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, i, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next i
        
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    Proc_Sc.Add Item:=sc1, Key:="Sc"

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0

End Sub

Private Sub cmd_Sql_Search_Click()
    Screen.MousePointer = vbHourglass
    Call Form_Ref
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'    If KeyAscii = KEY_RETURN Then
'        KeyAscii = 0
'        SendKeys "{TAB}"
'    End If

End Sub

Private Sub Form_Load()
    Dim iDR    As Long
    
    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(ss1)
'    Call Gp_Sp_Setting(ss2)
    Call Gp_Sp_Setting(ss3)
    Call Gp_Sp_Setting(ss4)
    
    Call Gf_Sp_Cls(sc1)
        
    Call Gp_Sp_ColGet(ss1, "C-System.INI", Me.Name)
'    Call Gp_Sp_ColGet(ss2, "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss3, "C-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss4, "C-System.INI", Me.Name)
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Screen.MousePointer = vbDefault
    
    iSelectCnt = 0
    iSelectUserCnt = 0
    CondBetween = 7
    
    Call SerchTable
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(ss1, "C-System.INI", Me.Name)
'    Call Gp_Sp_ColSet(ss2, "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss3, "C-System.INI", Me.Name)
    Call Gp_Sp_ColSet(ss4, "C-System.INI", Me.Name)

    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
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
    Dim iDR As Integer
    
    If Gf_Sp_Cls(sc1) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        Call cmd_all_clear_Click
        Call ssCond_Click(3)
        txt_User_Cond.Text = ""
        lbl_count.Caption = ""
        cbo_SQL_Name.Text = ""
        ss1.MaxCols = 0
        sFrameUserCond.Visible = False
    End If

End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ref()

    Dim iDR     As Long
    Dim iLoc    As Long
    Dim sText   As String
    
    iLoc = 0

    On Error Resume Next

    sText = Replace(Trim(txt_User_Cond.Text), vbCrLf, "")

    If sText = "" Then Exit Sub
    
    If Mid(sAuthority, 1, 1) <> "1" Then
       Call Gp_MsgBoxDisplay("您没有查询权限", "", "错误提示")
       Exit Sub
    End If
    
    iLoc = InStr(1, UCase(sText), "WHERE")
    If iLoc = 0 Then
       Call Gp_MsgBoxDisplay("查询语句中必须有条件限制 WHERE", "", "错误提示")
       Exit Sub
    End If

    iLoc = InStr(1, UCase(sText), "FROM")
    If iLoc = 0 Then
       Call Gp_MsgBoxDisplay("查询语句中必须有 FROM", "", "错误提示")
       Exit Sub
    End If
        
    Call txt_User_Cond_Change
    Call cmd_Select
    
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

Private Sub cmd_User_Select_Click()

    If sFrameUserCond.Visible = True Then
        sFrameUserCond.Visible = False
        cmd_User_Select.ForeColor = &H808080
        Exit Sub
    End If
    
    sFrameUserCond.Top = 570
    sFrameUserCond.Left = 150
    sFrameUserCond.Width = 15180
    sFrameUserCond.Height = 8790
    sFrameUserCond.Visible = True
    sFrameUserCond.ZOrder (0)
    
    cmd_User_Select.ForeColor = &HFF0000
    
    If Trim(lblTableID.Caption) = "" Then iSelectUserCnt = 0
End Sub


'Private Sub Direct_Select()
'
'    Dim SQL             As String
'    Dim iDR             As Long
'    Dim iDc             As Long
'
'    On Error GoTo Error_Rtn
'
'    Set AdoRs = New ADODB.Recordset
'
'    sFrameUserCond.Visible = False
'
'    Screen.MousePointer = vbHourglass
'    ss1.MaxRows = 0
'
'    SQL = Trim(txt_User_Cond.Text)
'    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
'
'    ss1.MaxCols = AdoRs.Fields.Count
'    iDc = AdoRs.Fields.Count - 1
'
'    With ss1
'        .Row = 0
'        For iDR = 0 To iDc
'            .Col = iDR + 1
'            .Text = AdoRs.Fields(iDR).Name
''            .Text = Mid(AdoRs.Fields(iDR).Name, 1, Len(Trim(AdoRs.Fields(iDR).Name)) - 1)
'            .ColWidth(iDR + 1) = Len(AdoRs.Fields(iDR).Name) + 1
'            .TypeHAlign = SS_CELL_H_ALIGN_CENTER
'        Next iDR
'    End With
'
'    If AdoRs.EOF = True Then
'        MDIMain.StatusBar1.Panels(1) = "提示信息: 没有资料"
'    Else
'        Call Fill_Spread(AdoRs, ss1)
'        MDIMain.StatusBar1.Panels(1) = "提示信息: 资料已被查询"
'    End If
'
'    If ss1.MaxRows > 0 Then
'        lbl_count.Caption = Format(ss1.MaxRows, "#,##0") & "(件) "
'    End If
'
'    Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows, True)
'    AdoRs.Close
'
'    ss1.OperationMode = OperationModeNormal
'    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'    Screen.MousePointer = vbDefault
'
'    Exit Sub
'
'Error_Rtn:
'
'    Call Gp_MsgBoxDisplay("查询错误 : " & Error)
'
'    Screen.MousePointer = vbDefault
'
'End Sub
'
Private Sub cmd_Select()
    
    Dim SQL             As String
    Dim iDR             As Long
    Dim iDc             As Long
    Dim iLoc            As Long
    Dim iFrom           As Long
    Dim iTo             As Long
    Dim sText           As String
    Dim sItem_Name      As String
    Dim sSumItemName    As String

    On Error GoTo Error_Rtn
    
    Set AdoRs = New ADODB.Recordset
    
'    sFrameCond.Visible = False
    sFrameUserCond.Visible = False
    
    Screen.MousePointer = vbHourglass
        
    If txt_Sel_Count.Value = 0 Then txt_Sel_Count.Value = 2000
           
    sText = Replace(Trim(txt_User_Cond.Text), vbCrLf, "")
    
    SQL = "SELECT * FROM ( " & sText & ") " & " WHERE ROWNUM <= " & txt_Sel_Count.Value

    ss1.MaxCols = 0
    ss1.MaxRows = 0
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    ss1.MaxCols = AdoRs.Fields.Count
    iDc = AdoRs.Fields.Count - 1
    
    With ss1
        .Row = 0
        For iDR = 0 To iDc
            .Col = iDR + 1
'                .Text = AdoRs.Fields(iDr).Name
            .Text = Mid(AdoRs.Fields(iDR).Name, 1, Len(Trim(AdoRs.Fields(iDR).Name)))
            .ColWidth(iDR + 1) = Len(AdoRs.Fields(iDR).Name) + 1
            .TypeHAlign = SS_CELL_H_ALIGN_CENTER
        Next iDR
    End With
    
    If AdoRs.EOF = True Then
        MDIMain.StatusBar1.Panels(1) = "提示信息: 没有资料"
    Else
        Call Fill_Spread(AdoRs, ss1)
        MDIMain.StatusBar1.Panels(1) = "提示信息: 资料已被查询"
    End If
    
    iLoc = InStr(1, lbl_count.Caption, "中")
    If ss1.MaxRows = 0 Then
        lbl_count.Caption = ""
    ElseIf iLoc > 0 Then
        lbl_count.Caption = Left(lbl_count.Caption, iLoc - 1) & "中 " & Format(ss1.MaxRows, "#,##0") & "(件) "
    Else
        lbl_count.Caption = lbl_count.Caption & "中 " & Format(ss1.MaxRows, "#,##0") & "(件) "
    End If
    
    Call Gp_Sp_BlockLock(ss1, 1, ss1.MaxCols, 1, ss1.MaxRows, True)
    AdoRs.Close
    
    ss1.OperationMode = OperationModeNormal
    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
Error_Rtn:
    
    Call Gp_MsgBoxDisplay("查询错误 : " & Error)
    
    Screen.MousePointer = vbDefault
    
End Sub

'Private Sub cmd_SQL_Click()
'    Dim sItem_Name  As String
'
'    chk_Direct_Flag.Value = ssCBUnchecked
'
'    If Trim(txt_Table_Name.Text) <> "" Then
'
'        Call ItemSelect(sItem_Name)
'
'        Call Comma_Clear
'
'        SQL = ""
'        SQL = SQL & " SELECT  " & vbCrLf & sItem_Name
'        SQL = SQL & "   FROM  " & txt_Table_Name.Text & vbCrLf
'        If Trim(txt_Where_Cond.Text) <> "" Then
'            SQL = SQL & "  WHERE  " & Trim(txt_Where_Cond.Text)
'        End If
'
'        txt_User_Cond.Text = SQL
'    End If
'
'    sFrame.Visible = True
'    sFrame.ZOrder (0)
'End Sub

Private Sub cmd_all_clear_Click()

    Dim iDR             As Integer
    
    iSelectCnt = 0

End Sub

'Private Sub cmdCancel_Click()
'    sFrameCond.Visible = False
'End Sub

'Private Sub cmdOk_Click()
'    Dim sItem_Name  As String
'
'    chk_Direct_Flag.Value = ssCBUnchecked
'
'    Call ItemSelect(sItem_Name)
'
'    If Trim(sItem_Name) = "" Then
'        txt_User_Cond.Text = ""
''        sFrameCond.Visible = False
'        Exit Sub
'    End If
'
'    If Trim(txt_Where_Cond.Text) = "" Then
'        Call Gp_MsgBoxDisplay("输入条件项目", "", "错误提示")
'        Exit Sub
'    End If
'
'    Call Comma_Clear
'
'    SQL = ""
'    SQL = SQL & " SELECT  " & vbCrLf & sItem_Name
'    SQL = SQL & "   FROM  " & txt_Table_Name.Text & vbCrLf
'    SQL = SQL & "  WHERE  " & Trim(txt_Where_Cond.Text)
'
'    txt_User_Cond.Text = SQL
''    sFrameCond.Visible = False
'
'    Call Form_Ref
'End Sub

'Private Sub cmd_ADD_Click()
'    Call CondEdit(" AND ")
'End Sub

'Private Sub cmd_AND_Click()
'    Call CondEdit(" AND ")
'End Sub

'Private Sub cmd_Clear_Click()
'    txt_Where_Cond.Text = ""
'End Sub

'Private Sub cmd_OR_Click()
'    Call CondEdit(" OR ")
'End Sub

Private Sub cmdUserCancel_Click()
    sFrameUserCond.Visible = False
End Sub

Private Sub cmdUserOk_Click()

    sFrameUserCond.Visible = False
    
    Call Form_Ref
    
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txt_User_Cond_Change()
    Dim iDR     As Long
    Dim SQL     As String
    Dim sText   As String
    Dim iLoc    As Long
    Dim iDc     As Long

    lbl_count.Caption = ""
    iLoc = 0

    On Error Resume Next

    sText = Replace(Trim(txt_User_Cond.Text), vbCrLf, "")

    If sText = "" Then Exit Sub

    If UCase(Left(sText, 1)) <> "S" Then
       Call Gp_MsgBoxDisplay("条件必须输入", "", "错误提示")
       Exit Sub
    End If
    
    iLoc = InStr(1, UCase(sText), "UPDATE")
    If iLoc > 0 Then
       Call Gp_MsgBoxDisplay("不允许进行更新数据操作", "", "错误提示")
       Exit Sub
    End If
    
    iLoc = InStr(1, UCase(sText), "INSERT")
    If iLoc > 0 Then
       Call Gp_MsgBoxDisplay("不允许进行插入数据操作", "", "错误提示")
       Exit Sub
    End If
    
    iLoc = InStr(1, UCase(sText), "DELETE")
    If iLoc > 0 Then
       Call Gp_MsgBoxDisplay("不允许进行删除数据操作", "", "错误提示")
       Exit Sub
    End If
    
    iLoc = InStr(1, UCase(sText), "WHERE")
    If iLoc = 0 Then Exit Sub

    iLoc = InStr(1, UCase(sText), "FROM")
    If iLoc = 0 Then Exit Sub

    If txt_Sel_Count.Value = 0 Then txt_Sel_Count.Value = 2000

End Sub

Private Sub cmd_Sql_Insert_Click()
    Dim SQL             As String
    Dim sName           As String
    Dim sText           As String
    Dim sInsName        As String
    

    On Error GoTo Error_cmd_Sql_Insert
        
    If Mid(sAuthority, 2, 1) <> "1" Then
       Call Gp_MsgBoxDisplay("您没有保存权限", "", "错误提示")
       Exit Sub
    End If

    sText = Replace(Trim(txt_User_Cond.Text), "'", "@")
    
    If sText = "" Then Exit Sub
    
    If Trim(cbo_SQL_Name.Text) = "" Then
        Call Gp_MsgBoxDisplay("请先输入查询 SQL 名称")
        Exit Sub
    End If
    
    sName = Trim(cbo_SQL_Name.Text)
        
    Set AdoRs = New ADODB.Recordset

    sInsName = STAND_ID
        
    SQL = " SELECT      USER_SQL     " & vbCrLf
    SQL = SQL & "  FROM ZP_USER_SQL  " & vbCrLf
    SQL = SQL & " WHERE EMP_ID   =  '" & sInsName & "'" & vbCrLf
    SQL = SQL & "   AND SQL_NAME =  '" & sName & "'" & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    If Left(sUserID, 6) <> "1JS600" And AdoRs.RecordCount > 0 Then
        Call Gp_MsgBoxDisplay("已经有相同的名称")
        Exit Sub
    End If
    
    M_CN1.BeginTrans
    
    If AdoRs.RecordCount = 0 Then
        SQL = ""
        SQL = " Insert Into ZP_USER_SQL ( EMP_ID, SQL_NAME,  USER_SQL, INS_EMP_ID  )          " & vbCrLf
        SQL = SQL & "  Values ( '" & sInsName & "','" & sName & "','" & sText & "','" & sUserID & "')"
    Else
        SQL = " UPDATE  ZP_USER_SQL  SET                " & vbCrLf
        SQL = SQL & "  USER_SQL      =  '" & sText & "'" & vbCrLf
        SQL = SQL & " WHERE EMP_ID   =  '" & sInsName & "'" & vbCrLf
        SQL = SQL & "   AND SQL_NAME =  '" & sName & "'" & vbCrLf
    End If
    
    M_CN1.Execute SQL
    
    M_CN1.CommitTrans
    
    AdoRs.Close
    
    Call SerchUserSql
    cbo_SQL_Name.Text = sName
    
    Exit Sub
    
Error_cmd_Sql_Insert:
    
    Call Gp_MsgBoxDisplay(" 保存错误 : " & Error)
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmd_Sql_Delete_Click()
    
    Dim SQL             As String
    Dim sName           As String
    
    On Error GoTo Error_cmd_Sql_Delete
        
    If Left(sUserID, 6) <> "1JS600" Then
        If Mid(sAuthority, 4, 1) <> "1" Then
           Call Gp_MsgBoxDisplay("您没有删除权限", "", "错误提示")
           Exit Sub
        End If
    End If
    
    If Trim(cbo_SQL_Name.Text) = "" Then Exit Sub
    
    txt_User_Cond.Text = ""
    
    sName = Trim(cbo_SQL_Name.Text)
    
    M_CN1.BeginTrans
    
    SQL = ""
    SQL = " DELETE FROM  ZP_USER_SQL " & vbCrLf
    SQL = SQL & " WHERE INS_EMP_ID   IN ('1JS1005','" & sUserID & "')" & vbCrLf
    SQL = SQL & "   AND SQL_NAME = '" & sName & "'" & vbCrLf
    
    M_CN1.Execute SQL
    
    M_CN1.CommitTrans
    
    Call SerchUserSql
    
    Exit Sub
    
Error_cmd_Sql_Delete:
    
    Call Gp_MsgBoxDisplay(" 删除错误 : " & Error)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub cbo_SQL_Name_Click()
    Dim sSql            As String
    Dim SQL             As String
    Dim sName           As String
    Dim iDR             As Long
    Dim iDc             As Long
    Dim iLoc            As Long
    Dim iLoc2           As Long
    Dim iLoc3           As Long
    Dim iDx             As Integer
    Dim iSeq            As Integer
    Dim sTable_Name     As String
    Dim sColum_Name     As String
    Dim sAliasName      As String
    Dim sMessg          As String

    On Error Resume Next
    
    If Trim(cbo_SQL_Name.Text) = "" Then Exit Sub
    
'    If Trim(txt_User_Cond.Text) <> "" Then
'        sMessg = "表格中还有数据未处理，" + vbCrLf
'        sMessg = sMessg + "放弃并继续吗？"
'
'        If Not Gf_MessConfirm(sMessg, "Q") Then
'            Exit Sub
'        End If
'    End If
               
    sName = Trim(cbo_SQL_Name.Text)
    iSeq = cbo_SQL_Name.ListIndex
    
    Set AdoRs = New ADODB.Recordset
    
    SQL = " SELECT      USER_SQL       " & vbCrLf
    SQL = SQL & "  FROM ZP_USER_SQL    " & vbCrLf
    SQL = SQL & " WHERE EMP_ID   =  '" & STAND_ID & "'" & vbCrLf
    SQL = SQL & "   AND SQL_NAME =  '" & sName & "'" & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly

    If AdoRs.RecordCount = 0 Then txt_User_Cond.Text = "":    Exit Sub
        txt_User_Cond.Text = Replace(Trim(AdoRs.Fields(0) & ""), "@", "'")
    AdoRs.Close
    
    Screen.MousePointer = vbHourglass
    
'    Call cmd_all_clear_Click
    
'    sFrameCond.Visible = False
'    Call cmd_Auto_Select_Click
    
    sSql = Replace(UCase(txt_User_Cond.Text), vbCrLf, "")
    iLoc = InStr(1, sSql, "FROM")
    sSql = UCase(Trim(Mid(sSql, iLoc + 4, Len(sSql))))
          
    Do Until sSql = ""
        iLoc = InStr(1, sSql, "WHERE")
        iLoc2 = InStr(1, sSql, ",")
        
        If iLoc2 = 0 Or iLoc2 > iLoc Then
            iLoc = InStr(1, sSql, " ")
            sTable_Name = Trim(Left(sSql, iLoc))
            sSql = ""
        Else
            sTable_Name = Trim(Mid(sSql, 1, iLoc2 - 2))
            sSql = UCase(Trim(Mid(sSql, iLoc2 + 1, Len(sSql))))
        End If
        
    Loop
    
    sSql = Replace(UCase(txt_User_Cond.Text), vbCrLf, "")
    iLoc = InStr(1, sSql, "SELECT")
    sSql = UCase(Trim(Mid(sSql, iLoc + 6, Len(sSql))))
          
    Do Until sSql = ""
        iLoc = InStr(1, sSql, " ")
        iLoc2 = InStr(1, sSql, ".")
        iLoc3 = InStr(1, sSql, "(")
        
        If iLoc > iLoc3 And iLoc3 > 0 Then
            iLoc2 = InStr(iLoc3, sSql, ".")
            sAliasName = Left(sSql, iLoc2 - 1)
            iLoc3 = InStr(iLoc3, sSql, ")")
            sColum_Name = Trim(Mid(sSql, iLoc2 + 1, iLoc3 - iLoc2 - 1))
        Else
            sAliasName = Left(sSql, iLoc2 - 1)
            sColum_Name = Trim(Mid(sSql, iLoc2 + 1, iLoc - iLoc2))
        End If
        
        Select Case sAliasName
            Case "A":  iDx = 1
            Case "B":  iDx = 2
            Case "C":  iDx = 3
            Case "D":  iDx = 4
        End Select
        
        iLoc = InStr(1, sSql, "FROM")
        iLoc2 = InStr(1, sSql, ",")
        
        If iLoc2 = 0 Or iLoc2 > iLoc Then
            sSql = ""
        Else
            sSql = UCase(Trim(Mid(sSql, iLoc2 + 1, Len(sSql))))
        End If
        
    Loop
    
    sSql = UCase(txt_User_Cond.Text)
    iLoc = InStr(1, sSql, "WHERE")
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub ss3_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sTable  As String
    Dim iLoc    As Integer
    
    If Row < 1 Then Exit Sub

    ss3.Col = 2
    ss3.Row = Row
    sTable = Trim(ss3.Text)
    
    ss3.Col = 1
'    lblTableName.Caption = sTable & "(" & Trim(ss3.Text) & ")"
    lblTableName.Caption = Trim(ss3.Text)
    
    If Trim(txt_TabName.Text) = "" Then iSelectUserCnt = 0
    
    iLoc = InStr(1, txt_TabName.Text, sTable & " ")
    If iLoc = 0 Then
        iSelectUserCnt = iSelectUserCnt + 1
        lblTableID.Caption = sTable & " " & Chr(iSelectUserCnt + 64)
    Else
        lblTableID.Caption = sTable & " " & Mid(txt_TabName.Text, iLoc + Len(sTable) + 1, 1)
    End If
            
    Call Column_Edit(sTable, ss4)

    Call Gp_Sp_BlockLock(ss4, 1, ss4.MaxCols, 1, ss4.MaxRows, True)
    
'    ss3.OperationMode = OperationModeNormal
'    ss4.OperationMode = OperationModeNormal
End Sub

Private Sub ss4_DblClick(ByVal Col As Long, ByVal Row As Long)
    Dim sColId   As String
    Dim sColName As String
    Dim iLoc     As Integer
    
    If Row < 1 Then Exit Sub
    ss4.Row = Row
    ss4.Col = 2:    sColName = Trim(ss4.Text)
    Call ItemNameEdit(sColName)
    
    ss4.Col = 3:    sColId = Trim(ss4.Text)
    
    If Trim(txt_User_Cond.Text) = "" Then
        txt_User_Cond.Text = "SELECT " & vbCrLf
    End If
    
    iLoc = InStr(1, UCase(txt_User_Cond.Text), "WHERE")
    
    txt_User_Cond.Text = txt_User_Cond.Text & " " & Right(Trim(lblTableID.Caption), 1) & "." & sColId
    
    If iLoc = 0 Then
        If sColName <> "" Then
            txt_User_Cond.Text = txt_User_Cond.Text & " " & sColName & Right(Trim(lblTableID.Caption), 1) & vbCrLf & ","
        Else
            txt_User_Cond.Text = txt_User_Cond.Text & " " & sColId & "_" & Right(Trim(lblTableID.Caption), 1) & vbCrLf & ","
        End If
    End If
    
    If Trim(txt_TabName.Text) = "" Then
        txt_TabName.Text = lblTableID.Caption
    Else
        iLoc = InStr(1, txt_TabName.Text, Mid(Trim(lblTableID.Caption), 1, Len(Trim(lblTableID.Caption)) - 1))
        If iLoc = 0 Then
            txt_TabName.Text = txt_TabName.Text & "," & vbCrLf & lblTableID.Caption
        End If
    End If
    
End Sub





Private Sub Column_Edit(Table_id As String, oSpr As vaSpread)
    Dim SQL             As String
    Dim ColName         As String
    Dim iDR             As Integer
    
    Set AdoRs = New ADODB.Recordset
    
    SQL = " SELECT      COMMENTS, COLNO, CNAME, COLTYPE,                            " & vbCrLf
    SQL = SQL & "       DECODE(COLTYPE,'NUMBER',PRECISION||','||SCALE,WIDTH) WIDTH  " & vbCrLf
    SQL = SQL & "  FROM COL  A, USER_COL_COMMENTS B " & vbCrLf
    SQL = SQL & " WHERE A.TNAME  =  '" & Table_id & "'" & vbCrLf
    SQL = SQL & "   AND A.TNAME  = B.TABLE_NAME     " & vbCrLf
    SQL = SQL & "   AND A.CNAME  = B.COLUMN_NAME    " & vbCrLf
    SQL = SQL & " ORDER BY COLNO " & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly

    oSpr.MaxRows = AdoRs.RecordCount

    iDR = 1
    Do Until AdoRs.EOF
        oSpr.Row = iDR
        oSpr.Col = 3
        oSpr.Text = AdoRs.Fields("CNAME") & ""
        ColName = AdoRs.Fields("CNAME")
        
        oSpr.Col = 2
        If Trim(AdoRs.Fields("COMMENTS") & "") = "" Then
            oSpr.Text = ColName
        Else
            oSpr.Text = AdoRs.Fields("COMMENTS") & ""
        End If
        
        oSpr.Col = 4
        oSpr.Text = AdoRs.Fields("COLTYPE")
'        oSpr.Col = 5
'        oSpr.Text = AdoRs.Fields("WIDTH")

        AdoRs.MoveNext
        iDR = iDR + 1
    Loop

    AdoRs.Close
        
End Sub

Private Sub ItemNameEdit(sItemEditName As String)
    
    If sItemEditName = "" Then Exit Sub
    
    sItemEditName = Replace(sItemEditName, " ", "_")
    sItemEditName = Replace(sItemEditName, "(", "")
    sItemEditName = Replace(sItemEditName, ")", "")
    sItemEditName = Replace(sItemEditName, ",", "")
    sItemEditName = Replace(sItemEditName, "/", "")
    sItemEditName = Replace(sItemEditName, "*", "")
    sItemEditName = Replace(sItemEditName, "+", "")
    sItemEditName = Replace(sItemEditName, "-", "")
    sItemEditName = Replace(sItemEditName, "'", "")
    sItemEditName = Replace(sItemEditName, ":", "")
    sItemEditName = Replace(sItemEditName, ";", "")
    sItemEditName = Replace(sItemEditName, vbCrLf, "")
    
End Sub

Private Sub SelectItemEdit(sItem_Name As String, sSumItemName As String)
    Dim sItem  As String
    Dim sTable As String
    Dim sCont  As String
    Dim iMax   As Integer
    Dim iCnt   As Integer
    Dim iPeri  As Integer
    Dim iComm  As Integer
    Dim iSpace As Integer
     
    sSumItemName = ""
    iMax = Len(sItem_Name)
    sItem = Trim(Mid(sItem_Name, 7, iMax))
        
    Do
        iPeri = InStr(1, sItem, ".")
        If iPeri = 0 Then sSumItemName = sItem:  Exit Do
        
        sItem = Mid(sItem, iPeri + 1, iMax)
            
        iSpace = InStr(1, sItem, " ")
        iComm = InStr(1, sItem, ",")
        If iComm <> iSpace And iComm <> 0 Then sCont = Trim(Mid(sItem, iSpace, iComm - iSpace))
        
        If iComm = iSpace Then
            sSumItemName = sSumItemName & sItem
            iMax = 0
        ElseIf iComm > iSpace And sCont <> "" Then
            sItem = Trim(Mid(sItem, iSpace, iMax))

            iComm = InStr(1, sItem, ",")
            sSumItemName = sSumItemName & Left(sItem, iComm)
        ElseIf iComm < iSpace And iComm = 0 Then
            sItem = Trim(Mid(sItem, iSpace, iMax))
            sSumItemName = sSumItemName & sItem
            iMax = 0
        Else
            sSumItemName = sSumItemName & Left(sItem, iComm)
        End If
    Loop Until iMax = 0
    
End Sub

'Recordset To Spread Fill Data
Public Sub Fill_Spread(ByRef objRS As Recordset, ByRef oSpr As vaSpread, Optional ByVal bColorCol As Integer = -1)
    Dim iFieldCnt       As Integer
    Dim lRecordCnt      As Long
    Dim sFirstString    As String
    Dim iChangeRow      As Integer
    Dim bColor          As Boolean
    Dim lDr             As Long
    Dim iDc             As Integer
    Dim iTextLength     As Integer
    Dim lStartNO        As Long
    
    iFieldCnt = objRS.Fields.Count
    lRecordCnt = objRS.RecordCount
    
    lStartNO = oSpr.MaxRows
    
    oSpr.MaxRows = oSpr.MaxRows + lRecordCnt
    
    For lDr = lStartNO + 1 To lRecordCnt + lStartNO
        For iDc = 1 To iFieldCnt
            oSpr.Row = lDr:   oSpr.Col = iDc
            If oSpr.CellType = SS_CELL_TYPE_BUTTON Then
                oSpr.TypeButtonText = objRS(iDc - 1).Value & ""
            Else
                If objRS(iDc - 1).Type = adNumeric Then
                    If InStr(1, objRS(iDc - 1), ".") Then
                        oSpr.Text = Format(objRS(iDc - 1).Value, "#,##0.0##")
                    Else
                        oSpr.Text = Format(objRS(iDc - 1).Value)
                    End If
                Else
                    If objRS(iDc - 1).Type = adDate Then
                        oSpr.Text = Format(objRS(iDc - 1).Value & "", "####-##-## ##:##:##")
                    Else
                        oSpr.Text = objRS(iDc - 1).Value & ""
                    End If
                End If
            End If
            
            iTextLength = Val(Len(oSpr.Text) & "")
            
            oSpr.Row = 0
            If iTextLength > oSpr.ColWidth(iDc) Then
                oSpr.ColWidth(iDc) = iTextLength
            End If
        Next iDc
        
        objRS.MoveNext
    Next lDr

End Sub

Private Sub SerchTable()
    Dim iDR   As Long
    
    Set AdoRs = New ADODB.Recordset

    SQL = " SELECT      TABLE_NAME,                     " & vbCrLf
    SQL = SQL & "       DECODE(SUBSTR(COMMENTS,1,6),'(OPEN)',SUBSTR(COMMENTS,7,LENGTH(COMMENTS)),COMMENTS) COMMENTS" & vbCrLf
    SQL = SQL & "  FROM USER_TAB_COMMENTS               " & vbCrLf
    SQL = SQL & " WHERE COMMENTS   IS NOT NULL          " & vbCrLf
    SQL = SQL & " AND SUBSTR(COMMENTS,1,6) = '(OPEN)' " & vbCrLf
    SQL = SQL & " ORDER BY TABLE_NAME           " & vbCrLf
    
    AdoRs.Open SQL, M_CN1, adOpenForwardOnly, adLockReadOnly
    
    ss3.MaxRows = AdoRs.RecordCount
    
    iDR = 1
    Do Until AdoRs.EOF
        
        ss3.Row = iDR
        ss3.Col = 1
        ss3.Text = AdoRs.Fields("COMMENTS")
        ss3.Col = 2
        ss3.Text = AdoRs.Fields("TABLE_NAME")
            
        AdoRs.MoveNext
        iDR = iDR + 1
    Loop
    
    AdoRs.Close
                
    Call SerchUserSql
    
    lblTableName.Caption = ""
    Call Gp_Sp_BlockLock(ss3, 1, ss3.MaxCols, 1, ss3.MaxRows, True)
    
End Sub

Private Sub SerchUserSql()
    Dim iDR   As Long
             
    cbo_SQL_Name.Clear
    
    SQL = " SELECT      SQL_NAME                     " & vbCrLf
    SQL = SQL & "  FROM ZP_USER_SQL                  " & vbCrLf
    SQL = SQL & " WHERE EMP_ID  =  '" & STAND_ID & "'" & vbCrLf
    
    Call Gf_ComboAdd(M_CN1, cbo_SQL_Name, SQL)
        
End Sub

'Private Sub CondEdit(sAndOr As String)
'    Dim sColumnName As String
'    Dim iLoc        As Integer
'    Dim sCond       As String
'
''    If Right(sCond, 1) = "=" Or Right(sCond, 3) = "AND" Then
''       Call Gp_MsgBoxDisplay("错误条件项目(没输入条件值)..", "", "错误提示")
''       Exit Sub
''    End If
'
'    If Trim(CBO_FILTER.Text) = "" Then Exit Sub
'
'    iLoc = InStr(InStr(1, CBO_ITEM.Text, ".") + 1, CBO_ITEM.Text, ".")
'
'    sColumnName = Trim(Left(CBO_ITEM.Text, iLoc - 1))
'
'    sCond = Replace(Trim(txt_Where_Cond.Text), vbCrLf, "")
'
'    If Right(sCond, 1) = "=" Then
'       txt_Where_Cond.Text = Left(txt_Where_Cond.Text, Len(txt_Where_Cond.Text) - 2)
'       sColumnName = ""
'    ElseIf UCase(Right(sCond, 2)) = "OR" Or UCase(Right(sCond, 3)) = "AND" Then
'       txt_Where_Cond.Text = Left(txt_Where_Cond.Text, Len(txt_Where_Cond.Text) - 4)
'       txt_Where_Cond.Text = txt_Where_Cond.Text & vbCrLf & sAndOr
'    ElseIf Trim(txt_Where_Cond.Text) <> "" Then
'       txt_Where_Cond.Text = txt_Where_Cond.Text & vbCrLf & sAndOr
'    End If
'
'    If CBO_FILTER.ListIndex = 6 Then
'        txtCond1.Text = txtCond1.Text & "%"
'    End If
'
'    Select Case UCase(Trim(Right(CBO_ITEM.Text, 50)))
'        Case "NUMBER", "LONG"
'            txt_Where_Cond.Text = txt_Where_Cond.Text & sColumnName & " " & CBO_FILTER.Text & " " & txtCond1.Text
'        Case "DATE", "TIMESTAMP(6)"
'            txt_Where_Cond.Text = txt_Where_Cond.Text & sColumnName & " '" & CBO_FILTER.Text & " TO_DATE('" & txtCond1.Text & "','YYYY-MM-DD HH24:MI:SS')"
'        Case Else
'            txt_Where_Cond.Text = txt_Where_Cond.Text & sColumnName & " " & CBO_FILTER.Text & " '" & txtCond1.Text & "'"
'    End Select
'
'    If CBO_FILTER.ListIndex = CondBetween Then
'        Select Case UCase(Trim(Right(CBO_ITEM.Text, 50)))
'            Case "NUMBER", "LONG"
'                txt_Where_Cond.Text = txt_Where_Cond.Text & " AND " & txtCond2.Text
'            Case "DATE", "TIMESTAMP(6)"
'                txt_Where_Cond.Text = txt_Where_Cond.Text & " AND TO_DATE('" & txtCond2.Text & "','YYYY-MM-DD HH24:MI:SS')"
'            Case Else
'                txt_Where_Cond.Text = txt_Where_Cond.Text & " AND '" & txtCond2.Text & "'"
'        End Select
'    End If
'
'    CBO_ITEM.ListIndex = -1: CBO_FILTER.ListIndex = -1: txtCond1.Text = "": txtCond2.Text = ""
'End Sub
'
'Private Sub CBO_FILTER_Change()
'    Call txtCondEdit
'End Sub

'Private Sub CBO_FILTER_Click()
'    Call txtCondEdit
'End Sub

'Private Sub txtCondEdit()
'
'    If CBO_FILTER.ListIndex > CondBetween Then
'        txtCond1.Enabled = False
'    End If
'    If CBO_FILTER.ListIndex = CondBetween Then
'        txtCond2.Enabled = True
'    End If
'End Sub

'Private Sub CBO_ITEM_Click()
'    CBO_ITEM.RightToLeft = True
'    CBO_FILTER.SetFocus
'End Sub
'
'Private Sub CBO_ITEM_DropDown()
'    CBO_ITEM.RightToLeft = True
'End Sub

Private Sub lblCond_Click(Index As Integer)
    
    If lblCond(Index).Caption = "=" Then
        txt_User_Cond.Text = txt_User_Cond.Text & " " & lblCond(Index).Caption & " "
    Else
        txt_User_Cond.Text = txt_User_Cond.Text & vbCrLf & " " & lblCond(Index).Caption & " "
    End If
    
    txt_User_Cond.SetFocus
    txt_User_Cond.SelStart = Len(txt_User_Cond.Text) + 1
End Sub

Private Sub ssCond_Click(Index As Integer)

    txt_User_Cond.Text = Trim(txt_User_Cond.Text)
    
    If ssCond(Index).Caption = "SELECT" Then
        txt_User_Cond.Text = ssCond(Index).Caption & " " & txt_User_Cond.Text
    ElseIf ssCond(Index).Caption = "FROM" Then
        If Mid(txt_User_Cond.Text, Len(txt_User_Cond.Text), 1) = "," Then
           txt_User_Cond.Text = Mid(txt_User_Cond.Text, 1, Len(txt_User_Cond.Text) - 1)
        End If
        txt_User_Cond.Text = txt_User_Cond.Text & " FROM " & txt_TabName.Text & vbCrLf
    ElseIf ssCond(Index).Caption = "CLEAR" Then
        txt_User_Cond.Text = ""
        txt_TabName.Text = ""
        ss4.MaxRows = 0
        iSelectUserCnt = 0
    Else
        txt_User_Cond.Text = txt_User_Cond.Text & " " & ssCond(Index).Caption
    End If
    
End Sub


Private Sub txt_User_Cond1_Change()

End Sub


