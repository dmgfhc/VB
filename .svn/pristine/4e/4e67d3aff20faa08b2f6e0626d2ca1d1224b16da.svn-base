VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AED4010C 
   Caption         =   "确定轧钢作业生产管制指示_AED4010C"
   ClientHeight    =   9615
   ClientLeft      =   195
   ClientTop       =   2730
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   14715
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_to 
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
      Left            =   8655
      TabIndex        =   14
      Top             =   540
      Width           =   1365
   End
   Begin VB.TextBox txt_target 
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
      Left            =   11430
      TabIndex        =   13
      Top             =   540
      Width           =   1365
   End
   Begin VB.TextBox txt_from 
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
      Left            =   6900
      TabIndex        =   12
      Top             =   540
      Width           =   1365
   End
   Begin VB.TextBox TXT_PLT 
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
      Height          =   315
      Left            =   1500
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   540
      Width           =   540
   End
   Begin VB.TextBox TXT_PLT_NAME 
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
      Height          =   315
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   540
      Width           =   3420
   End
   Begin Threed.SSPanel SSPsend 
      Height          =   315
      Left            =   13005
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "已下达"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPpdt 
      Height          =   315
      Left            =   14115
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "生产中"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin CSTextLibCtl.sidbEdit SDB_SLAB_EDT_SEQ 
      Height          =   315
      Left            =   3090
      TabIndex        =   4
      Tag             =   "炉次编制号"
      Top             =   540
      Visible         =   0   'False
      Width           =   375
      _Version        =   262145
      _ExtentX        =   661
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   16
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
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_PRC_LINE 
      Height          =   315
      Left            =   3510
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   180
      _Version        =   262145
      _ExtentX        =   317
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      RawData         =   "1"
      Text            =   " 0"
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   16
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
      Undo            =   0
      Data            =   1
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   105
      Top             =   540
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
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
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   5505
      Top             =   540
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "起始板坯号"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   10035
      Top             =   540
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "目标板坯号"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   8265
      Top             =   540
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Caption         =   "->"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
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
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   420
      Left            =   105
      TabIndex        =   7
      Top             =   60
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   741
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_move 
         Height          =   330
         Left            =   525
         TabIndex        =   8
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "调 整"
      End
      Begin Threed.SSOption opt_delete 
         Height          =   330
         Left            =   2040
         TabIndex        =   9
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   582
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
         Caption         =   "删 除"
      End
      Begin Threed.SSOption opt_sent 
         Height          =   330
         Left            =   75
         TabIndex        =   10
         Top             =   -165
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
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
         Caption         =   "发 送"
      End
      Begin Threed.SSOption opt_cancel 
         Height          =   330
         Left            =   1050
         TabIndex        =   11
         Top             =   -165
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
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
         Caption         =   "取 消"
      End
      Begin Threed.SSOption opt_cnf 
         Height          =   330
         Left            =   3375
         TabIndex        =   20
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
         Caption         =   "生产管制指示"
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   420
      Left            =   5490
      TabIndex        =   15
      Top             =   60
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   741
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_target 
         Height          =   330
         Left            =   5925
         TabIndex        =   16
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "目标板坯号"
      End
      Begin Threed.SSOption opt_from 
         Height          =   330
         Left            =   1410
         TabIndex        =   17
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
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
         Caption         =   "起始板坯号"
      End
      Begin Threed.SSOption opt_to 
         Height          =   330
         Left            =   3150
         TabIndex        =   18
         Top             =   60
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   582
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
         Caption         =   "->"
      End
   End
   Begin VB.TextBox TXT_MPLATE_NO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10155
      MaxLength       =   12
      TabIndex        =   6
      Tag             =   "炉次管理号"
      Top             =   75
      Visible         =   0   'False
      Width           =   1395
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8370
      Left            =   90
      TabIndex        =   19
      Top             =   900
      Width           =   15120
      _Version        =   393216
      _ExtentX        =   26670
      _ExtentY        =   14764
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
      SpreadDesigner  =   "AED4010C.frx":0000
   End
End
Attribute VB_Name = "AED4010C"
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
'-- Program Name      指示调整
'-- Program ID        AGG2040C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Yang meng
'-- Coder             Yang meng
'-- Date              2003.7.23
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
Dim Mode As String

'Public Complete As Boolean           'Move Status Setting

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

Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sSlab_Edt_Seq_Fr As String
Dim sSlab_Edt_Seq_To As String
Dim sSlab_Edt_Seq_Tg As String

Private Sub Form_Define()
        
    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    Call Gp_Ms_Collection(TXT_PLT, "p", "n", "m", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    For i = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, i, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next i
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AED4010C.P_REFER1", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 5, True)
    Call Gp_Sp_ColHidden(ss1, 22, True)
    'Call Gp_Sp_ColHidden(ss1, 23, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
    End With

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
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"), False)
    
    Call Gp_Sp_ReadOnlySet(Sc1.Item("Spread"))
   
    Call Gf_Sp_Cls(Sc1)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "E-System.INI", Me.Name)
    
    TXT_PLT.Text = "C1"
    
    Call txt_plt_KeyUp(0, 0)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "E-System.INI", Me.Name)
    
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
    Set Sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Sc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        MDIMain.MenuTool.Buttons(4).Enabled = True
        TXT_PLT.Text = "C1"
        Call txt_plt_KeyUp(0, 0)
        opt_cnf.VALUE = False
        opt_sent.VALUE = False
        opt_cancel.VALUE = False
        opt_move.VALUE = False
        opt_delete.VALUE = False
        opt_from.VALUE = False
        opt_to.VALUE = False
        opt_target.VALUE = False
        opt_cnf.ForeColor = &H808080
        opt_sent.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_from.ForeColor = &H808080
        opt_to.ForeColor = &H808080
        opt_target.ForeColor = &H808080
        txt_from = ""
        txt_to = ""
        txt_target = ""
        TXT_MPLATE_NO = ""
        sSlab_Edt_Seq_Fr = 0
        sSlab_Edt_Seq_To = 0
        sSlab_Edt_Seq_Tg = 0
    End If
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = False                'Excel
    End With
        
    
End Sub

Public Sub Form_Ref()

    Dim sTemp As String
    Dim sL2_Send As String
    Dim sSlab_No As String
    Dim sPrc_Sts As String
    Dim iRow As Integer
    Dim iCol As Integer

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
       
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
       
        sSlab_Edt_Seq_Fr = 0
        sSlab_Edt_Seq_To = 0
        sSlab_Edt_Seq_Tg = 0
    
        With MDIMain.MenuTool
            .Buttons(7).Enabled = False                 'Row Insert
            .Buttons(8).Enabled = False                 'Row Delete
            .Buttons(9).Enabled = False                 'Row Cancel
            .Buttons(11).Enabled = False                'Copy
            .Buttons(12).Enabled = False                'Paste
            .Buttons(14).Enabled = True                 'Excel
        End With
        
    End If
    
    ss1.OperationMode = OperationModeNormal

End Sub

Public Sub Form_Pro()

    Dim mResult As String
    Dim sMsg As String
    
    Mode = ""
 
    If opt_move = True Then
    
        Mode = "M"
        
        If txt_from.Text <> "" Or txt_to.Text <> "" Or txt_target.Text <> "" Then
            sMsg = "确定要把板坯从(" + txt_from.Text + ")->(" + txt_to.Text + ")" + "调整到板坯(" + txt_target.Text + ")后边吗？"
        Else
            sMsg = "必须输入起始板坯号和目标板坯号！"
            Call Gp_MsgBoxDisplay(sMsg)
            Exit Sub
        End If
        
        sMsg = sMsg + "调整后相应的作业指示将被取消！"
        mResult = MsgBox(sMsg, vbYesNo)
        
        If mResult = vbYes Then
            If Gp_Process_Exec = "" Then
               MsgBox ("作业指示调整完毕 ！")
               Call Form_Ref
            Else
               MsgBox (Gp_Process_Exec + " 作业指示调整失败！")
            End If
        End If
    
    End If
 
    If opt_delete = True Then
        
        Mode = "D"
        
        If txt_from.Text = "" Then
           sMsg = "必须输入起始板坯号！"
           Call Gp_MsgBoxDisplay(sMsg)
           Exit Sub
        End If
        sMsg = "确定要删除选定板坯(" + txt_from.Text + ")" + ")吗？"
        
        If txt_to.Text <> "" Then
           sMsg = "确定要删除选定板坯(" + txt_from.Text + ")->(" + txt_to.Text + ")吗？"
        End If
        mResult = MsgBox(sMsg, vbYesNo)
        
        If mResult = vbYes Then
           If Gp_Process_Exec = "" Then
              MsgBox ("作业指示删除完毕 ！")
              Call Form_Ref
           Else
              MsgBox (Gp_Process_Exec + " 作业指示删除失败！")
           End If
        End If
    End If
 
    If opt_cnf = True Then
        
        Mode = "F"
        
        If txt_from.Text = "" Then
           sMsg = "必须输入起始板坯号！"
           Call Gp_MsgBoxDisplay(sMsg)
           Exit Sub
        End If
        sMsg = "确定要指示选定板坯(" + txt_from.Text + ")" + ")吗？"
        
        If txt_to.Text <> "" Then
           sMsg = "确定要指示选定板坯(" + txt_from.Text + ")->(" + txt_to.Text + ")吗？"
        End If
        mResult = MsgBox(sMsg, vbYesNo)
        
        If mResult = vbYes Then
           If Gp_Process_Exec = "" Then
              MsgBox ("作业指示完毕 ！")
              Call Form_Ref
           Else
              MsgBox (Gp_Process_Exec + " 作业指示失败！")
           End If
        End If
    End If
 
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
    End With
    
End Sub

Public Sub Form_Ins()
    
'    Call Gp_Sp_Ins(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Cpy()

'    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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
    
'    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub opt_cancel_Click(VALUE As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_cancel.VALUE = True Then
        opt_cancel.ForeColor = &HFF&
        opt_sent.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_from.Enabled = True
        opt_to.Enabled = False
        opt_target.Enabled = False
    Else
        opt_cancel.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_cnf_Click(VALUE As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_cnf.VALUE = True Then
    
        opt_cnf.ForeColor = &HFF&
        opt_delete.ForeColor = &H808080
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_from.Enabled = True
        opt_to.Enabled = True
        opt_target.Enabled = False
    Else
        opt_cnf.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0

End Sub

Private Sub opt_delete_Click(VALUE As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_delete.VALUE = True Then
    
        opt_delete.ForeColor = &HFF&
        opt_cnf.ForeColor = &H808080
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_from.Enabled = True
        opt_to.Enabled = True
        opt_target.Enabled = False
    Else
        opt_delete.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_from_Click(VALUE As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_from.VALUE = True Then
        opt_from.ForeColor = &HFF&
        opt_to.ForeColor = &H808080
        opt_target.ForeColor = &H808080
    Else
        opt_from.ForeColor = &H808080
    End If
    
End Sub

Private Sub opt_move_Click(VALUE As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_move.VALUE = True Then
        opt_move.ForeColor = &HFF&
        opt_cnf.ForeColor = &H808080
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_from.Enabled = True
        opt_to.Enabled = True
        opt_target.Enabled = True
    Else
        opt_move.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_sent_Click(VALUE As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_sent.VALUE = True Then
        opt_sent.ForeColor = &HFF&
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_from.Enabled = False
        opt_to.Enabled = True
        opt_target.Enabled = False
    Else
        opt_sent.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_target_Click(VALUE As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_target.VALUE = True Then
        opt_target.ForeColor = &HFF&
        opt_from.ForeColor = &H808080
        opt_to.ForeColor = &H808080
    Else
        opt_target.ForeColor = &H808080
    End If
    
End Sub

Private Sub opt_to_Click(VALUE As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_to.VALUE = True Then
        opt_to.ForeColor = &HFF&
        opt_from.ForeColor = &H808080
        opt_target.ForeColor = &H808080
    Else
        opt_to.ForeColor = &H808080
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim SE As String
    Dim C, M As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    Dim SEND_SLAB As String

    If Gf_Sp_Change(Proc_Sc, Sc1) Then
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0
    End If
    
    If Row < 1 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 6
    
    If opt_from.VALUE = True Then
        txt_from.Text = ss1.Text
        
        ss1.Col = 23
        sSlab_Edt_Seq_Fr = ss1.Text
    End If
    
    If opt_to.VALUE = True Then
        txt_to.Text = ss1.Text
        
        ss1.Col = 23
        sSlab_Edt_Seq_To = ss1.Text
    End If
    
    If opt_target.VALUE = True Then
        txt_target.Text = ss1.Text
        
        ss1.Col = 23
        sSlab_Edt_Seq_Tg = ss1.Text
    End If
    
End Sub

Private Sub SSPanel1_Click()
    
    opt_sent.VALUE = False
    opt_cancel.VALUE = False
    opt_move.VALUE = False
    opt_delete.VALUE = False
    opt_from.VALUE = False
    opt_to.VALUE = False
    opt_target.VALUE = False
    opt_sent.ForeColor = &H808080
    opt_move.ForeColor = &H808080
    opt_delete.ForeColor = &H808080
    opt_cancel.ForeColor = &H808080
    opt_from.ForeColor = &H808080
    opt_to.ForeColor = &H808080
    opt_target.ForeColor = &H808080
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=TXT_PLT
        DD.rControl.Add Item:=TXT_PLT_NAME
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(TXT_PLT.Text)) = TXT_PLT.MaxLength Then
        TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(TXT_PLT.Text), 2)
    Else
        TXT_PLT_NAME.Text = ""
    End If

End Sub

Public Function Gp_Process_Exec() As String

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iRow As Integer
    Dim adoCmd As ADODB.Command
    
    Dim sSlab_Seq_Fr As String
    Dim sSlab_Seq_To As String
    Dim sSlab_Seq_Tg As String
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sSlab_Seq_Fr = sSlab_Edt_Seq_Fr
    sSlab_Seq_To = sSlab_Edt_Seq_To
    sSlab_Seq_Tg = sSlab_Edt_Seq_Tg
    
    sQuery = "{call AFZ1000P ('" + Mode + "','" + "M" + "','" + sSlab_Seq_Fr + "','" + sSlab_Seq_To + "','" + sSlab_Seq_Tg + "','" + sUserID + "',?)}"
    
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
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        
        Screen.MousePointer = vbDefault
        Gp_Process_Exec = sErrMessg
        Set adoCmd = Nothing
        Exit Function
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Gp_Process_Exec = ""
    Exit Function

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Gp_Process_Exec = "Process_Exec_ERROR"
    Err.Raise Err.Number, Err.Description & sQuery
    
End Function
