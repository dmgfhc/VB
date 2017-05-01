VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "CSText32.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form ACB4050C 
   Caption         =   "板坯使用实绩查询_ACB4050C"
   ClientHeight    =   9225
   ClientLeft      =   555
   ClientTop       =   1575
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   11295
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text_out_plt_cd_name 
      Height          =   315
      Left            =   1470
      TabIndex        =   14
      Top             =   495
      Width           =   1245
   End
   Begin VB.TextBox txt_out_plt_cd 
      Height          =   315
      Left            =   1125
      MaxLength       =   1
      TabIndex        =   13
      Tag             =   "板坯去向"
      Top             =   495
      Width           =   330
   End
   Begin VB.TextBox Text_out_plt_name 
      Height          =   315
      Left            =   1470
      TabIndex        =   12
      Top             =   105
      Width           =   1245
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8220
      Left            =   60
      TabIndex        =   11
      Top             =   975
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   14499
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
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
      MaxCols         =   13
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "ACB4050C.frx":0000
   End
   Begin VB.TextBox txt_prod_cd 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4365
      MaxLength       =   2
      TabIndex        =   8
      Tag             =   "产品"
      Text            =   "SL"
      Top             =   480
      Width           =   345
   End
   Begin VB.TextBox txt_slab_no 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4365
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "产品"
      Top             =   90
      Width           =   1170
   End
   Begin VB.TextBox txt_stlgrd 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6870
      TabIndex        =   1
      Tag             =   "钢种"
      Top             =   105
      Width           =   1260
   End
   Begin VB.TextBox Text_out_plt 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1125
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "使用工厂"
      Top             =   105
      Width           =   330
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   180
      Top             =   105
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "使用工厂"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   180
      Top             =   495
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "板坯去向"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   5895
      Top             =   105
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "钢种"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   3420
      Top             =   90
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "板坯号"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   12135
      Top             =   480
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "重量合计"
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
   Begin CSTextLibCtl.sidbEdit Text_TOT_WGT 
      Height          =   315
      Left            =   13065
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   1470
      _Version        =   262145
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   255
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
      ReadOnly        =   -1  'True
      Insert          =   0   'False
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.000"
      Text            =   " 0.00"
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   5895
      Top             =   480
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "使用日期"
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
   Begin InDate.UDate UDate_IN_PLT_DATE_a 
      Height          =   315
      Left            =   6870
      TabIndex        =   4
      Tag             =   "使用日期"
      Top             =   480
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
      BackColor       =   12648447
   End
   Begin InDate.UDate UDate_IN_PLT_DATE_b 
      Height          =   315
      Left            =   8520
      TabIndex        =   5
      Tag             =   "使用日期"
      Top             =   480
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
      BackColor       =   12648447
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   3420
      Top             =   480
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "产品"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   12135
      Top             =   105
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   556
      Caption         =   "数量合计"
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
   Begin CSTextLibCtl.sidbEdit Text_TOT_SHEETS 
      Height          =   315
      Left            =   13065
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   105
      Width           =   1470
      _Version        =   262145
      _ExtentX        =   2593
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0.00"
      ForeColor       =   16711680
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
      ReadOnly        =   -1  'True
      Insert          =   0   'False
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0.00"
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
      MaxValue        =   9999999.9
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   165
      X2              =   15120
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   165
      X2              =   15120
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14565
      TabIndex        =   10
      Top             =   225
      Width           =   195
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "吨"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14565
      TabIndex        =   7
      Top             =   600
      Width           =   195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   135
      Left            =   8295
      TabIndex        =   6
      Top             =   570
      Width           =   255
   End
End
Attribute VB_Name = "ACB4050C"
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
'-- Program ID        ACB4050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          MENGDAN
'-- Coder             MENGDAN
'-- Date              2005.8.30
'-- Description
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
Public STR1 As String
Public BASE As String
Public AIMNO As String
Dim sQuery As String

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

Private Sub Form_Define()
        
     'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     ' FormType = "Msheet"
   FormType = "Refer"
         
           Call Gp_Ms_Collection(Text_out_plt, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_out_plt_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udate_in_plt_date_a, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(udate_in_plt_date_b, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_slab_no, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_prod_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                                              
      'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    ' control part   Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
 '  Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    'Sc1.Add Item:="ACB4050C.P_REFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

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
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    
    Call Form_Define
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
   
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)

    udate_in_plt_date_a.Text = Format(Date, "YYYY-MM-01")

    udate_in_plt_date_b.RawData = Gf_GetLastDay(udate_in_plt_date_b.RawData)
    Screen.MousePointer = vbDefault
    Text_out_plt.Text = "C1"
    'text_cur_inv_code.Text = "00"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '查询结束

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    Set rControl = Nothing
    
    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    End If
 
    udate_in_plt_date_a.Text = Format(Date, "YYYY-MM-01")

    udate_in_plt_date_b.RawData = Gf_GetLastDay(udate_in_plt_date_b.RawData)
    Text_out_plt_name.Text = ""
    text_tot_sheets.Value = 0
    text_tot_wgt.Value = 0
    Text_out_plt_cd_name.Text = ""
    
End Sub

Public Sub Form_Exc()
    
    If Trim(Text_out_plt.Text) = "" Then
        Call Gp_MsgBoxDisplay(Text_out_plt.Tag & "必须输入")
        Exit Sub
    End If
    
'    If Trim(txt_TO_INV_name.Text) = "" Then
'        Call Gp_MsgBoxDisplay(txt_TO_INV.Tag & "必须输入")
'        Exit Sub
'    End If
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)


End Sub


Public Sub Form_Ref()
    Dim I As Integer
    Dim TotalWeight As Double
    Dim TotalSheets As Double
    Dim minDATE As String
    Dim maxDATE As String
    
'    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
'
'    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
'        'Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'        With MDIMain.MenuTool
'            '.Buttons(8).Enabled = True                  'Row Delete
'            .Buttons(9).Enabled = False                 'Row Cancel
'            .Buttons(14).Enabled = True                 'Excel
'        End With
'        ss1.OperationMode = OperationModeNormal
'    End If

     Dim SMESG As String
     Dim S As String
     

    minDATE = udate_in_plt_date_a.RawData
    maxDATE = udate_in_plt_date_b.RawData
    
    If udate_in_plt_date_a.RawData <> "" And udate_in_plt_date_b.RawData <> "" Then
        If maxDATE >= minDATE Then
          sQuery = "Select SLAB_NO,GF_STLGRD_DETAIL(STLGRD),THK,WID,LEN,WGT,Gf_ComnNameFind('B0043',SIZE_KND),MIXED_FL,TO_DATE(PROD_DATE,'YYYY-MM-DD'),TO_DATE(OUT_PLT_DATE,'YYYY-MM-DD'),Gf_ComnNameFind('C0011',OUT_PLT_CD),PROD_CD,GF_COMNNAMEFIND('C0013',CUR_INV)"
          sQuery = sQuery + "  From FP_SLAB "
          sQuery = sQuery + " Where OUT_PLT    Like '" + Trim(Text_out_plt.Text) + "' ||'%' "
          sQuery = sQuery + "   AND OUT_PLT_DATE BETWEEN '" + minDATE + "' AND '" + maxDATE + "' "
          sQuery = sQuery + "AND REC_STS  = '3' "
          If Trim(txt_out_plt_cd.Text) <> "" Then
          sQuery = sQuery + "AND OUT_PLT_CD       = '" + Trim(txt_out_plt_cd.Text) + "' "
          End If
          If Trim(txt_stlgrd.Text) <> "" Then
          sQuery = sQuery + "AND STLGRD       = '" + Trim(txt_stlgrd.Text) + "' "
          End If
          If Trim(txt_slab_no.Text) <> "" Then
          sQuery = sQuery + "AND SLAB_NO       LIKE '" + Trim(txt_slab_no.Text) + "' ||'%' "
          End If
          If Trim(txt_prod_cd.Text) <> "" Then
          sQuery = sQuery + "AND PROD_CD       = '" + Trim(txt_prod_cd.Text) + "' "
          End If
          sQuery = sQuery + " ORDER BY OUT_PLT_DATE DESC, SLAB_NO ASC "
          
        Else
             Call MsgBox("输入日期不符合规范!" & Chr(10) & "请重试。", vbExclamation + vbOKOnly, "警告")
        End If
        
        SMESG = Gf_Ms_NeceCheck(nControl)
        If SMESG = "OK" Then
        
            SMESG = Gf_Ms_NeceCheck2(mControl)
            If SMESG = "OK" Then
            
                If Gf_Sp_Display(M_CN1, ss1, sQuery) Then
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
            Else
                SMESG = SMESG + " Must input according to length of item"
                Call Gp_MsgBoxDisplay(SMESG)
            End If
        Else
            SMESG = SMESG + " Must input necessarily"
            Call Gp_MsgBoxDisplay(SMESG)
        End If

        
        With ss1
            If .MaxRows = 0 Then
                text_tot_sheets.Text = "0"
                text_tot_wgt.Value = 0
            Else
                For I = 1 To .MaxRows
                    .Row = I
                    .Col = 6: TotalWeight = .Value + TotalWeight
                Next I
                text_tot_sheets.Text = Str$(TotalSheets)
                text_tot_wgt.Text = Str$(TotalWeight)
            End If
        End With
        text_tot_sheets.Text = ss1.MaxRows
    Else
       Call MsgBox("查询日期范围不能为空!", vbExclamation + vbOKOnly, "警告")
    End If

    
    
    
End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

  '  Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
'   Call ss1_row_Click(Col, Row)

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    ss1.Row = ss1.ActiveRow
    
    ss1.Col = 1
    txt_slab_no.Text = ss1.Text
    
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
Private Function Gf_GetLastDay(Optional DTDay As String = "") As Variant

On Error GoTo DGet_Error

    Dim sQuery As String
    Dim strDay As String
    
    If DTDay = "" Then
        sQuery = "SELECT TO_CHAR(LAST_DAY(SYSDATE),'YYYYMMDD') FROM DUAL"
    Else
       strDay = DTDay
       sQuery = "SELECT TO_CHAR(LAST_DAY(TO_DATE('" + strDay + "','YYYYMMDD')),'YYYYMMDD') FROM DUAL"
    End If
       
    Dim AdoRs As ADODB.Recordset
    
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
    
        If Not AdoRs.EOF Then
            If VarType(AdoRs.Fields(0)) = vbNull Then
                Gf_GetLastDay = ""
            Else
                Gf_GetLastDay = AdoRs.Fields(0)
            End If
        End If
        
    Else
        Gf_GetLastDay = "00000000"
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    Exit Function

DGet_Error:

    Set AdoRs = Nothing
    Gf_GetLastDay = "00000000"

End Function

Private Sub Text_out_plt_Change()
    If Len(Trim(Text_out_plt.Text)) = Text_out_plt.MaxLength Then
      Text_out_plt_name.Text = Gf_ComnNameFind(M_CN1, "C0001", Text_out_plt.Text, 2)
      Exit Sub
Else
      Text_out_plt_name.Text = ""
End If
End Sub

Private Sub Text_out_plt_DblClick()

    Call Text_out_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_out_plt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Text_out_plt_name.Text = ""
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0001"

        DD.rControl.Add Item:=Text_out_plt
        DD.rControl.Add Item:=Text_out_plt_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If
End Sub

Private Sub txt_out_plt_cd_Change()
    If Len(Trim(txt_out_plt_cd.Text)) = txt_out_plt_cd.MaxLength Then
      Text_out_plt_cd_name.Text = Gf_ComnNameFind(M_CN1, "C0011", txt_out_plt_cd.Text, 2)
      Exit Sub
    Else
      Text_out_plt_cd_name.Text = ""
    End If

End Sub

Private Sub txt_out_plt_cd_DblClick()

    Call txt_out_plt_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_out_plt_cd_KeyUp(KeyCode As Integer, Shift As Integer)
    'Text_out_plt_name.Text = ""
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "C0011"

        DD.rControl.Add Item:=txt_out_plt_cd
        DD.rControl.Add Item:=Text_out_plt_cd_name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If
'    If Len(Trim(txt_out_plt_cd.Text)) = txt_out_plt_cd.MaxLength Then
'        Text_out_plt_cd_name.Text = Gf_ComnNameFind(M_CN1, "C0011", txt_out_plt_cd.Text, 2)
'    Else
'        Text_out_plt_cd_name.Text = ""
'    End If
End Sub

Private Sub txt_prod_cd_DblClick()

    Call txt_prod_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

   'Text_PROD_CD_Name.Text = ""
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=txt_prod_cd
        'DD.rControl.Add Item:=Text_PROD_CD_Name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If

    'If Len(Trim(Text_PROD_CD.Text)) = Text_PROD_CD.MaxLength Then
       ' Text_PROD_CD_Name.Text = Gf_ComnNameFind(M_CN1, "B0005", Text_PROD_CD.Text, 2)
    'Else
        'Text_PROD_CD_Name.Text = ""
    'End If
    
End Sub

Private Sub txt_slab_no_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=txt_slab_no
        'DD.rControl.Add Item:=Text_PROD_CD_Name
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
        
    End If
End Sub

Private Sub txt_stlgrd_DblClick()

    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
        
        DD.sWitch = "MS"
        DD.rControl.Add Item:=txt_stlgrd
        
        DD.nameType = "1"
        Call Gf_Stlgrd_DD(M_CN1, KeyCode)
        
    End If

End Sub
