VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1200C 
   Caption         =   "销售计划录入_AAA1200C"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.ComboBox txt_prod_cd 
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
      ItemData        =   "AAA1200C.frx":0000
      Left            =   5910
      List            =   "AAA1200C.frx":000D
      TabIndex        =   16
      Tag             =   "产品代码"
      Text            =   "PP"
      Top             =   120
      Width           =   870
   End
   Begin VB.ComboBox cbo_plt 
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
      ItemData        =   "AAA1200C.frx":001D
      Left            =   3765
      List            =   "AAA1200C.frx":0027
      TabIndex        =   15
      Tag             =   "工厂"
      Top             =   120
      Width           =   870
   End
   Begin VB.TextBox txt_ref_type 
      Height          =   270
      Left            =   12090
      TabIndex        =   13
      Text            =   "1"
      Top             =   9930
      Visible         =   0   'False
      Width           =   615
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7980
      Left            =   105
      TabIndex        =   10
      Top             =   1170
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   14076
      _Version        =   196609
      PaneTree        =   "AAA1200C.frx":0033
      Begin FPSpread.vaSpread ss1 
         Height          =   3270
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   14985
         _Version        =   393216
         _ExtentX        =   26432
         _ExtentY        =   5768
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
         MaxCols         =   0
         MaxRows         =   0
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AAA1200C.frx":0085
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   4560
         Left            =   30
         TabIndex        =   12
         Top             =   3390
         Width           =   14985
         _Version        =   393216
         _ExtentX        =   26432
         _ExtentY        =   8043
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
         MaxCols         =   13
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AAA1200C.frx":032E
      End
   End
   Begin VB.TextBox txt_stlgrd_des 
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
      Left            =   2670
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   645
      Width           =   4125
   End
   Begin VB.TextBox txt_stlgrd 
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
      Left            =   1335
      MaxLength       =   11
      TabIndex        =   4
      Tag             =   "钢种"
      Top             =   645
      Width           =   1335
   End
   Begin VB.TextBox txt_cust_cd 
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
      Left            =   8100
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "客户"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txt_cust_name 
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
      Left            =   8970
      MaxLength       =   40
      TabIndex        =   2
      Top             =   120
      Width           =   3405
   End
   Begin Threed.SSCommand SCmd2 
      Height          =   375
      Left            =   12480
      TabIndex        =   0
      Top             =   645
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "上传Excel"
   End
   Begin Threed.SSCommand SCmd1 
      Height          =   375
      Left            =   12480
      TabIndex        =   1
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "计划详细查询"
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   105
      Top             =   645
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "钢种"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   4665
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "产品"
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
   Begin InDate.UDate dtp_yy_mm 
      Height          =   315
      Left            =   1335
      TabIndex        =   6
      Tag             =   "日期"
      Top             =   120
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Text            =   "____-__"
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
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   105
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "年月"
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
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   6870
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      Caption         =   "客户"
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
   Begin InDate.UDate dtp_copy_to 
      Height          =   315
      Left            =   10140
      TabIndex        =   7
      Top             =   645
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Text            =   "____-__"
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
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel9 
      Height          =   315
      Left            =   9165
      Top             =   645
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "复制到"
      Alignment       =   1
      BackColor       =   16777088
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
   Begin InDate.UDate dtp_copy_from 
      Height          =   315
      Left            =   7935
      TabIndex        =   8
      Top             =   645
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      Text            =   "____-__"
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
      Mask            =   "%%%%-%%"
      MaxLength       =   7
   End
   Begin InDate.ULabel ULabel7 
      Height          =   315
      Left            =   6960
      Top             =   645
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "从"
      Alignment       =   1
      BackColor       =   16777088
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   315
      Left            =   11370
      TabIndex        =   9
      Top             =   645
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "复制"
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   2520
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
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
   Begin Threed.SSCommand SCmd3 
      Height          =   375
      Left            =   13935
      TabIndex        =   14
      Top             =   120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "订单查询"
   End
   Begin VB.Shape Shape1 
      Height          =   540
      Left            =   6870
      Top             =   525
      Width           =   5505
   End
End
Attribute VB_Name = "AAA1200C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       production plan
'-- Sub_System Name
'-- Program Name
'-- Program ID        AAA1200C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
'-- Date              2009.6.10
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

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
    
    Dim sQuery As String
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
             Call Gp_Ms_Collection(dtp_yy_mm, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(cbo_plt, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_cust_cd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stlgrd, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_ref_type, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
                     
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    Sc1.Add Item:="AAA1200C.P_SREFER", Key:="P-R"
    Sc1.Add Item:="AAA1200C.P_SMODIFY", Key:="P-M"
    Sc1.Add Item:=ss1, Key:="Spread"
    
    'Spread_Collection
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="AAA1200C.P_SREFER", Key:="P-R"
    Sc2.Add Item:="AAA1200C.P_UPLOAD", Key:="P-M"
    
    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call Menu_Setting

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
    Call Menu_Setting
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))

    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    
    Call Gp_Sp_ColGet(ss2, "A-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
    cbo_plt.Text = "C1": txt_prod_cd.Text = "PP"
    
    If Mid(sAuthority, 1, 3) = "111" Then
       SCmd1.Enabled = True
       SCmd3.Enabled = True
       SSCommand1.Enabled = True
       SCmd2.Enabled = True
    ElseIf Mid(sAuthority, 1, 1) = "1" Then
       SCmd1.Enabled = True
       SCmd3.Enabled = False
       SSCommand1.Enabled = False
       SCmd2.Enabled = False
    Else
       SCmd1.Enabled = False
       SCmd3.Enabled = False
       SSCommand1.Enabled = False
       SCmd2.Enabled = False
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(ss2, "A-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    
    ss1.MaxCols = 0
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(Sc2)
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Menu_Setting
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    rControl(1).SetFocus
    
    txt_cust_name.Text = ""
    txt_stlgrd_des.Text = ""

End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    Dim iCol As Integer
    Dim ColTot As Double
    Dim RowTot As Double
    txt_ref_type.Text = "R"
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
        
        If Sp_Header_Refer() Then
            With ss1
                 If txt_cust_cd.Text = "" Or txt_stlgrd.Text = "" Then
                    .Protect = True
                    .Col = 1:      .Col2 = -1
                    .Row = 1:      .Row2 = -1
                    .BlockMode = True
                    .Lock = True
                    .BackColor = &H80000005
                    .BlockMode = False
                    .OperationMode = OperationModeNormal
                 Else
                    .Protect = False
                    .Col = 1:      .Col2 = -1
                    .Row = 1:      .Row2 = -1
                    .BlockMode = True
                    .Lock = False
                    .BackColor = &HC0FFFF
                    .BlockMode = False
                    .OperationMode = OperationModeNormal
                 End If
            End With

            If Sp_Data_Refer() Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Menu_Setting
'                Call Gp_Ms_ControlLock(Mc1!lControl, True)
                With ss1
                    .MaxRows = .MaxRows + 1
                    .MaxCols = .MaxCols + 1
                    '列合计
                    .Row = .MaxRows
                    .Col = SpreadHeader
                    .Text = "合计"
                    .Col = SpreadHeader + 1
                    .Text = "合计"
                    
                    .Row = .MaxRows:       .Row2 = .MaxRows
                    .Col = SpreadHeader:   .Col2 = SpreadHeader + 1
                    .ColMerge = MergeAlways
                    .RowMerge = MergeAlways
                    
                    For iCol = 1 To .MaxCols - 1
                        .Col = iCol
                        ColTot = Gf_Sp_ColSum(ss1, .Col, 1, .MaxRows - 1)
                        .Row = .MaxRows
                        If ColTot > 0 Then
                            .Text = ColTot
                        Else
                            .Text = ""
                        End If
                        .CellType = CellTypeNumber
                        .TypeNumberDecPlaces = 3
                        .TypeNumberMax = 999999999
                        .TypeNumberMin = 0
                        .TypeNumberShowSep = True
                        .TypeHAlign = TypeHAlignRight
                        .TypeVAlign = TypeVAlignCenter
                    Next iCol
                    
                    '行合计
                    .Col = .MaxCols
                    .Row = SpreadHeader
                    .Text = "合计"
                    .Row = SpreadHeader + 1
                    .Text = "合计"
                    
                    .Col = .MaxCols:       .Col2 = .MaxCols
                    .Row = SpreadHeader:   .Row2 = SpreadHeader
                    .ColMerge = MergeAlways
                    .RowMerge = MergeAlways
                    
                    For iCol = 1 To .MaxRows
                        .Row = iCol
                        RowTot = Gf_Sp_RowSum(ss1, .Row, 1, .MaxCols - 1)
                        .Col = .MaxCols
                        If RowTot > 0 Then
                            .Text = RowTot
                        Else
                            .Text = ""
                        End If
                        .CellType = CellTypeNumber
                        .TypeNumberDecPlaces = 3
                        .TypeNumberMax = 999999999
                        .TypeNumberMin = 0
                        .TypeNumberShowSep = True
                        .TypeHAlign = TypeHAlignRight
                        .TypeVAlign = TypeVAlignCenter
                
                        .ColWidth(.Col) = 11
                    Next iCol
                    
                End With
            End If
        End If
    Else
        sMesg = sMesg + " 必须输入...."
        Call Gp_MsgBoxDisplay(sMesg)
    End If
End Sub

Public Sub Form_Pro()

    Dim sMesg As String
    Dim sInput As String
    Dim sDate As String

    ss2.Col = 0
    ss2.Row = 1
    sInput = ss2.Text

If sInput = "Input" Then '导入
   ss2.Col = 1
   sDate = Mid(ss2.Text, 1, 4) + Mid(ss2.Text, 6, 2)
   If sDate <> dtp_yy_mm.RawData Then
      MsgBox "查询年月与导入年月不一致！", vbCritical, "系统提示信息"
      Exit Sub
   End If
   
   If dtp_yy_mm.RawData = "" Or cbo_plt.Text = "" Or txt_prod_cd.Text = "" Then
        MsgBox "导入年月，工厂，产品必须输入！"
        Exit Sub
   End If
   
   If dtp_yy_mm.RawData < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Then
        sMesg = " 历史月份的数据不可再次导入！"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    Call History_Del(dtp_yy_mm.RawData, cbo_plt.Text, txt_prod_cd.Text)
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc2"), Mc1, True) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
        Call Gp_Sp_BlockLock(ss2, 9, 9, 1, ss2.MaxRows, True)
    End If
Else                    '保存上表数据
    If dtp_yy_mm.RawData = "" Or Trim(cbo_plt.Text) = "" Or Trim(txt_prod_cd.Text) = "" Or Trim(txt_cust_cd.Text) = "" Or Trim(txt_stlgrd.Text) = "" Then
        sMesg = "年月，工厂，产品，客户，钢种必须全部输入！"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If Sp_Process(M_CN1, Proc_Sc("Sc1")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
End If
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
If txt_ref_type.Text = "1" Then
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc1")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
ElseIf txt_ref_type.Text = "2" Then
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End If
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub SCmd1_Click()
On Error GoTo Refer_Err
    
    txt_ref_type.Text = "D"
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc2").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        txt_cust_cd.Enabled = True
        txt_stlgrd.Enabled = True
        ss2.OperationMode = OperationModeNormal
        Exit Sub
    End If
            
    Exit Sub

Refer_Err:
End Sub

Private Sub SCmd2_Click()
   Load frm_Excel
   frm_Excel.txt_load_file.Text = "AAA1200C"
   frm_Excel.Show 1
End Sub

Private Sub SCmd3_Click()
Dim iCol As Integer
Dim ColTot As Double
Dim RowTot As Double

If dtp_yy_mm.RawData = "" Or Trim(cbo_plt.Text) = "" Or Trim(txt_prod_cd.Text) = "" Then
    MsgBox "年月，工厂，产品必须输入！", vbCritical, "系统提示信息"
    Exit Sub
End If

If Sp_Header_Refer() Then
   Call Sp_Order_Refer
   txt_ref_type.Text = "P"
   Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc1, Mc1("nControl"), Mc1("mControl"))
   Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
   Call Menu_Setting
   txt_cust_cd.Enabled = True
   txt_stlgrd.Enabled = True
   
   With ss1
     If txt_cust_cd.Text = "" Or txt_stlgrd.Text = "" Then
        Call Gp_Sp_BlockLock(ss1, 1, .MaxCols, 1, .MaxRows, True)
        Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, 1, .MaxRows, BLACK, &H80000005)
        ss2.OperationMode = OperationModeNormal
     Else
        Call Gp_Sp_BlockLock(ss1, 1, .MaxCols, 1, .MaxRows, False)
        Call Gp_Sp_BlockColor(ss1, 1, .MaxCols, 1, .MaxRows, BLACK, &HC0FFFF)
        ss2.OperationMode = OperationModeNormal
     End If
     
    .MaxRows = .MaxRows + 1
    .MaxCols = .MaxCols + 1
    '列合计
    .Row = .MaxRows
    .Col = SpreadHeader
    .Text = "合计"
    .Col = SpreadHeader + 1
    .Text = "合计"
    
    .Row = .MaxRows:       .Row2 = .MaxRows
    .Col = SpreadHeader:   .Col2 = SpreadHeader
    .ColMerge = MergeAlways
    .RowMerge = MergeAlways
    
    For iCol = 1 To .MaxCols - 1
        .Col = iCol
        ColTot = Gf_Sp_ColSum(ss1, .Col, 1, .MaxRows - 1)
        .Row = .MaxRows
        .Text = ColTot
    Next iCol
    
    '行合计
    .Col = .MaxCols
    .Row = SpreadHeader
    .Text = "合计"
    .Row = SpreadHeader + 1
    .Text = "合计"
    
    .Col = .MaxCols:       .Col2 = .MaxCols
    .Row = SpreadHeader:   .Row2 = SpreadHeader
    .ColMerge = MergeAlways
    .RowMerge = MergeAlways
    
    For iCol = 1 To .MaxRows
        .Row = iCol
        RowTot = Gf_Sp_RowSum(ss1, .Row, 1, .MaxCols - 1)
        .Col = .MaxCols
        .Text = RowTot
        .CellType = CellTypeNumber
        .TypeNumberDecPlaces = 3
        .TypeNumberMax = 999999999
        .TypeNumberMin = 0
        .TypeNumberShowSep = True
        .TypeHAlign = TypeHAlignRight

        .ColWidth(.Col) = 11
    Next iCol

End With

End If
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
txt_ref_type.Text = "1"
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
 '      Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc1")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc1")("Spread").EditMode = True

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
        MDIMain.Mnu_Sorting.Visible = False
        MDIMain.Line1.Visible = False
        
        PopupMenu MDIMain.PopUp_Spread
        
        MDIMain.Mnu_Sorting.Visible = True
        MDIMain.Line1.Visible = True
    End If

End Sub

Public Sub Menu_Setting()

    MDIMain.MenuTool.Buttons(5).Enabled = False    'Delete
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Row cancel
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Row cancel
    
End Sub


Public Function Sp_Header_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
    Dim sQuery2 As String
    
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    Dim AdoRs2 As ADODB.Recordset
    Dim ArrayRecords2 As Variant

    Set adoRs = New ADODB.Recordset
    
    sQuery = "SELECT THK_CD, FR_THK, TO_THK "
    sQuery = sQuery + "   FROM BP_THICK_GRP "
    sQuery = sQuery + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    If txt_prod_cd.Text <> "**" Then
       sQuery = sQuery + "    AND THK_CD <> '*' "
    End If
    sQuery = sQuery + "  ORDER BY THK_CD "
    
    With ss1

        Sp_Header_Refer = True
        .ReDraw = False
        .MaxRows = 0:  .MaxCols = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        adoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If adoRs.BOF Or adoRs.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            adoRs.Close
            Set adoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = adoRs.GetRows
        adoRs.Close
        Set adoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = UBound(ArrayRecords, 2) + 1
            For iCol = 0 To .MaxCols - 1
            
               .Col = iCol + 1
               .Row = SpreadHeader + 1
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt))
                    .Row = SpreadHeader
                    .Text = ArrayRecords(0, iCnt)
                End If
                                  
                'Column Type Setting
                .Col = iCol + 1: .Col2 = iCol + 1
                .Row = 1: .Row2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 3
                .TypeNumberMax = 999999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                .ColWidth(iCol + 1) = 11
                iCnt = iCnt + 1
                
            Next iCol
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Set AdoRs2 = New ADODB.Recordset
    
    sQuery2 = "SELECT WID_CD, FR_WID, TO_WID "
    sQuery2 = sQuery2 + "   FROM BP_WIDTH_GRP "
    sQuery2 = sQuery2 + "  WHERE PROD_CD = '" + txt_prod_cd.Text + "' "
    If txt_prod_cd.Text <> "**" Then
       sQuery2 = sQuery2 + "    AND WID_CD <> '*' "
    End If
    sQuery2 = sQuery2 + "  ORDER BY WID_CD "
    
    With ss1

        Sp_Header_Refer = True
        .ColWidth(SpreadHeader + 1) = 13
        Screen.MousePointer = vbHourglass
        'Ado Execute
        AdoRs2.Open sQuery2, M_CN1, adOpenKeyset
        
        If AdoRs2.BOF Or AdoRs2.EOF Then
        
            Sp_Header_Refer = False
            '.ReDraw = True
            AdoRs2.Close
            Set AdoRs2 = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords2 = AdoRs2.GetRows
        AdoRs2.Close
        Set AdoRs2 = Nothing

        If UBound(ArrayRecords2, 2) + 1 <> 0 Then
        
            .MaxRows = UBound(ArrayRecords2, 2) + 1
            iCnt = 0
            
            For iRow = 1 To .MaxRows
            
                .Row = iRow
                .Col = SpreadHeader + 1
                
                If VarType(ArrayRecords2(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords2(1, iCnt)) & " ~ " & Trim(ArrayRecords2(2, iCnt))
                    .Col = SpreadHeader
                    .Text = ArrayRecords2(0, iCnt)
                End If
                
                .Row = iRow + 2: .Row2 = iRow + 2
                .Col = 1: .Col2 = -1
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 3
                .TypeNumberMax = 999999999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
                
                iCnt = iCnt + 1
            Next iRow
                
        End If
        
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
        
    'Call Gp_Sp_EvenRowBackcolor(Sc1.Item("Spread"), 0)
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Set AdoRs2 = Nothing
    ss1.ReDraw = True
    Sp_Header_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
   ' Dim SPARA As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    
    sQuery = "{CALL AAA1200C.P_SREFER('" + dtp_yy_mm.RawData + "','" + cbo_plt.Text + "','" + txt_prod_cd.Text + "','" + txt_cust_cd.Text + "','" + txt_stlgrd.Text + "','R')}"
    
    'Ado Execute
    adoRs.Open sQuery, M_CN1, adOpenKeyset
    
    With ss1

        Sp_Data_Refer = True
        .ReDraw = False
       ' .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        If adoRs.BOF Or adoRs.EOF Then
        
            Sp_Data_Refer = False
            .ReDraw = True
            adoRs.Close
            Set adoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = adoRs.GetRows
        adoRs.Close
        Set adoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
            
            For iCnt = 0 To UBound(ArrayRecords, 2)
                If Not (VarType(ArrayRecords(0, iCnt)) = vbNull) Then
                    
                    .Row = Asc(ArrayRecords(0, iCnt)) - 64
                    
                    For iCol = 1 To .MaxCols
                        .Col = iCol
                         
                            If VarType(ArrayRecords(iCol, iCnt)) = vbNull Or ArrayRecords(iCol, iCnt) = 0 Then
                                .Text = ""
                            Else
                                .Text = Trim(ArrayRecords(iCol, iCnt))
                            End If
    
                    Next iCol

                End If
            Next iCnt
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Public Function Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim sMesg As String
    Dim sTemp As String
    Dim sPara As String
    
    Dim adoCmd As ADODB.Command

    Sp_Process = True
    
    With ss1
    
        'MaxRow = 0 is Exit Function Or iCount = 0
        If .MaxRows < 1 Then
            Sp_Process = False
            Exit Function
        End If
        
        Screen.MousePointer = vbHourglass
        .ReDraw = False
        
        'Db Connection Check
        If Conn Is Nothing Then
            If GF_DbConnect = False Then Sp_Process = False: Exit Function
        End If
        
        'Ado Setting
        Conn.CursorLocation = adUseServer
        Set adoCmd = New ADODB.Command
        
        Set adoCmd.ActiveConnection = Conn
        adoCmd.CommandType = adCmdStoredProc
        adoCmd.CommandText = Sc.Item("P-M")
        
        Conn.BeginTrans
        
        'Ceate Parameter (Input) iType + iColumn
        For iCount = 1 To 9
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        adoCmd.Parameters(0).Value = dtp_yy_mm.RawData                   'YEAR_MONTH
        adoCmd.Parameters(1).Value = cbo_plt.Text                        'PLT
        adoCmd.Parameters(2).Value = txt_prod_cd.Text                    'PROD_CD
        adoCmd.Parameters(3).Value = Trim(txt_cust_cd.Text)              'CUST_CD
        adoCmd.Parameters(4).Value = Trim(txt_stlgrd.Text)               'STLGRD
        adoCmd.Parameters(8).Value = sUserID                             'EMP_ID
        
        For iRow = 1 To .MaxRows
            
            .Row = iRow
            .Col = SpreadHeader
            adoCmd.Parameters(6).Value = .Text                           'THK_GRP
            
            'Parameters Setting
            For iCol = 1 To .MaxCols
            
                .Col = iCol
                If Trim(.Text) <> "" Then
                    adoCmd.Parameters(7).Value = Val(.Text)
                    
                    .Row = SpreadHeader
                    adoCmd.Parameters(5).Value = .Text                   'WID_GRP
                    
                    adoCmd.Execute
                    
                    'Error Check
                    If adoCmd("Error") <> "0" Then
               
                        ret_Result_ErrCode = adoCmd("Error")
                        ret_Result_ErrMsg = adoCmd("Messg")
                        sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
               
                        Call Gp_MsgBoxDisplay(sErrMessg)
                        Screen.MousePointer = vbDefault
                        Set adoCmd = Nothing
                        Conn.RollbackTrans
                        Sp_Process = False
                        Exit Function
               
                     End If
                
                End If
                
                .Row = iRow
            
            Next iCol
            
        Next iRow
        
        Conn.CommitTrans
        .ReDraw = True
        MDIMain.StatusBar1.Panels(1) = "提示信息: 数据处理完成"
        Screen.MousePointer = vbDefault
 '       Call txt_aply_item_Change
        Exit Function
    
    End With

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("SpreadPro_Error : " & Error)

End Function

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
txt_ref_type.Text = "2"
End Sub

Private Sub SSCommand1_Click()

On Error GoTo Cmd1_Error

    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim sMesg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command

    If dtp_copy_from.RawData = "" Or dtp_copy_to.RawData = "" Or dtp_copy_to.RawData <= dtp_copy_from.RawData Then
        Call Gp_MsgBoxDisplay("必须输入正确的日期...")
        Exit Sub
    End If
    
    If Trim(cbo_plt.Text) = "" Or Trim(txt_prod_cd.Text) = "" Then
        Call Gp_MsgBoxDisplay("必须输入 工厂 和 产品 ...")
        Exit Sub
    End If
    
    
    Screen.MousePointer = vbHourglass
        
    'Db Connection Check
'    If GF_DbConnect = False Then
'       Exit Sub
'    End If
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = M_CN1
    adoCmd.CommandType = adCmdText
    
    'Ceate Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    sQuery = "{CALL AAA1200C.P_COPY('" + dtp_copy_from.RawData + "','" + dtp_copy_to.RawData + "','" + cbo_plt.Text + "','" + txt_prod_cd.Text + "','" + txt_cust_cd.Text + "','" + txt_stlgrd.Text + "', '" + sUserID + "',?,? )}"
    
    adoCmd.CommandText = sQuery
    adoCmd.Execute , , adExecuteNoRecords

     'Error Check
     If adoCmd("Error") <> "0" Then

         ret_Result_ErrCode = adoCmd("Error")
         ret_Result_ErrMsg = adoCmd("Messg")
         sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg

         Call Gp_MsgBoxDisplay(sErrMessg)
         Screen.MousePointer = vbDefault
         Set adoCmd = Nothing
         Exit Sub
     End If
        
     Screen.MousePointer = vbDefault
     Set adoCmd = Nothing
     Call Gp_MsgBoxDisplay("复制成功，请查询!", "I", "")
     Exit Sub
    
Cmd1_Error:

    Screen.MousePointer = vbDefault
    M_CN1.RollbackTrans
    Set adoCmd = Nothing
    Call Gp_MsgBoxDisplay("Cmd1_Error : " & Error)

End Sub

Private Sub txt_cust_cd_DblClick()
    Call txt_cust_cd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_cust_cd_KeyUp(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"

        DD.rControl.Add Item:=txt_cust_cd
        DD.rControl.Add Item:=txt_cust_name

        DD.nameType = "1"
        Call Gf_Customer_DD(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_cust_cd)) = txt_cust_cd.MaxLength Then
        txt_cust_name.Text = Gf_CustNameFind(M_CN1, Trim(txt_cust_cd.Text), 1)
    Else
        txt_cust_name.Text = ""
    End If

End Sub


Private Sub txt_stlgrd_DblClick()
    Call txt_stlgrd_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_stlgrd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        
        DD.rControl.Add Item:=txt_stlgrd
        DD.rControl.Add Item:=txt_stlgrd_des

        DD.nameType = "2"
        Call Gf_Stlgrd_DD_AC(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_stlgrd)) >= 10 Then
        txt_stlgrd_des.Text = Gf_StlgrdNameFind(M_CN1, Trim(txt_stlgrd.Text))
    Else
        txt_stlgrd_des.Text = ""
    End If

End Sub

Private Sub txt_prod_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "B0005"
        
        DD.rControl.Add Item:=txt_prod_cd
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

End Sub

'---------------------------------------------------------------------------------------
'   1.ID           : Gf_Stlgrd_DD_AC
'   2.Name         : Stlgrd Code Code Data Dictionary Make Query
'   3.Input  Value : Conn Connection, KeyCode Integer
'   4.Return Value : Boolean
'   5.Writer       : Kim Sung Ho
'   6.Create Date  : 2003. 06 .20
'   7.Modify Date  :
'   8.Comment      : Stlgrd Code Code Data Dictionary Make Query
'---------------------------------------------------------------------------------------
Public Function Gf_Stlgrd_DD_AC(Conn As ADODB.Connection, KeyCode As Integer) As Boolean
    
    Dim sOld_Code, sNew_Code  As String
    Dim sOld_Name, sNew_Name  As String
    
    Dim iCount As Integer
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If

    If DD.rControl.Count = 0 Or DD.rControl.Count > 2 Then
        Call Gp_MsgBoxDisplay("DataDic Condition Invaild.....", "I")
        DD.DataDicType = ""
        DD.DicRefType = ""
        DD.nameType = ""
        DD.sQuery = ""
        DD.sWitch = ""
        DD.sWhere = ""
        DD.sSelect = False
        DD.sKey = ""
        Set DD.rControl = Nothing
        Set DD.wControl = Nothing
        Set DD.sPname = Nothing
        Exit Function
    End If
    
    DD.DataDicType = "S"        'Stlgrd Code
    DD.DicRefType = "C"         'Active Form DataDic Call
    
    If DD.sWitch = "MS" Then
    
        DD.sQuery = "SELECT STLGRD ""钢种"", STEEL_GRD_DETAIL ""目标说明"" FROM  NISCO.QP_NISCO_CHMC "
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.rControl.Item(1).Text) & "%' AND NVL(STLGRD_FL,'N') <> 'H'  "
            
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(STEEL_GRD_DETAIL,'%')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  STLGRD  ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "SELECT STLGRD ""钢种"", STEEL_GRD_DETAIL ""目标说明"" FROM  NISCO.QP_NISCO_CHMC "
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.sPname.Text) & "%' AND NVL(STLGRD_FL,'N') <> 'H' "
            
        If DD.rControl.Count > 1 Then
            DD.sPname.Col = DD.rControl.Item(2)
            sOld_Name = DD.sPname.Text
            DD.sWhere = DD.sWhere + " AND NVL(STEEL_GRD_DETAIL,'%')   like '" & Trim(DD.sPname.Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  STLGRD  ASC "
   
    End If
    
    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
    
        If DD.sWitch = "SP" Then
            
            DD.sPname.Col = DD.rControl.Item(1)
            sNew_Code = DD.sPname.Text
            
            If DD.rControl.Count > 1 Then
                DD.sPname.Col = DD.rControl.Item(2)
                sNew_Name = DD.sPname.Text
            End If
            
            DD.sPname.TabStop = True
            DD.sPname.SetFocus
            DD.sPname.SetActiveCell DD.rControl.Item(1), DD.sPname.ActiveRow
            DD.sPname.Action = SS_ACTION_ACTIVE_CELL
            DD.sPname.EditMode = True
            DD.sPname.TabStop = False
            
            If DD.sSelect Then
                If sOld_Code <> sNew_Code Then Call Gp_Sp_UpdateMake(DD.sPname, False)
            End If
            
        End If
    
    End If
    
    DD.sWitch = ""
    DD.sSelect = False
    
    Set DD.sPname = Nothing
    Set DD.rControl = Nothing

End Function



Public Function Sp_Order_Refer() As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iRow As Integer
    Dim iCnt As Integer
    Dim sQuery As String
    Dim sEdate As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant
    
    Set adoRs = New ADODB.Recordset
    
    sQuery = "{ call AAA1200C.P_SREFER('" + dtp_yy_mm.RawData + "','" + cbo_plt.Text + "', '" + txt_prod_cd.Text + "', '" + txt_cust_cd.Text + "', '" + txt_stlgrd.Text + "','O') }"
    
    adoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If adoRs.BOF Or adoRs.EOF Then
    
        Sp_Order_Refer = False
        adoRs.Close
        Set adoRs = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
        
    End If
        
    ArrayRecords = adoRs.GetRows
    
    adoRs.Close
    Set adoRs = Nothing
    
    With ss1
        iCnt = 0
        For iCnt = 0 To UBound(ArrayRecords, 2)
            .Row = Asc(ArrayRecords(0, iCnt)) - 64

            For iCol = 1 To .MaxCols
                .Col = iCol
                
                .BlockMode = True
                .CellType = 13      'SS_CELL_TYPE_NUMBER
                .TypeNumberDecPlaces = 3
                .TypeNumberMax = 99999999999.999
                .TypeNumberMin = 0
                .TypeNumberShowSep = True
                .TypeNumberLeadingZero = TypeLeadingZeroNo
                .TypeHAlign = TypeHAlignRight
                .BlockMode = False
    
                If ArrayRecords(iCol, iCnt) = vbNull Or ArrayRecords(iCol, iCnt) = 0 Then
                    .Text = ""
                Else
                    .Text = CStr(ArrayRecords(iCol, iCnt))
                End If
            Next iCol
        Next iCnt
        .ReDraw = True
        .Refresh
        Screen.MousePointer = vbDefault
        
    End With
    
    Exit Function

SpreadDisplay_Error:
    
    Set adoRs = Nothing
    ss1.ReDraw = True
    Sp_Order_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function

Private Sub History_Del(ByVal sYearMonth As String, ByVal sPlt As String, ByVal sPROD_CD As String)
    Dim sQuery      As String
    Dim sMesg       As String
    
    On Error GoTo UPDATE_ERROR
    
    M_CN1.BeginTrans
    
    sQuery = "          DELETE  FROM  AP_SALES_PLAN                  " & vbCrLf
    sQuery = sQuery & "  WHERE  YEAR_MONTH   = '" & sYearMonth & "'  " & vbCrLf
    sQuery = sQuery & "    AND  PLT          = '" & sPlt & "'        " & vbCrLf
    sQuery = sQuery & "    AND  PROD_CD      = '" & sPROD_CD & "'    " & vbCrLf

    M_CN1.Execute sQuery
        
    M_CN1.CommitTrans
    
    Exit Sub

UPDATE_ERROR:

    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay(Err.Description & sQuery)
    
    M_CN1.RollbackTrans
End Sub

