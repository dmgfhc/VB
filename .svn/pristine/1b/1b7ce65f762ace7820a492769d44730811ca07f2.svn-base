VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AAA1020C 
   Caption         =   "技术参数录入_AAA1020C"
   ClientHeight    =   9195
   ClientLeft      =   165
   ClientTop       =   1635
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   15225
   WindowState     =   2  'Maximized
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7965
      Left            =   90
      TabIndex        =   16
      Top             =   1170
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14049
      _Version        =   196609
      PaneTree        =   "AAA1020C.frx":0000
      Begin FPSpread.vaSpread ss2 
         Height          =   4590
         Left            =   30
         TabIndex        =   18
         Top             =   3345
         Width           =   15105
         _Version        =   393216
         _ExtentX        =   26644
         _ExtentY        =   8096
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
         MaxCols         =   16
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AAA1020C.frx":0052
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   3225
         Left            =   30
         TabIndex        =   17
         Top             =   30
         Width           =   15105
         _Version        =   393216
         _ExtentX        =   26644
         _ExtentY        =   5689
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
         SpreadDesigner  =   "AAA1020C.frx":07B0
      End
   End
   Begin VB.TextBox txt_excel 
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Text            =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
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
      ItemData        =   "AAA1020C.frx":09C9
      Left            =   9120
      List            =   "AAA1020C.frx":09D6
      TabIndex        =   13
      Tag             =   "产品代码"
      Top             =   450
      Width           =   870
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   330
      Left            =   10140
      TabIndex        =   11
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      Caption         =   "详细查询"
   End
   Begin InDate.UDate dtp_copy_to 
      Height          =   330
      Left            =   13035
      TabIndex        =   10
      Top             =   630
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      Height          =   330
      Left            =   11940
      Top             =   630
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
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
      Height          =   330
      Left            =   13035
      TabIndex        =   9
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      Height          =   330
      Left            =   11940
      Top             =   180
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
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
      Height          =   750
      Left            =   14265
      TabIndex        =   8
      Top             =   180
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1323
      _Version        =   196609
      Font3D          =   1
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
      Caption         =   "复制"
   End
   Begin VB.ComboBox cbo_line 
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
      Left            =   3795
      TabIndex        =   7
      Tag             =   "PRC_LINE"
      Top             =   450
      Width           =   600
   End
   Begin VB.ComboBox cbo_prc 
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
      Left            =   1260
      TabIndex        =   6
      Tag             =   "工序"
      Top             =   450
      Width           =   825
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
      ItemData        =   "AAA1020C.frx":09E6
      Left            =   9120
      List            =   "AAA1020C.frx":09E8
      TabIndex        =   5
      Tag             =   "工厂"
      Top             =   90
      Width           =   870
   End
   Begin VB.TextBox txt_aply_item_name 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4425
      TabIndex        =   2
      Top             =   90
      Width           =   3300
   End
   Begin VB.TextBox txt_aply_item 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3795
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "项目"
      Top             =   90
      Width           =   600
   End
   Begin InDate.ULabel ULabel8 
      Height          =   300
      Left            =   2625
      Top             =   90
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      Caption         =   "项目"
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
      Height          =   300
      Left            =   2625
      TabIndex        =   4
      Top             =   810
      Width           =   5235
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
      Height          =   300
      Left            =   1260
      MaxLength       =   11
      TabIndex        =   3
      Tag             =   "钢种"
      Top             =   810
      Width           =   1365
   End
   Begin InDate.ULabel ULabel6 
      Height          =   285
      Left            =   90
      Top             =   810
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
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
      Height          =   300
      Left            =   7950
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
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
   Begin InDate.ULabel ULabel4 
      Height          =   300
      Left            =   90
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      Caption         =   "工序"
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
      Height          =   300
      Left            =   7950
      Top             =   90
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
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
   Begin InDate.UDate dtp_yy_mm 
      Height          =   300
      Left            =   1260
      TabIndex        =   0
      Tag             =   "日期"
      Top             =   90
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   529
      Text            =   "____-__"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
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
      Height          =   300
      Left            =   90
      Top             =   90
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
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
   End
   Begin InDate.ULabel ULabel2 
      Height          =   300
      Left            =   2625
      Tag             =   "机号"
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      Caption         =   "机号"
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
   Begin Threed.SSCommand SSCommand3 
      Height          =   330
      Left            =   15225
      TabIndex        =   12
      Top             =   660
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   582
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "实绩计算"
   End
   Begin Threed.SSCommand SCmd2 
      Height          =   330
      Left            =   10140
      TabIndex        =   14
      Top             =   630
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
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
   Begin VB.Shape Shape1 
      Height          =   1050
      Left            =   11760
      Top             =   45
      Width           =   3450
   End
End
Attribute VB_Name = "AAA1020C"
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
'-- Program ID        AAA1020C
'-- Document No       Q-00-0010(Specification)
'-- Designer
'-- Coder
'-- Date              2003.7.9
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
         Call Gp_Ms_Collection(txt_aply_item, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_aply_item_name, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(cbo_plt, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(cbo_prc, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(cbo_line, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_prod_cd, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_stlgrd, "p", "n", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_stlgrd_des, " ", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
                     
     Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
                                
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Spread_Collection
    Sc1.Add Item:="AAA1020C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=ss1, Key:="Spread"
    
    Sc2.Add Item:=pColumn2, Key:="pColumn"
    Sc2.Add Item:=nColumn2, Key:="nColumn"
    Sc2.Add Item:=aColumn2, Key:="aColumn"
    Sc2.Add Item:=mColumn2, Key:="mColumn"
    Sc2.Add Item:=iColumn2, Key:="iColumn"
    Sc2.Add Item:=lColumn2, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=ss2.MaxCols, Key:="Last"
    Sc2.Add Item:=ss2, Key:="Spread"
    Sc2.Add Item:="AAA1020C.P_UPLOAD", Key:="P-M"
       
    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
    
    sQuery = "SELECT DISTINCT SUBSTR(CD,1,2) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' "
    Call Gf_ComboAdd(M_CN1, cbo_plt, sQuery)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cbo_plt_Change()
 
    Dim sQuery As String
    
    sQuery = "SELECT DISTINCT SUBSTR(CD,3,2) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND SUBSTR(CD,3,1) = '" + Mid(cbo_plt.Text, 1, 1) + "'"
    Call Gf_ComboAdd(M_CN1, cbo_prc, sQuery, True)
 
End Sub

Private Sub cbo_plt_Click()

    Dim sQuery As String
    
    If Trim(cbo_plt.Text) = "**" Then
       sQuery = "SELECT DISTINCT SUBSTR(CD,3,2) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND (SUBSTR(CD,3,1) = 'B' OR SUBSTR(CD,3,1) = '*') "
       cbo_line.Clear
       cbo_line.Text = "*"
    Else
       sQuery = "SELECT DISTINCT SUBSTR(CD,3,2) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND SUBSTR(CD,3,1) = '" + Mid(cbo_plt.Text, 1, 1) + "' "
    End If
    
    Call Gf_ComboAdd(M_CN1, cbo_prc, sQuery, True)

End Sub

Private Sub cbo_prc_Change()

    Dim sQuery As String
      
    sQuery = "SELECT DISTINCT SUBSTR(CD,5,1) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND SUBSTR(CD,3,2) = SUBSTR('" + cbo_prc.Text + "', 1, 2) "
    Call Gf_ComboAdd(M_CN1, cbo_line, sQuery, True)
 
End Sub

Private Sub cbo_prc_Click()

   Dim sQuery As String
   
   If Trim(cbo_plt.Text) = "**" Then
      cbo_line.Clear
      cbo_line.Text = "*"
   Else
        sQuery = "SELECT DISTINCT SUBSTR(CD,5,1) FROM ZP_CD WHERE CD_MANA_NO = 'A0002' AND SUBSTR(CD,3,2) = SUBSTR('" + cbo_prc.Text + "', 1, 2) "
        Call Gf_ComboAdd(M_CN1, cbo_line, sQuery, True)
       
        If Trim(cbo_prc.Text) = "" Then
        Else
            Select Case Trim(cbo_prc.Text)
                Case "BA", "BB", "BC", "BD", "BE"
                     txt_prod_cd.Text = "**"
                     txt_prod_cd.Enabled = False
                     txt_stlgrd.Text = "***********"
                     txt_stlgrd_des.Text = ""
                     txt_stlgrd.Enabled = False
        
                Case "BF"
                     txt_prod_cd.Text = "SL"
                     txt_prod_cd.Enabled = False
                     txt_stlgrd.Text = "***********"
                     txt_stlgrd.Enabled = False
                     
                Case "CA", "CB"
                     txt_prod_cd.Text = ""
                     txt_prod_cd.Enabled = True
                     txt_stlgrd.Text = "***********"
                     txt_stlgrd.Enabled = True
                  
            End Select
           
        End If
   End If
 
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

'    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    Call Gp_Sp_Setting(Sc1.Item("Spread"))
    Call Gp_Sp_Setting(Sc2.Item("Spread"))
    
    Call Sp_Setting
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(ss2, "A-System.INI", Me.Name)

    Screen.MousePointer = vbDefault
    
    If Mid(sAuthority, 1, 3) = "111" Then
       SSCommand2.Enabled = True
       SCmd2.Enabled = True
       SSCommand1.Enabled = True
    ElseIf Mid(sAuthority, 1, 1) = "1" Then
       SSCommand2.Enabled = True
       SCmd2.Enabled = False
       SSCommand1.Enabled = False
    Else
       SSCommand2.Enabled = False
       SCmd2.Enabled = False
       SSCommand1.Enabled = False
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

End Sub

Public Sub Form_Ref()

    Dim sMesg As String
    Dim iCol As Integer
    
    sMesg = Gf_Ms_NeceCheck(nControl)
    If sMesg = "OK" Then
        
        If Sp_Header_Refer() Then
            If Sp_Data_Refer() Then
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call Menu_Setting
                Call Gp_Ms_ControlLock(Mc1!lControl, True)
            End If
        End If
        
        If Left(dtp_yy_mm.RawData, 6) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Then
           Call Gp_Sp_BlockLock(ss1, 1, -1, 1, -1, True)
        Else
           
            If Trim(txt_aply_item.Text) = "001" Or Trim(txt_aply_item.Text) = "002" Or Trim(txt_aply_item.Text) = "003" Then
                         
                    With ss1
                         For iCol = 1 To .MaxCols - 1 Step 2
                             Call Gp_Sp_BlockLock(ss1, iCol, iCol, 1, .MaxRows, True)
                         Next iCol
                    
                         For iCol = 2 To .MaxCols Step 2
                             .Col = iCol
                             .Row = 1
                             .Col2 = iCol
                             .Row2 = .MaxRows
                             .BlockMode = True
                             .Lock = False
                             .BackColor = &HC0FFFF
                             .BlockMode = False
                             .Protect = True
                         Next iCol
                     End With
                     
                 Else
                 
                    With ss1
                         For iCol = 2 To .MaxCols Step 2
                             .Col = iCol
                             .Row = 1
                             .Col2 = iCol
                             .Row2 = .MaxRows
                             .BlockMode = True
                             .Lock = True
                             .BlockMode = False
                             .Protect = True
                         Next iCol
                     End With
                     
                 End If

        End If
        
        Call Gp_Ms_ControlLock(Mc1!lControl, True)
        
    Else
        sMesg = sMesg + " 必须输入...."
        Call Gp_MsgBoxDisplay(sMesg)
    End If
    
    txt_stlgrd.Enabled = True
    
   '    Call Gp_Ms_ControlLock(Mc1!lControl, True)

End Sub

Public Sub Form_Pro()

    Dim sMesg As String
    Dim sInput As String

    ss2.Col = 0
    ss2.Row = 1
    sInput = ss2.Text

If sInput = "Input" Then '导入
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc2"), Mc1, True) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
        Call Gp_Sp_BlockLock(ss2, 14, 14, 1, ss2.MaxRows, True)
        ss2.OperationMode = OperationModeNormal
    End If
Else '保存上表数据
    'BEFORE CURRENT DAY PROCESS IMPOSSIBLE
    If Left(dtp_yy_mm.RawData, 6) < Gf_CodeFind(M_CN1, "SELECT TO_CHAR(SYSDATE,'YYYYMM') FROM DUAL") Then
        sMesg = " 只能录入当月数据！"
        Call Gp_MsgBoxDisplay(sMesg)
        Exit Sub
    End If
    
    If Sp_Process(M_CN1, Proc_Sc("Sc")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Menu_Setting
    End If
End If

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
If txt_excel.Text = "1" Then
   Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
ElseIf txt_excel.Text = "2" Then
   Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
End If
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub SCmd2_Click()
   Load frm_Excel
   frm_Excel.txt_load_file.Text = "AAA1020C"
   frm_Excel.Show 1
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
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
    txt_excel.Text = "1"
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
 '      Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
'        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
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
        MDIMain.Mnu_Sorting.Visible = False
        MDIMain.Line1.Visible = False
        
        PopupMenu MDIMain.PopUp_Spread
        
        MDIMain.Mnu_Sorting.Visible = True
        MDIMain.Line1.Visible = True
    End If

End Sub

Public Sub Sp_Setting()
 
    With ss1

        .ColHeaderRows = 3
        .RowHeaderCols = 2
        .Col = -1
        .Row = SpreadHeader + 1
        .FontBold = True
        
        .RowHeight(SpreadHeader) = 15
        .RowHeight(SpreadHeader + 1) = 15
        
        .Row = SpreadHeader + 2
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = 0
        
        .BlockMode = True
        .RowMerge = MergeAlways
        .ColMerge = MergeAlways
        .BlockMode = False
        
        .Row = SpreadHeader
        .Col = SpreadHeader
        .Text = "宽度组\厚度组"
        .Row = SpreadHeader + 1
        .Col = SpreadHeader
        .Text = "宽度组\厚度组"
        
        .Row = SpreadHeader + 2
        .RowHidden = True
        
        .Col = SpreadHeader + 1
        .ColHidden = True
        
    End With
    

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
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) * 2
            For iCol = 0 To .MaxCols - 1 Step 2
            
               .Col = iCol + 1
               .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                  
                .Col = iCol + 2
                .Row = SpreadHeader
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, iCnt)) & " ~ " & Trim(ArrayRecords(2, iCnt)) & "mm"
                End If
                           
                .Col = iCol + 1:  .Row = SpreadHeader + 1:  .Text = "实绩"
                .Col = iCol + 2:  .Row = SpreadHeader + 1:  .Text = "计划"
                
                .Col = iCol + 1
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
                End If
                
                .Col = iCol + 2
                .Row = SpreadHeader + 2
                
                If VarType(ArrayRecords(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCnt))
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
                
                .ColWidth(iCol + 1) = 12
                
                .Col = iCol + 2: .Col2 = iCol + 2
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
                
                .ColWidth(iCol + 2) = 12
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
     '   .ReDraw = False
     '   .MaxRows = 0:  .MaxCols = 0
         .ColWidth(0) = 15
      '  .ColWidth(1) = 20
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
        
            .MaxRows = (UBound(ArrayRecords2, 2) + 1)
            iCnt = 0
            
            For iRow = 1 To .MaxRows
            
                .Row = iRow
                .Col = SpreadHeader
                
                If VarType(ArrayRecords2(0, iCnt)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords2(1, iCnt)) & " ~ " & Trim(ArrayRecords2(2, iCnt)) & "mm"
                End If
                
                .Col = SpreadHeader + 1
                .Text = Trim(ArrayRecords2(0, iCnt))
                
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
    
    With ss1

        For iCol = 1 To .MaxCols Step 2
            .Col = iCol
            .Row = 1
            .Col2 = iCol
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = True
            .BlockMode = False
            .Protect = True
        Next iCol

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
    Dim sTdate As String
    Dim sQuery As String
    Dim sEdate As String
    Dim sWID_GRP As String
    Dim sTHK_GRP As String
   ' Dim SPARA As String
    Dim adoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set adoRs = New ADODB.Recordset
    
    '3 Month After
    sEdate = Mid(dtp_yy_mm.Text, 1, 4) + Mid(dtp_yy_mm.Text, 6, 2)
  
    sQuery = "SELECT WID_GRP, THK_GRP, RST_VALUE,PLAN_VALUE"
    sQuery = sQuery + "   FROM AP_PROD_PLAN "
    sQuery = sQuery + "  WHERE YEAR_MONTH  = '" + sEdate + "' "
    sQuery = sQuery + "    AND APLY_ITEM   = '" + Trim(txt_aply_item.Text) + "' "
    sQuery = sQuery + "    AND PLT         = '" + Trim(cbo_plt.Text) + "' "
    sQuery = sQuery + "    AND PRC         = '" + Trim(cbo_prc.Text) + "' "
    sQuery = sQuery + "    AND PRC_LINE    = '" + Trim(cbo_line.Text) + "' "
    sQuery = sQuery + "    AND PROD_CD     = '" + Trim(txt_prod_cd.Text) + "' "
    sQuery = sQuery + "    AND STLGRD      = '" + Trim(txt_stlgrd.Text) + "' "
    
    If txt_prod_cd.Text <> "**" Then
       sQuery = sQuery + "    AND wid_grp <> '*' and thk_grp<> '*'"
    End If
    sQuery = sQuery + "  ORDER BY WID_GRP, THK_GRP "
    
    With ss1

        Sp_Data_Refer = True
        .ReDraw = False
       ' .MaxRows = 0
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        adoRs.Open sQuery, M_CN1, adOpenKeyset
        
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
            iRow = 1
            For iCnt = 0 To UBound(ArrayRecords, 2)
                .Row = iRow
                .Col = SpreadHeader + 1
                 sWID_GRP = .Text
                 Do While iRow <= .MaxRows And sWID_GRP <> Trim(ArrayRecords(0, iCnt))
                    iRow = iRow + 1
                    .Row = iRow
                    sWID_GRP = .Text
                 Loop
                           
                 For iCol = 1 To .MaxCols - 1 Step 2
                    .Col = iCol
                    .Row = SpreadHeader + 2
                    sTHK_GRP = .Text

                    If sTHK_GRP = ArrayRecords(1, iCnt) Then
                        .Row = iRow
                     
                        If VarType(ArrayRecords(2, iCnt)) = vbNull Or ArrayRecords(2, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(2, iCnt))
                        End If
                        .Col = iCol + 1
                        If VarType(ArrayRecords(3, iCnt)) = vbNull Or ArrayRecords(3, iCnt) = 0 Then
                            .Text = ""
                        Else
                            .Text = Trim(ArrayRecords(3, iCnt))
                        End If
                
                    End If

                Next iCol
                
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
        For iCount = 1 To 11
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        For iRow = 1 To .MaxRows
            
            .Row = iRow
            
            'Parameters Setting
            For iCol = 2 To .MaxCols Step 2
            
                .Col = iCol
                If Trim(.Text) <> "" Then
                
                    .Row = SpreadHeader + 2
                    .Col = iCol
                    adoCmd.Parameters(6).Value = .Text     'thk_grp
              
                    .Row = iRow
                    .Col = SpreadHeader + 1
                    adoCmd.Parameters(7).Value = .Text     'wid_grp
                    
                    .Col = iCol
                    If Trim(.Text) = "" Then                'plan_value
                        adoCmd.Parameters(9).Value = 0
                    Else
                        dTempInt = .Text
                        adoCmd.Parameters(9).Value = dTempInt
                    End If
                    
                    adoCmd.Parameters(10).Value = sUserID                            'User-id
                    
                    adoCmd.Parameters(0).Value = Mid(dtp_yy_mm.Text, 1, 4) + _
                                                 Mid(dtp_yy_mm.Text, 6, 2)           'YEAR_MONTH
                                                 
                    adoCmd.Parameters(1).Value = cbo_plt.Text                        'PLT
                    adoCmd.Parameters(2).Value = cbo_prc.Text                        'PRC
                    adoCmd.Parameters(3).Value = cbo_line.Text                       'PRC_LINE
                    adoCmd.Parameters(4).Value = txt_prod_cd.Text                    'PROD_CD
                    adoCmd.Parameters(5).Value = txt_stlgrd.Text                     'STLGRD
                    adoCmd.Parameters(8).Value = txt_aply_item.Text                  'APLY_ITEM
                                   
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
'
'Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)
'
'    Dim sQuery As String
'    Dim sYearMonth As String
'    Dim sItem As String
'    Dim sPLT As String
'    Dim sPrc As String
'    Dim sLine As String
'    Dim sProd As String
'    Dim sStlgrd As String
'    Dim sThk As String
'    Dim sWid As String
'
'    If ss2.MaxRows < 1 Then
'       Exit Sub
'    End If
'
'    ss2.Row = Row
'    ss2.Col = 1
'    sYearMonth = ss2.Text
'    ss2.Col = 2
'    sItem = ss2.Text
'    ss2.Col = 3
'    sPLT = ss2.Text
'    ss2.Col = 4
'    sPrc = ss2.Text
'    ss2.Col = 5
'    sLine = ss2.Text
'    ss2.Col = 6
'    sProd = ss2.Text
'    ss2.Col = 7
'    sStlgrd = ss2.Text
'    ss2.Col = 8
'    sThk = ss2.Text
'    ss2.Col = 9
'    sWid = ss2.Text
'
'    sQuery = "SELECT SUBSTR(YEAR_MONTH,1,4)||'-'||SUBSTR(YEAR_MONTH,5,2) AS YY_MM, RST_VALUE FROM AP_PROD_PLAN "
'    sQuery = sQuery + " WHERE SUBSTR(YEAR_MONTH,1,4) = '" + Mid(sYearMonth, 1, 4) + "' "
'    sQuery = sQuery + "   AND APLY_ITEM = '" + sItem + "'"
'    sQuery = sQuery + "   AND PLT       = '" + sPLT + "'"
'    sQuery = sQuery + "   AND PRC       = '" + sPrc + "'"
'    sQuery = sQuery + "   AND PRC_LINE  = '" + sLine + "'"
'    sQuery = sQuery + "   AND PROD_CD   = '" + sProd + "'"
'    sQuery = sQuery + "   AND STLGRD    = '" + sStlgrd + "'"
'    sQuery = sQuery + "   AND THK_GRP   = '" + sThk + "'"
'    sQuery = sQuery + "   AND WID_GRP   = '" + sWid + "'"
'    sQuery = sQuery + " ORDER BY YY_MM "
'
'  '  Debug.Print sQuery
'    If Gf_Only_Display(M_CN1, sc3, sQuery) Then
'
'    End If
'
'End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    txt_excel.Text = "2"
End Sub

Private Sub SSCommand1_Click()

'   Dim sQuery As String
'   If dtp_copy_from.RawData <> "" And dtp_copy_to.RawData <> "" And dtp_copy_to.RawData > dtp_copy_from.RawData Then
'      sQuery = "{call AAA1020C.P_MODIFY1('" + dtp_copy_from.RawData + "','" + dtp_copy_to.RawData + "','" + sUserID + "',?,?)}"
'
'   End If
On Error GoTo Cmd1_Error

    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    Dim sMesg As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command

    If dtp_copy_from.RawData = "" Or dtp_copy_to.RawData = "" Or dtp_copy_to.RawData <= dtp_copy_from.RawData Then
        Call Gp_MsgBoxDisplay("必须输入正确的日期...")
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
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = "AAA1020C.P_MODIFY2"
    
    M_CN1.BeginTrans
    
    'Ceate Parameter (Input) iType + iColumn
    For iCount = 1 To 8
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Ceate Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    adoCmd.Parameters(0).Value = dtp_copy_from.RawData
    adoCmd.Parameters(1).Value = dtp_copy_to.RawData
    
    adoCmd.Parameters(2).Value = cbo_plt.Text
    adoCmd.Parameters(3).Value = cbo_prc.Text
    adoCmd.Parameters(4).Value = cbo_line.Text
    adoCmd.Parameters(5).Value = txt_prod_cd.Text
    adoCmd.Parameters(6).Value = txt_aply_item.Text
    
    adoCmd.Parameters(7).Value = sUserID                            'User-id
    adoCmd.Execute
     
     'Error Check
     If adoCmd("Error") <> "0" Then

         ret_Result_ErrCode = adoCmd("Error")
         ret_Result_ErrMsg = adoCmd("Messg")
         sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg

         Call Gp_MsgBoxDisplay(sErrMessg)
         Screen.MousePointer = vbDefault
         Set adoCmd = Nothing
         M_CN1.RollbackTrans
         Exit Sub

      End If
        
      M_CN1.CommitTrans
      Screen.MousePointer = vbDefault
      Exit Sub
    
Cmd1_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    M_CN1.RollbackTrans
    Call Gp_MsgBoxDisplay("Cmd1_Error : " & Error)

End Sub

Private Sub SSCommand2_Click()

    Dim sQuery As String

     If dtp_yy_mm.RawData <> "" And Trim(txt_aply_item.Text) <> "" And Trim(cbo_plt.Text) <> "" Then
        sQuery = "select year_month,aply_item,gf_comnnamefind('A0001',aply_item),plt,prc,prc_line,prod_cd, "
        sQuery = sQuery + " stlgrd,GF_STLGRD_DETAIL(stlgrd),thk_grp,GF_PLAN_SIZE_RANGE(PROD_CD,THK_GRP,'T'),wid_grp,GF_PLAN_SIZE_RANGE(PROD_CD,WID_GRP,'W'),plan_value,NVL(UPD_DATE,INS_DATE),nvl(upd_emp,ins_emp) "
        sQuery = sQuery + "  FROM AP_PROD_PLAN"
        sQuery = sQuery + " WHERE YEAR_MONTH = '" + Mid(dtp_yy_mm.RawData, 1, 6) + "'"
        sQuery = sQuery + "   AND APLY_ITEM  = '" + Trim(txt_aply_item.Text) + "'"
        sQuery = sQuery + "   AND PLT        LIKE '" + Trim(cbo_plt.Text) + "%'"
        sQuery = sQuery + "   AND PRC        LIKE '" + Trim(cbo_prc.Text) + "%'"
        sQuery = sQuery + "   AND PRC_LINE   LIKE '" + Trim(cbo_line.Text) + "%'"
        sQuery = sQuery + "   AND PROD_CD    LIKE '" + Trim(txt_prod_cd.Text) + "%'"
        sQuery = sQuery + "   AND STLGRD     LIKE '" + Trim(txt_stlgrd.Text) + "%'"
        sQuery = sQuery + "   ORDER BY year_month,APLY_ITEM,PLT,PRC,PRC_LINE,PROD_CD,STLGRD,THK_GRP,WID_GRP "
        
        If Gf_Only_Display(M_CN1, Sc2, sQuery) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           ss2.OperationMode = OperationModeNormal
        End If
     Else
        MsgBox "年月，项目，工厂必须输入!", vbCritical, "系统提示信息"
        Exit Sub
     End If

End Sub

Private Sub SSCommand3_Click()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iCount As Integer
    
    Dim adoCmd As ADODB.Command
    
    Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    'Procedure Name Not Define
    sQuery = "{call ACG1000P ('" + Left(dtp_yy_mm.RawData, 6) + "',?)}"
    
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
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Call Gp_MsgBoxDisplay("实绩编成完了..!!", "I")
        Call Form_Ref
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error)
    
End Sub

Private Sub txt_aply_item_DblClick()
    Call txt_aply_item_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_aply_item_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "A0001"
        
        DD.rControl.Add Item:=txt_aply_item
        DD.rControl.Add Item:=txt_aply_item_name

        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If
    
    If Len(Trim(txt_aply_item)) = txt_aply_item.MaxLength Then
        txt_aply_item_name.Text = Gf_ComnNameFind(M_CN1, "A0001", Trim(txt_aply_item.Text), 2)
    Else
        txt_aply_item_name.Text = ""
    End If

End Sub

Private Sub txt_prod_cd_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub txt_prod_cd_LostFocus()

    If Trim(txt_prod_cd.Text) = "" Then
    Else
        If Trim(cbo_prc.Text) = "CA" Or Trim(cbo_prc.Text) = "CB" Then
            If Trim(txt_prod_cd.Text) = "HC" Or Trim(txt_prod_cd.Text) = "PP" Then
            Else
                Call Gp_MsgBoxDisplay("必须输入正确的产品代码 ...")
                txt_prod_cd.Text = ""
                txt_prod_cd.SetFocus
            End If
        End If
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
    
    If Len(Trim(txt_stlgrd)) = txt_stlgrd.MaxLength Then
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
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.rControl.Item(1).Text) & "%' AND STLGRD_FL <> 'H'  "
            
        If DD.rControl.Count > 1 Then
            DD.sWhere = DD.sWhere + " AND NVL(STEEL_GRD_DETAIL,'%')   like '" & Trim(DD.rControl.Item(2).Text) & "%' "
        End If
        
        DD.sWhere = DD.sWhere + " ORDER  BY  STLGRD  ASC "
        
    Else
    
        DD.sPname.Col = DD.rControl.Item(1)
        sOld_Code = DD.sPname.Text
            
        DD.sQuery = "SELECT STLGRD ""钢种"", STEEL_GRD_DETAIL ""目标说明"" FROM  NISCO.QP_NISCO_CHMC "
        DD.sWhere = " WHERE STLGRD like '" & Trim(DD.sPname.Text) & "%' AND STLGRD_FL <> 'H' "
            
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



