VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AKH1040C 
   Caption         =   "非定型耐材检验实绩录入及查询界面_AKH1040C"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.ComboBox CBO_ALL_NO 
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
      ItemData        =   "AKH1040C.frx":0000
      Left            =   6030
      List            =   "AKH1040C.frx":0002
      TabIndex        =   16
      Tag             =   "ROLL_NO"
      Top             =   120
      Width           =   1845
   End
   Begin VB.TextBox txt_all_name 
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
      Left            =   2355
      MaxLength       =   50
      TabIndex        =   5
      Tag             =   "工厂"
      Top             =   120
      Width           =   1845
   End
   Begin VB.TextBox txt_all_cd 
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
      Left            =   1650
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "铁合金代码"
      Top             =   120
      Width           =   705
   End
   Begin VB.TextBox test_all_cd 
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
      Left            =   2490
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "铁合金代码"
      Top             =   4800
      Width           =   705
   End
   Begin VB.TextBox test_all_name 
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
      Left            =   3195
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "工厂"
      Top             =   4800
      Width           =   1725
   End
   Begin VB.ComboBox CBO_ALL_NO1 
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
      ItemData        =   "AKH1040C.frx":0004
      Left            =   6630
      List            =   "AKH1040C.frx":0006
      TabIndex        =   1
      Tag             =   "ROLL_NO"
      Top             =   4800
      Width           =   1845
   End
   Begin VB.TextBox txt_heat_no 
      Alignment       =   2  'Center
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
      Left            =   14280
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "炉号"
      Top             =   4800
      Width           =   1050
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   8160
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "上料日期"
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
   Begin InDate.UDate SDT_TO_DATE 
      Height          =   315
      Left            =   11250
      TabIndex        =   6
      Tag             =   "终止日期"
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
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
   Begin InDate.UDate SDT_FROM_DATE 
      Height          =   315
      Left            =   9495
      TabIndex        =   7
      Tag             =   "起始日期"
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
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
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      Caption         =   "非定型材料代码"
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
   Begin FPSpread.vaSpread ss2 
      Height          =   3735
      Left            =   120
      TabIndex        =   8
      Top             =   5160
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   6588
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
      MaxCols         =   10
      MaxRows         =   31
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AKH1040C.frx":0008
   End
   Begin Threed.SSCheck Chk_ss1 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   720
      _ExtentX        =   1270
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
      Caption         =   "验收"
      Value           =   1
   End
   Begin Threed.SSCheck Chk_ss2 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   503
      _Version        =   196609
      Font3D          =   2
      ForeColor       =   0
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
      Caption         =   "匹配"
      Value           =   1
   End
   Begin InDate.UDate TEST_FROM_DATE 
      Height          =   315
      Left            =   9870
      TabIndex        =   11
      Tag             =   "发生日期"
      Top             =   4800
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
   Begin InDate.UDate TEST_TO_DATE 
      Height          =   315
      Left            =   11640
      TabIndex        =   12
      Tag             =   "日期"
      Top             =   4800
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   8640
      Top             =   4800
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      Caption         =   "匹配日期"
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
      Left            =   960
      Top             =   4800
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      Caption         =   "非定型材料代码"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   5040
      Top             =   4800
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "非定型材料批号"
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
   Begin InDate.ULabel ULabel63 
      Height          =   315
      Left            =   13200
      Top             =   4800
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "炉号"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   3855
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   15180
      _Version        =   393216
      _ExtentX        =   26776
      _ExtentY        =   6800
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
      MaxCols         =   19
      MaxRows         =   31
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AKH1040C.frx":082C
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   4440
      Top             =   120
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      Caption         =   "非定型材料批号"
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   11010
      TabIndex        =   14
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   11400
      TabIndex        =   13
      Top             =   4800
      Width           =   255
   End
End
Attribute VB_Name = "AKH1040C"
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
'-- Program Name      卷筒使用实绩查询及修改界面
'-- Program ID        AGF3020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          ZHANG
'-- Coder             ZHANG
'-- Date              2009.10.10
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
Public sDateTime As String          'Active Form Time Setting
Public sQuery_load As String        'Active Form sQuery Setting
Public QueryYN      As Boolean

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


Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim pColumn2 As New Collection      'Spread Primary Key Collection
Dim nColumn2 As New Collection      'Spread necessary Column Collection
Dim mColumn2 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2 As New Collection      'Spread Insert Column Collection
Dim aColumn2 As New Collection      'Master -> Spread Column Collection
Dim lColumn2 As New Collection      'Spread Lock Column Collection


Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection

'Dim Sc3 As New Collection           'Spread Collection




  Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Sheet"


    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_all_cd, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_all_name, " ", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(CBO_ALL_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(SDT_FROM_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(SDT_TO_DATE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    
    Call Gp_Sp_Collection(ss1, 1, "P", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
 
   
  
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AKH1040C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AKH1040C.P_REFER", Key:="P-R"
'    sc1.Add Item:="AKH1020C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Call Gp_Sp_ColHidden(ss1, 4, True)

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    

   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
         Call Gp_Ms_Collection(test_all_cd, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(test_all_name, " ", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(CBO_ALL_NO1, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(TEST_FROM_DATE, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
        Call Gp_Ms_Collection(TEST_TO_DATE, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(txt_heat_no, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  

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
    Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
 
  
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AKH1040C.P_REFER1", Key:="P-R"
    sc2.Add Item:="AKH1040C.P_MODIFY1", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
  
    sc2.Item("Spread").Col = 0
    sc2.Item("Spread").Row = 0
    sc2.Item("Spread").Text = "◎"

    Call Gp_Sp_ColHidden(ss2, 5, True)
    Call Gp_Sp_ColHidden(ss2, 6, True)
    Call Gp_Sp_ColHidden(ss2, 7, True)
    Call Gp_Sp_ColHidden(ss2, 8, True)
     
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

    sAuthority = Gf_Pgm_Authority(Me.Name)


    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
     Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Gp_Sp_Setting(sc1("Spread"))
    Call Gf_Sp_Cls(sc1)
    Call Gp_Sp_ColGet(sc1("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_HdColColor(sc1.Item("Spread"), 1)
    
    
    Call Gp_Sp_Setting(sc2("Spread"))
    Call Gf_Sp_Cls(sc2)
    Call Gp_Sp_ColGet(sc2("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_HdColColor(sc2.Item("Spread"), 1)
   
   
    SDT_FROM_DATE.RawData = Mid(SDT_FROM_DATE.RawData, 1, 6) & "01"
 
    Screen.MousePointer = vbDefault
  
    Chk_ss1.ForeColor = &HFF&
    Chk_ss2.ForeColor = &H808080
 
  
    Chk_ss1.Value = ssCBChecked
    Chk_ss2.Value = ssCBUnchecked
  
   sQuery_load = "SELECT CHECKNO FROM FP_ALL_MAIN  ORDER BY CHECKNO  "
   Call Gf_ComboAdd(M_CN1, CBO_ALL_NO, sQuery_load)
    
   Call Gf_ComboAdd(M_CN1, CBO_ALL_NO1, sQuery_load)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)

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
    
 
    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Spread_Can()

   Dim sCid As String

     If Chk_ss1.Value = -1 Then
            
            Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
            
         
                
                ss1.Row = ss1.ActiveRow
                ss1.Col = 4
                sCid = ss1.Text
                If sCid <> "" Then
                   ss1.Col = 3
                   ss1.Text = sCid
                End If
   End If
   
     If Chk_ss2.Value = -1 Then
            
            Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC2"))
              
   End If
   
   
        
End Sub

Public Sub Form_Cls()
    
    
    If Chk_ss1.Value = -1 Then
         Call Gf_Sp_Cls(sc1)
         txt_all_cd.Text = ""
         txt_all_name.Text = ""
         CBO_ALL_NO.Text = ""
         SDT_FROM_DATE = ""
         SDT_TO_DATE = ""
         
    End If
        
       
    If Chk_ss2.Value = -1 Then
     
        Call Gf_Sp_Cls(sc2)
        test_all_cd.Text = ""
        test_all_name.Text = ""
        CBO_ALL_NO1.Text = ""
        TEST_FROM_DATE = ""
        TEST_TO_DATE = ""
        txt_heat_no = ""
         
    End If
 

End Sub

Public Sub Form_Ref()

  Dim i As Integer
  Dim sCid As String


 QueryYN = False
 
        If Chk_ss1.Value = -1 Then

    
              If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
         
         
              If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
                 Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
        
        
                            For i = 1 To ss1.MaxRows
                               ss1.Col = 4
                                ss1.Row = i
                               sCid = ss1.Text
                               If sCid <> "" Then
                               ss1.Col = 3
                               ss1.Text = sCid
        
                              End If
                            
                            Next i
         
              End If
              
      End If
      
      
      If Chk_ss2.Value = -1 Then

             If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
         
         
              If Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl")) Then
                 Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        
                 Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
        
                                
        
              End If
              
      End If
      
      
 Exit Sub

Refer_Err:

End Sub



Public Sub Form_Pro()
   Dim iCount As Integer
   Dim MsgBox As String

  
    If Chk_ss1.Value = -1 Then

        If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
             Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
             Call Form_Ref
        End If
        
    End If
   
   
   
    If Chk_ss2.Value = -1 Then

        If Gf_Sp_Process(M_CN1, sc2, Mc2) Then
             Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
             Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
             Call Form_Ref
        End If
        
    End If
   

End Sub


Public Sub Form_Ins()

    
    If Chk_ss1.Value = -1 Then
    
            
            Call Gp_Sp_Ins(Proc_Sc("Sc"))
            ss1.Col = 16
            ss1.Row = ss1.ActiveRow
            ss1.Text = sUserID
        '
        '    Call Gp_Sp_ColLock(ss1, 1, False)
            
            ss1.Row = ss1.ActiveRow
            ss1.Col = 3
            ss1.BackColor = &HC0FFFF
            
             
            Call Pf_ComboAdd(M_CN1, ss1, 3, "SELECT CHECKNO FROM FP_ALL_MAIN  ORDER BY CHECKNO")
            
            
   End If
            
            
    If Chk_ss2.Value = -1 Then

           
            Call Gp_Sp_Ins(sc2)
            ss2.Col = 8
            ss2.Row = ss2.ActiveRow
            ss2.Text = sUserID
               
            
   End If
            
            
  
End Sub
Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
    
End Sub
Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
    ss1.Col = 16
    ss1.Row = ss1.ActiveRow
    ss1.Text = sUserID

'    Call Gp_Sp_ColLock(ss1, 1, False)
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 3
    ss1.BackColor = &HC0FFFF
    
    Call Pf_ComboAdd(M_CN1, ss1, 3, "SELECT CHECKNO FROM FP_ALL_MAIN  ORDER BY CHECKNO")
    ss1.OperationMode = OperationModeNormal

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

 If Chk_ss1.Value = -1 Then
    
    Call Gp_Sp_Del(Proc_Sc("SC"))
    
    
 End If
 
 If Chk_ss2.Value = -1 Then
    
    Call Gp_Sp_Del(Proc_Sc("SC2"))
    
    
 End If



End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
Dim sCid As String
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub
Private Sub SDT_FROM_DATE_DblClick()
    If SDT_FROM_DATE.RawData = "" Then
     SDT_FROM_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_TO_DATE.RawData = "" Then
        SDT_TO_DATE.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub
Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code As String
    
    Dim i As Integer
      Dim sCid As String


   
              If ss1.ActiveCol = 1 Then
            
                        If ss1.MaxRows < 1 Then Exit Sub
                        
                        If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
                            Exit Sub
                        End If
                    
                        Select Case ss1.ActiveCol
                        
                            Case 1
                            
                                If KeyCode = vbKeyF4 Then
                                
                                    Set DD.sPname = Me.ss1
                                    
                                    DD.sWitch = "SP"
                                    DD.sKey = "F0049"
                                    DD.rControl.Add Item:=1
                                    DD.rControl.Add Item:=2
                                    
                                    DD.nameType = "2"
                                    Call Gf_Common_DD(M_CN1, KeyCode)
                                    
                                Else
                    
                                    ss1.Col = ss1.ActiveCol
                    
                                     If Len(Trim(ss1.Text)) = 4 Then
                                     
                                            sTemp_Code = ss1.Text
                                            ss1.Col = 2
                                            ss1.Text = Gf_ComnNameFind(M_CN1, "F0049", Trim(sTemp_Code), 2)
                                     Else
                                            ss1.Col = 2
                                            ss1.Text = ""
                                     End If
                                    
                                End If
                                
                    
                        End Select
                        
                End If
        
        
End Sub
Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim iRow  As Integer
Dim i, j, Scr_wgt, hm_wgt, Steel_wgt As Integer

   If Row <> 0 Then
   
        If Col = 5 Then
                ss1.Col = Col
                ss1.Row = Row
                If ss1.Lock = False Then
                   ss1.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
                   ss1.Col = 0
               
                
                    If ss1.Text <> "Input" And ss1.Text <> "Delete" Then
                       ss1.Text = "Update"
                    End If
                   
        
                End If
          
      End If

  End If
  
   If Row <> 0 Then
   
        If Col = 17 Then
                ss1.Col = Col
                ss1.Row = Row
                If ss1.Lock = False Then
                   ss1.Text = Format(Now, "YYYY-MM-DD HH:MM:SS")
                   ss1.Col = 0
               
                
                    If ss1.Text <> "Input" And ss1.Text <> "Delete" Then
                       ss1.Text = "Update"
                    End If
                   
        
                End If
          
      End If

  End If
  

  
End Sub
Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)

  Dim sCid As String
  Dim i As Integer
  
  If Col = 6 Then
           ss1.Col = 6
           sCid = ss1.Text
           ss1.Col = 7
           ss1.Text = sCid
  End If
        

  End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim sCid As String

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)

        If Col = 3 Then
           ss1.Col = 3
           sCid = ss1.Text
           ss1.Col = 4
           ss1.Text = sCid
        End If
        
        
       If Col = 6 Then
           ss1.Col = 6
           sCid = ss1.Text
           ss1.Col = 7
           ss1.Text = sCid
        End If
        

    End If
    

End Sub
Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss1
        PopupMenu MDIMain.PopUp_Spread
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
Public Function Pf_ComboAdd(Conn As ADODB.Connection, ss As vaSpread, Col As Integer, sQuery As String, Optional ClsChk As Boolean = True) As Boolean

On Error GoTo ComboAdd_Error

    Dim AdoRs As ADODB.Recordset
    Dim sList As String
    
    'Db Connection Check
    If Conn.State = 0 Then
        If GF_DbConnect = False Then Pf_ComboAdd = False: Exit Function
    End If
    
'    If ClsChk Then
'        Cbo.Clear
'    End If
    
    Set AdoRs = New ADODB.Recordset

    'Ado Execute
    AdoRs.Open sQuery, Conn, adOpenKeyset
    
    If Not AdoRs.BOF And Not AdoRs.EOF Then
        While Not AdoRs.EOF
            
            If AdoRs.Fields(0) <> vbNull Then
                sList = sList & AdoRs.Fields(0) & vbTab
                'Cbo.AddItem AdoRs.Fields(0)
            End If
            AdoRs.MoveNext
            
        Wend
        Pf_ComboAdd = True
    Else
        Pf_ComboAdd = False
    End If
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    ss.Col = Col
    ss.TypeComboBoxList = sList
    
    Exit Function

ComboAdd_Error:

    Set AdoRs = Nothing
    Pf_ComboAdd = False

End Function

Private Sub txt_all_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
           DD.sWitch = "MS"
           DD.sKey = "F0049"
    
           DD.rControl.Add Item:=txt_all_cd
           DD.rControl.Add Item:=txt_all_name
           
           DD.nameType = "2"
           Call Gf_Common_DD(M_CN1, KeyCode)
    
    Else
    
        If Len(Trim(txt_all_cd.Text)) = txt_all_cd.MaxLength Then
            txt_all_name.Text = Gf_ComnNameFind(M_CN1, "F0049", txt_all_cd.Text, 2)
        Else
            txt_all_name.Text = ""
        End If
    End If
    
End Sub
Private Sub test_all_cd_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
           DD.sWitch = "MS"
           DD.sKey = "F0049"
    
           DD.rControl.Add Item:=test_all_cd
           DD.rControl.Add Item:=test_all_name
           
           DD.nameType = "2"
           Call Gf_Common_DD(M_CN1, KeyCode)
    
    Else
    
        If Len(Trim(test_all_cd.Text)) = test_all_cd.MaxLength Then
            test_all_name.Text = Gf_ComnNameFind(M_CN1, "F0049", test_all_cd.Text, 2)
        Else
            test_all_name.Text = ""
        End If
    End If
    
End Sub

Private Sub txt_all_cd_DblClick()

    Call txt_all_cd_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Chk_ss1_Click(Value As Integer)

   If Chk_ss1.Value = ssCBUnchecked Then
       If Chk_ss2.Value = ssCBUnchecked Then
            Chk_ss1.Value = ssCBChecked
       End If
       Exit Sub
    End If
   
    Chk_ss1.ForeColor = &HFF&
    Chk_ss2.ForeColor = &H808080
    Chk_ss2.Value = ssCBUnchecked
    
    Call Gf_Sp_Cls(sc2)
              test_all_cd.Text = ""
              test_all_name.Text = ""
              CBO_ALL_NO1.Text = ""
              TEST_FROM_DATE = ""
              TEST_TO_DATE = ""
              txt_heat_no = ""
   
   If SDT_FROM_DATE.RawData = "" Then
       SDT_FROM_DATE.Text = Format(Now, "YYYY-MM") + "-01"
   End If
   If SDT_TO_DATE.RawData = "" Then
       SDT_TO_DATE.Text = Format(Now, "YYYY-MM-DD")
   End If

   
End Sub
Private Sub Chk_ss2_Click(Value As Integer)

    If Chk_ss2.Value = ssCBUnchecked Then
       If Chk_ss1.Value = ssCBUnchecked Then
            Chk_ss2.Value = ssCBChecked
       End If
       Exit Sub
    End If
   
    Chk_ss2.ForeColor = &HFF&
    Chk_ss1.ForeColor = &H808080
    Chk_ss1.Value = ssCBUnchecked
    
    Call Gf_Sp_Cls(sc1)
                txt_all_cd.Text = ""
                txt_all_name.Text = ""
                CBO_ALL_NO.Text = ""
                SDT_FROM_DATE = ""
                SDT_TO_DATE = ""
    
   If TEST_FROM_DATE.RawData = "" Then
       TEST_FROM_DATE.Text = Format(Now, "YYYY-MM-DD")
   End If
   If TEST_TO_DATE.RawData = "" Then
       TEST_TO_DATE.Text = Format(Now, "YYYY-MM-DD")
   End If
  
End Sub
Private Sub ss2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code As String

    If ss2.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss2.ActiveCol
    
        Case 1
           
            ss2.Row = ss2.ActiveRow
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss2
                
                DD.sWitch = "SP"
                DD.sKey = "F0049"
                DD.rControl.Add Item:=1
                DD.rControl.Add Item:=2
                
                DD.nameType = "2"
                Call Gf_Common_DD(M_CN1, KeyCode)
                
                ss2.Col = 1
                If ss2.Text = "B" Then
                    ss2.Text = ""
                    ss2.Col = 2
                    ss2.Text = ""
                    Exit Sub
                End If
                
            Else
            
                ss2.Col = ss2.ActiveCol
                
                 If Len(Trim(ss2.Text)) = 4 Then

                   sTemp_Code = ss2.Text
                   ss2.Col = 2
                   ss2.Text = Gf_ComnNameFind(M_CN1, "F0049", Trim(sTemp_Code), 2)
                Else
                       ss2.Col = 2
                       ss2.Text = ""
                End If
            
            End If
            
    End Select
    
End Sub



Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim iRow  As Integer
Dim i, j, Scr_wgt, hm_wgt, Steel_wgt As Integer

  If Row <> 0 Then
     
      If Col = 9 Then
                ss2.Col = Col
                ss2.Row = Row
       
             If ss2.Lock = False Then
                 
                   ss2.Col = 0
                   If ss2.Text <> "Input" And ss2.Text <> "Delete" Then
                       ss2.Text = "Update"
                    End If
        
            End If
                
       End If
     
   
       If Col = 10 Then
                ss2.Col = Col
                ss2.Row = Row
                
       
             If ss2.Lock = False Then
                 
                   ss2.Col = 0
                   If ss2.Text <> "Input" And ss2.Text <> "Delete" Then
                       ss2.Text = "Update"
                    End If
        
            End If
                
       End If
     End If

  
End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        MDIMain.Mnu_Sorting.Enabled = False
        PopupMenu MDIMain.PopUp_Spread
        MDIMain.Mnu_Sorting.Enabled = False
    End If

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim sCid As String

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)

    End If

End Sub
