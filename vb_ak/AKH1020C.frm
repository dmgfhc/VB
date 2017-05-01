VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AKH1020C 
   Caption         =   "铁合金检验实绩录入及查询界面_AKH1020C"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_all1 
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
      Left            =   15240
      MaxLength       =   12
      TabIndex        =   11
      Tag             =   "铁合金代码"
      Top             =   120
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txt_all 
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
      Left            =   13680
      MaxLength       =   12
      TabIndex        =   9
      Tag             =   "铁合金代码"
      Top             =   120
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txt_all_cd1 
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
      Left            =   12720
      MaxLength       =   6
      TabIndex        =   7
      Tag             =   "铁合金代码"
      Top             =   120
      Visible         =   0   'False
      Width           =   705
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
      Left            =   2115
      MaxLength       =   50
      TabIndex        =   6
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
      Left            =   1410
      MaxLength       =   4
      TabIndex        =   5
      Tag             =   "铁合金代码"
      Top             =   120
      Width           =   705
   End
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
      ItemData        =   "AKH1020C.frx":0000
      Left            =   5550
      List            =   "AKH1020C.frx":0002
      TabIndex        =   0
      Tag             =   "ROLL_NO"
      Top             =   120
      Width           =   1965
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   4200
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "铁合金批号"
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   7800
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
      Left            =   10890
      TabIndex        =   1
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
      Left            =   9135
      TabIndex        =   2
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
   Begin FPSpread.vaSpread ss1 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   16380
      _Version        =   393216
      _ExtentX        =   28893
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
      MaxCols         =   23
      MaxRows         =   31
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AKH1020C.frx":0004
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      Caption         =   "铁合金代码"
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
   Begin FPSpread.vaSpread ss3 
      Height          =   2235
      Left            =   120
      TabIndex        =   8
      Top             =   7440
      Width           =   16410
      _Version        =   393216
      _ExtentX        =   28945
      _ExtentY        =   3942
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
      MaxCols         =   7
      MaxRows         =   2
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AKH1020C.frx":0D6C
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   3075
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   16410
      _Version        =   393216
      _ExtentX        =   28945
      _ExtentY        =   5424
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
      MaxCols         =   11
      MaxRows         =   2
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AKH1020C.frx":120C
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
      Left            =   10650
      TabIndex        =   3
      Top             =   180
      Width           =   255
   End
End
Attribute VB_Name = "AKH1020C"
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

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

Dim pColumn3 As New Collection      'Spread Primary Key Collection
Dim nColumn3 As New Collection      'Spread necessary Column Collection
Dim mColumn3 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn3 As New Collection      'Spread Insert Column Collection
Dim aColumn3 As New Collection      'Master -> Spread Column Collection
Dim lColumn3 As New Collection      'Spread Lock Column Collection



Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Sc3 As New Collection           'Spread Collection
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
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
 
 
  
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AKH1020C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="AKH1020C.P_REFER", Key:="P-R"
'    sc1.Add Item:="AKH1020C.P_ONEROW", Key:="P-O"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Call Gp_Sp_ColHidden(ss1, 6, True)

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "◎"
    

    
     'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
 
    
   
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AKH1020C.P_REFER1", Key:="P-R"
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
    
    
       'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn3, nColumn3, mColumn3, iColumn3, aColumn3, lColumn3)
   
    
        
   'Spread_Collection
    Sc3.Add Item:=ss3, Key:="Spread"
    Sc3.Add Item:="AKH1020C.P_REFER2", Key:="P-R"
    Sc3.Add Item:=pColumn2, Key:="pColumn"
    Sc3.Add Item:=nColumn2, Key:="nColumn"
    Sc3.Add Item:=aColumn2, Key:="aColumn"
    Sc3.Add Item:=mColumn2, Key:="mColumn"
    Sc3.Add Item:=iColumn2, Key:="iColumn"
    Sc3.Add Item:=lColumn2, Key:="lColumn"
    Sc3.Add Item:=1, Key:="First"
    Sc3.Add Item:=ss3.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=Sc3, Key:="Sc3"
    

    Sc3.Item("Spread").Col = 0
    Sc3.Item("Spread").Row = 0
    Sc3.Item("Spread").Text = "◎"
    

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

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1("Spread"))
    Call Gf_Sp_Cls(sc1)
    Call Gp_Sp_ColGet(sc1("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_HdColColor(sc1.Item("Spread"), 1)
    
    
    Call Gp_Sp_Setting(sc2("Spread"))
    Call Gf_Sp_Cls(sc2)
    Call Gp_Sp_ColGet(sc2("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_HdColColor(sc2.Item("Spread"), 1)
    
     Call Gp_Sp_Setting(Sc3("Spread"))
    Call Gf_Sp_Cls(Sc3)
    Call Gp_Sp_ColGet(Sc3("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_HdColColor(Sc3.Item("Spread"), 1)
    
    SDT_FROM_DATE.RawData = Mid(SDT_FROM_DATE.RawData, 1, 6) & "01"
   
    Screen.MousePointer = vbDefault
    
   sQuery_load = "SELECT CHECKNO FROM FP_ALL_MAIN  ORDER BY CHECKNO  "
   Call Gf_ComboAdd(M_CN1, CBO_ALL_NO, sQuery_load)

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
    
    Set iColumn3 = Nothing
    Set pColumn3 = Nothing
    Set lColumn3 = Nothing
    Set nColumn3 = Nothing
    Set mColumn3 = Nothing
    Set aColumn3 = Nothing

    Set Mc1 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Sc3 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Spread_Can()
    
    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
    
    Dim sCid As String
        
        ss1.Row = ss1.ActiveRow
        ss1.Col = 6
        sCid = ss1.Text
        If sCid <> "" Then
           ss1.Col = 5
           ss1.Text = sCid
        End If
        
End Sub

Public Sub Form_Cls()
    
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    
    txt_all_cd1.Text = ""
    txt_all.Text = ""
    txt_all1.Text = ""
    txt_all_cd = ""
    txt_all_name = ""
    CBO_ALL_NO = ""
   

End Sub

Public Sub Form_Ref()

  Dim i As Integer
  Dim sCid As String


QueryYN = False

      Call Gf_Sp_Cls(sc1)
      Call Gf_Sp_Cls(sc2)
      Call Gf_Sp_Cls(Sc3)



      If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
 
 
      If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
         Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)

         Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal


                    For i = 1 To ss1.MaxRows
                       ss1.Col = 6
                        ss1.Row = i
                       sCid = ss1.Text
                       If sCid <> "" Then
                       ss1.Col = 5
                       ss1.Text = sCid

                       End If
                      
                      ss1.Col = 8
                      ss1.Row = i
                      ss1.Lock = True
                      
                      ss1.Col = 9
                      ss1.Row = i
                      ss1.Lock = True
                    
                    Next i
                    
                         ss1.Col = 3
                         ss1.Row = 1
                         txt_all_cd1.Text = ss1.Text
                    
                         Call Gp_Scrap_Send

                         ss2.Col = 1
                         ss2.Row = ss2.ActiveRow
                         txt_all.Text = ss2.Text
                         ss2.Col = 4
                         ss2.Row = ss2.ActiveRow
                         txt_all1.Text = ss2.Text
                        
                        
                         Call Gp_Scrap_cnd
                    
                    

      End If
            


 Exit Sub

Refer_Err:

End Sub



Public Sub Form_Pro()
   Dim icount As Integer
   Dim MsgBox As String


   If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
        Call Form_Ref
   End If

End Sub



Public Sub Form_Ins()

    
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    txt_all_cd1.Text = ""
    txt_all.Text = ""
    txt_all1.Text = ""
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    ss1.Col = 20
    ss1.Row = ss1.ActiveRow
    ss1.Text = sUserID
'
'    Call Gp_Sp_ColLock(ss1, 1, False)
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 5
    ss1.BackColor = &HC0FFFF
    
    Call Pf_ComboAdd(M_CN1, ss1, 5, "SELECT CHECKNO FROM FP_ALL_MAIN  ORDER BY CHECKNO  ")
 
End Sub
Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    Proc_Sc("Sc").Item("Spread").OperationMode = OperationModeNormal
    
End Sub
Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
    ss1.Col = 15
    ss1.Row = ss1.ActiveRow
    ss1.Text = sUserID

'    Call Gp_Sp_ColLock(ss1, 1, False)
    
    ss1.Row = ss1.ActiveRow
    ss1.Col = 5
    ss1.BackColor = &HC0FFFF
    
       Call Pf_ComboAdd(M_CN1, ss1, 5, "SELECT CHECKNO FROM FP_ALL_MAIN  ORDER BY CHECKNO ")
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
    
    Call Gp_Sp_Del(Proc_Sc("SC"))


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
      
    Dim sQuery1 As String
    Dim sQuery2 As String
    Dim sQuery3 As String
    Dim MAXSEQ As String
    Dim MAXSEQ_1 As Integer
    Dim BEF_GRID As String
    
     
    Call Gf_Sp_Cls(sc2)
    Call Gf_Sp_Cls(Sc3)
    txt_all_cd1.Text = ""
    txt_all.Text = ""
    txt_all1.Text = ""
  

              If ss1.ActiveCol = 1 Then

                                    Select Case ss1.ActiveCol

                                        Case 1

                                            If KeyCode = vbKeyF4 Then

                                                Set DD.sPname = Me.ss1

                                                DD.sWitch = "SP"
                                                DD.sKey = "F0001"
                                                DD.rControl.Add Item:=1
                                                DD.rControl.Add Item:=2

                                                DD.nameType = "2"
                                                Call Gf_Common_DD(M_CN1, KeyCode)

                                            Else

                                                ss1.Col = ss1.ActiveCol

                                                 If Len(Trim(ss1.Text)) = 4 Then

                                                        sTemp_Code = ss1.Text
                                                        ss1.Col = 2
                                                        ss1.Text = Gf_ComnNameFind(M_CN1, "F0001", Trim(sTemp_Code), 2)
                                                 Else
                                                        ss1.Col = 2
                                                        ss1.Text = ""
                                                 End If

                                            End If


                                     End Select
                                    
'                        ss1.Row = ss1.ActiveRow
                        ss1.Col = 1
                        sQuery1 = "SELECT ERP_ALL_CD FROM FP_ALL_RELATIONSHIP  WHERE MES_ALL_CD = '" + Trim(ss1.Text) + "' "
                        ss1.Col = 3
                        ss1.Text = Gf_FloatFind(M_CN1, sQuery1)
                        txt_all_cd1 = Gf_CodeFind(M_CN1, sQuery1)
                        
                        ss1.Col = 1
                        sQuery1 = "SELECT ERP_ALL_NAME FROM FP_ALL_RELATIONSHIP  WHERE MES_ALL_CD = '" + Trim(ss1.Text) + "' "
                        ss1.Col = 4
                        ss1.Text = Gf_FloatFind(M_CN1, sQuery1)

                        Call Gp_Scrap_Send
                        
                     
                       ss2.Col = 1
                       ss2.Row = ss2.ActiveRow
                       txt_all.Text = ss2.Text
                       If txt_all.Text = "" Then
                           ss2.Col = 1
                           ss2.Row = 1
                           txt_all.Text = ss2.Text
                       End If
                       
                       
                       
                       ss2.Col = 4
                       ss2.Row = ss2.ActiveRow
                       txt_all1.Text = ss2.Text
                        If txt_all1.Text = "" Then
                           ss2.Col = 4
                           ss2.Row = 4
                           txt_all1.Text = ss2.Text
                       End If
                       
                       Call Gp_Scrap_cnd

                    


             End If

                
    

  
End Sub
        
        
'End Sub
Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim iRow  As Integer
Dim i, j, Scr_wgt, hm_wgt, Steel_wgt As Integer
  Dim sQuery2 As String

   If Row <> 0 Then
   
        If Col = 8 Then
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
   
        If Col = 21 Then
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
  
  If ss1.ActiveCol = 2 Then

         ss1.Col = 3
               ss1.Row = ss1.ActiveRow
               txt_all_cd1.Text = ss1.Text
               If txt_all_cd1.Text = "" Then
                   ss1.Col = 3
                   ss1.Row = 3
                 txt_all_cd1.Text = ss1.Text
               End If
        Call Gp_Scrap_Send
        
      
        
        ss2.Col = 1
        ss2.Row = ss2.ActiveRow
        txt_all.Text = ss2.Text
        ss2.Col = 4
        ss2.Row = ss2.ActiveRow
        txt_all1.Text = ss2.Text
        
        
        Call Gp_Scrap_cnd

   End If
   
 
  
End Sub


Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)

  Dim sCid As String
  Dim i As Integer
  Dim sQuery As String
  
  
  If Col = 9 Then
           ss1.Col = 9
           sCid = ss1.Text
           ss1.Col = 10
           ss1.Text = sCid
  End If

        
        
  If ss1.ActiveCol = 5 Then
  
       ss1.Row = ss1.ActiveRow
       ss1.Col = ss1.ActiveCol
       If Len(Trim(ss1.Text)) > 1 Then
       
         
          ss1.Col = 5
          sQuery = "SELECT JUDG_RST FROM FP_ALL_MAIN   WHERE CHECKNO = '" + Trim(ss1.Text) + "' "
          ss1.Col = 11
          ss1.Text = Gf_FloatFind(M_CN1, sQuery)


       Else
       
          ss1.Col = 11
          ss1.Text = ""

    End If
  End If
        
        

  End Sub
Private Sub ss2_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim iRow  As Integer
Dim i, j, Scr_wgt, hm_wgt, Steel_wgt As Integer

  
  ss2.Col = 1
  ss2.Row = ss2.ActiveRow
  txt_all.Text = ss2.Text
  ss2.Col = 4
  ss2.Row = ss2.ActiveRow
  txt_all1.Text = ss2.Text
  
  
  Call Gp_Scrap_cnd

  
End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    Dim sCid As String

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)

        If Col = 5 Then
           ss1.Col = 5
           sCid = ss1.Text
           ss1.Col = 6
           ss1.Text = sCid
        End If
        
        
       If Col = 9 Then
           ss1.Col = 9
           sCid = ss1.Text
           ss1.Col = 10
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
           DD.sKey = "F0001"
    
           DD.rControl.Add Item:=txt_all_cd
           DD.rControl.Add Item:=txt_all_name
           
           DD.nameType = "2"
           Call Gf_Common_DD(M_CN1, KeyCode)
    
    Else
    
        If Len(Trim(txt_all_cd.Text)) = txt_all_cd.MaxLength Then
            txt_all_name.Text = Gf_ComnNameFind(M_CN1, "F0001", txt_all_cd.Text, 2)
        Else
            txt_all_name.Text = ""
        End If
    End If
    
End Sub

Private Sub txt_all_cd_DblClick()

    Call txt_all_cd_KeyUp(vbKeyF4, 0)
    
End Sub


Public Sub Gp_Scrap_Send()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    

    sQuery = "{CALL AKH1020C.P_REFER1('" & Trim(txt_all_cd1.Text) & "')}"
    
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
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
    Else
'        Call MsgBox("成功排程！", vbInformation, "系统提示信息")
'        Call Form_Ref
        
         Call Gf_Sp_Display(M_CN1, ss2, sQuery)

    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub

Public Sub Gp_Scrap_cnd()

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    
    Dim adoCmd As ADODB.Command
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    

'    sQuery = "{CALL AKH1020C.P_REFER2('" & Trim(txt_all.Text) + Trim(txt_all1.Text) & "')}"
    
    sQuery = "{call AKH1020C.P_REFER2 ('" + txt_all + "', '" + txt_all1 + "')}"
    
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
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Sub
    Else
'        Call MsgBox("成功排程！", vbInformation, "系统提示信息")
'        Call Form_Ref
        
         Call Gf_Sp_Display(M_CN1, ss3, sQuery)

    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Sub
