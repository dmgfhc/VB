VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQC0090C 
   Caption         =   "最终成分实绩修改及查询界面_AQC0090C"
   ClientHeight    =   9210
   ClientLeft      =   855
   ClientTop       =   1935
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   14955
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   315
      Left            =   6960
      TabIndex        =   13
      Top             =   180
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      _Version        =   196609
      BackColor       =   14804173
      Begin Threed.SSOption SSOption_All 
         Height          =   225
         Left            =   30
         TabIndex        =   14
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   196609
         BackColor       =   14804173
         Caption         =   "全部"
         Value           =   -1
      End
      Begin Threed.SSOption SSOption_Checked 
         Height          =   225
         Left            =   2460
         TabIndex        =   15
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   196609
         BackColor       =   14804173
         Caption         =   "已确认"
      End
      Begin Threed.SSOption SSOption_NoCheck 
         Height          =   225
         Left            =   1245
         TabIndex        =   16
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   397
         _Version        =   196609
         BackColor       =   14804173
         Caption         =   "未确认"
      End
   End
   Begin VB.TextBox txt_CHECK_STS 
      Height          =   315
      Left            =   14460
      TabIndex        =   12
      Text            =   "A"
      Top             =   45
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmd_ANALYST_CHECK 
      Caption         =   "成分检查确认"
      Height          =   315
      Left            =   10620
      TabIndex        =   11
      Top             =   180
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "最终成分实绩"
      Height          =   2430
      Left            =   120
      TabIndex        =   9
      Top             =   3285
      Width           =   15000
      Begin FPSpread.vaSpread ss3 
         Height          =   2055
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14745
         _Version        =   393216
         _ExtentX        =   26009
         _ExtentY        =   3625
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
         MaxCols         =   37
         MaxRows         =   4
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0090C.frx":0000
      End
   End
   Begin VB.TextBox txt_emp_cd 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   14790
      Locked          =   -1  'True
      TabIndex        =   8
      Tag             =   "作业人员"
      Top             =   405
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.ComboBox cbo_LINE_NO 
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
      ItemData        =   "AQC0090C.frx":0BFC
      Left            =   6135
      List            =   "AQC0090C.frx":0C09
      TabIndex        =   7
      Tag             =   "机号"
      Top             =   180
      Width           =   600
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   2340
      Left            =   150
      TabIndex        =   5
      Top             =   705
      Width           =   15000
      _Version        =   393216
      _ExtentX        =   26458
      _ExtentY        =   4128
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
      MaxCols         =   14
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0090C.frx":0C15
   End
   Begin VB.TextBox txt_HEAT_OLC_NO 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   1245
      MaxLength       =   8
      TabIndex        =   4
      Tag             =   "炉号"
      Top             =   180
      Width           =   850
   End
   Begin InDate.UDate txt_Charge_Date 
      Height          =   300
      Left            =   3375
      TabIndex        =   3
      Tag             =   "出钢日期"
      Top             =   180
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
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
      Left            =   14520
      TabIndex        =   2
      Text            =   "B1"
      Top             =   405
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox txt_charge_no 
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
      Left            =   15480
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.TextBox txt_oper 
      Height          =   315
      Left            =   14730
      TabIndex        =   0
      Text            =   "1"
      Top             =   45
      Visible         =   0   'False
      Width           =   195
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   180
      Top             =   180
      Width           =   1005
      _ExtentX        =   1773
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   2340
      Top             =   180
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "出钢日期"
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
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabeL3 
      Height          =   315
      Left            =   5100
      Top             =   180
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "机  号"
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
   Begin FPSpread.vaSpread ss2 
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   15000
      _Version        =   393216
      _ExtentX        =   26458
      _ExtentY        =   5733
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
      MaxRows         =   11
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0090C.frx":1347
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E1E4CD&
      Caption         =   "成分确认后除气体元素外   其他元素不允许修改"
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   12270
      TabIndex        =   17
      Top             =   180
      Width           =   1995
   End
End
Attribute VB_Name = "AQC0090C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Quality Management System
'-- Sub_System Name   Quality System
'-- Program Name      CHEMISTRY
'-- Program ID        AQC0090C
'-- Document No
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2005.12
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
Public sDateTime As String              'Active Form Authority Setting

Dim pControl1 As New Collection      'Master Primary Key Collection
Dim nControl1 As New Collection      'Master Necessary Collection
Dim mControl1 As New Collection      'Master Maxlength check Collection
Dim iControl1 As New Collection      'Master Insert Collection
Dim rControl1 As New Collection      'Master Refer Collection
Dim cControl1 As New Collection      'Master Copy Collection
Dim aControl1 As New Collection      'Master -> Spread Collection
Dim lControl1 As New Collection      'Master Lock Collection

Dim Mc1 As New Collection           'Master Collection

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim Mc2 As New Collection           'Master Collection

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


Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim str_STLGRD_DETAIL As String
Dim str_STLGRD As String
Dim lngActiveRow As Long
Dim sOldAuthority As String         'Save First Load Authority

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Hsheet"              'form类型
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       
      Call Gp_Ms_Collection(txt_charge_no, "p", " ", " ", "i", " ", "a", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)

                
    'MASTER Collection
     'Mc1.Add Item:="AQC0090C.P_MODIFY", Key:="P-M"
     'Mc1.Add Item:="AQC0090C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"


       Call Gp_Ms_Collection(txt_HEAT_OLC_NO, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
       Call Gp_Ms_Collection(txt_Charge_Date, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           Call Gp_Ms_Collection(cbo_LINE_NO, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
         Call Gp_Ms_Collection(txt_CHECK_STS, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
           'Call Gp_Ms_Collection(cbo_LINE_NO, " ", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl1, iControl1, rControl1, aControl1, lControl1)
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"

 
     'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQC0090C.P_SREFER", Key:="P-R"
'    Sc1.Add Item:="AQC0090T.P_SREFER", Key:="P-R"
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
    Call Gp_Sp_Collection(ss3, 1, "p", "n", " ", "i", "a", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 2, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 7, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 8, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss3, 9, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 10, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 11, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 12, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 14, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 15, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 17, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 20, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 21, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 23, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 26, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 27, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 28, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 29, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 30, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 31, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 32, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 33, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 34, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 35, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 36, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss3, 37, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   'Call Gp_Sp_Collection(ss3, 38, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

    'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AQC0090C.P_SMODIFY", Key:="P-M"
    sc3.Add Item:="AQC0090C.P_SREFER3", Key:="P-R"
    sc3.Add Item:="AQC0090C.P_SONEROW", Key:="P-O"
'    sc3.Add Item:="AQC0090T.P_SMODIFY", Key:="P-M"
'    sc3.Add Item:="AQC0090T.P_SREFER3", Key:="P-R"
'    sc3.Add Item:="AQC0090T.P_SONEROW", Key:="P-O"
    sc3.Add Item:=pColumn2, Key:="pColumn"
    sc3.Add Item:=nColumn2, Key:="nColumn"
    sc3.Add Item:=aColumn2, Key:="aColumn"
    sc3.Add Item:=mColumn2, Key:="mColumn"
    sc3.Add Item:=iColumn2, Key:="iColumn"
    sc3.Add Item:=lColumn2, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=sc3, Key:="Sc3"
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = "◎"
    Call Gp_Sp_ColHidden(ss1, 12, True)
    Call Gp_Sp_ColHidden(ss1, 13, True)

    Call Gp_Sp_ColHidden(ss3, 5, True)
    Call Gp_Sp_ColHidden(ss3, 6, True)
    
    Call Gp_Sp_ColHidden(ss3, 11, True)
    Call Gp_Sp_ColHidden(ss3, 12, True)
    
    Call Gp_Sp_ColHidden(ss3, 17, True)
    Call Gp_Sp_ColHidden(ss3, 18, True)
 
    Call Gp_Sp_ColHidden(ss3, 23, True)
    Call Gp_Sp_ColHidden(ss3, 24, True)
    
    Call Gp_Sp_ColHidden(ss3, 29, True)
    Call Gp_Sp_ColHidden(ss3, 30, True)
    
    Call Gp_Sp_ColHidden(ss3, 35, True)
    Call Gp_Sp_ColHidden(ss3, 36, True)
    
    Call Gp_Sp_ColHidden(ss3, 37, True)
    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQC0090C.P_SREFER2", Key:="P-R"
'     sc2.Add Item:="AQC0090T.P_SREFER2", Key:="P-R"
'    sc2.Add Item:="AFK2030C.P_REFER", Key:="P-R"
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    lngActiveRow = 0
    str_STLGRD_DETAIL = "标准"
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
End Sub

Private Sub cmd_ANALYST_CHECK_Click()
    Dim OutParam(2, 4)      As Variant
    Dim ret_Result_ErrMsg   As String
    Dim sHeatNo              As String
    Dim sQuery              As String
    
    
    Dim adoCmd As adodb.Command
    
    With ss3
        .Col = 1
        .Row = .ActiveRow
        sHeatNo = Trim(.Text)
    End With
    If Proc_Sc("Sc3")("Spread").MaxRows < 1 Then Exit Sub
    
'    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
   sAuthority = Chem_Check_AUTH(sUserID, sOldAuthority)
   
  
   If sAuthority = "1000" Or sAuthority = "0000" Then
       Call MsgBox("你没有当前操作权限！", vbOKOnly, "系统提示")
       Exit Sub
   End If
   
  
   If MsgBox("成分确认后将不可修改，是否确认？", vbOKCancel, "系统提示") <> vbOK Then   '2011.9.15 liuxiang
       Exit Sub
   End If
    
    On Error GoTo Process_Exec_ERROR

        
    Screen.MousePointer = vbHourglass
    
    'Return Error Code Parameter
    OutParam(1, 1) = "arg_e_code"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 2

    'Return Error Messsage Parameter
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    Set adoCmd.ActiveConnection = M_CN1
            
    '---------squery(CALL AQT1320P)----------------------
    sQuery = "{CALL AQC1920P('" & sHeatNo & "','" & sUserID & "'," & "?,?)}"   '20090923 坯料成分判定程序
'    sQuery = "{CALL AQD0610P('" & sHeatNo & "','" & sUserID & "'," & "?,?)}"
    '-------------------------------------------------------
    
    adoCmd.CommandType = adCmdText
    adoCmd.CommandText = sQuery
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_code") <> "YY" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        GoTo Process_Exec_ERROR
    End If
    Set adoCmd = Nothing
    
    Call Gp_MsgBoxDisplay("处理完了..!!", "I")
    Screen.MousePointer = vbDefault
    Call Form_Ref
    Exit Sub
    

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("Process_Exec_ERROR : " & Error & "   " & ret_Result_ErrMsg)



End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    With MDIMain.MenuTool
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(13).Enabled = False                'Separator
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
    txt_emp_cd.Text = sUserID
    sOldAuthority = sAuthority
    Call Form_Define
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_ControlLock(Mc2("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc3")("Spread"), 2)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc3")("Spread"), 8)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc3")("Spread"), 14)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc3")("Spread"), 20)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc3")("Spread"), 26)
    Call Gp_Sp_HdColColor(Proc_Sc("Sc3")("Spread"), 32)
    txt_emp_cd.Text = sUserID
    
    Screen.MousePointer = vbDefault
    Call Sp_Header_display(Proc_Sc("Sc2")("Spread"))
    Call LC_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    
     Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    
    
    With MDIMain.MenuTool
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(13).Enabled = False                'Separator
        .Buttons(14).Enabled = True                 'Excel
    End With


End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Gf_Sp_ProceExist(Proc_Sc("Sc3")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl1 = Nothing
    Set nControl1 = Nothing
    Set iControl1 = Nothing
    Set rControl1 = Nothing
    Set cControl1 = Nothing
    Set aControl1 = Nothing
    Set lControl1 = Nothing
    Set mControl1 = Nothing
    
    Set Mc1 = Nothing

    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing

    Set Mc2 = Nothing
    
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
    
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set sc3 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Exit()

    Unload Me
    
End Sub

Public Sub Form_Cls()

    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    ss2.ClearRange 3, 1, ss2.MaxCols, ss2.MaxRows, True
    Call Gp_Sp_BlockColor(Proc_Sc("Sc1")("Spread"), 3, ss2.MaxCols, 1, ss2.MaxRows)

    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_ControlLock(Mc2("pControl"), False)
         
    txt_HEAT_OLC_NO.Text = ""
    txt_Charge_Date.Text = ""
    cbo_LINE_NO.Text = ""
    txt_HEAT_OLC_NO.SetFocus
    txt_charge_no.Text = ""
    str_STLGRD_DETAIL = "标准"
    
End Sub

Public Sub Form_Ref()
        
    On Error GoTo Refer_Err
    If txt_HEAT_OLC_NO.Text = "" And txt_Charge_Date.RawData = "" Then
        Call Gp_MsgBoxDisplay("炉号和查询日期不能同时为空")
        txt_HEAT_OLC_NO.SetFocus
        Exit Sub
    ElseIf txt_HEAT_OLC_NO.Text <> "" And txt_Charge_Date.RawData <> "" Then
        txt_HEAT_OLC_NO.Text = ""
    End If
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc2, Mc2("nControl")) Then
        Call Gp_Sp_ReadOnlySet(ss1)
    End If
    If ss1.MaxRows > 0 Then
        ss1.Row = 1: ss1.Col = 1: txt_charge_no.Text = Trim(ss1.Text)
        ss1.Col = 3: str_STLGRD = Trim(ss1.Text)
        ss1.Col = 4: str_STLGRD_DETAIL = Trim(ss1.Text)
        Call Sp_Refer2
        'Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl"))
        'Call Ms_Chm_BitsSet(str_STLGRD, Mc1("rControl"))
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc3"), Mc1, Mc1("nControl"))
        Call GS_SetChemicalSpreadLineColor(ss3, "06111621")
        Call spChemLenSet(str_STLGRD)
        Call SP_ColSet
        txt_HEAT_OLC_NO.Text = txt_charge_no.Text
        ss1.SetFocus
        lngActiveRow = 1
        ss3.OperationMode = OperationModeNormal
    End If
    
    Call CHEM_CHECK
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
   
    Dim sMesg As String
    Dim iRow  As Integer
    Dim iLastRow As Integer
    
    txt_charge_no.Text = Mid(txt_charge_no.Text, 1, 8)
     
'    For iRow = 1 To ss3.MaxRows Step 1
'        ss3.Row = iRow
'        ss3.Col = ss3.MaxCols:  ss3.Text = ""
'    Next iRow
'
    For iRow = ss3.MaxRows To 1 Step -1
        ss3.Row = iRow
        ss3.Col = ss3.MaxCols:  ss3.Text = ""
        ss3.Col = 0:
        If Trim(ss3.Text) <> "" And iLastRow = 0 Then
            iLastRow = ss3.Row
            ss3.Col = ss3.MaxCols
            ss3.Text = "Y"
        End If
    Next iRow
    
    If Len(Trim(txt_charge_no.Text)) <> 8 Then
        sMesg = sMesg + " 炉号必须是8位"
        Call Gp_MsgBoxDisplay(sMesg)
       Exit Sub
    Else
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc3"), Mc1) Then
            Call MDIMain.FormMenuSetting(Me, FormType, "SE", sAuthority)
            Call spChemLenSet(str_STLGRD)
            Call SP_ColSet
        End If
    End If
    
   Call Form_Ref
    
End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc3"))
    Call spRowChemLenSet(str_STLGRD, ss3.ActiveRow)
    Call Sp_CellLock(2, ss3.ActiveRow)
    Call Sp_CellLock(8, ss3.ActiveRow)
    Call Sp_CellLock(14, ss3.ActiveRow)
    Call Sp_CellLock(20, ss3.ActiveRow)
    Call Sp_CellLock(26, ss3.ActiveRow)
    Call Sp_CellLock(32, ss3.ActiveRow)
End Sub



Private Sub txt_Charge_Date_DblClick()

    txt_Charge_Date.RawData = Format(Now, "YYYYMMDD")
        
End Sub

Public Sub Sp_Header_display(sPname As Variant)

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim sQuery As String
    
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New adodb.Recordset
    
    sQuery = " SELECT CHEM_COMP_CD From QP_CHEM_SEQ ORDER BY CHEM_COMP_SEQ ASC "
    
    With sPname

        .ReDraw = False
        .MaxCols = 2
        .MaxRows = 7
        Screen.MousePointer = vbHourglass
        
        'Title Setting
        .Col = 1
        .Row = 0
        .Text = "工序(标准)\成分"

        .Row = 1
        .Text = str_STLGRD_DETAIL

        .Row = 4
        .Text = "转炉"

        .Row = 5
        .Text = "LF"

        .Row = 6
        .Text = "VD/RH"

        .Row = 7
        .Text = "CCM"

        .Col = 2

        .Row = 1
        .Text = "最小值"
        .Row = 2
        .Text = "最大值"
        .Row = 3
        .Text = "目标值"

        .Row = 4
        .Text = "实绩"

        .Row = 5
        .Text = "实绩"

        .Row = 6
        .Text = "实绩"

        .Row = 7
        .Text = "实绩"
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) + 1 <> 0 Then
        
            .MaxCols = (UBound(ArrayRecords, 2) + 1) + 2
            .Row = 0
        
            For iCol = 2 To .MaxCols - 1
            
                .Col = iCol + 1
                .ColWidth(.Col) = 8
                
                If VarType(ArrayRecords(0, iCol - 2)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCol - 2))
                End If
                    
            Next iCol
            
        End If
        
        Call .AddCellSpan(1, 0, 2, 1)
        Call .AddCellSpan(1, 1, 1, 3)
        
        .BlockMode = True
        .Row = 0
        .Col = 1
        .Row2 = -1
        .Col2 = 2
        .TypeHAlign = TypeHAlignCenter
        .TypeVAlign = TypeVAlignCenter
        .BlockMode = False

        .ColsFrozen = 2
        .ReDraw = True
        
        Screen.MousePointer = vbDefault
        
    End With
    
Exit Sub

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Screen.MousePointer = vbDefault
    
End Sub

Public Sub LC_Sp_Setting(ByVal sPname As Variant)

    Dim iRow As Integer

    With sPname
    
        .RowHeight(-1) = 14
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 12
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 12
        Else
            .RowHeight(0) = 24
        End If
        
        .RowHeadersShow = False
        .BackColorStyle = BackColorStyleUnderGrid
        
        .GrayAreaBackColor = &HE0E0E0
        .GridColor = &H808040
        
        .ShadowColor = &HE1E4CD
        .ShadowDark = &H808040
        .SelBackColor = &HCEECFF     ''&HE3F4FF      ''&HFFFF80     '&H808040
     
        .OperationMode = OperationModeNormal
        .RetainSelBlock = True
        .UserResize = UserResizeColumns
        
        .ProcessTab = True
        .ScrollBarExtMode = True
        .ButtonDrawMode = 1
        .TabStop = False
        
        .Col = 0: .Col2 = -1
        .Row = 0: .Row2 = -1
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .Row = 0
        .FontBold = True
        
        For iRow = 1 To .MaxRows
            
            .Col = 3: .Col2 = .MaxCols
            .Row = iRow: .Row2 = iRow
            .BlockMode = True
                    
'            Select Case iRow
'                Case 1, 2, 3, 4, 6, 8, 10
                    .CellType = CellTypeNumber
                    .TypeNumberDecPlaces = 6
                    .TypeNumberMax = 99.999999
                    .TypeNumberMin = 0
                    .TypeNumberLeadingZero = TypeLeadingZeroYes
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignCenter
'                Case Else
'                    .CellType = CellTypeEdit
'                    .TypeHAlign = SS_CELL_H_ALIGN_CENTER
'                    .TypeVAlign = TypeVAlignCenter
'            End Select
            
            .BlockMode = False
                    
        Next iRow
        
    End With
    
End Sub
Private Function LC_Sp_Display(Conn As adodb.Connection, sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim icount As Integer
    Dim iRowCount As Long
    Dim iColCount As Long
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant

    LC_Sp_Display = True
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then LC_Sp_Display = False: Exit Function
    End If
    
    Set AdoRs = New adodb.Recordset
    
    With sPname

        .ReDraw = False
        icount = 0
        
        .ClearRange 3, 1, .MaxCols, .MaxRows, True
        Call Gp_Sp_BlockColor(Proc_Sc("Sc2")("Spread"), 3, .MaxCols, 1, .MaxRows)
    
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
            
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            LC_Sp_Display = False
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) <> 0 Then
        
            For iColCount = 2 To .MaxCols - 1
            
                .Col = iColCount + 1
                
                For iRowCount = 1 To .MaxRows
                
                    .Row = iRowCount
                    
                    If VarType(ArrayRecords(iRowCount, iColCount - 2)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iRowCount, iColCount - 2))
                    End If
                    
                Next iRowCount
                
            Next iColCount
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    LC_Sp_Display = False
    Call Gp_MsgBoxDisplay("Query Failed..." & sQuery)
    Screen.MousePointer = vbDefault

End Function

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    If ss1.MaxRows < 1 Or Row < 1 Then
        Exit Sub
    End If
    If Gf_Sp_ProceExist(Proc_Sc("Sc3")("Spread")) Then
        Call ss1.SetActiveCell(0, lngActiveRow)
        Exit Sub
    End If

    ss1.Col = 1: ss1.Row = Row: txt_charge_no.Text = Trim(ss1.Text)
    ss1.Col = 3: str_STLGRD = Trim(ss1.Text)
    ss1.Col = 4: str_STLGRD_DETAIL = Trim(ss1.Text)
    Call Sp_Refer2
    'Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl"))
    'Call Ms_Chm_BitsSet(str_STLGRD, Mc1.Item("rControl"))
    Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc3"), Mc1, Mc1("nControl"))
    Call GS_SetChemicalSpreadLineColor(ss3, "06111621")
    Call spChemLenSet(str_STLGRD)
    Call SP_ColSet
    Call CHEM_CHECK
    
    txt_HEAT_OLC_NO.Text = txt_charge_no.Text
    lngActiveRow = ss1.ActiveRow
    ss3.OperationMode = OperationModeNormal
End Sub
Private Sub Sp_Refer2()
    On Error GoTo Refer_Err

    Dim sMsg As String
    Dim sQuery As String
    Dim sQuery_cnt As String
    txt_charge_no.Text = Mid(txt_charge_no.Text, 1, 8)
    
    sMsg = Gf_Ms_NeceCheck(Mc1("nControl"))
    If sMsg <> "OK" Then
        sMsg = sMsg + "必须输入"
        Call Gp_MsgBoxDisplay(sMsg)
        Exit Sub
    End If
    Call Sp_Header_Set(ss2)
    Call LC_Sp_Display(M_CN1, Proc_Sc("Sc2")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
                
Refer_Err:

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Debug.Print KeyCode
    Select Case KeyCode
    Case 33, 34, 38, 40
        If Gf_Sp_ProceExist(Proc_Sc("Sc3")("Spread")) Then
             Call ss1.SetActiveCell(0, lngActiveRow)
            Exit Sub
        End If

        ss1.Col = 1: ss1.Row = ss1.ActiveRow: txt_charge_no.Text = Trim(ss1.Text) + "标准"
        ss1.Col = 3: str_STLGRD = Trim(ss1.Text)
        ss1.Col = 4: str_STLGRD_DETAIL = Trim(ss1.Text)
        Call Sp_Refer2
        'Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl"))
        'Call Ms_Chm_BitsSet(str_STLGRD, Mc1("rControl"))
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc3"), Mc1, Mc1("nControl"))
        Call GS_SetChemicalSpreadLineColor(ss3, "0611162126")
        Call spChemLenSet(str_STLGRD)

        txt_HEAT_OLC_NO.Text = txt_charge_no.Text
        lngActiveRow = ss1.ActiveRow
    End Select
End Sub

Public Sub Sp_Header_Set(sPname As Variant)

On Error GoTo SpreadDisplay_Error

    With sPname

        .ReDraw = False
        Screen.MousePointer = vbHourglass
        
        'Title Setting
        .Col = 1
        .Row = 1
        .Text = str_STLGRD_DETAIL

        .ReDraw = True
        
        Screen.MousePointer = vbDefault
        
    End With
    
Exit Sub

SpreadDisplay_Error:
    
    ss1.ReDraw = True
    Screen.MousePointer = vbDefault
    
End Sub
Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc3"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), 5)
    Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), 11)
    Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), 17)
    Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), 23)
    Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), 29)
    Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), 35)
    With ss3
        .Protect = True
        .Row = .ActiveRow: .Row2 = .ActiveRow
        .Col = 1: .Col2 = .MaxCols
        
        .BlockMode = True
        .Lock = False
        .BlockMode = False

    End With

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

'
'Private Sub Ms_Chm_BitsSet(ByVal stl_GRD As String, Retcol As Collection)
'On Error GoTo BitsSet_Error
'
'    Dim sQuery As String
'    Dim sTemp As String
'    Dim Ctrl As Control
'    Dim int_Format As Integer
'    Dim str_Format As String
'    Dim AdoRs As adodb.Recordset
'    Dim dValue As Double
'
'    Set AdoRs = New adodb.Recordset
'
'    sTemp = "Select GF_CHEM_LEN( '" + stl_GRD + "','"
'    If Not Retcol Is Nothing Then
'
'        For Each Ctrl In Retcol
'
'            If TypeOf Ctrl Is TextBox Then
'
'                sQuery = sTemp + Trim(Ctrl.Text) + "') FROM DUAL"
'                'Ado Execute
'                AdoRs.Open sQuery, M_CN1, adOpenKeyset
'                If Not AdoRs.BOF And Not AdoRs.EOF Then
'
'                            If Not AdoRs.EOF Then
'                                If VarType(AdoRs.Fields(0)) = vbNull Then
'                                   str_Format = "0.0000"
'                                   int_Format = 4
'                            Else
'                                str_Format = Trim(AdoRs.Fields(0))
'                                int_Format = InStr(str_Format, ".") - 1
'                                If int_Format < 0 Then
'                                    int_Format = 0
'                                End If
''                                If Ctrl.Text = "H" Then
''                                    str_Format = "0.0"
''                                ElseIf Ctrl.Text = "O" Or Ctrl.Text = "N" Then
''                                    str_Format = "0"
''                                End If
'                            End If
'                        End If
'                End If
'                AdoRs.Close
'
'          ElseIf TypeOf Ctrl Is sidbEdit Then
'             dValue = Ctrl.Value
'             If int_Format = 0 Then
'                Ctrl.FmtDecDigits = 0
'             Else
'                 Ctrl.FmtDecDigits = Len(str_Format) - InStr(str_Format, ".")
'             End If
'             Ctrl.FmtIntDigits = int_Format
'             'Ctrl.FmtDecDigits = int_Format
'             Ctrl.RawData = str_Format
'             Ctrl.Value = dValue
'        End If
'        Next Ctrl
'
'    End If
'    Set AdoRs = Nothing
'    Exit Sub
'
'BitsSet_Error:
'    Set AdoRs = Nothing
'    Exit Sub
'
'End Sub
 
Private Sub subSetChemLength(ByVal stl_GRD As String, ByVal vSP As vaSpread, ByVal Col As Long, ByVal Row As Long)
On Error GoTo BitsSet_Error
    Dim sQuery As String
    Dim sTemp As String
    Dim int_Format As Integer
    Dim str_Format As String
    Dim AdoRs As adodb.Recordset
    Dim dValue As Double
    
    Set AdoRs = New adodb.Recordset
    
    sTemp = "Select GF_CHEM_KND( '" + stl_GRD + "','"
    
    With vSP
        
        .Row = Row
        
        .Col = Col
        sQuery = sTemp + Trim(.Text) + "') FROM DUAL"
         'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        If Not AdoRs.BOF And Not AdoRs.EOF Then

            If Not AdoRs.EOF Then
                If VarType(AdoRs.Fields(0)) = vbNull Then
                   str_Format = "9.9999"
                   int_Format = 4
            Else
                str_Format = Replace(Trim(AdoRs.Fields(0)), "0", "9")
                int_Format = InStr(str_Format, ".") - 1
                If int_Format < 0 Then
                    int_Format = 0
                End If

            End If
        End If
        End If
        AdoRs.Close
        .Col = Col + 1
        .TypeNumberMin = 0
        .TypeNumberMax = Val(str_Format)
        .TypeNumberDecPlaces = GF_GET_SPREAD_DECIMAL(Trim(str_Format))
            
    End With
    Set AdoRs = Nothing
    Exit Sub

BitsSet_Error:
    Set AdoRs = Nothing
    Exit Sub
End Sub

Private Sub spChemLenSet(ByVal stl_GRD As String)
Dim lngRow As Long
    If ss3.MaxRows < 1 Then Exit Sub
    lngRow = 1
    With ss3
        For lngRow = 1 To .MaxRows
            .Row = lngRow
            Call subSetChemLength(stl_GRD, ss3, 2, .Row)
            Call subSetChemLength(stl_GRD, ss3, 8, .Row)
            Call subSetChemLength(stl_GRD, ss3, 14, .Row)
            Call subSetChemLength(stl_GRD, ss3, 20, .Row)
            Call subSetChemLength(stl_GRD, ss3, 26, .Row)
            Call subSetChemLength(stl_GRD, ss3, 32, .Row)
        Next
    End With
End Sub

Private Sub spRowChemLenSet(ByVal stl_GRD As String, ByVal Row As Long)
    If Row < 1 Then Exit Sub
        Call subSetChemLength(stl_GRD, ss3, 2, Row)
        Call subSetChemLength(stl_GRD, ss3, 8, Row)
        Call subSetChemLength(stl_GRD, ss3, 14, Row)
        Call subSetChemLength(stl_GRD, ss3, 20, Row)
        Call subSetChemLength(stl_GRD, ss3, 26, Row)
        Call subSetChemLength(stl_GRD, ss3, 32, Row)
End Sub

Private Sub SP_ColSet()
    Call Gp_Sp_ColLock(ss3, 2, True)
    Call Gp_Sp_ColLock(ss3, 8, True)
    Call Gp_Sp_ColLock(ss3, 14, True)
    Call Gp_Sp_ColLock(ss3, 20, True)
    Call Gp_Sp_ColLock(ss3, 26, True)
    Call Gp_Sp_ColLock(ss3, 32, True)
    
    Call GS_SetChemicalSpreadLineColor(ss3, "0713")
    Call GS_SetChemicalSpreadLineColor(ss3, "1925")
    Call GS_SetChemicalSpreadLineColor(ss3, "3137")
End Sub

Private Sub Sp_CellLock(ByVal Col As Long, ByVal Row As Long)
    With ss3
        .Protect = True
        .Row = Row: .Row2 = Row
        .Col = Col: .Col2 = Col
        
        .BlockMode = True
        .Lock = True
        .BlockMode = False

    End With
End Sub
'Private Sub ss3_EditChange(ByVal Col As Long, ByVal Row As Long)
'    Select Case Col
'        Case 2, 7, 12, 17, 22
'        Call Sp_Chem_Check(ss3, ss3.ActiveCol, ss3.ActiveRow)
'    End Select
'End Sub

Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc3")("Spread"), Mode)
        Select Case Col
            Case 2, 8, 14, 20, 26, 32
                Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), Col + 2)
                
            Case 3, 9, 15, 21, 27, 33
                Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), Col + 1)
        End Select
    End If

End Sub

Private Sub ss3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Proc_Sc("Sc3")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_InAuthority(Proc_Sc("Sc3"), ss3.ActiveCol + 1)
    End If

    If Shift = 0 Then Proc_Sc("Sc3")("Spread").EditMode = True

End Sub

Private Sub ss3_KeyUp(KeyCode As Integer, Shift As Integer)
Dim strChemCD As String
Dim lngRow As Long
Dim lngCurRow As Long
Dim lngCurCol As Long
    If ss3.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss3.ActiveCol

        Case 2, 8, 14, 20, 26, 32

            If KeyCode = vbKeyF4 Then

                Set DD.sPname = Me.ss3

                DD.sWitch = "SP"
                DD.rControl.Add Item:=ss3.ActiveCol
                DD.nameType = "2"

                Call GF_CHEM_SEQ(M_CN1, KeyCode)
                Call Sp_Chem_Check(ss3, ss3.ActiveCol, ss3.ActiveRow)
            End If

    End Select
    
End Sub

Private Sub Sp_Chem_Check(ByVal vSP As vaSpread, ByVal Col As Long, ByVal Row As Long)
    Dim strChemCD As String
    Dim lngRow As Long
    Dim lngCurRow As Long
    Dim lngCurCol As Long
    
    With vSP
        .Col = Col: .Row = Row: strChemCD = Trim(.Text)
        If strChemCD = "" Then Exit Sub
        lngCurRow = Row: lngCurCol = Col
        For lngRow = 1 To .MaxRows
            .Row = lngRow
            .Col = 2
            If Trim(.Text) = strChemCD Then
                If .Col <> lngCurCol Or lngRow <> lngCurRow Then
                    Call Gp_MsgBoxDisplay("该成分已经存在！")
                    .Col = Col: .Row = Row: .Text = ""
                    Call .SetActiveCell(3, lngRow)
                    Exit Sub
                End If
            End If
            .Col = 8
            If Trim(.Text) = strChemCD Then
                If .Col <> lngCurCol Or lngRow <> lngCurRow Then
                     Call Gp_MsgBoxDisplay("该成分已经存在！")
                    .Col = Col: .Row = Row: .Text = ""
                    Call .SetActiveCell(9, lngRow)
                    Exit Sub
                End If
            End If
            .Col = 14
            If Trim(.Text) = strChemCD Then
                If .Col <> lngCurCol Or lngRow <> lngCurRow Then
                    Call Gp_MsgBoxDisplay("该成分已经存在！")
                    .Col = Col: .Row = Row: .Text = ""
                    Call .SetActiveCell(15, lngRow)
                    Exit Sub
                End If
            End If
            .Col = 20
            If Trim(.Text) = strChemCD Then
                If .Col <> lngCurCol Or lngRow <> lngCurRow Then
                     Call Gp_MsgBoxDisplay("该成分已经存在！")
                    .Col = Col: .Row = Row: .Text = ""
                    Call .SetActiveCell(21, lngRow)
                    Exit Sub
                End If
            End If

            .Col = 26
            If Trim(.Text) = strChemCD Then
                If .Col <> lngCurCol Or lngRow <> lngCurRow Then
                    Call Gp_MsgBoxDisplay("该成分已经存在！")
                    .Col = Col: .Row = Row: .Text = ""
                    Call .SetActiveCell(27, lngRow)
                    Exit Sub
                End If
            End If
            .Col = 32
            If Trim(.Text) = strChemCD Then
                If .Col <> lngCurCol Or lngRow <> lngCurRow Then
                    Call Gp_MsgBoxDisplay("该成分已经存在！")
                    .Col = Col: .Row = Row: .Text = ""
                    Call .SetActiveCell(33, lngRow)
                    Exit Sub
                End If
            End If
        Next
         Call subSetChemLength(str_STLGRD, vSP, Col + 1, Row)
    End With
    
End Sub

Private Sub ss3_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Select Case Col
        Case 2, 8, 14, 20, 26, 32
            Call Sp_Chem_Check(ss3, Col, Row)
    End Select
End Sub

Private Sub SSOption_All_Click(Value As Integer)
    If SSOption_All.Value = True Then
        txt_CHECK_STS.Text = "A"
    End If
End Sub

Private Sub SSOption_Checked_Click(Value As Integer)
    If SSOption_Checked.Value = True Then
        txt_CHECK_STS.Text = "Y"
    End If
End Sub

Private Sub SSOption_NoCheck_Click(Value As Integer)
    If SSOption_NoCheck.Value = True Then
        txt_CHECK_STS.Text = "N"
    End If
End Sub
Private Sub CHEM_CHECK()
Dim lngRow As Long

   With ss3
     For lngRow = 1 To .MaxRows
            .Row = lngRow
            .Col = 4
            If .Text = "N" Then
               .Col = 3
               .BackColor = &HFF&
            End If
            .Col = 10
            If .Text = "N" Then
               .Col = 9
               .BackColor = &HFF&
            End If
            .Col = 16
            If .Text = "N" Then
               .Col = 15
               .BackColor = &HFF&
            End If
            .Col = 22
            If .Text = "N" Then
               .Col = 21
               .BackColor = &HFF&
            End If
            .Col = 28
            If .Text = "N" Then
               .Col = 27
               .BackColor = &HFF&
            End If
            .Col = 34
            If .Text = "N" Then
               .Col = 33
               .BackColor = &HFF&
            End If
     Next
   End With
End Sub
