VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQE2050C 
   Caption         =   "市场质量异议台帐_AQE2050C"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_claim_no 
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
      Left            =   1020
      MaxLength       =   14
      TabIndex        =   18
      Top             =   5490
      Width           =   1545
   End
   Begin VB.TextBox txt_dept_pers 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9555
      MaxLength       =   3
      TabIndex        =   7
      Tag             =   "dept"
      Top             =   560
      Width           =   600
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   555
      Width           =   1680
   End
   Begin VB.TextBox txt_PLT_CD_NAME 
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
      Left            =   1680
      MaxLength       =   14
      TabIndex        =   1
      Tag             =   "CD_MANA_NO"
      Top             =   75
      Width           =   915
   End
   Begin VB.TextBox txt_PLT_CD 
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
      Left            =   840
      MaxLength       =   14
      TabIndex        =   0
      Tag             =   "工厂代码"
      Top             =   75
      Width           =   825
   End
   Begin VB.TextBox txt_KND 
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
      Left            =   6360
      MaxLength       =   14
      TabIndex        =   12
      Tag             =   "CD_MANA_NO"
      Text            =   "F"
      Top             =   45
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   2820
      TabIndex        =   9
      Top             =   -45
      Width           =   3045
      Begin VB.OptionButton opt_KND_C 
         BackColor       =   &H00E0E0E0&
         Caption         =   "完成"
         Height          =   285
         Left            =   2010
         TabIndex        =   14
         Top             =   135
         Width           =   915
      End
      Begin VB.OptionButton opt_KND_F 
         BackColor       =   &H00E0E0E0&
         Caption         =   "全部"
         Height          =   285
         Left            =   60
         TabIndex        =   11
         Top             =   120
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton opt_KND_P 
         BackColor       =   &H00E0E0E0&
         Caption         =   " "
         Height          =   285
         Left            =   990
         TabIndex        =   10
         Top             =   120
         Width           =   915
      End
   End
   Begin VB.TextBox txt_dept 
      BackColor       =   &H00FFFFFF&
      Height          =   310
      Left            =   1380
      MaxLength       =   3
      TabIndex        =   2
      Top             =   585
      Width           =   600
   End
   Begin VB.TextBox txt_dept_name 
      Height          =   315
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   570
      Width           =   1680
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   3060
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   53
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Left            =   30
      Top             =   75
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   529
      Caption         =   "工厂"
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
   Begin InDate.ULabel ULabel12 
      Height          =   300
      Index           =   2
      Left            =   3840
      Top             =   585
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "发生日期"
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
   Begin InDate.UDate dtp_ISU_DATE_FROM 
      Height          =   300
      Left            =   5160
      TabIndex        =   5
      Tag             =   "订单确认日期"
      Top             =   585
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
   Begin InDate.UDate dtp_ISU_DATE_TO 
      Height          =   300
      Left            =   6600
      TabIndex        =   6
      Tag             =   "订单确认日期"
      Top             =   585
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
   Begin InDate.ULabel ULabel7 
      Height          =   300
      Left            =   30
      Top             =   585
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      Caption         =   "责任单位"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   8220
      Top             =   570
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Caption         =   "责任人"
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
   Begin Threed.SSCheck Chk_ss1 
      Height          =   285
      Left            =   14325
      TabIndex        =   13
      Top             =   585
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
      Caption         =   "表1"
      Value           =   1
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4320
      Left            =   0
      TabIndex        =   15
      Top             =   990
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   7620
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
      MaxRows         =   15
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE2050C.frx":0000
   End
   Begin InDate.ULabel ULabel3 
      Height          =   300
      Left            =   30
      Top             =   5490
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      Caption         =   "CLAIM_NO"
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
   Begin Threed.SSCheck Chk_ss2 
      Height          =   285
      Left            =   14370
      TabIndex        =   16
      Top             =   5550
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   503
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
      Caption         =   "表2"
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   3825
      Left            =   0
      TabIndex        =   17
      Top             =   5895
      Width           =   15075
      _Version        =   393216
      _ExtentX        =   26591
      _ExtentY        =   6747
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
      MaxCols         =   26
      MaxRows         =   15
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE2050C.frx":0BC5
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   15120
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "AQE2050C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       市场质量异议台帐_AQE2050C
'-- Sub_System Name
'-- Program Name
'-- Program ID        AQE2050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Sun Bin
'-- Coder
'-- Date              2008.8.16
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

Dim pControl2 As New Collection     'Master Primary Key Collection
Dim nControl2 As New Collection     'Master Necessary Collection
Dim mControl2 As New Collection     'Master Maxlength check Collection
Dim iControl2 As New Collection     'Master Insert Collection
Dim rControl2 As New Collection     'Master Refer Collection
Dim cControl2 As New Collection     'Master Copy Collection
Dim aControl2 As New Collection     'Master -> Spread Collection
Dim lControl2 As New Collection     'Master Lock Collection

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

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(txt_PLT_CD, "P", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(dtp_ISU_DATE_FROM, "P", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(dtp_ISU_DATE_TO, "P", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(txt_KND, "P", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_dept, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_dept_pers, "P", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
         Call Gp_Ms_Collection(txt_claim_no, "p", " ", " ", " ", "r", "a", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 5, " ", "n", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
'
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQE2050C.P_REFER1", Key:="P-R"
    Sc1.Add Item:="AQE2050C.P_ONEROW1", Key:="P-O"
    Sc1.Add Item:="AQE2050C.P_MODIFY1", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss2, 1, "p", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
     Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 26, "P", " ", " ", "i", "a", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
'
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQE2050C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:="AQE2050C.P_ONEROW2", Key:="P-O"
    sc2.Add Item:="AQE2050C.P_REFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=2, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = "◎"
    
    
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
    
    Call Gp_Sp_Setting(Sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(Sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(Sc1.Item("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "Q-System.INI", Me.Name)
    
'    Call Gp_Sp_HdColColor(Sc1.Item("Spread"), 3)
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Sc1.Item("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl = Nothing
    Set nControl = Nothing
    Set iControl = Nothing
    Set rControl = Nothing
    Set cControl = Nothing
    Set aControl = Nothing
    Set lControl = Nothing
    Set mControl = Nothing
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
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
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub
Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("Sc"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc2) Then
        If Gf_Sp_Cls(Sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
'            rControl(1).SetFocus
        End If
    End If

End Sub

'Private Sub Form_KeyPress(KeyAscii As Integer)
'
'        If KeyAscii = KEY_RETURN Then
'            KeyAscii = 0
'            SendKeys "{TAB}"
'        End If
'
'End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case Me.ActiveControl.Name
    
            Case "txt_PLT_CD"
            
               If KeyCode = vbKeyF4 Then
               
               DD.sWitch = "MS"
               DD.rControl.Add Item:=txt_PLT_CD
               DD.rControl.Add Item:=txt_PLT_CD_NAME
               DD.nameType = "2"
               DD.sKey = "C0001"
               
               Call Gf_Common_DD(M_CN1, KeyCode)
              End If
              
             Case "txt_dept"
            
               If KeyCode = vbKeyF4 Then
               
               DD.sWitch = "MS"
               DD.rControl.Add Item:=txt_dept
               DD.rControl.Add Item:=txt_dept_name
               DD.nameType = "2"
               DD.sKey = "Q0076"
               
               Call Gf_Common_DD(M_CN1, KeyCode)
              End If
    End Select
      
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call Gf_Sp_Cls(sc2)
    End If
            
    Exit Sub

Refer_Err:

End Sub
Public Sub Form_Ins()

    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    
    If Chk_ss1.Value = -1 Then
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)
    Else
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)
        Call SPREAD_ITEM_COPY
    End If

End Sub
Public Sub Form_Pro()
        
    Call STS_SET
    If Chk_ss1.Value = -1 Then
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
            Call Gp_Ms_Cls(Mc2("rControl"))
            Call Gf_Sp_Cls(sc2)
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        End If
    Else
        If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc2) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
End Sub

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("Sc"))

End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub

Public Sub Form_Exit()
    Unload Me
End Sub


Private Sub opt_KND_C_Click()
   If opt_KND_C.Value = True Then
        txt_KND.Text = "C"
       Call Form_Ref
   End If
End Sub

Private Sub opt_KND_F_Click()
   If opt_KND_F.Value = True Then
        txt_KND.Text = "F"
       Call Form_Ref
   End If
End Sub

Private Sub opt_KND_P_Click()
   If opt_KND_P.Value = True Then
        txt_KND.Text = "P"
       Call Form_Ref
   End If
End Sub

'Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
'
'    lBlkcol1 = BlockCol
'    lBlkcol2 = BlockCol2
'    lBlkrow1 = BlockRow
'    lBlkrow2 = BlockRow2
'
'End Sub

'Private Sub ss1_Change(ByVal Col As Long, ByVal Row As Long)
'
'    If Gf_Sc_Authority(sAuthority, "U") Then
'
'        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), 0)
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 27)
'
'    End If
'    With ss1
'      .Col = 3
'      If .Text = "" Then
'      .Col = 4
'      .Text = ""
'      End If
'      .Col = 5
'       If .Text = "" Then
'      .Col = 6
'      .Text = ""
'      End If
'      .Col = 10
'      If .Text = "" Then
'      .Col = 11
'      .Text = ""
''      .Col = 12
''      .Text = ""
'      End If
'      .Col = 17
'      If .Text = "" Then
'      .Col = 18
'      .Text = ""
'      End If
'      .Col = 19
'      If .Text = "" Then
'      .Col = 20
'      .Text = ""
'      End If
'    End With
'End Sub

'Public Sub Spread_Del()
'
'    Call Gp_Sp_Del(Proc_Sc("Sc"))
'
'End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Sc1.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub
    
    ss1.Row = Row
    
    If Gf_Sp_ProceExist(sc2.Item("Spread")) Then Exit Sub
    
    ss1.Col = 1
    txt_claim_no.Text = ss1.Text
    
    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    
End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 24)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sTemp_Code As String

    If ss1.MaxRows < 1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Or KeyCode = 229 Then
        Exit Sub
    End If

    Select Case ss1.ActiveCol
    
        Case 3
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "C0001"
                DD.rControl.Add Item:=3
                
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            End If
       Case 4
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "B0005"
                DD.rControl.Add Item:=4
                
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            End If
     Case 7
        
            If KeyCode = vbKeyF4 Then
            
                Set DD.sPname = Me.ss1
                
                DD.sWitch = "SP"
                DD.sKey = "C0001"
                DD.rControl.Add Item:=7
                DD.rControl.Add Item:=8
                
                DD.nameType = "2"
                
                Call Gf_Common_DD(M_CN1, KeyCode)
                
            End If
      Case 9
            
                If KeyCode = vbKeyF4 Then
                
                    Set DD.sPname = Me.ss1
                    
                    DD.sWitch = "SP"
                    DD.sKey = "C0001"
                    DD.rControl.Add Item:=9
                    DD.rControl.Add Item:=10
                    
                    DD.nameType = "2"
                    
                    Call Gf_Common_DD(M_CN1, KeyCode)
                
            End If
     Case 11
            
                If KeyCode = vbKeyF4 Then
                
                    Set DD.sPname = Me.ss1
                    
                    DD.sWitch = "SP"
                    DD.sKey = "C0001"
                    DD.rControl.Add Item:=11
                    DD.rControl.Add Item:=12
                    
                    DD.nameType = "2"
                    
                    Call Gf_Common_DD(M_CN1, KeyCode)
                    
                End If
     Case 16
            
                If KeyCode = vbKeyF4 Then
                
                    Set DD.sPname = Me.ss1
                    
                    DD.sWitch = "SP"
                    DD.sKey = "C0001"
                    DD.rControl.Add Item:=16
                    DD.rControl.Add Item:=17
                    
                    DD.nameType = "2"
                    
                    Call Gf_Common_DD(M_CN1, KeyCode)
                    
                End If
    Case 18
            
                If KeyCode = vbKeyF4 Then
                
                    Set DD.sPname = Me.ss1
                    
                    DD.sWitch = "SP"
                    DD.sKey = "C0001"
                    DD.rControl.Add Item:=18
                    DD.rControl.Add Item:=19
                    
                    DD.nameType = "2"
                    
                    Call Gf_Common_DD(M_CN1, KeyCode)
                 End If
            
    End Select

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


Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(sc2.Item("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 24)
    End If
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 21)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss2_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss2_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

    If Row > 0 Then
        Set Active_Spread = Me.ss2
        PopupMenu MDIMain.PopUp_Spread
    End If

End Sub

Private Sub Chk_ss1_Click(Value As Integer)
    
    If Chk_ss1.Value = ssCBUnchecked Then
       If Chk_ss2.Value = ssCBUnchecked Then
            Chk_ss1.Value = ssCBChecked
       End If
       Exit Sub
    End If
   
    If Gf_Sp_Change(Proc_Sc, Sc1) Then
        Chk_ss1.ForeColor = &HFF&
        Chk_ss2.ForeColor = &H808080
        Chk_ss2.Value = ssCBUnchecked
    Else
        Chk_ss1.Value = ssCBUnchecked
        Chk_ss2.Value = ssCBChecked
    End If
        
End Sub

Private Sub Chk_ss2_Click(Value As Integer)
    
    If Chk_ss2.Value = ssCBUnchecked Then
        If Chk_ss1.Value = ssCBUnchecked Then
            Chk_ss2.Value = ssCBChecked
        End If
        Exit Sub
    End If
    
    If Gf_Sp_Change(Proc_Sc, sc2) Then
        Chk_ss1.ForeColor = &H808080
        Chk_ss2.ForeColor = &HFF&
        Chk_ss1.Value = ssCBUnchecked
    Else
        Chk_ss2.Value = ssCBUnchecked
        Chk_ss1.Value = ssCBChecked
    End If
        
End Sub

'Private Sub Code_Name(Conn As adodb.Connection, KeyCode As Integer)
'    Dim sOld_Code, sNew_Code  As String
'    Dim sOld_Name, sNew_Name  As String
'
'    DD.DataDicType = "EMP"      'Program ID
'    DD.DicRefType = "C"         'Active Form DataDic Call
'
'        DD.sPname.Col = DD.rControl.Item(1)
'        sOld_Code = DD.sPname.Text
'
'        DD.sQuery = "            SELECT EMP_ID ""人员 ID"", EMP_NAME ""人员名称"" FROM  NISCO.ZP_EMPLOYEE "
''        DD.sWhere = "             WHERE DESCRIPTION  LIKE '" & Trim(DD.sKey) & "' "
'
'    If Gf_DD_Display(Conn, DD.sQuery + DD.sWhere, False) Then
'
'        If DD.sWitch = "SP" Then
'
'            DD.sPname.Col = DD.rControl.Item(1)
'            sNew_Code = DD.sPname.Text
'
'            If DD.rControl.COUNT > 1 Then
'                DD.sPname.Col = DD.rControl.Item(2)
'                sNew_Name = DD.sPname.Text
'            End If
'
'            DD.sPname.TabStop = True
'            DD.sPname.SetFocus
'            DD.sPname.SetActiveCell DD.rControl.Item(1), DD.sPname.ActiveRow
'            DD.sPname.Action = SS_ACTION_ACTIVE_CELL
'            DD.sPname.EditMode = True
'            DD.sPname.TabStop = False
'
'            If DD.sSelect Then
'                If sOld_Code <> sNew_Code Then Call Gp_Sp_UpdateMake(DD.sPname, False)
'            End If
'        End If
'
'    End If
'
'    DD.sWitch = ""
'    DD.sSelect = False
'
'    Set DD.sPname = Nothing
'    Set DD.rControl = Nothing
'
'
'End Sub
Private Sub STS_SET()
    Dim REASON_CD_STS As String
    Dim RSLT_STS As String
    
     With ss1
     .Col = 14
     .Row = .ActiveRow
     REASON_CD_STS = .Text
     .Col = 15
     RSLT_STS = .Text
     .Col = 2
     If REASON_CD_STS = "" And RSLT_STS = "" Then
        .Text = "P"
     ElseIf REASON_CD_STS <> "" Or RSLT_STS <> "" And .Text = "A" Then
        .Text = "C"
    Else: Exit Sub
    End If
    
     End With
End Sub
'Private Sub Master_To_Spread()
'With ss1
'    .Col = 2
'    .Row = .ActiveRow
'    .Text = txt_KND.Text
'
'End With
'
'End Sub

'
'Private Sub txt_emp_cd_Change()
'    Dim ID_CD As String
'    Dim sQuery As String
'    ID_CD = txt_EMP_CD.Text
'
'    If txt_EMP_CD.Text = "" Then
'       txt_EMP_NAME.Text = ""
'    End If
'
'    If txt_EMP_CD.Text <> "" And Len(txt_EMP_CD) = 7 Then
'       sQuery = "SELECT EMP_NAME FROM NISCO.ZP_EMPLOYEE WHERE EMP_ID='" + ID_CD + "'"
'       txt_EMP_NAME.Text = Gf_FloatFind(M_CN1, sQuery)
'     End If
'End Sub


'Private Sub txt_dept_KeyUp(KeyCode As Integer, Shift As Integer)
'
'    If KeyCode = vbKeyF4 Then
'
'        DD.sWitch = "MS"
'        DD.sKey = "Z0002"
'        DD.rControl.Add Item:=txt_dept
'        DD.rControl.Add Item:=txt_dept_name
'
'        DD.nameType = "2"
'
'        Call Gf_Common_DD(M_CN1, KeyCode)
'
'        Exit Sub
'
'    End If
'
'
'    If Len(Trim(txt_dept.Text)) = txt_dept.MaxLength Then
'        txt_dept_name.Text = Gf_ComnNameFind(M_CN1, "Z0002", txt_dept.Text, 2)
'    Else
'        txt_dept_name.Text = ""
'    End If
'
'End Sub

Private Sub SPREAD_ITEM_COPY()
Dim CLAIM_NO As String

    With ss1
        .Row = .ActiveRow
        .Col = 1
        CLAIM_NO = .Text
    End With
    
    With ss2
        .Row = .ActiveRow
        .Col = 26
        .Text = CLAIM_NO
    End With
End Sub






