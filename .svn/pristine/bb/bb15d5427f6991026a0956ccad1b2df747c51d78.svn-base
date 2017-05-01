VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form AQE1051C 
   Caption         =   "南钢中厚板卷厂板坯质量情况_AQE1051C"
   ClientHeight    =   7620
   ClientLeft      =   675
   ClientTop       =   4230
   ClientWidth     =   15135
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   15135
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_PRC_LINE 
      Height          =   315
      Left            =   6945
      TabIndex        =   15
      Top             =   105
      Width           =   405
   End
   Begin VB.TextBox Txt_SLAB_THK 
      Height          =   315
      Left            =   12825
      TabIndex        =   14
      Top             =   105
      Width           =   570
   End
   Begin VB.TextBox txt_STLGRD 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8250
      MaxLength       =   11
      TabIndex        =   13
      Tag             =   "钢种"
      Top             =   105
      Width           =   1305
   End
   Begin VB.TextBox txt_STLGRD_DETAIL 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9555
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   105
      Width           =   2250
   End
   Begin VB.Frame Frame2 
      Height          =   15
      Left            =   90
      TabIndex        =   8
      Top             =   4875
      Width           =   14865
   End
   Begin VB.TextBox TXT_BOF_CC 
      Height          =   330
      Left            =   -15
      MaxLength       =   1
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   135
      TabIndex        =   2
      Top             =   -45
      Width           =   1560
      Begin VB.OptionButton opt_bof 
         BackColor       =   &H00E0E0E0&
         Caption         =   "转炉"
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   180
         Width           =   720
      End
      Begin VB.OptionButton opt_ccm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "连铸"
         Height          =   255
         Left            =   765
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   14940
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4110
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   14910
      _Version        =   393216
      _ExtentX        =   26300
      _ExtentY        =   7250
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
      MaxCols         =   34
      MaxRows         =   13
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE1051C.frx":0000
   End
   Begin Threed.SSCommand cmd_Edit 
      Height          =   360
      Left            =   13800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   16711680
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "班组查询"
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   4440
      Left            =   90
      TabIndex        =   7
      Top             =   4815
      Width           =   14910
      _Version        =   393216
      _ExtentX        =   26300
      _ExtentY        =   7832
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
      MaxCols         =   35
      MaxRows         =   13
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQE1051C.frx":1237
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   1740
      Top             =   90
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      Caption         =   "时间从"
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
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   4110
      Top             =   90
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   556
      Caption         =   "时间到"
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
   Begin InDate.UDate txt_to_date 
      Height          =   315
      Left            =   4905
      TabIndex        =   9
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
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
   Begin InDate.UDate txt_from_date 
      Height          =   315
      Left            =   2535
      TabIndex        =   10
      Top             =   90
      Width           =   1380
      _ExtentX        =   2434
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
      Left            =   7440
      Top             =   105
      Width           =   795
      _ExtentX        =   1402
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   11850
      Top             =   105
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "板坯厚度"
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
      Height          =   315
      Left            =   6360
      Top             =   105
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Caption         =   "机号"
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
   Begin VB.Label Label1 
      Caption         =   "--"
      Height          =   180
      Left            =   3915
      TabIndex        =   11
      Top             =   150
      Width           =   255
   End
End
Attribute VB_Name = "AQE1051C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Slab Quality Analysis and Stat.
'-- Sub_System Name
'-- Program Name
'-- Program ID        AQE1050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2006.09.07
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
Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCnt As Integer
Dim iSumCol As New Collection       'Sum Column

Dim SumCnt1 As Integer
Dim SumCol1 As New Collection       'Sum Column

Dim Cur_Spread As Object            'Spread Object
Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    Dim i As Long
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Refer"

       'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
           Call Gp_Ms_Collection(TXT_BOF_CC, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_from_date, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_to_date, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(TXT_PRC_LINE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_STLGRD, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_STLGRD_DETAIL, " ", " ", " ", " ", " ", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(Txt_SLAB_THK, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  BELOW EDIT ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    'Sc1.Add Item:="AQE1050C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQE1052C.P_SREFER1", Key:="P-R"
    'Sc1.Add Item:="AQE1050C.P_ONEROW", Key:="P-O"
    
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
'------------------------------------  EDIT  End      ---------------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
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
   Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 23, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 24, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 25, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 26, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 27, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 28, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 29, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 30, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 31, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 32, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 33, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 34, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
   Call Gp_Sp_Collection(ss2, 35, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)

    
   'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQE1052C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    'Duplicate Count
    iDupCnt = 1
    
    'Sum Column Count
    SumCnt1 = 24
    
    'Sum Column Setting
    For i = 3 To 8
        SumCol1.Add Item:=i
    Next i
    
    SumCol1.Add Item:=12
    
    'Sum Column Setting
    For i = 14 To 22
        SumCol1.Add Item:=i
    Next i
    
    SumCol1.Add Item:=24
    
    'Sum Column Setting
    For i = 26 To 28
        SumCol1.Add Item:=i
    Next i
    
    SumCol1.Add Item:=30
    
    
    'Sum Column Setting
    For i = 32 To 34
        SumCol1.Add Item:=i
    Next i
    
    
    'Sum Column Count
    iSumCnt = 25
    
    'Sum Column Setting
    For i = 3 To 9
        iSumCol.Add Item:=i
    Next i
    
    iSumCol.Add Item:=13
    
    For i = 15 To 23
        iSumCol.Add Item:=i
    Next i
    
    iSumCol.Add Item:=25
    
    For i = 27 To 29
        iSumCol.Add Item:=i
    Next i
    
    iSumCol.Add Item:=31
    iSumCol.Add Item:=33
    iSumCol.Add Item:=34
    iSumCol.Add Item:=35
    
    
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cmd_Edit_Click()
    Dim i, j As Integer
    If Trim(txt_from_date.RawData) = "" Or Trim(txt_to_date.RawData) = "" Then
       MsgBox "查询日期未输入!", vbCritical, "系统提示信息"
       Exit Sub
    End If
    Call Gf_Sp_Display(M_CN1, ss2, Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
    Call Gp_Sp_EvenRowBackcolor(sc2.Item("Spread"))
    With ss2
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            Select Case .Text
            Case "A"
               .Col = 0
               .Text = "甲"
            Case "B"
               .Col = 0
               .Text = "乙"
            Case "C"
               .Col = 0
               .Text = "丙"
            Case "D"
               .Col = 0
               .Text = "丁"
            End Select
            
            For j = 3 To .MaxCols
                .Col = j
            If (j <> 12 And j <> 14 And j <> 24 And j <> 26 And j <> 30 And j <> 32) _
             And Val(.Text) = 0 Then
                .Text = ""
             End If
            Next j
        Next i
    End With

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
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc2")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    Set Cur_Spread = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)

    
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
    Set Sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing
    Set Cur_Spread = Nothing
   Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub


Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) And Gf_Sp_Cls(Proc_Sc("Sc2")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        'rControl(1).SetFocus
        Set Cur_Spread = Nothing
    End If

End Sub

Public Sub Form_Ref()

    Dim i, j As Integer
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
'    If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), 0, SumCnt1, SumCol1) Then
    If Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc"), Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", Mc1("pControl")), 1, 1, SumCnt1, SumCol1, False) Then
     
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        'Call Gp_Sp_EvenRowBackcolor(sc1.Item("Spread"))
        Call SS_DW(Sc1.Item("Spread"))
        With ss1
         For i = 1 To .MaxRows
             .Row = i
             For j = 2 To .MaxCols
                 .Col = j
                 If (j <> 11 And j <> 13 And j <> 23 And j <> 25 And j <> 29 And j <> 31) _
                 And Val(.Text) = 0 Then
                    .Text = ""
                 End If
             Next j
         Next i
        End With
    End If
    Call Group_Refer
    Call SS_DW(ss2)
    Set Cur_Spread = ss1
End Sub

'Public Sub Form_Pro()
'
'    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
'        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'    End If
'
'End Sub

'Public Sub Spread_ColumnsSort()
'
'    Spread_ColSort.Show 1
'
'End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    If Not Cur_Spread Is Nothing Then
        Call Gp_Sp_Excel(Me, Cur_Spread, lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Else
        Call Gp_MsgBoxDisplay("请先点击选择你要导出的数据表！", "I")
    End If
End Sub

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub opt_bof_Click()
    TXT_BOF_CC.Text = "1"
End Sub

Private Sub opt_ccm_Click()
    TXT_BOF_CC.Text = "2"
End Sub

Private Sub txt_STLGRD_Change()
    Dim sQuery As String
    If Len(Trim(txt_STLGRD.Text)) = 11 Then
        sQuery = "SELECT STEEL_GRD_DETAIL FROM QP_NISCO_CHMC WHERE STLGRD = '" + Trim(txt_STLGRD.Text) + "'"
        txt_STLGRD_DETAIL.Text = Gf_FloatFind(M_CN1, sQuery)
    Else
        txt_STLGRD_DETAIL.Text = ""
    End If
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        
        Case "txt_STLGRD"
            sCode = "STLGRD"
            Set oCodeName = txt_STLGRD_DETAIL
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    Set Cur_Spread = ss1
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub


'Private Sub ss1_LostFocus()
'
'    lBlkcol1 = 0
'    lBlkcol2 = 0
'    lBlkrow1 = 0
'    lBlkrow2 = 0
'
'End Sub

'Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)
'
'    If Row > 0 Then
'        Set Active_Spread = Me.ss1
'        PopupMenu MDIMain.PopUp_Spread
'    End If
'
'End Sub

Private Sub Sp_Zero_Clear()
    Dim i As Integer
    Dim j As Integer
    With ss1
     For i = 1 To .MaxRows
         .Row = i
         For j = 2 To .MaxCols
             .Col = j
             If .CellType = CellTypeNumber And _
             Val(.Text) = 0 Then
                .Text = ""
             End If
         Next j
     Next i
    End With
    With ss2
         For i = 1 To .MaxRows
             .Row = i
             For j = 2 To .MaxCols
                 .Col = j
                 If .CellType = CellTypeNumber And _
                 Val(.Text) = 0 Then
                    .Text = ""
                 End If
             Next j
         Next i
        End With
End Sub

Private Sub Group_Refer()
    Dim i, j As Integer
    'Call Gf_Sp_Display(M_CN1, ss2, Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
    Call Gf_Multi_Stotal_Display(M_CN1, Proc_Sc("Sc2"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")), 1, 1, iSumCnt, iSumCol, False)
    'Call Gp_Sp_EvenRowBackcolor(sc2.Item("Spread"))
    With ss2
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            Select Case .Text
            Case "A"
               .Col = 0
               .Text = "甲"
            Case "B"
               .Col = 0
               .Text = "乙"
            Case "C"
               .Col = 0
               .Text = "丙"
            Case "D"
               .Col = 0
               .Text = "丁"
            Case Else
               .Col = 0
               .Text = " "
            End Select

            For j = 3 To .MaxCols
                .Col = j
                If (j <> 12 And j <> 14 And j <> 24 And j <> 26 And j <> 30 And j <> 32) _
                 And Val(.Text) = 0 Then
                    .Text = ""
                 End If
            Next j
        Next i
    End With
End Sub

Private Sub ss2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss2_Click(ByVal Col As Long, ByVal Row As Long)
    Set Cur_Spread = ss2
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub SS_SUM_PERCENT(ByVal vaMySpread As vaSpread, Col_Total_PROC As Integer, Col_Acture_PROC As Integer, Col_Certif_PROC As Integer, Col_Once_Certif_PERC As Integer, Col_ShouD_PERC As Integer, Row As Long)

Dim Vaul_Col_Total_PROC        As Double
Dim Vaul_Col_Acture_PROC       As Double
Dim Scrap                      As Double
Dim Vaul_Col_Certif_PROC       As Double

  With vaMySpread
     .Row = Row
     .Col = Col_Total_PROC
      Vaul_Col_Total_PROC = .Text
     .Col = Col_Acture_PROC
      Vaul_Col_Acture_PROC = .Text
      If Vaul_Col_Total_PROC <> 0 Then
         If Vaul_Col_Acture_PROC <> 0 Then
            Scrap = Vaul_Col_Total_PROC - Vaul_Col_Acture_PROC
        Else: Scrap = Vaul_Col_Total_PROC
        End If
     Else
        Call MsgBox("生产总量错误！", vbOKOnly)
     End If
     
     If Vaul_Col_Acture_PROC <> 0 Then
        .Col = Col_Once_Certif_PERC
        .Text = (Vaul_Col_Acture_PROC - Scrap) / Vaul_Col_Acture_PROC * 100
     Else
        .Text = 0
'--        Call MsgBox("实际产量错误！", vbOKOnly)
     End If
         

     .Col = Col_Certif_PROC
      Vaul_Col_Certif_PROC = .Text
     .Col = Col_ShouD_PERC
     .Text = Vaul_Col_Certif_PROC / Vaul_Col_Total_PROC * 100
  End With
 
End Sub


Private Sub SS_DW(ByVal vaMySpread As vaSpread)
Dim iRow As Long
Dim iSaveRowNo() As Long
Dim iArrayNum As Integer

With vaMySpread
    .Col = 1
    ReDim iSaveRowNo(.MaxRows)
    For iRow = 1 To .MaxRows
        .Row = iRow
        If InStr(1, .Text, "小计") > 0 Then
            iSaveRowNo(iRow - 1) = iRow
        ElseIf InStr(1, .Text, "合计") > 0 Then
            iSaveRowNo(iRow - 1) = iRow
        Else
            iSaveRowNo(iRow - 1) = 0
        End If
        
    Next iRow

    For iArrayNum = 0 To .MaxRows
        If iSaveRowNo(iArrayNum) > 0 Then
            If .Name = "ss1" Then
                Call SS_SUM_PERCENT(vaMySpread, 3, 4, 5, 9, 10, iSaveRowNo(iArrayNum))
            ElseIf .Name = "ss2" Then
                Call SS_SUM_PERCENT(vaMySpread, 4, 5, 6, 10, 11, iSaveRowNo(iArrayNum))
            Else
                Exit Sub
            End If
        End If
    Next iArrayNum
End With

End Sub

