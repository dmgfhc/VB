VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "spr32x60.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AKP1011C 
   Caption         =   "中厚板卷厂生产简报_AKP1011C"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9495
   ScaleWidth      =   12465
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text_UPD 
      Height          =   270
      Left            =   4140
      TabIndex        =   7
      Top             =   60
      Visible         =   0   'False
      Width           =   930
   End
   Begin Threed.SSCommand SSCommand_PRINT 
      Height          =   345
      Left            =   13065
      TabIndex        =   6
      Top             =   45
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      _Version        =   196609
      Enabled         =   0   'False
      Caption         =   "打印报表"
   End
   Begin Threed.SSCommand SSCommand_CREAT 
      Height          =   345
      Left            =   11670
      TabIndex        =   5
      Top             =   45
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   609
      _Version        =   196609
      Enabled         =   0   'False
      Caption         =   "生成报表"
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   9330
      Left            =   180
      TabIndex        =   1
      Top             =   450
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   16457
      _Version        =   196609
      PaneTree        =   "AKP1011C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   3885
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   14910
         _Version        =   393216
         _ExtentX        =   26300
         _ExtentY        =   6853
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         MaxRows         =   10
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP1011C.frx":0072
      End
      Begin FPSpread.vaSpread SS2 
         Height          =   1965
         Left            =   30
         TabIndex        =   3
         Top             =   4005
         Width           =   14910
         _Version        =   393216
         _ExtentX        =   26300
         _ExtentY        =   3466
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         ColHeaderDisplay=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   24
         MaxRows         =   10
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP1011C.frx":16D2
      End
      Begin FPSpread.vaSpread SS3 
         Height          =   3240
         Left            =   30
         TabIndex        =   4
         Top             =   6060
         Width           =   14910
         _Version        =   393216
         _ExtentX        =   26300
         _ExtentY        =   5715
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
         MaxCols         =   12
         MaxRows         =   7
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AKP1011C.frx":26FA
      End
   End
   Begin InDate.UDate txt_DATE 
      Height          =   315
      Left            =   1305
      TabIndex        =   0
      Tag             =   "起始日期"
      Top             =   90
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
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   180
      Top             =   90
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "日期"
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
End
Attribute VB_Name = "AKP1011C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       NISCO Production Management System
'-- Sub_System Name   Steel Making System
'-- Program Name      PROD REPORT
'-- Program ID        AFP1011C
'-- Designer          ZHENGWEN
'-- Coder             ZHENGWEN
'-- Date              2004.12.10
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
Public QueryYN As Boolean
Public psPrintDateFrom, psPrintDateto As String
Public crxApplication As New CRAXDRT.Application
Public Report As CRAXDRT.Report

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

Dim pColumn11 As New Collection      'Spread Primary Key Collection
Dim nColumn11 As New Collection      'Spread necessary Column Collection
Dim mColumn11 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn11 As New Collection      'Spread Insert Column Collection
Dim aColumn11 As New Collection      'Master -> Spread Column Collection
Dim lColumn11 As New Collection      'Spread Lock Column Collection

Dim pColumn12 As New Collection      'Spread Primary Key Collection
Dim nColumn12 As New Collection      'Spread necessary Column Collection
Dim mColumn12 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn12 As New Collection      'Spread Insert Column Collection
Dim aColumn12 As New Collection      'Master -> Spread Column Collection
Dim lColumn12 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Sc2 As New Collection           'Spread Collection
Dim sc3 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection
Dim iSumCol As New Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()
    Dim I As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Sheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
    Call Gp_Ms_Collection(txt_DATE, "p", "n", " ", "i", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    Call Gp_Sp_Collection1(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 8, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection1(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 12, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 14, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection1(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AFP1011C.P_SREFER1", Key:="P-R"
    Sc1.Add Item:="AFP1011C.P_MODIFY1", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    For I = 1 To 21
        Call Sp_ColLock(ss1, I, 2, True)
        Call Sp_ColLock(ss1, I, 4, True)
        Call Sp_ColLock(ss1, I, 6, True)
        Call Sp_ColLock(ss1, I, 8, True)
        Call Sp_ColLock(ss1, I, 9, True)
        Call Sp_ColLock(ss1, I, 10, True)
    Next I
    

    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection1(SS2, 1, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
    Call Gp_Sp_Collection1(SS2, 2, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
    Call Gp_Sp_Collection1(SS2, 3, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
    Call Gp_Sp_Collection1(SS2, 4, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
    Call Gp_Sp_Collection1(SS2, 5, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
    Call Gp_Sp_Collection1(SS2, 6, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
    Call Gp_Sp_Collection1(SS2, 7, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
    Call Gp_Sp_Collection1(SS2, 8, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
    Call Gp_Sp_Collection1(SS2, 9, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 10, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 11, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 12, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 13, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 14, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 15, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 16, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 17, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 18, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 19, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 20, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 21, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 22, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 23, " ", " ", " ", "i", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
   Call Gp_Sp_Collection1(SS2, 24, " ", " ", " ", " ", " ", " ", pColumn11, nColumn11, mColumn11, iColumn11, aColumn11, lColumn11)
    
    'Spread_Collection
    Sc2.Add Item:=SS2, Key:="Spread"
    Sc2.Add Item:="AFP1011C.P_SREFER2", Key:="P-R"
    Sc2.Add Item:="AFP1011C.P_MODIFY2", Key:="P-M"
    Sc2.Add Item:=pColumn11, Key:="pColumn"
    Sc2.Add Item:=nColumn11, Key:="nColumn"
    Sc2.Add Item:=aColumn11, Key:="aColumn"
    Sc2.Add Item:=mColumn11, Key:="mColumn"
    Sc2.Add Item:=iColumn11, Key:="iColumn"
    Sc2.Add Item:=lColumn11, Key:="lColumn"
    Sc2.Add Item:=1, Key:="First"
    Sc2.Add Item:=SS2.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=Sc2, Key:="Sc2"
    For I = 1 To 21
        Call Sp_ColLock(SS2, I, 2, True)
        Call Sp_ColLock(SS2, I, 4, True)
        Call Sp_ColLock(SS2, I, 6, True)
        Call Sp_ColLock(SS2, I, 8, True)
        Call Sp_ColLock(SS2, I, 9, True)
        Call Sp_ColLock(SS2, I, 10, True)
    Next I
    

    
    
        'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(SS3, 1, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS3, 2, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS3, 3, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS3, 4, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS3, 5, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS3, 6, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS3, 7, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS3, 8, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS3, 9, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS3, 10, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS3, 11, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS3, 12, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    
    'Spread_Collection
    sc3.Add Item:=SS3, Key:="Spread"
    sc3.Add Item:="AFP1011C.P_MODIFY3", Key:="P-M"
    sc3.Add Item:=pColumn12, Key:="pColumn"
    sc3.Add Item:=nColumn12, Key:="nColumn"
    sc3.Add Item:=aColumn12, Key:="aColumn"
    sc3.Add Item:=mColumn12, Key:="mColumn"
    sc3.Add Item:=iColumn12, Key:="iColumn"
    sc3.Add Item:=lColumn12, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss1.MaxRows, Key:="Last"
    Proc_Sc.Add Item:=sc3, Key:="Sc3"
    
    Call Sp_ColLock(SS3, 1, 1, True)
    Call Sp_ColLock(SS3, 1, 2, True)
    Call Sp_ColLock(SS3, 2, 1, True)
    Call Sp_ColLock(SS3, 2, 2, True)
    Call Sp_ColLock(SS3, 3, 1, True)
    Call Sp_ColLock(SS3, 3, 2, True)
    For I = 1 To 4
        Call Sp_ColLock(SS3, 4, I, True)
        Call Sp_ColLock(SS3, 5, I, True)
        Call Sp_ColLock(SS3, 7, I, True)
        Call Sp_ColLock(SS3, 9, I, True)
        Call Sp_ColLock(SS3, 11, I, True)
    Next I
    Call Sp_ColLock(SS3, 1, 6, True)
    Call Sp_ColLock(SS3, 3, 6, True)
    Call Sp_ColLock(SS3, 5, 6, True)
    Call Sp_ColLock(SS3, 7, 6, True)
    Call Sp_ColLock(SS3, 9, 6, True)
    Call Sp_ColLock(SS3, 11, 6, True)
    Call Sp_ColLock(SS3, 1, 7, True)
    Call Sp_ColLock(SS3, 3, 7, True)
    Call Sp_ColLock(SS3, 6, 7, True)
    Call Sp_ColLock(SS3, 9, 7, True)

    Call Gp_Sp_ColHidden(ss1, 22, True)
    Call Gp_Sp_ColHidden(ss1, 23, True)
    Call Gp_Sp_ColHidden(ss1, 24, True)
    Call Gp_Sp_ColHidden(SS2, 22, True)
    Call Gp_Sp_ColHidden(SS2, 23, True)
    Call Gp_Sp_ColHidden(SS2, 24, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()
Dim A, B, C As Variant

    Screen.MousePointer = vbHourglass
    
    sAuthority = Gf_Pgm_Authority(Me.Name)
    'sAuthority = "1234567"
    Call Form_Define
        
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Sp_Header_display(Proc_Sc("Sc1")("Spread"))
    Call Sp_Header_display(Proc_Sc("Sc2")("Spread"))
    Call Sp_Header_display(Proc_Sc("Sc3")("Spread"))
    
    Call Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Sp_Setting(Proc_Sc("Sc3")("Spread"))
    
'    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
'    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
'    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "F-System.INI", Me.Name)
    'Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "F-System.INI", Me.Name)
    If Gf_Sc_Authority(sAuthority, "U") Then
       SSCommand_CREAT.Enabled = True
       SSCommand_PRINT.Enabled = True
       
    End If
    

    txt_DATE.RawData = Format(Date, "yyyymmdd")
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist1(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    If Gf_Sp_ProceExist1(Proc_Sc("Sc2")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    If Sp_ProceExist(Proc_Sc("Sc3")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc2")("Spread"), "F-System.INI", Me.Name)
    Call Gp_Sp_ColSet(Proc_Sc("Sc3")("Spread"), "F-System.INI", Me.Name)
    
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
   
    Set iColumn11 = Nothing
    Set pColumn11 = Nothing
    Set lColumn11 = Nothing
    Set nColumn11 = Nothing
    Set mColumn11 = Nothing
    Set aColumn11 = Nothing
   
    Set iColumn12 = Nothing
    Set pColumn12 = Nothing
    Set lColumn12 = Nothing
    Set nColumn12 = Nothing
    Set mColumn12 = Nothing
    Set aColumn12 = Nothing
   
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set Sc2 = Nothing
    Set sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()
    ss1.ClearRange 1, 1, ss1.MaxCols, ss1.MaxRows, True
    SS2.ClearRange 1, 1, SS2.MaxCols, SS2.MaxRows, True
    Call ss3_clear
    Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_ControlLock(Mc1("lControl"), False)


End Sub

Public Sub Form_Ref()
Dim I, j As Integer
Dim iRow  As Integer
Dim sQuery As String
Dim iHmWgt As Double
Dim iOlcWgt1 As Double
Dim iOlcWgt2 As Double
Dim iOlcWgt3 As Double
On Error GoTo Refer_Err
    If Sp_Display(M_CN1, Proc_Sc("Sc1")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc1").Item("P-R"), "R", Mc1("pControl"))) Then
        Call Sp_Display(M_CN1, Proc_Sc("Sc2")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
        Call Sp_Data_Refer
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
        
Refer_Err:
End Sub
Public Sub Form_Pro()

    If Gf_Sp_Process1(M_CN1, Proc_Sc("SC1")) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    If Gf_Sp_Process1(M_CN1, Proc_Sc("Sc2"), Mc1) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    If Sp_Process(M_CN1, Proc_Sc("Sc3")) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
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
    
    'Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake1(Proc_Sc("Sc1")("Spread"), Mode)
        With ss1
             If .ActiveRow = 1 Then
             .Col = 22
             .Text = txt_DATE.RawData
             .Col = 23
             .Text = "A"
             End If
             
             If .ActiveRow = 3 Then
             .Col = 22
             .Text = txt_DATE.RawData
             .Col = 23
             .Text = "B"
             End If
             
             If .ActiveRow = 5 Then
             .Col = 22
             .Text = txt_DATE.RawData
             .Col = 23
             .Text = "C"
             End If
             
             If .ActiveRow = 7 Then
             .Col = 22
             .Text = txt_DATE.RawData
             .Col = 23
             .Text = "D"
             End If
             
        End With
    End If
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc1")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    


End Sub

Public Sub Sp_Header_display(sPname As Variant)

On Error GoTo SpreadDisplay_Error

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim sQuery As String

    
    With sPname

        .ReDraw = False
       ' .MaxCols = 17
        Screen.MousePointer = vbHourglass
        
          Screen.MousePointer = vbDefault
        
    End With
    
Exit Sub

SpreadDisplay_Error:
    

    Screen.MousePointer = vbDefault
    
End Sub
Public Sub Sp_Setting(ByVal sPname As Variant)

Dim iRow As Integer

  With sPname
'
  .RowHeight(-1) = 13
        
        If .ColHeaderRows > 1 Then
            .RowHeight(SpreadHeader + (.ColHeaderRows - 2)) = 13
            .RowHeight(SpreadHeader + (.ColHeaderRows - 1)) = 13
        Else
            .RowHeight(0) = 24
        End If
       
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
        .Row = 0: .row2 = -1
        
        
        .BlockMode = True
        .FontBold = False
        .FontName = "SimSun"
        .FontSize = 10
        .BlockMode = False
        
        .Col = -1
        .Row = 0
        .FontBold = True
    End With
    With SS3
         .ColHeadersShow = False
         .Col = 1
         .Row = 5
         .RowHeight(5) = 24
    End With
    
End Sub
Public Function Sp_Data_Refer() As Boolean

On Error GoTo SpreadDisplay_Error


    Dim sTdate As String
    Dim sQuery As String

    Dim NEW1, END1, DIFF, PLAN, ACT As Variant
    

    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New ADODB.Recordset
    

  
    sQuery = "SELECT PROD_DATE,CON_LIFE,TAP_TOT_WGT,TAP_AVE_WGT,BOF_D_RATE,MILL_D_RATE,CC_H_WGT,MILL_H_WGT,MON_S_PLN_PROD,MON_S_ACT_PROD,MON_M_PLN_PROD,"
    sQuery = sQuery + "MON_M_ACT_PROD,SLAB_WGT_PROG,AVE_DAY_SLAB_WGT,DAY_ADD_SLAB_WGT,PLATE_WGT_PROG ,AVE_DAY_PLATE_WGT,"
    sQuery = sQuery + "DAY_ADD_PLATE_WGT,MAX_G_BOF_CNT,MAX_G_BOF_DATE,MAX_G_BOF_GROUP,MAX_G_SLAB_WGT,MAX_G_SLAB_DATE,"
    sQuery = sQuery + "MAX_G_SLAB_GROUP,MAX_G_MILL_WGT,MAX_G_MILL_DATE ,MAX_G_MILL_GROUP,MAX_D_SLAB_WGT,MAX_D_SLAB_DATE ,"
    sQuery = sQuery + "MAX_D_MILL_WGT,MAX_D_MILL_DATE,MAX_M_SLAB_WGT,MAX_M_SLAB_DATE,MAX_M_MILL_WGT,MAX_M_MILL_DATE,DESCRIPTION,SCRAP_REM_WGT,YARD_SLAB_WGT"

    sQuery = sQuery + " FROM  FP_PROD_SUM_REPORT "
    sQuery = sQuery + "  WHERE PROD_DATE  = '" + txt_DATE.RawData + "' "
        
    With SS3

        Sp_Data_Refer = True
        .ReDraw = False

        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
        
            Sp_Data_Refer = False
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        AdoRs.Close
        Set AdoRs = Nothing
        
        .MaxRows = 7
    
        NEW1 = CDate(txt_DATE.Text)
        END1 = DateAdd("m", 1, NEW1)
        DIFF = DateDiff("d", NEW1, END1) - Val(Mid(txt_DATE.RawData, 7, 2))
        
'        For iRowCount = 0 To .MaxRows - 1
'
'            .Row = iRowCount + 1
'
'            For iColcount = 0 To .MaxCols - 1

'                        If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
'                            .Text = ""
'                        Else
'                            .Text = Trim(ArrayRecords(iColcount, iRowCount))
'                        End If

                .Col = 1
                .Row = 3
                If VarType(ArrayRecords(1, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(1, 0))
                End If
                
                .Col = 2
                .Row = 3
                If VarType(ArrayRecords(2, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(2, 0))
                End If
                
                .Col = 3
                .Row = 3
                If VarType(ArrayRecords(3, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(3, 0))
                End If
                
                .Col = 6
                .Row = 1
                If VarType(ArrayRecords(4, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(4, 0))
                End If
                
                .Col = 6
                .Row = 2
                If VarType(ArrayRecords(5, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(5, 0))
                End If

                .Col = 6
                .Row = 3
                If VarType(ArrayRecords(6, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(6, 0))
                End If


                .Col = 6
                .Row = 4
                If VarType(ArrayRecords(7, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(7, 0))
                End If

                .Col = 8
                .Row = 1
                If VarType(ArrayRecords(8, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(8, 0))
                    PLAN = ArrayRecords(8, 0)
                End If

                .Col = 8
                .Row = 2
                If VarType(ArrayRecords(9, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(9, 0))
                    ACT = ArrayRecords(9, 0)
                End If

                .Col = 8
                .Row = 3
                .Text = DIFF

                .Col = 8
                .Row = 4
                If DIFF <> 0 And VarType(ArrayRecords(8, 0)) <> vbNull And VarType(ArrayRecords(9, 0)) <> vbNull Then
                .Text = Trim(Round((PLAN - ACT) / DIFF, 3))

                Else
                 .Text = 0
                End If
                
                
                .Col = 10
                .Row = 1
                If VarType(ArrayRecords(10, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(10, 0))
                    PLAN = ArrayRecords(10, 0)
                End If

                .Col = 10
                .Row = 2
                If VarType(ArrayRecords(11, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(11, 0))
                    ACT = ArrayRecords(11, 0)
                End If
                

                .Col = 10
                .Row = 3
                .Text = DIFF

                .Col = 10
                .Row = 4
                If DIFF <> 0 And VarType(ArrayRecords(10, 0)) <> vbNull And VarType(ArrayRecords(11, 0)) <> vbNull Then
                .Text = Trim(Round(PLAN - ACT / DIFF, 3))
                Else
                 .Text = 0
                End If

'
                .Col = 2
                .Row = 6
                If VarType(ArrayRecords(12, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(12, 0))
                End If

                .Col = 4
                .Row = 6
                If VarType(ArrayRecords(13, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(13, 0))
                End If

                .Col = 6
                .Row = 6
                If VarType(ArrayRecords(14, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(14, 0))
                End If

                .Col = 8
                .Row = 6
                If VarType(ArrayRecords(15, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(15, 0))
                End If

                .Col = 10
                .Row = 6
                If VarType(ArrayRecords(16, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(16, 0))
                End If

                .Col = 12
                .Row = 6
                If VarType(ArrayRecords(17, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(17, 0))
                End If

                .Col = 2
                .Row = 7
                If VarType(ArrayRecords(18, 0)) = vbNull Then
                    .Text = "0"
                Else
                    .Text = Trim(ArrayRecords(18, 0))
                End If
                If VarType(ArrayRecords(19, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + "炉" + Trim(ArrayRecords(19, 0))
                End If
                If VarType(ArrayRecords(20, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + " " + Trim(ArrayRecords(20, 0)) + "班"
                End If
                

                .Col = 4
                .Row = 7
                If VarType(ArrayRecords(21, 0)) = vbNull Then
                    .Text = "0"
                Else
                    .Text = Trim(ArrayRecords(21, 0))
                End If
                If VarType(ArrayRecords(22, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + "吨" + Trim(ArrayRecords(22, 0))
                End If
                If VarType(ArrayRecords(23, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + " " + Trim(ArrayRecords(23, 0)) + "班"
                End If
                
                .Col = 5
                .Row = 7
                If VarType(ArrayRecords(24, 0)) = vbNull Then
                    .Text = "0"
                Else
                    .Text = Trim(ArrayRecords(24, 0))
                End If
                If VarType(ArrayRecords(25, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + "吨" + Trim(ArrayRecords(25, 0))
                End If
                If VarType(ArrayRecords(26, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + " " + Trim(ArrayRecords(26, 0)) + "班"
                End If
                
                .Col = 7
                .Row = 7
                If VarType(ArrayRecords(27, 0)) = vbNull Then
                    .Text = "0"
                Else
                    .Text = Trim(ArrayRecords(27, 0))
                End If
                If VarType(ArrayRecords(28, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + "吨" + Trim(ArrayRecords(28, 0))
                End If

                .Col = 8
                .Row = 7
                If VarType(ArrayRecords(29, 0)) = vbNull Then
                    .Text = "0"
                Else
                    .Text = Trim(ArrayRecords(29, 0))
                End If
                If VarType(ArrayRecords(30, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + "吨" + Trim(ArrayRecords(30, 0))
                End If

                .Col = 10
                .Row = 7
                If VarType(ArrayRecords(31, 0)) = vbNull Then
                    .Text = "0"
                Else
                    .Text = Trim(ArrayRecords(31, 0))
                End If
                If VarType(ArrayRecords(32, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + "吨" + Trim(ArrayRecords(32, 0)) + "月"
                End If

                .Col = 11
                .Row = 7
                If VarType(ArrayRecords(33, 0)) = vbNull Then
                    .Text = "0"
                Else
                    .Text = Trim(ArrayRecords(33, 0))
                End If
                If VarType(ArrayRecords(34, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = .Text + "吨" + Trim(ArrayRecords(34, 0)) + "月"
                End If

 
                .Col = 1
                .Row = 5
                If VarType(ArrayRecords(35, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(35, 0))
                End If
               
                .Col = 12
                .Row = 1
                If VarType(ArrayRecords(36, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(36, 0))
                End If

                
                .Col = 12
                .Row = 3
                If VarType(ArrayRecords(37, 0)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(37, 0))
                End If
                
               
               
                
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With
    
    MDIMain.StatusBar1.Panels(1) = "提示信息: 数据查询完成"
    Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Data_Refer = False
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("SpreadDisplay_Error : " & Error)
    
End Function
Public Function ss3_clear() As Boolean
    With SS3
                .Col = 1
                .Row = 3
                .Text = ""
                
                .Col = 2
                .Row = 3
                .Text = ""
                
                .Col = 3
                .Row = 3
                .Text = ""
                
                .Col = 6
                .Row = 1
                .Text = ""
                
                .Col = 6
                .Row = 2
                .Text = ""
                
                .Col = 6
                .Row = 3
                .Text = ""
                
                
                .Col = 6
                .Row = 4
                .Text = ""
                
                .Col = 8
                .Row = 1
                .Text = ""
                
                .Col = 8
                .Row = 2
                .Text = ""

                
                .Col = 8
                .Row = 3
                .Text = ""
               
                .Col = 8
                .Row = 4
                .Text = ""

              
                .Col = 10
                .Row = 1
                .Text = ""

               
                .Col = 10
                .Row = 2
                .Text = ""

              
                .Col = 10
                .Row = 3
                .Text = ""
                
                .Col = 10
                .Row = 4
                .Text = ""
                
                .Col = 12
                .Row = 1
                .Text = ""

                
                .Col = 12
                .Row = 3
                .Text = ""
                
                .Col = 1
                .Row = 5
                .Text = ""
               
                .Col = 2
                .Row = 6
                .Text = ""
                
                .Col = 4
                .Row = 6
                .Text = ""
                
                .Col = 6
                .Row = 6
                .Text = ""

                .Col = 8
                .Row = 6
                .Text = ""
                
                .Col = 10
                .Row = 6
                .Text = ""
                
                .Col = 12
                .Row = 6
                .Text = ""

                .Col = 2
                .Row = 7
                .Text = ""
                
                .Col = 4
                .Row = 7
                .Text = ""
                
                .Col = 5
                .Row = 7
                .Text = ""

                .Col = 7
                .Row = 7
                .Text = ""
                
                .Col = 8
                .Row = 7
                .Text = ""
                
                .Col = 10
                .Row = 7
                .Text = ""
                
                .Col = 11
                .Row = 7
                .Text = ""

    End With
End Function

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake1(Proc_Sc("Sc1")("Spread"), Mode)
        With ss1
             If .ActiveRow = 1 Then
             .Col = 22
             .Text = txt_DATE.RawData
             .Col = 23
             .Text = "A"
             End If
             
             If .ActiveRow = 3 Then
             .Col = 22
             .Text = txt_DATE.RawData
             .Col = 23
             .Text = "B"
             End If
             
             If .ActiveRow = 5 Then
             .Col = 22
             .Text = txt_DATE.RawData
             .Col = 23
             .Text = "C"
             End If
             
             If .ActiveRow = 7 Then
             .Col = 22
             .Text = txt_DATE.RawData
             .Col = 23
             .Text = "D"
             End If
             
        End With
    End If
End Sub

Private Sub ss2_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc2")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    

End Sub
Public Function Sp_Display(Conn As ADODB.Connection, sPname As Variant, sQuery As String) As Boolean

On Error GoTo SpreadDisplay_Error

    Dim iCount As Integer
    Dim iRowCount As Long
    Dim iColcount As Long
    Dim AdoRs As ADODB.Recordset
    Dim ArrayRecords As Variant

    Sp_Display = True
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Sp_Display = False: Exit Function
    End If
    
    Set AdoRs = New ADODB.Recordset
    
    With sPname

        .ReDraw = False
        iCount = 0
        
        '.ClearRange 1, 1, .MaxCols, .MaxRows, True
    
        Screen.MousePointer = vbHourglass
        
        'Ado Execute
        AdoRs.Open sQuery, Conn, adOpenKeyset
        
        If AdoRs.BOF Or AdoRs.EOF Then
            
            .ReDraw = True
            AdoRs.Close
            Set AdoRs = Nothing
            Sp_Display = False
            Call Gp_MsgBoxDisplay("无相关记录", "I")
            Call Form_Cls
            Screen.MousePointer = vbDefault
            Exit Function
            
        End If
        
        ArrayRecords = AdoRs.GetRows
        
        AdoRs.Close
        Set AdoRs = Nothing

        If UBound(ArrayRecords, 2) <> 0 Then
        
        For iRowCount = 0 To .MaxRows - 1
        
            .Row = iRowCount + 1
            
            For iColcount = 0 To .MaxCols - 1
            
                .Col = iColcount + 1
                    
                    If VarType(ArrayRecords(iColcount, iRowCount)) = vbNull Then
                        .Text = ""
                    Else
                        .Text = Trim(ArrayRecords(iColcount, iRowCount))
                    End If
                    
                Next iColcount
                
            Next iRowCount
            
        End If
        
        .ReDraw = True
        Screen.MousePointer = vbDefault
        
    End With

Exit Function

SpreadDisplay_Error:
    
    Set AdoRs = Nothing
    Sp_Display = False
    'Call Gp_MsgBoxDisplay("Query Failed..." & sQuery)
        Call Gp_MsgBoxDisplay("Gf_Sp_Refer Error : " & Error)

    Screen.MousePointer = vbDefault

End Function
Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

        If Gf_Sc_Authority(sAuthority, "U") Then
           Text_UPD.Text = "Update"
        End If

End Sub
Private Sub ss3_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc3")("Spread").MaxRows < 1 Then Exit Sub
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
End Sub

Public Sub Sp_ColLock(sPname As Variant, ColNum As Variant, RowNum As Variant, LockType As Boolean)

    With sPname
    
        .Protect = True
        .Col = ColNum: .Col2 = ColNum
        .Row = RowNum: .row2 = RowNum
        
        .BlockMode = True
        .Lock = LockType
        .BlockMode = False
    
    End With
    
End Sub

Public Function Sp_Process(Conn As ADODB.Connection, Sc As Collection, Optional RefChek As Boolean) As Boolean

On Error GoTo SpreadPro_Error

    If txt_DATE.RawData = "" Then
       Sp_Process = False
       Call Form_Ref
       Exit Function
    End If
    If Text_UPD.Text = "" Then
        Sp_Process = False
        Call Form_Ref
        Exit Function
    End If

    Dim iCol, iRow, iCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    
    Dim sMesg As String
    Dim sTemp As String
    
    Dim adoCmd As ADODB.Command

    Sp_Process = True
     
'    With ss1
'        For iRow = 1 To .MaxRows
'            For iCol = 2 To .MaxCols Step 2
'                .Row = iRow
'                .Col = iCol
'                If .Text = "" Then
'                   .Col = iCol - 1
'                   .Text = ""
'                End If
'            Next iCol
'        Next iRow
'    End With
    With SS3
    
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
        For iCount = 0 To 12
            adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
        Next iCount
        
        'Ceate Parameter (Output)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
        adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
        
        If Text_UPD.Text <> "" Then
          adoCmd.Parameters(0).Value = "U"
        End If
        .Col = 1
        .Row = 3
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(1).Value = Str(Trim(.Text))                'BOF_LIFE
        Else
            adoCmd.Parameters(1).Value = 0
        End If
        
        .Col = 2
        .Row = 3
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(2).Value = Str(Trim(.Text))                'plt
        Else
            adoCmd.Parameters(2).Value = 0
        End If
        
        .Col = 6
        .Row = 1
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(3).Value = Str(Trim(.Text))             'plt
        Else
            adoCmd.Parameters(3).Value = 0
        End If

        .Col = 6
        .Row = 2
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(4).Value = Str(Trim(.Text))                'plt
        Else
            adoCmd.Parameters(4).Value = 0
        End If
                                   
        .Col = 6
        .Row = 3
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(5).Value = Str(Trim(.Text))                'plt
        Else
            adoCmd.Parameters(5).Value = 0
        End If
        
        .Col = 6
        .Row = 4
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(6).Value = Str(Trim(.Text))                'plt
        Else
            adoCmd.Parameters(6).Value = 0
        End If
        
        .Col = 8
        .Row = 1
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(7).Value = Str(Trim(.Text))                'plt
        Else
            adoCmd.Parameters(7).Value = 0
        End If
        
        .Col = 10
        .Row = 1
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(8).Value = Str(Trim(.Text))                'plt
        Else
            adoCmd.Parameters(8).Value = 0
        End If
        
        
        .Col = 12
        .Row = 1
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(9).Value = Str(Trim(.Text))                'plt
        Else
            adoCmd.Parameters(9).Value = 0
        End If
        
        .Col = 12
        .Row = 3
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(10).Value = Str(Trim(.Text))                'plt
        Else
            adoCmd.Parameters(10).Value = 0
        End If
        
        .Col = 1
        .Row = 5
        If Trim(.Text) <> "" Then
            adoCmd.Parameters(11).Value = Trim(.Text)                'DESCRIPTION
        Else
            adoCmd.Parameters(11).Value = ""
        End If
        

        adoCmd.Parameters(12).Value = txt_DATE.RawData
        
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
                
        
        Conn.CommitTrans
        MDIMain.StatusBar1.Panels(1) = "提示信息: 数据处理完成"
        Text_UPD = ""
        Screen.MousePointer = vbDefault
        Call Form_Ref
    
    End With
    
    Exit Function

SpreadPro_Error:

    Screen.MousePointer = vbDefault
    Set adoCmd = Nothing
    Conn.RollbackTrans
    Sp_Process = False
    Call Gp_MsgBoxDisplay("SpreadPro_Error : " & Error)

End Function

Public Sub Gp_Sp_UpdateMake1(sPname As Variant, Mode As Integer)

    With sPname
    
        If .MaxRows < 1 Then Exit Sub
        
        .Col = .ActiveCol
        .Row = IIf(.ActiveRow > 0, .ActiveRow, 0)
        
        If Mode = 1 Then
            .Tag = .Text
        Else
            If Trim(.Tag) <> Trim(.Text) Then
                .Col = 24
                Select Case Trim(.Text)
                    Case "Input", "Update", "Delete"
                    Case Else
                        .Text = "Update"
                End Select
            End If
        End If
    
    End With

End Sub
Public Function Gf_Sp_Process1(Conn As ADODB.Connection, Sc As Collection, Optional Mc As Collection, _
                              Optional RefChek As Boolean = False) As Boolean

On Error GoTo SpreadPro_Error

    Dim iCol, iCount, iProcessCount As Integer
    Dim ret_Result_ErrCode As Integer
    Dim ret_Result_ErrMsg As String
    
    Dim dTempInt As Double
    Dim dTempFloat As Double
    
    Dim sMesg As String
    Dim sTemp As String
    Dim ProcessChk As String
    Dim DelYN As Boolean
    Dim Msg_Count As Integer
    Dim Msg_Yes As String
    
    Dim adoCmd As ADODB.Command

    Gf_Sp_Process1 = True
    iProcessCount = 0
    
    'MaxRow = 0 is Exit Function Or iCount = 0
    If Sc.Item("Spread").MaxRows < 1 Or Sc.Item("iColumn").Count = 0 Then
        Gf_Sp_Process1 = False
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    Sc.Item("Spread").ReDraw = False
    
    'NeceCheck
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 24, iCount))
            
            Case "Input", "Update"
            
                If Not Mc Is Nothing Then
                    Call Gp_Sp_Move(iCount, Sc, Mc)
                End If
                
                'Maxlength Check
                sMesg = Gf_Sp_NeceCheck2(Sc.Item("Spread"), Sc.Item("mColumn"), iCount, Sc.Item("nColumn"))
                        
                If Trim(sMesg) = "OK" Then
                    
                ElseIf Mid(sMesg, 1, 5) = "FALSE" Then
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = Mid(sMesg, 6, Len(sMesg))
                    sMesg = sMesg + "长度不正确"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process1 = False
                    Exit Function
                Else
                    Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                    sMesg = sMesg + "必须输入"
                    Call Gp_MsgBoxDisplay(sMesg)
                    Screen.MousePointer = vbDefault
                    Set adoCmd = Nothing
                    Gf_Sp_Process1 = False
                    Exit Function
                End If
        
        End Select
    
    Next iCount
    
    'Db Connection Check
    If Conn Is Nothing Then
        If GF_DbConnect = False Then Gf_Sp_Process1 = False: Exit Function
    End If
    
    'Ado Setting
    Conn.CursorLocation = adUseServer
    Set adoCmd = New ADODB.Command
    
    Set adoCmd.ActiveConnection = Conn
    adoCmd.CommandType = adCmdStoredProc
    adoCmd.CommandText = Sc.Item("P-M")
    
    Conn.BeginTrans
    
    'Create Parameter (Input) iType + iColumn
    For iCount = 0 To Sc.Item("iColumn").Count
        adoCmd.Parameters.Append adoCmd.CreateParameter("", adVariant, adParamInput)
    Next iCount
    
    'Create Parameter (Output)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Error", adVariant, adParamOutput)
    adoCmd.Parameters.Append adoCmd.CreateParameter("Messg", adVariant, adParamOutput)
    
    Msg_Count = 1
    For iCount = 1 To Sc.Item("Spread").MaxRows
        
        ProcessChk = "NO"
        DelYN = False
        
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 24, iCount))
        
            Case "Input"
                adoCmd.Parameters(0).Value = "I"
                ProcessChk = "YES"
                
            Case "Update"
                adoCmd.Parameters(0).Value = "U"
                ProcessChk = "YES"
                
            Case "Delete"
                adoCmd.Parameters(0).Value = "D"
                If Msg_Count = 1 Then
                   DelYN = Gf_MessConfirm("您确定要删除状态为[Delete]的数据吗？", "Q")
                   If DelYN Then Msg_Yes = "yes"
                   Msg_Count = Msg_Count + 1
                End If
                If Msg_Yes = "yes" Then DelYN = True
        End Select
          
        If ProcessChk = "YES" Or DelYN Then
            
            'Parameters Setting
            For iCol = 1 To Sc.Item("iColumn").Count
            
                Sc.Item("Spread").Col = Sc.Item("iColumn").Item(iCol)
                
                Select Case Sc.Item("Spread").CellType
                
                    Case SS_CELL_TYPE_CURRENCY
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempFloat = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempFloat)
                        End If
                        
                    Case SS_CELL_TYPE_NUMBER
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = 0
                        Else
                            dTempInt = Sc.Item("Spread").Text
                            adoCmd.Parameters(iCol).Value = Str(dTempInt)
                        End If
                        
                    Case SS_CELL_TYPE_CHECKBOX
                        If Sc.Item("Spread").Text = "1" Then
                            adoCmd.Parameters(iCol).Value = "1"
                        Else
                            adoCmd.Parameters(iCol).Value = "0"
                        End If
                        
                    Case SS_CELL_TYPE_COMBOBOX
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = "0"
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_PIC, SS_CELL_TYPE_TIME
                        If Trim(Sc.Item("Spread").Value) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Trim(Str(Sc.Item("Spread").Value))
                        End If
                        
                    Case SS_CELL_TYPE_DATE
                        If Trim(Sc.Item("Spread").Text) = "" Then
                            adoCmd.Parameters(iCol).Value = ""
                        Else
                            adoCmd.Parameters(iCol).Value = Mid(Trim(Sc.Item("Spread").Text), 1, 4) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 6, 2) & _
                                                            Mid(Trim(Sc.Item("Spread").Text), 9, 2)
                        End If
                       
                    Case Else
                        sTemp = Replace(Sc.Item("Spread").Text, "'", "''")
                        adoCmd.Parameters(iCol).Value = Trim(sTemp)
                        
                End Select
           
            Next iCol
                           
            iProcessCount = iProcessCount + 1
            adoCmd.Execute
            
            'Error Check
            If adoCmd("Error") <> "0" Then
            
                ret_Result_ErrCode = adoCmd("Error")
                ret_Result_ErrMsg = adoCmd("Messg")
        
                sErrMessg = "Error Code : " & ret_Result_ErrCode & vbCrLf & "Error Mesg : " & ret_Result_ErrMsg
                
                Call Gp_Sp_RowColor(Sc.Item("Spread"), iCount, , vbYellow)
                Call Gp_MsgBoxDisplay(sErrMessg)
                
                Screen.MousePointer = vbDefault
                Set adoCmd = Nothing
                
                Conn.RollbackTrans
                Gf_Sp_Process1 = False
                Exit Function
        
             End If
        
        End If
        
    Next iCount
    
    Conn.CommitTrans
    
    ' 0 Column Space
    For iCount = 1 To Sc.Item("Spread").MaxRows
    
        Select Case Trim(Gf_Sp_RcvData(Sc.Item("Spread"), 24, iCount))
        
            Case "Input", "Update"
                Call Gp_Sp_SendData(Sc.Item("Spread"), "", 24, iCount)
                
            Case "Delete"
                If DelYN Then
                   Call Gp_Sp_SendData(Sc.Item("Spread"), "", 24, iCount)
                   Call Gp_Sp_DeleteRow(Sc.Item("Spread"), iCount)
                   iCount = iCount - 1
                End If
        End Select
        
    Next iCount
    
    Sc.Item("Spread").ReDraw = True
    
'    If iProcessCount > 0 Then
'        If Not Mc Is Nothing Then
'            If RefChek = False Then Call Gf_Sp_Display(Conn, Sc.Item("Spread"), _
'                                                    Gf_Ms_MakeQuery(Sc.Item("P-R"), "R", Mc("pControl")), Sc.Item("pColumn"), False)
'
'        Else
'            If RefChek = False Then Call Gf_Sp_Display(Conn, Sc.Item("Spread"), _
'                           Gf_Sp_MakeQuery(Sc.Item("Spread"), Sc.Item("P-R"), "R", Sc.Item("aColumn"), 1), Sc.Item("pColumn"), False)
'        End If
'
'        MDIMain.StatusBar1.Panels(1) = "提示信息：成功处理了" & iProcessCount & "条记录"
'        'Call Gp_MsgBoxDisplay("Data that handle is " & iProcessCount & " items", "I")
'
'    End If
            
    If iProcessCount > 0 Then
        If Not Mc Is Nothing Then
            Call Gp_Ms_ControlLock(Mc.Item("lControl"), True)
        End If
        Call Form_Ref
    Else
        Gf_Sp_Process1 = False
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function

SpreadPro_Error:

    Set adoCmd = Nothing
    Conn.RollbackTrans
    Gf_Sp_Process1 = False
    Call Gp_MsgBoxDisplay("Gf_Sp_Process Error : " & Error)
    Screen.MousePointer = vbDefault

End Function

Private Sub SSCommand_CREAT_Click()
On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim adoCmd As ADODB.Command
    Dim Response As Variant

    Response = MsgBox("生成" + Mid(txt_DATE.RawData, 1, 4) + "年" + Mid(txt_DATE.RawData, 5, 2) + "月" + Mid(txt_DATE.RawData, 7, 2) + "日  " + "新报表吗?", vbYesNo, "系统提示信息")
    If Response = vbYes Then
      
            Screen.MousePointer = vbHourglass
            
            OutParam(1, 1) = "arg_e_msg"
            OutParam(1, 2) = adVarChar
            OutParam(1, 3) = adParamOutput
            OutParam(1, 4) = 256
             
            sQuery = "{call AFP1012P ('" + txt_DATE.RawData + "' ,?)}"
            
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
                   Call MsgBox("生成报表！", vbInformation, "系统提示信息")
                   Call Form_Ref
            End If
            
            Set adoCmd = Nothing
            Screen.MousePointer = vbDefault
    Else
            Call Form_Ref
    End If
    Exit Sub

Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    ERR.Raise ERR.Number, ERR.Description & sQuery
    
End Sub


Private Sub SSCommand_PRINT_Click()
    If Not GF_Print("中厚板卷厂生产简报.rpt", "") Then
       MsgBox "打印出错！", vbCritical, "系统提示信息"
       Exit Sub
    End If

End Sub
Public Function GF_Print(sReportFileName As String, DefaultDate As String) As Boolean
Dim psPrintDate, psPrintshift, psprintgroup As Variant

On Error GoTo PrintError:
    Dim adocTemp As New ADODB.Command
    psPrintDate = txt_DATE.RawData


    adocTemp.ActiveConnection = M_CN1
    
    adocTemp.CommandText = "Delete from Tab_Print_Param where REPORT_NAME='" + sReportFileName + "'"
    adocTemp.Execute
    
    adocTemp.CommandText = "Insert Into Tab_Print_Param(Report_Name,Param_Name1) Values('" + sReportFileName + "','" + psPrintDate + "')"
    adocTemp.Execute
    
    adocTemp.CommandText = "Commit"
    adocTemp.Execute
    
    Set Report = crxApplication.OpenReport(App.Path & "\" & sReportFileName, 1)

   
    Report.EnableParameterPrompting = False
    Report.Database.LogOnServerEx "ado", "ora9", "", "nisco", "nisco01", "", M_CN1.ConnectionString
    Report.Database.Verify
    Report.ReadRecords
    Call frmReport.form_init(Me)
    frmReport.Show
    
    GF_Print = True
         
    Set Report = Nothing
    Exit Function
PrintError:
     Set Report = Nothing

     Set adocTemp = Nothing
     GF_Print = False
End Function
Public Sub Gp_Sp_Collection1(sPname As Variant, Num As Integer, pcol As String, ncol As String, mcol As String, _
                                                               iCol As String, acol As String, lCol As String, _
                            pColumn As Collection, nColumn As Collection, mColumn As Collection, iColumn As Collection, _
                            aColumn As Collection, lColumn As Collection)
   
    If LCase(Trim(pcol)) = "p" Then       'PK Column
        pColumn.Add Item:=Num
    End If
    
    If LCase(Trim(ncol)) = "n" Then       'Necessary Column
        nColumn.Add Item:=Num
        'Call Gp_Sp_ColColor(SpName, Num, , &H80FF80)
    End If
    
    If LCase(Trim(mcol)) = "m" Then       'Spread Maxlength check Column
        mColumn.Add Item:=Num
    End If
    
    If LCase(Trim(iCol)) = "i" Then       'Spread Insert Column
        iColumn.Add Item:=Num
       ' Call Gp_Sp_ColColor(sPname, Num, , &HC0FFFF)
    End If
    
    If LCase(Trim(acol)) = "a" Then       'Master -> Spread Column
        aColumn.Add Item:=Num
        Call Gp_Sp_ColHidden(sPname, Num, True)
    End If
    
    If LCase(Trim(lCol)) = "l" Then       'Spread Lock Column
        lColumn.Add Item:=Num
        Call Gp_Sp_ColLock(sPname, Num, True)
    End If

    
End Sub
Public Function Gf_Sp_ProceExist1(sPname As Variant, Optional Tf As Boolean = True) As Boolean

    Dim sMessg As String
    Dim lCount As Long
    Dim Proc As Long

    With sPname
    
        Proc = 0
        
        For lCount = 1 To .MaxRows
            .Col = 24: .Row = lCount
            If Trim(.Text) = "Update" Then
                Proc = Proc + 1
                Exit For
            End If
        Next lCount
        
        If Proc > 0 Then
            If Tf Then
                sMessg = "表格中还有数据未处理，" + vbCrLf
                sMessg = sMessg + "放弃并继续吗？"
                
                If Gf_MessConfirm(sMessg, "Q") Then
                    Gf_Sp_ProceExist1 = False
                Else
                    Gf_Sp_ProceExist1 = True
                End If
                
            Else
                Gf_Sp_ProceExist1 = True
            End If
            
        Else
            Gf_Sp_ProceExist1 = False
        End If
    
        Exit Function
        
    End With

End Function
Public Function Sp_ProceExist(sPname As Variant, Optional Tf As Boolean = True) As Boolean

    Dim sMessg As String
    Dim lCount As Long
    Dim Proc As Long

    With sPname
    
        Proc = 0
        
    
        If Trim(Text_UPD.Text) = "Update" Then
            Proc = Proc + 1
        End If

        
        If Proc > 0 Then
            If Tf Then
                sMessg = "表格中还有数据未处理，" + vbCrLf
                sMessg = sMessg + "放弃并继续吗？"
                
                If Gf_MessConfirm(sMessg, "Q") Then
                    Sp_ProceExist = False
                Else
                    Sp_ProceExist = True
                End If
                
            Else
                Sp_ProceExist = True
            End If
            
        Else
            Sp_ProceExist = False
        End If
    
        Exit Function
        
    End With

End Function


