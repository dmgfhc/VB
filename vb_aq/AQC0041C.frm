VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AQC0041C 
   Caption         =   "PWHT����ʵ��_AQC0041C"
   ClientHeight    =   9330
   ClientLeft      =   195
   ClientTop       =   1140
   ClientWidth     =   15405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   15405
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd 
      Caption         =   "����ָʾ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13440
      MaskColor       =   &H00808080&
      TabIndex        =   9
      Top             =   60
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txt_smp_cut_loc 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_smp_no 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   15195
      _Version        =   393216
      _ExtentX        =   26802
      _ExtentY        =   2566
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
      MaxCols         =   34
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0041C.frx":0000
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   15195
      _Version        =   393216
      _ExtentX        =   26802
      _ExtentY        =   2566
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
      MaxCols         =   50
      MaxRows         =   1
      OperationMode   =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0041C.frx":0A54
   End
   Begin FPSpread.vaSpread ss3 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   7275
      _Version        =   393216
      _ExtentX        =   12832
      _ExtentY        =   2143
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
      MaxCols         =   6
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0041C.frx":21F3
   End
   Begin FPSpread.vaSpread ss4 
      Height          =   1215
      Left            =   7440
      TabIndex        =   5
      Top             =   4920
      Width           =   7875
      _Version        =   393216
      _ExtentX        =   13891
      _ExtentY        =   2143
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
      MaxCols         =   8
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0041C.frx":26A9
   End
   Begin FPSpread.vaSpread ss5 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   15195
      _Version        =   393216
      _ExtentX        =   26802
      _ExtentY        =   2143
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
      MaxCols         =   17
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0041C.frx":2B94
   End
   Begin FPSpread.vaSpread ss7 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   7800
      Width           =   15195
      _Version        =   393216
      _ExtentX        =   26802
      _ExtentY        =   2566
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
      MaxCols         =   37
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0041C.frx":3429
   End
   Begin FPSpread.vaSpread ss6 
      Height          =   1455
      Left            =   120
      TabIndex        =   8
      Top             =   6240
      Width           =   15195
      _Version        =   393216
      _ExtentX        =   26802
      _ExtentY        =   2566
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
      MaxCols         =   42
      MaxRows         =   1
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQC0041C.frx":3E3A
   End
End
Attribute VB_Name = "AQC0041C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       ��������
'-- Sub_System Name   ������׼����
'-- Program Name      ������Ƽ�����
'-- Program ID        AQC0041C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Lee Qing Yu
'-- Coder             Lee Qing Yu
'-- Date              2003.5.19
'-- Description       ������Ƽ�����
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

Dim pControl2 As New Collection      'Master Primary Key Collection
Dim nControl2 As New Collection      'Master Necessary Collection
Dim mControl2 As New Collection      'Master Maxlength check Collection
Dim iControl2 As New Collection      'Master Insert Collection
Dim rControl2 As New Collection      'Master Refer Collection
Dim cControl2 As New Collection      'Master Copy Collection
Dim aControl2 As New Collection      'Master -> Spread Collection
Dim lControl2 As New Collection      'Master Lock Collection

Dim pControl3 As New Collection      'Master Primary Key Collection
Dim nControl3 As New Collection      'Master Necessary Collection
Dim mControl3 As New Collection      'Master Maxlength check Collection
Dim iControl3 As New Collection      'Master Insert Collection
Dim rControl3 As New Collection      'Master Refer Collection
Dim cControl3 As New Collection      'Master Copy Collection
Dim aControl3 As New Collection      'Master -> Spread Collection
Dim lControl3 As New Collection      'Master Lock Collection

Dim pControl4 As New Collection      'Master Primary Key Collection
Dim nControl4 As New Collection      'Master Necessary Collection
Dim mControl4 As New Collection      'Master Maxlength check Collection
Dim iControl4 As New Collection      'Master Insert Collection
Dim rControl4 As New Collection      'Master Refer Collection
Dim cControl4 As New Collection      'Master Copy Collection
Dim aControl4 As New Collection      'Master -> Spread Collection
Dim lControl4 As New Collection      'Master Lock Collection

Dim pControl5 As New Collection      'Master Primary Key Collection
Dim nControl5 As New Collection      'Master Necessary Collection
Dim mControl5 As New Collection      'Master Maxlength check Collection
Dim iControl5 As New Collection      'Master Insert Collection
Dim rControl5 As New Collection      'Master Refer Collection
Dim cControl5 As New Collection      'Master Copy Collection
Dim aControl5 As New Collection      'Master -> Spread Collection
Dim lControl5 As New Collection      'Master Lock Collection

Dim pControl6 As New Collection      'Master Primary Key Collection
Dim nControl6 As New Collection      'Master Necessary Collection
Dim mControl6 As New Collection      'Master Maxlength check Collection
Dim iControl6 As New Collection      'Master Insert Collection
Dim rControl6 As New Collection      'Master Refer Collection
Dim cControl6 As New Collection      'Master Copy Collection
Dim aControl6 As New Collection      'Master -> Spread Collection
Dim lControl6 As New Collection      'Master Lock Collection

Dim pControl7 As New Collection      'Master Primary Key Collection
Dim nControl7 As New Collection      'Master Necessary Collection
Dim mControl7 As New Collection      'Master Maxlength check Collection
Dim iControl7 As New Collection      'Master Insert Collection
Dim rControl7 As New Collection      'Master Refer Collection
Dim cControl7 As New Collection      'Master Copy Collection
Dim aControl7 As New Collection      'Master -> Spread Collection
Dim lControl7 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim pColumn12 As New Collection      'Spread Primary Key Collection
Dim nColumn12 As New Collection      'Spread necessary Column Collection
Dim mColumn12 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn12 As New Collection      'Spread Insert Column Collection
Dim aColumn12 As New Collection      'Master -> Spread Column Collection
Dim lColumn12 As New Collection      'Spread Lock Column Collection

Dim pColumn13 As New Collection      'Spread Primary Key Collection
Dim nColumn13 As New Collection      'Spread necessary Column Collection
Dim mColumn13 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn13 As New Collection      'Spread Insert Column Collection
Dim aColumn13 As New Collection      'Master -> Spread Column Collection
Dim lColumn13 As New Collection      'Spread Lock Column Collection

Dim pColumn14 As New Collection      'Spread Primary Key Collection
Dim nColumn14 As New Collection      'Spread necessary Column Collection
Dim mColumn14 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn14 As New Collection      'Spread Insert Column Collection
Dim aColumn14 As New Collection      'Master -> Spread Column Collection
Dim lColumn14 As New Collection      'Spread Lock Column Collection

Dim pColumn15 As New Collection      'Spread Primary Key Collection
Dim nColumn15 As New Collection      'Spread necessary Column Collection
Dim mColumn15 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn15 As New Collection      'Spread Insert Column Collection
Dim aColumn15 As New Collection      'Master -> Spread Column Collection
Dim lColumn15 As New Collection      'Spread Lock Column Collection

Dim pColumn16 As New Collection      'Spread Primary Key Collection
Dim nColumn16 As New Collection      'Spread necessary Column Collection
Dim mColumn16 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn16 As New Collection      'Spread Insert Column Collection
Dim aColumn16 As New Collection      'Master -> Spread Column Collection
Dim lColumn16 As New Collection      'Spread Lock Column Collection

Dim pColumn17 As New Collection      'Spread Primary Key Collection
Dim nColumn17 As New Collection      'Spread necessary Column Collection
Dim mColumn17 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn17 As New Collection      'Spread Insert Column Collection
Dim aColumn17 As New Collection      'Master -> Spread Column Collection
Dim lColumn17 As New Collection      'Spread Lock Column Collection


Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
Dim Mc3 As New Collection           'Master Collection
Dim Mc4 As New Collection           'Master Collection
Dim Mc5 As New Collection           'Master Collection
Dim Mc6 As New Collection           'Master Collection
Dim Mc7 As New Collection           'Master Collection

Dim Sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection
Dim sc3 As New Collection
Dim sc4 As New Collection
Dim sc5 As New Collection
Dim sc6 As New Collection
Dim sc7 As New Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim COL_ARR1
Dim COL_ARR2
Dim COL_ARR3
Dim COL_ARR4
Dim COL_ARR5
Dim COL_ARR6
Dim COL_ARR7
Dim i As Integer




Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

       Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
  Call Gp_Ms_Collection(txt_SMP_CUT_LOC, "p", " ", " ", " ", "r", " ", "l", pControl, nControl, mControl, iControl, rControl, aControl, lControl)

    
    'MASTER Collection
    Mc1.Add Item:=pControl, Key:="pControl"
    Mc1.Add Item:=nControl, Key:="nControl"
    Mc1.Add Item:=mControl, Key:="mControl"
    Mc1.Add Item:=iControl, Key:="iControl"
    Mc1.Add Item:=rControl, Key:="rControl"
    Mc1.Add Item:=cControl, Key:="cControl"
    Mc1.Add Item:=aControl, Key:="aControl"
    Mc1.Add Item:=lControl, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(SS1, 1, "p", "n", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 2, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 3, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 4, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 5, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 6, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 7, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 8, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(SS1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 10, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 18, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 22, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 23, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 24, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 25, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 26, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 27, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 28, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 29, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 30, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 31, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 32, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 33, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 34, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 35, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 36, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 37, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 38, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 39, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(SS1, 40, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)

    
    'Spread_Collection
    Sc1.Add Item:=SS1, Key:="Spread"
    'Sc1.Add Item:="AQC0041C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:="AQC0041C.P_REFER", Key:="P-R"
    'Sc1.Add Item:="AQC0041C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=SS1.MaxCols, Key:="Last"
    
    
    
       Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  Call Gp_Ms_Collection(txt_SMP_CUT_LOC, "p", " ", " ", " ", "r", " ", "l", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
  
    Mc2.Add Item:=pControl2, Key:="pControl"
    Mc2.Add Item:=nControl2, Key:="nControl"
    Mc2.Add Item:=mControl2, Key:="mControl"
    Mc2.Add Item:=iControl2, Key:="iControl"
    Mc2.Add Item:=rControl2, Key:="rControl"
    Mc2.Add Item:=cControl2, Key:="cControl"
    Mc2.Add Item:=aControl2, Key:="aControl"
    Mc2.Add Item:=lControl2, Key:="lControl"
    
    
    Call Gp_Sp_Collection(SS2, 1, "p", "n", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS2, 2, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS2, 3, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS2, 4, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS2, 5, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS2, 6, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS2, 7, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS2, 8, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
    Call Gp_Sp_Collection(SS2, 9, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 10, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 11, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 12, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 13, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 14, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 15, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 16, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 17, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 18, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 19, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 20, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 21, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 22, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 23, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 24, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 25, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 26, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 27, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 28, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 29, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 30, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 31, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 32, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 33, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
   Call Gp_Sp_Collection(SS2, 34, " ", " ", " ", " ", " ", " ", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
'   Call Gp_Sp_Collection(ss2, 35, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
'   Call Gp_Sp_Collection(ss2, 36, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
'   Call Gp_Sp_Collection(ss2, 37, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
'   Call Gp_Sp_Collection(ss2, 38, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
'   Call Gp_Sp_Collection(ss2, 39, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
'   Call Gp_Sp_Collection(ss2, 40, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
'   Call Gp_Sp_Collection(ss2, 41, " ", " ", " ", " ", " ", "l", pColumn12, nColumn12, mColumn12, iColumn12, aColumn12, lColumn12)
     
     'Spread_Collection
    sc2.Add Item:=SS2, Key:="Spread"
    sc2.Add Item:="AQC0041C.P_REFER_2", Key:="P-R"
    sc2.Add Item:=pColumn12, Key:="pColumn"
    sc2.Add Item:=nColumn12, Key:="nColumn"
    sc2.Add Item:=aColumn12, Key:="aColumn"
    sc2.Add Item:=mColumn12, Key:="mColumn"
    sc2.Add Item:=iColumn12, Key:="iColumn"
    sc2.Add Item:=lColumn12, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=SS2.MaxCols, Key:="Last"
    
    
    
       Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
  Call Gp_Ms_Collection(txt_SMP_CUT_LOC, "p", " ", " ", " ", "r", " ", "l", pControl3, nControl3, mControl3, iControl3, rControl3, aControl3, lControl3)
  
    Mc3.Add Item:=pControl3, Key:="pControl"
    Mc3.Add Item:=nControl3, Key:="nControl"
    Mc3.Add Item:=mControl3, Key:="mControl"
    Mc3.Add Item:=iControl3, Key:="iControl"
    Mc3.Add Item:=rControl3, Key:="rControl"
    Mc3.Add Item:=cControl3, Key:="cControl"
    Mc3.Add Item:=aControl3, Key:="aControl"
    Mc3.Add Item:=lControl3, Key:="lControl"
    
    Call Gp_Sp_Collection(ss3, 1, "p", "n", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
    Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
    Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
    Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
    Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
    Call Gp_Sp_Collection(ss3, 6, " ", " ", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
    
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AQC0041C.P_REFER_3", Key:="P-R"
    sc3.Add Item:=pColumn13, Key:="pColumn"
    sc3.Add Item:=nColumn13, Key:="nColumn"
    sc3.Add Item:=aColumn13, Key:="aColumn"
    sc3.Add Item:=mColumn13, Key:="mColumn"
    sc3.Add Item:=iColumn13, Key:="iColumn"
    sc3.Add Item:=lColumn13, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"
  
         Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
  Call Gp_Ms_Collection(txt_SMP_CUT_LOC, "p", " ", " ", " ", "r", " ", "l", pControl4, nControl4, mControl4, iControl4, rControl4, aControl4, lControl4)
  
    Mc4.Add Item:=pControl4, Key:="pControl"
    Mc4.Add Item:=nControl4, Key:="nControl"
    Mc4.Add Item:=mControl4, Key:="mControl"
    Mc4.Add Item:=iControl4, Key:="iControl"
    Mc4.Add Item:=rControl4, Key:="rControl"
    Mc4.Add Item:=cControl4, Key:="cControl"
    Mc4.Add Item:=aControl4, Key:="aControl"
    Mc4.Add Item:=lControl4, Key:="lControl"
    
    Call Gp_Sp_Collection(ss4, 1, "p", "n", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 2, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 3, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 4, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 5, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 6, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    Call Gp_Sp_Collection(ss4, 7, " ", " ", " ", " ", " ", " ", pColumn14, nColumn14, mColumn14, iColumn14, aColumn14, lColumn14)
    
    sc4.Add Item:=ss4, Key:="Spread"
    sc4.Add Item:="AQC0041C.P_REFER_4", Key:="P-R"
    sc4.Add Item:=pColumn14, Key:="pColumn"
    sc4.Add Item:=nColumn14, Key:="nColumn"
    sc4.Add Item:=aColumn14, Key:="aColumn"
    sc4.Add Item:=mColumn14, Key:="mColumn"
    sc4.Add Item:=iColumn14, Key:="iColumn"
    sc4.Add Item:=lColumn14, Key:="lColumn"
    sc4.Add Item:=1, Key:="First"
    sc4.Add Item:=ss4.MaxCols, Key:="Last"
  
         Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", "l", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
  Call Gp_Ms_Collection(txt_SMP_CUT_LOC, "p", " ", " ", " ", "r", " ", "l", pControl5, nControl5, mControl5, iControl5, rControl5, aControl5, lControl5)
  
    Mc5.Add Item:=pControl5, Key:="pControl"
    Mc5.Add Item:=nControl5, Key:="nControl"
    Mc5.Add Item:=mControl5, Key:="mControl"
    Mc5.Add Item:=iControl5, Key:="iControl"
    Mc5.Add Item:=rControl5, Key:="rControl"
    Mc5.Add Item:=cControl5, Key:="cControl"
    Mc5.Add Item:=aControl5, Key:="aControl"
    Mc5.Add Item:=lControl5, Key:="lControl"
    
    Call Gp_Sp_Collection(ss5, 1, "p", "n", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
    Call Gp_Sp_Collection(ss5, 2, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
    Call Gp_Sp_Collection(ss5, 3, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
    Call Gp_Sp_Collection(ss5, 4, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
    Call Gp_Sp_Collection(ss5, 5, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
    Call Gp_Sp_Collection(ss5, 6, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
    Call Gp_Sp_Collection(ss5, 7, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
    Call Gp_Sp_Collection(ss5, 8, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
    Call Gp_Sp_Collection(ss5, 9, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
   Call Gp_Sp_Collection(ss5, 10, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
   Call Gp_Sp_Collection(ss5, 11, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
   Call Gp_Sp_Collection(ss5, 12, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
   Call Gp_Sp_Collection(ss5, 13, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
   Call Gp_Sp_Collection(ss5, 14, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
   Call Gp_Sp_Collection(ss5, 15, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
   Call Gp_Sp_Collection(ss5, 16, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
   Call Gp_Sp_Collection(ss5, 17, " ", " ", " ", " ", " ", " ", pColumn15, nColumn15, mColumn15, iColumn15, aColumn15, lColumn15)
    
    sc5.Add Item:=ss5, Key:="Spread"
    sc5.Add Item:="AQC0041C.P_REFER_5", Key:="P-R"
    sc5.Add Item:=pColumn15, Key:="pColumn"
    sc5.Add Item:=nColumn15, Key:="nColumn"
    sc5.Add Item:=aColumn15, Key:="aColumn"
    sc5.Add Item:=mColumn15, Key:="mColumn"
    sc5.Add Item:=iColumn15, Key:="iColumn"
    sc5.Add Item:=lColumn15, Key:="lColumn"
    sc5.Add Item:=1, Key:="First"
    sc5.Add Item:=ss5.MaxCols, Key:="Last"
  
         Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
  Call Gp_Ms_Collection(txt_SMP_CUT_LOC, "p", " ", " ", " ", "r", " ", "l", pControl6, nControl6, mControl6, iControl6, rControl6, aControl6, lControl6)
  
    Mc6.Add Item:=pControl6, Key:="pControl"
    Mc6.Add Item:=nControl6, Key:="nControl"
    Mc6.Add Item:=mControl6, Key:="mControl"
    Mc6.Add Item:=iControl6, Key:="iControl"
    Mc6.Add Item:=rControl6, Key:="rControl"
    Mc6.Add Item:=cControl6, Key:="cControl"
    Mc6.Add Item:=aControl6, Key:="aControl"
    Mc6.Add Item:=lControl6, Key:="lControl"
    
    Call Gp_Sp_Collection(ss6, 1, "p", "n", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
    Call Gp_Sp_Collection(ss6, 2, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
    Call Gp_Sp_Collection(ss6, 3, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
    Call Gp_Sp_Collection(ss6, 4, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
    Call Gp_Sp_Collection(ss6, 5, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
    Call Gp_Sp_Collection(ss6, 6, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
    Call Gp_Sp_Collection(ss6, 7, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
    Call Gp_Sp_Collection(ss6, 8, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
    Call Gp_Sp_Collection(ss6, 9, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 10, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 11, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 12, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 13, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 14, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 15, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 16, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 17, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 18, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 19, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 20, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 21, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 22, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 23, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 24, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 25, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 26, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 27, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 28, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 29, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 30, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 31, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 32, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 33, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 34, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 35, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 36, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 37, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 38, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 39, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 40, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 41, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
   Call Gp_Sp_Collection(ss6, 42, " ", " ", " ", " ", " ", " ", pColumn16, nColumn16, mColumn16, iColumn16, aColumn16, lColumn16)
    
    sc6.Add Item:=ss6, Key:="Spread"
    sc6.Add Item:="AQC0041C.P_REFER_6", Key:="P-R"
    sc6.Add Item:=pColumn16, Key:="pColumn"
    sc6.Add Item:=nColumn16, Key:="nColumn"
    sc6.Add Item:=aColumn16, Key:="aColumn"
    sc6.Add Item:=mColumn16, Key:="mColumn"
    sc6.Add Item:=iColumn16, Key:="iColumn"
    sc6.Add Item:=lColumn16, Key:="lColumn"
    sc6.Add Item:=1, Key:="First"
    sc6.Add Item:=ss6.MaxCols, Key:="Last"
  
         Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", "r", " ", "l", pControl7, nControl7, mControl7, iControl7, rControl7, aControl7, lControl7)
  Call Gp_Ms_Collection(txt_SMP_CUT_LOC, "p", " ", " ", " ", "r", " ", "l", pControl7, nControl7, mControl7, iControl7, rControl7, aControl7, lControl7)
  
    Mc7.Add Item:=pControl7, Key:="pControl"
    Mc7.Add Item:=nControl7, Key:="nControl"
    Mc7.Add Item:=mControl7, Key:="mControl"
    Mc7.Add Item:=iControl7, Key:="iControl"
    Mc7.Add Item:=rControl7, Key:="rControl"
    Mc7.Add Item:=cControl7, Key:="cControl"
    Mc7.Add Item:=aControl7, Key:="aControl"
    Mc7.Add Item:=lControl7, Key:="lControl"
    
    Call Gp_Sp_Collection(ss7, 1, "p", "n", " ", " ", " ", " ", pColumn17, nColumn17, mColumn17, iColumn17, aColumn17, lColumn17)
    Call Gp_Sp_Collection(ss7, 2, " ", " ", " ", " ", " ", " ", pColumn17, nColumn17, mColumn17, iColumn17, aColumn17, lColumn17)
    Call Gp_Sp_Collection(ss7, 3, " ", " ", " ", " ", " ", " ", pColumn17, nColumn17, mColumn17, iColumn17, aColumn17, lColumn17)
    Call Gp_Sp_Collection(ss7, 4, " ", " ", " ", " ", " ", " ", pColumn17, nColumn17, mColumn17, iColumn17, aColumn17, lColumn17)
    Call Gp_Sp_Collection(ss7, 5, " ", " ", " ", " ", " ", " ", pColumn17, nColumn17, mColumn17, iColumn17, aColumn17, lColumn17)
    Call Gp_Sp_Collection(ss7, 6, " ", " ", " ", " ", " ", " ", pColumn17, nColumn17, mColumn17, iColumn17, aColumn17, lColumn17)
    
    sc7.Add Item:=ss7, Key:="Spread"
    sc7.Add Item:="AQC0041C.P_REFER_7", Key:="P-R"
    sc7.Add Item:=pColumn17, Key:="pColumn"
    sc7.Add Item:=nColumn17, Key:="nColumn"
    sc7.Add Item:=aColumn17, Key:="aColumn"
    sc7.Add Item:=mColumn17, Key:="mColumn"
    sc7.Add Item:=iColumn17, Key:="iColumn"
    sc7.Add Item:=lColumn17, Key:="lColumn"
    sc7.Add Item:=1, Key:="First"
    sc7.Add Item:=ss7.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="sc1"
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    Proc_Sc.Add Item:=sc3, Key:="Sc3"
    Proc_Sc.Add Item:=sc4, Key:="Sc4"
    Proc_Sc.Add Item:=sc5, Key:="Sc5"
    Proc_Sc.Add Item:=sc6, Key:="Sc6"
    Proc_Sc.Add Item:=sc7, Key:="Sc7"
    
    'Call Gp_Sp_BlockColor(ss1, 1, 1, 1, 1, , &HFFFF&)
    'Call Gp_Sp_BlockColor(ss1, 5, 5, 1, ss1.MaxRows, , &HFFFF&)
    'Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, iRow, iRow, , &HFFFF&)
    'Call Gp_Sp_CellColor(ss1, 5, 5, 1, ss1.MaxRows, RED)
    'Call Gp_Sp_CellColor(ss1, 1, 2, RED, RED)
    
    COL_ARR1 = Array(7, 11, 15, 22, 26, 30, 37, 40, 45, 49)
    COL_ARR2 = Array(14, 24, 34)
    COL_ARR5 = Array(7, 11, 15)
    COL_ARR6 = Array(8, 12, 16, 20, 24, 28, 32, 36, 39, 42)
    
    
    
    For i = 0 To 9
      Call Gp_Sp_ColHidden(SS1, COL_ARR1(i), True)
    Next i
    
    For i = 0 To 2
      Call Gp_Sp_ColHidden(SS2, COL_ARR2(i), True)
    Next i
    
    For i = 0 To 2
      Call Gp_Sp_ColHidden(ss5, COL_ARR5(i), True)
    Next i
    
    For i = 0 To 9
      Call Gp_Sp_ColHidden(ss6, COL_ARR6(i), True)
    Next i
    
    Call Gp_Sp_ColHidden(ss4, 8, True)
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    
End Sub



Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

 '   Call Cobox_Item_Add("enduse_cd", "qp_ord_usage", " ", Cob_ENDUSE_CD)


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
    'sAuthority = "1111"
     
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("pControl"))
    Call Gp_Ms_Cls(Mc3("pControl"))
    Call Gp_Ms_Cls(Mc4("rControl"))
    Call Gp_Ms_Cls(Mc5("pControl"))
    Call Gp_Ms_Cls(Mc6("pControl"))
    Call Gp_Ms_Cls(Mc7("rControl"))

    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc3")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc4")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc5")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc6")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc7")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc1"))
    Call Gf_Sp_Cls(Proc_Sc("Sc2"))
    Call Gf_Sp_Cls(Proc_Sc("Sc3"))
    Call Gf_Sp_Cls(Proc_Sc("Sc4"))
    Call Gf_Sp_Cls(Proc_Sc("Sc5"))
    Call Gf_Sp_Cls(Proc_Sc("Sc6"))
    Call Gf_Sp_Cls(Proc_Sc("Sc7"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc3")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc4")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc5")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc6")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc7")("Spread"), "Q-System.INI", Me.Name)
    
'    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 1)
'
'    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 3)
    
    Screen.MousePointer = vbDefault
    'Call Gp_Sp_BlockColor(ss1, 2, 2, 1, ss1.MaxRows, , &HC0E0FF)


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc1")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    
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
    
    Set Mc1 = Nothing
    Set Sc1 = Nothing
    
    Call Gp_Sp_ColSet(Proc_Sc("sc2")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl2 = Nothing
    Set nControl2 = Nothing
    Set iControl2 = Nothing
    Set rControl2 = Nothing
    Set cControl2 = Nothing
    Set aControl2 = Nothing
    Set lControl2 = Nothing
    Set mControl2 = Nothing
    
    Set iColumn12 = Nothing
    Set pColumn12 = Nothing
    Set lColumn12 = Nothing
    Set nColumn12 = Nothing
    Set mColumn12 = Nothing
    Set aColumn12 = Nothing
    
    Set Mc2 = Nothing
    Set sc2 = Nothing
    
    Call Gp_Sp_ColSet(Proc_Sc("sc3")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl3 = Nothing
    Set nControl3 = Nothing
    Set iControl3 = Nothing
    Set rControl3 = Nothing
    Set cControl3 = Nothing
    Set aControl3 = Nothing
    Set lControl3 = Nothing
    Set mControl3 = Nothing
    
    Set iColumn13 = Nothing
    Set pColumn13 = Nothing
    Set lColumn13 = Nothing
    Set nColumn13 = Nothing
    Set mColumn13 = Nothing
    Set aColumn13 = Nothing
    
    Set Mc3 = Nothing
    Set sc3 = Nothing
    
    Call Gp_Sp_ColSet(Proc_Sc("sc4")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl4 = Nothing
    Set nControl4 = Nothing
    Set iControl4 = Nothing
    Set rControl4 = Nothing
    Set cControl4 = Nothing
    Set aControl4 = Nothing
    Set lControl4 = Nothing
    Set mControl4 = Nothing
    
    Set iColumn14 = Nothing
    Set pColumn14 = Nothing
    Set lColumn14 = Nothing
    Set nColumn14 = Nothing
    Set mColumn14 = Nothing
    Set aColumn14 = Nothing
    
    Set Mc4 = Nothing
    Set sc4 = Nothing
    
    Call Gp_Sp_ColSet(Proc_Sc("sc5")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl5 = Nothing
    Set nControl5 = Nothing
    Set iControl5 = Nothing
    Set rControl5 = Nothing
    Set cControl5 = Nothing
    Set aControl5 = Nothing
    Set lControl5 = Nothing
    Set mControl5 = Nothing
    
    Set iColumn15 = Nothing
    Set pColumn15 = Nothing
    Set lColumn15 = Nothing
    Set nColumn15 = Nothing
    Set mColumn15 = Nothing
    Set aColumn15 = Nothing
    
    Set Mc5 = Nothing
    Set sc5 = Nothing
    
    Call Gp_Sp_ColSet(Proc_Sc("sc6")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl6 = Nothing
    Set nControl6 = Nothing
    Set iControl6 = Nothing
    Set rControl6 = Nothing
    Set cControl6 = Nothing
    Set aControl6 = Nothing
    Set lControl6 = Nothing
    Set mControl6 = Nothing
    
    Set iColumn16 = Nothing
    Set pColumn16 = Nothing
    Set lColumn16 = Nothing
    Set nColumn16 = Nothing
    Set mColumn16 = Nothing
    Set aColumn16 = Nothing
    
    Set Mc6 = Nothing
    Set sc6 = Nothing
    
    Call Gp_Sp_ColSet(Proc_Sc("sc7")("Spread"), "Q-System.INI", Me.Name)
    
    Set pControl7 = Nothing
    Set nControl7 = Nothing
    Set iControl7 = Nothing
    Set rControl7 = Nothing
    Set cControl7 = Nothing
    Set aControl7 = Nothing
    Set lControl7 = Nothing
    Set mControl7 = Nothing
    
    Set iColumn17 = Nothing
    Set pColumn17 = Nothing
    Set lColumn17 = Nothing
    Set nColumn17 = Nothing
    Set mColumn17 = Nothing
    Set aColumn17 = Nothing
    
    Set Mc7 = Nothing
    Set sc7 = Nothing
    
    
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))
      
End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        rControl(1).SetFocus
    End If
 

 
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    Dim i As Integer
    Dim iRow As Integer

    
    If Gf_Sp_ProceExist(Proc_Sc("Sc1").Item("Spread")) Then Exit Sub
    If Gf_Sp_ProceExist(Proc_Sc("Sc2").Item("Spread")) Then Exit Sub
    If Gf_Sp_ProceExist(Proc_Sc("Sc3").Item("Spread")) Then Exit Sub
    If Gf_Sp_ProceExist(Proc_Sc("Sc4").Item("Spread")) Then Exit Sub
    If Gf_Sp_ProceExist(Proc_Sc("Sc5").Item("Spread")) Then Exit Sub
    If Gf_Sp_ProceExist(Proc_Sc("Sc6").Item("Spread")) Then Exit Sub
    If Gf_Sp_ProceExist(Proc_Sc("Sc7").Item("Spread")) Then Exit Sub
    
'    sMesg = Gf_Ms_NeceCheck(nControl)
'    If sMesg = "OK" Then
'
'        sMesg = Gf_Ms_NeceCheck2(mControl)
'        If sMesg = "OK" Then


        
'            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
'                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
'
'                'Exit Sub
'            End If
            
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc1, Mc1("nControl"), Mc1("mControl"), False)
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc2"), Mc2, Mc2("nControl"), Mc2("mControl"), False)
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc3"), Mc3, Mc3("nControl"), Mc3("mControl"), False)
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc4"), Mc4, Mc4("nControl"), Mc4("mControl"), False)
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc5"), Mc5, Mc5("nControl"), Mc5("mControl"), False)
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc6"), Mc6, Mc6("nControl"), Mc6("mControl"), False)
        Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc7"), Mc7, Mc7("nControl"), Mc7("mControl"), False)
        
        SS1.OperationMode = OperationModeNormal
        SS2.OperationMode = OperationModeNormal
        ss3.OperationMode = OperationModeNormal
        ss4.OperationMode = OperationModeNormal
        ss5.OperationMode = OperationModeNormal
        ss6.OperationMode = OperationModeNormal
        ss7.OperationMode = OperationModeNormal
        
'        Else
'            sMesg = sMesg + " Must input according to length of item"
'            Call Gp_MsgBoxDisplay(sMesg)
'        End If
'
'    Else
'        sMesg = sMesg + " Must input necessarily"
'        Call Gp_MsgBoxDisplay(sMesg)
'
'    End If

      With SS1
        For iRow = 1 To .MaxRows
          .Row = iRow
          For i = 0 To 9
            .Col = COL_ARR1(i)
            If .Value = 1 Then
              Call Gp_Sp_BlockColor(SS1, COL_ARR1(i) - 1, COL_ARR1(i) - 1, iRow, iRow, , RED)
            Else
              Call Gp_Sp_BlockColor(SS1, COL_ARR1(i) - 1, COL_ARR1(i) - 1, iRow, iRow, , &HD0D0D0)
            End If
          Next i
        Next iRow
      End With
      
      With SS2
        For iRow = 1 To .MaxRows
          .Row = iRow
          For i = 0 To 2
            .Col = COL_ARR2(i)
            If .Value = 1 Then
              Call Gp_Sp_BlockColor(SS2, COL_ARR2(i) - 1, COL_ARR2(i) - 1, iRow, iRow, , RED)
            Else
              Call Gp_Sp_BlockColor(SS2, COL_ARR2(i) - 1, COL_ARR2(i) - 1, iRow, iRow, , &HD0D0D0)
            End If
          Next i
        Next iRow
      End With
      
      With ss5
        For iRow = 1 To .MaxRows
          .Row = iRow
          For i = 0 To 2
            .Col = COL_ARR5(i)
            If .Value = 1 Then
              Call Gp_Sp_BlockColor(ss5, COL_ARR5(i) - 1, COL_ARR5(i) - 1, iRow, iRow, , RED)
            Else
              Call Gp_Sp_BlockColor(ss5, COL_ARR5(i) - 1, COL_ARR5(i) - 1, iRow, iRow, , &HD0D0D0)
            End If
          Next i
        Next iRow
      End With
      
      With ss3
        For iRow = 1 To .MaxRows
          .Row = iRow
            .Col = 6
            If .Value = "N" Then
              Call Gp_Sp_BlockColor(ss3, 6, 6, iRow, iRow, , RED)
            ElseIf .Value = "Y" Then
              Call Gp_Sp_BlockColor(ss3, 6, 6, iRow, iRow, , &HD0D0D0)
            End If
        Next iRow
      End With
      
      With ss4
        For iRow = 1 To .MaxRows
          .Row = iRow
            .Col = 8
            If .Value = 1 Then
              Call Gp_Sp_BlockColor(ss4, 7, 7, iRow, iRow, , RED)
            Else
              Call Gp_Sp_BlockColor(ss4, 7, 7, iRow, iRow, , &HD0D0D0)
            End If
        Next iRow
      End With
      
      With ss6
        For iRow = 1 To .MaxRows
          .Row = iRow
          For i = 0 To 9
            .Col = COL_ARR6(i)
            If .Value = 1 Then
              Call Gp_Sp_BlockColor(ss6, COL_ARR6(i) - 1, COL_ARR6(i) - 1, iRow, iRow, , RED)
            Else
              Call Gp_Sp_BlockColor(ss6, COL_ARR6(i) - 1, COL_ARR6(i) - 1, iRow, iRow, , &HD0D0D0)
            End If
          Next i
        Next iRow
      End With

    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
     
 
         If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
         End If


    
End Sub

Public Sub Form_Ins()
    
    Call Gp_Sp_Ins(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 10)
    'ss1.SetFocus
    

End Sub

Public Sub Spread_Cpy()

    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 6)
    
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



Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)

End Sub




