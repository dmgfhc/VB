VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0092C 
   Caption         =   "���ճɷ�ʵ����ѯ����_AQC0092C"
   ClientHeight    =   9210
   ClientLeft      =   855
   ClientTop       =   1935
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   14955
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_charge_no 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12510
      TabIndex        =   10
      Top             =   150
      Visible         =   0   'False
      Width           =   1185
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   8580
      Left            =   120
      TabIndex        =   7
      Top             =   540
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   15134
      _Version        =   196609
      PaneTree        =   "AQC0092C.frx":0000
      Begin FPSpread.vaSpread ss1 
         Height          =   5055
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   14895
         _Version        =   393216
         _ExtentX        =   26273
         _ExtentY        =   8916
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
         MaxCols         =   61
         MaxRows         =   1
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "AQC0092C.frx":0052
      End
      Begin FPSpread.vaSpread ss2 
         Height          =   3375
         Left            =   30
         TabIndex        =   9
         Top             =   5175
         Width           =   14895
         _Version        =   393216
         _ExtentX        =   26273
         _ExtentY        =   5953
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
         SpreadDesigner  =   "AQC0092C.frx":105E
      End
   End
   Begin VB.TextBox txt_to_heat 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10455
      MaxLength       =   8
      TabIndex        =   5
      Top             =   120
      Width           =   1005
   End
   Begin VB.TextBox txt_from_heat 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8895
      MaxLength       =   8
      TabIndex        =   4
      Top             =   120
      Width           =   1020
   End
   Begin VB.ComboBox cbo_LINE_NO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AQC0092C.frx":3FF2
      Left            =   6210
      List            =   "AQC0092C.frx":3FFF
      TabIndex        =   1
      Tag             =   "����"
      Top             =   120
      Width           =   600
   End
   Begin InDate.UDate txt_Charge_Date_Fr 
      Height          =   300
      Left            =   1155
      TabIndex        =   0
      Tag             =   "��������"
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "��������"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   1
      Left            =   5175
      Top             =   120
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "��  ��"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.UDate txt_Charge_Date_to 
      Height          =   300
      Left            =   2850
      TabIndex        =   2
      Tag             =   "��������"
      Top             =   120
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Index           =   0
      Left            =   7710
      Top             =   120
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Caption         =   "¯��"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
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
      Left            =   10020
      TabIndex        =   6
      Top             =   165
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
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
      Left            =   2655
      TabIndex        =   3
      Top             =   165
      Width           =   255
   End
End
Attribute VB_Name = "AQC0092C"
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
'-- Program ID        AQC0092C
'-- Document No
'-- Designer          HJD
'-- Coder             HJD
'-- Date              2006.11.01
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
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim str_STLGRD_DETAIL As String
Dim str_STLGRD As String
Dim lngActiveRow As Long

Private Sub Form_Define()
       
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Refer"              'form����
     
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       
      Call Gp_Ms_Collection(txt_charge_no, "p", " ", " ", " ", " ", " ", " ", pControl1, nControl1, mControl1, iControl1, rControl1, aControl1, lControl1)

                
    'MASTER Collection
     'Mc1.Add Item:="AQC0092C.P_MODIFY", Key:="P-M"
     'Mc1.Add Item:="AQC0092C.P_REFER", Key:="P-R"
     Mc1.Add Item:=pControl1, Key:="pControl"
     Mc1.Add Item:=nControl1, Key:="nControl"
     Mc1.Add Item:=mControl1, Key:="mControl"
     Mc1.Add Item:=iControl1, Key:="iControl"
     Mc1.Add Item:=rControl1, Key:="rControl"
     Mc1.Add Item:=cControl1, Key:="cControl"
     Mc1.Add Item:=aControl1, Key:="aControl"
     Mc1.Add Item:=lControl1, Key:="lControl"

       
     Call Gp_Ms_Collection(txt_Charge_Date_Fr, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
     Call Gp_Ms_Collection(txt_Charge_Date_to, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(cbo_LINE_NO, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
          Call Gp_Ms_Collection(txt_from_heat, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
            Call Gp_Ms_Collection(txt_to_heat, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

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
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
   'louyannan 20101109 start
   Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 50, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 51, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 52, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 53, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 54, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 55, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 56, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 57, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 58, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 59, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 60, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 61, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    'louyannan 20101109 end
  
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQC0092C.P_SREFER", Key:="P-R"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxRows, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc1"
    
    
    
    Sc1.Item("Spread").Col = 0
    Sc1.Item("Spread").Row = 0
    Sc1.Item("Spread").Text = "��"
    Call Gp_Sp_ColHidden(ss1, 60, True)
    Call Gp_Sp_ColHidden(ss1, 61, True)

    
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AQC0092C.P_SREFER2", Key:="P-R"
'    sc2.Add Item:="AFK2030C.P_REFER", Key:="P-R"
    
    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    lngActiveRow = 0
    str_STLGRD_DETAIL = "��׼"
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
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

    Call Form_Define
    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    Call Gp_Ms_Cls(Mc2("rControl"))
    
    Call Gp_Ms_ControlLock(Mc1("lControl"), True)
    Call Gp_Ms_ControlLock(Mc2("lControl"), True)
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    Call Gp_Ms_NeceColor(Mc2("nControl"))
    
    
    Screen.MousePointer = vbDefault
    
    Call Sp_Header_display(Proc_Sc("Sc2")("Spread"))
    Call LC_Sp_Setting(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc2")("Spread"))
    Call Gp_Sp_Setting(Proc_Sc("Sc1")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc1")("Spread"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc2")("Spread"), "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    
    With MDIMain.MenuTool
        .Buttons(10).Enabled = False                'Separator
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(13).Enabled = False                'Separator
        .Buttons(14).Enabled = True                 'Excel
    End With

    Call SS1_HEAD_SET
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc1")("Spread"), "Q-System.INI", Me.Name)
    
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
    txt_charge_no.Text = ""
End Sub

Public Sub Form_Ref()
    
    On Error GoTo Refer_Err
    Dim lngRow As Long
    
    If txt_Charge_Date_Fr.RawData = "" And (txt_from_heat.Text = "" Or txt_to_heat.Text = "") Then
        txt_Charge_Date_Fr.RawData = Format(Now, "YYYYMMDD")
    End If
    
    If txt_Charge_Date_to.RawData = "" And (txt_from_heat.Text = "" Or txt_to_heat.Text = "") Then
        txt_Charge_Date_to.RawData = Format(Now, "YYYYMMDD")
    End If
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc1"), Mc2, Mc2("nControl")) Then
        Call Gp_Sp_ReadOnlySet(ss1)
    End If
    If ss1.MaxRows > 0 Then
        Screen.MousePointer = vbHourglass
        For lngRow = 1 To ss1.MaxRows
            Call std_ChemValChk(lngRow)
        Next
        ss1.Row = 1: ss1.Col = 1: txt_charge_no.Text = Trim(ss1.Text)
        ss1.Col = 5: str_STLGRD = Trim(ss1.Text)
        ss1.Col = 6: str_STLGRD_DETAIL = Trim(ss1.Text)
        Call Sp_Refer2
        'Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl"))
        'Call Ms_Chm_BitsSet(str_STLGRD, Mc1("rControl"))
        
        ss1.SetFocus
        lngActiveRow = 1
        Screen.MousePointer = vbDefault
    End If
    
    Exit Sub
    
Refer_Err:

End Sub

Public Sub Form_Pro()
   
'
End Sub
Public Sub Spread_Can()
    '
End Sub


Private Sub txt_Charge_Date_FR_DblClick()

    txt_Charge_Date_Fr.RawData = Format(Now, "YYYYMMDD")
        
End Sub
Private Sub txt_Charge_Date_TO_DblClick()

    txt_Charge_Date_to.RawData = Format(Now, "YYYYMMDD")
        
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
        .Text = "����(��׼)\�ɷ�"

        .Row = 1
        .Text = str_STLGRD_DETAIL

        .Row = 4
        .Text = "ת¯"

        .Row = 5
        .Text = "LF"

        .Row = 6
        .Text = "VD/RH"

        .Row = 7
        .Text = "CCM"

        .Col = 2

        .Row = 1
        .Text = "��Сֵ"
        .Row = 2
        .Text = "���ֵ"
        .Row = 3
        .Text = "Ŀ��ֵ"

        .Row = 4
        .Text = "ʵ��"

        .Row = 5
        .Text = "ʵ��"

        .Row = 6
        .Text = "ʵ��"

        .Row = 7
        .Text = "ʵ��"
        
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
    ss2.ReDraw = True
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
            Call Gp_MsgBoxDisplay("����ؼ�¼", "I")
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

    ss1.Col = 1: ss1.Row = Row: txt_charge_no.Text = Trim(ss1.Text)
    ss1.Col = 5: str_STLGRD = Trim(ss1.Text)
    ss1.Col = 6: str_STLGRD_DETAIL = Trim(ss1.Text)
    Call Sp_Refer2
    'Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl"))
    'Call Ms_Chm_BitsSet(str_STLGRD, Mc1.Item("rControl"))
    'Call std_ChemValChk(lngActiveRow)
    lngActiveRow = ss1.ActiveRow
End Sub
Private Sub Sp_Refer2()
    On Error GoTo Refer_Err

    Dim sMsg As String
    Dim sQuery As String
    Dim sQuery_cnt As String
    txt_charge_no.Text = Mid(txt_charge_no.Text, 1, 8)
    
    sMsg = Gf_Ms_NeceCheck(Mc1("nControl"))
    If sMsg <> "OK" Then
        sMsg = sMsg + "��������"
        Call Gp_MsgBoxDisplay(sMsg)
        Exit Sub
    End If
    Call Sp_Header_Set(ss2)
    Call LC_Sp_Display(M_CN1, Proc_Sc("Sc2")("Spread"), Gf_Ms_MakeQuery(Proc_Sc("Sc2").Item("P-R"), "R", Mc1("pControl")))
                
Refer_Err:

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    If lngActiveRow > 0 Then
        ss1.Row = lngActiveRow: ss1.Col = 1
        AQC0090C.txt_HEAT_OLC_NO = Trim(ss1.Text)
        AQC0090C.txt_charge_no = Trim(ss1.Text)
        AQC0090C.Show
        AQC0090C.SetFocus
        Call AQC0090C.Form_Ref
    Else
        Exit Sub
    End If
End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
'    Debug.Print KeyCode
    Select Case KeyCode
    Case 33, 34, 38, 40

        ss1.Col = 1: ss1.Row = ss1.ActiveRow: txt_charge_no.Text = Trim(ss1.Text) + "��׼"
        ss1.Col = 5: str_STLGRD = Trim(ss1.Text)
        ss1.Col = 6: str_STLGRD_DETAIL = Trim(ss1.Text)
        Call Sp_Refer2
        'Call Gf_Ms_Refer(M_CN1, Mc1, Mc1("pControl"), Mc1("mControl"))
        'Call Ms_Chm_BitsSet(str_STLGRD, Mc1("rControl"))

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
    
    'Call Gp_Sp_Excel(Me, Proc_Sc("Sc2")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    Call Sp_Excel(ss1)
End Sub

Public Sub Form_Ins()
    '
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

Private Sub SS1_HEAD_SET()

On Error GoTo SS1_HEAD_SET_ERR

    Dim iCol As Integer
    Dim iCnt As Integer
    Dim iColCnt As Integer
    Dim sQuery As String
    
    Dim AdoRs As adodb.Recordset
    Dim ArrayRecords As Variant

    Set AdoRs = New adodb.Recordset
    
    sQuery = " SELECT CHEM_COMP_CD From QP_CHEM_SEQ ORDER BY CHEM_COMP_SEQ ASC "
    
    With ss1

        .ReDraw = False
 
        
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
        
            .Row = 1
            
            For iCol = 11 To 59
            
                .Col = iCol
                .ColWidth(.Col) = 8
                
                If VarType(ArrayRecords(0, iCol - 11)) = vbNull Then
                    .Text = ""
                Else
                    .Text = Trim(ArrayRecords(0, iCol - 11))
                End If
                    
            Next iCol
            
        End If
 
        .ReDraw = True
        
        Screen.MousePointer = vbDefault
        
    End With
    
Exit Sub

SS1_HEAD_SET_ERR:
    
    Set AdoRs = Nothing
    ss1.ReDraw = True
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub std_ChemValChk(ByVal Row As Long)
On Error GoTo std_ChemValChk_Error
    Dim sQuery As String
    Dim strTemp As String
    Dim strTemp2 As String
    Dim lngCol As Long
    Dim strCHK As String
    Dim AdoRs As adodb.Recordset
    
    Set AdoRs = New adodb.Recordset
    ss1.Row = Row: ss1.Col = 5
    strTemp = "Select QP_ELEMENT_DEC( '" + Trim(ss1.Text) + "','"
    
    With ss1
        For lngCol = 11 To 59
            .Row = SpreadHeader + 1: .Col = lngCol
            strTemp2 = strTemp + Trim(ss1.Text) + "','"
            .Row = Row
            sQuery = strTemp2 + Trim(.Text) + "') FROM DUAL"
             'Ado Execute
            AdoRs.Open sQuery, M_CN1, adOpenKeyset
            If Not AdoRs.BOF And Not AdoRs.EOF Then
    
                If Not AdoRs.EOF Then
                strCHK = Trim(AdoRs.Fields(0))
                End If
                If strCHK = "N" Then
                    .ForeColor = vbRed
                End If
            End If
            AdoRs.Close
        Next lngCol

    End With
    Set AdoRs = Nothing
    Exit Sub

std_ChemValChk_Error:
    Set AdoRs = Nothing
    Exit Sub
End Sub

Private Sub Sp_Excel(sPname As Variant)

On Error GoTo Excel_Error

    Dim ret         As Boolean
    Dim xlApp       As Object
    Dim xlBpp       As Object
    Dim xlBook      As Object
    Dim xlSheet     As Object
    Dim ColIndex    As Integer
    Dim sExlRange   As String
    Dim sExlRange1  As String
    Dim iExlCol     As Integer
    
    Dim RowIndex     As Long
    
    With sPname
    
        If .MaxRows = 0 Then Exit Sub
        
      
        Clipboard.Clear
        
        .Col = 1: .Col2 = .MaxCols - 2
        .Row = SpreadHeader: .Row2 = .MaxRows
        Clipboard.SetText .Clip
        
        'Call Excel
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.Worksheets(1)
    
        xlApp.Visible = True
        
        xlSheet.Cells.NumberFormatLocal = "G/ͨ�ø�ʽ"
        xlSheet.Range("A1").Select
        xlSheet.Paste
        xlSheet.Cells.EntireColumn.AutoFit       'Column AutoFit
        
        sExlRange1 = ""
        
        For ColIndex = 1 To .MaxCols - 2
            .Col = ColIndex
            .Row = 1
            
            iExlCol = ColIndex
            If ColIndex > 10 Or (IsNumeric(.Text) And Left(.Text, 1) = "0" And _
               (Len(.Text) = 8 Or Len(.Text) = 10 Or Len(.Text) = 12 Or Len(.Text) = 14)) Then
                If ColIndex > 104 Then
                    sExlRange1 = "D" & sExlRange1
                    iExlCol = ColIndex - 104
                ElseIf ColIndex > 78 Then
                    sExlRange1 = "C" & sExlRange1
                    iExlCol = ColIndex - 78
                ElseIf ColIndex > 52 Then
                    sExlRange1 = "B" & sExlRange1
                    iExlCol = ColIndex - 52
                ElseIf ColIndex > 26 Then
                    sExlRange1 = "A"
                    iExlCol = ColIndex - 26
                End If
                
                sExlRange = sExlRange1 & Chr(iExlCol + 64) & "1:" & sExlRange1 & Chr(iExlCol + 64) & .MaxRows + 5
                If Len(.Text) = 8 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000"
                ElseIf Len(.Text) = 10 Then
                    xlSheet.Range(sExlRange).NumberFormat = "0000000000"
                ElseIf Len(.Text) = 12 Then
                    xlSheet.Range(sExlRange).NumberFormat = "000000000000"
                ElseIf Len(.Text) = 14 Then
                    xlSheet.Range(sExlRange).NumberFormat = "00000000000000"
                End If
                
                For RowIndex = 1 To .MaxRows
                    .Row = RowIndex
                    sExlRange = Trim(sExlRange1) & Chr(iExlCol + 64) & Trim(str(RowIndex + 2))
                    If .ForeColor = vbRed Then
                        xlSheet.Range(sExlRange).Font.ColorIndex = 3
                        xlSheet.Range(sExlRange).Font.Name = "����"
                        xlSheet.Range(sExlRange).Font.Size = 16
                    End If
                Next
            End If
        Next
       
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
    End With
    
    Exit Sub
    
Excel_Error:
    Call Gp_MsgBoxDisplay("���Ļ�����δ��װExcel", "W")

End Sub
