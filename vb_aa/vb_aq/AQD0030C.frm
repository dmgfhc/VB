VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form AQD0030C 
   Caption         =   "����֤������η���_AQD0030C"
   ClientHeight    =   9090
   ClientLeft      =   -15
   ClientTop       =   2340
   ClientWidth     =   15585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15585
   WindowState     =   2  'Maximized
   Begin VB.TextBox TXT_CONTROL_NO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   12330
      TabIndex        =   18
      Top             =   1080
      Width           =   1410
   End
   Begin VB.TextBox txt_plt 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   5655
      MaxLength       =   2
      TabIndex        =   16
      Tag             =   "plt"
      Top             =   135
      Width           =   465
   End
   Begin VB.TextBox txt_SAVE_DIR 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   7290
      TabIndex        =   11
      Top             =   1080
      Width           =   2835
   End
   Begin VB.TextBox TXT_PONO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   12330
      TabIndex        =   9
      Top             =   600
      Width           =   1410
   End
   Begin VB.TextBox txt_STDSPEC 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   12330
      MaxLength       =   15
      TabIndex        =   8
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox txt_CERT_NO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   2010
      MaxLength       =   14
      TabIndex        =   5
      Top             =   600
      Width           =   2235
   End
   Begin VB.TextBox txt_CUST_CD 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   6195
      MaxLength       =   6
      TabIndex        =   4
      Top             =   585
      Width           =   1125
   End
   Begin VB.TextBox txt_ORD_NO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   9300
      MaxLength       =   11
      TabIndex        =   3
      Top             =   135
      Width           =   1305
   End
   Begin VB.TextBox txt_CUST_NAME 
      Enabled         =   0   'False
      Height          =   310
      Left            =   7335
      TabIndex        =   2
      Top             =   585
      Width           =   3285
   End
   Begin VB.TextBox txt_PROD_CD 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   7620
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "��Ʒ����"
      Top             =   135
      Width           =   540
   End
   Begin Threed.SSCommand cmdReport 
      Height          =   375
      Left            =   13920
      TabIndex        =   0
      Top             =   1080
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�����ʱ���"
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   1
      Left            =   150
      Top             =   135
      Width           =   1020
      _ExtentX        =   1799
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
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   0
      Left            =   150
      Top             =   600
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "����֤������"
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
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   4335
      Top             =   585
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "�ͻ�"
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
      Left            =   8265
      Top             =   135
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   556
      Caption         =   "������"
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
   Begin InDate.UDate dtp_fr_date 
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      Tag             =   "��������"
      Top             =   135
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.74
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16777215
   End
   Begin InDate.UDate dtp_to_date 
      Height          =   315
      Left            =   2745
      TabIndex        =   7
      Tag             =   "��������"
      Top             =   135
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.74
         Charset         =   134
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
      Index           =   2
      Left            =   6225
      Top             =   135
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "��Ʒ"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   10860
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      Caption         =   "�ƺ�"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   10860
      Top             =   585
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      Caption         =   "��ͬ��"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand ssc_DIR_FIND 
      Height          =   345
      Left            =   10200
      TabIndex        =   10
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      _Version        =   196609
      PictureFrames   =   1
      Picture         =   "AQD0030C.frx":0000
      ButtonStyle     =   1
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   315
      Left            =   2010
      TabIndex        =   12
      Top             =   1080
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   556
      _Version        =   196609
      Font3D          =   1
      ShadowStyle     =   1
      Begin Threed.SSOption ssp_ELE 
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   30
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   0
         Caption         =   "�����ʱ�"
      End
      Begin Threed.SSOption ssp_PRN 
         Height          =   255
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   16448
         Caption         =   "ֱ�Ӵ�ӡ"
         Value           =   -1
      End
      Begin Threed.SSOption ssp_SAVE_PRN 
         Height          =   255
         Left            =   1170
         TabIndex        =   14
         Top             =   30
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   64
         Caption         =   "���沢��ӡ"
      End
      Begin Threed.SSOption ssp_SAVE 
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   30
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         ForeColor       =   0
         Caption         =   "���治��ӡ"
      End
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   150
      Top             =   1050
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      Caption         =   "��������ĵ�"
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
   Begin InDate.ULabel ULabel7 
      Height          =   300
      Left            =   4335
      Top             =   150
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   529
      Caption         =   "����"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   13920
      TabIndex        =   17
      Top             =   360
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����ͳ��"
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   10860
      Top             =   1080
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      Caption         =   "���ƺ�"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdReport_sms 
      Height          =   375
      Left            =   13920
      TabIndex        =   19
      Top             =   720
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����˵����"
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   7725
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1440
      Width           =   15120
      _Version        =   393216
      _ExtentX        =   26670
      _ExtentY        =   13626
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
      MaxCols         =   20
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "AQD0030C.frx":0352
   End
   Begin Threed.SSCommand cmd_All 
      Height          =   375
      Left            =   13920
      TabIndex        =   22
      Top             =   0
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ȫѡ"
   End
End
Attribute VB_Name = "AQD0030C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       ��������
'-- Sub_System Name   �ж�����
'-- Program Name      ����֤������η���
'-- Program ID        AQD0030C
'-- Document No       Q-00-0010(Specification)
'-- Designer          Chu Kyo Su
'-- Coder             Chu Kyo Su
'-- Date              2003.07. 25
'-- Description       ����֤������η���
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
Public sPLT_Authority As String     'Active User Plant Authority Setting

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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection
Dim bPrintCheck As Boolean

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

'---------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------
'------------------------------ Report Variable ----------------------------------------------
'---------------------------------------------------------------------------------------------
Dim xlApp       As Object
Dim xlSheet     As Object

Dim arrRecords1 As Variant      'sQueryHeadC
Dim arrRecords2 As Variant      'sQueryDetailC
Dim arrRecords8 As Variant      'sQueryDetailC

Dim arrRecords3 As Variant      'sQueryHeadS
Dim arrRecords4 As Variant      'sQueryDetailS

Dim arrRecords5 As Variant      'sQueryHeadP
Dim arrRecords6 As Variant      'sQueryDetailP
Dim arrRecords7 As Variant      'sQueryDetailP

Dim arrRecords10 As Variant      'sQueryHeadb
Dim arrRecords11 As Variant      'sQueryDetailb

Dim sQuery      As String
Dim sErrMsg     As String
Dim sDate       As String
Dim AdoRs       As adodb.Recordset
Dim sPICTURE    As String
Dim oPICTURE    As Variant

Dim Report_KND As String        '˵���飺T

'---------------------------------------------------------------------------------------------

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
       Call Gp_Ms_Collection(txt_CERT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(txt_PROD_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_fr_date, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(dtp_to_date, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(TXT_CUST_CD, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_ORD_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_PONO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_CONTROL_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    
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
    
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", "n", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQD0030C.P_MODIFY_SEND", Key:="P-M"
    Sc1.Add Item:="AQD0030C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQD0030C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
     
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cmd_All_Click()
  Dim i As Integer
  Dim v_text As String

  If cmd_All.Caption = "ȫѡ" Then
    v_text = "1"
    cmd_All.Caption = "ȡ��ȫѡ"
  Else
    v_text = "0"
    cmd_All.Caption = "ȫѡ"
  End If
    
  With ss1
    For i = 1 To .MaxRows
      .Row = i
      .Col = 1
      .Text = v_text
    Next i
 End With
End Sub

Private Sub cmdReport_sms_Click()
    Report_KND = "T"
    Call cmdReport_Click
    Report_KND = ""
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------- Code Name Find --------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
            
        Case "txt_PROD_CD"             '��Ʒ
            sCode = "B0005"
                    
        Case "txt_CUST_CD"              '�ͻ�����
            sCode = "CUST_CD"
            Set oCodeName = txt_CUST_NAME
            
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing
Err_Track:
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    Call subButtonHide
    
    If dtp_fr_date.RawData = "" Or dtp_to_date.RawData = "" Then
       dtp_fr_date.Text = Date
       dtp_to_date.Text = Date
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass
    
    
    sAuthority = Gf_Pgm_Authority(Me.Name, True)

'     If sAuthority = "1000" Then
'       cmdReport.Visible = False
'    End If
    
    sPLT_Authority = Gf_PLT_Authority(Me.Name)
    If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
       txt_plt.Text = sPLT_Authority
    Else
       txt_plt.Text = ""
    End If
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    
    Call Gp_Ms_Cls(Mc1("rControl"))
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    txt_PROD_CD.Text = "PP"
    
    Screen.MousePointer = vbDefault
    
    Call subButtonHide

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "Q-System.INI", Me.Name)
    
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
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
    Call subButtonHide
    
End Sub



Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
     '  rControl(1).SetFocus
        If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
           txt_plt.Text = sPLT_Authority
        Else
           txt_plt.Text = ""
        End If
    End If
    txt_CERT_NO.Text = ""
    TXT_CUST_CD.Text = ""
    txt_CUST_NAME.Text = ""
    txt_ORD_NO.Text = ""
    txt_PROD_CD.Text = ""
    txt_STDSPEC.Text = ""
    dtp_fr_date.Text = Date
    dtp_to_date.Text = Date
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    Dim sMesg As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
     If subCheck = True Then
        
            If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
                ss1.OperationMode = OperationModeNormal
                Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                Call subButtonHide
                Exit Sub
            End If
            
    Else
                
        GoTo Refer_Err
        
    End If
    
    Call subButtonHide
    
    bPrintCheck = False
    
    Exit Sub

Refer_Err:

End Sub
Public Sub Form_Pro()

'         Call Spread_Cheack
         If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
           Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
           Call Form_Ref
         End If

End Sub

Public Sub Spread_ColumnsSort()

    Spread_ColSort.Show 1
    
End Sub

Public Sub Spread_Forzens_Setting()

    Me.ActiveControl.ColsFrozen = Me.ActiveControl.ActiveCol
    
End Sub

Public Sub Spread_Forzens_Cancel()

    Me.ActiveControl.ColsFrozen = 0
    
End Sub

Public Sub Form_Exc()
    
    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

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

Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim b_PRT_FL        As Boolean
Dim s_Msg           As String
Dim s_Shp_ist_no    As String

    If Row < 1 Then Exit Sub
    
    b_PRT_FL = Print_FL(Row)
    
    With ss1
        .Row = Row
        .Col = 12
        s_Shp_ist_no = .Text
        
        If b_PRT_FL Then
            .Col = 1
            If .Text = "1" Then
               .Col = 0:    .Text = "Update"
            Else
               .Col = 0:    .Text = ""
            End If
        Else
            .Col = 1
            If .Text = "1" Then
                .Col = 1: .Text = "0"
                .Col = 0: .Text = ""
                s_Msg = "������� " + s_Shp_ist_no + " ����δ���յ��ʱ���,Ŀǰ���ᵥ�ʱ��鲻�ɷ���!"
                Call Gf_MessConfirm(s_Msg, "I")
            End If
        End If
    End With
    
    Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 15)
    
End Sub



Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)
    
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
    If ss1.MaxRows > 0 And ss1.ActiveRow > 0 Then
        
        If Col <> 1 And Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 1) <> "" Then
            
            AQD0020C.Show
            AQD0020C.SetFocus
            AQD0020C.txt_CERT_NO.Text = Gf_Get_Cell_Value(ss1, ss1.ActiveRow, 2)
            
            Call AQD0020C.Form_Ref
            
        End If
        
    End If
End Sub

'Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'
'  '  If Gf_Sc_Authority(sAuthority, "U") Then
'        Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
'        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 13)
'   ' End If
'
'End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 15)
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'    Call GP_SetRowHeaderClear(ss1, NewRow)
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


Private Sub subButtonHide()

'    MDIMain.MenuTool.Buttons(4).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(5).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Inssert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row cancel
    
    MDIMain.MenuTool.Buttons(11).Enabled = False    'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False    'paste
    

End Sub

'-----------------------------------------------------------------------
'---------------------------- Report Main ------------------------------
'-----------------------------------------------------------------------
Private Sub cmdReport_Click()
    Dim sCertNo As String
    Dim sFlag   As String
    Dim i       As Integer
    Dim Save_Path As String
    Dim Save_State As Integer
    Dim sEMP_ID As String
    Screen.MousePointer = vbHourglass
    
    
    Save_State = 0
    
       If ssp_PRN.Value = True Then  'Or ssp_ELE.Value = True
            Save_State = 0
        End If
        If ssp_SAVE_PRN.Value = True Then
            Save_State = 1
        End If
        If ssp_SAVE.Value = True Then
            Save_State = 2
        End If
        If ssp_ELE.Value = True Then
            Save_State = 3
        End If
    
    Save_Path = Trim(txt_SAVE_DIR)
    
    sErrMsg = ""
    
    Set AdoRs = New adodb.Recordset
    
    sQuery = "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    sDate = AdoRs.Fields(0)
    
    AdoRs.Close
    Set AdoRs = Nothing
    
    With ss1
        
        For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .Text = "1" Then
                .Col = 2:     sCertNo = Trim(.Text)
                .Col = 14:    sFlag = Trim(.Text)
                If Report_KND = "T" Then
                    sErrMsg = Cert_type_check("T", sCertNo, Save_State, Save_Path)
                Else
                    sErrMsg = Cert_type_check(sFlag, sCertNo, Save_State, Save_Path)
                End If
                
                '��ӡ�����ʱ���
                If ssp_ELE.Value = True Then
                  sErrMsg = Electronic_certificate(sCertNo)
                End If
                
                .Col = 15:    sEMP_ID = Trim(.Text)         '��ӡ��
                 Call SEND_ERP("U", 1, sCertNo, sEMP_ID)    '��¼��ӡ����
                                    
                If sErrMsg <> "" Then
                    i = .MaxRows
                End If
            End If
        Next i
        
'        If sErrMsg = "" Then
'            If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) = False Then Exit Sub
'        End If
'
    End With

    Call subButtonHide
    
    With ss1
       For i = 1 To .MaxRows
            .Col = 0
            .Row = i
          If Trim(.Text) = "Input" Or Trim(.Text) = "Update" Or Trim(.Text) = "Delete" Then
            .Text = ""
            .Col = 1
            .Value = 0
          End If
       Next i
    End With
    
    Call Form_Ref
    
    Report_KND = ""
    
    Screen.MousePointer = vbDefault
        
        
End Sub

'--------------------------------------------------------------------------------------------------------
'------------------------------------------- Local Procedure --------------------------------------------
'--------------------------------------------------------------------------------------------------------

Private Function subCheck() As Boolean

    Dim sMesg As String
    Dim sFrDate As String
    Dim sToDate As String
    Dim sProdCd As String
    Dim sCertNo As String
    Dim sOrdNo As String
    Dim sTrnsNo As String
    Dim sCustCD As String
    
    sProdCd = Trim(txt_PROD_CD.Text)
    sCertNo = Trim(txt_CERT_NO.Text)
    sOrdNo = Trim(txt_ORD_NO.Text)
    sTrnsNo = Trim(txt_STDSPEC.Text)
    sCustCD = Trim(TXT_CUST_CD.Text)
    
    sFrDate = Trim(dtp_fr_date.Text)
    sToDate = Trim(dtp_to_date.Text)
    
    sFrDate = Replace(sFrDate, "_", "")
    sToDate = Replace(sToDate, "_", "")

    sFrDate = Replace(sFrDate, "-", "")
    sToDate = Replace(sToDate, "-", "")
    
    If sCertNo = "" Then
        If sFrDate = "" Or sToDate = "" Then
            sMesg = "���������뷢�����ڣ���ʼ�ͽ������ڣ�"
            Call Gp_MsgBoxDisplay(sMesg)
            subCheck = False
            Exit Function
        Else
            If sCustCD = "" And sOrdNo = "" And sProdCd = "" And sTrnsNo = "" Then
                sMesg = "�����롰��Ʒ���롱��������š��򡰿ͻ����롱�򡰶����š��е�����һ��"
                Call Gp_MsgBoxDisplay(sMesg)
                subCheck = False
                Exit Function
            End If
        End If
    End If
        
    subCheck = True

End Function

Private Sub ssc_DIR_FIND_Click()
    Load Form_DIR_SELECT
    Form_DIR_SELECT.Show 1
   txt_SAVE_DIR.Text = sEXLSavePATH
End Sub

'Private Sub SSCommand1_Click()
'
'End Sub

Private Sub SSCommand1_Click()
    AQD0031C.Show
    AQD0031C.SetFocus
End Sub

Private Sub ssp_ELE_Click(Value As Integer)
  If ssp_ELE.Value = 1 Then
    txt_SAVE_DIR.Text = "C:\ELE_CERT"
    txt_SAVE_DIR.Enabled = False
  End If
  
End Sub

Private Sub txt_CERT_NO_Change()
    Call Gf_Control_text_Up(txt_CERT_NO)
End Sub

Private Sub txt_PROD_CD_Change()
    Call Gf_Control_text_Up(txt_PROD_CD)
End Sub

Private Function Print_FL(ByVal iRow As Long) As Boolean
     With ss1
     
          .Row = iRow
          .Col = 17
          If .Text = "n" Or .Text = "N" Then
             Print_FL = False
          Else
             Print_FL = True
          End If
                           
      End With
End Function

Private Sub TXT_PLT_Change()

    If txt_plt.Text = "C3" Then
       txt_PROD_CD.Text = "PP"
    End If

End Sub

Private Sub txt_PLT_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_plt

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If

End Sub

Private Function SEND_ERP(iType As String, iCheck As String, Cert_No As String, EMP_ID As String) As Boolean

    Dim OutParam(2, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim sMesg As String
    
    
    Dim adoCmd As adodb.Command
    Screen.MousePointer = vbHourglass
    

    OutParam(1, 1) = "arg_CD"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 1
    OutParam(2, 1) = "arg_e_msg"
    OutParam(2, 2) = adVarChar
    OutParam(2, 3) = adParamOutput
    OutParam(2, 4) = 256
    

    sQuery = "{call AQD0030C.P_MODIFY ('" + iType + "', '" + iCheck + "','" + Cert_No + "','" + EMP_ID + "',?,?)}"
'AQD0520P(P_PROD_CD,P_INSP_CD,P_STDSPEC,P_CON_NO)
    'Ado Setting
    M_CN1.CursorLocation = adUseServer
    Set adoCmd = New adodb.Command
    
    adoCmd.CommandType = adCmdText
    Set adoCmd.ActiveConnection = M_CN1
    
    adoCmd.CommandText = sQuery
    
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(1, 1), OutParam(1, 2), OutParam(1, 3), OutParam(1, 4))
    adoCmd.Parameters.Append adoCmd.CreateParameter(OutParam(2, 1), OutParam(2, 2), OutParam(2, 3), OutParam(2, 4))
    
    adoCmd.Execute , , adExecuteNoRecords
    
    'Process Error Check
    If adoCmd("arg_e_msg") <> "" Then
        ret_Result_ErrMsg = adoCmd("arg_e_msg")
        
        sErrMessg = "Error Mesg : " & ret_Result_ErrMsg
        
        Screen.MousePointer = vbDefault
        Call Gp_MsgBoxDisplay(sErrMessg)
        Set adoCmd = Nothing
        Exit Function
        
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Exit Function
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    
    Err.Raise Err.Number, Err.Description & sQuery
    
End Function

 'LR �����ʱ��鴦��
 
Private Function Electronic_certificate(sCertNo As String) As String
                

        Name "c:\ele_cert\cert.pdf" As "c:\ELE_CERT\" & sCertNo & ".pdf"


End Function

