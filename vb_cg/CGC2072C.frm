VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGC2072C 
   Caption         =   "ʵ��ȡ��¼�����_CGC2072C"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_place 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14370
      MaxLength       =   12
      TabIndex        =   11
      Top             =   1710
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox txt_in_car_no 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14910
      MaxLength       =   12
      TabIndex        =   10
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ComboBox CBO_GROUP 
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
      ItemData        =   "CGC2072C.frx":0000
      Left            =   7155
      List            =   "CGC2072C.frx":0002
      TabIndex        =   3
      Tag             =   "���"
      Top             =   60
      Width           =   855
   End
   Begin VB.ComboBox CBO_SHIFT 
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
      ItemData        =   "CGC2072C.frx":0004
      Left            =   6300
      List            =   "CGC2072C.frx":0006
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox TXT_MAT_NO 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      MaxLength       =   14
      TabIndex        =   1
      Top             =   480
      Width           =   2160
   End
   Begin VB.TextBox TXT_SEQ 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6300
      MaxLength       =   12
      TabIndex        =   0
      Top             =   480
      Width           =   870
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   4095
      Left            =   0
      TabIndex        =   4
      Top             =   1290
      Width           =   15060
      _Version        =   393216
      _ExtentX        =   26564
      _ExtentY        =   7223
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   38
      MaxRows         =   10
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGC2072C.frx":0008
   End
   Begin InDate.ULabel ULabel16 
      Height          =   315
      Left            =   60
      Top             =   480
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "��ѯ��"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   4980
      Top             =   480
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "�ֶκ�"
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
      Left            =   4980
      Top             =   60
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "���/��"
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
   Begin InDate.UDate SDT_PROD_DATE_FROM 
      Height          =   315
      Left            =   1380
      TabIndex        =   5
      Tag             =   "��ʼ����"
      Top             =   60
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin InDate.UDate SDT_PROD_DATE_TO 
      Height          =   315
      Left            =   3135
      TabIndex        =   6
      Tag             =   "��ʼ����"
      Top             =   60
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
   Begin InDate.ULabel ULabel27 
      Height          =   315
      Left            =   60
      Top             =   60
      Width           =   1290
      _ExtentX        =   2275
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
   Begin Threed.SSPanel SSP1 
      Height          =   315
      Left            =   13650
      TabIndex        =   7
      Top             =   60
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   16711680
      BackColor       =   16761087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "һ���ඩ��"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSP2 
      Height          =   315
      Left            =   13650
      TabIndex        =   8
      Top             =   420
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   0
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "�ȴ���ָʾ"
      FloodColor      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin InDate.ULabel ULabel8 
      Height          =   315
      Left            =   11760
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      Caption         =   "���"
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
   Begin CSTextLibCtl.sidbEdit SDB_THK 
      Height          =   315
      Left            =   11190
      TabIndex        =   12
      Top             =   510
      Width           =   1065
      _Version        =   262145
      _ExtentX        =   1879
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      ShowZero        =   0   'False
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_THK_TO 
      Height          =   315
      Left            =   12660
      TabIndex        =   13
      Top             =   510
      Width           =   945
      _Version        =   262145
      _ExtentX        =   1667
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      FocusSelect     =   -1  'True
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "0.00"
      Text            =   ""
      StartText.x     =   3
      StartText.y     =   3
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   15
      CurNumDataChars =   0
      MaxDataChars    =   0
      FirstDataPos    =   0
      CurPos          =   0
      MaxLen          =   0
      DataReadOnly    =   0   'False
      Mask            =   ""
      Justification   =   2
      BorderStyle     =   0
      FmtControl      =   1
      NumDecDigits    =   2
      NumIntDigits    =   4
      ShowZero        =   0   'False
      MaxValue        =   9999.99
      MinValue        =   0
      Undo            =   0
      Data            =   0
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1185
      Left            =   8130
      TabIndex        =   15
      Top             =   60
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2090
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.ComboBox CBO_CAR_NO 
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
         ItemData        =   "CGC2072C.frx":1240
         Left            =   1440
         List            =   "CGC2072C.frx":124A
         TabIndex        =   19
         Top             =   630
         Width           =   1020
      End
      Begin VB.ComboBox CBO_LINE 
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
         ItemData        =   "CGC2072C.frx":125E
         Left            =   180
         List            =   "CGC2072C.frx":1271
         TabIndex        =   17
         Top             =   630
         Width           =   1020
      End
      Begin Threed.SSOption opt_on 
         Height          =   255
         Left            =   300
         TabIndex        =   16
         Top             =   30
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "����"
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   180
         Top             =   300
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         Caption         =   "����"
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
      Begin Threed.SSOption opt_trk 
         Height          =   255
         Left            =   1530
         TabIndex        =   20
         Top             =   30
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   450
         _Version        =   196609
         Font3D          =   1
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ʵ��"
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   1440
         Top             =   300
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   556
         Caption         =   "ȡ��"
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
   End
   Begin Threed.SSPanel SSP4 
      Height          =   315
      Left            =   13650
      TabIndex        =   18
      Top             =   810
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      _Version        =   196609
      ForeColor       =   0
      BackColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��ȡ��"
      FloodColor      =   0
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin FPSpread.vaSpread ss2 
      Height          =   3435
      Left            =   0
      TabIndex        =   21
      Top             =   5400
      Width           =   15060
      _Version        =   393216
      _ExtentX        =   26564
      _ExtentY        =   6059
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
      ColsFrozen      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   13
      MaxRows         =   10
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGC2072C.frx":1284
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   12360
      TabIndex        =   14
      Top             =   630
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "~"
      Height          =   120
      Left            =   2955
      TabIndex        =   9
      Top             =   180
      Width           =   195
   End
End
Attribute VB_Name = "CGC2072C"
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
'-- Program Name      ʵ��ȡ��¼�����
'-- Program ID        CGC2072C
'-- Document No       Q-00-0010(Specification)
'-- Designer          LI CHAO
'-- Coder             LI CHAO
'-- Date              2013.12.05
'-- Description       ʵ��ȡ��¼�����
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

'Dim pControl1 As New Collection      'Master Primary Key Collection
'Dim nControl1 As New Collection      'Master Necessary Collection
'Dim mControl1 As New Collection      'Master Maxlength check Collection
'Dim iControl1 As New Collection      'Master Insert Collection
'Dim rControl1 As New Collection      'Master Refer Collection
'Dim cControl1 As New Collection      'Master Copy Collection
'Dim aControl1 As New Collection      'Master -> Spread Collection
'Dim lControl1 As New Collection      'Master Lock Collection

Dim pColumn  As New Collection      'Spread Primary Key Collection
Dim nColumn  As New Collection      'Spread necessary Column Collection
Dim mColumn  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn  As New Collection      'Spread Insert Column Collection
Dim aColumn  As New Collection      'Master -> Spread Column Collection
Dim lColumn  As New Collection      'Spread Lock Column Collection

Dim pColumn2  As New Collection      'Spread Primary Key Collection
Dim nColumn2  As New Collection      'Spread necessary Column Collection
Dim mColumn2  As New Collection      'Spread Maxlength check Column Collection
Dim iColumn2  As New Collection      'Spread Insert Column Collection
Dim aColumn2  As New Collection      'Master -> Spread Column Collection
Dim lColumn2  As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim Mc2 As New Collection           'Master Collection
'Dim Mc2 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim mOplate_No As String

Const SS1_PLAN_SMP = 1  'ȡ��
Const SS1_ORD_CNT = 2
Const SS1_MOTHER_NO = 3
Const SS1_MV_DATE = 4
Const SS1_SHIFT = 5
Const SS1_TRNS_CMPY_CD = 6
Const SS1_OUT_SHEET_NO = 7
Const SS1_TRIM_FL = 9
Const SS1_PLATE_NO = 10
Const SS1_ORD_THK = 11
Const SS1_ORD_WID = 12
Const SS1_ORD_LEN = 13
Const SS1_SIZE_KND = 14
Const SS1_LEN_LIM = 15
Const SS1_THK_LIM = 16
Const SS1_APLY_STDSPEC = 18
Const SS1_LEN = 22
Const SS1_UST_STATUS = 24
Const SS1_GAS_STATUS = 25
Const SS1_CL_STATUS = 26
Const SS1_HTM_METH = 27
Const SS1_QT = 28
Const SS1_ORD_NO = 29
Const SS1_CUST_CD = 31
Const SS1_CUST_NAME = 32
Const SS1_ORD_REMARK = 33
Const SS1_STDSPEC_STLGRD = 34
Const SS1_STDSPEC_ORG_KND = 35
Const SS1_CD_MANA_NO = 36
Const SS1_SUSERID = 37
Const SS1_IN_CAR_NO = 38


Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
     FormType = "Msheet"

     'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(TXT_SEQ, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(txt_place, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
              Call Gp_Ms_Collection(CBO_LINE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_in_car_no, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(SDT_PROD_DATE_FROM, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(SDT_PROD_DATE_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(CBO_SHIFT, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
             Call Gp_Ms_Collection(CBO_GROUP, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               Call Gp_Ms_Collection(SDB_THK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(SDB_THK_TO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
               
     Mc1.Add Item:=pControl, Key:="pControl"
     Mc1.Add Item:=nControl, Key:="nControl"
     Mc1.Add Item:=mControl, Key:="mControl"
     Mc1.Add Item:=iControl, Key:="iControl"
     Mc1.Add Item:=rControl, Key:="rControl"
     Mc1.Add Item:=cControl, Key:="cControl"
     Mc1.Add Item:=aControl, Key:="aControl"
     Mc1.Add Item:=lControl, Key:="lControl"
     
     Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", " ", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)

    'MASTER Collection
    'Mc2.Add Item:="CGC2072C.P_SREFER2", Key:="P-R"
     Mc2.Add Item:=pControl2, Key:="pControl"
     Mc2.Add Item:=nControl2, Key:="nControl"
     Mc2.Add Item:=mControl2, Key:="mControl"
     Mc2.Add Item:=iControl2, Key:="iControl"
     Mc2.Add Item:=rControl2, Key:="rControl"
     Mc2.Add Item:=cControl2, Key:="cControl"
     Mc2.Add Item:=aControl2, Key:="aControl"
     Mc2.Add Item:=lControl2, Key:="lControl"
    
     
        Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
        Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", "l", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", "i", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
       Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", " ", pColumn, nColumn, mColumn, iColumn, aColumn, lColumn)
'
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGC2072C.P_SREFER", Key:="P-R"
    sc1.Add Item:="CGC2072C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn, Key:="pColumn"
    sc1.Add Item:=nColumn, Key:="nColumn"
    sc1.Add Item:=aColumn, Key:="aColumn"
    sc1.Add Item:=mColumn, Key:="mColumn"
    sc1.Add Item:=iColumn, Key:="iColumn"
    sc1.Add Item:=lColumn, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", " ", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
  
    'Spread_Collection
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="CGC2072C.P_SREFER2", Key:="P-R"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"

     Me.KeyPreview = True
     Me.BackColor = &HE0E0E0
     
End Sub

Private Sub CBO_CAR_NO_Click()
    If CBO_CAR_NO.Text = "��ȡ��" Then
        txt_in_car_no.Text = "Y"
    ElseIf CBO_CAR_NO.Text = "δȡ��" Then
        txt_in_car_no.Text = "N"
    End If
End Sub

Private Sub opt_on_Click(Value As Integer)
    txt_place.Text = "1"
    If opt_on.Value = True Then
       opt_on.ForeColor = &HFF&
       opt_trk.ForeColor = &H0&
    Else
        opt_on.ForeColor = &H0&
    End If
End Sub

Private Sub opt_trk_Click(Value As Integer)
    txt_place.Text = "2"
    If opt_trk.Value = True Then
        opt_trk.ForeColor = &HFF&
        opt_on.ForeColor = &H0&
    Else
        opt_trk.ForeColor = &H0&
    End If
End Sub

Private Sub SDT_PROD_DATE_FROM_GotFocus()
     If SDT_PROD_DATE_FROM.RawData = "" Then
        SDT_PROD_DATE_FROM.RawData = Gf_DTSet(M_CN1, "D")
     End If
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub SDT_PROD_DATE_TO_GotFocus()
     If SDT_PROD_DATE_TO.RawData = "" Then
        SDT_PROD_DATE_TO.RawData = Gf_DTSet(M_CN1, "D")
     End If
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)

    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "G-System.INI", Me.Name)
    
    CBO_SHIFT.AddItem "1"
    CBO_SHIFT.AddItem "2"
    CBO_SHIFT.AddItem "3"
    
    CBO_GROUP.AddItem "A"
    CBO_GROUP.AddItem "B"
    CBO_GROUP.AddItem "C"
    CBO_GROUP.AddItem "D"
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    opt_on.Value = True
    
    'Call Gp_Sp_ColHidden(ss1, SS1_ORD_CNT, True)
    'Call Gp_Sp_ColHidden(ss1, SS1_SUSERID, True)
'    Call Gp_Sp_ColHidden(ss1, SS1_IN_CAR_NO, True)
    
    Screen.MousePointer = vbDefault

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
    
    Set iColumn = Nothing
    Set pColumn = Nothing
    Set lColumn = Nothing
    Set nColumn = Nothing
    Set mColumn = Nothing
    Set aColumn = Nothing

    Set Mc1 = Nothing
    Set Mc2 = Nothing
    Set sc1 = Nothing
    Set sc2 = Nothing
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub
Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)

End Sub
Public Sub Form_Exit()

    Unload Me

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(sc1) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
       Call Gp_Ms_ControlLock(Mc1("lControl"), False)
       Call Gp_Ms_Cls(Mc1("pControl"))
       SDT_PROD_DATE_FROM.RawData = ""
       SDT_PROD_DATE_TO.RawData = ""
    End If

End Sub

Public Sub Form_Ref()
    
    Dim SMESG As String
    Dim lRow As Long
    '����һ��������ǣ�����������ɫ��ʾ
    Dim iColor As Integer
    Dim sord_cnt As Integer
    Dim sHtm_Meth As String
    Dim sIncarno As String
    
    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Nothing) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    For lRow = 1 To ss1.MaxRows

        ss1.ROW = lRow:       ss1.Col = SS1_MOTHER_NO
        'ȡĸ��ţ���ʼֵΪ�գ�����ɫ�����Ϊ1
        If mOplate_No = "" Then
            iColor = 1
        Else
        'ĸ��Ų�Ϊ��ʱ���������һĸ����Ƿ�Ϊ��ͬĸ���
            If ss1.Text <> mOplate_No Then
            '����ǲ�ͬĸ��ţ�������ɫ���Ϊ1����ô��ɫ��Ǹ�Ϊ2����ʾ�ı���ɫ
                If iColor = 1 Then
                   iColor = 2
                   '���ĸ�����ͬ����ô��ɫ��ǻ�Ϊ1����ʾ��ɫ����
                Else
                   iColor = 1
                End If
            End If
       End If
       '��1��ʾ��ɫ��Ϊǳ��ɫ����2��ʾ��ɫ��Ϊ��ɫ
       'ÿ��ѭ�����������iColorΪ1������ɫΪǳ��ɫ��������ɫΪ��ɫ
       If iColor = 1 Then
         'ȡ����ɫ�ı䣬���ΪY��ʾȡ���������������ɫ���ɫ
          ss1.Col = SS1_PLAN_SMP
          If ss1.Text = "Y" Then
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, &HFF&, &HE0E0E0) 'ǳ��ɫ
          Else
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, , &HE0E0E0) 'ǳ��ɫ
          End If
       Else
          'ȡ����ɫ�ı䣬���ΪY��ʾȡ���������������ɫ���ɫ
          ss1.Col = SS1_PLAN_SMP
          If ss1.Text = "Y" Then
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, &HFF&, &HFFFFFF) '��
          Else
             Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, ss1.ROW, ss1.ROW, , &HFFFFFF) '��
          End If
       End If
       '��ֵ��ԭΪforѭ����ĸ��ŵ�ȡֵ
       ss1.Col = SS1_MOTHER_NO
       mOplate_No = ss1.Text
       
       ss1.ROW = lRow:          ss1.Col = SS1_ORD_CNT:          sord_cnt = Val(ss1.Text)
            If sord_cnt > 1 Then
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRow, lRow, , SSP1.BackColor)
            End If
            
       ss1.ROW = lRow:          ss1.Col = SS1_IN_CAR_NO:          sIncarno = ss1.Text
        If txt_place.Text = "2" Then
            If sIncarno = "Y" Then
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRow, lRow, , SSP4.BackColor)
            End If
        Else
            If sIncarno <> "" Then
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRow, lRow, , SSP4.BackColor)
            End If
        End If
        
        '�ȴ���ָʾ��ɫ��ʾ
        ss1.ROW = lRow:          ss1.Col = SS1_HTM_METH:          sHtm_Meth = Val(ss1.Text)
        If Mid(sHtm_Meth, 1, 1) = "N" And Mid(sHtm_Meth, 1, 1) <> "/" Then
            Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, lRow, lRow, , &HFF0000)
        End If
       
    Next lRow

End Sub


Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
         Call Gp_Sp_UpdateMake(Proc_Sc("SC")("Spread"), Mode)
         ss1.ROW = ss1.ActiveRow
         ss1.Col = SS1_SUSERID
         ss1.Text = sUserID
    End If
   
End Sub
Public Sub Form_Pro()
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub


Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    
    Dim lRow As Long
    Dim sBlockSeq As String
    Dim sSeq As String
    
'    Call Gp_Sp_Sort(sc1.Item("Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
    If ROW <= 0 Then Exit Sub
    If Col > 1 Then Exit Sub
    
    ss1.ROW = ROW
    ss1.Col = 3
    TXT_MAT_NO.Text = ss1.Text
    

    Call Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl"), False)
    ss2.OperationMode = OperationModeNormal
    TXT_MAT_NO.Text = ""
    

End Sub
