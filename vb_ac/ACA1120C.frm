VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "indate.ocx"
Begin VB.Form ACA1120C 
   Caption         =   "��ʷ�������̲�ѯ_ACA1120C"
   ClientHeight    =   9420
   ClientLeft      =   375
   ClientTop       =   2460
   ClientWidth     =   15630
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   15630
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   1191
      _Version        =   196609
      BackColor       =   14737632
      ShadowStyle     =   1
      Begin VB.ComboBox Combo_ORD_ITEM 
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
         Left            =   4350
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.TextBox Text_BB_ORD_NO 
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
         Left            =   3000
         MaxLength       =   11
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox Text_BB_PROD_CD_mate 
         Height          =   315
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.TextBox txt_cfm_mill_plt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   3660
         MaxLength       =   2
         TabIndex        =   2
         Top             =   120
         Width           =   540
      End
      Begin VB.TextBox Text_BB_PROD_CD 
         Alignment       =   2  'Center
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
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "��Ʒ"
         Top             =   120
         Width           =   645
      End
      Begin InDate.ULabel ULabel9 
         Height          =   315
         Left            =   120
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         Caption         =   "��Ʒ"
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
         ForeColor       =   16711680
      End
      Begin InDate.ULabel ULabel3 
         Height          =   315
         Left            =   4560
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         Caption         =   "�û�������"
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
      Begin InDate.UDate Udate_BB_DEL_TO 
         Height          =   315
         Left            =   7500
         TabIndex        =   3
         Tag             =   "������"
         Top             =   120
         Width           =   1440
         _ExtentX        =   2540
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
         MaxLength       =   10
      End
      Begin InDate.ULabel ULabel01 
         Height          =   315
         Index           =   14
         Left            =   2370
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         Caption         =   "��������"
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
         ForeColor       =   16711680
      End
      Begin InDate.UDate Udate_BB_DEL_FR 
         Height          =   315
         Left            =   5850
         TabIndex        =   4
         Tag             =   "������"
         Top             =   120
         Width           =   1440
         _ExtentX        =   2540
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
         MaxLength       =   10
      End
      Begin InDate.ULabel ULabel2 
         Height          =   315
         Left            =   9360
         Top             =   120
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         Caption         =   "������������"
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
      Begin InDate.UDate Udate_BB_INS 
         Height          =   315
         Left            =   10680
         TabIndex        =   9
         Tag             =   "������������"
         Top             =   120
         Width           =   1440
         _ExtentX        =   2540
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
         MaxLength       =   10
      End
      Begin Threed.SSPanel SSP90 
         Height          =   375
         Left            =   12360
         TabIndex        =   16
         Top             =   120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÿ��1��6��11��21��������"
         FloodColor      =   65535
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   120
         Left            =   7320
         TabIndex        =   11
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   120
         Left            =   2790
         TabIndex        =   10
         Top             =   120
         Width           =   90
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "~"
         Height          =   120
         Left            =   2790
         TabIndex        =   8
         Top             =   120
         Width           =   90
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "�û�������"
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
   Begin InDate.UDate UDate1 
      Height          =   315
      Left            =   2940
      TabIndex        =   5
      Tag             =   "������"
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      MaxLength       =   10
   End
   Begin InDate.UDate UDate2 
      Height          =   315
      Left            =   1290
      TabIndex        =   6
      Tag             =   "������"
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      MaxLength       =   10
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8460
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   720
      Width           =   15495
      _Version        =   393216
      _ExtentX        =   27331
      _ExtentY        =   14923
      _StockProps     =   64
      ButtonDrawMode  =   4
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
      MaxRows         =   2
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      SpreadDesigner  =   "ACA1120C.frx":0000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "~"
      Height          =   120
      Left            =   2790
      TabIndex        =   7
      Top             =   90
      Width           =   90
   End
End
Attribute VB_Name = "ACA1120C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       PROCESS MANAGEMENT
'-- Sub_System Name
'-- Program Name
'-- Program ID        ACA1120C
'-- Document No       Q-00-0010(Specification)
'-- Designer          CaoLei
'-- Coder             CaoLei
'-- Date              2013.04.26
'-- Description nnnn
'-------------------------------------------------------------------------------
'-- UPDATE HISTORY  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- VER   DATE     EDITOR       DESCRIPTION
'-------------------------------------------------------------------------------
'-- DECLARATION     ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'  -------------------------------------------------------------------------------

Public FormType As String           'Form Type
Public Toolbar_St As String         'Active Form ToolBar Setting
Public sAuthority As String         'Active Form Authority Setting
Public ORD_NO As String             'Transfer to ACA1030C
Public ORD_ITEM As String           'Transfer to ACA1030C


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
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim iSumCol As New Collection       'Sum Column

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim sCheck1 As String
Dim sCheck2 As String
Dim iCount As Integer

Const iSumColCnt = 12
Const iSumCol1 = 24
Const iSumCol2 = 25
Const iSumCol3 = 26
Const iSumCol4 = 27
Const iSumCol5 = 28
Const iSumCol6 = 29
Const iSumCol7 = 30
Const iSumCol8 = 31
Const iSumCol9 = 32
Const iSumCol10 = 33
Const iSumCol11 = 34
Const iSumCol12 = 35

Const SS1_URGNT_FL = 49          '����������ɫ���
Const SS1_ORD_NO = 1
Const SS1_ORD_ITEM = 2


Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
 '   FormType = "Msheet"
    FormType = "Refer"

   'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
      Call Gp_Ms_Collection(Text_BB_PROD_CD, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Udate_BB_DEL_FR, "p", "n", "m", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(Udate_BB_DEL_TO, "p", "n", "m", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Text_BB_ORD_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(Combo_ORD_ITEM, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     Call Gp_Ms_Collection(txt_cfm_mill_plt, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(Udate_BB_INS, "p", "n", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          
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
     Call Gp_Sp_Collection(ss1, 1, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, "p", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '������������
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
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '�ƻ�Ͷ�빤��
    
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '��������
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)  '����������ɫ���
     
   'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"

    sc1.Add Item:="ACA1120C.P_SREFER", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    
    sc1.Item("Spread").Col = 0
    sc1.Item("Spread").Row = 0
    sc1.Item("Spread").Text = "��"
    
    
    'Sum Column Count
    iSumCnt = iSumColCnt
    
    'Sum Column Setting
    iSumCol.Add Item:=iSumCol1
    iSumCol.Add Item:=iSumCol2
    iSumCol.Add Item:=iSumCol3
    iSumCol.Add Item:=iSumCol4
    iSumCol.Add Item:=iSumCol5
    iSumCol.Add Item:=iSumCol6
    iSumCol.Add Item:=iSumCol7
    iSumCol.Add Item:=iSumCol8
    iSumCol.Add Item:=iSumCol9
    iSumCol.Add Item:=iSumCol10
    iSumCol.Add Item:=iSumCol11
    iSumCol.Add Item:=iSumCol12
    
        
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
    
    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"), False)
    Call Gp_Sp_ReadOnlySet(Proc_Sc("Sc")("Spread"))
    Call Gf_Sp_Cls(Proc_Sc("Sc"))
    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
    
    Text_BB_PROD_CD.Text = "PP"
    
    Udate_BB_DEL_FR.Text = Mid(DateAdd("m", -1, Date), 1, 8) + "01"

    Udate_BB_DEL_TO.Text = Format(DateAdd("m", 1, Udate_BB_DEL_FR.Text), "YYYY-MM-DD")
    Udate_BB_DEL_TO.Text = DateAdd("d", -1, Udate_BB_DEL_TO.Text)
    
    Udate_BB_INS.Text = Mid(Date, 1, 8) + "06"

    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "C-System.INI", Me.Name)
    
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
    Set sc1 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub

Public Sub Form_Cls()

    If Gf_Sp_Cls(Proc_Sc("Sc")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If
    
    
    
    Udate_BB_DEL_TO.Text = DateAdd("m", -1, Now)   '����
    Udate_BB_DEL_TO.Text = DateAdd("d", -1, Udate_BB_DEL_TO.Text)     '�������һ��
    
    Udate_BB_DEL_FR.Text = Mid(Udate_BB_DEL_TO.Text, 1, 8) + "01"     '���µ�һ��
    
    
    Udate_BB_INS.Text = Mid(Date, 1, 8) + "06"    '����6��
   
   
End Sub

Public Sub Form_Exc()

    Call Gp_Sp_Excel(Me, Proc_Sc("Sc")("Spread"), lBlkcol1, lBlkcol2, lBlkrow1, lBlkrow2)
    
End Sub

Public Sub Form_Ref()
    
    Dim SMESG As String
    Dim sQuery As String
    
    If Udate_BB_DEL_TO.RawData >= Udate_BB_DEL_FR.RawData Or Udate_BB_DEL_TO.RawData = "" Then
    
        If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
        SMESG = Gf_Ms_NeceCheck(nControl)
        If SMESG = "OK" Then
        
            SMESG = Gf_Ms_NeceCheck2(mControl)
            If SMESG = "OK" Then
                
                sQuery = Gf_Ms_MakeQuery(Proc_Sc("Sc").Item("P-R"), "R", pControl)
                If Gf_Total_Display(M_CN1, Proc_Sc("Sc"), sQuery, 0, iSumCnt, iSumCol) Then
                    ss1.OperationMode = OperationModeNormal
                    Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
                End If
        
            Else
                SMESG = SMESG + " Must input according to length of item"
                Call Gp_MsgBoxDisplay(SMESG)
            End If
                
        Else
           SMESG = SMESG + " Must input necessarily"
           Call Gp_MsgBoxDisplay(SMESG)
        End If
                 
    Else
       Call MsgBox("�������ڲ����Ϲ淶!" & Chr(10) & "�������", vbExclamation + vbOKOnly, "����")
    End If
    
    
     '����������ɫ���
    Call SS1_CHANGE_COLOR
    
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

Public Sub Form_Exit()

    Unload Me
    
End Sub

Private Sub SS1_CHANGE_COLOR()

    With ss1
      
        If .MaxRows <= 0 Then
           Exit Sub
        End If
        For iCount = 1 To .MaxRows
            .Row = iCount
            
             '����������ɫ��� 2012-11-07  by  GengXueyu
            ss1.Row = .Row:       ss1.Col = SS1_URGNT_FL
            If ss1.Text = "Y" Then
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_NO, SS1_ORD_NO, .Row, .Row, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_ORD_ITEM, SS1_ORD_ITEM, .Row, .Row, &HC000&)
                 Call Gp_Sp_BlockColor(ss1, SS1_URGNT_FL, SS1_URGNT_FL, .Row, .Row, &HC000&)
            End If

        Next iCount

    End With
    
End Sub

Private Sub text_bb_ord_no_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim sQuery As String
    
    If Len(Trim(Text_BB_ORD_NO.Text)) = Text_BB_ORD_NO.MaxLength Then
    
        If Combo_ORD_ITEM.Text <> "" Then Exit Sub
        
        Text_BB_ORD_NO.Text = StrConv(Text_BB_ORD_NO.Text, vbUpperCase)
        sQuery = " SELECT ORD_ITEM FROM CP_PRC WHERE ORD_NO = '" & Trim(Text_BB_ORD_NO.Text) & "'"
        Call Gf_ComboAdd(M_CN1, Combo_ORD_ITEM, sQuery)
       
       'If Combo_ORD_ITEM.ListCount <> 0 Then
       '   Combo_ORD_ITEM.ListIndex = 0
       'End If
    Else
    
        Combo_ORD_ITEM.Clear
        
    End If

End Sub

Private Sub Text_BB_PROD_CD_Change()

    Select Case Text_BB_PROD_CD.Text
         Case "S", "s", "SL"
             Text_BB_PROD_CD.Text = "SL"
         Case "P", "p", "PP"
             Text_BB_PROD_CD.Text = "PP"
         Case "H", "h", "HC"
             Text_BB_PROD_CD.Text = "HC"
         Case "", "**"
             Text_BB_PROD_CD.Text = ""
         Case Else
             Text_BB_PROD_CD.Text = ""
             Call MsgBox("��Ʒ�������" & Chr(10) & "�����Ϲ淶! �������", vbExclamation + vbOKOnly, "����")
    End Select
           
End Sub

Private Sub Text_BB_PROD_CD_DblClick()

    Call Text_BB_PROD_CD_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub Text_BB_PROD_CD_KeyUp(KeyCode As Integer, Shift As Integer)
   
   Text_BB_PROD_CD_mate = ""
   
   If KeyCode = vbKeyF4 Then
 
        DD.sWitch = "MS"
        DD.sKey = "B0005"

        DD.rControl.Add Item:=Text_BB_PROD_CD
        DD.rControl.Add Item:=Text_BB_PROD_CD_mate
   
        DD.nameType = "2"
        'DD.nameType="1" ���������Ʋ�ѯ
        'DD.nameType="2" ��Ӣ�����Ʋ�ѯ
       
        Call Gf_Common_DD(M_CN1, KeyCode)

        'Call Gf_Customer_DD(M_CN1, KeyCode)
        'Gf_Customer_DD() ���ڿͻ�����

        Exit Sub
        
    End If

    If Len(Trim(Text_BB_PROD_CD.Text)) = Text_BB_PROD_CD.MaxLength Then
       '  Gf_ComnNAME_Find( �����ַ���, DD.sKEy���� ,DD.nameType)
       ' Gf_CustNameFind( �����ַ���, �ͻ���������,DD.nameType)
        Text_BB_PROD_CD_mate.Text = Gf_ComnNameFind(M_CN1, "B0005", Text_BB_PROD_CD.Text, 2)
    Else
        Text_BB_PROD_CD_mate.Text = ""
    End If
    
End Sub

Private Sub TXT_CFM_MILL_PLT_DblClick()

    Call txt_cfm_mill_plt_KeyUp(vbKeyF4, 0)
    
End Sub

Private Sub txt_cfm_mill_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=txt_cfm_mill_plt
        
        DD.nameType = "2"
        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub

    End If

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

    If ss1.MaxRows < 1 Or Row = 0 Or Row = -999 Or ss1.MaxRows = Row Then Exit Sub
    
        Unload ACA1030C
        Load ACA1030C
        
        ACA1030C.txt_prod_cd = ACA1120C.Text_BB_PROD_CD
        
        ss1.Row = Row
        ss1.Col = 1
        ACA1030C.txt_ord_no.Text = Trim(ss1.Value)
        
        ss1.Row = Row
        ss1.Col = 2
        ACA1030C.cbo_ord_item.Text = Trim(ss1.Value)
        
        ACA1030C.Active_CForm = "ACA1030C"
        ACA1030C.Show
        ACA1030C.SetFocus
    
End Sub