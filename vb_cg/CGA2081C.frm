VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form CGA2081C 
   Caption         =   "�������з�ʵ��¼�����_CGA2081C"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin FPSpread.vaSpread ss1 
      Height          =   8490
      Left            =   90
      TabIndex        =   0
      Top             =   780
      Width           =   15285
      _Version        =   393216
      _ExtentX        =   26961
      _ExtentY        =   14975
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
      MaxCols         =   25
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CGA2081C.frx":0000
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   600
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   1058
      _Version        =   196609
      BackColor       =   14737632
      Begin VB.TextBox TXT_FLAG 
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
         Left            =   14670
         MaxLength       =   1
         TabIndex        =   11
         Top             =   255
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txt_mat_no 
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
         Left            =   7710
         MaxLength       =   10
         TabIndex        =   10
         Top             =   150
         Width           =   1260
      End
      Begin VB.ComboBox CBO_PLT 
         BackColor       =   &H00C0FFFF&
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
         Height          =   315
         ItemData        =   "CGA2081C.frx":1F7F
         Left            =   14880
         List            =   "CGA2081C.frx":1F86
         TabIndex        =   4
         Text            =   "C1"
         Top             =   825
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.ComboBox CBO_LINE 
         BackColor       =   &H00C0FFFF&
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
         ItemData        =   "CGA2081C.frx":1F8D
         Left            =   14895
         List            =   "CGA2081C.frx":1F94
         TabIndex        =   3
         Text            =   "1"
         Top             =   1140
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.ComboBox CBO_SHIFT_REF 
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
         ItemData        =   "CGA2081C.frx":1F9B
         Left            =   5580
         List            =   "CGA2081C.frx":1FAB
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   150
         Width           =   735
      End
      Begin InDate.ULabel ULabel10 
         Height          =   315
         Left            =   135
         Tag             =   "������FROM"
         Top             =   150
         Width           =   1275
         _ExtentX        =   2249
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
      Begin CSTextLibCtl.sitxEdit TXT_From_Date 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   150
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__"
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
         Mask            =   "____-__-__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin CSTextLibCtl.sitxEdit TXT_To_Date 
         Height          =   315
         Left            =   2805
         TabIndex        =   6
         Tag             =   "������TO"
         Top             =   150
         Width           =   1185
         _Version        =   262145
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   125
         Text            =   "____-__-__ __-__-__"
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
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   ""
         Text            =   "____-__-__"
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
         Mask            =   "____-__-__"
         Justification   =   1
         CharacterTable  =   ""
         BorderStyle     =   0
         MaxLength       =   0
         ValidateMask    =   0   'False
      End
      Begin InDate.ULabel ULabel5 
         Height          =   315
         Left            =   9195
         Top             =   150
         Width           =   1185
         _ExtentX        =   2090
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
      Begin CSTextLibCtl.sidbEdit SDB_TOT_WGT 
         Height          =   315
         Left            =   10410
         TabIndex        =   7
         Tag             =   "�ϸ�����"
         Top             =   150
         Width           =   1470
         _Version        =   262145
         _ExtentX        =   2593
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
         ReadOnly        =   -1  'True
         FocusSelect     =   -1  'True
         Modified        =   -1  'True
         HideSelection   =   -1  'True
         RawData         =   "0.000"
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
         NumIntDigits    =   5
         ShowZero        =   0   'False
         MaxValue        =   99999.999
         MinValue        =   0
         Undo            =   0
         Data            =   0
      End
      Begin InDate.ULabel ULabel13 
         Height          =   315
         Left            =   4260
         Top             =   150
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         Caption         =   "�зϰ��"
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
         Left            =   6495
         Top             =   150
         Width           =   1170
         _ExtentX        =   2064
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
      Begin Threed.SSOption OPT_SCRAP 
         Height          =   330
         Left            =   13725
         TabIndex        =   12
         Top             =   150
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
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
         Caption         =   "���з�"
      End
      Begin Threed.SSOption OPT_SCRAP_WAIT 
         Height          =   330
         Left            =   12735
         TabIndex        =   13
         Top             =   150
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
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
         Caption         =   "���з�"
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ton"
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
         Left            =   11955
         TabIndex        =   9
         Top             =   210
         Width           =   360
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "~"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   2655
         TabIndex        =   8
         Top             =   270
         Width           =   195
      End
   End
End
Attribute VB_Name = "CGA2081C"
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
'-- Program Name      �ϸ�ʵ��
'-- Program ID        AGF2080C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2003.7.23
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

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

Dim pContro2 As New Collection      'Master Primary Key Collection
Dim nContro2 As New Collection      'Master Necessary Collection
Dim mContro2 As New Collection      'Master Maxlength check Collection
Dim iContro2 As New Collection      'Master Insert Collection
Dim rContro2 As New Collection      'Master Refer Collection
Dim cContro2 As New Collection      'Master Copy Collection
Dim aContro2 As New Collection      'Master -> Spread Collection
Dim lContro2 As New Collection      'Master Lock Collection

Dim pColumn1 As New Collection      'Spread Primary Key Collection
Dim nColumn1 As New Collection      'Spread necessary Column Collection
Dim mColumn1 As New Collection      'Spread Maxlength check Column Collection
Dim iColumn1 As New Collection      'Spread Insert Column Collection
Dim aColumn1 As New Collection      'Master -> Spread Column Collection
Dim lColumn1 As New Collection      'Spread Lock Column Collection

Dim Mc1 As New Collection           'Master Collection
Dim sc1 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
     Call Gp_Ms_Collection(TXT_From_Date, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(TXT_To_Date, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     Call Gp_Ms_Collection(CBO_SHIFT_REF, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
        Call Gp_Ms_Collection(txt_mat_no, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(TXT_FLAG, "p", " ", " ", " ", " ", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
     
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 8, "p", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 22, "p", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CGA2081C.P_SREFER", Key:="P-R"
    sc1.Add Item:="CGA2081C.P_SONEROW", Key:="P-O"
    sc1.Add Item:="AGF2080C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 1, True)
    Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 2, True)
    Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 3, True)
    Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 7, True)
    Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 10, True)
   Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 11, True)
   Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 17, True)
   Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 18, True)
   Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 22, True)
   Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 23, True)
   Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 24, True)
    
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

    Call Gp_Sp_Setting(Proc_Sc("Sc")("Spread"))
    
    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(Proc_Sc("Sc")("Spread"), "CG-System.INI", Me.Name)
    
'    Call Gp_Sp_HdColColor(Proc_Sc("Sc")("Spread"), 2)
        
    TXT_From_Date.RawData = Format(Date, "yyyymm") + "01"
    TXT_To_Date.RawData = Format(Date, "yyyymmdd")
    
    OPT_SCRAP_WAIT.Value = True
        
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If

    Call Gp_Sp_ColSet(Proc_Sc("Sc")("Spread"), "CG-System.INI", Me.Name)

    Set pContro1 = Nothing
    Set nContro1 = Nothing
    Set iContro1 = Nothing
    Set rContro1 = Nothing
    Set cContro1 = Nothing
    Set aContro1 = Nothing
    Set lContro1 = Nothing
    Set mContro1 = Nothing

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

Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

Public Sub Form_Cls()

        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gf_Sp_Cls(Proc_Sc("SC"))
'        Call Gp_Ms_Cls(Mc1("pControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("rControl"), False)
End Sub

Public Sub Form_Ref()
    
    Dim iRow As Integer
    On Error Resume Next

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
 
    SDB_TOT_WGT.Value = 0
    
    If Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl")) Then
        If ss1.MaxRows > 0 Then
           For iRow = 1 To ss1.MaxRows
               ss1.ROW = iRow
               ss1.Col = 16
               SDB_TOT_WGT.Value = SDB_TOT_WGT.Value + Val(ss1.Value)
           Next iRow
        End If
        
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        If TXT_FLAG = "1" Then
           MDIMain.MenuTool.Buttons(8).Enabled = False
        ElseIf TXT_FLAG = "2" Then
           MDIMain.MenuTool.Buttons(8).Enabled = True
        End If
    End If
  
End Sub

Public Sub Form_Pro()

    Dim sMesg As String
    Dim iCount As Integer

    For iCount = 1 To ss1.MaxRows
        Select Case Trim(Gf_Sp_RcvData(ss1, 0, iCount))
            
            Case "Update"
            
                  With ss1
                      .Col = 4
                      If Not Gp_DateCheck(.Text, "X") Then
                         Call Gp_MsgBoxDisplay("����ȷ�����з�����")
                         Exit Sub
                      End If
                      .Col = 5
                      If Trim(.Text) <> "1" And Trim(.Text) <> "2" And Trim(.Text) <> "3" Then
                        MsgBox "����ȷ�����зϰ��!", vbCritical, "ϵͳ��ʾ��Ϣ"
                        Exit Sub
                      End If
                  End With
                  
        End Select
    Next iCount
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("SC"), Mc1) Then
       Call Gf_Sp_Refer(M_CN1, Proc_Sc("Sc"), Mc1, Mc1("nControl"), Mc1("mControl"))
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If

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

Private Sub OPT_SCRAP_WAIT_Click(Value As Integer)
    OPT_SCRAP_WAIT.ForeColor = &HFF&
    OPT_SCRAP.ForeColor = &H808080
    TXT_FLAG.Text = "1"
    ss1.MaxRows = 0
    ULabel10.Caption = "��������"
End Sub

Private Sub OPT_SCRAP_Click(Value As Integer)
    OPT_SCRAP_WAIT.ForeColor = &H808080
    OPT_SCRAP.ForeColor = &HFF&
    TXT_FLAG.Text = "2"
    ss1.MaxRows = 0
    ULabel10.Caption = "�з�����"
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Change(ByVal Col As Long, ByVal ROW As Long)
Dim RES As String
If ss1.ActiveCol = 10 Then
    ss1.Col = ss1.ActiveCol
    ss1.ROW = ss1.ActiveRow
    RES = Trim(ss1.Text)
    If Len(Trim(ss1.Text)) = 1 Then
        ss1.Col = 11
        ss1.Text = Gf_ComnNameFind(M_CN1, "F0011", RES, 1)
    Else
        ss1.Col = 11
        ss1.Text = ""
    End If
End If
End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal ROW As Long)
If Trim(CBO_SHIFT_REF.Text) = "" Then
   MsgBox "����ȷ�����зϰ��!", vbCritical, "ϵͳ��ʾ��Ϣ"
   Exit Sub
End If

    With ss1
        .ROW = ROW
        .Col = 4
        .Text = Format(Now, "YYYY-MM-DD")
        .Col = 5
        .Text = CBO_SHIFT_REF.Text
        .Col = 21
        .Text = sUserID
        If Gf_Sc_Authority(sAuthority, "U") Then
           .Col = 0
            Select Case Trim(.Text)
                   Case "Input", "Update", "Delete"
                   Case Else
                        .Text = "Update"
            End Select
        End If

    End With
    
End Sub

Private Sub ss1_EditMode(ByVal Col As Long, ByVal ROW As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
    End If
    
End Sub

Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub

    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True

End Sub

Private Sub ss1_KeyUp(KeyCode As Integer, Shift As Integer)
If ss1.ActiveCol = 10 Then
    If KeyCode = vbKeyF4 Then
        ss1.ROW = ss1.ActiveRow
        ss1.Col = ss1.ActiveCol
        DD.sWitch = "MS"
        DD.sKey = "F0011"
        DD.rControl.Add Item:=ss1.Text
        ss1.Col = 11
        DD.rControl.Add Item:=ss1.Text

        DD.nameType = "1"

        Call Gf_Common_DD(M_CN1, KeyCode)
        Exit Sub
    End If
End If
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub TXT_From_Date_DblClick()
    TXT_From_Date.RawData = Gf_DTSet(M_CN1, "D")
    TXT_To_Date.RawData = Gf_DTSet(M_CN1, "D")
End Sub

Private Sub TXT_To_Date_DblClick()
    TXT_To_Date.RawData = Gf_DTSet(M_CN1, "D")
End Sub
