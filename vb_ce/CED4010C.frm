VERSION 5.00
Object = "{A5CC20C4-B5F5-11CD-98EC-0020AF234C9D}#4.1#0"; "cstext32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Begin VB.Form CED4010C 
   Caption         =   "ȷ��������ҵ��������ָʾ_CED4010C"
   ClientHeight    =   9615
   ClientLeft      =   885
   ClientTop       =   4215
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_search_slabno 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   315
      Left            =   11715
      MaxLength       =   10
      TabIndex        =   22
      ToolTipText     =   "�س�����"
      Top             =   540
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.TextBox target_y 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   21
      Tag             =   "����"
      Top             =   9300
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox to_y 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1590
      MaxLength       =   2
      TabIndex        =   20
      Tag             =   "����"
      Top             =   9300
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox from_y 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   570
      MaxLength       =   2
      TabIndex        =   19
      Tag             =   "����"
      Top             =   9300
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.TextBox txt_to 
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
      Left            =   7275
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   540
      Width           =   1365
   End
   Begin VB.TextBox txt_target 
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
      Left            =   10050
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   540
      Width           =   1365
   End
   Begin VB.TextBox txt_from 
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
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   540
      Width           =   1365
   End
   Begin VB.TextBox TXT_PLT 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "����"
      Top             =   540
      Width           =   540
   End
   Begin VB.TextBox TXT_PLT_NAME 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2055
      MaxLength       =   50
      TabIndex        =   0
      Tag             =   "����"
      Top             =   540
      Width           =   2025
   End
   Begin Threed.SSPanel SSPsend 
      Height          =   315
      Left            =   13320
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
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
      Caption         =   "���´�"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin Threed.SSPanel SSPpdt 
      Height          =   315
      Left            =   14280
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
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
      Caption         =   "������"
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
   End
   Begin CSTextLibCtl.sidbEdit SDB_SLAB_EDT_SEQ 
      Height          =   315
      Left            =   3090
      TabIndex        =   4
      Tag             =   "¯�α��ƺ�"
      Top             =   540
      Visible         =   0   'False
      Width           =   375
      _Version        =   262145
      _ExtentX        =   661
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   ""
      Text            =   " 0"
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   16
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
      NumDecDigits    =   0
      NumIntDigits    =   8
      Undo            =   0
      Data            =   0
   End
   Begin CSTextLibCtl.sidbEdit SDB_PRC_LINE 
      Height          =   315
      Left            =   3510
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   180
      _Version        =   262145
      _ExtentX        =   317
      _ExtentY        =   556
      _StockProps     =   125
      Text            =   " 0"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.76
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderEffect    =   2
      DataProperty    =   2
      Modified        =   0   'False
      HideSelection   =   -1  'True
      RawData         =   "1"
      Text            =   " 0"
      StartText.x     =   3
      StartText.y     =   2
      FirstVisPos     =   0
      HiAnchor        =   0
      HiNew           =   0
      CaretHeight     =   16
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
      NumDecDigits    =   0
      NumIntDigits    =   5
      Undo            =   0
      Data            =   1
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   60
      Top             =   540
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
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
   End
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   4125
      Top             =   540
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "��ʼ������"
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
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   8655
      Top             =   540
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      Caption         =   "Ŀ�������"
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
      Left            =   6885
      Top             =   540
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Caption         =   "->"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   435
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   767
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_move 
         Height          =   330
         Left            =   285
         TabIndex        =   8
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "�� ��"
      End
      Begin Threed.SSOption opt_delete 
         Height          =   330
         Left            =   1350
         TabIndex        =   9
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
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
         Caption         =   "ɾ ��"
      End
      Begin Threed.SSOption opt_sent 
         Height          =   285
         Left            =   135
         TabIndex        =   10
         Top             =   -165
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
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
         Caption         =   "�� ��"
      End
      Begin Threed.SSOption opt_cancel 
         Height          =   285
         Left            =   1110
         TabIndex        =   11
         Top             =   -165
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   503
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
         Caption         =   "ȡ ��"
      End
      Begin Threed.SSOption opt_cnf 
         Height          =   330
         Left            =   2385
         TabIndex        =   23
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
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
         Caption         =   "��������ָʾ"
      End
   End
   Begin Threed.SSPanel SSPanel2 
      Height          =   435
      Left            =   4110
      TabIndex        =   15
      Top             =   60
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   767
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_target 
         Height          =   330
         Left            =   5955
         TabIndex        =   16
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ŀ�������"
      End
      Begin Threed.SSOption opt_from 
         Height          =   330
         Left            =   1410
         TabIndex        =   17
         Top             =   60
         Width           =   1305
         _ExtentX        =   2302
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
         Caption         =   "��ʼ������"
      End
      Begin Threed.SSOption opt_to 
         Height          =   330
         Left            =   3150
         TabIndex        =   18
         Top             =   60
         Width           =   1365
         _ExtentX        =   2408
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
         Caption         =   "��ֹ������"
      End
   End
   Begin VB.TextBox TXT_MPLATE_NO 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10155
      MaxLength       =   12
      TabIndex        =   6
      Tag             =   "¯�ι�����"
      Top             =   75
      Visible         =   0   'False
      Width           =   1395
   End
   Begin FPSpread.vaSpread ss1 
      Height          =   8370
      Left            =   45
      TabIndex        =   24
      Top             =   900
      Width           =   15195
      _Version        =   393216
      _ExtentX        =   26802
      _ExtentY        =   14764
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
      MaxCols         =   38
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "CED4010C.frx":0000
   End
End
Attribute VB_Name = "CED4010C"
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
'-- Program Name      ָʾ����
'-- Program ID        CKG2010C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2007.11.19
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
Dim Mode As String

Dim pContro1 As New Collection      'Master Primary Key Collection
Dim nContro1 As New Collection      'Master Necessary Collection
Dim mContro1 As New Collection      'Master Maxlength check Collection
Dim iContro1 As New Collection      'Master Insert Collection
Dim rContro1 As New Collection      'Master Refer Collection
Dim cContro1 As New Collection      'Master Copy Collection
Dim aContro1 As New Collection      'Master -> Spread Collection
Dim lContro1 As New Collection      'Master Lock Collection

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

Dim sSlab_Edt_Seq_Fr As String
Dim sSlab_Edt_Seq_To As String
Dim sSlab_Edt_Seq_Tg As String

Private Sub Form_Define()
        
    Dim i As Integer
    
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

             Call Gp_Ms_Collection(TXT_PLT, "p", "n", "m", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    Call Gp_Ms_Collection(SDB_SLAB_EDT_SEQ, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    For i = 1 To ss1.MaxCols
        Call Gp_Sp_Collection(ss1, i, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Next i
  
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="CED4010C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 2, True)
    Call Gp_Sp_ColHidden(ss1, 3, True)
    Call Gp_Sp_ColHidden(ss1, 4, True)
    Call Gp_Sp_ColHidden(ss1, 5, True)
    Call Gp_Sp_ColHidden(ss1, 33, True)
    Call Gp_Sp_ColHidden(ss1, 38, True)

    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
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
    
    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(sc1.Item("Spread"), False)
    Call Gp_Sp_ReadOnlySet(sc1.Item("Spread"))
    Call Gf_Sp_Cls(sc1)
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    
    TXT_PLT.Text = "C3"
    
    Call txt_plt_KeyUp(0, 0)
    
    Screen.MousePointer = vbDefault
    
    txt_search_slabno.Text = "����������"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "E-System.INI", Me.Name)
    
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
    
    If Gf_Sp_Cls(sc1) Then
        
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        MDIMain.MenuTool.Buttons(4).Enabled = True
        TXT_PLT.Text = "C3"
        Call txt_plt_KeyUp(0, 0)
        opt_cnf.Value = False
        opt_sent.Value = False
        opt_cancel.Value = False
        opt_move.Value = False
        opt_delete.Value = False
        opt_from.Value = False
        opt_to.Value = False
        opt_target.Value = False
        opt_sent.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_from.ForeColor = &H808080
        opt_to.ForeColor = &H808080
        opt_target.ForeColor = &H808080
        txt_from = ""
        from_y.Text = ""
        txt_to = ""
        to_y.Text = ""
        txt_target = ""
        target_y.Text = ""
        TXT_MPLATE_NO = ""
        sSlab_Edt_Seq_Fr = 0
        sSlab_Edt_Seq_To = 0
        sSlab_Edt_Seq_Tg = 0
        
    End If
    
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = False                'Excel
    End With
    
End Sub

Public Sub Form_Ref()

    Dim sTemp As String
    Dim sL2_Send As String
    Dim sSlab_No As String
    Dim sPrc_Sts As String
    Dim iRow As Integer
    Dim iCol As Integer

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
       
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        ss1.OperationMode = OperationModeNormal
        
        sSlab_Edt_Seq_Fr = 0
        sSlab_Edt_Seq_To = 0
        sSlab_Edt_Seq_Tg = 0
    
        With MDIMain.MenuTool
            .Buttons(7).Enabled = False                 'Row Insert
            .Buttons(8).Enabled = False                 'Row Delete
            .Buttons(9).Enabled = False                 'Row Cancel
            .Buttons(11).Enabled = False                'Copy
            .Buttons(12).Enabled = False                'Paste
            .Buttons(14).Enabled = True                 'Excel
        End With
        
    End If

End Sub

Public Sub Form_Pro()

    Dim mResult As String
    Dim sMsg As String
    
    Mode = ""

    If opt_move = True Then
        
        If Not ((from_y.Text = "Y" And to_y.Text = "Y" And target_y.Text = "Y") _
        Or (from_y.Text = "" And to_y.Text = "" And target_y.Text = "")) Then
            MsgBox ("���´��ָʾ��δ�´��ָʾ���ܻ���һ�������")
            Exit Sub
        End If
        
    End If
 
    If opt_move = True Then
        
        Mode = "M"
        If txt_from.Text <> "" And txt_to.Text <> "" And txt_target.Text <> "" Then  '˳����
            sMsg = "ȷ��Ҫ�Ѱ�����(" + txt_from.Text + ")->(" + txt_to.Text + ")" + "����������(" + txt_target.Text + ")�����"
        Else
            sMsg = "����������ʼ�����š���ֹ�����ź�Ŀ������ţ�"
            Call Gp_MsgBoxDisplay(sMsg)
            Exit Sub
        End If
    
        mResult = MsgBox(sMsg, vbYesNo)
        
        If mResult = vbYes Then
            If Gp_Process_Exec = "" Then
                MsgBox ("��ҵָʾ������� ��")
                Call Form_Ref
            Else
                MsgBox (Gp_Process_Exec + " ��ҵָʾ����ʧ�ܣ�")
            End If
        End If
     
     End If
 
     If opt_delete = True Then
        
        Mode = "D"
        If txt_from.Text = "" Then
           sMsg = "����������ʼ�����ţ�"
           Call Gp_MsgBoxDisplay(sMsg)
           Exit Sub
        End If
        
        If txt_to.Text = "" Then
           sMsg = "����������ֹ�����ţ�"
           Call Gp_MsgBoxDisplay(sMsg)
           Exit Sub
        End If
        
        sMsg = "ȷ��Ҫɾ��ѡ������(" + txt_from.Text + ")" + ")��"
        If txt_to.Text <> "" Then
           sMsg = "ȷ��Ҫɾ��ѡ������(" + txt_from.Text + ")->(" + txt_to.Text + ")��"
        End If
        
        mResult = MsgBox(sMsg, vbYesNo)
        If mResult = vbYes Then
           If Gp_Process_Exec = "" Then
              MsgBox ("��ҵָʾɾ����� ��")
              Call Form_Ref
           Else
              MsgBox (Gp_Process_Exec + " ��ҵָʾɾ��ʧ�ܣ�")
           End If
        End If
     
     End If
 
     If opt_cnf = True Then
        
        Mode = "F"
        
        If txt_from.Text = "" Then
           sMsg = "����������ʼ�����ţ�"
           Call Gp_MsgBoxDisplay(sMsg)
           Exit Sub
        End If
        
        sMsg = "ȷ��Ҫָʾѡ������(" + txt_from.Text + ")" + ")��"
        If txt_to.Text <> "" Then
           sMsg = "ȷ��Ҫָʾѡ������(" + txt_from.Text + ")->(" + txt_to.Text + ")��"
        End If
        
        mResult = MsgBox(sMsg, vbYesNo)
        If mResult = vbYes Then
           If Gp_Process_Exec = "" Then
              MsgBox ("��ҵָʾ��� ��")
              Call Form_Ref
           Else
              MsgBox (Gp_Process_Exec + " ��ҵָʾʧ�ܣ�")
           End If
        End If
     
    End If
 
    opt_cnf.Value = False
    opt_sent.Value = False
    opt_cancel.Value = False
    opt_move.Value = False
    opt_delete.Value = False
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    
    opt_cnf.ForeColor = &H808080
    opt_sent.ForeColor = &H808080
    opt_move.ForeColor = &H808080
    opt_delete.ForeColor = &H808080
    opt_cancel.ForeColor = &H808080
    opt_from.ForeColor = &H808080
    opt_to.ForeColor = &H808080
    opt_target.ForeColor = &H808080
    txt_from = ""
    from_y.Text = ""
    txt_to = ""
    to_y.Text = ""
    txt_target = ""
    target_y.Text = ""
    TXT_MPLATE_NO = ""
            
    With MDIMain.MenuTool
        .Buttons(7).Enabled = False                 'Row Insert
        .Buttons(8).Enabled = False                 'Row Delete
        .Buttons(9).Enabled = False                 'Row Cancel
        .Buttons(11).Enabled = False                'Copy
        .Buttons(12).Enabled = False                'Paste
        .Buttons(14).Enabled = True                 'Excel
    End With

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

Private Sub opt_cancel_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_cancel.Value = True Then
        opt_cancel.ForeColor = &HFF&
        opt_cnf.ForeColor = &H808080
        opt_sent.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_from.Enabled = True
        opt_from.Value = True
        opt_to.Enabled = False
        opt_target.Enabled = False
    Else
        opt_cancel.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_cnf_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_cnf.Value = True Then
    
        opt_cnf.ForeColor = &HFF&
        opt_delete.ForeColor = &H808080
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_from.Enabled = True
        opt_from.Value = True
        opt_to.Enabled = True
        opt_target.Enabled = False
    Else
        opt_delete.ForeColor = &H808080
    End If
    
    opt_from.Value = True
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0

End Sub

Private Sub opt_delete_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_delete.Value = True Then
    
        opt_delete.ForeColor = &HFF&
        opt_cnf.ForeColor = &H808080
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_from.Enabled = True
        opt_from.Value = True
        opt_to.Enabled = True
        opt_target.Enabled = False
    Else
        opt_delete.ForeColor = &H808080
    End If
    
    opt_from.Value = True
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_from_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_from.Value = True Then
        opt_from.ForeColor = &HFF&
        opt_to.ForeColor = &H808080
        opt_target.ForeColor = &H808080
    Else
        opt_from.ForeColor = &H808080
    End If
    
End Sub

Private Sub opt_move_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_move.Value = True Then
        opt_move.ForeColor = &HFF&
        opt_cnf.ForeColor = &H808080
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_from.Enabled = True
        opt_from.Value = True
        opt_to.Enabled = True
        opt_target.Enabled = True
    Else
        opt_move.ForeColor = &H808080
    End If
    
    opt_from.Value = True
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_sent_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_sent.Value = True Then
        opt_sent.ForeColor = &HFF&
        opt_cnf.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
        opt_from.Enabled = False
        opt_to.Enabled = True
        opt_to.Value = True
        opt_target.Enabled = False
    Else
        opt_sent.ForeColor = &H808080
    End If
    
    txt_from = ""
    txt_to = ""
    txt_target = ""
    
    sSlab_Edt_Seq_Fr = 0
    sSlab_Edt_Seq_To = 0
    sSlab_Edt_Seq_Tg = 0
    
End Sub

Private Sub opt_target_Click(Value As Integer)
    
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_target.Value = True Then
        opt_target.ForeColor = &HFF&
        opt_from.ForeColor = &H808080
        opt_to.ForeColor = &H808080
    Else
        opt_target.ForeColor = &H808080
    End If
    
End Sub

Private Sub opt_to_Click(Value As Integer)

    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_to.Value = True Then
        opt_to.ForeColor = &HFF&
        opt_from.ForeColor = &H808080
        opt_target.ForeColor = &H808080
    Else
        opt_to.ForeColor = &H808080
    End If
    
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim C, M As Integer
    Dim iRow As Integer
    Dim iCol As Integer
    Dim SEND_SLAB As String

    If Gf_Sp_Change(Proc_Sc, sc1) Then
        lBlkcol1 = 0
        lBlkcol2 = 0
        lBlkrow1 = 0
        lBlkrow2 = 0
    End If
    
    If Row < 1 Then Exit Sub
    
    ss1.Row = Row
    ss1.Col = 6
            
    If opt_from.Value = True Then
       txt_from.Text = ss1.Text
       ss1.Col = 33
       from_y.Text = ss1.Text
       
       ss1.Col = 38
       sSlab_Edt_Seq_Fr = ss1.Text
       
    ElseIf opt_to.Value = True Then
       txt_to.Text = ss1.Text
       ss1.Col = 33
       to_y.Text = ss1.Text
       
       ss1.Col = 38
       sSlab_Edt_Seq_To = ss1.Text
       
    ElseIf opt_target.Value = True Then
       txt_target.Text = ss1.Text
       ss1.Col = 33
       target_y.Text = ss1.Text
       
       ss1.Col = 38
       sSlab_Edt_Seq_Tg = ss1.Text
    End If

    ss1.Col = 33
    If (opt_sent.Value = True Or opt_delete.Value = True) Then
    
        If ss1.Text = "Y" Then
            ss1.Col = 6
            MsgBox ("������ " + ss1.Text + " ��ҵָʾ���´")
            If opt_from.Value = True Then
               txt_from.Text = ""
               from_y.Text = ""
            ElseIf opt_to.Value = True Then
               txt_to.Text = ""
               to_y.Text = ""
            ElseIf opt_target.Value = True Then
               txt_target.Text = ""
               target_y.Text = ""
            End If
            
            Exit Sub
            
        End If
        
    End If
    
    ss1.Col = 5
    If ss1.Text = "B" And (opt_sent.Value = True Or opt_move.Value = True Or opt_cancel.Value = True Or opt_delete.Value = True) Then
        
        ss1.Col = 6
        MsgBox ("������ " + ss1.Text + " ����¯�����ܵ�����")
        
        If opt_from.Value = True Then
           txt_from.Text = ""
           from_y.Text = ""
        
        ElseIf opt_to.Value = True Then
           txt_to.Text = ""
           to_y.Text = ""
        
        ElseIf opt_target.Value = True Then
           txt_target.Text = ""
           target_y.Text = ""
        End If
        
        Exit Sub
    End If
                              
End Sub

Private Sub SSPanel1_Click()
    
    opt_sent.Value = False
    opt_cancel.Value = False
    opt_move.Value = False
    opt_delete.Value = False
    opt_from.Value = False
    opt_to.Value = False
    opt_target.Value = False
    opt_sent.ForeColor = &H808080
    opt_move.ForeColor = &H808080
    opt_delete.ForeColor = &H808080
    opt_cancel.ForeColor = &H808080
    opt_from.ForeColor = &H808080
    opt_to.ForeColor = &H808080
    opt_target.ForeColor = &H808080
    
End Sub

Private Sub txt_plt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF4 Then
    
        DD.sWitch = "MS"
        DD.sKey = "C0001"
        DD.rControl.Add Item:=TXT_PLT
        DD.rControl.Add Item:=TXT_PLT_NAME
        
        DD.nameType = "2"
        
        Call Gf_Common_DD(M_CN1, KeyCode)
        
        Exit Sub
        
    End If

    If Len(Trim(TXT_PLT.Text)) = TXT_PLT.MaxLength Then
        TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", Trim(TXT_PLT.Text), 2)
    Else
        TXT_PLT_NAME.Text = ""
    End If

End Sub

Public Function Gp_Process_Exec() As String

On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim ret_Result_ErrMsg As String
    Dim sQuery As String
    Dim iRow As Integer
    Dim adoCmd As ADODB.Command
    
    Dim sSlab_Seq_Fr As String
    Dim sSlab_Seq_To As String
    Dim sSlab_Seq_Tg As String
    
    Screen.MousePointer = vbHourglass
    
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    sSlab_Seq_Fr = sSlab_Edt_Seq_Fr
    sSlab_Seq_To = sSlab_Edt_Seq_To
    sSlab_Seq_Tg = sSlab_Edt_Seq_Tg
    
    Select Case Mode
    
        Case "M"  'MOVE
        
            sQuery = "{call CGG2044P ('" + "M" + "','" + sSlab_Seq_Fr + "','" + sSlab_Seq_To + "','" + sSlab_Seq_Tg + "',?)}"
        
        Case "D"  'DELETE
        
            sQuery = "{call CGG2043P ('" + sSlab_Seq_Fr + "','" + sSlab_Seq_To + "','" + "M" + "',?)}"
            
        Case "F"  'CONFIRM
        
            sQuery = "{call AFZ1091P ('C3','" + sSlab_Seq_Fr + "','" + sSlab_Seq_To + "',?)}"
        
    End Select
    
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
        Gp_Process_Exec = sErrMessg
        Set adoCmd = Nothing
        Exit Function
    End If
    
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Gp_Process_Exec = ""
    Exit Function

Process_Exec_ERROR:
    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Gp_Process_Exec = "Process_Exec_ERROR"
    Err.Raise Err.Number, Err.Description & sQuery
    
End Function


Private Sub txt_search_slabno_Click()

    If txt_search_slabno.Text = "����������" Then
        txt_search_slabno.Text = ""
    End If
    
End Sub

Private Sub txt_search_slabno_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer
    
    If KeyCode = 13 Then
    
        For i = 1 To ss1.MaxRows
            ss1.Row = i
            ss1.Col = 6
            If ss1.Text = Trim(txt_search_slabno.Text) Then
               Call ss1.SetActiveCell(6, i)
               Exit For
            End If
        Next i
        
    End If
        
End Sub