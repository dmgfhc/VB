VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form AGG2060C 
   Caption         =   "���ּƻ���ѯ����_AGG2060C"
   ClientHeight    =   9675
   ClientLeft      =   1110
   ClientTop       =   2025
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
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
      ItemData        =   "AGG2060C.frx":0000
      Left            =   3495
      List            =   "AGG2060C.frx":000D
      TabIndex        =   2
      Top             =   9450
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox CBO_PLT 
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
      ItemData        =   "AGG2060C.frx":001D
      Left            =   1380
      List            =   "AGG2060C.frx":0027
      TabIndex        =   1
      Text            =   "C1"
      Top             =   9450
      Visible         =   0   'False
      Width           =   735
   End
   Begin Threed.SSCommand Cmd_Seq_Update 
      Height          =   360
      Left            =   4680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   9420
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      _Version        =   196609
      Font3D          =   1
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "����ָʾ"
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   210
      Top             =   9450
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
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
   Begin InDate.ULabel ULabel43 
      Height          =   315
      Left            =   2610
      Top             =   9450
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9075
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   16007
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "������ҵ�ƻ�"
      TabPicture(0)   =   "AGG2060C.frx":0033
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSSplitter1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "���ִ�ƻ�"
      TabPicture(1)   =   "AGG2060C.frx":004F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SSSplitter2"
      Tab(1).ControlCount=   1
      Begin SSSplitter.SSSplitter SSSplitter1 
         Height          =   8775
         Left            =   0
         TabIndex        =   4
         Top             =   300
         Width           =   15345
         _ExtentX        =   27067
         _ExtentY        =   15478
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   1
         PaneTree        =   "AGG2060C.frx":006B
         Begin Threed.SSFrame SSFrame1 
            Height          =   555
            Left            =   15
            TabIndex        =   5
            Top             =   15
            Width           =   15315
            _ExtentX        =   27014
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737632
            Begin VB.TextBox txt_Slab_no 
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
               Left            =   2065
               TabIndex        =   6
               Top             =   120
               Width           =   1575
            End
            Begin InDate.ULabel ULabel3 
               Height          =   315
               Left            =   480
               Top             =   120
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   556
               Caption         =   "��ʼ������"
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
            Begin Threed.SSCommand Cmd_Edit 
               Height          =   330
               Left            =   7620
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   120
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   582
               _Version        =   196609
               Font3D          =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "��������"
            End
            Begin Threed.SSCommand Cmd_exl 
               Height          =   330
               Left            =   4710
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   150
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   582
               _Version        =   196609
               Font3D          =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "���ּƻ���"
            End
         End
         Begin FPSpread.vaSpread ss1 
            Height          =   8130
            Left            =   15
            TabIndex        =   8
            Top             =   630
            Width           =   15315
            _Version        =   393216
            _ExtentX        =   27014
            _ExtentY        =   14340
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
            MaxCols         =   47
            MaxRows         =   1
            Protect         =   0   'False
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AGG2060C.frx":00BD
         End
      End
      Begin SSSplitter.SSSplitter SSSplitter2 
         Height          =   8775
         Left            =   -75000
         TabIndex        =   9
         Top             =   300
         Width           =   15345
         _ExtentX        =   27067
         _ExtentY        =   15478
         _Version        =   196609
         SplitterBarWidth=   3
         BorderStyle     =   1
         PaneTree        =   "AGG2060C.frx":132A
         Begin FPSpread.vaSpread ss2 
            Height          =   8130
            Left            =   15
            TabIndex        =   11
            Top             =   630
            Width           =   15315
            _Version        =   393216
            _ExtentX        =   27014
            _ExtentY        =   14340
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
            MaxCols         =   22
            MaxRows         =   1
            RetainSelBlock  =   0   'False
            SpreadDesigner  =   "AGG2060C.frx":137C
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   555
            Left            =   15
            TabIndex        =   10
            Top             =   15
            Width           =   15315
            _ExtentX        =   27014
            _ExtentY        =   979
            _Version        =   196609
            BackColor       =   14737632
            Begin InDate.UDate udFmDate 
               Height          =   315
               Left            =   2070
               TabIndex        =   12
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
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
            Begin InDate.ULabel labDate 
               Height          =   315
               Left            =   480
               Tag             =   "��ѯ��FROM"
               Top             =   120
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   556
               Caption         =   "��ѯ����"
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
            Begin InDate.UDate udToDate 
               Height          =   315
               Left            =   3840
               TabIndex        =   13
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
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
            Begin Threed.SSCommand SCmd2 
               Height          =   330
               Left            =   7620
               TabIndex        =   14
               Top             =   120
               Width           =   1965
               _ExtentX        =   3466
               _ExtentY        =   582
               _Version        =   196609
               Font3D          =   1
               ForeColor       =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9.75
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "�ϴ��ƻ�"
            End
            Begin VB.Label Label1 
               BackColor       =   &H00E0E0E0&
               Caption         =   "~"
               Height          =   120
               Left            =   3630
               TabIndex        =   15
               Top             =   240
               Width           =   195
            End
         End
      End
   End
End
Attribute VB_Name = "AGG2060C"
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
'-- Program Name      ������ҵ�ƻ���ѯ����
'-- Program ID        AGG2060C
'-- Document No       Q-00-0010(Specification)
'-- Designer          KIM.S.H
'-- Coder             KIM.S.H
'-- Date              2005.8.10
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
Dim sc1 As New Collection           'Spread Collection
Dim sc2 As New Collection           'Spread Collection
Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim SE, MPLATE_NO, plate_no As String

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim lRow        As Long
Dim lRowRange   As Long

Private Sub Form_Define()

    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
   Call Gp_Ms_Collection(TXT_SLAB_NO, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
       Call Gp_Ms_Collection(CBO_PLT, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(CBO_LINE, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(udFmDate, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      Call Gp_Ms_Collection(udToDate, "p", " ", " ", " ", "r", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
      
      Call Gp_Ms_Collection(udFmDate, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
      Call Gp_Ms_Collection(udToDate, "p", " ", " ", " ", "r", " ", " ", pControl2, nControl2, mControl2, iControl2, rControl2, aControl2, lControl2)
                   
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
    
    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
   Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '��ĥ
   Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
   
    'Spread_Collection
    
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="AGG2060C.P_SREFER", Key:="P-R"
    sc1.Add Item:="AGG2060C.P_MODIFY", Key:="P-M"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    'Spread_Collection
    'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
    Call Gp_Sp_Collection(ss2, 1, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 2, "p", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 3, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) 'add by LiQian 2012-08-14 ������������
    Call Gp_Sp_Collection(ss2, 4, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) '������ 20141230
    Call Gp_Sp_Collection(ss2, 5, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 6, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 7, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 8, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 9, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 10, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 11, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 12, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 13, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2) 'add by LiQian 2012-08-14 �Ƿ�̽��
    Call Gp_Sp_Collection(ss2, 14, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 15, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 16, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 17, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 18, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 19, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 20, " ", " ", " ", "i", " ", " ", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 21, " ", " ", " ", "i", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)
    Call Gp_Sp_Collection(ss2, 22, " ", " ", " ", " ", " ", "l", pColumn2, nColumn2, mColumn2, iColumn2, aColumn2, lColumn2)


    'Spread_Collection
    
    sc2.Add Item:=ss2, Key:="Spread"
    sc2.Add Item:="AGG2060C.P_SREFER2", Key:="P-R"
    sc2.Add Item:="AGG2060C.P_MODIFY2", Key:="P-M"
    sc2.Add Item:=pColumn2, Key:="pColumn"
    sc2.Add Item:=nColumn2, Key:="nColumn"
    sc2.Add Item:=aColumn2, Key:="aColumn"
    sc2.Add Item:=mColumn2, Key:="mColumn"
    sc2.Add Item:=iColumn2, Key:="iColumn"
    sc2.Add Item:=lColumn2, Key:="lColumn"
    sc2.Add Item:=1, Key:="First"
    sc2.Add Item:=ss2.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc2, Key:="Sc2"
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
    
    MDIMain.MenuTool.Buttons(7).Enabled = False
'    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

End Sub

Private Sub Cmd_exl_Click()
    Call ExcelFl
End Sub

Private Sub Form_Activate()

    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    MDIMain.MenuTool.Buttons(7).Enabled = False
'    MDIMain.MenuTool.Buttons(8).Enabled = False
    MDIMain.MenuTool.Buttons(9).Enabled = False
    MDIMain.MenuTool.Buttons(11).Enabled = False
    MDIMain.MenuTool.Buttons(12).Enabled = False

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
    Call Gp_Ms_NeceColor(Mc2("nControl"))

    Call Gp_Sp_Setting(sc1.Item("Spread"))
    Call Gp_Sp_Setting(sc2.Item("Spread"))
    
    Call Gf_Sp_Cls(sc1)
    Call Gf_Sp_Cls(sc2)
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColGet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
'    Call Gp_Sp_ColHidden(Proc_Sc("Sc")("Spread"), 26, True)


    CBO_PLT.ListIndex = 0
    CBO_LINE.ListIndex = 0
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call Gp_Sp_ColSet(sc1.Item("Spread"), "G-System.INI", Me.Name)
    Call Gp_Sp_ColSet(sc2.Item("Spread"), "G-System.INI", Me.Name)
    
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
    Set sc1 = Nothing
    Set sc2 = Nothing
    
    Set Proc_Sc = Nothing

    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")

End Sub

Public Sub Form_Cls()

    If SSTab1.Tab = 0 Then
        If Gf_Sp_Cls(sc1) Then
            Call Gp_Ms_Cls(Mc1("rControl"))
            Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
            Call Gp_Ms_ControlLock(Mc1("lControl"), False)
        End If
    Else
        Call Gf_Sp_Cls(sc2)
        Call Gp_Ms_Cls(Mc2("rControl"))
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        Call Gp_Ms_ControlLock(Mc1("lControl"), False)
    End If
    
End Sub

Public Sub Form_Exc()

    Select Case SSTab1.Tab
           
           Case 0
     
                Call ExcelPrn
            
           Case 1
           
                Call ExcelPrn1
    
    End Select

    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

    If SSTab1.Tab = 0 Then
        If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
        
        If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
            ss1.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
'            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        End If
    Else
        
        If Gf_Sp_Refer(M_CN1, sc2, Mc2, Mc2("nControl"), Mc2("mControl")) Then
            ss2.OperationMode = OperationModeNormal
            Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
            MDIMain.MenuTool.Buttons(7).Enabled = False
'            MDIMain.MenuTool.Buttons(8).Enabled = False
            MDIMain.MenuTool.Buttons(9).Enabled = False
            MDIMain.MenuTool.Buttons(11).Enabled = False
            MDIMain.MenuTool.Buttons(12).Enabled = False
        End If
    End If
    
    Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()

    If Gf_Sp_Process(M_CN1, sc2, Mc2) Then
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

Public Sub Form_Exit()
    Unload Me
End Sub

Private Sub SCmd2_Click()
   Load LoadExcel
   LoadExcel.Show 1
End Sub

Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2

End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)
'    Dim iRow  As Integer
'
'    If Row < 1 Then Exit Sub
''    ss1.Col = 1
''    ss1.Row = Row
''    txt_Slab_no.Text = ss1.Text
'
'    If lRowRange > 0 Then
'
'        'Copy the currently selected block of cells to the clipboard
'        ss1.SetSelection 1, lRow, ss1.MaxCols, lRow + lRowRange - 1
'        ss1.ClipboardCopy
'
'        'Delete a row in the spreadsheet
'        ss1.DeleteRows lRow, lRowRange
'
'        'Insert a row into the spreadsheet
''        Row = Row + 1
'
'        If lRow < Row Then Row = Row - lRowRange
'
'        ss1.InsertRows Row, lRowRange
'
'        ss1.BlockMode = True
'        ss1.Col = 1:    ss1.Col2 = ss1.MaxRows
'        ss1.Row = 1:    ss1.Row2 = ss1.MaxRows
'        ss1.Lock = False
'        ss1.BlockMode = False
'
'        'Paste the data FROM the clipboard to the currently selected block of cells
'        ss1.SetSelection 1, Row, ss1.MaxCols, Row + lRowRange - 1
'        ss1.ClipboardPaste
'
'        For iRow = Row To Row + lRowRange - 1
'            ss1.Row = iRow:   ss1.Col = 0:    ss1.Text = "Update"
'        Next iRow
'
'        ss1.BlockMode = True
'        ss1.Col = 1:    ss1.Col2 = ss1.MaxRows
'        ss1.Row = 1:    ss1.Row2 = ss1.MaxRows
'        ss1.Lock = True
'        ss1.Col = 24:   ss1.Col2 = 24
'        ss1.Lock = False
'        ss1.BlockMode = False
'
'        lRow = 0: lRowRange = 0
'    End If
    
End Sub

Private Sub ss1_LostFocus()

    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0

End Sub

Private Sub ss1_RightClick(ByVal ClickType As Integer, ByVal Col As Long, ByVal Row As Long, ByVal MouseX As Long, ByVal MouseY As Long)

'    If Row > 0 Then
'        Set Active_Spread = Me.ss1
'        PopupMenu MDIMain.PopUp_Spread
'    End If

End Sub

Private Sub ss1_SelChange(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long, ByVal CurCol As Long, ByVal curRow As Long)
  
'    lRowRange = BlockRow2 - BlockRow + 1
'    lRow = BlockRow
    
End Sub

Private Sub ExcelPrn()

    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDateText       As String
    Dim sDate           As String
    Dim sDateNext       As String
    
    If ss1.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AGG2060C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
       
    sDate = Format(Date, "YYYY-MM-DD")
    sDateText = Left(sDate, 4) & "��"
    sDateText = sDateText & Mid(sDate, 6, 2) & "��"
    sDateText = sDateText & Mid(sDate, 9, 2) & "��"
    
    xlApp.Range("A2").Value = sDateText

    If lBlkrow1 = 0 Then
        lBlkrow1 = 1
    End If
    If lBlkrow2 = 0 Then
        lBlkrow2 = ss1.MaxRows
    End If
    
    For i = 2 To lBlkrow2 - lBlkrow1 + 1
          xlApp.Rows("4:4").Select
          xlApp.Selection.Copy
          xlApp.Selection.Insert Shift:=1
    Next i
    
    Clipboard.Clear
    
    ss1.Col = 1:        ss1.Col2 = 1 'ss1.MaxCols - 1
    ss1.Row = lBlkrow1: ss1.Row2 = lBlkrow2
    
    Clipboard.SetText ss1.Clip
        
    For i = 1 To lBlkrow2 - lBlkrow1 + 1
          xlApp.Range("A" & i + 3).Value = i
    Next i
    
    
    Clipboard.Clear
    ss1.SetSelection 1, 1, 1, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("B4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 2, 1, 2, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("C4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 3, 1, 3, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("D4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 5, 1, ss1.MaxCols - 1, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("E4").Select
    xlApp.ActiveSheet.Paste
    
'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
'     xlApp.Application.Visible = False
'     xlSheet.Close False
'     xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub
Private Sub ExcelPrn1()

    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDateText       As String
    Dim sDate           As String
    Dim sDateNext       As String
    
    If ss2.MaxRows < 1 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
     
    On Error Resume Next
    
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    
    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AGG2061C.xls")
    
    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select
    
    Clipboard.Clear
    ss2.SetSelection 1, 1, 1, ss2.MaxRows
    ss2.ClipboardCopy
    xlApp.Range("A4").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    Clipboard.Clear
    ss2.SetSelection 3, 1, ss2.MaxCols - 2, ss2.MaxRows
    ss2.ClipboardCopy
    xlApp.Range("B4").Select
    xlApp.ActiveSheet.Paste
    Clipboard.Clear
    
    ss2.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub Cmd_Edit_Click()
    'On Error GoTo Process_Exec_ERROR

    Dim OutParam(1, 4) As Variant
    Dim strRet_Result_ErrMsg As String
    'Dim strEdit_date As String
    Dim sQuery As String
          
    If Trim(CBO_PLT.Text) = "" Then
        Call Gp_MsgBoxDisplay(CBO_PLT.Tag + "��������")
        Exit Sub
    End If
    
    If Trim(CBO_PLT.Text) <> "C1" And Trim(CBO_PLT.Text) <> "C2" Then
   '     Call Gp_MsgBoxDisplay(txt_woo_rsn.Tag + " Must input according to length of item")
         Call Gp_MsgBoxDisplay(CBO_PLT.Tag + " �������")
        Exit Sub
    End If
    
    Dim adoCmd As ADODB.Command
    
     Screen.MousePointer = vbHourglass
    
    'Return Error Messsage Parameter
    OutParam(1, 1) = "arg_e_msg"
    OutParam(1, 2) = adVarChar
    OutParam(1, 3) = adParamOutput
    OutParam(1, 4) = 256
    
    'strEdit_date = CStr(TXT_CHK_TIME.RawData)
    
    sQuery = "{call AGG2060P ('" + Trim(CBO_PLT.Text) + "',?)}"

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
        strRet_Result_ErrMsg = adoCmd("arg_e_msg")
        sErrMessg = "Error Mesg : " & strRet_Result_ErrMsg
        Call Gp_MsgBoxDisplay(sErrMessg)
    Else
        Set adoCmd = Nothing
        Screen.MousePointer = vbDefault
        
        Call Gp_MsgBoxDisplay("���³ɹ�..!!", "I")
        Call Form_Ref
        Exit Sub
    End If
Process_Exec_ERROR:

    Set adoCmd = Nothing
    Screen.MousePointer = vbDefault
    Call Gp_MsgBoxDisplay("����ʧ�ܣ���")

End Sub

'Private Sub Cmd_Seq_Update_Click()
'
'    Dim AdoRs As ADODB.Recordset
'    Dim I             As Integer
'    Dim sSlabNo       As String
'    Dim SqlCmd        As String
'    Dim ProcessCnt    As Integer
'
'    On Error GoTo ErrHandle
'
'    If ss1.MaxRows < 1 Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'
'    M_CN1.BeginTrans
'
'    For I = 1 To ss1.MaxRows
'        ss1.Row = I
'        ss1.Col = 1
'        sSlabNo = Trim(ss1.Text)
'
'        SqlCmd = ""
'        SqlCmd = " UPDATE  GP_MILL_PLAN   " & vbCrLf
'        SqlCmd = SqlCmd & "   SET  CHG_SEQ_NO  = " & I & vbCrLf
'        SqlCmd = SqlCmd & "  Where SLAB_NO     = '" & sSlabNo & "' " & vbCrLf
'
'        M_CN1.Execute SqlCmd
'    Next I
'
'    M_CN1.CommitTrans
'    Call Gp_MsgBoxDisplay("���³ɹ�..!!", "I")
'    Screen.MousePointer = vbDefault
'
'    Exit Sub
'
'ErrHandle:
'
'    M_CN1.RollbackTrans
'    Screen.MousePointer = vbDefault
'    Call Gp_MsgBoxDisplay("����ʧ�ܣ���")
'
'End Sub

Private Sub vaSpread1_Advance(ByVal AdvanceNext As Boolean)

End Sub

'Private Sub SSTab1_Click(PreviousTab As Integer)
'
'    If SSTab1.Tab = 1 Then
'        AGG2060C.Cmd_Seq_Update.Visible = False
'        AGG2060C.Cmd_Edit.Visible = False
'    Else
'        AGG2060C.Cmd_Seq_Update.Visible = True
'        AGG2060C.Cmd_Edit.Visible = True
'    End If
'End Sub

Private Sub ss2_Advance(ByVal AdvanceNext As Boolean)

End Sub

Private Sub ss2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

    If Gf_Sc_Authority(sAuthority, "U") Then
        Call Gp_Sp_UpdateMake(ss2, Mode)
    End If

End Sub
Public Sub Spread_Del()

    Call Gp_Sp_Del(sc2)

End Sub


Private Sub ExcelFl()

    Dim i               As Integer
    Dim xlApp           As Object
    Dim xlSheet         As Object
    Dim sDateText       As String
    Dim sDate           As String
    Dim sDateNext       As String
        
        
    If ss1.MaxRows < 1 Then Exit Sub

    Screen.MousePointer = vbHourglass

    On Error Resume Next

    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
    End If

    Err.Clear

    xlApp.Workbooks.Open (App.Path & "\AGG2062C.xls")

    Set xlSheet = xlApp.Worksheets("Sheet1")
    xlApp.Sheets("Sheet1").Select

    sDate = Format(Date, "YYYY-MM-DD")
    sDateText = Left(sDate, 4) & "��"
    sDateText = sDateText & Mid(sDate, 6, 2) & "��"
    sDateText = sDateText & Mid(sDate, 9, 2) & "��"

    xlApp.Range("A2").Value = sDateText
    xlApp.Range("AE2").Value = sDateText

    If lBlkrow1 = 0 Then
        lBlkrow1 = 1
    End If
    If lBlkrow2 = 0 Then
        lBlkrow2 = ss1.MaxRows
    End If

    For i = 2 To lBlkrow2 - lBlkrow1 + 1
          xlApp.Rows("4:4").Select
          xlApp.Selection.Copy
          xlApp.Selection.Insert Shift:=1
    Next i

    Clipboard.Clear

    ss1.Col = 1:        ss1.Col2 = 1 'ss1.MaxCols - 1
    ss1.Row = lBlkrow1: ss1.Row2 = lBlkrow2

    Clipboard.SetText ss1.Clip

    For i = 1 To lBlkrow2 - lBlkrow1 + 1
          xlApp.Range("A" & i + 3).Value = i
          xlApp.Range("AE" & i + 3).Value = i
    Next i
    
    
    Clipboard.Clear
    ss1.SetSelection 1, 1, 1, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("B4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 2, 1, 2, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("C4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 3, 1, 3, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("D4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 5, 1, 5, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("E4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 7, 1, 7, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("F4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 8, 1, 8, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("G4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 9, 1, 9, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("H4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 11, 1, 11, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("I4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 12, 1, 12, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("J4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 14, 1, 14, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("K4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 15, 1, 15, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("L4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 16, 1, 16, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("M4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 17, 1, 17, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("N4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 18, 1, 18, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("O4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 19, 1, 19, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("P4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 20, 1, 20, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("Q4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 21, 1, 21, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("R4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 23, 1, 23, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("S4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 24, 1, 24, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("T4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 25, 1, 25, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("U4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 28, 1, 28, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("V4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 32, 1, 32, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("W4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 35, 1, 35, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("X4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 36, 1, 36, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("Y4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 37, 1, 37, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("Z4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 40, 1, 40, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AA4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 41, 1, 41, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AB4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 43, 1, 43, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AC4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 45, 1, 45, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AD4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 1, 1, 1, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AF4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 2, 1, 2, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AG4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 3, 1, 3, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AH4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 6, 1, 6, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AI4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 7, 1, 7, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AJ4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 8, 1, 8, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AK4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 9, 1, 9, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AL4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 12, 1, 12, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AM4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 14, 1, 14, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AN4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 16, 1, 16, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AO4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 17, 1, 17, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AP4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 18, 1, 18, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AQ4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 19, 1, 19, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AR4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 20, 1, 20, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AS4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 21, 1, 21, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AT4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 24, 1, 24, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AU4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 25, 1, 25, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AV4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 26, 1, 26, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AW4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 27, 1, 27, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AX4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 28, 1, 28, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AY4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 32, 1, 32, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("AZ4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 33, 1, 33, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("BA4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 34, 1, 34, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("BB4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 35, 1, 35, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("BC4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 36, 1, 36, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("BD4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 37, 1, 37, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("BE4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 40, 1, 40, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("BF4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 41, 1, 41, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("BG4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 43, 1, 43, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("BH4").Select
    xlApp.ActiveSheet.Paste
    
    Clipboard.Clear
    ss1.SetSelection 45, 1, 45, ss1.MaxRows
    ss1.ClipboardCopy
    xlApp.Range("BI4").Select
    xlApp.ActiveSheet.Paste
    
'    xlApp.ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    
    ss1.ClearSelection
       
    Screen.MousePointer = vbDefault
    
    xlApp.Application.Visible = True
'     xlApp.Application.Visible = False
'     xlSheet.Close False
'     xlApp.Quit
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
        
    Exit Sub

ErrHandle:
    MsgBox Error
'    xlApp.Application.Visible = True
    
    Set xlSheet = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbDefault
End Sub
