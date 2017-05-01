VERSION 5.00
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form EGD1020C 
   Caption         =   "热处理作业指示下达及调整_EGD1020C"
   ClientHeight    =   8235
   ClientLeft      =   90
   ClientTop       =   2310
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text_TOT_SHEETS 
      Alignment       =   1  'Right Justify
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   310
      Left            =   13980
      Locked          =   -1  'True
      TabIndex        =   17
      Tag             =   "机号"
      Top             =   540
      Width           =   1020
   End
   Begin VB.TextBox TXT_MAT_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13605
      MaxLength       =   14
      TabIndex        =   15
      Top             =   90
      Width           =   1620
   End
   Begin VB.TextBox txt_HTM_COND1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11250
      TabIndex        =   12
      Top             =   90
      Width           =   825
   End
   Begin VB.TextBox txt_HTM_METH1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10695
      TabIndex        =   11
      Text            =   " "
      Top             =   90
      Width           =   555
   End
   Begin VB.TextBox txt_SB 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7470
      TabIndex        =   10
      Text            =   " "
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox txt_PrcLine 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4035
      MaxLength       =   2
      TabIndex        =   9
      Tag             =   "工厂"
      Top             =   105
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.ComboBox cbo_PrcLine 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "EGD1020C.frx":0000
      Left            =   4275
      List            =   "EGD1020C.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "炉座号"
      Top             =   90
      Width           =   1635
   End
   Begin VB.TextBox TXT_PLT_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "工厂"
      Top             =   90
      Width           =   1020
   End
   Begin VB.TextBox txt_Plt 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1125
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   0
      Tag             =   "工厂"
      Top             =   90
      Width           =   540
   End
   Begin InDate.ULabel ULabel2 
      Height          =   315
      Left            =   90
      Top             =   90
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "工厂"
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
   Begin Threed.SSPanel SSPanel1 
      Height          =   420
      Left            =   90
      TabIndex        =   2
      Top             =   480
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   741
      _Version        =   196609
      BackColor       =   14737632
      BevelOuter      =   1
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSOption opt_move 
         Height          =   330
         Left            =   3930
         TabIndex        =   3
         Top             =   60
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "调 整"
      End
      Begin Threed.SSOption opt_delete 
         Height          =   330
         Left            =   2940
         TabIndex        =   4
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "删 除"
      End
      Begin Threed.SSOption opt_sent 
         Height          =   330
         Left            =   225
         TabIndex        =   5
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "发 送"
      End
      Begin Threed.SSOption opt_cancel 
         Height          =   300
         Left            =   1575
         TabIndex        =   6
         Top             =   60
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   529
         _Version        =   196609
         Font3D          =   2
         ForeColor       =   8421504
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "取 消"
      End
      Begin Threed.SSPanel SSPsend 
         Height          =   315
         Left            =   5130
         TabIndex        =   13
         Top             =   60
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   8454143
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "已下达"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel SSPpdt 
         Height          =   315
         Left            =   6330
         TabIndex        =   14
         Top             =   60
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   196609
         ForeColor       =   16711680
         BackColor       =   16761087
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "生产中"
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Left            =   3240
      Top             =   90
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "产线别"
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
   Begin FPSpread.vaSpread ss1 
      Height          =   8280
      Left            =   90
      TabIndex        =   8
      Top             =   930
      Width           =   15135
      _Version        =   393216
      _ExtentX        =   26696
      _ExtentY        =   14605
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
      MaxCols         =   25
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "EGD1020C.frx":0004
   End
   Begin InDate.ULabel ULabel22 
      Height          =   315
      Index           =   2
      Left            =   6435
      Top             =   90
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "抛丸"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel11 
      Height          =   315
      Left            =   8850
      Top             =   90
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      Caption         =   "热处理方法/条件"
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
      ForeColor       =   16711680
   End
   Begin InDate.ULabel ULabel6 
      Height          =   315
      Left            =   12555
      Top             =   90
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "物料号"
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
   Begin InDate.ULabel ULabel3 
      Height          =   315
      Left            =   11610
      Top             =   540
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      Caption         =   "中板热处理厂(等待/生产)"
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E1FE&
      Caption         =   "个"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15000
      TabIndex        =   16
      Top             =   600
      Width           =   195
   End
End
Attribute VB_Name = "EGD1020C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       Nisco Production Management System
'-- Sub_System Name   HTM System
'-- Program Name      指示调整
'-- Program ID        EGD1020C
'-- Document No       Q-00-0010(Specification)
'-- Designer          GUOLI
'-- Coder             GUOLI
'-- Date              2010.07.20
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

'Public Complete As Boolean           'Move Status Setting

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

Dim procCNT As Long

Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"

             Call Gp_Ms_Collection(txt_Plt, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
         Call Gp_Ms_Collection(txt_PrcLine, "p", "n", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
              Call Gp_Ms_Collection(txt_SB, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(txt_HTM_METH1, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
       Call Gp_Ms_Collection(txt_HTM_COND1, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
          Call Gp_Ms_Collection(TXT_MAT_NO, "p", " ", " ", " ", "r", " ", " ", pContro1, nContro1, mContro1, iContro1, rContro1, aContro1, lContro1)
    
    'MASTER Collection
    Mc1.Add Item:=pContro1, Key:="pControl"
    Mc1.Add Item:=nContro1, Key:="nControl"
    Mc1.Add Item:=mContro1, Key:="mControl"
    Mc1.Add Item:=iContro1, Key:="iControl"
    Mc1.Add Item:=rContro1, Key:="rControl"
    Mc1.Add Item:=cContro1, Key:="cControl"
    Mc1.Add Item:=aContro1, Key:="aControl"
    Mc1.Add Item:=lContro1, Key:="lControl"
    
    Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1, False)
    Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
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
 
    'Spread_Collection
    sc1.Add Item:=ss1, Key:="Spread"
    sc1.Add Item:="EGD1020C.P_MODIFY", Key:="P-M"
    sc1.Add Item:="EGD1020C.P_REFER1", Key:="P-R"
    sc1.Add Item:=pColumn1, Key:="pColumn"
    sc1.Add Item:=nColumn1, Key:="nColumn"
    sc1.Add Item:=aColumn1, Key:="aColumn"
    sc1.Add Item:=mColumn1, Key:="mColumn"
    sc1.Add Item:=iColumn1, Key:="iColumn"
    sc1.Add Item:=lColumn1, Key:="lColumn"
    sc1.Add Item:=1, Key:="First"
    sc1.Add Item:=ss1.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=sc1, Key:="Sc"
    
    Call Gp_Sp_ColHidden(ss1, 1, True)
    Call Gp_Sp_ColHidden(ss1, 2, True)
    Call Gp_Sp_ColHidden(ss1, 17, True)
    Call Gp_Sp_ColHidden(ss1, 23, True)
    Call Gp_Sp_ColHidden(ss1, 24, True)


    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
        
End Sub

Private Sub cbo_PrcLine_Click()
    If cbo_PrcLine.ListIndex = 0 Then
        txt_PrcLine = "1"
    ElseIf cbo_PrcLine.ListIndex = 1 Then
        txt_PrcLine = "2"
    End If
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    
        With MDIMain.MenuTool
            .Buttons(7).Enabled = False                 'Row Insert
            .Buttons(8).Enabled = False                 'Row Delete
            .Buttons(9).Enabled = False                 'Row Cancel
            .Buttons(11).Enabled = False                'Copy
            .Buttons(12).Enabled = False                'Paste
            .Buttons(14).Enabled = False                 'Excel
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
    
    Call Gp_Sp_ColGet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    
    txt_Plt.Text = "C3"
    TXT_PLT_NAME.Text = Gf_ComnNameFind(M_CN1, "C0001", txt_Plt.Text, 2)
    
    cbo_PrcLine.AddItem "一号线"
    cbo_PrcLine.AddItem "二号线"
    cbo_PrcLine.ListIndex = 1
    txt_PrcLine = "2"
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Gf_Sp_ProceExist(Proc_Sc("Sc")("Spread")) Then
        Cancel = 1
        Exit Sub
    End If
    
    Call Gp_Sp_ColSet(sc1.Item("Spread"), "EG-System.INI", Me.Name)
    
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
            opt_sent.Value = False
            opt_cancel.Value = False
            opt_move.Value = False
            opt_delete.Value = False
            opt_sent.ForeColor = &H808080
            opt_move.ForeColor = &H808080
            opt_delete.ForeColor = &H808080
            opt_cancel.ForeColor = &H808080
        End If
        
        With MDIMain.MenuTool
            .Buttons(7).Enabled = False                 'Row Insert
            .Buttons(8).Enabled = False                 'Row Delete
            .Buttons(9).Enabled = False                 'Row Cancel
            .Buttons(11).Enabled = False                'Copy
            .Buttons(12).Enabled = False                'Paste
            .Buttons(14).Enabled = False                 'Excel
        End With
        
    procCNT = 0
    SSPanel1.Enabled = True
    opt_sent.Value = False
    opt_cancel.Value = False
    opt_delete.Value = False
    
End Sub

Public Sub Form_Ref()

On Error GoTo Refer_Err

Dim sTemp As String
Dim sL2_Send As String
Dim sSlab_No As String
Dim sPrc_Sts As String
Dim iRow As Integer
Dim iCol As Integer
Dim i As Integer
Dim TIME As String
Dim sQuery As String
Dim sQuery1 As String

    If Gf_Sp_ProceExist(Proc_Sc("Sc").Item("Spread")) Then Exit Sub
    
      
    sQuery1 = "SELECT SUM(CASE WHEN PRC_STS = 'A' THEN 1 ELSE 0 END) || '/' ||SUM(CASE WHEN PRC_STS = 'B' THEN 1 ELSE 0 END)  FROM EP_HTM_INS  WHERE PRC_STS IN ('A','B') AND INS_LOC = 'H' AND PLT = 'C3'"
    Text_TOT_SHEETS.Text = Gf_FloatFind(M_CN1, sQuery1)
    
    If Gf_Sp_Refer(M_CN1, sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
       Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
        With MDIMain.MenuTool
            .Buttons(7).Enabled = False                 'Row Insert
            .Buttons(8).Enabled = False                 'Row Delete
            .Buttons(9).Enabled = False                 'Row Cancel
            .Buttons(11).Enabled = False                'Copy
            .Buttons(12).Enabled = False                'Paste
            .Buttons(14).Enabled = True                 'Excel
        End With
    End If
    
    TIME = Format(Now, "YYYY-MM")
    procCNT = 0
    SSPanel1.Enabled = True
    opt_sent.Value = False
    opt_cancel.Value = False
    opt_delete.Value = False
    
    For iRow = 1 To ss1.MaxRows
    
      ss1.ROW = iRow
      ss1.Col = 23
        If Mid(ss1.Text, 1, 7) < TIME Then
          For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.ForeColor = &HFF&
          Next
        End If

        If ss1.Text = "" Then
           Exit For
        End If
      
        ss1.ROW = iRow
        ss1.Col = 21
        If ss1.Text = "B" Then
           For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = SSPpdt.BackColor
           Next
        Else
           If ss1.Text = "Y" Then
           
           For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = SSPsend.BackColor
           Next
              
           End If
        End If
        
        ss1.ROW = iRow
        ss1.Col = 22
        If ss1.Text = "B" Then
           For i = 1 To ss1.MaxCols
               ss1.Col = i
               ss1.BackColor = SSPpdt.BackColor
           Next
        End If
        
        If ss1.Text = "" Then
           Exit For
        End If
     
        
    Next iRow
    
    ss1.OperationMode = OperationModeNormal
    
Exit Sub

Refer_Err:

End Sub

Public Sub Form_Pro()
Dim mResult As String
Dim sMsg As String
If opt_delete.Value = True Then
   If Gf_MessConfirm("确定要删除标记为" & "'Update'" & "的作业指示吗？", "W", "系统提示信息确认") Then
      Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1)
   Else
      Exit Sub
   End If
End If
    
Call Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1)
Call Form_Ref

End Sub

Public Sub Form_Ins()
    
'    Call Gp_Sp_Ins(Proc_Sc("Sc"))

End Sub

Public Sub Spread_Cpy()

'    Call Gp_Sp_Copy(Proc_Sc("Sc"))
    
End Sub

Public Sub Spread_Pst()

'    Call Gp_Sp_Paste(Proc_Sc("Sc"))
    
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
    
'    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub opt_cancel_Click(Value As Integer)
    Dim iRow As Integer
    Dim sTemp As String
    Dim ForCnt As Long
    
    If ss1.MaxRows <= 0 Then opt_cancel.Value = False: Exit Sub
    
    If opt_cancel.Value = True Then
        opt_cancel.ForeColor = &HFF&
        opt_sent.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
    Else
        opt_cancel.ForeColor = &H808080
    End If
    
    For ForCnt = 1 To ss1.MaxRows
        ss1.ROW = ForCnt
        ss1.Col = 1
        ss1.Text = "C"
        If ss1.Text = "C" Then
           ss1.Col = 4
           ss1.Lock = True
        End If
    Next
    
End Sub

Private Sub opt_delete_Click(Value As Integer)
    Dim iRow As Integer
    Dim sTemp As String
    Dim ForCnt As Long
    
    If ss1.MaxRows <= 0 Then opt_delete.Value = False: Exit Sub
    
    If opt_delete.Value = True Then
    
        opt_delete.ForeColor = &HFF&
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
    Else
        opt_delete.ForeColor = &H808080
    End If
    
    For ForCnt = 1 To ss1.MaxRows
        ss1.ROW = ForCnt
        ss1.Col = 1
        ss1.Text = "D"
        If ss1.Text = "D" Then
           ss1.Col = 4
           ss1.Lock = True
        End If
    Next

End Sub


Private Sub opt_move_Click(Value As Integer)
    Dim iRow As Integer
    Dim sTemp As String
    
    If opt_move.Value = True Then
        opt_move.ForeColor = &HFF&
        opt_sent.ForeColor = &H808080
        opt_cancel.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
    Else
        opt_move.ForeColor = &H808080
    End If

    
End Sub

Private Sub opt_sent_Click(Value As Integer)
    Dim iRow As Integer
    Dim sTemp As String
    Dim ForCnt As Long
    
    If ss1.MaxRows <= 0 Then opt_sent.Value = False: Exit Sub
    
    If opt_sent.Value = True Then
        opt_sent.ForeColor = &HFF&
        opt_cancel.ForeColor = &H808080
        opt_move.ForeColor = &H808080
        opt_delete.ForeColor = &H808080
    Else
        opt_sent.ForeColor = &H808080
    End If
    
    For ForCnt = 1 To ss1.MaxRows
        ss1.ROW = ForCnt
        ss1.Col = 1
        ss1.Text = "S"
        If ss1.Text = "S" Then
           ss1.Col = 4
           ss1.Lock = False
        End If
    Next

    
End Sub


Private Sub ss1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Dim iRow As Integer
    Dim i As Integer
    
    lBlkcol1 = BlockCol
    lBlkcol2 = BlockCol2
    lBlkrow1 = BlockRow
    lBlkrow2 = BlockRow2
    
    If BlockRow < 0 Then Exit Sub
    
    If Not opt_sent And Not opt_cancel And Not opt_delete Then
        MsgBox "请先确认作业功能......!", vbCritical, "系统提示信息"
        Exit Sub
    End If
    
    If Gf_Sc_Authority(sAuthority, "U") Then
    
        For iRow = BlockRow To BlockRow2
        
            ss1.ROW = iRow
            ss1.Col = 0
            If ss1.Text = "Update" Then
                ss1.Text = ""
                ss1.Col = 21
                If ss1.Text = "B" Then
                    For i = 1 To ss1.MaxCols
                        ss1.Col = i
                        ss1.BackColor = SSPpdt.BackColor
                    Next
                ElseIf ss1.Text = "Y" Then
                    For i = 1 To ss1.MaxCols
                        ss1.Col = i
                        ss1.BackColor = SSPsend.BackColor
                    Next
                Else
                    Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow)
                End If
                
                ss1.Col = 22
                If ss1.Text = "B" Then
                   For i = 1 To ss1.MaxCols
                       ss1.Col = i
                       ss1.BackColor = SSPpdt.BackColor
                   Next
                End If
                
                procCNT = procCNT - 1
            Else
                ss1.Text = "Update"
                Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFFFF80)
                procCNT = procCNT + 1
            End If
            
        Next iRow
        
    End If
    
    If procCNT > 0 Then
        SSPanel1.Enabled = False
    Else
        SSPanel1.Enabled = True
    End If
       
End Sub

Private Sub ss1_Click(ByVal Col As Long, ByVal ROW As Long)
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, ROW)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
End Sub

Private Sub txt_HTM_COND1_Change()
    If Len(Trim(txt_HTM_COND1.Text)) = 4 Then
       If Trim(txt_HTM_METH1) <> Mid(Trim(txt_HTM_COND1), 1, 1) Then
          MsgBox "热处理方法与热处理条件不一样", vbCritical, "系统提示信息"
          txt_HTM_COND1 = ""
       End If
    End If
End Sub

Private Sub txt_HTM_COND1_DblClick()
    Call txt_HTM_COND1_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_HTM_COND1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = ""
        DD.rControl.Add Item:=txt_HTM_COND1

        DD.nameType = "2"

        Call Gf_HEAT_COND_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub txt_HTM_METH1_DblClick()
    Call txt_HTM_METH1_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_HTM_METH1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0073"
        DD.rControl.Add Item:=txt_HTM_METH1

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub

Private Sub txt_SB_DblClick()
    Call txt_SB_KeyUp(vbKeyF4, 0)
End Sub

Private Sub txt_SB_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then

        DD.sWitch = "MS"
        DD.sKey = "Q0074"
        DD.rControl.Add Item:=txt_SB

        DD.nameType = "2"

        Call Gf_Common_DD(M_CN1, KeyCode)

        Exit Sub

    End If
End Sub
