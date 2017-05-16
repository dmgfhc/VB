VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{D1F54538-FC6B-4AC6-9655-2FB5170110A8}#1.0#0"; "Indate.ocx"
Object = "{8C3D4AA0-2599-11D2-BAF1-00104B9E0792}#3.0#0"; "sssplt30.ocx"
Begin VB.Form AQC0050C 
   Caption         =   "复样指示及下达界面_AQC0050C"
   ClientHeight    =   10890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10890
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Row_Check 
      BackColor       =   &H00E0E0E0&
      Caption         =   "双击试样号多行标记不下达"
      Height          =   300
      Left            =   11760
      TabIndex        =   22
      Top             =   45
      Width           =   2835
   End
   Begin VB.CheckBox cbo_loc 
      Caption         =   "B"
      Height          =   375
      Index           =   1
      Left            =   13920
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox cbo_loc 
      Caption         =   "T"
      Height          =   375
      Index           =   0
      Left            =   13440
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txt_OUT_SHEET_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1020
      MaxLength       =   14
      TabIndex        =   19
      Top             =   840
      Width           =   2145
   End
   Begin VB.CommandButton cmd_AllCheck 
      Caption         =   "标记不下达"
      Height          =   345
      Left            =   10200
      TabIndex        =   18
      Top             =   45
      Width           =   1275
   End
   Begin VB.TextBox SAVE_SMP_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13440
      MaxLength       =   14
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.CheckBox txt_CHECK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "已下达"
      Height          =   300
      Left            =   9270
      TabIndex        =   16
      Top             =   450
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   315
      Left            =   9375
      TabIndex        =   14
      Top             =   825
      Width           =   2385
      Begin VB.OptionButton opt_KND_Y 
         BackColor       =   &H00E0E0E0&
         Caption         =   "热处理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1215
         TabIndex        =   9
         Top             =   30
         Width           =   1125
      End
      Begin VB.OptionButton opt_KND_N 
         BackColor       =   &H00E0E0E0&
         Caption         =   "非热处理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.TextBox txt_PRC 
      Height          =   315
      Left            =   10200
      MaxLength       =   2
      TabIndex        =   13
      Top             =   405
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.ComboBox CBO_LINE 
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
      ItemData        =   "AQC0050C.frx":0000
      Left            =   12840
      List            =   "AQC0050C.frx":0002
      TabIndex        =   10
      Top             =   825
      Width           =   525
   End
   Begin VB.TextBox txt_SMP_SEND_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6285
      MaxLength       =   13
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txt_SMP_NO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1020
      MaxLength       =   14
      TabIndex        =   4
      Top             =   450
      Width           =   2145
   End
   Begin VB.TextBox txt_STDSPEC_NAME 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   8550
      MaxLength       =   18
      TabIndex        =   3
      Top             =   30
      Width           =   645
   End
   Begin VB.TextBox txt_STDSPEC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   6255
      MaxLength       =   18
      TabIndex        =   2
      Top             =   30
      Width           =   2295
   End
   Begin VB.TextBox TXT_PLT 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1005
      MaxLength       =   2
      TabIndex        =   0
      Top             =   30
      Width           =   420
   End
   Begin VB.TextBox txt_PLT_NAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1455
      TabIndex        =   1
      Top             =   30
      Width           =   1710
   End
   Begin InDate.UDate dtp_date_t 
      Height          =   315
      Left            =   7785
      TabIndex        =   6
      Top             =   435
      Width           =   1440
      _ExtentX        =   2540
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
   Begin InDate.UDate dtp_date_f 
      Height          =   315
      Left            =   6255
      TabIndex        =   5
      Top             =   450
      Width           =   1500
      _ExtentX        =   2646
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
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   2
      Left            =   4950
      Top             =   435
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "生产日期"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   315
      Index           =   3
      Left            =   4965
      Top             =   30
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   556
      Caption         =   "标准编号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin SSSplitter.SSSplitter SSSplitter1 
      Height          =   7905
      Left            =   60
      TabIndex        =   11
      Top             =   1215
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   13944
      _Version        =   196609
      BorderStyle     =   0
      PaneTree        =   "AQC0050C.frx":0004
      Begin FPSpread.vaSpread ss3 
         Height          =   7905
         Left            =   4830
         TabIndex        =   12
         Top             =   0
         Width           =   10305
         _Version        =   393216
         _ExtentX        =   18177
         _ExtentY        =   13944
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
         SpreadDesigner  =   "AQC0050C.frx":0056
      End
      Begin FPSpread.vaSpread ss1 
         Height          =   7905
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   4740
         _Version        =   393216
         _ExtentX        =   8361
         _ExtentY        =   13944
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
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
         MaxCols         =   49
         MaxRows         =   1
         ProcessTab      =   -1  'True
         Protect         =   0   'False
         SpreadDesigner  =   "AQC0050C.frx":03C7
      End
   End
   Begin InDate.ULabel ULabel4 
      Height          =   315
      Left            =   4950
      Top             =   840
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Caption         =   "试样委托单号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   0
      Left            =   60
      Top             =   450
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   529
      Caption         =   "试样编号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   6
      Left            =   60
      Top             =   30
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   529
      Caption         =   "生产厂"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      Left            =   8070
      Top             =   825
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   556
      Caption         =   "是否热处理"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel5 
      Height          =   315
      Left            =   11790
      Top             =   825
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      Caption         =   "热处理线"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin InDate.ULabel ULabel1 
      Height          =   300
      Index           =   1
      Left            =   60
      Top             =   840
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   529
      Caption         =   "轧制批号"
      Alignment       =   1
      BackColor       =   14804173
      BackgroundStyle =   1
      ChiselText      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "AQC0050C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------
'-- PROGRAM HEADER  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'-------------------------------------------------------------------------------
'-- System Name       质量管理
'-- Sub_System Name   判定管理
'-- Program Name      复样指示及下达界面
'-- Program ID        AQC0050C
'-- Document No       Q-00-0010(Specification)
'-- Designer          HAN.Y.S
'-- Coder             Sun Bin
'-- Date              2008.03. 11
'-- Description       复样指示及下达界面
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

Dim Mc1 As New Collection           'Master Collection
Dim Sc1 As New Collection           'Spread Collection
'Dim sc2 As New Collection
Dim sc3 As New Collection
Dim xy()


Dim Proc_Sc As New Collection       'Spread Struc Collection

Dim lBlkcol1 As Long                'To Excel Block Col1
Dim lBlkcol2 As Long                'To Excel Block Col2
Dim lBlkrow1 As Long                'To Excel Block Row1
Dim lBlkrow2 As Long                'To Excel Block Row2

Dim arrChem(3, 35) As String
Private Sub Form_Define()
        
    'Form Type : Start , Master, Sheet, Msheet, PopSheet, Refer
    FormType = "Msheet"
    'Call Master_Collection("Control_Name", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "r(refer)", "a(append)", "l(lock)")
            Call Gp_Ms_Collection(txt_plt, "p", "n", " ", " ", " ", "a", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
        Call Gp_Ms_Collection(txt_STDSPEC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(txt_SMP_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(dtp_date_f, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
         Call Gp_Ms_Collection(dtp_date_t, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
    Call Gp_Ms_Collection(txt_SMP_SEND_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
           Call Gp_Ms_Collection(CBO_LINE, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
            Call Gp_Ms_Collection(txt_PRC, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
          Call Gp_Ms_Collection(txt_CHECK, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
   Call Gp_Ms_Collection(txt_OUT_SHEET_NO, "p", " ", " ", " ", " ", " ", " ", pControl, nControl, mControl, iControl, rControl, aControl, lControl)
     
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
'     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 1, " ", " ", " ", "i", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 2, " ", " ", " ", "i", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 3, " ", " ", " ", "i", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 4, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '块数
     Call Gp_Sp_Collection(ss1, 5, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)    '重量
     Call Gp_Sp_Collection(ss1, 6, " ", " ", " ", "i", " ", "", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 7, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 8, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
     Call Gp_Sp_Collection(ss1, 9, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 10, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 11, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 12, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 13, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 14, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 15, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 16, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 17, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 18, " ", " ", " ", " ", " ", "l", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 19, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 20, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 21, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 22, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 23, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 24, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 25, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 26, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 27, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 28, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 29, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 30, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 31, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 32, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 33, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 34, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) 'Z向加试   34
    Call Gp_Sp_Collection(ss1, 35, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1) '追加z向  35
    Call Gp_Sp_Collection(ss1, 36, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 37, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 38, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 39, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 40, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 41, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 42, " ", " ", " ", " ", "a", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 43, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 44, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 45, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 46, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 47, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 48, " ", " ", " ", " ", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    Call Gp_Sp_Collection(ss1, 49, " ", " ", " ", "i", " ", " ", pColumn1, nColumn1, mColumn1, iColumn1, aColumn1, lColumn1)
    
    'Spread_Collection
    Sc1.Add Item:=ss1, Key:="Spread"
    Sc1.Add Item:="AQC0050C.P_REFER", Key:="P-R"
    Sc1.Add Item:="AQC0050C.P_ONEROW", Key:="P-O"
    Sc1.Add Item:="AQC0050C.P_MODIFY", Key:="P-M"
    Sc1.Add Item:=pColumn1, Key:="pColumn"
    Sc1.Add Item:=nColumn1, Key:="nColumn"
    Sc1.Add Item:=aColumn1, Key:="aColumn"
    Sc1.Add Item:=mColumn1, Key:="mColumn"
    Sc1.Add Item:=iColumn1, Key:="iColumn"
    Sc1.Add Item:=lColumn1, Key:="lColumn"
    Sc1.Add Item:=1, Key:="First"
    Sc1.Add Item:=ss1.MaxCols, Key:="Last"
    
'     Call SS1.AddCellSpan(5, 0, 1, 2)


    
      'Call Spread_Collection("Column_Num", "p(primary)", "n(Necessary)", "m(maxlength)", "i(insert)", "a(append_down), "l(lock)")
     Call Gp_Sp_Collection(ss3, 1, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 2, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 3, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 4, " ", " ", " ", " ", " ", " ", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     Call Gp_Sp_Collection(ss3, 5, " ", " ", " ", " ", " ", "l", pColumn13, nColumn13, mColumn13, iColumn13, aColumn13, lColumn13)
     
     'Spread_Collection
    sc3.Add Item:=ss3, Key:="Spread"
    sc3.Add Item:="AQC0050C.P_SREFER", Key:="P-R"
    sc3.Add Item:=pColumn13, Key:="pColumn"
    sc3.Add Item:=nColumn13, Key:="nColumn"
    sc3.Add Item:=aColumn13, Key:="aColumn"
    sc3.Add Item:=mColumn13, Key:="mColumn"
    sc3.Add Item:=iColumn13, Key:="iColumn"
    sc3.Add Item:=lColumn13, Key:="lColumn"
    sc3.Add Item:=1, Key:="First"
    sc3.Add Item:=ss3.MaxCols, Key:="Last"

    Proc_Sc.Add Item:=Sc1, Key:="Sc"
    
    'Call Gp_Sp_ColHidden(ss1, 37, True)
    'Call Gp_Sp_ColHidden(ss1, 43, True)
    'Call Gp_Sp_ColHidden(Ss3, 0, True)
    'Call Gp_Sp_ColHidden(Ss3, 5, True)
    
    Me.KeyPreview = True
    Me.BackColor = &HE0E0E0
'    Call Gp_Sp_BlockColor(ss1, 2, ss1.MaxCols, 1, ss1.MaxRows, , &HFFFF&)


End Sub

Private Sub MenuToolSet()

    MDIMain.MenuTool.Buttons(7).Enabled = False    'Row Insert
    MDIMain.MenuTool.Buttons(8).Enabled = False    'Row Delete
    MDIMain.MenuTool.Buttons(9).Enabled = False    'Row Cancle
    MDIMain.MenuTool.Buttons(11).Enabled = False   'Copy
    MDIMain.MenuTool.Buttons(12).Enabled = False   'Paste
'    MDIMain.MenuTool.Buttons(14).Enabled = False

End Sub

Private Sub cbo_loc_Click(Index As Integer)


 'louyannan 20101215
If cbo_loc(Index).Value = "1" Then
  cbo_loc(Abs(Index - 1)) = "0"
Else
  cbo_loc(Abs(Index - 1)) = "1"
End If


ss1_Click ss1.ActiveCol, ss1.ActiveRow


End Sub






Private Sub cmd_AllCheck_Click()

  If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If
 
    Call DataSave
    
End Sub

Private Sub Form_Activate()
     
    Call MDIMain.FormMenuSetting(Me, FormType, Toolbar_St, sAuthority)
    Call MenuToolSet

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = KEY_RETURN Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo Err_Track:
    Dim oCodeName As Object
    Dim sCode As String
    
    Select Case Me.ActiveControl.Name
        Case "TXT_PLT"                     '工厂
            sCode = "C0001"
            Set oCodeName = txt_PLT_NAME
        Case "txt_STDSPEC"                 '标准
            sCode = "STDSPEC"
            Set oCodeName = txt_STDSPEC_NAME
    End Select
    
    If sCode = "" Then Exit Sub
    
    Call Gp_MS_CodeNameFind(KeyCode, sCode, Me.ActiveControl, oCodeName)
    
    Set oCodeName = Nothing

Err_Track:
    
    Set oCodeName = Nothing

End Sub

Private Sub Form_Load()

    Screen.MousePointer = vbHourglass

    sAuthority = Gf_Pgm_Authority(Me.Name)
'    sAuthority = "1111"
    sPLT_Authority = Gf_PLT_Authority(Me.Name)
    If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
       txt_plt.Text = sPLT_Authority
    Else
       txt_plt.Text = ""
    End If
    
    txt_PRC.Text = "DS"
    
    Call Form_Define

    Call MDIMain.FormMenuSetting(Me, FormType, "FS", sAuthority)
    Call MenuToolSet

    Call Gp_Ms_Cls(Mc1("rControl"))

    Call Gp_Ms_NeceColor(Mc1("nControl"))
    
    Call Gp_Sp_Setting(ss1)
    Call Gp_Sp_Setting(ss3)

    Call Gf_Sp_Cls(Proc_Sc("Sc"))

    Call Gp_Sp_ColGet(ss1, "Q-System.INI", Me.Name)
    Call Gp_Sp_ColGet(ss3, "Q-System.INI", Me.Name)
    
    CBO_LINE.AddItem ""
    CBO_LINE.AddItem "1"
    CBO_LINE.AddItem "2"
    
    CBO_LINE.ListIndex = 0
    
    cbo_loc(0).Value = "1" 'louyannan 20110110
    Screen.MousePointer = vbDefault

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
    
    Set iColumn12 = Nothing
    Set pColumn12 = Nothing
    Set lColumn12 = Nothing
    Set nColumn12 = Nothing
    Set mColumn12 = Nothing
    Set aColumn12 = Nothing
    
    Set iColumn13 = Nothing
    Set pColumn13 = Nothing
    Set lColumn13 = Nothing
    Set nColumn13 = Nothing
    Set mColumn13 = Nothing
    Set aColumn13 = Nothing

    Set Mc1 = Nothing
    Set Sc1 = Nothing
    Set sc3 = Nothing
    Set Proc_Sc = Nothing
    
    Call MDIMain.FormMenuSetting(Me, "Start", Toolbar_St, "")
    
End Sub


Public Sub Spread_Can()

    Call Gp_Sp_Cancel(M_CN1, Proc_Sc("SC"))

End Sub

Public Sub Form_Cls()
    
    If Gf_Sp_Cls(Proc_Sc("SC")) Then
        Call Gp_Ms_Cls(Mc1("rControl"))
        Call Gf_Sp_Cls(sc3)
        Call MDIMain.FormMenuSetting(Me, FormType, "CLS", sAuthority)
        
        If sPLT_Authority <> "**" And sPLT_Authority <> "" Then
           txt_plt.Text = sPLT_Authority
        Else
           txt_plt.Text = ""
        End If
        
    End If
    
    txt_PLT_NAME = ""
    txt_STDSPEC_NAME.Text = ""
    SAVE_SMP_NO = ""

End Sub

Public Sub Form_Ref()
    Dim iRow, iCol  As Integer
    Dim sQuery      As String
    Dim sMesg       As String
    Dim AdoRs       As adodb.Recordset

    On Error GoTo Refer_Err
    
    If dtp_date_f.RawData = "" Then
       'dtp_date_f.RawData = Format(Now, "yyyymm") + "01"
       dtp_date_f.RawData = ""
    End If
    
    If dtp_date_t.RawData = "" Then
       dtp_date_t.RawData = Format(Now, "yyyymmdd")
    End If
    
    If Gf_Sp_Refer(M_CN1, Sc1, Mc1, Mc1("nControl"), Mc1("mControl")) Then
        ss1.OperationMode = OperationModeNormal
        Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
        Call MenuToolSet
    End If
'
    Call Gf_Sp_Cls(sc3)
    ReDim xy(1, 0)
    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub
    
    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 45  '43
            If .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFF80FF)
            End If
         Next iRow
    End With
    
    Row_Check = 0
    
Refer_Err:
    
    SAVE_SMP_NO.Text = ""
    Screen.MousePointer = vbDefault

End Sub

Private Sub ss1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

    If Col <> 10 And Col <> 11 Then Exit Sub

    ss1.Row = Row

    If Col = 19 And ButtonDown = 1 Then
        ss1.Col = 20
        ss1.Text = 0
    ElseIf Col = 20 And ButtonDown = 1 Then
        ss1.Col = 19
        ss1.Text = 0
    End If
    
'    If Col <> 8 And Col <> 9 Then Exit Sub
'
'    ss1.Row = Row
'
'    If Col = 17 And ButtonDown = 1 Then
'        ss1.Col = 18
'        ss1.Text = 0
'    ElseIf Col = 18 And ButtonDown = 1 Then
'        ss1.Col = 17
'        ss1.Text = 0
'    End If

End Sub

''
Private Sub ss1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)


    If Gf_Sc_Authority(sAuthority, "U") Then
'    设定第一列update
        Call Gp_Sp_UpdateMake(Proc_Sc("Sc")("Spread"), Mode)
'      设定下达人员
        Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 40)

    End If
    
'    With ss1
'
'         .Row = .ActiveRow
'         .Col = .ActiveCol
'         If .Col = 7 And ButtonDown = 1 Then
'            If .Text = 1 Then
'               .Col = 8
'               .Text = 0
'            End If
'         ElseIf .Col = 8 And ButtonDown = 1 Then
'            If .Text = 1 Then
'               .Col = 7
'               .Text = 0
'            End If
'         End If
'    End With
    
    Call MenuToolSet

End Sub

Public Sub Form_Pro()
    Dim iRow, iCol, i, j, k As Integer
    Dim otherItem  As String
    
    If sPLT_Authority <> "**" And sPLT_Authority <> txt_plt.Text Then
       Call Gp_MsgBoxDisplay("这个工厂的产品 你没有修改功能", "I")
       Exit Sub
    End If
    
 '--------------------配置化项目复样  王成  2012.12.12-----------把值填到ss1其它项目中----------------------
    If Gf_Sc_Authority(sAuthority, "U") Then
    For i = 0 To UBound(xy, 2)
    If xy(0, i) <> "" And xy(1, i) <> "" Then
      With ss3
          .Row = xy(1, i)
          .Col = 1
           With ss1
           For iRow = 1 To ss1.MaxRows
           .Row = iRow
           .Col = 1
           If .Text = xy(0, i) Then
                .Col = 33
                .Text = .Text + ss3.Text + ";"
                 Call Gp_Sp_InAuthority(Proc_Sc("Sc"), 35)
           End If
           Next iRow
         End With
       End With
    End If
    Next i
    End If
    '---------------------------------------------------------------------------------------------------
    
    
    If Gf_Sp_Process(M_CN1, Proc_Sc("Sc"), Mc1) Then
      Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    End If
    
    Call MenuToolSet
    
    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub

    With ss1
    
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 43
            If .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFF80FF)
            End If
         Next iRow
         
    End With
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

Public Sub Spread_Del()
    
    Call Gp_Sp_Del(Proc_Sc("SC"))

End Sub

Private Sub opt_KND_N_Click()
        opt_KND_N.Value = True
        opt_KND_Y.Value = False
        txt_PRC.Text = "DS"
        CBO_LINE.ListIndex = 0
End Sub

Private Sub opt_KND_Y_Click()
        opt_KND_N.Value = False
        opt_KND_Y.Value = True
        txt_PRC.Text = "DH"
End Sub


Private Sub ss1_Click(ByVal Col As Long, ByVal Row As Long)

    Dim sQuery          As String
    Dim sMesg           As String
    Dim i               As Integer
    Dim j               As Integer
    Dim AdoRs           As adodb.Recordset
    Dim ArrayRecords    As Variant
    Dim arr             As Variant
    Dim SMP_NO, smp_loc As Variant
    
 On Error GoTo Error_Rtn

    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)

    If ss1.MaxRows < 1 Or Row = 0 Then Exit Sub

    With ss1
        .Col = 1
        .Row = .ActiveRow
        SMP_NO = .Text
        .Col = 44
        smp_loc = .Text
    End With
    

    'If SMP_NO = SAVE_SMP_NO Then Exit Sub
    
    ss3.MaxRows = 0

    ss1.ReDraw = False
    ss3.ReDraw = False

    Set AdoRs = New adodb.Recordset
    
    
    If smp_loc = "Y" Then
     cbo_loc(0).Visible = True
     cbo_loc(1).Visible = True

     If cbo_loc(0).Value = "1" Then 'louyannan  20101215
     smp_loc = "T"
     Else
     smp_loc = "B"
     End If
   Else
    cbo_loc(0).Visible = False
    cbo_loc(1).Visible = False

   End If

   sQuery = "{call AQC0050C.P_SREFER_1('" + Trim(SMP_NO) + "','" + Trim(smp_loc) + "')}"
                    
   
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        ArrayRecords = AdoRs.GetRows
        Call subSpreadView1(ArrayRecords)
        Erase ArrayRecords
    End If
     
    sQuery = "{call AQC0050C.P_SREFER('" + Trim(SMP_NO) + "')}"
    
    AdoRs.Close
                    
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
        
    If Not (AdoRs.BOF And AdoRs.EOF) Then
        ArrayRecords = AdoRs.GetRows
        Call subSpreadView3(ArrayRecords)
        Erase ArrayRecords
    End If
'
    Call Gp_Sp_EvenRowBackcolor(ss3)
    
    
    '--------------------配置化项目显示  刘翔  2012.11.20-----------------------------------------------------
    
    Erase ArrayRecords
    
    AdoRs.Close
    sQuery = "{call AQC0040C.P_SREFER_CONFIG('" + Trim(SMP_NO) + "','" + Trim(smp_loc) + "')}"
    AdoRs.Open sQuery, M_CN1, adOpenKeyset
    
    Set AdoRs = M_CN1.Execute(sQuery)
    
    If Not AdoRs.EOF And Not AdoRs.BOF Then
      ArrayRecords = AdoRs.GetRows
      Call subSpreadView_Config(ArrayRecords)
    End If
    
    AdoRs.Close
    Erase ArrayRecords
    
    '-----------------------------------------------------------------------------------------------------------
     '--------------------配置化项目复样  王成  2012.12.12-------------把选择的配置化项目填到ss3中--------------
    If ss1.ActiveRow > 0 Then
       With ss1
       .Col = 1
       .Row = ss1.ActiveRow
          For i = 0 To UBound(xy, 2)
             If ss1.Text = xy(0, i) Then
                With ss3
                ss3.Row = xy(1, i)
                ss3.Col = 6
                ss3.Text = "1"
                End With
             End If
          Next i
      End With
    End If
    '-----------------------------------------------------------------------------------------------------------

    Set AdoRs = Nothing
    Set ArrayRecords = Nothing
    ss1.ReDraw = True
    ss3.ReDraw = True
    
    SAVE_SMP_NO = SMP_NO

    Exit Sub

Error_Rtn:

    Set AdoRs = Nothing
    Set ArrayRecords = Nothing
    Screen.MousePointer = vbDefault
    ss1.ReDraw = True
    ss3.ReDraw = True

End Sub


Private Sub ss1_DblClick(ByVal Col As Long, ByVal Row As Long)

If ss1.ActiveCol = 1 Or ss1.ActiveCol = 2 Or ss1.ActiveCol = 3 Or ss1.ActiveCol = 6 Then 'louyannan 20110112
 txt_SMP_NO.SetFocus
 End If
    Call Gp_Sp_Sort(Proc_Sc("Sc")("Spread"), Col, Row)
    
    lBlkcol1 = 0
    lBlkcol2 = 0
    lBlkrow1 = 0
    lBlkrow2 = 0
    
  If Row_Check = 1 And txt_CHECK = 0 And ss1.ActiveCol = 1 Then
    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = "Update"
           .Col = 39
           .Text = sUserID
    End With
  End If
    
End Sub


Private Sub ss1_KeyDown(KeyCode As Integer, Shift As Integer)

    If Proc_Sc("Sc")("Spread").MaxRows < 1 Then Exit Sub
    
    
    If ss1.ActiveCol = 1 Or ss1.ActiveCol = 2 Or ss1.ActiveCol = 3 Or ss1.ActiveCol = 6 Then 'louyannan 20110112
     txt_SMP_NO.SetFocus
     End If

    
    If Gf_Sc_Authority(sAuthority, "I") = False Then Exit Sub
    
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyTab And Shift <> 1) Then
        Call Gp_Sp_AutoInsert(Proc_Sc("Sc"))
    End If

    If Shift = 0 Then Proc_Sc("Sc")("Spread").EditMode = True
    

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
Private Sub subSpreadView1(ByVal strArr As Variant)

    Dim i           As Integer
    Dim iRow        As Integer
    Dim sMatr(188)   As String
    
    If UBound(strArr, 2) < 0 Then Exit Sub
        
        
    sMatr(0) = "屈服点实绩                            "
    sMatr(1) = "拉伸规定总伸长应力实绩                "
    sMatr(2) = "抗拉强度实绩                          "
    sMatr(3) = "屈强比实绩                            "
    sMatr(4) = "断后伸长率实绩                        "
    sMatr(5) = "断面收缩率实绩1                       "
    sMatr(6) = "断面收缩率实绩2                       "
    sMatr(7) = "断面收缩率实绩3                       "
       
    sMatr(8) = "断面收缩率实绩平均                    "
    sMatr(9) = "冷弯试验实绩                          "
    sMatr(10) = "冲击试验温度                         "
    sMatr(11) = "冲击试样尺寸                         "
    sMatr(12) = "冲击试验实绩 1                       "
    sMatr(13) = "冲击试验实绩 2                       "
    sMatr(14) = "冲击试验实绩 3                       "
    sMatr(15) = "冲击试验实绩 4                       "
    sMatr(16) = "冲击试验实绩 5                       "
    sMatr(17) = "冲击试验实绩 6                       "
    sMatr(18) = "冲击试验实绩平均                     "
   
    sMatr(19) = "冲击剪切面积实绩平均                 "
    sMatr(20) = "冲击剪切面积实绩 1                   "
    sMatr(21) = "冲击剪切面积实绩 2                   "
    sMatr(22) = "冲击剪切面积实绩 3                   "
    sMatr(23) = "冲击剪切面积实绩 4                   "
    sMatr(24) = "冲击剪切面积实绩 5                   "
    sMatr(25) = "冲击剪切面积实绩 6                   "
    sMatr(26) = "时效冲击试验温度                     "
    sMatr(27) = "时效冲击试样尺寸                     "
    sMatr(28) = "时效冲击功实绩1                      "
    sMatr(29) = "时效冲击功实绩2                      "
    sMatr(30) = "时效冲击功实绩3                      "
    sMatr(31) = "时效冲击功实绩4                      "
    sMatr(32) = "时效冲击功实绩5                      "
    sMatr(33) = "时效冲击功实绩6                      "
    sMatr(34) = "时效冲击实绩平均                     "
                  
    sMatr(35) = "时效冲击纤维断面率实绩               "
    sMatr(36) = "重力撕裂温度                         "
    sMatr(37) = "重力撕裂实绩1                        "
    sMatr(38) = "重力撕裂实绩2                        "
    sMatr(39) = "重力撕裂实绩平均                     "
    sMatr(40) = "硬度实绩                             "
    sMatr(41) = "拉伸规定非比例伸长应力实绩           "
    sMatr(42) = "拉伸规定残余伸长应力实绩实绩         "
    sMatr(43) = "高温拉伸屈服强度实绩                 "
    sMatr(44) = "高温拉伸抗拉强度实绩                 "
    sMatr(45) = "高温拉伸断面收缩率实绩1              "
'20090806 SUN BIN
    sMatr(46) = "高温拉伸断面收缩率实绩2              "
    sMatr(47) = "高温拉伸断面收缩率实绩3              "
    sMatr(48) = "高温拉伸断面收缩率实绩平均           "
'20090806 SUN BIN END
    sMatr(49) = "高温拉伸断后伸长率实绩               "
    sMatr(50) = "高温拉伸规定非比例伸长应力实绩       "
    sMatr(51) = "高温拉伸规定残余伸长应力实绩         "
    sMatr(52) = "焊接硬度实绩                         "
    sMatr(53) = "焊缝弯曲实绩                         "
    sMatr(54) = "反复弯曲实绩                         "
    sMatr(55) = "锻平试验实绩                         "
    sMatr(56) = "抗氢裂能力CSR实绩                    "
    sMatr(57) = "抗氢裂能力CLR实绩                    "
    sMatr(58) = "抗氢裂能力CWR实绩                    "
    sMatr(59) = "硫化物腐蚀裂纹实绩                   "
    sMatr(60) = "追加冲击试验温度                     "
    sMatr(61) = "追加击试样尺寸                       "
    sMatr(62) = "追加冲击试验实绩平均                 "
    sMatr(63) = "追加冲击试验实绩 1                   "
    sMatr(64) = "追加冲击试验实绩 2                   "
    sMatr(65) = "追加冲击试验实绩 3                   "
    sMatr(66) = "追加冲击试验实绩 4                   "
    sMatr(67) = "追加冲击试验实绩 5                   "
    sMatr(68) = "追加冲击试验实绩 6                   "
    sMatr(69) = "追加冲击剪切面积实绩平均             "
    sMatr(70) = "追加冲击剪切面积实绩 1               "
    sMatr(71) = "追加冲击剪切面积实绩 2               "
    sMatr(72) = "追加冲击剪切面积实绩 3               "
    sMatr(73) = "追加冲击剪切面积实绩 4               "
    sMatr(74) = "追加冲击剪切面积实绩 5               "
    sMatr(75) = "追加冲击剪切面积实绩 6               "
    sMatr(76) = "追加时效冲击试验温度                 "
    sMatr(77) = "追加时效冲击试样尺寸                 "
    sMatr(78) = "追加时效冲击实绩平均                 "
    sMatr(79) = "追加时效冲击功实绩1                  "
    sMatr(80) = "追加时效冲击功实绩2                  "
    sMatr(81) = "追加时效冲击功实绩3                  "
    sMatr(82) = "追加时效冲击功实绩4                  "
    sMatr(83) = "追加时效冲击功实绩5                  "
    sMatr(84) = "追加时效冲击功实绩6                  "
    sMatr(85) = "追加时效冲击纤维断面率实绩           "
    sMatr(86) = "晶粒度实绩                           "
    sMatr(87) = "脱碳层实绩                           "
    sMatr(88) = "硫印实绩                             "
    sMatr(89) = "断口检验实绩1                        "
    sMatr(90) = "断口检验实绩2                        "
    sMatr(91) = "断口检验实绩3                        "
    sMatr(92) = "断口检验实绩4                        "
    sMatr(93) = "断口检验实绩5                        "
    sMatr(94) = "酸浸检验实绩1                        "
    sMatr(95) = "酸浸检验实绩2                        "
    sMatr(96) = "酸浸检验实绩3                        "
    sMatr(97) = "酸浸检验实绩4                        "
    sMatr(98) = "酸浸检验实绩5                        "
    sMatr(99) = "带状组织实绩                         "
    sMatr(100) = "淬透性试验实绩1                     "
    sMatr(101) = "淬透性试验实绩2                      "
    sMatr(102) = "淬透性试验实绩3                      "
    sMatr(103) = "非金属夹杂物(粗)实绩1                "
    sMatr(104) = "非金属夹杂物(粗)实绩2                "
    sMatr(105) = "非金属夹杂物(粗)实绩3                "
    sMatr(106) = "非金属夹杂物(粗)实绩4                "
    sMatr(107) = "非金属夹杂物(细)实绩1                "
    sMatr(108) = "非金属夹杂物(细)实绩2                "
    sMatr(109) = "非金属夹杂物(细)实绩3                "
    sMatr(110) = "非金属夹杂物(细)实绩4                "
    sMatr(111) = "奥氏体晶粒度实绩                     "
    sMatr(112) = "DS类非金属夹杂实绩                   "
    sMatr(113) = "TIN类非金属夹杂实绩                  "
'20090804 sun bin start
    sMatr(114) = "追加屈服点实绩                           "
    sMatr(115) = "追加拉伸规定总伸长应力实绩               "
    sMatr(116) = "追加抗拉强度实绩                         "
    sMatr(117) = "追加屈强比实绩                           "
    sMatr(118) = "追加断后伸长率实绩                       "
    sMatr(119) = "追加断面收缩率实绩1                      "
    sMatr(120) = "追加断面收缩率实绩2                      "
    sMatr(121) = "追加断面收缩率实绩3                      "
    sMatr(122) = "追加断面收缩率实绩平均                   "
    sMatr(123) = "追加冷弯试验实绩                         "
    sMatr(124) = "追加硬度实绩                             "
    sMatr(125) = "追加拉伸规定非比例伸长应力实绩           "
    sMatr(126) = "追加拉伸规定残余伸长应力实绩实绩         "
    sMatr(127) = "追加高温拉伸屈服强度实绩                 "
    sMatr(128) = "追加高温拉伸抗拉强度实绩                 "
    sMatr(129) = "追加高温拉伸断面收缩率实绩1              "
'20090806 sun bin start
    sMatr(130) = "追加高温拉伸断面收缩率实绩2              "
    sMatr(131) = "追加高温拉伸断面收缩率实绩3              "
    sMatr(132) = "追加高温拉伸断面收缩率实绩平均           "
'20090806 sun bin end
    sMatr(133) = "追加高温拉伸断后伸长率实绩               "
    sMatr(134) = "追加高温拉伸规定非比例伸长应力实绩       "
    sMatr(135) = "追加高温拉伸规定残余伸长应力实绩         "
'20090804 sun bin end
  
    'louyanan 20101121 start

   sMatr(136) = "拉升厚度方向断面收缩率实绩1"
   sMatr(137) = "拉升厚度方向断面收缩率实绩2"
   sMatr(138) = "拉升厚度方向断面收缩率实绩3"
   
'2016-11-22  ljn  start
   sMatr(139) = "拉升厚度方向断面收缩率实绩4"
   sMatr(140) = "拉升厚度方向断面收缩率实绩5"
   sMatr(141) = "拉升厚度方向断面收缩率实绩6"
'2016-11-22  ljn  end
   
   sMatr(142) = "拉升厚度方向断面收缩率实绩平均"
   

'2016-11-22  ljn  start
   sMatr(143) = "追加拉升厚度方向断面收缩率实绩1"
   sMatr(144) = "追加拉升厚度方向断面收缩率实绩2"
   sMatr(145) = "追加拉升厚度方向断面收缩率实绩3"
   sMatr(146) = "追加拉升厚度方向断面收缩率实绩4"
   sMatr(147) = "追加拉升厚度方向断面收缩率实绩5"
   sMatr(148) = "追加拉升厚度方向断面收缩率实绩6"
   sMatr(149) = "追加拉升厚度方向断面收缩率实绩平均"
   
   sMatr(150) = "厚度方向抗拉强度1"
   sMatr(151) = "厚度方向抗拉强度2"
   sMatr(152) = "厚度方向抗拉强度3"
   '2016-12-2 LJN
   sMatr(153) = "厚度方向抗拉强度4"
   sMatr(154) = "厚度方向抗拉强度5"
   sMatr(155) = "厚度方向抗拉强度6"
   sMatr(156) = "追加厚度方向抗拉强度1"
   sMatr(157) = "追加厚度方向抗拉强度2"
   sMatr(158) = "追加厚度方向抗拉强度3"
   '2016-12-2 LJN
   sMatr(159) = "追加厚度方向抗拉强度4"
   sMatr(160) = "追加厚度方向抗拉强度5"
   sMatr(161) = "追加厚度方向抗拉强度6"
'2016-11-22  ljn  end
   sMatr(162) = "高温拉升厚度方向断面收缩率实绩1"
   sMatr(163) = "高温拉升厚度方向断面收缩率实绩2"
   sMatr(164) = "高温拉升厚度方向断面收缩率实绩3"
   sMatr(165) = "高温拉升厚度方向断面收缩率实绩平均"
   sMatr(166) = "冲击侧向膨胀值实绩1"
   sMatr(167) = "冲击侧向膨胀值实绩2"
   sMatr(168) = "冲击侧向膨胀值实绩3"
   sMatr(169) = "冲击侧向膨胀值实绩4"
   sMatr(170) = "冲击侧向膨胀值实绩5"
   sMatr(171) = "冲击侧向膨胀值实绩6"
   sMatr(172) = "冲击侧向膨胀值实绩平均"
   sMatr(173) = "追加冲击侧向膨胀值实绩1"
   sMatr(174) = "追加冲击侧向膨胀值实绩2"
   sMatr(175) = "追加冲击侧向膨胀值实绩3"
   sMatr(176) = "追加冲击侧向膨胀值实绩4"
   sMatr(177) = "追加冲击侧向膨胀值实绩5"
   sMatr(178) = "追加冲击侧向膨胀值实绩6"
   sMatr(179) = "追加冲击侧向膨胀值实绩平均"
   sMatr(180) = "NDT重力撕裂实绩"
 'edit by gengxueyu 20110212 for kangda start
   sMatr(181) = "均匀变形伸长率UEL"
   sMatr(182) = "追加均匀变形伸长率UEL"
   sMatr(183) = "追加应力比项目1"
   sMatr(184) = "追加应力比项目2"
   sMatr(185) = "追加应力比项目3"
   sMatr(186) = "追加应力比项目4"
   sMatr(187) = "追加应力比项目5"   '165 '181
'edit by gengxueyu 20110212 for kangda end
   sMatr(188) = "断口"
   
    With ss3
        .MaxRows = 189
    
        For i = 1 To 189
            .Row = i
            .Col = 1: .Text = sMatr(i - 1)
        Next i
                
        For i = 1 To UBound(strArr, 1) + 1
        
            .Row = i: .Col = 4
            .Text = NullCheck(strArr(i - 1, 0), "")
            
        Next i
    End With

End Sub

Private Sub subSpreadView3(ByVal strArr As Variant)

    Dim i                     As Integer
    Dim iRow                  As Integer
    Dim sMatr(3, 189)         As Variant
    Dim sMatrCON(6, 189)      As Variant
    Dim sMin, sMax, sFL, sRE  As Variant
    
    If UBound(strArr, 2) < 0 Then Exit Sub
      
    If UBound(strArr, 2) = 0 Then
        For i = 0 To 188
            sMatr(0, i) = NullCheck(strArr(i, 0), "")
        Next i
        
        For i = 0 To 188
            sMatr(1, i) = NullCheck(strArr(i + 189, 0))
        Next i
    
        For i = 0 To 188
            sMatr(2, i) = NullCheck(strArr(i + 378, 0))
        Next i
        
        
        With ss3
                
            For i = 1 To 189
                .Row = i
                .Col = 2: .Text = sMatr(1, i - 1)
                .Col = 3: .Text = sMatr(2, i - 1)
                .Col = 5: .Text = sMatr(0, i - 1)
            Next i
         End With
    End If
     
    If UBound(strArr, 2) = 1 Then
        For i = 0 To 188
            sMatrCON(0, i) = NullCheck(strArr(i, 0), "")
            sMatrCON(3, i) = NullCheck(strArr(i, 1), "")
        Next i
        
        For i = 0 To 188
            sMatrCON(1, i) = NullCheck(strArr(i + 188, 0))
            sMatrCON(4, i) = NullCheck(strArr(i + 188, 1))
        Next i
    
        For i = 0 To 188
            sMatrCON(2, i) = NullCheck(strArr(i + 378, 0))
            sMatrCON(5, i) = NullCheck(strArr(i + 378, 1))
        Next i
        
            
        For i = 1 To 188
            If sMatrCON(0, i - 1) = "A" Or sMatrCON(0, i - 1) = "B" Then
                If sMatrCON(3, i - 1) = "A" Or sMatrCON(3, i - 1) = "B" Then
                   If Val(sMatrCON(1, i - 1)) >= Val(sMatrCON(4, i - 1)) Then
                      sMin = sMatrCON(1, i - 1)
                   Else
                      sMin = sMatrCON(4, i - 1)
                   End If
                   If Val(sMatrCON(2, i - 1)) = 0 Then
                        sMax = sMatrCON(5, i - 1)
                   Else
                        If Val(sMatrCON(2, i - 1)) >= Val(sMatrCON(5, i - 1)) Then
                           sMax = sMatrCON(5, i - 1)
                        Else
                           sMax = sMatrCON(2, i - 1)
                        End If
                   End If
                   sFL = "A"
                Else
                   sFL = "A"
                   sMin = sMatrCON(1, i - 1)
                   sMax = sMatrCON(2, i - 1)
                End If
               
            Else
                  If sMatrCON(3, i - 1) = "A" Or sMatrCON(3, i - 1) = "B" Then
                     sFL = "A"
                     sMin = sMatrCON(4, i - 1)
                     sMax = sMatrCON(5, i - 1)
                  Else
                     sFL = ""
                     sMin = ""
                     sMax = ""
                  End If
                  
            End If
            With ss3
                .Row = i
                .Col = 2: .Text = sMin
                .Col = 3: .Text = sMax
                .Col = 5: .Text = sFL
            End With
            
         Next i
    End If
     
     Call subSpreadCheck1
     Call subSpreadERROR(ss3)
      With ss3
        For i = 1 To .MaxRows
            sRE = Gf_Get_Cell_Value(ss3, i, 4)
            sFL = Gf_Get_Cell_Value(ss3, i, 5)
            If sFL = "A" And sRE = "" Then
             .Col = 4
             .BackColor = RED
            End If
        Next i
      End With
    

End Sub


Private Sub subSpreadCheck1()
    
    Dim i As Long
    Dim j As Long
    
    With ss3
       
       For i = 1 To 189

           If Gf_Get_Cell_Value(ss3, i, 5) <> "A" And Gf_Get_Cell_Value(ss3, i, 5) <> "B" Then
               .Row = i
               .RowHidden = True
           Else
                .RowHidden = False
                j = j + 1
                .Col = 0: .Text = j

           End If
           
           '2016-12-2  LJN
            If Mid(Trim(txt_STDSPEC), 1, 3) <> "API" And Mid(Trim(txt_STDSPEC), 1, 10) <> "GB/T9711.2" Then
                If i = 20 Or i = 21 Or i = 22 Or i = 23 Or i = 24 _
                   Or i = 25 Or i = 26 Or i = 70 Or i = 71 Or i = 72 _
                   Or i = 73 Or i = 74 Or i = 75 Or i = 76 Then
                   .RowHidden = True
                End If
            End If
       Next i
                
    End With
End Sub


Private Sub subSpreadERROR(sPname As vaSpread)
    
    Dim i As Long
    Dim C_DSC_CD, C_RSLT_VAL, C_MAX, C_MIN, C_RESULT, C_FL As Variant

    With sPname
    
       If .MaxRows < 1 Then Exit Sub
       
       For i = 1 To .MaxRows
           .Row = i
           C_DSC_CD = Gf_Get_Cell_Value(sPname, i, 5)
           C_RSLT_VAL = Gf_Get_Cell_Value(sPname, i, 4)
           C_MIN = Val(Gf_Get_Cell_Value(sPname, i, 2))
           C_MAX = Val(Gf_Get_Cell_Value(sPname, i, 3))
           C_RESULT = Val(Gf_Get_Cell_Value(sPname, i, 4))
           If C_MIN <> 0 And C_MAX <> 0 Then
              If C_RESULT > C_MAX Or C_RESULT < C_MIN Then
                 Call Gp_Sp_CellColor(sPname, 4, i, RED)
              End If
           Else
              If C_MIN = 0 And C_MAX <> 0 Then
                 If C_RESULT > C_MAX Then
                    Call Gp_Sp_CellColor(sPname, 4, i, RED)
                 End If
              Else
                 If C_MIN <> 0 And C_MAX = 0 Then
                    If C_RESULT < C_MIN Then
                      Call Gp_Sp_CellColor(sPname, 4, i, RED)
                    End If
                 End If
              End If
           End If
           If C_DSC_CD = "A" Or C_DSC_CD = "B" Then
              If C_RSLT_VAL = "N" Then
                 Call Gp_Sp_CellColor(sPname, 4, i, RED)
              End If
           End If
       Next i
 
    End With
    
End Sub

Public Sub DataSave()
    Dim iRow, iCol As Integer
    
    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0050C.P_MODIFY_1", Key:="P-M"
  
    With ss1
           .Row = .ActiveRow
           .Col = 0
           .Text = "Update"
           .Col = 39
           .Text = sUserID
    End With
    
    If Gf_Sp_Process(M_CN1, Sc1, Mc1) Then Call MDIMain.FormMenuSetting(Me, FormType, "RE", sAuthority)
    
    ss1.OperationMode = OperationModeNormal
    Call MenuToolSet
    
    If ss1.MaxRows < 1 Or ss1.ActiveRow = 0 Then Exit Sub
    
    With ss1
         For iRow = 1 To .MaxRows
            .Row = iRow
            .Col = 45
            If .Text = "N" Then
               Call Gp_Sp_BlockColor(ss1, 1, ss1.MaxCols, iRow, iRow, , &HFF80FF)
            End If
         Next iRow
    End With
    
    Sc1.Remove ("P-M")
    Sc1.Add Item:="AQC0050C.P_MODIFY", Key:="P-M"

End Sub

Private Sub subSpreadView_Config(ByVal strArr As Variant)

    Dim i As Integer
    Dim OLD_MAXROWS As Integer
    Dim sMin, sMax, sFL, sRE  As Variant
    
    If UBound(strArr, 2) < 0 Then Exit Sub
    
    With ss3
        OLD_MAXROWS = .MaxRows
        .MaxRows = .MaxRows + UBound(strArr, 2) + 1

        For i = 1 To UBound(strArr, 2) + 1
            .Row = OLD_MAXROWS + i
            .Col = 1: .Text = GF_NullChange(strArr(0, i - 1))
            .Col = 2: .Text = GF_NullChange(strArr(1, i - 1)) & ""
            .Col = 3: .Text = GF_NullChange(strArr(2, i - 1)) & ""
            .Col = 4: .Text = GF_NullChange(strArr(3, i - 1)) & ""
            .Col = 5: .Text = GF_NullChange(strArr(4, i - 1)) & ""
            .Col = 6: .CellType = CellTypeCheckBox
        Next i
            
    End With
    
   'subSpreadCheck1 隐藏试验项目或判定代码为空的行，且重写第一列顺序号
    'Call subSpreadCheck1
    
    '检查上下限，超标实绩红字
    Call subSpreadERROR(ss3)
    
    '检查有无实绩，如果有判定代码但没实绩，底色变红
      With ss3
        For i = 1 To .MaxRows
            sRE = Gf_Get_Cell_Value(ss3, i, 4)
            sFL = Gf_Get_Cell_Value(ss3, i, 5)
            If sFL = "A" And sRE = "" Then
             .Col = 4
             .BackColor = RED
            End If
        Next i
      End With

End Sub

 '--------------------配置化项目复样  王成  2012.12.12-------------保存配置化项目同时添加Update标志------------------------------
 
Private Sub ss3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    Dim i As Integer
    Dim j As Integer
 If ss1.ActiveRow > 0 Then
        For i = 188 To ss3.MaxRows
        With ss3
        .Row = i
        .Col = 6
         With ss1
         For j = 0 To UBound(xy, 2)
               ss1.Row = ss1.ActiveRow
               ss1.Col = 1
               If xy(1, j) = i And xy(0, j) = ss1.Text Then
               xy(0, j) = ""
               xy(1, j) = ""
               End If
        Next j
            If ss3.Text = "1" Then
               xy(0, UBound(xy, 2)) = ss1.Text
               xy(1, UBound(xy, 2)) = i
               ReDim Preserve xy(1, UBound(xy, 2) + 1)
               ss1.Row = ss1.ActiveRow
               ss1.Col = 0
               ss1.Text = "Update"
            End If
         End With
         End With
         Next i
  End If
  
End Sub


